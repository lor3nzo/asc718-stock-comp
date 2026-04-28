"""
ASC 718 Stock Compensation MVP
Single-file Streamlit app

Run:
    pip install streamlit pandas numpy scipy openpyxl xlsxwriter
    streamlit run app.py

Scope:
- Local-first MVP for startup CFOs/controllers
- Excel upload plus built-in sample data
- Downloadable Excel upload template
- Options, RSUs, performance awards
- Black-Scholes for options
- Straight-line vesting
- Manual forfeiture events
- Custom reporting window
- CSV and Excel exports

Important:
This is a prototype. It is not audit-certified ASC 718 software.
"""

from __future__ import annotations

import io
import math
from datetime import date
from typing import Optional

import pandas as pd
import streamlit as st
from scipy.stats import norm


# ==========================================================
# App setup
# ==========================================================

st.set_page_config(page_title="ASC 718 Stock Comp MVP", layout="wide")

st.title("ASC 718 Stock Compensation MVP")
st.caption("Local-first valuation + monthly waterfall engine for startup CFOs/controllers")


# ==========================================================
# Default assumptions
# ==========================================================

DEFAULT_RISK_FREE_RATE = 0.04
DEFAULT_VOLATILITY = 0.55
DEFAULT_EXPECTED_TERM_YEARS = 6.0
DEFAULT_DIVIDEND_YIELD = 0.00
DEFAULT_PROBABLE_THRESHOLD = 0.70

REQUIRED_COLUMNS = [
    "grant_id",
    "employee_name",
    "award_type",
    "grant_date",
    "vest_start_date",
    "vest_end_date",
    "shares",
    "strike_price",
    "grant_date_fmv",
    "performance_probability",
]

OPTIONAL_ASSUMPTION_OVERRIDE_COLUMNS = [
    "risk_free_rate_override",
    "volatility_override",
    "expected_term_override",
    "dividend_yield_override",
]

ALLOWED_AWARD_TYPES = {"option", "rsu", "performance"}
MAX_UPLOAD_BYTES = 5 * 1024 * 1024  # 5 MB

KNOWN_LIMITATIONS = [
    "No graded vesting",
    "No tranche-level modeling",
    "No probability change tracking",
    "No cumulative catch-up logic",
    "No modifications / repricing",
    "No integrations",
    "No persistence",
]


# ==========================================================
# Sample data
# ==========================================================

SAMPLE_GRANTS = pd.DataFrame(
    [
        {
            "grant_id": "OPT-001",
            "employee_name": "Ava Chen",
            "award_type": "option",
            "grant_date": "2024-01-15",
            "vest_start_date": "2024-01-15",
            "vest_end_date": "2028-01-15",
            "shares": 10000,
            "strike_price": 1.00,
            "grant_date_fmv": 1.25,
            "performance_probability": 1.00,
        },
        {
            "grant_id": "RSU-001",
            "employee_name": "Marcus Lee",
            "award_type": "rsu",
            "grant_date": "2024-03-01",
            "vest_start_date": "2024-03-01",
            "vest_end_date": "2027-03-01",
            "shares": 5000,
            "strike_price": 0.00,
            "grant_date_fmv": 2.10,
            "performance_probability": 1.00,
        },
        {
            "grant_id": "PERF-001",
            "employee_name": "Priya Shah",
            "award_type": "performance",
            "grant_date": "2024-06-01",
            "vest_start_date": "2024-06-01",
            "vest_end_date": "2026-06-01",
            "shares": 8000,
            "strike_price": 0.00,
            "grant_date_fmv": 3.00,
            "performance_probability": 0.65,
        },
        {
            "grant_id": "PERF-002",
            "employee_name": "Noah Patel",
            "award_type": "performance",
            "grant_date": "2024-06-01",
            "vest_start_date": "2024-06-01",
            "vest_end_date": "2026-06-01",
            "shares": 6000,
            "strike_price": 0.00,
            "grant_date_fmv": 3.00,
            "performance_probability": 0.80,
        },
        {
            "grant_id": "OPT-002",
            "employee_name": "Jordan Smith",
            "award_type": "option",
            "grant_date": "2023-09-01",
            "vest_start_date": "2023-09-01",
            "vest_end_date": "2027-09-01",
            "shares": 12000,
            "strike_price": 0.75,
            "grant_date_fmv": 1.10,
            "performance_probability": 1.00,
        },
        {
            "grant_id": "RSU-002",
            "employee_name": "Elena Garcia",
            "award_type": "rsu",
            "grant_date": "2022-05-01",
            "vest_start_date": "2022-05-01",
            "vest_end_date": "2026-05-01",
            "shares": 3000,
            "strike_price": 0.00,
            "grant_date_fmv": 0.90,
            "performance_probability": 1.00,
        },
    ]
)

SAMPLE_FORFEITURES = pd.DataFrame(
    [
        {
            "grant_id": "OPT-002",
            "forfeiture_date": "2025-02-15",
            "reason": "Employee termination",
        }
    ]
)


# ==========================================================
# Date and math helpers
# ==========================================================

def parse_date(value) -> pd.Timestamp:
    if pd.isna(value):
        return pd.NaT
    return pd.to_datetime(value, errors="coerce").normalize()


def days_between(start: pd.Timestamp, end: pd.Timestamp) -> int:
    if pd.isna(start) or pd.isna(end):
        return 0
    return max((end - start).days, 0)


def month_range(start_date: date, end_date: date) -> pd.DatetimeIndex:
    start = pd.to_datetime(start_date).to_period("M").to_timestamp("M")
    end = pd.to_datetime(end_date).to_period("M").to_timestamp("M")
    return pd.date_range(start=start, end=end, freq="M")


def historical_month_range(grants: pd.DataFrame, report_end: date) -> pd.DatetimeIndex:
    """
    Compute from earliest vesting start date so beginning balances inside
    a custom reporting window reflect prior recognized expense.
    """
    earliest_vest_start = pd.to_datetime(grants["vest_start_date"]).min().date()
    return month_range(earliest_vest_start, report_end)


def overlap_days(
    period_start: pd.Timestamp,
    period_end: pd.Timestamp,
    service_start: pd.Timestamp,
    service_end: pd.Timestamp,
) -> int:
    start = max(period_start, service_start)
    end = min(period_end, service_end)
    return days_between(start, end)


def black_scholes_call_value(
    stock_price: float,
    strike_price: float,
    expected_term_years: float,
    risk_free_rate: float,
    volatility: float,
    dividend_yield: float,
) -> float:
    """
    Black-Scholes call option value.

    ASC 718 concept:
    Plain-vanilla employee stock options are commonly measured at grant-date
    fair value using an option-pricing model such as Black-Scholes.
    """
    if stock_price <= 0 or strike_price <= 0 or expected_term_years <= 0 or volatility <= 0:
        return 0.0

    d1 = (
        math.log(stock_price / strike_price)
        + (risk_free_rate - dividend_yield + 0.5 * volatility**2) * expected_term_years
    ) / (volatility * math.sqrt(expected_term_years))
    d2 = d1 - volatility * math.sqrt(expected_term_years)

    value = (
        stock_price * math.exp(-dividend_yield * expected_term_years) * norm.cdf(d1)
        - strike_price * math.exp(-risk_free_rate * expected_term_years) * norm.cdf(d2)
    )
    return max(value, 0.0)


# ==========================================================
# Validation and normalization
# ==========================================================

def validate_grants(df: pd.DataFrame) -> list[str]:
    errors = []

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        return [f"Missing required columns: {', '.join(missing)}"]

    temp = df.copy()
    temp["award_type"] = temp["award_type"].astype(str).str.lower().str.strip()

    invalid_types = set(temp["award_type"].dropna()) - ALLOWED_AWARD_TYPES
    if invalid_types:
        errors.append(f"Invalid award_type values: {', '.join(sorted(invalid_types))}. Use option, rsu, or performance.")

    for col in ["grant_date", "vest_start_date", "vest_end_date"]:
        parsed = pd.to_datetime(temp[col], errors="coerce")
        bad_count = parsed.isna().sum()
        if bad_count > 0:
            errors.append(f"Column '{col}' has {bad_count} invalid or blank date value(s). Use YYYY-MM-DD.")

    for col in ["shares", "strike_price", "grant_date_fmv", "performance_probability"]:
        parsed = pd.to_numeric(temp[col], errors="coerce")
        bad_count = parsed.isna().sum()
        if bad_count > 0:
            errors.append(f"Column '{col}' has {bad_count} invalid or blank numeric value(s).")

    for col in OPTIONAL_ASSUMPTION_OVERRIDE_COLUMNS:
        if col in temp.columns:
            parsed = pd.to_numeric(temp[col], errors="coerce")
            bad_count = parsed[temp[col].notna()].isna().sum()
            if bad_count > 0:
                errors.append(f"Optional column '{col}' has {bad_count} invalid numeric override value(s).")

    if "risk_free_rate_override" in temp.columns:
        parsed = pd.to_numeric(temp["risk_free_rate_override"], errors="coerce")
        if ((parsed < 0) | (parsed > 1)).any():
            errors.append("risk_free_rate_override must be between 0 and 1 when provided.")

    if "volatility_override" in temp.columns:
        parsed = pd.to_numeric(temp["volatility_override"], errors="coerce")
        if ((parsed <= 0) | (parsed > 3)).any():
            errors.append("volatility_override must be greater than 0 and no more than 3 when provided.")

    if "expected_term_override" in temp.columns:
        parsed = pd.to_numeric(temp["expected_term_override"], errors="coerce")
        if ((parsed <= 0) | (parsed > 10)).any():
            errors.append("expected_term_override must be greater than 0 and no more than 10 when provided.")

    if "dividend_yield_override" in temp.columns:
        parsed = pd.to_numeric(temp["dividend_yield_override"], errors="coerce")
        if ((parsed < 0) | (parsed > 1)).any():
            errors.append("dividend_yield_override must be between 0 and 1 when provided.")

    shares = pd.to_numeric(temp["shares"], errors="coerce")
    if (shares <= 0).any():
        errors.append("Shares must be greater than zero for every grant.")

    fmv = pd.to_numeric(temp["grant_date_fmv"], errors="coerce")
    if (fmv <= 0).any():
        errors.append("grant_date_fmv must be greater than zero for every grant.")

    strike = pd.to_numeric(temp["strike_price"], errors="coerce")
    option_rows = temp["award_type"] == "option"
    non_option_rows = temp["award_type"].isin(["rsu", "performance"])

    if option_rows.any() and (strike[option_rows] <= 0).any():
        errors.append("Option awards must have strike_price greater than zero.")

    if non_option_rows.any() and (strike[non_option_rows] < 0).any():
        errors.append("RSU and performance strike_price cannot be negative. Use 0 if not applicable.")

    probs = pd.to_numeric(temp["performance_probability"], errors="coerce")
    if ((probs < 0) | (probs > 1)).any():
        errors.append("performance_probability must be between 0 and 1.")

    vest_start = pd.to_datetime(temp["vest_start_date"], errors="coerce")
    vest_end = pd.to_datetime(temp["vest_end_date"], errors="coerce")
    grant_date = pd.to_datetime(temp["grant_date"], errors="coerce")

    if (vest_end <= vest_start).any():
        errors.append("vest_end_date must be after vest_start_date for every grant.")

    if (grant_date > vest_end).any():
        errors.append("grant_date should not be after vest_end_date.")

    if (grant_date > vest_start).any():
        errors.append("grant_date should not be after vest_start_date.")

    duplicate_count = temp["grant_id"].astype(str).duplicated().sum()
    if duplicate_count > 0:
        errors.append(f"Found {duplicate_count} duplicate grant_id value(s). grant_id must be unique.")

    return errors



def validate_forfeitures(forfeitures: pd.DataFrame, grants: pd.DataFrame) -> list[str]:
    errors = []
    if forfeitures.empty:
        return errors

    valid_grant_ids = set(grants["grant_id"].astype(str))
    f = forfeitures.copy()

    # Rule 1: grant_id must exist in grants
    invalid_ids = set(f["grant_id"].astype(str).dropna()) - valid_grant_ids
    if invalid_ids:
        errors.append(f"Forfeiture grant_id(s) not found in grants: {', '.join(sorted(invalid_ids))}")

    # Rule 2: forfeiture_date must not be null
    if f["forfeiture_date"].isna().any():
        errors.append("One or more forfeiture events are missing a forfeiture_date.")

    # Rule 3: forfeiture_date must be >= vest_start_date for that grant
    merged = f.merge(
        grants[["grant_id", "vest_start_date"]].astype({"grant_id": str}),
        on="grant_id", how="left"
    )
    bad_dates = merged[
        merged["forfeiture_date"].notna()
        & (pd.to_datetime(merged["forfeiture_date"]) < pd.to_datetime(merged["vest_start_date"]))
    ]
    if not bad_dates.empty:
        errors.append(
            f"forfeiture_date is before vest_start_date for grant(s): {', '.join(bad_dates['grant_id'].tolist())}"
        )

    # Rule 4: only one forfeiture per grant
    duplicates = f["grant_id"].astype(str).dropna()
    duplicates = duplicates[duplicates.duplicated()]
    if not duplicates.empty:
        errors.append(
            f"Multiple forfeiture events found for grant(s): {', '.join(duplicates.unique().tolist())}. "
            "Only one forfeiture per grant is supported."
        )

    return errors
def normalize_grants(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["award_type"] = out["award_type"].astype(str).str.lower().str.strip()

    for col in ["grant_date", "vest_start_date", "vest_end_date"]:
        out[col] = out[col].apply(parse_date)

    for col in ["shares", "strike_price", "grant_date_fmv", "performance_probability"]:
        out[col] = pd.to_numeric(out[col], errors="coerce")

    out["performance_probability"] = out["performance_probability"].clip(0.0, 1.0)

    for col in OPTIONAL_ASSUMPTION_OVERRIDE_COLUMNS:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce")

    return out


def assumption_override_value(grant: dict | pd.Series, override_col: str, default_value: float) -> float:
    value = grant.get(override_col, None)
    if pd.notna(value):
        return float(value)
    return float(default_value)


# ==========================================================
# Template and export helpers
# ==========================================================

@st.cache_data
def to_template_excel_bytes() -> bytes:
    """Create downloadable Excel upload template."""
    template = pd.DataFrame(
        [
            {
                "grant_id": "OPT-EXAMPLE-001",
                "employee_name": "Example Employee",
                "award_type": "option",
                "grant_date": "2024-01-15",
                "vest_start_date": "2024-01-15",
                "vest_end_date": "2028-01-15",
                "shares": 10000,
                "strike_price": 1.00,
                "grant_date_fmv": 1.25,
                "performance_probability": 1.00,
                "risk_free_rate_override": None,
                "volatility_override": None,
                "expected_term_override": None,
                "dividend_yield_override": None,
            },
            {
                "grant_id": "RSU-EXAMPLE-001",
                "employee_name": "Example Employee",
                "award_type": "rsu",
                "grant_date": "2024-03-01",
                "vest_start_date": "2024-03-01",
                "vest_end_date": "2027-03-01",
                "shares": 5000,
                "strike_price": 0.00,
                "grant_date_fmv": 2.10,
                "performance_probability": 1.00,
                "risk_free_rate_override": None,
                "volatility_override": None,
                "expected_term_override": None,
                "dividend_yield_override": None,
            },
            {
                "grant_id": "PERF-EXAMPLE-001",
                "employee_name": "Example Employee",
                "award_type": "performance",
                "grant_date": "2024-06-01",
                "vest_start_date": "2024-06-01",
                "vest_end_date": "2026-06-01",
                "shares": 8000,
                "strike_price": 0.00,
                "grant_date_fmv": 3.00,
                "performance_probability": 0.75,
                "risk_free_rate_override": None,
                "volatility_override": None,
                "expected_term_override": None,
                "dividend_yield_override": None,
            },
        ]
    )

    instructions = pd.DataFrame(
        [
            {"field": "grant_id", "rule": "Unique grant identifier. Required."},
            {"field": "employee_name", "rule": "Employee or service provider name. Required."},
            {"field": "award_type", "rule": "Allowed values: option, rsu, performance."},
            {"field": "grant_date", "rule": "Use YYYY-MM-DD format."},
            {"field": "vest_start_date", "rule": "Use YYYY-MM-DD format."},
            {"field": "vest_end_date", "rule": "Must be after vest_start_date. Use YYYY-MM-DD format."},
            {"field": "shares", "rule": "Must be greater than zero."},
            {"field": "strike_price", "rule": "Options require greater than zero. Use 0 for RSUs/performance awards."},
            {"field": "grant_date_fmv", "rule": "Grant-date fair market value. Must be greater than zero."},
            {"field": "performance_probability", "rule": "Decimal between 0 and 1. Use 1 for options and RSUs."},
            {"field": "risk_free_rate_override", "rule": "Optional decimal override for option awards. Blank uses sidebar assumption."},
            {"field": "volatility_override", "rule": "Optional decimal override for option awards. Blank uses sidebar assumption."},
            {"field": "expected_term_override", "rule": "Optional years override for option awards. Blank uses sidebar assumption."},
            {"field": "dividend_yield_override", "rule": "Optional decimal override for option awards. Blank uses sidebar assumption."},
        ]
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        template.to_excel(writer, index=False, sheet_name="Upload Template")
        instructions.to_excel(writer, index=False, sheet_name="Instructions")

        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
        money_format = workbook.add_format({"num_format": "$#,##0.00"})
        percent_format = workbook.add_format({"num_format": "0.0%"})

        template_ws = writer.sheets["Upload Template"]
        instructions_ws = writer.sheets["Instructions"]

        for col_num, value in enumerate(template.columns.values):
            template_ws.write(0, col_num, value, header_format)

        for col_num, value in enumerate(instructions.columns.values):
            instructions_ws.write(0, col_num, value, header_format)

        template_ws.set_column(0, 1, 22)
        template_ws.set_column(2, 2, 16)
        template_ws.set_column(3, 5, 16, date_format)
        template_ws.set_column(6, 6, 12)
        template_ws.set_column(7, 8, 16, money_format)
        template_ws.set_column(9, 13, 22, percent_format)
        instructions_ws.set_column(0, 0, 28)
        instructions_ws.set_column(1, 1, 90)

    return output.getvalue()


@st.cache_data
def to_excel_bytes(
    summary: pd.DataFrame,
    detail: pd.DataFrame,
    grants: pd.DataFrame,
    forfeitures: pd.DataFrame,
    assumptions: pd.DataFrame,
    audit_awareness_mode: bool = False,
) -> bytes:
    disclaimer_rows = [
        {"item": "Tool name", "detail": "ASC 718 Stock Compensation MVP"},
        {"item": "Status", "detail": "Internal use only. Not audit-certified ASC 718 software."},
        {"item": "Valuation assumptions", "detail": "User-entered. Must be supported by external sources before audit or board use."},
        {"item": "Expected term", "detail": "Must comply with SAB 107 / SAB 110 policy."},
        {"item": "Volatility", "detail": "Must be supported by peer group volatility study."},
        {"item": "Risk-free rate", "detail": "Must reference Treasury yield curve at grant date."},
        {"item": "Scope limitations", "detail": "No graded vesting, no tranche modeling, no modifications, no repricing, no integrations."},
    ]

    if audit_awareness_mode:
        disclaimer_rows.append({"item": "Limitations", "detail": "; ".join(KNOWN_LIMITATIONS)})

    disclaimer = pd.DataFrame(disclaimer_rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        disclaimer.to_excel(writer, index=False, sheet_name="Disclaimer")
        summary.to_excel(writer, index=False, sheet_name="Summary Waterfall")
        detail.to_excel(writer, index=False, sheet_name="Grant Detail")
        grants.to_excel(writer, index=False, sheet_name="Valued Grants")
        assumptions.to_excel(writer, index=False, sheet_name="Assumptions")
        forfeitures.to_excel(writer, index=False, sheet_name="Forfeitures")

        workbook = writer.book
        money_format = workbook.add_format({"num_format": "$#,##0.00"})
        percent_format = workbook.add_format({"num_format": "0.0%"})
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})

        for sheet_name in writer.sheets:
            writer.sheets[sheet_name].set_column(0, 30, 18)

        for sheet_name in ["Summary Waterfall", "Grant Detail", "Valued Grants"]:
            ws = writer.sheets[sheet_name]
            ws.set_column(0, 0, 14, date_format)
            ws.set_column(6, 25, 18, money_format)

        writer.sheets["Assumptions"].set_column(1, 1, 24, percent_format)

    return output.getvalue()



def sanitize_for_export(value: str) -> str:
    """Prevent CSV/Excel formula injection by prefixing dangerous leading characters."""
    dangerous = ("=", "+", "-", "@")
    s = str(value)
    if s.startswith(dangerous):
        return "'" + s
    return s


def sanitize_dataframe_for_export(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    string_cols = out.select_dtypes(include="object").columns.tolist()
    for col in string_cols:
        out[col] = out[col].astype(str).apply(sanitize_for_export)
    return out

def format_currency_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    cols = [
        "fair_value_per_share",
        "total_fair_value",
        "recognized_fair_value_basis",
        "beginning_unrecognized_comp",
        "gross_monthly_expense",
        "early_termination_reversal",
        "current_month_expense",
        "ending_unrecognized_comp",
        "cumulative_expense",
        "forfeited_unrecognized_cost",
    ]
    for col in cols:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0).round(2)
    return out


# ==========================================================
# Valuation engine
# ==========================================================

@st.cache_data
def calculate_fair_values(
    grants: pd.DataFrame,
    risk_free_rate: float,
    volatility: float,
    expected_term_years: float,
    dividend_yield: float,
) -> pd.DataFrame:
    rows = []

    for grant_tuple in grants.itertuples(index=False):
        grant = grant_tuple._asdict()
        award_type = grant["award_type"]

        if award_type == "option":
            risk_free_rate_for_grant = assumption_override_value(grant, "risk_free_rate_override", risk_free_rate)
            volatility_for_grant = assumption_override_value(grant, "volatility_override", volatility)
            expected_term_for_grant = assumption_override_value(grant, "expected_term_override", expected_term_years)
            dividend_yield_for_grant = assumption_override_value(grant, "dividend_yield_override", dividend_yield)

            fair_value_per_share = black_scholes_call_value(
                stock_price=float(grant["grant_date_fmv"]),
                strike_price=float(grant["strike_price"]),
                expected_term_years=expected_term_for_grant,
                risk_free_rate=risk_free_rate_for_grant,
                volatility=volatility_for_grant,
                dividend_yield=dividend_yield_for_grant,
            )
            valuation_method = "Black-Scholes"

        elif award_type == "rsu":
            # ASC 718 concept:
            # RSUs usually use grant-date FMV because there is no exercise price.
            fair_value_per_share = float(grant["grant_date_fmv"])
            valuation_method = "Grant-date FMV"

        elif award_type == "performance":
            # ASC 718 concept:
            # Performance awards are measured at grant-date fair value.
            # Expense recognition depends on whether the performance condition is probable.
            fair_value_per_share = float(grant["grant_date_fmv"])
            valuation_method = "Grant-date FMV, probable recognition test"

        else:
            fair_value_per_share = 0.0
            valuation_method = "Unknown"

        if award_type != "option":
            risk_free_rate_for_grant = None
            volatility_for_grant = None
            expected_term_for_grant = None
            dividend_yield_for_grant = None

        total_fair_value = fair_value_per_share * float(grant["shares"])

        row = dict(grant)
        row.update(
            {
                "valuation_method": valuation_method,
                "fair_value_per_share": fair_value_per_share,
                "total_fair_value": total_fair_value,
                "risk_free_rate_used": risk_free_rate_for_grant,
                "volatility_used": volatility_for_grant,
                "expected_term_years_used": expected_term_for_grant,
                "dividend_yield_used": dividend_yield_for_grant,
            }
        )
        rows.append(row)

    return pd.DataFrame(rows)


# ==========================================================
# Expense engine
# ==========================================================

def build_forfeiture_map(forfeitures: pd.DataFrame) -> dict[str, pd.Timestamp]:
    if forfeitures.empty:
        return {}

    f = forfeitures.copy()
    f["forfeiture_date"] = f["forfeiture_date"].apply(parse_date)
    f = f.dropna(subset=["grant_id", "forfeiture_date"])
    f = f.sort_values("forfeiture_date")
    return f.groupby("grant_id")["forfeiture_date"].first().to_dict()


def calculate_grant_monthly_schedule(
    grant: dict | pd.Series,
    months: pd.DatetimeIndex,
    forfeiture_date: Optional[pd.Timestamp],
    probable_threshold: float,
) -> pd.DataFrame:
    service_start = grant["vest_start_date"]
    service_end = grant["vest_end_date"]

    if pd.notna(forfeiture_date):
        effective_service_end = min(service_end, forfeiture_date)
        forfeited = True
    else:
        effective_service_end = service_end
        forfeited = False

    total_service_days = days_between(service_start, service_end)
    effective_service_days = days_between(service_start, effective_service_end)
    total_fair_value = float(grant["total_fair_value"])

    # Performance award recognition:
    # If not probable, no expense is recognized in this MVP.
    # If probable, full grant-date fair value is recognized over the service period.
    performance_probability = float(grant.get("performance_probability", 1.0))
    is_performance_award = grant["award_type"] == "performance"
    is_probable = (not is_performance_award) or (performance_probability >= probable_threshold)
    recognizable_fair_value = total_fair_value if is_probable else 0.0

    daily_expense = recognizable_fair_value / total_service_days if total_service_days > 0 else 0.0
    max_recognizable_cost = daily_expense * effective_service_days

    rows = []
    cumulative_gross_expense = 0.0
    cumulative_forfeiture_adjustment = 0.0
    forfeiture_adjustment_recorded = False

    for month_end in months:
        period_start = month_end.to_period("M").to_timestamp()
        period_end = month_end

        service_days_in_month = overlap_days(
            period_start=period_start,
            period_end=period_end,
            service_start=service_start,
            service_end=effective_service_end,
        )

        gross_expense = daily_expense * service_days_in_month

        if cumulative_gross_expense + gross_expense > max_recognizable_cost:
            gross_expense = max(max_recognizable_cost - cumulative_gross_expense, 0.0)

        beginning_cumulative_net_expense = cumulative_gross_expense + cumulative_forfeiture_adjustment
        beginning_unrecognized = max(recognizable_fair_value - beginning_cumulative_net_expense, 0.0) if is_probable else 0.0

        cumulative_gross_expense += gross_expense

        forfeiture_adjustment = 0.0
        forfeited_unrecognized_cost = 0.0

        # Forfeiture adjustment:
        # In the forfeiture month, reverse any cumulative expense above cost earned
        # through the actual service period. This is simplified grant-level logic.
        if forfeited and not forfeiture_adjustment_recorded and period_start <= forfeiture_date <= period_end:
            earned_cost_through_forfeiture = max_recognizable_cost
            cumulative_before_adjustment = cumulative_gross_expense + cumulative_forfeiture_adjustment
            forfeiture_adjustment = min(earned_cost_through_forfeiture - cumulative_before_adjustment, 0.0)
            cumulative_forfeiture_adjustment += forfeiture_adjustment
            forfeiture_adjustment_recorded = True

        cumulative_net_expense = cumulative_gross_expense + cumulative_forfeiture_adjustment
        ending_unrecognized = max(recognizable_fair_value - cumulative_net_expense, 0.0) if is_probable else 0.0

        if forfeited and month_end >= forfeiture_date:
            forfeited_unrecognized_cost = max(total_fair_value - max_recognizable_cost, 0.0)

        rows.append(
            {
                "month": month_end.date(),
                "grant_id": grant["grant_id"],
                "employee_name": grant["employee_name"],
                "award_type": grant["award_type"],
                "valuation_method": grant["valuation_method"],
                "shares": grant["shares"],
                "fair_value_per_share": grant["fair_value_per_share"],
                "total_fair_value": total_fair_value,
                "performance_probability": performance_probability,
                "probable_threshold": probable_threshold if is_performance_award else None,
                "performance_condition_probable": is_probable if is_performance_award else None,
                "recognized_fair_value_basis": recognizable_fair_value,
                "forfeiture_date": forfeiture_date.date() if pd.notna(forfeiture_date) else None,
                "beginning_unrecognized_comp": beginning_unrecognized,
                "gross_monthly_expense": gross_expense,
                "early_termination_reversal": forfeiture_adjustment,
                "current_month_expense": gross_expense + forfeiture_adjustment,
                "ending_unrecognized_comp": ending_unrecognized,
                "cumulative_expense": cumulative_net_expense,
                "forfeited_unrecognized_cost": forfeited_unrecognized_cost,
            }
        )

    return pd.DataFrame(rows)


@st.cache_data
def build_full_schedule(
    valued_grants: pd.DataFrame,
    forfeitures: pd.DataFrame,
    start_date: date,
    end_date: date,
    probable_threshold: float,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    months = historical_month_range(valued_grants, end_date)
    forfeiture_map = build_forfeiture_map(forfeitures)

    detail_frames = []
    for grant_tuple in valued_grants.itertuples(index=False):
        grant = grant_tuple._asdict()
        forfeiture_date = forfeiture_map.get(grant["grant_id"], pd.NaT)
        detail_frames.append(
            calculate_grant_monthly_schedule(
                grant=grant,
                months=months,
                forfeiture_date=forfeiture_date,
                probable_threshold=probable_threshold,
            )
        )

    detail = pd.concat(detail_frames, ignore_index=True) if detail_frames else pd.DataFrame()

    if detail.empty:
        return pd.DataFrame(), pd.DataFrame()

    report_start_month = pd.to_datetime(start_date).to_period("M").to_timestamp("M").date()
    report_end_month = pd.to_datetime(end_date).to_period("M").to_timestamp("M").date()
    detail = detail[(detail["month"] >= report_start_month) & (detail["month"] <= report_end_month)].copy()

    summary = (
        detail.groupby("month", as_index=False)
        .agg(
            beginning_unrecognized_comp=("beginning_unrecognized_comp", "sum"),
            gross_monthly_expense=("gross_monthly_expense", "sum"),
            early_termination_reversal=("early_termination_reversal", "sum"),
            current_month_expense=("current_month_expense", "sum"),
            ending_unrecognized_comp=("ending_unrecognized_comp", "sum"),
            cumulative_expense=("cumulative_expense", "sum"),
            forfeited_unrecognized_cost=("forfeited_unrecognized_cost", "sum"),
        )
    )

    return summary, detail


# ==========================================================
# Sidebar
# ==========================================================

st.sidebar.header("Settings")

st.sidebar.download_button(
    label="Download Excel upload template",
    data=to_template_excel_bytes(),
    file_name="asc718_grant_upload_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

use_sample_data = st.sidebar.toggle("Use sample dataset", value=True)
audit_awareness_mode = st.sidebar.toggle("Audit Awareness Mode", value=False)

uploaded_file = st.sidebar.file_uploader(
    "Upload grant file (.xlsx)",
    type=["xlsx"],
    help="Use the template above or match its column structure.",
)

st.sidebar.subheader("Valuation assumptions")
st.sidebar.caption("MVP defaults. Replace with sourced support before audit use.")

risk_free_rate = st.sidebar.number_input("Risk-free rate", 0.0, 1.0, DEFAULT_RISK_FREE_RATE, 0.005, format="%.3f")
volatility = st.sidebar.number_input("Volatility", 0.0, 3.0, DEFAULT_VOLATILITY, 0.025, format="%.3f")
expected_term_years = st.sidebar.number_input("Expected term years", 0.1, 10.0, DEFAULT_EXPECTED_TERM_YEARS, 0.25)
dividend_yield = st.sidebar.number_input("Dividend yield", 0.0, 1.0, DEFAULT_DIVIDEND_YIELD, 0.005, format="%.3f")
probable_threshold = st.sidebar.number_input("Performance probable threshold", 0.0, 1.0, DEFAULT_PROBABLE_THRESHOLD, 0.05, format="%.2f")

assumption_source_note = st.sidebar.text_area(
    "Assumption source note",
    value="MVP placeholder assumptions. Replace with Treasury curve, peer volatility study, SAB 107/110 expected-term policy, and dividend yield support before audit use.",
    height=110,
)

st.sidebar.subheader("Reporting window")
report_start = st.sidebar.date_input("Start date", value=date(2024, 1, 1))
report_end = st.sidebar.date_input("End date", value=date(2026, 12, 31))

if report_end < report_start:
    st.error("End date must be after start date.")
    st.stop()


# ==========================================================
# Load data
# ==========================================================

if uploaded_file is not None:
    if uploaded_file.size > MAX_UPLOAD_BYTES:
        st.error(
            f"Uploaded file exceeds 5 MB limit ({uploaded_file.size / 1024 / 1024:.1f} MB). "
            "Please upload a smaller file."
        )
        st.stop()

    try:
        raw_grants = pd.read_excel(uploaded_file, engine="openpyxl")
        st.success("Uploaded grant file loaded.")
    except Exception:
        st.error(
            "Could not read the uploaded file. "
            "Check that it is a valid .xlsx file and matches the upload template column structure. "
            "Download the template from the sidebar if needed."
        )
        st.stop()
elif use_sample_data:
    raw_grants = SAMPLE_GRANTS.copy()
else:
    st.info("Upload an Excel file or enable sample data.")
    st.stop()

validation_errors = validate_grants(raw_grants)
if validation_errors:
    st.error("Upload validation failed.")
    for error in validation_errors:
        st.write(f"- {error}")
    st.stop()

grants = normalize_grants(raw_grants)


# ==========================================================
# UI sections
# ==========================================================

st.subheader("1. Grant data")
grants_display = grants.copy()
for col in ["grant_date", "vest_start_date", "vest_end_date"]:
    if col in grants_display.columns:
        grants_display[col] = grants_display[col].dt.date

st.dataframe(grants_display, use_container_width=True)

st.subheader("2. Forfeiture events")
st.caption("Add or edit forfeiture events manually. Forfeiture stops future recognition for that grant.")

forfeitures_input = st.data_editor(
    SAMPLE_FORFEITURES.copy() if use_sample_data else pd.DataFrame(columns=SAMPLE_FORFEITURES.columns),
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "grant_id": st.column_config.SelectboxColumn("grant_id", options=sorted(grants["grant_id"].astype(str).unique())),
        "forfeiture_date": st.column_config.DateColumn("forfeiture_date"),
        "reason": st.column_config.TextColumn("reason"),
    },
)

forfeitures = forfeitures_input.copy()
if not forfeitures.empty:
    forfeitures["forfeiture_date"] = forfeitures["forfeiture_date"].apply(parse_date)

forfeiture_errors = validate_forfeitures(forfeitures, grants)
if forfeiture_errors:
    st.error("Forfeiture validation failed.")
    for err in forfeiture_errors:
        st.write(f"- {err}")
    st.stop()



# ==========================================================
# Calculate
# ==========================================================

valued_grants = calculate_fair_values(
    grants=grants,
    risk_free_rate=risk_free_rate,
    volatility=volatility,
    expected_term_years=expected_term_years,
    dividend_yield=dividend_yield,
)

summary, detail = build_full_schedule(
    valued_grants=valued_grants,
    forfeitures=forfeitures,
    start_date=report_start,
    end_date=report_end,
    probable_threshold=probable_threshold,
)

if summary.empty:
    st.info(
        "No expense recognized in the selected reporting window. "
        "Check that the reporting window overlaps with at least one grant vesting period."
    )
    st.stop()

assumptions = pd.DataFrame(
    [
        {"assumption": "risk_free_rate", "value": risk_free_rate},
        {"assumption": "volatility", "value": volatility},
        {"assumption": "expected_term_years", "value": expected_term_years},
        {"assumption": "dividend_yield", "value": dividend_yield},
        {"assumption": "probable_threshold", "value": probable_threshold},
        {"assumption": "vesting_method", "value": "straight-line"},
        {"assumption": "day_count_convention", "value": "Actual/Actual (calendar days)"},
        {"assumption": "period_end_date_treatment", "value": "Inclusive - expense accrues through vest_end_date"},
        {"assumption": "reporting_window_method", "value": "compute from earliest vesting start, then filter display window"},
        {"assumption": "performance_award_method", "value": "recognize only when probable"},
        {"assumption": "forfeiture_method", "value": "manual grant-level forfeiture date"},
        {"assumption": "assumption_source_note", "value": sanitize_for_export(assumption_source_note)},
    ]
)

valued_grants_display = format_currency_columns(valued_grants)
summary_display = format_currency_columns(summary)
detail_display = format_currency_columns(detail)


# ==========================================================
# Output
# ==========================================================

if audit_awareness_mode:
    st.warning("Audit Awareness Mode: This MVP does not support " + "; ".join(KNOWN_LIMITATIONS) + ".")

st.subheader("3. Grant-date fair values")
st.warning(
    "MVP assumption warning: option valuations use user-entered placeholder assumptions. "
    "Before audit or board use, support risk-free rate, volatility, expected term, and dividend yield with documented sources."
)
st.dataframe(valued_grants_display, use_container_width=True)

metric_cols = st.columns(4)
metric_cols[0].metric("Total grant-date fair value", f"${valued_grants['total_fair_value'].sum():,.0f}")
metric_cols[1].metric("Total window net expense", f"${summary['current_month_expense'].sum():,.0f}" if not summary.empty else "$0")
metric_cols[2].metric("Total grants", f"{len(valued_grants):,.0f}")
metric_cols[3].metric("Forfeiture events", f"{len(forfeitures.dropna(subset=['grant_id'])) if not forfeitures.empty else 0:,.0f}")

st.subheader("4. Monthly waterfall")
st.caption("Beginning unrecognized comp, gross expense, forfeiture adjustment, current expense, and ending unrecognized comp.")
st.dataframe(summary_display, use_container_width=True)

st.subheader("5. Grant-level detail")
with st.expander("Show grant-level monthly detail", expanded=False):
    st.dataframe(detail_display, use_container_width=True)

st.subheader("6. Exports")

summary_export = sanitize_dataframe_for_export(summary_display)
detail_export = sanitize_dataframe_for_export(detail_display)
grants_export = sanitize_dataframe_for_export(valued_grants_display)
forfeitures_export = sanitize_dataframe_for_export(forfeitures)

assumption_header = assumptions.set_index("assumption")["value"].to_frame().T
assumption_csv = assumption_header.to_csv(index=False)
waterfall_csv = summary_export.to_csv(index=False)
csv_bytes = (
    "# ASC 718 MVP - Valuation Assumptions\n"
    + assumption_csv
    + "\n# Monthly Waterfall\n"
    + waterfall_csv
).encode("utf-8")

excel_bytes = to_excel_bytes(
    summary=summary_export,
    detail=detail_export,
    grants=grants_export,
    forfeitures=forfeitures_export,
    assumptions=assumptions,
    audit_awareness_mode=audit_awareness_mode,
)

export_cols = st.columns(2)
with export_cols[0]:
    st.download_button(
        label="Download CSV summary",
        data=csv_bytes,
        file_name="asc718_monthly_waterfall.csv",
        mime="text/csv",
    )

with export_cols[1]:
    st.download_button(
        label="Download Excel workbook",
        data=excel_bytes,
        file_name="asc718_stock_comp_mvp.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ==========================================================
# Help
# ==========================================================

with st.expander("Upload template requirements"):
    st.write("Your Excel file should include these columns:")
    st.code(", ".join(REQUIRED_COLUMNS))
    st.write("Allowed award_type values: option, rsu, performance")
    st.write("Optional option assumption override columns are supported: risk_free_rate_override, volatility_override, expected_term_override, dividend_yield_override.")
    st.write("Performance probability should be a decimal, e.g. 0.65 for 65%.")
    st.markdown(
        """
Validation rules:
- grant_id must be unique.
- shares and grant_date_fmv must be greater than zero.
- option strike_price must be greater than zero.
- RSU and performance strike_price can be zero, but not negative.
- vest_end_date must be after vest_start_date.
- performance_probability must be between 0 and 1.
- Dates should use YYYY-MM-DD format.
        """
    )

with st.expander("Accounting logic notes"):
    st.markdown(
        """
- **Options:** Grant-date fair value is calculated using Black-Scholes.
- **Black-Scholes assumptions:** Risk-free rate, volatility, expected term, and dividend yield are exposed and exported for review.
- **RSUs:** Grant-date fair value equals grant-date FMV multiplied by shares.
- **Performance awards:** Grant-date FMV is measured at full FMV, but expense is recognized only when the probability input meets the probable threshold.
- **Straight-line vesting:** Expense is recognized evenly over the requisite service period using Actual/Actual calendar days, with expense accruing through the vest_end_date.
- **Forfeitures:** Manual forfeiture date stops future expense recognition and records a one-time forfeiture adjustment in the forfeiture month.
- **Forfeiture method:** When a grant is forfeited, the requisite service period is truncated at the forfeiture date. Expense accrues only through that date. No reversal of prior-period expense is required under this approach. The column `early_termination_reversal` will show $0 unless cumulative recognized expense exceeded the cost earned through the forfeiture date, which is rare under daily straight-line accrual.
- **Custom reporting window:** The app computes from the earliest vesting start date first, then filters to the selected window so opening balances are not understated.
        """
    )
