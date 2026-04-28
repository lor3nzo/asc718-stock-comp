# ASC 718 Stock Compensation MVP

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Streamlit](https://img.shields.io/badge/Built%20with-Streamlit-red)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Status](https://img.shields.io/badge/status-MVP-orange)
![Local First](https://img.shields.io/badge/local--first-no%20cloud-green)

**Version: V.005**

**Local-first ASC 718 modeling tool for startup CFOs, controllers, and builders.**  
Convert equity grant data into defensible expense schedules (options, RSUs, performance awards) with full transparency.

> Not audit-certified software. Designed for internal modeling and audit support.

---

## Why this exists

ASC 718 modeling is:
- Error-prone in spreadsheets
- Hard to audit
- Time-consuming to rebuild

This tool provides a **transparent, reproducible, and inspectable engine** for equity compensation accounting.

---

## Features

- Black-Scholes valuation (options)
- Grant-date FMV (RSUs + performance awards)
- Daily straight-line expense (Actual/Actual)
- Manual forfeiture handling
- Per-grant assumption overrides
- Monthly waterfall with full traceability
- CSV + Excel export (audit-friendly)
- Formula injection protection
- Local-first (no data leaves your machine)

---

## Quick Start (2 minutes)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then:

1. Enable **"Use sample dataset"**
2. Keep default assumptions
3. Set reporting window to 2024–2026

Expected:
- Monthly expense appears
- Total expense < total FV (due to performance awards)

---

## Example Output

| Month | Expense | Ending Unrecognized |
|------|--------|--------------------|
| Jan 2024 | 12,500 | 487,500 |
| Feb 2024 | 12,500 | 475,000 |

---

## Model Integrity

- Total expense ties to total fair value (absent forfeitures)
- Forfeitures truncate service period correctly
- No over-recognition beyond earned cost
- Deterministic outputs given same inputs

---

## Architecture

```
Input → Validation → Valuation → Expense Engine → Reporting → Export
```

---

## Limitations

- No graded vesting
- No tranche modeling
- No modifications / repricing
- No probability change tracking
- No integrations or persistence

---

## Contributing

Contributions are welcome.

Suggested areas:
- Performance optimization
- Additional vesting models (graded, tranches)
- Better visualization (charts, dashboards)
- Test coverage expansion
- UI/UX improvements

To contribute:
1. Fork the repo
2. Create a feature branch
3. Submit a PR with clear description

---

## Testing

```bash
pytest test_asc718.py
```

---

## Disclaimer

This is not audit-certified ASC 718 software.

- Assumptions must be externally supported
- Outputs should be reviewed before audit or board use

---

## License

MIT License
