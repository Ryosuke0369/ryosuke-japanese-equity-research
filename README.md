# Japanese Equity Research — Automated Financial Modeling Platform

Config-driven Python platform that generates professional Excel workbooks with live formulas for Japanese equity research. Each model reads primary Japanese filings (決算短信, 有価証券報告書, 適時開示) and produces institutional-grade output.

## Platform Architecture

```
overrides JSON → generate_dcf.py → Excel (8 sheets, 500+ formulas)
                      ↓
              EDINET API → Parse → Config → Template → .xlsx
```

### Automated Pipeline
1. **EDINET Fetcher** — Downloads 有価証券報告書 and 決算短信 via API
2. **Financial Parser** — Extracts IS/BS/CF, NWC components, segment data
3. **Config Builder** — Merges parsed data with manual overrides (JSON)
4. **Template Engine** — Generates Excel with cross-sheet references, scenario analysis, sensitivity tables

### Key Design Principles
- **Config-driven**: Change one JSON file → regenerate entire model for any company
- **Method flexibility**: Capex/D&A (revenue_pct or direct), NWC (days or revenue_pct)
- **Segment/Driver analysis**: Configurable per-segment revenue decomposition (backlog, manmonth, growth_rate, manual)
- **5-scenario framework**: Base/Upside/Management/Downside1/Downside2 with per-year assumptions

## Models

### DCF & Comparable Company Analysis (Template)
**Reusable template** — applied to multiple companies via overrides JSON

| Sheet | Content |
|-------|---------|
| Executive Summary | Investment thesis, target price, recommendation |
| Financial Statements | Historical IS/CF/BS |
| DCF Model | WACC, 5-scenario FCF, PGM + Exit valuation |
| NWC Schedule | Days method or Revenue% method |
| Comps Analysis | Peer comparison, implied valuation |
| Sensitivity Analysis | WACC×TG, WACC×Exit Multiple (7×7) |
| Segment Analysis | Revenue/OP/OPM by segment + reconciliation |
| Driver Analysis | Per-segment revenue decomposition by driver_type |

**Applied to:**
- **Core Corporation (2359.T)** — IT/Defense, 3 segments (manmonth + manual)
- **TechnoSmart (6246.T)** — Machinery, 5 product lines (backlog + manual)

### KFC Holdings Japan — LBO Analysis
Reconstructing Carlyle Group's 2024 take-private. 8-sheet model with 3-statement integration, debt schedule, returns analysis (2.4x MOIC / 19.2% IRR).

→ See [`examples/kfc-japan-lbo/`](./examples/kfc-japan-lbo/)

### Headwaters × BBD Initiative — M&A Accretion/Dilution
Stock-for-stock merger analysis. Pro forma EPS across 3 scenarios. Dilutive under all scenarios; ¥131mn+ synergies required.

→ See [`examples/headwaters-bbd-merger/`](./examples/headwaters-bbd-merger/)

## Quick Start

```bash
# Generate DCF model for a new company
python scripts/generate_dcf.py \
  --ticker 2359 \
  --overrides data/overrides/2359_overrides.json \
  --comps-csv data/comps/2359_comps.csv \
  --years 4 \
  --force

# Verify formulas
python scripts/recalc.py models/2359_DCF_Model.xlsx
```

## Tech Stack
- Python 3.13 + openpyxl (Excel generation with live formulas)
- EDINET API (有価証券報告書, 決算短信)
- yfinance (real-time price, beta, shares outstanding)
- Consistent styling: blue=input, black=formula, green=cross-sheet reference

## About
Ryosuke Sato — Freelance Financial Analyst (Japan)
- 8+ years in finance & business operations
- Native Japanese + English for cross-border equity research
- Reading primary Japanese filings for international investors

GitHub: [github.com/Ryosuke0369](https://github.com/Ryosuke0369)

## Disclaimer
All analyses are for informational and demonstration purposes only. They do not constitute investment advice.
