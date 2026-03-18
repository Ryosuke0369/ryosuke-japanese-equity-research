# Japanese Equity Research — Automated Financial Modeling Platform

Config-driven Python platform that generates institutional-grade Excel workbooks for Japanese equity research. Reads primary Japanese filings (決算短信, 有価証券報告書, 適時開示) directly — no machine translation.

## How It Works

```
JSON config → generate_dcf.py → 8-sheet Excel (500+ formulas, zero errors)
```

**Pipeline:**
1. EDINET API fetches 有価証券報告書 and 決算短信
2. Parser extracts IS / BS / CF / NWC / Segment data
3. Config builder merges parsed data with manual overrides (JSON)
4. Template engine generates Excel with cross-sheet references and scenario analysis

## Generated Excel Structure

| # | Sheet | Content |
|---|-------|---------|
| 1 | Executive Summary | Investment thesis, target price, BUY/HOLD/SELL |
| 2 | Financial Statements | Historical IS / CF / BS |
| 3 | DCF Model | WACC, 5-scenario FCF projection, PGM + Exit valuation |
| 4 | NWC Schedule | Days method (DSO/DIH/DPO) or Revenue % method |
| 5 | Comps Analysis | Peer comparison with implied valuation |
| 6 | Sensitivity Analysis | WACC × TG and WACC × Exit Multiple (7×7 tables) |
| 7 | Segment Analysis | Revenue / OP / OPM by segment + reconciliation |
| 8 | Driver Analysis | Per-segment revenue decomposition |

## Key Design Features

**Config-driven flexibility** — change one JSON file, regenerate the entire model for any company:
- Capex & D&A: revenue % or direct input (company-dependent)
- NWC: DSO/DIH/DPO days method or simple revenue % method
- 5 scenario framework: Base / Upside / Management / Downside 1 / Downside 2

**Segment & Driver Analysis** — each segment gets its own revenue driver model:
- `backlog` — orders → backlog → revenue (equipment manufacturers)
- `manmonth` — headcount × utilization × rate (IT services)
- `growth_rate` — YoY growth driven (general purpose)
- `manual` — direct revenue input

## Repository Structure

```
├── scripts/                  # Pipeline scripts
│   ├── generate_dcf.py       # Main entry point
│   ├── edinet_fetcher.py     # EDINET API client
│   ├── edinet_parser.py      # Financial statement parser
│   ├── yfinance_quarterly.py # LTM data supplement
│   ├── comps_fetcher.py      # Comparable company data
│   └── recalc.py             # LibreOffice formula verification
│
├── templates/
│   └── dcf_comps_template.py # Template engine (2,300+ lines)
│
├── data/
│   ├── overrides/            # Per-company JSON configs
│   │   ├── 2359_overrides.json
│   │   └── 6246_overrides.json
│   └── comps/                # Comparable company CSVs
│
├── models/                   # Generated Excel outputs
└── CLAUDE.md                 # AI-assisted development config
```

## Applied To

**Core Corporation (2359.T)** — IT / Defense
- 3 segments: manmonth + manual driver types
- Direct Capex/D&A method (non-revenue-linked)
- NWC via DSO/DIH/DPO days method

**TechnoSmart (6246.T)** — Machinery / Equipment
- 5 product lines: backlog + manual driver types
- Revenue % Capex/D&A method
- NWC via revenue % method (equipment maker with large contract liabilities)

## Quick Start

```bash
# Generate DCF model
python scripts/generate_dcf.py \
  --ticker 2359 \
  --overrides data/overrides/2359_overrides.json \
  --comps-csv data/comps/2359_comps.csv \
  --years 4 --force

# Verify formulas (requires LibreOffice)
python scripts/recalc.py models/2359_DCF_Model.xlsx
```

## Tech Stack

- Python 3.13 + openpyxl (Excel generation with live formulas)
- EDINET API (有価証券報告書, 決算短信)
- yfinance (real-time price, beta, shares outstanding)
- Styling convention: blue = input, black = formula, green = cross-sheet reference

## About

**Ryosuke Sato** — Freelance Financial Analyst (Japan)
- 8+ years in finance & business operations
- Native Japanese speaker reading primary filings for international investors
- GitHub: [github.com/Ryosuke0369](https://github.com/Ryosuke0369)

## Disclaimer

All analyses are for informational and demonstration purposes only. They do not constitute investment advice.
