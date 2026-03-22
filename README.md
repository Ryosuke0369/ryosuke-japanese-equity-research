# Japanese Equity Research — Automated Financial Modeling Platform

Config-driven Python platform that generates institutional-grade Excel workbooks for Japanese equity research. Reads primary Japanese filings (決算短信, 有価証券報告書, 適時開示) directly — no machine translation.

## How It Works

```
JSON config → generate_dcf.py → 8-sheet Excel (500+ formulas, segment-linked DCF)
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
| 3 | DCF Model | Revenue/EBIT auto-linked from Segment Analysis, 5-scenario valuation |
| 4 | NWC Schedule | Days method (DSO/DIH/DPO) or Revenue % method |
| 5 | Comps Analysis | Peer comparison with implied valuation |
| 6 | Sensitivity Analysis | WACC × TG and WACC × Exit Multiple (7×7 tables) |
| 7 | Segment Analysis | Bottom-up forecast engine: Revenue Growth% + OPM per segment × 5 scenarios |
| 8 | Driver Analysis | Per-segment revenue decomposition |

## Key Design Features

**Bottom-up segment-linked DCF** — all forecast inputs live in Segment Analysis:
- Per-segment Revenue Growth (YoY%) and Operating Margin across 5 scenarios
- Segment totals flow directly to DCF Model via cell references
- COGS% is back-calculated from segment mix — not a manual input
- Consolidated SGA% and NWC% input at group level
- Switch one dropdown → every segment, Revenue, EBIT, FCF, and valuation update instantly

**Data flow:**
```
Segment Input Matrix (Growth% + OPM × 5 scenarios)
  → CHOOSE(scenario) → Segment Revenue / OP
    → Total Revenue / Total OP
      → DCF Model (Revenue, EBIT, FCF, Valuation)
        → Sensitivity Analysis (auto-updated)
```

**Config-driven flexibility** — change one JSON file, regenerate the entire model for any company:
- Capex & D&A: revenue % or direct input (company-dependent)
- NWC: DSO/DIH/DPO days method or simple revenue % method
- 5 scenario framework: Base / Upside / Management / Downside 1 / Downside 2

**Segment driver types** — each segment gets its own revenue driver model:
- `backlog` — orders → backlog → revenue (equipment manufacturers)
- `manmonth` — headcount × utilization × rate (IT services)
- `growth_rate` — YoY growth driven (general purpose)
- `manual` — direct revenue input

**Overrides JSON format** — per-company configuration:
- `segments[].projections.revenue_growth`: YoY% array (Base case)
- `segments[].projections.op_margin`: OPM% array (Base case)
- `segments[].scenario_projections`: other 4 scenarios (optional, falls back to Base)
- `scenarios.sga_pct` / `scenarios.nwc_pct`: consolidated-level inputs per scenario
- Backward compatible: tickers without `segments` use legacy top-down growth rate method

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
├── reports/                  # Finalized equity research outputs
│   └── TechnoSmart_6246_Equity_Research.pdf
└── CLAUDE.md                 # AI-assisted development config
```

## Applied To

**Core Corporation (2359.T)** — IT / Defense
- 3 segments: manmonth + manual driver types
- Direct Capex/D&A method (non-revenue-linked)
- NWC via DSO/DIH/DPO days method

**TechnoSmart (6246.T)** — Machinery / Equipment
- 5 product lines: backlog + manual driver types
- Segment-linked DCF: per-segment Revenue Growth% + OPM × 5 scenarios
- NWC via revenue % method (equipment maker with large contract liabilities)
- Full equity research PDF report available (7 pages, English)

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
