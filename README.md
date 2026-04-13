# Japanese Equity Research — Automated DCF / SOTP Pipeline

Python automation pipeline for generating institutional-quality equity research models for Japanese listed companies. All primary data is sourced from original Japanese-language regulatory documents (有価証券報告書, 決算短信) via EDINET XBRL.

## Features
- **DCF Model**: 5-scenario (Bull/Upside/Base/Downside1/Downside2) with PGM + Exit Multiple
- **SOTP Model**: Segment-level EV/EBITDA valuation with Peer Comps
- **NWC**: 7-item itemized buildup method (not percentage-of-revenue)
- **Comps**: Automated peer comparable analysis
- **Sensitivity**: WACC × Terminal Growth + WACC × Exit Multiple matrices
- **Cross-validation**: Automated shares outstanding consistency check between DCF and SOTP

## Current Coverage

| Ticker | Company | SOTP Fair Value | Rating | Report |
|--------|---------|----------------|--------|--------|
| 6365.T | DMW Corporation (電業社機械製作所) | ¥9,366 | BUY | — |
| 2359.T | Core Corporation | ¥3,386 | BUY | [PDF](reports/Core_2359_Equity_Research_v2.pdf) |

## Tech Stack
- Python 3.x (openpyxl, yfinance, requests)
- Claude Code for pipeline orchestration
- Data sources: EDINET, IR Bank, Yahoo Finance Japan

## Pipeline

```
EDINET XBRL → edinet_fetcher.py → edinet_parser.py → generate_dcf.py → 8-sheet Excel
```

**Input:** One JSON config file per company (overrides + scenarios)
**Output:** Dynamic Excel model with 500+ live formulas, zero hardcoded values

### Generated Excel Structure

| # | Sheet | Content |
|---|-------|---------|
| 1 | Executive Summary | Investment thesis, target price, BUY/HOLD/SELL |
| 2 | Financial Statements | 6-year historical IS / CF / BS |
| 3 | DCF Model | WACC, 5-scenario FCF projection, PGM + Exit valuation |
| 4 | NWC Schedule | Days method or Revenue % method |
| 5 | Comps Analysis | Peer comparison with implied valuation |
| 6 | Sensitivity Analysis | WACC × TG and WACC × Exit Multiple tables |
| 7 | Segment Analysis | Revenue / OP / OPM by segment |
| 8 | Driver Analysis | Per-segment revenue decomposition |

### Revenue Driver Types

- `backlog` — orders → backlog → revenue (equipment manufacturers)
- `manmonth` — headcount × utilization × rate (IT services)
- `growth_rate` — YoY growth driven (general purpose)
- `manual` — direct revenue input

## Repository Structure

```
├── reports/          # Published equity research PDFs
├── models/           # Generated Excel DCF models
├── scripts/          # Pipeline scripts (fetcher, parser, generator)
├── templates/        # Excel template engine
├── data/
│   ├── overrides/    # Per-company JSON configs
│   └── comps/        # Comparable company CSVs
└── notes/            # Weekly coverage notes
```

## Usage

```bash
# 1. Edit overrides JSON with assumptions
# data/overrides/{ticker}_overrides.json

# 2. Generate DCF model
python scripts/generate_dcf.py {ticker} --output-dir models --force

# 3. Generate SOTP model (if sotp section exists in overrides)
python scripts/generate_sotp.py {ticker}
```

## Key Design Decisions

- **D/E ratio must be set manually** — EDINET XBRL taxonomy uses a narrow definition of debt that can understate leverage by 10x (discovered on IHI 7013.T where EDINET D/E = 0.097 vs actual 0.98)
- **5 scenarios** — Upside / Base / Management / Downside 1 / Downside 2
- **Segment-level modeling** — each segment gets its own driver type and projection matrix
- **COGS% override** — when segments are present, COGS% is derived from segment OPM to maintain consistency

## About

**Ryosuke Sato** — Independent Japanese equity researcher. 8+ years in finance and business operations. Native Japanese speaker providing primary-source research that most foreign investors cannot access directly.

- LinkedIn: [linkedin.com/in/ryosukesato0369](https://linkedin.com/in/ryosukesato0369)
- Upwork: Available for commissioned research

*Disclaimer: The author may hold positions in covered securities. Reports are for informational purposes only and do not constitute investment advice.*
