# Japanese Equity Research — Ryosuke Sato

Independent equity research on Japanese-listed companies, targeting foreign institutional investors who lack access to Japanese-language securities filings.

## What This Is

A Python-automated equity research pipeline that generates institutional-quality DCF and comparable company analysis from EDINET XBRL filings. All primary data is sourced from original Japanese-language regulatory documents (有価証券報告書, 決算短信).

## Active Coverage

| Company | Ticker | Recommendation | Target Price | Report |
|---------|--------|---------------|-------------|--------|
| **Core Corporation** | 2359.T | **BUY** | ¥3,386 (+46.9%) | [PDF](reports/Core_2359_Equity_Research_v2.pdf) |

**Upcoming catalyst:** April 28, 2026 — FY2026/3 full-year earnings. Our estimate: OP ¥4,050M vs guidance ¥3,500M (+15.7% beat). EPS ¥201 vs consensus ~¥174.

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

## Quick Start

```bash
# 1. Fetch EDINET XBRL data
python scripts/edinet_fetcher.py 2359

# 2. Place comps CSV
# data/comps/2359.T_comps_.csv

# 3. Create overrides JSON
# data/overrides/2359_overrides.json
# IMPORTANT: Always set de_ratio, capex_pct, da_pct manually

# 4. Generate Excel model
python scripts/generate_dcf.py 2359
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
