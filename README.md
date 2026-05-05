# Japanese Equity Research — Python-Automated Financial Modeling

Automated equity research pipeline for Japanese listed companies. Built with Python + Claude Code.

Reads original Japanese filings directly (決算短信, 有価証券報告書, 適時開示) and generates institutional-grade financial models in ~30 minutes per company.

---

## Automation Pipeline

The core of this project is a Python pipeline that auto-generates DCF and SOTP Excel models from structured JSON inputs.

### Templates
| File | Description |
|---|---|
| `templates/dcf_comps_template.py` | DCF (PGM + Exit Multiple) + Comparable Company Analysis — generates 5-sheet Excel workbook |
| `templates/sotp_template.py` | Sum-of-the-Parts valuation — generates 6-sheet Excel workbook with cross-check to DCF |

### Key Features
- **DCF Model**: 5-year FCF projection, WACC (CAPM + size premium), terminal value (perpetuity growth + exit multiple), scenario matrix, sensitivity analysis, NWC schedule
- **SOTP Model**: Segment EBITDA buildup, peer comps per segment, conglomerate/liquidity discount, sensitivity tables, D&A allocation
- **Cross-Check**: SOTP automatically reads DCF Excel outputs for 5-method valuation comparison
- **Comps Analysis**: EV/EBITDA and PER with median-based implied valuation
- All calculations use **Excel formulas** (not hardcoded values) — fully auditable

### How It Works
```
data/overrides/{ticker}_overrides.json   ← Company-specific assumptions
data/comps/{ticker}_comps.csv            ← Comparable companies data
        ↓
scripts/generate_dcf.py                  ← Generates DCF + Comps Excel
scripts/generate_sotp.py                 ← Generates SOTP Excel (reads DCF output)
        ↓
models/{ticker}_DCF_Model_*.xlsx         ← Output: DCF workbook
models/{ticker}_SOTP_Model.xlsx          ← Output: SOTP workbook
```

---

## Coverage

### DMW Corporation (6365.T) — BUY | TP ¥8,576 (+48%)
Pump & blower manufacturer with proprietary DeROs® energy recovery device for desalination (99.7% efficiency, world-leading). Market prices it as a boring domestic pump maker — SOTP reveals hidden overseas/desalination value.

| Method | Fair Value | vs ¥5,800 |
|---|---|---|
| SOTP (Base) | ¥8,196 | +41% |
| DCF — Exit Multiple | ¥9,660 | +67% |
| Comps EV/EBITDA | ¥9,772 | +69% |
| Comps PER | ¥7,496 | +29% |
| DCF — PGM | ¥5,882 | +1% |
| **Weighted Target** | **¥8,576** | **+48%** |

**Catalysts**: Middle East desalination reconstruction (2026-28), India factory doubling (2027E), METI Energy Conservation Grand Prize (Jan 2026)

### Core Corporation (2359.T) — BUY (Catalyst-Driven) | TP ¥2,784 (+26%)
IT services + defense/space tech (GNSS, satellite systems). Market values it as a generic SIer — SOTP shows defense tech segment at near-zero implied value.

| Method | Fair Value | vs ¥2,218 |
|---|---|---|
| SOTP (Base) | ¥4,786 | +116% |
| DCF — Exit Multiple | ¥3,781 | +70% |
| Comps EV/EBITDA | ¥2,476 | +12% |
| DCF — PGM | ¥2,824 | +27% |
| Comps PER | ¥2,055 | -7% |

**Catalyst (delivered)**: Q4 FY3/26 results (4/28/2026) — OP ¥3.82B, +9% above revised MTP target. Next catalyst: 15th Mid-Term Plan announcement (FY3/27-29).

**MTP Track Record Analysis**

Core Corp delivers consistent ~35-39% operating profit growth across each three-year mid-term plan, while keeping revenue guidance conservative. The 14th MTP target was revised down on April 28, 2025 — a footnote that appears only in the Japanese MTP document, not in English IR materials.

| MTP | Period | OP Growth (3Y) | Outcome |
|---|---|---|---|
| 12th | FY3/18-FY3/20 | +38% | Delivered |
| 13th | FY3/21-FY3/23 | +35% | Delivered |
| 14th | FY3/24-FY3/26 | +39% | Delivered (revised target +9% beat) |

Pattern: lower the bar → clear it. The 15th MTP announcement is likely to follow the same template — a conservative headline, with delivery skewing higher by FY3/29.

2040 vision: ¥100B revenue (vs FY3/26 actual ¥26.5B), implying ¥50B by 2030. The 15th MTP is the bridge.

### Torishima Pump (6363.T) — HOLD | Fair Value ¥3,392 (+2%)
Pump manufacturer with desalination exposure. 5-method average converges near current price — limited upside.

---

## Repository Structure

```
├── data/
│   ├── overrides/       # Company-specific model assumptions (JSON)
│   └── comps/           # Comparable company data (CSV)
├── models/              # Generated Excel models (DCF, SOTP)
├── templates/           # Python templates for Excel generation
│   ├── dcf_comps_template.py
│   └── sotp_template.py
├── scripts/             # Generation & utility scripts
├── reports/             # PDF equity research reports
├── docs/                # Documentation & specs
└── notes/               # Analysis notes & memos
```

---

## Tech Stack

- **Python**: openpyxl, yfinance, pandas, numpy
- **Claude Code**: AI-assisted model generation and code automation
- **Data Sources**: EDINET, TDnet (決算短信), yfinance
- **Output**: Formula-driven Excel models (fully auditable, zero hardcoded values)

---

## About

Financial analyst based in Japan. Specializing in Japanese equity research for foreign investors, bridging the language and information gap between global capital and Japan's undercovered small/mid-cap market.

## Disclaimer

All analyses are for informational and demonstration purposes only. They do not constitute investment advice.
