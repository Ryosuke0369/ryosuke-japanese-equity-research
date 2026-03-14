# Japanese Equity Research Automation

**Python x Excel — Fully Automated Equity Research Workflow for Japanese Stocks**

End-to-end automation that fetches financial data directly from Japan's regulatory API, intelligently fills data gaps with market sources, and generates institutional-grade Excel models with a single command. Built for analysts covering Japanese equities who need speed without sacrificing rigor.

---

## What's New — Phase 4: Intelligent WACC & Scenario-Driven Valuation

> **Professional-grade DCF inputs, automatically.** Phase 4 brings intelligent WACC calculation, multi-scenario analysis, and working capital modeling — eliminating hours of manual setup while keeping the analyst in control of forward assumptions.

### Intelligent WACC Auto-Calculation

- **Beta** — Fetched from yfinance with abnormal-value guard (clamped to 0.6–1.5 range; defaults to 1.0 fallback)
- **Size Premium** — Automatically determined by market cap: 0.0% (large-cap), 1.5% (mid-cap), 3.0% (small-cap)
- **D/E Ratio** — Dynamically computed from Net Debt / Market Cap (live market data)

### D&A / Capex Auto-Estimation

- Historical revenue ratios calculated from EDINET filings (average of available periods)
- Graceful fallback to conservative default ratios when EDINET data is unavailable

### 5-Scenario Matrix

```
Scenario Dropdown (Excel Data Validation)
┌─────────────────────────────────────────────┐
│  Base  │ Upside │ Management │ Down 1 │ Down 2 │
├─────────────────────────────────────────────┤
│  Revenue Growth, COGS%, SGA%, NWC days      │
│  — all independently configurable per       │
│    scenario via Excel dropdown              │
└─────────────────────────────────────────────┘
```

### NWC Schedule

- DSO / DIH / DPO-based working capital schedule integrated into the DCF model
- Change in NWC automatically flows into Free Cash Flow calculation

---

## What's New — Phase 3: Hybrid LTM with Adaptive Fallback

> **The 2024 regulatory wall, solved.** Japan's April 2024 金融商品取引法改正 (Financial Instruments and Exchange Act amendment) eliminated mandatory XBRL submissions for Q1/Q3 quarterly reports. This created a critical data gap — mid-cycle LTM calculations became impossible from EDINET alone.

### Our Solution: yfinance Adaptive Fallback

When EDINET lacks interim XBRL data, the system **automatically detects the gap** and falls back to yfinance quarterly financials to construct a **hybrid LTM** — combining the precision of EDINET annual data with the timeliness of market-sourced quarterly data.

```
EDINET (XBRL)          yfinance (quarterly)         Hybrid LTM
┌──────────────┐       ┌──────────────────┐       ┌──────────────────┐
│ FY annual    │       │ Q1, Q2, Q3, Q4   │       │ FY + Q_new       │
│ H1 semi-ann  │  ───► │ (auto-detected)  │  ───► │   − Q_old        │
│ (if available)│       │ Income/CF/BS     │       │ = Accurate LTM   │
└──────────────┘       └──────────────────┘       └──────────────────┘
```

- **Zero manual intervention** — Gap detection, data sourcing, and LTM recomputation happen automatically
- **Graceful degradation** — If yfinance is unavailable, the system proceeds with the best available EDINET data
- **Full coverage** — Income statement, cash flow, and balance sheet items are all enriched

---

## Key Features

| Feature | Description |
|---------|-------------|
| **EDINET API Data Fetching** | Fully automated 5-year financial data retrieval from the FSA's EDINET API (v2) — just provide a securities code |
| **XBRL Financial Parser** | Extracts all DCF-critical variables from `.xbrl` instance documents with 100% accuracy (JPY millions) |
| **Hybrid LTM Generation** | Automatically detects Q1/Q3 data gaps and constructs hybrid LTM by combining EDINET + yfinance quarterly data |
| **Adaptive Fallback** | Seamlessly switches between EDINET XBRL and yfinance data sources depending on availability |
| **Live Market Data** | Fetches real-time stock prices, shares outstanding, and market cap via yfinance API |
| **DCF + Comps Model** | Generates a 5-sheet Excel workbook with 150+ live formulas: Executive Summary, Financial Statements, DCF Valuation, Comparable Company Analysis, and Sensitivity Tables |
| **LBO Model** | Full 3-statement LBO model (8 sheets) with debt schedules, IRR/MOIC returns analysis |
| **M&A Accretion/Dilution** | Stock-for-stock merger analysis (7 sheets) with pro forma EPS impact across scenarios |
| **5-Scenario Analysis** | Base / Upside / Management / Downside 1 & 2 — switch via Excel dropdown. Revenue Growth, COGS%, SGA%, NWC days all independently configurable per scenario |
| **NWC Schedule** | DSO / DIH / DPO-based working capital schedule. Change in NWC automatically flows into DCF Free Cash Flow |
| **Sensitivity Analysis** | Auto-generated 2-axis sensitivity tables: WACC × Terminal Growth Rate and WACC × Exit Multiple |
| **Intelligent WACC** | Beta abnormal-value guard, market-cap-linked Size Premium, and dynamic D/E Ratio calculation for professional WACC estimation |
| **Comps CSV Workflow** | Drop a CSV into `data/comps/` and comparable company analysis is automatically integrated into the model |

---

## Design Philosophy: Why Not Fully Automate Projections?

> **"Automation handles the past. The analyst owns the future."**

This tool deliberately automates **only historical data collection and LTM computation**. Forward-looking assumptions — revenue growth rates, COGS trajectories, margin expansion, capex intensity — are left as **manual inputs in the Excel model**.

Why? Because **building the projection is the job**. Equity research exists to form a differentiated view on a company's future cash flows. That view is what drives the target price, the BUY/SELL recommendation, and ultimately the investment decision. Automating projections would strip away the very thing that makes research valuable.

The generated Excel model provides:
- Historical data pre-populated with 100% accuracy (no copy-paste errors)
- A clean projection framework with formulas ready for your assumptions
- Sensitivity tables that instantly reflect your scenario changes

**You bring the thesis. The tool brings the infrastructure.**

However, WACC components (Beta, Size Premium, D/E Ratio) and D&A/Capex ratios — parameters that can be objectively derived from historical and market data — are auto-calculated. We draw a clear boundary between subjective forward projections and objective market-observable inputs.

---

## Quick Start

```bash
# 1. Clone the repository
git clone https://github.com/Ryosuke0369/ryosuke-japanese-equity-research.git
cd ryosuke-japanese-equity-research

# 2. Install dependencies
pip install openpyxl yfinance requests beautifulsoup4 lxml python-dotenv pandas

# 3. Set your EDINET API key
#    Get a free key at: https://disclosure2dl.edinet-fsa.go.jp/guide/static/register
echo "EDINET_API_KEY=your-subscription-key-here" > .env

# 4. Generate a complete DCF model (one command)
python scripts/generate_dcf.py 2359          # Core Corporation (3月決算)
python scripts/generate_dcf.py 2359 --years 3  # 3 years only
```

### Adding Comparable Companies

1. Create a CSV in the `data/comps/` directory (e.g., `7974_comps.csv`)
2. CSV format: `Ticker,Company Name` (e.g., `6758.T,Sony Group`)
3. Loaded automatically via `--comps-csv` option or auto-detection by securities code

### What Happens Under the Hood

```
Securities Code (e.g. 2359)
        │
        ▼
┌─────────────────────────┐
│   edinet_fetcher.py     │  ← EDINET API v2 (documents.json)
│   • Adaptive Search     │     Finds 5 annual + 1 interim report
│   • ZIP download        │     Downloads & extracts XBRL files
└────────┬────────────────┘
         │ .xbrl files
         ▼
┌─────────────────────────┐
│   edinet_parser.py      │  ← BeautifulSoup + lxml
│   • Context mapping     │     Maps XBRL contexts to fiscal periods
│   • Data extraction     │     Extracts all financial line items
│   • LTM calculation     │     FY + H1_current − H1_prior = LTM
└────────┬────────────────┘
         │ Structured data (OrderedDict)
         ▼
┌─────────────────────────┐
│  yfinance_quarterly.py  │  ← Adaptive Fallback (Phase 3)
│  • Gap detection        │     Detects missing Q1/Q3 data
│  • Quarterly fetch      │     Pulls IS/CF/BS from yfinance
│  • Hybrid LTM           │     Merges EDINET + yfinance → LTM
└────────┬────────────────┘
         │ Enriched data
         ▼
┌─────────────────────────┐
│   generate_dcf.py       │  ← One-click orchestrator
│   • Config builder      │     Converts data → model parameters
│   • Comps integration   │     Fetches peer group multiples
│   • Excel generation    │     5-sheet workbook, 150+ formulas
└────────┬────────────────┘
         │
         ▼
    output/2359_DCF_Model_YYYYMMDD.xlsx
```

### Sample Output

```
  Item                                      LTM(Q3 2025-12)    FY2025    FY2024    FY2023    FY2022    FY2021
  Revenue (売上高)                                   25,800    24,599    23,999    22,848    21,798    20,785
  Operating Income (営業利益)                          3,650     3,175     3,141     2,744     2,368     2,032
  Net Income (当期純利益)                               2,550     2,242     2,271     1,968     1,623     1,423
  Operating Cash Flow (営業CF)                       2,900     2,373     2,190     1,944     1,799     1,851
  Net Debt (ネットデット)                               -7,500    -6,174    -4,565    -3,775    -2,737    -1,526
```

> All values in JPY millions (百万円). LTM is automatically computed via EDINET XBRL or yfinance hybrid fallback.

---

## Repository Structure

```
ryosuke-japanese-equity-research/
├── scripts/                               # Core automation engine
│   ├── edinet_fetcher.py                  #   EDINET API client — fetches & downloads XBRL
│   ├── edinet_parser.py                   #   XBRL parser — extracts financials & computes LTM
│   ├── yfinance_quarterly.py              #   Hybrid LTM fallback via yfinance (Phase 3)
│   ├── generate_dcf.py                    #   One-click DCF model generator
│   ├── comps_fetcher.py                   #   Comparable company data fetcher
│   ├── recalc.py                          #   Excel formula recalculation utility
│   └── pdf_parser.py                      #   Legacy PDF financial data extractor
│
├── templates/                             # Reusable Excel model generators
│   ├── dcf_comps_template.py              #   DCF + Comparable Company Analysis
│   ├── lbo_template.py                    #   Leveraged Buyout Analysis
│   ├── ma_accretion_template.py           #   M&A Accretion / Dilution Analysis
│   └── comps_input_template.csv           #   Peer company input template
│
├── data/comps/                            # Input data (peer company CSVs)
├── output/                                # Generated Excel models
├── tmp/edinet_data/                       # Auto-cleaned download cache (gitignored)
│
├── examples/                              # Completed case studies
│   ├── core-corporation-2359/             #   DCF/Comps — GIS & Defense IT
│   ├── kudan-4425/                        #   DCF/Comps — Deep Learning SLAM
│   ├── kfc-japan-lbo/                     #   LBO — Carlyle Take-Private
│   └── headwaters-bbd-merger/             #   M&A — Stock-for-Stock Merger
│
├── .env                                   # API keys (gitignored)
└── README.md
```

---

## Completed Case Studies

### DCF & Comparable Company Analysis

**Core Corporation (2359.T)** — GIS/Defense IT Services
- Target price: JPY 2,734 (BUY, +22% upside)
- 152 Excel formulas with cross-sheet references and BUY/HOLD/SELL recommendation logic

**Kudan (4425.T)** — Deep Learning Visual SLAM
- PDF auto-parsing from 3 years of annual reports (TDnet filings)
- Live yfinance integration for real-time equity valuation

**Nintendo (7974.T)** — Global Gaming & Entertainment
- Switch 2 launch cycle analysis with 5-scenario framework
- Comps: Sony, Capcom, EA, Take-Two
- WACC auto-calculated: Beta guard (1.0 fallback), Size Premium 0.0% (large-cap)
- Sensitivity Analysis: WACC vs Terminal Growth / Exit Multiple

### LBO Analysis

**KFC Holdings Japan (9873)** — Carlyle Take-Private (May 2024)
- Full 3-statement model with BS balance check and CF consistency check
- Returns: 2.4x MOIC / 19.2% IRR (15x exit, Year 5)

### M&A Accretion/Dilution

**Headwaters (4011.T) x BBD Initiative (5259.T)** — Stock-for-Stock Merger
- Pro forma EPS impact across Base / Downside / Synergy scenarios

---

## Roadmap

- [x] **Phase 1** — Template-based Excel model generation (DCF, LBO, M&A)
- [x] **Phase 2** — EDINET API integration for automated XBRL data extraction + LTM
- [x] **Phase 3** — Hybrid LTM generation with yfinance adaptive fallback (Q1/Q3 data gap solution)
- [x] **Phase 4** — Intelligent WACC, 5-scenario analysis, NWC schedule, sensitivity tables, comps CSV workflow
- [ ] **Phase 5** — EDINET XBRL tag expansion (non-standard Capex/D&A extraction), company guidance auto-integration

---

## Tech Stack

- **Python 3** — Core automation engine
- **EDINET API v2** — Official FSA disclosure API for XBRL financial filings
- **yfinance** — Quarterly financial data & real-time market data (adaptive fallback source)
- **BeautifulSoup + lxml** — XBRL/XML parsing
- **openpyxl** — Excel generation with live formulas (no hardcoded values)
- **pandas** — Comps CSV reading and data manipulation
- **python-dotenv** — Secure API key management

---

## About the Author

**Ryosuke Sato** — Financial Analyst & Python Developer

- 8+ years in finance and business operations
- Specializing in Japanese equity research automation and financial modeling
- Python-driven workflow optimization: transforming manual Excel processes into reproducible, auditable pipelines
- Native Japanese + English — bridging Japanese market data for international investors

[GitHub](https://github.com/Ryosuke0369)
