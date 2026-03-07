# Japanese Equity Research Automation

**Python x Excel — Fully Automated Equity Research Workflow for Japanese Stocks**

End-to-end automation that fetches financial data directly from Japan's regulatory API and generates institutional-grade Excel models with a single command. Built for analysts covering Japanese equities who need speed without sacrificing rigor.

---

## What's New — Phase 2: EDINET API Integration

> **No more PDFs.** Financial data is now pulled directly from the Financial Services Agency's EDINET API (v2), parsed from official XBRL filings, and aggregated into a multi-year matrix — fully automated.

### Highlights

- **One command, 5 years of data** — Enter a securities code, get 5 years of annual reports + the latest semi-annual report, automatically downloaded, extracted, and parsed
- **XBRL precision** — Revenue, operating income, net income, D&A, working capital (AR/AP/inventory), debt, and cash flow extracted at 100% accuracy (JPY millions) directly from regulatory XBRL instance documents
- **LTM auto-calculation** — Last Twelve Months financials computed automatically by combining the latest interim period with the most recent full-year data
- **Post-2024 reform support** — Handles both the legacy quarterly system (docType 140) and the new semi-annual system (docType 160) introduced by the April 2024 金融商品取引法改正
- **Adaptive Search** — Predicts filing dates from fiscal year-end and uses spiral search (±20 days) to find documents in ~3 API calls instead of scanning hundreds of dates

---

## Key Features

| Feature | Description |
|---------|-------------|
| **EDINET API Data Fetching** | Fully automated 5-year financial data retrieval from the FSA's EDINET API (v2) — just provide a securities code |
| **XBRL Financial Parser** | Extracts all DCF-critical variables from `.xbrl` instance documents with 100% accuracy (JPY millions) |
| **LTM Calculation** | Automatically computes Last Twelve Months financials from the latest interim + annual data |
| **Live Market Data** | Fetches real-time stock prices, shares outstanding, and market cap via yfinance API |
| **DCF + Comps Model** | Generates a 5-sheet Excel workbook with 150+ live formulas: Executive Summary, Financial Statements, DCF Valuation, Comparable Company Analysis, and Sensitivity Tables |
| **LBO Model** | Full 3-statement LBO model (8 sheets) with debt schedules, IRR/MOIC returns analysis |
| **M&A Accretion/Dilution** | Stock-for-stock merger analysis (7 sheets) with pro forma EPS impact across scenarios |

---

## Quick Start — EDINET Financial Data Extraction

```bash
# 1. Clone the repository
git clone https://github.com/Ryosuke0369/ryosuke-japanese-equity-research.git
cd ryosuke-japanese-equity-research

# 2. Install dependencies
pip install openpyxl yfinance requests beautifulsoup4 lxml python-dotenv

# 3. Set your EDINET API key
#    Get a free key at: https://disclosure2dl.edinet-fsa.go.jp/guide/static/register
echo "EDINET_API_KEY=your-subscription-key-here" > .env

# 4. Run — fetches 5 years of annual data + LTM for any listed company
python scripts/edinet_fetcher.py 2359          # Core Corporation (3月決算)
python scripts/edinet_fetcher.py 2359 --years 3  # 3 years only
```

### Sample Output

```
  Item                                      LTM(2Q 2025-09)    FY2025    FY2024    FY2023    FY2022    FY2021
  Revenue (売上高)                                   25,117    24,599    23,999    22,848    21,798    20,785
  Operating Income (営業利益)                          3,483     3,175     3,141     2,744     2,368     2,032
  Net Income (当期純利益)                               2,426     2,242     2,271     1,968     1,623     1,423
  Operating Cash Flow (営業CF)                       2,749     2,373     2,190     1,944     1,799     1,851
  Net Debt (ネットデット)                               -7,297    -6,174    -4,565    -3,775    -2,737    -1,526
```

> All values in JPY millions (百万円). LTM is computed from the latest H1 semi-annual + full-year annual data.

---

## Repository Structure

```
ryosuke-japanese-equity-research/
├── scripts/                               # Core automation engine
│   ├── edinet_fetcher.py                  #   EDINET API client — fetches & downloads XBRL
│   ├── edinet_parser.py                   #   XBRL parser — extracts financials & computes LTM
│   ├── comps_fetcher.py                   #   Comparable company data fetcher
│   └── pdf_parser.py                      #   Legacy PDF financial data extractor
│
├── templates/                             # Reusable Excel model generators
│   ├── dcf_comps_template.py              #   DCF + Comparable Company Analysis
│   ├── lbo_template.py                    #   Leveraged Buyout Analysis
│   └── ma_accretion_template.py           #   M&A Accretion / Dilution Analysis
│
├── examples/                              # Completed case studies
│   ├── core-corporation-2359/             #   DCF/Comps — GIS & Defense IT
│   ├── kudan-4425/                        #   DCF/Comps — Deep Learning SLAM
│   ├── kfc-japan-lbo/                     #   LBO — Carlyle Take-Private
│   └── headwaters-bbd-merger/             #   M&A — Stock-for-Stock Merger
│
└── tmp/edinet_data/                       # Auto-cleaned download cache (gitignored)
```

---

## How It Works

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
│   dcf_comps_template.py │  ← openpyxl
│   • 5-sheet Excel model │     DCF, Comps, Sensitivity, etc.
│   • 150+ live formulas  │     No hardcoded values
└─────────────────────────┘
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
- [ ] **Phase 3** — End-to-end pipeline: securities code → finished Excel model (zero manual input)
- [ ] **Phase 4** — Multi-company batch processing and automated comps table generation

---

## Tech Stack

- **Python 3** — Core automation engine
- **EDINET API v2** — Official FSA disclosure API for XBRL financial filings
- **BeautifulSoup + lxml** — XBRL/XML parsing
- **openpyxl** — Excel generation with live formulas (no hardcoded values)
- **yfinance** — Real-time stock price and market data
- **python-dotenv** — Secure API key management

---

## About the Author

**Ryosuke Sato** — Financial Analyst & Python Developer

- 8+ years in finance and business operations
- Specializing in Japanese equity research automation and financial modeling
- Python-driven workflow optimization: transforming manual Excel processes into reproducible, auditable pipelines
- Native Japanese + English — bridging Japanese market data for international investors

[GitHub](https://github.com/Ryosuke0369)
