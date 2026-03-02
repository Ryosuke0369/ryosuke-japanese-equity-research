# Japanese Equity Research Automation

**Python x Excel — Fully Automated Equity Research Workflow for Japanese Stocks**

End-to-end automation that transforms raw financial PDFs into institutional-grade Excel models with a single command. Built for analysts covering Japanese equities who need speed without sacrificing rigor.

---

## Key Features

| Feature | Description |
|---------|-------------|
| **PDF Financial Scraping** | Automatically extracts revenue, operating income, COGS, SGA, cash flow, and balance sheet items from Japanese corporate filings (TDnet format) using spatial PDF parsing |
| **Live Market Data** | Fetches real-time stock prices, shares outstanding, and market cap via yfinance API |
| **DCF + Comps Model** | Generates a 5-sheet Excel workbook with 150+ live formulas: Executive Summary, Financial Statements, DCF Valuation, Comparable Company Analysis, and Sensitivity Tables |
| **LBO Model** | Full 3-statement LBO model (8 sheets) with debt schedules, IRR/MOIC returns analysis, and balance sheet integrity checks |
| **M&A Accretion/Dilution** | Stock-for-stock merger analysis (7 sheets) with pro forma EPS impact across multiple scenarios |
| **One-Click Generation** | Config-driven architecture — edit a Python dict, run the script, get a complete Excel model |

---

## Repository Structure

```
ryosuke-japanese-equity-research/
├── templates/                          # Reusable model generators (start here!)
│   ├── dcf_comps_template.py           #   DCF + Comparable Company Analysis
│   ├── lbo_template.py                 #   Leveraged Buyout Analysis
│   └── ma_accretion_template.py        #   M&A Accretion / Dilution Analysis
│
├── scripts/                            # Shared utilities
│   ├── pdf_parser.py                   #   PDF financial data extractor
│   └── recalc.py                       #   Excel formula verifier (LibreOffice)
│
└── examples/                           # Completed case studies
    ├── core-corporation-2359/          #   DCF/Comps — GIS & Defense IT
    ├── kudan-4425/                     #   DCF/Comps — Deep Learning SLAM
    ├── kfc-japan-lbo/                  #   LBO — Carlyle Take-Private
    └── headwaters-bbd-merger/          #   M&A — Stock-for-Stock Merger
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

## Quick Start — Generate a DCF Model for Any Company

```bash
# 1. Clone the repository
git clone https://github.com/Ryosuke0369/ryosuke-japanese-equity-research.git
cd ryosuke-japanese-equity-research

# 2. Install dependencies
pip install openpyxl yfinance PyMuPDF

# 3. Create a folder for the new company and add PDF reports
mkdir examples/my-company-1234
# Place TDnet annual report PDFs into the folder

# 4. Copy the template
cp templates/dcf_comps_template.py examples/my-company-1234/

# 5. Edit the config dict in the template
#    - Set ticker, company name, fiscal years
#    - Adjust growth assumptions and WACC parameters
#    - Add comparable companies with their multiples

# 6. Run the script
python examples/my-company-1234/dcf_comps_template.py

# 7. Verify formulas (requires LibreOffice)
python scripts/recalc.py examples/my-company-1234/1234_Equity_Research_V3.xlsx

# 8. Open in Excel — all formulas auto-calculate
```

---

## Tech Stack

- **Python 3** — Core automation engine
- **openpyxl** — Excel generation with live formulas (no hardcoded values)
- **PyMuPDF (fitz)** — Spatial PDF parsing for Japanese financial filings
- **yfinance** — Real-time stock price and market data
- **LibreOffice** — Headless formula recalculation and error checking

---

## About the Author

**Ryosuke Sato** — Financial Analyst & Python Developer

- 8+ years in finance and business operations
- Specializing in Japanese equity research automation and financial modeling
- Python-driven workflow optimization: transforming manual Excel processes into reproducible, auditable pipelines
- Native Japanese + English — bridging Japanese market data for international investors

[GitHub](https://github.com/Ryosuke0369)
