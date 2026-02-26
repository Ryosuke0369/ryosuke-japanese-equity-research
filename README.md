# Japanese Equity Research & Financial Modeling

Python-automated financial models for Japanese listed companies. Built by a native Japanese analyst who reads original filings (決算短信, 有価証券報告書, 適時開示) directly.

## Projects

### 1. Core Corporation (2359.T) — DCF & Comparable Company Analysis
- **DCF Valuation**: 5-year FCF projection, WACC (CAPM + size premium), terminal value (perpetuity growth + exit multiple), sensitivity analysis
- **Comps Analysis**: 8 Japanese IT/SIer peers, EV/EBITDA, PER, PBR with percentile-based valuation range
- **Deliverables**: Dynamic Excel model (227 formula cells, zero errors) + Python scripts
- **Key Finding**: Integrated valuation range ¥2,466 - ¥2,946 (+10% to +32% upside)

### 2. Headwaters (4011.T) × BBD Initiative (5259.T) — M&A Accretion/Dilution Analysis
- **Transaction**: Stock-for-stock absorption merger, ratio 0.50
- **Analysis**: Pro forma EPS (3 scenarios), transaction multiples, pro forma balance sheet, goodwill calculation, sensitivity analysis
- **Deliverables**: Dynamic Excel model (7 sheets) + Python scripts
- **Key Finding**: Dilutive under all scenarios (-33% to -106%); ¥131mn+ annual synergies required to break even

### [3. KFC Holdings Japan (9873.T) — LBO Analysis](./kfc-japan-lbo/)
- **Transaction**: Carlyle Group's take-private of KFC Japan from Mitsubishi Corporation (May 2024)
- **Analysis**: Full LBO model with 5-year projections, debt schedule (TLA/TLB), returns analysis (MOIC/IRR), and sensitivity tables
- **Deliverables**: Dynamic Excel model (710 formula cells, zero errors) + 8-page PDF report
- **Key Finding**: 2.4x MOIC / 19.2% IRR at base case (15x exit, Year 5)

## Tech Stack
- Python (pandas, numpy, yfinance, openpyxl)
- Excel (formula-driven dynamic models)
- Data Sources: EDINET, TDnet, yfinance

## About
Financial analyst based in Japan with 8+ years of experience in finance and business operations. Specializing in Japanese equity research for foreign investors, bridging the language and information gap between global capital and Japan's undercovered small/mid-cap market.

## Disclaimer
All analyses are for informational and demonstration purposes only. They do not constitute investment advice.
