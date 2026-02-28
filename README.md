# Japanese Equity Research — Financial Models & Templates

Automated financial model generators for Japanese equity analysis.
Each model uses a config-driven Python build script that generates
professional Excel workbooks with live formulas.

## Models

### 1. DCF & Comparable Company Analysis
**Base case:** Core Corporation (2359.T) — GIS/Defense IT
- 5-sheet Excel: Executive Summary → Financial Statements → DCF Model → Comps Analysis → Sensitivity
- Config-driven: Edit `config` dict → run `python dcf_comps_build.py` → Excel generated
- 152 Excel formulas, cross-sheet references, BUY/HOLD/SELL recommendation logic
- Target: ¥2,734 (BUY, +22% upside from ¥2,240)

### 2. LBO Analysis
**Base case:** KFC Holdings Japan (9873) — Carlyle Take-Private (May 2024)
- 8-sheet Excel: Cover → Transaction Assumptions → IS → Transaction Summary → BS → CF → Debt Schedule → Returns
- Config-driven: Edit `config` dict → run `python lbo_build.py` → Excel generated
- Full 3-statement model with BS balance check and CF consistency check
- Returns: 2.4x MOIC / 19.2% IRR (15x exit, Year 5)

### 3. M&A Accretion/Dilution Analysis
**Base case:** Headwaters (4011.T) × BBD Initiative (5259.T) — Stock-for-Stock Merger
- 7-sheet Excel: Executive Summary → Transaction Overview → Pre-Deal Financials → Pro Forma EPS → Transaction Multiples → Pro Forma BS → Sensitivity
- Pro Forma EPS impact across 3 scenarios (Base/Downside/Synergy)

## Tech Stack
- Python + openpyxl (Excel generation with live formulas)
- recalc.py (LibreOffice-based formula recalculation & error checking)
- Consistent styling across all models (blue=input, black=formula, green=cross-sheet ref)

## How to Use Templates

```bash
# Example: Generate DCF/Comps for a new ticker
# 1. Edit config dict in dcf_comps_build.py with new company data
# 2. Run
python core-corporation-2359/dcf_comps_build.py
# 3. Verify
python scripts/recalc.py output.xlsx
# 4. Open in Excel — formulas auto-calculate
```

## Repository Structure

```
├── core-corporation-2359/
│   ├── Core_Corporation_2359T_Equity_Research.xlsx
│   └── dcf_comps_build.py
├── headwaters-bbd-merger/
│   ├── HW_BBD_MA_Accretion_Dilution.xlsx
│   └── HW_BBD_MA_build.py
├── kfc-japan-lbo/
│   ├── KFC_Japan_LBO_EN.xlsx
│   └── lbo_build.py
└── scripts/
    └── recalc.py
```

## About
Built by Ryosuke Sato — Freelance Financial Analyst (Japan)
- 8+ years in finance & business operations
- Native Japanese + English for cross-border equity research
- Specializing in Japanese market analysis for international investors
