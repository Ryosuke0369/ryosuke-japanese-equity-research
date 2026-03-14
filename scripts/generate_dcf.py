"""
generate_dcf.py - One-click DCF model generation from EDINET data.

Usage:
    python scripts/generate_dcf.py 2359
    python scripts/generate_dcf.py 2359 --years 3
    python scripts/generate_dcf.py 2359 --output-dir output
"""

import argparse
import os
import sys
import re
from datetime import datetime
from collections import OrderedDict

# Ensure imports work from project root
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scripts.edinet_fetcher import fetch_and_parse_multi_year
from scripts.comps_fetcher import get_comps_data
from scripts.yfinance_quarterly import enrich_merged_data_with_yfinance
from templates.dcf_comps_template import generate_dcf_workbook, get_live_market_data


# =====================================================================
# MERGED DATA -> CONFIG CONVERSION
# =====================================================================
def merged_data_to_config(company_info, merged_data):
    """Convert EDINET merged_data into the config dict expected by generate_dcf_workbook.

    Args:
        company_info: Dict from edinet_parser with company_name, securities_code, etc.
        merged_data: OrderedDict from fetch_and_parse_multi_year with FY/LTM keys.

    Returns:
        dict: Config dict ready for generate_dcf_workbook().
    """
    # Separate FY keys and LTM key
    fy_keys = [k for k in merged_data if k.startswith("FY")]
    ltm_keys = [k for k in merged_data if k.startswith("LTM")]

    # Filter out FY years with no meaningful data (revenue is None or 0)
    # XBRL prior3/prior4 contexts often exist but contain no extracted values
    fy_keys = [k for k in fy_keys if merged_data[k].get("revenue") is not None]

    # Sort FY keys oldest-first for historical arrays
    fy_keys_oldest_first = sorted(fy_keys)

    # Helper: safe get with 0 fallback
    def _val(data_dict, key, default=0):
        v = data_dict.get(key)
        return v if v is not None else default

    # Build historical arrays (oldest-first)
    hist_revenue = [_val(merged_data[k], "revenue") for k in fy_keys_oldest_first]
    hist_cogs = [_val(merged_data[k], "cogs") for k in fy_keys_oldest_first]
    hist_sga = [_val(merged_data[k], "sga") for k in fy_keys_oldest_first]
    hist_operating_income = [_val(merged_data[k], "operating_income") for k in fy_keys_oldest_first]
    hist_net_income = [_val(merged_data[k], "net_income") for k in fy_keys_oldest_first]
    hist_ocf = [_val(merged_data[k], "operating_cf") for k in fy_keys_oldest_first]
    hist_capex = [_val(merged_data[k], "capex") for k in fy_keys_oldest_first]
    hist_cash = [_val(merged_data[k], "cash") for k in fy_keys_oldest_first]
    hist_debt = [_val(merged_data[k], "total_debt") for k in fy_keys_oldest_first]

    # Handle None COGS: reverse-calculate from revenue - operating_income
    for i in range(len(hist_cogs)):
        if hist_cogs[i] == 0 and hist_revenue[i] != 0:
            # cogs = revenue - operating_income - sga
            hist_cogs[i] = round(hist_revenue[i] - hist_operating_income[i] - hist_sga[i], 1)
            if hist_cogs[i] < 0:
                hist_cogs[i] = 0

    # Base year values: LTM preferred, then latest FY
    latest_fy_key = fy_keys[0] if fy_keys else None  # newest FY (fy_keys are newest-first from merged_data)
    base_key = ltm_keys[0] if ltm_keys else latest_fy_key

    if base_key is None:
        raise ValueError("No FY or LTM data found in merged_data")

    base_data = merged_data[base_key]
    latest_fy_data = merged_data[latest_fy_key] if latest_fy_key else base_data

    base_year_revenue = _val(base_data, "revenue", 1)
    base_year_cogs = _val(base_data, "cogs")
    # If cogs is 0 in base, reverse-calc
    if base_year_cogs == 0 and base_year_revenue != 0:
        base_year_cogs = round(
            _val(base_data, "revenue") - _val(base_data, "operating_income") - _val(base_data, "sga"), 1
        )
        if base_year_cogs < 0:
            base_year_cogs = 0

    base_year_ar = _val(base_data, "accounts_receivable")
    base_year_inv = _val(base_data, "inventories")
    base_year_ap = _val(base_data, "accounts_payable")
    net_debt = _val(base_data, "net_debt")

    # Auto-calculate DCF assumptions: average da_pct/capex_pct across all FY years
    da_ratios = []
    capex_ratios = []
    for k in fy_keys_oldest_first:
        rev_k = _val(merged_data[k], "revenue")
        if rev_k > 0:
            dep_k = _val(merged_data[k], "depreciation")
            capex_k = _val(merged_data[k], "capex")
            if dep_k > 0:
                da_ratios.append(dep_k / rev_k)
            if capex_k > 0:
                capex_ratios.append(capex_k / rev_k)

    da_pct = round(sum(da_ratios) / len(da_ratios), 4) if da_ratios else 0.02
    capex_pct = round(sum(capex_ratios) / len(capex_ratios), 4) if capex_ratios else 0.03

    # Clamp to reasonable ranges
    capex_pct = max(0.005, min(capex_pct, 0.20))
    da_pct = max(0.005, min(da_pct, 0.15))

    latest_rev = _val(latest_fy_data, "revenue", 1)

    # Calculate CAGR from last 3 years of revenue
    if len(hist_revenue) >= 3 and hist_revenue[-3] and hist_revenue[-3] > 0 and hist_revenue[-1] > 0:
        cagr_3yr = (hist_revenue[-1] / hist_revenue[-3]) ** (1 / 3) - 1
    elif len(hist_revenue) >= 2 and hist_revenue[-2] and hist_revenue[-2] > 0 and hist_revenue[-1] > 0:
        cagr_3yr = hist_revenue[-1] / hist_revenue[-2] - 1
    else:
        cagr_3yr = 0.05  # default

    cagr_3yr = round(max(-0.10, min(cagr_3yr, 0.50)), 4)  # clamp

    # Latest FY ratios
    cogs_pct_latest = round(_val(latest_fy_data, "cogs") / latest_rev, 4) if latest_rev else 0.70
    if cogs_pct_latest <= 0 or cogs_pct_latest >= 1:
        cogs_pct_latest = round(base_year_cogs / base_year_revenue, 4) if base_year_revenue else 0.70
    sga_pct_latest = round(_val(latest_fy_data, "sga") / latest_rev, 4) if latest_rev else 0.13
    if sga_pct_latest <= 0 or sga_pct_latest >= 1:
        sga_pct_latest = 0.13

    # NWC day calculations for scenarios
    dso_days = round(base_year_ar / base_year_revenue * 365) if base_year_revenue else 60
    dih_days = round(base_year_inv / base_year_cogs * 365) if base_year_cogs else 30
    dpo_days = round(base_year_ap / base_year_cogs * 365) if base_year_cogs else 45

    # Clamp NWC days to reasonable ranges
    dso_days = max(10, min(dso_days, 180))
    dih_days = max(0, min(dih_days, 180))
    dpo_days = max(10, min(dpo_days, 180))

    # Build scenarios
    base_growth = [round(cagr_3yr, 4)] * 5
    base_cogs = [round(cogs_pct_latest, 4)] * 5
    base_sga = [round(sga_pct_latest, 4)] * 5
    base_dso = [dso_days] * 5
    base_dih = [dih_days] * 5
    base_dpo = [dpo_days] * 5

    scenarios = {
        "Base": {
            "revenue_growth": base_growth,
            "cogs_pct": base_cogs,
            "sga_pct": base_sga,
            "dso_days": base_dso,
            "dih_days": base_dih,
            "dpo_days": base_dpo,
        },
        "Upside": {
            "revenue_growth": [round(cagr_3yr * 1.5, 4)] * 5,
            "cogs_pct": [round(cogs_pct_latest * 0.95, 4)] * 5,
            "sga_pct": [round(sga_pct_latest * 0.90, 4)] * 5,
            "dso_days": [max(10, dso_days - 5)] * 5,
            "dih_days": [max(0, dih_days - 2)] * 5,
            "dpo_days": [dpo_days + 3] * 5,
        },
        "Management": {
            "revenue_growth": [round(cagr_3yr * 1.2, 4)] * 5,
            "cogs_pct": base_cogs,
            "sga_pct": base_sga,
            "dso_days": base_dso,
            "dih_days": base_dih,
            "dpo_days": base_dpo,
        },
        "Downside 1": {
            "revenue_growth": [round(max(cagr_3yr * 0.5, 0.0), 4)] * 5,
            "cogs_pct": [round(cogs_pct_latest * 1.05, 4)] * 5,
            "sga_pct": [round(sga_pct_latest * 1.10, 4)] * 5,
            "dso_days": [dso_days + 5] * 5,
            "dih_days": [dih_days + 3] * 5,
            "dpo_days": [max(10, dpo_days - 3)] * 5,
        },
        "Downside 2": {
            "revenue_growth": [0.0] * 5,
            "cogs_pct": [round(cogs_pct_latest * 1.10, 4)] * 5,
            "sga_pct": [round(sga_pct_latest * 1.15, 4)] * 5,
            "dso_days": [dso_days + 10] * 5,
            "dih_days": [dih_days + 5] * 5,
            "dpo_days": [max(10, dpo_days - 5)] * 5,
        },
    }

    # Company info
    company_name = company_info.get("company_name", "Unknown Company")
    securities_code = company_info.get("securities_code", "0000")
    # EDINET securities_code is 5 digits (e.g. "23590"), strip trailing "0"
    if len(securities_code) == 5 and securities_code.endswith("0"):
        ticker_4digit = securities_code[:4]
    else:
        ticker_4digit = securities_code
    ticker_str = f"{ticker_4digit}.T"

    # Latest operating income + depreciation for EBITDA approximation
    latest_oi = _val(latest_fy_data, "operating_income")
    latest_dep = _val(latest_fy_data, "depreciation")
    core_ebitda = latest_oi + latest_dep
    core_net_income = _val(latest_fy_data, "net_income")


    # ── Stub Period Calculation ──
    # Determines how far into the current FY we are, based on LTM/quarterly data
    ltm_label = ltm_keys[0] if ltm_keys else None
    stub_fraction = 1.0  # default: no stub (full year ahead)
    stub_months_elapsed = 0
    ltm_revenue = base_year_revenue  # fallback to latest FY

    if ltm_label:
        # Parse quarter number from LTM label like "LTM(2Q 2025-09)"
        m = re.search(r"(\d)Q", ltm_label)
        if m:
            quarter_number = int(m.group(1))
            # Months elapsed in the new FY = quarter_number * 3
            # e.g. Q2 data → 6 months elapsed → 6 months remaining
            stub_months_elapsed = quarter_number * 3
            stub_fraction = (12 - stub_months_elapsed) / 12
            # Edge case: stub_fraction = 0 means FY just ended → treat as full year
            if stub_fraction <= 0:
                stub_fraction = 1.0
                stub_months_elapsed = 0

        ltm_revenue = _val(merged_data[ltm_label], "revenue", base_year_revenue)

    # ── Projection Start FY Label ──
    # Derive next FY label from latest FY key
    if fy_keys_oldest_first:
        latest_fy_label = fy_keys_oldest_first[-1]  # e.g. "FY2025"
        m = re.search(r"FY(\d+)", latest_fy_label)
        if m:
            next_fy_year = int(m.group(1)) + 1
            projection_start_fy = f"FY{next_fy_year}(E)"
        else:
            projection_start_fy = "Year 1(E)"
    else:
        projection_start_fy = "Year 1(E)"

    config = {
        # Company Info
        "company_name": company_name,
        "ticker": ticker_str,
        "exchange": "TSE",
        "sector": "N/A",
        "current_price": 1000,  # placeholder, overridden by yfinance
        "shares_outstanding": 10_000_000,  # placeholder, overridden by yfinance
        "net_debt": net_debt,

        # Historical Financials (JPY mn, oldest-first)
        "hist_years": fy_keys_oldest_first,
        "hist_revenue": hist_revenue,
        "hist_operating_income": hist_operating_income,
        "hist_net_income": hist_net_income,
        "hist_cogs": hist_cogs,
        "hist_sga": hist_sga,
        "hist_ocf": hist_ocf,
        "hist_capex": hist_capex,
        "hist_cash": hist_cash,
        "hist_debt": hist_debt,

        # DCF Assumptions
        "scenarios": scenarios,
        "capex_pct": capex_pct,
        "da_pct": da_pct,
        "tax_rate": 0.30,
        "risk_free": 0.022,   # Japan 10Y JGB yield
        "beta": 1.20,
        "erp": 0.065,         # Japan equity risk premium
        "size_premium": 0.030,
        "cost_of_debt_at": 0.010,
        "de_ratio": 0.10,
        "terminal_growth": 0.02,
        "exit_multiple": 10.0,
        "projection_years": 5,

        # Stub Period
        "stub_fraction": stub_fraction,
        "stub_months_elapsed": stub_months_elapsed,
        "ltm_revenue": ltm_revenue,
        "projection_start_fy": projection_start_fy,

        # Base Year Values
        "base_year_revenue": base_year_revenue,
        "base_year_cogs": base_year_cogs,
        "base_year_ar": base_year_ar,
        "base_year_inv": base_year_inv,
        "base_year_ap": base_year_ap,

        # Comps (empty by default — can be populated separately)
        "comps": [],

        # Implied Valuation
        "core_ebitda": core_ebitda,
        "core_net_income": core_net_income,

        # Investment Thesis & Risks (placeholders)
        "investment_thesis": [
            "1. [Edit] Describe key competitive advantage",
            "2. [Edit] Describe growth driver",
            "3. [Edit] Describe margin expansion opportunity",
        ],
        "key_risks": [
            "1. [Edit] Describe primary risk factor",
            "2. [Edit] Describe secondary risk factor",
            "3. [Edit] Describe tertiary risk factor",
        ],

        # Settings
        "primary_multiple": "EV/EBITDA" if core_ebitda > 0 else "EV/Sales",
    }

    return config


# =====================================================================
# CLI
# =====================================================================
def main():
    parser = argparse.ArgumentParser(
        description="Generate DCF model from EDINET data",
        usage="python scripts/generate_dcf.py TICKER [--years N] [--output-dir DIR] [--comps-csv PATH]",
    )
    parser.add_argument("ticker", help="Securities code (e.g. 2359)")
    parser.add_argument("--years", type=int, default=5, help="Number of years to fetch (default: 5)")
    parser.add_argument("--output-dir", default="output", help="Output directory (default: output)")
    parser.add_argument("--comps-csv", default=None, help="Path to comps CSV (default: data/comps/<ticker>_comps.csv)")
    args = parser.parse_args()

    ticker_code = args.ticker.strip()
    num_years = min(args.years, 5)

    print(f"\n{'=' * 60}")
    print(f"DCF Model Generator - Ticker: {ticker_code}")
    print(f"{'=' * 60}")

    # Step 1: EDINET fetch + parse
    print(f"\n[Step 1/6] Fetching {num_years} years of financial data from EDINET...")
    company_info, merged_data = fetch_and_parse_multi_year(ticker_code, num_years)

    # Step 2: Check LTM coverage, yfinance fallback if needed
    print(f"\n[Step 2/6] Checking LTM data coverage...")
    ticker_4digit = re.sub(r"0$", "", (company_info.get("securities_code") or ticker_code)[:5])
    config_ticker_str = f"{ticker_4digit}.T"
    fiscal_year_end = company_info.get("fiscal_year_end")
    merged_data = enrich_merged_data_with_yfinance(
        merged_data, config_ticker_str, fiscal_year_end
    )

    # Step 3: Convert to config
    print(f"\n[Step 3/6] Building DCF configuration...")
    config = merged_data_to_config(company_info, merged_data)

    # Step 4: Fetch live market data via yfinance (price, shares, beta)
    print(f"\n[Step 4/6] Fetching market data...")
    ticker_str = config["ticker"]
    config["current_price"], config["shares_outstanding"], live_beta = get_live_market_data(
        ticker_str, config["current_price"], config["shares_outstanding"]
    )
    config["beta"] = live_beta  # raw value; template normalizes to [0.6, 1.5]

    # Auto-calculate D/E ratio from net_debt and market cap
    try:
        market_cap = config["current_price"] * config["shares_outstanding"] / 1_000_000  # JPY mn
        if config["net_debt"] > 0 and market_cap > 0:
            config["de_ratio"] = round(config["net_debt"] / market_cap, 4)
        else:
            config["de_ratio"] = 0.0
        config["de_ratio"] = min(config["de_ratio"], 2.0)
    except Exception:
        market_cap = 0
        config["de_ratio"] = 0.10

    # Beta & Size Premium are normalized inside generate_dcf_workbook (template-level)
    print(f"  Raw Beta: {live_beta:.2f}, D/E Ratio: {config['de_ratio']:.4f}, Mkt Cap: {market_cap:,.0f} mn")

    # Step 5: Load comparable companies data
    print(f"\n[Step 5/6] Loading comparable companies...")
    # Resolve comps CSV path: --comps-csv > data/comps/<ticker>_comps.csv
    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    if args.comps_csv:
        comps_csv_path = args.comps_csv
    else:
        comps_csv_path = os.path.join(project_root, "data", "comps", f"{ticker_code}_comps.csv")

    if os.path.isfile(comps_csv_path):
        try:
            config["comps"] = get_comps_data(comps_csv_path)
            print(f"  Loaded {len(config['comps'])} comps from {comps_csv_path}")
        except Exception as e:
            print(f"  WARNING: Failed to load comps from {comps_csv_path}: {e}")
            print(f"  Continuing without comps data.")
            config["comps"] = []
    else:
        print(f"  No comps CSV found at: {comps_csv_path}")
        print(f"  Continuing without comps data. To add comps, create the CSV or use --comps-csv.")
        config["comps"] = []

    # Step 6: Generate Excel
    print(f"\n[Step 6/6] Generating DCF workbook...")
    os.makedirs(args.output_dir, exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d")
    output_path = os.path.join(args.output_dir, f"{ticker_code}_DCF_Model_{date_str}.xlsx")

    saved_path = generate_dcf_workbook(config, output_path)

    print(f"\n{'=' * 60}")
    print(f"DCF Model saved: {saved_path}")
    print(f"{'=' * 60}")

    # Print summary
    print(f"\nSummary:")
    print(f"  Company:    {config['company_name']}")
    print(f"  Ticker:     {config['ticker']}")
    print(f"  Price:      {config['current_price']:,.0f}")
    print(f"  Shares:     {config['shares_outstanding']:,}")
    print(f"  Base Rev:   {config['base_year_revenue']:,.0f} mn")
    print(f"  Net Debt:   {config['net_debt']:,.0f} mn")
    print(f"  Hist Years: {len(config['hist_years'])}")
    print(f"  LTM Rev:    {config['ltm_revenue']:,.0f} mn")
    print(f"  Stub:       {config['stub_fraction']:.2f} ({config['stub_months_elapsed']}m elapsed)")
    print(f"  Proj Start: {config['projection_start_fy']}")
    print(f"  Comps:      {len(config['comps'])} companies")


if __name__ == "__main__":
    main()
