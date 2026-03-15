"""
build_core_2359.py - Core Corporation (2359.T) DCF Model Generator

Research-backed config with segment analysis, verified FCF bridge,
and defense-tech thesis integration.

Usage:
    python examples/core-corporation-2359/build_core_2359.py
"""
import os
import sys
from datetime import datetime

# ── Path setup ──
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, "..", ".."))
sys.path.insert(0, PROJECT_ROOT)

from templates.dcf_comps_template import generate_dcf_workbook, get_live_market_data
from scripts.comps_fetcher import get_comps_data


def build_config():
    """Build config dict for Core Corporation."""

    # ── Live market data ──
    price, shares, beta = get_live_market_data(
        "2359.T", fallback_price=2200, fallback_shares=14_370_201
    )

    # ── D/E ratio (net cash → clamp to 0.0) ──
    net_debt = -6174
    if net_debt < 0:
        de_ratio = 0.0
    else:
        mkt_cap = price * shares / 1_000_000
        de_ratio = min(net_debt / mkt_cap, 2.0) if mkt_cap > 0 else 0.1

    config = {
        # ── Company Info ──
        "company_name": "Core Corporation",
        "ticker": "2359.T",
        "exchange": "TSE Prime",
        "sector": "Information & Communication",
        "current_price": price,
        "shares_outstanding": shares,
        "net_debt": net_debt,

        # ── Historical Financials (JPY mn, oldest-first) ──
        "hist_years": ["FY2024/3", "FY2025/3"],
        "hist_revenue": [23999, 24599],
        "hist_cogs": [17408, 17821],
        "hist_sga": [3451, 3602],
        "hist_operating_income": [3141, 3175],
        "hist_net_income": [2182, 2207],
        "hist_ocf": [2500, 2400],
        "hist_capex": [136, 152],
        "hist_cash": [7200, 7586],
        "hist_debt": [1500, 1412],

        # ── DCF Assumptions ──
        "capex_pct": 0.0062,
        "da_pct": 0.0088,
        "tax_rate": 0.305,
        "risk_free": 0.012,
        "beta": beta,
        "erp": 0.065,
        "size_premium": 0.030,
        "cost_of_debt_at": 0.005,
        "de_ratio": de_ratio,
        "terminal_growth": 0.015,
        "exit_multiple": 10.0,
        "projection_years": 5,

        # ── Stub Period ──
        "stub_fraction": 0.25,
        "stub_months_elapsed": 9,
        "ltm_revenue": 24599,
        "projection_start_fy": "FY2026(E)",

        # ── Base Year Values (FY2025/3, NWC Schedule) ──
        "base_year_revenue": 24599,
        "base_year_cogs": 17821,
        "base_year_ar": 7861,
        "base_year_inv": 247,
        "base_year_ap": 1735,

        # ── Scenarios (5 cases × 5 years) ──
        "scenarios": {
            "Base": {
                "revenue_growth": [0.098, 0.104, 0.070, 0.050, 0.040],
                # Year1: 24,599→27,000 (9.8%), Year2: 27,000→29,800 (10.4% M&Aフル寄与)
                "cogs_pct": [0.717, 0.713, 0.710, 0.705, 0.700],
                # Year1 OPM=15.0% (1-0.717-0.133), Layer 2上昇で改善トレンド
                "sga_pct": [0.133, 0.130, 0.128, 0.126, 0.125],
                "dso_days": [115, 113, 112, 110, 110],
                "dih_days": [6, 6, 6, 6, 6],
                "dpo_days": [35, 35, 35, 35, 35],
            },
            "Upside": {
                "revenue_growth": [0.098, 0.140, 0.100, 0.080, 0.060],
                # Year1同じ（ほぼ確定）、Year2以降:防衛テック受注+Layer2加速
                "cogs_pct": [0.717, 0.705, 0.695, 0.690, 0.685],
                "sga_pct": [0.130, 0.125, 0.122, 0.120, 0.118],
                "dso_days": [112, 110, 108, 105, 105],
                "dih_days": [5, 5, 5, 5, 5],
                "dpo_days": [36, 37, 38, 38, 38],
            },
            "Management": {
                "revenue_growth": [0.098, 0.080, 0.060, 0.050, 0.040],
                # 会社予想OP 3,500Mベース（保守的）
                "cogs_pct": [0.724, 0.722, 0.720, 0.718, 0.716],
                "sga_pct": [0.146, 0.143, 0.140, 0.138, 0.135],
                "dso_days": [115, 115, 113, 112, 110],
                "dih_days": [6, 6, 6, 6, 6],
                "dpo_days": [35, 35, 35, 35, 35],
            },
            "Downside 1": {
                "revenue_growth": [0.098, 0.040, 0.030, 0.020, 0.020],
                # Layer 2伸び悩み
                "cogs_pct": [0.724, 0.730, 0.735, 0.738, 0.740],
                "sga_pct": [0.140, 0.142, 0.145, 0.145, 0.145],
                "dso_days": [118, 120, 122, 122, 122],
                "dih_days": [7, 7, 7, 7, 7],
                "dpo_days": [34, 33, 33, 33, 33],
            },
            "Downside 2": {
                "revenue_growth": [0.098, 0.000, -0.020, 0.000, 0.010],
                # 景気後退シナリオ
                "cogs_pct": [0.730, 0.745, 0.750, 0.748, 0.745],
                "sga_pct": [0.148, 0.150, 0.155, 0.155, 0.150],
                "dso_days": [120, 125, 128, 128, 125],
                "dih_days": [8, 8, 8, 8, 8],
                "dpo_days": [32, 30, 30, 30, 30],
            },
        },

        # ── Implied Valuation ──
        "core_ebitda": 3391,
        "core_net_income": 2207,

        # ── Investment Thesis ──
        "investment_thesis": [
            "1. Defense Tech: GNSS anti-spoofing tech for JASDF; 4-layer moat"
            " (proprietary receiver + QZSS auth + security clearance"
            " + space tech track record)",
            "2. AX Transformation: Layer 2 (solution) ratio rising"
            " 25.9% -> 33.3%, structurally improving OPM",
            "3. Earnings Surprise: Base case EPS >Y200 vs consensus ~Y174;"
            " Q3 OP progress rate 82.5% implies ~Y4,050M OP",
        ],

        # ── Key Risks ──
        "key_risks": [
            "1. Defense contracts still in R&D/validation phase"
            " -- mass production orders unconfirmed",
            "2. Key person risk in specialized GNSS/space technology division",
            "3. Revenue concentration in Japanese domestic IT services market",
        ],

        # ── Settings ──
        "primary_multiple": "EV/EBITDA",
    }

    # ── Comps (dynamic load from CSV) ──
    comps_csv = os.path.join(PROJECT_ROOT, "data", "comps", "2359_comps.csv")
    if os.path.isfile(comps_csv):
        config["comps"] = get_comps_data(comps_csv)
    else:
        print(f"Warning: Comps CSV not found at {comps_csv}")
        config["comps"] = []

    return config


def main():
    config = build_config()

    output_dir = os.path.join(PROJECT_ROOT, "output")
    os.makedirs(output_dir, exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d")
    output_path = os.path.join(output_dir, f"2359_DCF_Model_{date_str}.xlsx")

    saved = generate_dcf_workbook(config, output_path)
    print(f"\nDCF Model saved: {saved}")

    # Summary
    mkt_cap = config["current_price"] * config["shares_outstanding"] / 1e6
    print(f"\n--- Core Corporation (2359.T) Summary ---")
    print(f"  Price: {config['current_price']:,.0f} JPY")
    print(f"  Shares: {config['shares_outstanding']:,}")
    print(f"  Market Cap: {mkt_cap:,.0f} JPY mn")
    print(f"  Net Debt: {config['net_debt']:,} JPY mn")
    print(f"  EV: {mkt_cap + config['net_debt']:,.0f} JPY mn")
    print(f"  Beta: {config['beta']:.2f}")
    print(f"  Comps loaded: {len(config['comps'])}")


if __name__ == "__main__":
    main()
