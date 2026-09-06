"""
generate_dcf.py - One-click DCF model generation from EDINET data.

Usage:
    python scripts/generate_dcf.py 2359
    python scripts/generate_dcf.py 2359 --years 3
    python scripts/generate_dcf.py 2359 --output-dir output
"""

import argparse
import json
import os
import subprocess
import sys
import re
from datetime import datetime
from collections import OrderedDict

# The Windows console is cp932 here: a single un-encodable character (an em dash
# in a warning) would raise UnicodeEncodeError and kill an otherwise good run.
# Degrade those characters instead of the process.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass

# Ensure imports work from project root
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scripts.edinet_fetcher import fetch_and_parse_multi_year, fetch_tanshin
from scripts.comps_fetcher import get_comps_data
from scripts.yfinance_quarterly import enrich_merged_data_with_yfinance
from scripts.overrides_validator import validate_overrides, OverridesValidationError
from scripts.guidance_fetcher import get_guidance
from templates.dcf_comps_template import generate_dcf_workbook, get_live_market_data, calc_wacc

# Exit codes. A caller that pipes stdout (batch/regen.sh) sees only text, so
# the distinction between "no workbook was written" and "a workbook was
# written but failed validation" has to live in the exit status.
#   0 = generated and validated
#   1 = generated but validate_output.py reported FAIL (file kept for triage)
#   2 = bad invocation / contract violation (argparse, overrides validator)
#   3 = nothing generated: the output file exists and --force was not given
EXIT_SKIPPED = 3


def _resolve_date_stamp(raw):
    """Validate --date and fall back to today.

    The date in the filename is the ANALYSIS BASIS date, not the wall-clock
    run date: a batch that starts before midnight and finishes after it used
    to split into <ticker>_DCF_Model_20260906.xlsx and ..._20260907.xlsx, and
    the earlier file stayed behind as a stale twin. part4 of the 2026-09-05
    batch worked around this with a TARGET_DATE environment variable and a
    rename step in batch/regen.sh; --date makes it a first-class argument.
    """
    if raw is None:
        return datetime.now().strftime("%Y%m%d")
    s = str(raw).strip()
    try:
        return datetime.strptime(s, "%Y%m%d").strftime("%Y%m%d")
    except ValueError:
        print(f"ERROR: --date must be YYYYMMDD (got {raw!r}).")
        sys.exit(2)


# =====================================================================
# MERGED DATA -> CONFIG CONVERSION
# =====================================================================
def merged_data_to_config(company_info, merged_data, forecast_data=None):
    """Convert EDINET merged_data into the config dict expected by generate_dcf_workbook.

    Args:
        company_info: Dict from edinet_parser with company_name, securities_code, etc.
        merged_data: OrderedDict from fetch_and_parse_multi_year with FY/LTM keys.
        forecast_data: Optional dict with forecast_revenue, forecast_operating_income, etc.
                       If provided, Management scenario Year 1 growth uses guidance.

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

    # A value EDINET did not extract is MISSING, not zero. `_val` used to
    # default to 0, and that single default produced four distinct classes of
    # wrong number in the 2026-09-05 batch:
    #   - operating income 0 in an old year -> "OI growth = (0-0)/0" -> #DIV/0!
    #     on the Financial Statements sheet (8 tickers)
    #   - net_debt 0 where the company actually held net cash (2801: -50,341 mn,
    #     ~JPY 54/share) or net debt (2897: 77,200 mn reported as 4,418)
    #   - core_ebitda = OI + D&A silently computed from a missing leg
    #   - the Reverse DCF derivations inheriting all of the above
    # `_val` now returns None for a missing key. `_num` is the explicit opt-in
    # for the places where a numeric default really is the intended semantics
    # (ratio denominators, `or`-chained fallbacks), so every zero in this
    # function is now a zero somebody chose.
    def _val(data_dict, key):
        return data_dict.get(key)

    def _num(data_dict, key, default=0):
        v = data_dict.get(key)
        return v if v is not None else default

    # Build historical arrays (oldest-first). None = the filing did not give it.
    hist_revenue = [_val(merged_data[k], "revenue") for k in fy_keys_oldest_first]
    hist_cogs = [_val(merged_data[k], "cogs") for k in fy_keys_oldest_first]
    hist_sga = [_val(merged_data[k], "sga") for k in fy_keys_oldest_first]
    hist_operating_income = [_val(merged_data[k], "operating_income") for k in fy_keys_oldest_first]
    hist_net_income = [_val(merged_data[k], "net_income") for k in fy_keys_oldest_first]
    hist_ocf = [_val(merged_data[k], "operating_cf") for k in fy_keys_oldest_first]
    hist_capex = [_val(merged_data[k], "capex") for k in fy_keys_oldest_first]
    hist_cash = [_val(merged_data[k], "cash") for k in fy_keys_oldest_first]
    hist_debt = [_val(merged_data[k], "total_debt") for k in fy_keys_oldest_first]
    hist_depreciation = [_val(merged_data[k], "depreciation") for k in fy_keys_oldest_first]

    _missing_pl = {
        name: [fy_keys_oldest_first[i] for i, v in enumerate(series) if v is None]
        for name, series in (("operating_income", hist_operating_income),
                             ("net_income", hist_net_income),
                             ("cogs", hist_cogs),
                             ("sga", hist_sga))
    }
    for name, years in _missing_pl.items():
        if years:
            print(f"  WARNING: EDINET has no {name} for {', '.join(years)} - "
                  f"left BLANK (was silently 0 before フェーズ2 #1). Supply "
                  f"hist_{name} in overrides if the number matters.")

    # Keep the EDINET series keyed by fiscal year, not by position. hist_years is
    # frequently replaced wholesale by overrides (different labels, different
    # count) while these series are not — copying them positionally is what put
    # FY2022 operating cash flow under the FY2025 column. align_hist_to_years()
    # re-attaches them by year key after the overrides are applied.
    edinet_fs_by_year = {
        k: {
            "hist_ocf": merged_data[k].get("operating_cf"),
            "hist_cash": merged_data[k].get("cash"),
            "hist_debt": merged_data[k].get("total_debt"),
            "hist_capex": merged_data[k].get("capex"),
            "hist_depreciation": merged_data[k].get("depreciation"),
        }
        for k in fy_keys_oldest_first
    }

    # Missing COGS: reverse-calculate from revenue - operating income - SGA,
    # but only when ALL THREE inputs are present. Back-solving from a missing
    # operating income treated as 0 is how a plausible-looking COGS got written
    # for a year the filing said nothing about (2802's COGS came out equal to
    # revenue). A year that cannot be reconstructed stays blank.
    for i in range(len(hist_cogs)):
        if hist_cogs[i] is not None:
            continue
        parts = (hist_revenue[i], hist_operating_income[i], hist_sga[i])
        if any(p is None for p in parts) or not hist_revenue[i]:
            continue
        hist_cogs[i] = max(0.0, round(parts[0] - parts[1] - parts[2], 1))

    # Base year values: LTM preferred, then latest FY
    latest_fy_key = fy_keys[0] if fy_keys else None  # newest FY (fy_keys are newest-first from merged_data)
    base_key = ltm_keys[0] if ltm_keys else latest_fy_key

    if base_key is None:
        raise ValueError("No FY or LTM data found in merged_data")

    base_data = merged_data[base_key]
    latest_fy_data = merged_data[latest_fy_key] if latest_fy_key else base_data

    # Base year revenue/cogs: use latest FY actuals (not LTM)
    # This ensures projection Year 1 connects naturally to the last historical FY
    # (e.g., FY2025: 21,579 → FY2026(E): 21,579 × 1.10 = 23,737)
    # LTM revenue is kept separately for reference/stub discounting.
    latest_annual_fy = merged_data[fy_keys_oldest_first[-1]] if fy_keys_oldest_first else base_data
    # base_year_revenue used to default to 1 (JPY 1mn) when absent, which turns
    # every ratio built on it into a four-orders-of-magnitude artefact instead
    # of an error. It is now None, and main() stops the run unless overrides
    # supply it.
    base_year_revenue = _val(latest_annual_fy, "revenue")
    base_year_cogs = _val(latest_annual_fy, "cogs")
    if base_year_cogs is None:
        _p = (base_year_revenue,
              _val(latest_annual_fy, "operating_income"),
              _val(latest_annual_fy, "sga"))
        if all(x is not None for x in _p) and _p[0]:
            base_year_cogs = max(0.0, round(_p[0] - _p[1] - _p[2], 1))

    # NWC base year: prefer latest FY annual BS over LTM snapshot
    # LTM BS is a point-in-time snapshot that may not be representative
    # (e.g., equipment makers have volatile AR depending on delivery timing)
    # The NWC base-year items keep numeric coalescing: they feed `or`-chained
    # fallbacks and day-count denominators that need falsy semantics, and no
    # wrong number in the batch traced back to them. What changes is that a
    # missing item is now reported instead of passing as a real zero.
    latest_annual_key = fy_keys_oldest_first[-1] if fy_keys_oldest_first else None
    if latest_annual_key:
        latest_annual = merged_data[latest_annual_key]
        _absent_bs = [k for k in ("accounts_receivable", "inventories", "accounts_payable")
                      if latest_annual.get(k) is None and base_data.get(k) is None]
        if _absent_bs:
            print(f"  WARNING: base-year BS items absent from EDINET: "
                  f"{', '.join(_absent_bs)} - treated as 0 for the NWC day counts. "
                  f"Set base_year_ar / base_year_inv / base_year_ap in overrides "
                  f"if the working-capital bridge matters for this ticker.")
        base_year_ar = _num(latest_annual, "accounts_receivable") or _num(base_data, "accounts_receivable")
        base_year_inv = _num(latest_annual, "inventories") or _num(base_data, "inventories")
        base_year_ap = _num(latest_annual, "accounts_payable") or _num(base_data, "accounts_payable")
    else:
        base_year_ar = _num(base_data, "accounts_receivable")
        base_year_inv = _num(base_data, "inventories")
        base_year_ap = _num(base_data, "accounts_payable")
    # ── Trade Receivables/Payables Total (for revenue_pct NWC method) ──
    if latest_annual_key:
        latest_annual_trt = _num(merged_data[latest_annual_key], "trade_receivables_total")
        latest_annual_tpt = _num(merged_data[latest_annual_key], "trade_payables_total")
    else:
        latest_annual_trt = 0
        latest_annual_tpt = 0
    # Fallback: if trade_receivables_total not available, use accounts_receivable
    base_year_trade_receivables = latest_annual_trt if latest_annual_trt else base_year_ar
    base_year_trade_payables = latest_annual_tpt if latest_annual_tpt else base_year_ap
    base_year_nwc = base_year_trade_receivables + base_year_inv - base_year_trade_payables

    # Historical NWC % of Revenue (for revenue_pct method)
    hist_nwc_pct = []
    for k in fy_keys_oldest_first:
        rev_k = _num(merged_data[k], "revenue")
        if rev_k > 0:
            trt_k = _num(merged_data[k], "trade_receivables_total") or _num(merged_data[k], "accounts_receivable")
            inv_k = _num(merged_data[k], "inventories")
            tpt_k = _num(merged_data[k], "trade_payables_total") or _num(merged_data[k], "accounts_payable")
            nwc_k = trt_k + inv_k - tpt_k
            hist_nwc_pct.append(round(nwc_k / rev_k, 4))
        else:
            hist_nwc_pct.append(0)

    # Net debt: never 0-by-default. A DCF's equity value is EV minus this number,
    # so a fabricated zero moves the per-share answer directly (2801 キッコーマン
    # held JPY 50,341 mn of NET CASH and was valued as if it held none). When
    # EDINET gives no net_debt line, derive it from the debt and cash balances;
    # when even that is impossible, leave it None so main() can stop the run.
    net_debt = _val(base_data, "net_debt")
    net_debt_source = "EDINET net_debt"
    if net_debt is None:
        _d, _c = base_data.get("total_debt"), base_data.get("cash")
        if _d is not None and _c is not None:
            net_debt = _d - _c
            net_debt_source = f"derived: total_debt {_d:,.0f} - cash {_c:,.0f}"
            print(f"  [net_debt] EDINET gave no net_debt line; {net_debt_source} "
                  f"= {net_debt:,.0f} mn")
        else:
            net_debt_source = "UNAVAILABLE (no net_debt, and total_debt/cash incomplete)"
            print(f"  WARNING: net_debt could not be determined from EDINET "
                  f"(total_debt={_d}, cash={_c}). It must be supplied in overrides.")

    # Auto-calculate DCF assumptions: average da_pct/capex_pct across all FY years
    da_ratios = []
    capex_ratios = []
    for k in fy_keys_oldest_first:
        rev_k = _num(merged_data[k], "revenue")
        if rev_k > 0:
            dep_k = _num(merged_data[k], "depreciation")
            capex_k = _num(merged_data[k], "capex")
            if dep_k > 0:
                da_ratios.append(dep_k / rev_k)
            if capex_k > 0:
                capex_ratios.append(capex_k / rev_k)

    da_pct = round(sum(da_ratios) / len(da_ratios), 4) if da_ratios else 0.02
    capex_pct = round(sum(capex_ratios) / len(capex_ratios), 4) if capex_ratios else 0.03

    # Clamp to reasonable ranges
    capex_pct = max(0.005, min(capex_pct, 0.20))
    da_pct = max(0.005, min(da_pct, 0.15))

    latest_rev = _num(latest_fy_data, "revenue", 0)

    # Calculate CAGR from last 3 years of revenue
    _r = hist_revenue
    if len(_r) >= 3 and _r[-3] and _r[-1] and _r[-3] > 0 and _r[-1] > 0:
        cagr_3yr = (_r[-1] / _r[-3]) ** (1 / 3) - 1
    elif len(_r) >= 2 and _r[-2] and _r[-1] and _r[-2] > 0 and _r[-1] > 0:
        cagr_3yr = _r[-1] / _r[-2] - 1
    else:
        cagr_3yr = 0.05  # default

    cagr_3yr = round(max(-0.10, min(cagr_3yr, 0.50)), 4)  # clamp

    # Latest FY ratios
    _latest_cogs = _num(latest_fy_data, "cogs", 0)
    _latest_sga = _num(latest_fy_data, "sga", 0)
    cogs_pct_latest = round(_latest_cogs / latest_rev, 4) if latest_rev else 0.70
    if cogs_pct_latest <= 0 or cogs_pct_latest >= 1:
        cogs_pct_latest = (round(base_year_cogs / base_year_revenue, 4)
                           if (base_year_revenue and base_year_cogs) else 0.70)
    sga_pct_latest = round(_latest_sga / latest_rev, 4) if latest_rev else 0.13
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

    # NWC % of Revenue for revenue_pct method
    latest_nwc_pct = hist_nwc_pct[-1] if hist_nwc_pct else 0.50
    base_nwc_pct = [round(latest_nwc_pct, 4)] * 5
    upside_nwc_pct = [round(latest_nwc_pct * 0.90, 4)] * 5
    mgmt_nwc_pct = [round(latest_nwc_pct, 4)] * 5
    ds1_nwc_pct = [round(latest_nwc_pct * 1.10, 4)] * 5
    ds2_nwc_pct = [round(latest_nwc_pct * 1.20, 4)] * 5

    scenarios = {
        "Base": {
            "revenue_growth": base_growth,
            "cogs_pct": base_cogs,
            "sga_pct": base_sga,
            "dso_days": base_dso,
            "dih_days": base_dih,
            "dpo_days": base_dpo,
            "nwc_pct": base_nwc_pct,
        },
        "Upside": {
            "revenue_growth": [round(cagr_3yr * 1.5, 4)] * 5,
            "cogs_pct": [round(cogs_pct_latest * 0.95, 4)] * 5,
            "sga_pct": [round(sga_pct_latest * 0.90, 4)] * 5,
            "dso_days": [max(10, dso_days - 5)] * 5,
            "dih_days": [max(0, dih_days - 2)] * 5,
            "dpo_days": [dpo_days + 3] * 5,
            "nwc_pct": upside_nwc_pct,
        },
        "Management": {
            "revenue_growth": [round(cagr_3yr * 1.2, 4)] * 5,
            "cogs_pct": base_cogs,
            "sga_pct": base_sga,
            "dso_days": base_dso,
            "dih_days": base_dih,
            "dpo_days": base_dpo,
            "nwc_pct": mgmt_nwc_pct,
        },
        "Downside 1": {
            "revenue_growth": [round(max(cagr_3yr * 0.5, 0.0), 4)] * 5,
            "cogs_pct": [round(cogs_pct_latest * 1.05, 4)] * 5,
            "sga_pct": [round(sga_pct_latest * 1.10, 4)] * 5,
            "dso_days": [dso_days + 5] * 5,
            "dih_days": [dih_days + 3] * 5,
            "dpo_days": [max(10, dpo_days - 3)] * 5,
            "nwc_pct": ds1_nwc_pct,
        },
        "Downside 2": {
            "revenue_growth": [0.0] * 5,
            "cogs_pct": [round(cogs_pct_latest * 1.10, 4)] * 5,
            "sga_pct": [round(sga_pct_latest * 1.15, 4)] * 5,
            "dso_days": [dso_days + 10] * 5,
            "dih_days": [dih_days + 5] * 5,
            "dpo_days": [max(10, dpo_days - 5)] * 5,
            "nwc_pct": ds2_nwc_pct,
        },
    }

    # ── Guidance Integration: override Management scenario Year 1 ──
    if forecast_data and forecast_data.get("forecast_revenue"):
        guidance_rev = forecast_data["forecast_revenue"]
        # Implied Year 1 growth = (guidance_revenue / latest_actual_revenue) - 1
        if base_year_revenue and base_year_revenue > 0:
            guidance_growth = round(guidance_rev / base_year_revenue - 1, 4)
            # Clamp to reasonable range
            guidance_growth = max(-0.50, min(guidance_growth, 3.0))
            # Update Management scenario Year 1 with guidance growth
            mgmt_growth = scenarios["Management"]["revenue_growth"]
            mgmt_growth[0] = guidance_growth
            # Taper remaining years toward CAGR
            for i in range(1, 5):
                blend = guidance_growth * (1 - i / 5) + cagr_3yr * (i / 5)
                mgmt_growth[i] = round(blend, 4)
            print(f"  Guidance integrated: FY Rev forecast = {guidance_rev:,.0f} mn "
                  f"-> Year 1 growth = {guidance_growth:.1%}")

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
    # core_ebitda drives both Comps legs. Summing a present OI with an absent
    # D&A (or vice versa) produced an "EBITDA" that was really just one of its
    # two components — 4502's comps-implied price came out at JPY -2,761 that
    # way. Either both legs are there or the value is None and the Comps legs
    # are labelled N/A downstream.
    latest_oi = _val(latest_fy_data, "operating_income")
    latest_dep = _val(latest_fy_data, "depreciation")
    if latest_oi is None or latest_dep is None:
        core_ebitda = None
        print(f"  WARNING: core_ebitda unavailable (operating_income="
              f"{latest_oi}, depreciation={latest_dep} for the latest FY). "
              f"Set core_ebitda in overrides, or the Comps EV/EBITDA leg is N/A.")
    else:
        core_ebitda = latest_oi + latest_dep
    core_net_income = _val(latest_fy_data, "net_income")


    # ── Stub Period Calculation ──
    # Determines how far into the current FY we are, based on LTM/quarterly data
    ltm_label = ltm_keys[0] if ltm_keys else None
    stub_fraction = 1.0  # default: no stub (full year ahead)
    stub_months_elapsed = 0
    ltm_revenue = base_year_revenue  # fallback to latest FY
    ltm_components = None

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

        ltm_revenue = _num(merged_data[ltm_label], "revenue", base_year_revenue)
        ltm_components = merged_data[ltm_label].get("_ltm_revenue_components")

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
        # None, never a placeholder: main() stops the run if neither yfinance
        # nor the overrides supply these (フェーズ2 #2).
        "current_price": None,
        "shares_outstanding": None,
        "net_debt": net_debt,
        "_net_debt_source": net_debt_source,

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
        "hist_depreciation": hist_depreciation,
        "_edinet_fs_by_year": edinet_fs_by_year,

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
        "_ltm_revenue_auto": ltm_revenue,
        "_ltm_revenue_components": (
            "; ".join(f"{k}={v}" for k, v in ltm_components.items())
            if ltm_components else None
        ),
        "_ltm_revenue_source": ("hybrid LTM (FY - prior cum + current cum)"
                                if ltm_components else
                                ("LTM row" if ltm_label else "latest FY (no LTM)")),
        "projection_start_fy": projection_start_fy,

        # Base Year Values
        "base_year_revenue": base_year_revenue,
        "base_year_cogs": base_year_cogs,
        "base_year_ar": base_year_ar,
        "base_year_inv": base_year_inv,
        "base_year_ap": base_year_ap,
        "base_year_trade_receivables": base_year_trade_receivables,
        "base_year_trade_payables": base_year_trade_payables,
        "base_year_nwc": base_year_nwc,
        "hist_nwc_pct": hist_nwc_pct,
        "nwc_method": "days",  # default; overridden to "revenue_pct" via overrides

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
        "primary_multiple": ("EV/EBITDA" if (core_ebitda or 0) > 0 else "EV/Sales"),
    }

    return config


# =====================================================================
# FISCAL-YEAR KEY ALIGNMENT (bug B1)
# =====================================================================
# EDINET-derived cash-flow / balance-sheet series were copied into the
# Financial Statements columns by POSITION. When overrides replace hist_years
# with a different label set (or a different number of years), position n stops
# meaning the same fiscal year in both lists, and the value lands under someone
# else's year header — silently, and by up to 3 years in the July 2026 models.
# Values are now attached by fiscal-year KEY; a year with no match is left blank.

_FS_SERIES = ("hist_ocf", "hist_cash", "hist_debt", "hist_capex", "hist_depreciation")


def _fy_key_parts(label):
    """Split a fiscal-year label into (year, month) strings.

    'FY2024/3(12m)' -> ('2024', '3'); 'FY2024' -> ('2024', None).
    """
    s = str(label)
    m = re.search(r"FY\s*(\d{4})", s)
    if not m:
        return (None, None)
    year = m.group(1)
    m2 = re.search(r"FY\s*\d{4}\s*[/-]\s*(\d{1,2})", s)
    return (year, m2.group(1).lstrip("0") if m2 else None)


def _resolve_year_keys(hist_years, source_keys):
    """Map each hist_years label to a source key, or None when it cannot match.

    Resolution ladder (each step must be unambiguous, else no match):
      1. identical label
      2. same fiscal year AND same fiscal-year-end month
      3. same fiscal year, when that year is unique on BOTH sides
    Anything else stays unmatched — no shifting, no padding, no "closest year".
    """
    resolved = {}
    remaining = list(source_keys)

    for lbl in hist_years:
        if lbl in remaining:
            resolved[lbl] = lbl
            remaining.remove(lbl)

    src_parts = {k: _fy_key_parts(k) for k in remaining}
    for lbl in hist_years:
        if lbl in resolved:
            continue
        ly, lm = _fy_key_parts(lbl)
        if ly is None:
            resolved[lbl] = None
            continue
        if lm is not None:
            exact = [k for k, (sy, sm) in src_parts.items() if sy == ly and sm == lm]
            if len(exact) == 1:
                resolved[lbl] = exact[0]
                src_parts.pop(exact[0])
                continue
        same_year_src = [k for k, (sy, _) in src_parts.items() if sy == ly]
        same_year_dst = [l for l in hist_years if _fy_key_parts(l)[0] == ly]
        if len(same_year_src) == 1 and len(same_year_dst) == 1:
            resolved[lbl] = same_year_src[0]
            src_parts.pop(same_year_src[0])
        else:
            resolved[lbl] = None
    return resolved


def align_hist_series_to_years(config, override_keys):
    """Re-attach the EDINET CF/BS series to config['hist_years'] by year key.

    Series supplied explicitly in overrides always win and are left untouched.
    Returns a human-readable coverage string (also written into the workbook).
    """
    hist_years = config.get("hist_years") or []
    source = config.get("_edinet_fs_by_year") or {}
    resolved = _resolve_year_keys(hist_years, list(source.keys()))

    overridden = [s for s in _FS_SERIES if s in override_keys]
    aligned = [s for s in _FS_SERIES if s not in override_keys]

    for series in aligned:
        config[series] = [
            (source.get(resolved.get(lbl)) or {}).get(series) if resolved.get(lbl) else None
            for lbl in hist_years
        ]

    # Coverage is reported PER SERIES on OCF/Cash/Debt — the three the FS sheet
    # shows and the ones that were mis-shifted. Counting a year as covered when
    # any one of the three is present would hide a series-specific hole (3687's
    # FY2025/9 debt), which is precisely the kind of gap this check exists for.
    n_years = len(hist_years)
    per_series = {}
    for s in ("hist_ocf", "hist_cash", "hist_debt"):
        if s in overridden:
            per_series[s] = n_years  # supplied wholesale by the analyst
        else:
            per_series[s] = sum(1 for v in (config.get(s) or []) if v is not None)
    covered = min(per_series.values()) if per_series else n_years
    coverage = (f"OCF {per_series['hist_ocf']}/{n_years}, "
                f"Cash {per_series['hist_cash']}/{n_years}, "
                f"Debt {per_series['hist_debt']}/{n_years}")

    # Per-year audit trail: which source year each column came from, and which of
    # OCF/Cash/Debt actually carry a value. validate_output.py asserts the sheet
    # matches this exactly, cell for cell.
    _short = {"hist_ocf": "ocf", "hist_cash": "cash", "hist_debt": "debt"}
    entries = []
    for i, lbl in enumerate(hist_years):
        filled = [
            _short[s] for s in ("hist_ocf", "hist_cash", "hist_debt")
            if (config.get(s) or [None] * n_years)[i] is not None
        ]
        entries.append(f"{lbl}<-{resolved.get(lbl) or 'BLANK'}[{'+'.join(filled)}]")
    year_map = "; ".join(entries)
    config["_fs_series_coverage"] = per_series
    config["_fs_year_coverage"] = coverage
    config["_fs_year_map"] = year_map
    config["_fs_year_sources"] = (
        f"overrides: {', '.join(overridden) or 'none'} | "
        f"EDINET year-key match: {', '.join(aligned) or 'none'}"
    )

    print(f"  [FS align] hist_years <- EDINET: {year_map}")
    print(f"  [FS align] overrides supplied: {', '.join(overridden) or 'none'}")
    if covered < n_years:
        print(f"  WARNING: OCF/Cash/Debt coverage {coverage} - verify against 短信 "
              f"(unmatched years left BLANK; set hist_ocf / hist_cash / hist_debt "
              f"in overrides to fill them)")
    return coverage, covered, n_years


# =====================================================================
# PEER FRESHNESS (bug C2)
# =====================================================================
PEER_STALE_DAYS = 45


def check_peer_freshness(comps, subject_ticker, max_age_days=PEER_STALE_DAYS):
    """Flag peers whose last traded price is stale (delisted / wrong ticker).

    8267 carried three peers that had been delisted for months; their frozen
    market caps kept feeding the medians. A peer whose last quote is older than
    `max_age_days` is marked so the template keeps the row (with a note) but
    drops it from every statistic.

    Safety rail: if more than half the lookups fail, the problem is the network
    or the yfinance install, not the peer set — nothing is excluded in that case.
    """
    try:
        import yfinance as yf
    except ImportError:
        print("  [Peers] yfinance not installed - freshness check skipped.")
        return

    def _subj(t):
        return str(t or "").strip().upper().split(".")[0]

    subject_key = _subj(subject_ticker)
    today = datetime.now().date()
    flagged, failures, checked = [], [], 0

    for comp in comps:
        if _subj(comp.get("ticker")) == subject_key:
            continue
        checked += 1
        tkr = str(comp.get("ticker", "")).strip()
        try:
            hist = yf.Ticker(tkr).history(period="3mo")
            if hist is None or hist.empty:
                raise ValueError("no price history")
            last_date = hist.index[-1].date()
            age = (today - last_date).days
            if age > max_age_days:
                comp["exclude_from_stats"] = True
                comp["exclude_reason"] = f"(last quote {last_date}, {age}d old)"
                flagged.append(f"{comp['name']} [{tkr}] {comp['exclude_reason']}")
            else:
                print(f"  [Peers] {tkr}: last quote {last_date} ({age}d) OK")
        except Exception as e:
            failures.append((comp, f"{type(e).__name__}: {e}"))

    if checked and len(failures) > checked / 2:
        print(f"  [Peers] WARNING: {len(failures)}/{checked} price lookups failed - "
              f"treating this as an environment problem, NOT as delistings. "
              f"No peer excluded. Re-run with network access to validate.")
    else:
        for comp, err in failures:
            comp["exclude_from_stats"] = True
            comp["exclude_reason"] = "(price unavailable — delisted or wrong ticker?)"
            flagged.append(f"{comp['name']} [{comp.get('ticker')}] price fetch failed: {err}")

    if flagged:
        print("  [Peers] WARNING: excluded from statistics (stale/unavailable):")
        for f in flagged:
            print(f"    - {f}")
    elif checked:
        print(f"  [Peers] All {checked} peer quotes are within {max_age_days} days.")


# =====================================================================
# OVERRIDES FALLBACK: inject 決算短信-derived latest FY actuals
# =====================================================================
def _is_placeholder(v):
    """True if an override value is an unfilled "__CONFIRM__" placeholder string.

    The overrides template marks values that the analyst must confirm from the
    決算短信 (e.g. shares_outstanding, current_price, net_debt) with a
    __CONFIRM__ sentinel. Applying these strings would break numeric math, so
    they are skipped and the auto-derived value (yfinance / EDINET) is kept.
    """
    return isinstance(v, str) and "__CONFIRM__" in v


def _inject_latest_fy_from_overrides(merged_data, overrides):
    """Inject/override the latest fiscal year with 決算短信-derived actuals.

    When overrides define a `fundamentals_fy<YYYY>` block (一次情報 from the
    earnings release), use it as the latest annual actuals — taking priority
    over EDINET for that period, which may not yet be published on EDINET.

    Only acts when such a key exists, so tickers without it (e.g. 4192) are
    completely unaffected.

    Args:
        merged_data: OrderedDict from fetch_and_parse_multi_year (may be empty).
        overrides: Loaded overrides dict.

    Returns:
        The (possibly new) merged_data OrderedDict with the latest FY set.
    """
    fund_keys = []
    for k in overrides:
        m = re.match(r"fundamentals_fy(\d{4})$", k)
        if m:
            fund_keys.append((int(m.group(1)), k))
    if not fund_keys:
        return merged_data

    fy_year, fund_key = max(fund_keys)  # newest fundamentals block
    f = overrides[fund_key]
    fy_label = f"FY{fy_year}"

    rev = f.get("revenue")
    oi = f.get("operating_income")
    ni = f.get("net_income")
    ebitda = f.get("ebitda")
    # D&A reverse-calc: EBITDA = OI + D&A  ->  D&A = EBITDA - OI
    dep = round(ebitda - oi, 1) if (ebitda is not None and oi is not None) else None
    # COGS/SGA for historical display, derived from override ratios if present
    cogs = round(rev * overrides["cogs_pct"], 1) if (rev is not None and overrides.get("cogs_pct")) else None
    sga = round(rev * overrides["sga_pct"], 1) if (rev is not None and overrides.get("sga_pct")) else None
    # net_debt only if numeric (override may hold a "__CONFIRM__" placeholder)
    nd = overrides.get("net_debt")
    nd = nd if isinstance(nd, (int, float)) else None

    record = {
        "revenue": rev,
        "operating_income": oi,
        "net_income": ni,
        "cogs": cogs,
        "sga": sga,
        "depreciation": dep,
        "net_debt": nd,
        "total_assets": f.get("total_assets"),
        "total_liabilities": f.get("total_liabilities"),
        "net_assets": f.get("net_assets"),
    }

    existing = merged_data.get(fy_label)
    if isinstance(existing, dict):
        # FY already present from EDINET → override only the 一次情報 fields,
        # preserving EDINET-derived balance-sheet detail (AR/Inv/AP, cash, debt).
        for k, v in record.items():
            if v is not None:
                existing[k] = v
        print(f"  [Fundamentals] Overrode {fy_label} latest-period actuals "
              f"from overrides.{fund_key} (revenue={rev:,.0f} mn)")
        return merged_data

    # FY not in EDINET data → insert as the newest FY column.
    new_merged = OrderedDict()
    for k, v in merged_data.items():  # keep LTM column(s) first
        if k.startswith("LTM"):
            new_merged[k] = v
    new_merged[fy_label] = record
    for k, v in merged_data.items():
        if k == "_meta" or k.startswith("LTM"):
            continue
        new_merged[k] = v
    new_merged["_meta"] = merged_data.get("_meta", {})
    print(f"  [Fundamentals] Injected {fy_label} as latest FY from "
          f"overrides.{fund_key} (revenue={rev:,.0f} mn, not on EDINET)")
    return new_merged


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
    parser.add_argument("--output-dir", default="models", help="Output directory (default: models)")
    parser.add_argument("--date", default=None, metavar="YYYYMMDD",
                        help="Date stamp for the output filename. Defaults to "
                             "today. Use it to pin the ANALYSIS BASIS date so a "
                             "batch that runs past midnight keeps writing to one "
                             "file per ticker instead of creating a second one.")
    parser.add_argument("--comps-csv", default=None, help="Path to comps CSV (default: data/comps/<ticker>_comps.csv)")
    parser.add_argument("--overrides", default=None,
                        help="Path to JSON override file (e.g. data/overrides/2359_overrides.json)")
    parser.add_argument("--force", action="store_true",
                        help="Overwrite existing output file without warning")
    parser.add_argument("--allow-unconfirmed", action="store_true",
                        help="Allow __CONFIRM__ placeholders in overrides (auto-derived "
                             "fallbacks are used; final runs should not need this)")
    parser.add_argument("--no-comps", action="store_true",
                        help="Explicitly generate without comparable companies "
                             "(otherwise a missing comps CSV is an error)")
    parser.add_argument("--no-peer-check", action="store_true",
                        help="Skip the yfinance peer price-freshness check "
                             "(offline runs)")
    parser.add_argument("--tanshin-fallback", action="store_true",
                        help="Also try the legacy EDINET 決算短信 search when no "
                             "guidance is found. Off by default: EDINET does not "
                             "host 決算短信 and the search costs ~50s to fail.")
    parser.add_argument("--no-recalc", action="store_true",
                        help="Do not recalculate the workbook via Excel COM before "
                             "validation (formula-value checks are then skipped)")
    parser.add_argument("--no-validate", action="store_true",
                        help="Skip the post-generation validate_output.py run")
    args = parser.parse_args()

    ticker_code = args.ticker.strip()
    num_years = min(args.years, 5)

    # -- Output-file protection, decided BEFORE any network work ----------
    # This check used to sit in Step 7, after ~6 minutes of EDINET/yfinance
    # traffic, and it announced the refusal as a "WARNING". Two things went
    # wrong with that: the operator paid the full fetch cost for a run that
    # produced nothing, and a wrapper that pipes stdout (batch/regen.sh) saw
    # only a warning line and went on to validate the STALE workbook - which
    # duly reported VERDICT: PASS for a model that had not been regenerated
    # (batch report 2026-09-05 SS15-6). The refusal now happens first, says
    # ERROR, and exits with a code of its own so a caller can tell "nothing
    # was generated" (3) from "generated but invalid" (1).
    os.makedirs(args.output_dir, exist_ok=True)
    date_str = _resolve_date_stamp(args.date)
    output_path = os.path.join(args.output_dir,
                               f"{ticker_code}_DCF_Model_{date_str}.xlsx")
    if os.path.exists(output_path) and not args.force:
        print()
        print(f"ERROR: {output_path} already exists - nothing was generated.")
        print("   Use --force to overwrite, or rename/move the existing file.")
        print("   Tip: Move finalized models to reports/ to protect them.")
        print(f"   (exit {EXIT_SKIPPED} = generation skipped; the existing file "
              f"is untouched and must NOT be read as a fresh run)")
        sys.exit(EXIT_SKIPPED)

    # Auto-detect overrides file if not specified
    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    if args.overrides is None:
        auto_override = os.path.join(project_root, "data", "overrides", f"{ticker_code}_overrides.json")
        if os.path.isfile(auto_override):
            args.overrides = auto_override
            print(f"  Auto-detected overrides: {auto_override}")

    # Load overrides once (reused in Steps 4.5 and 5)
    _overrides = None
    if args.overrides and os.path.isfile(args.overrides):
        with open(args.overrides, encoding="utf-8") as f:
            _overrides = json.load(f)
        # Fail fast on contract violations: unknown keys, nested WACC blocks,
        # legacy scenario names, wrong array lengths, __CONFIRM__ leftovers.
        # A run that completes must mean every override key was consumed.
        try:
            validate_overrides(_overrides, source_path=args.overrides,
                               allow_unconfirmed=args.allow_unconfirmed)
        except OverridesValidationError as e:
            print(f"\nERROR: {e}")
            sys.exit(1)
        print(f"  Overrides validated: {args.overrides}")

    print(f"\n{'=' * 60}")
    print(f"DCF Model Generator - Ticker: {ticker_code}")
    print(f"{'=' * 60}")

    # Repo hygiene: a ticker code in a script filename means somebody forked
    # the logic instead of adding a config file. Warn here (visible in the
    # normal workflow) rather than only in a standalone check nobody runs.
    try:
        from scripts.check_script_naming import check as _check_script_naming
    except ImportError:
        from check_script_naming import check as _check_script_naming
    _check_script_naming(warn_only=True)

    # Step 1: EDINET fetch + parse
    # Pull FY-end month from overrides so the fetcher searches the correct
    # filing season for non-standard fiscal years (e.g. November-FY companies
    # file 有報 in February, outside the default search windows).
    fy_end_month = _overrides.get("fiscal_year_end_month") if _overrides else None
    # If overrides supply a 一次情報 latest FY (決算短信-derived), an EDINET
    # shortfall (report not yet published, recently-listed) is non-fatal.
    has_fundamentals_override = bool(_overrides) and any(
        re.match(r"fundamentals_fy\d{4}$", k) for k in _overrides
    )
    print(f"\n[Step 1/9] Fetching {num_years} years of financial data from EDINET...")
    try:
        company_info, merged_data = fetch_and_parse_multi_year(
            ticker_code, num_years, fiscal_year_end_month=fy_end_month
        )
    except Exception as e:
        if not has_fundamentals_override:
            raise
        print(f"  WARNING: EDINET fetch failed ({type(e).__name__}: {e}).")
        print(f"  Falling back to overrides fundamentals (no EDINET history).")
        company_info = {
            "company_name": _overrides.get("company_name") or ticker_code,
            "securities_code": str(_overrides.get("ticker") or ticker_code),
            "fiscal_year_end": None,
        }
        merged_data = OrderedDict()
        merged_data["_meta"] = {}

    # Step 1.5: Inject 決算短信-derived latest FY actuals from overrides (if any),
    # taking priority over EDINET for the latest period. No-op when overrides
    # lack a fundamentals_fy<YYYY> key (e.g. 4192) → existing tickers unaffected.
    if _overrides:
        merged_data = _inject_latest_fy_from_overrides(merged_data, _overrides)

    # Step 2: Check LTM coverage, yfinance fallback if needed
    print(f"\n[Step 2/9] Checking LTM data coverage...")
    ticker_4digit = re.sub(r"0$", "", (company_info.get("securities_code") or ticker_code)[:5])
    config_ticker_str = f"{ticker_4digit}.T"
    fiscal_year_end = company_info.get("fiscal_year_end")
    merged_data = enrich_merged_data_with_yfinance(
        merged_data, config_ticker_str, fiscal_year_end
    )

    print()
    # Step 3: Company guidance (業績予想) - see scripts/guidance_fetcher.py for
    # why the old EDINET-only path could not work and what replaced it.
    print(f"[Step 3/9] Resolving company guidance (業績予想)...")
    _fy_years = [int(m.group(1)) for m in
                 (re.search(r"FY(\d{4})", str(k)) for k in merged_data)
                 if m]
    _latest_actual_fy = max(_fy_years) if _fy_years else None
    _xbrl_paths = (merged_data.get("_meta") or {}).get("xbrl_paths") or []
    forecast_data, _guidance_note, _guidance_source = get_guidance(
        ticker_code,
        min_fy_year=_latest_actual_fy,
        xbrl_paths=_xbrl_paths,
        allow_edinet_tanshin=args.tanshin_fallback,
    )
    if forecast_data:
        print(f"  guidance source: {_guidance_note}")
    else:
        print(f"  NO GUIDANCE - Management scenario falls back to the CAGR estimate.")
        for _line in _guidance_note.split(" | "):
            print(f"    - {_line}")

    # Step 4: Convert to config
    print(f"\n[Step 4/9] Building DCF configuration...")
    config = merged_data_to_config(company_info, merged_data, forecast_data=forecast_data)
    config["_guidance_source"] = _guidance_source
    config["_guidance_note"] = _guidance_note

    # Step 4.5: Apply manual overrides if provided
    if _overrides:
        print(f"\n[Step 4.5] Applying overrides from {args.overrides}...")
        for key, value in _overrides.items():
            if key == "scenarios" and isinstance(value, dict):
                # Deep merge: each scenario individually
                if "scenarios" not in config:
                    config["scenarios"] = {}
                for scen_name, scen_data in value.items():
                    if scen_name in config["scenarios"]:
                        config["scenarios"][scen_name].update(scen_data)
                    else:
                        config["scenarios"][scen_name] = scen_data
            elif _is_placeholder(value):
                print(f"  Skipped unfilled override '{key}' (__CONFIRM__ placeholder)")
                continue
            else:
                config[key] = value

        # Re-apply flat arrays from Base scenario after override
        if "scenarios" in _overrides and "Base" in _overrides["scenarios"]:
            _base = config["scenarios"]["Base"]
            if "revenue_growth" in _base:
                config["revenue_growth"] = _base["revenue_growth"]
            if "cogs_pct" in _base:
                config["cogs_pct"] = _base["cogs_pct"]
            if "sga_pct" in _base:
                config["sga_pct"] = _base["sga_pct"]

        override_keys = list(_overrides.keys())
        config["_override_keys"] = set(override_keys)
        print(f"  Applied {len(override_keys)} override fields: {', '.join(override_keys[:10])}")

        if config.get("nwc_method") == "revenue_pct":
            print(f"  NWC Method: revenue_pct (NWC % of Revenue) - DSO/DIH/DPO will not be used")

        # When segments define Revenue/EBIT, remove cogs_pct so template uses
        # back-calculation: COGS = Revenue - SGA - EBIT, EBIT from Segment Analysis
        if config.get("segments"):
            for sn in config.get("scenarios", {}):
                config["scenarios"][sn].pop("cogs_pct", None)
            config.pop("cogs_pct", None)
            print("  [Segments] Removed cogs_pct - COGS will be back-calculated from segment EBIT")

    # Step 4.55: the two inputs a DCF cannot be built without. Before フェーズ2 #1
    # a missing net_debt became 0 and a missing base-year revenue became JPY 1mn,
    # and the run completed with a workbook that looked finished. Neither can be
    # guessed, so the run stops here and names the override that fixes it.
    _blockers = []
    if config.get("net_debt") is None:
        _blockers.append(
            f"net_debt is unavailable ({config.get('_net_debt_source', '?')}). "
            f"Set \"net_debt\" in data/overrides/{ticker_code}_overrides.json "
            f"from the 短信/有報 balance sheet (interest-bearing debt - cash).")
    if not config.get("base_year_revenue"):
        _blockers.append(
            f"base_year_revenue is unavailable. Set \"base_year_revenue\" in "
            f"data/overrides/{ticker_code}_overrides.json from the latest FY actuals.")
    if _blockers:
        print()
        print("ERROR: the model cannot be built from the data available:")
        for _b in _blockers:
            print(f"  - {_b}")
        sys.exit(2)

    # Step 4.6: Attach the EDINET CF/BS series to the FINAL hist_years by year
    # key (must run after overrides, which may replace hist_years wholesale).
    print(f"\n[Step 4.6] Aligning OCF / Cash / Debt to fiscal-year keys...")
    _override_keys = set(_overrides.keys()) if _overrides else set()
    _fs_coverage, _fs_covered, _fs_n_years = align_hist_series_to_years(config, _override_keys)
    final_warnings = []
    if not forecast_data:
        final_warnings.append(
            "WARNING: no 会社予想 obtained - the Management scenario is the CAGR "
            f"estimate, not company guidance ({_guidance_note})")
    if _fs_covered < _fs_n_years:
        final_warnings.append(
            f"WARNING: OCF/Cash/Debt coverage {_fs_coverage} — verify against 短信"
        )

    # LTM Revenue (C20): report the construction and honour an explicit override.
    if config.get("_ltm_revenue_components"):
        print(f"  [LTM] components: {config['_ltm_revenue_components']}")
    else:
        print(f"  [LTM] source: {config.get('_ltm_revenue_source')} "
              f"(no quarterly decomposition available)")
    if _overrides and _overrides.get("ltm_revenue") is not None:
        _auto = config.get("_ltm_revenue_auto")
        config["ltm_revenue"] = _overrides["ltm_revenue"]
        config["_ltm_revenue_overridden"] = True
        print(f"  [LTM] override applied: {config['ltm_revenue']:,.1f} mn "
              f"(auto-constructed value was {(_auto or 0):,.1f} mn)")
        if _auto:
            _dev = abs(config["ltm_revenue"] - _auto) / abs(_auto)
            if _dev > 0.20:
                _msg = (f"WARNING: ltm_revenue override deviates {_dev:.1%} from the "
                        f"auto-constructed LTM ({config['ltm_revenue']:,.0f} vs "
                        f"{_auto:,.0f} mn) — confirm the scope is intentional")
                print(f"  {_msg}")
                final_warnings.append(_msg)
    print(f"  [LTM] C20 = {config.get('ltm_revenue', 0):,.1f} JPY mn "
          f"({'override' if config.get('_ltm_revenue_overridden') else 'auto'})")

    # Guard: re-calculate core_ebitda if overrides set it to None
    if config.get("core_ebitda") is None:
        _oi = config.get("hist_operating_income", [])
        _da = config.get("hist_depreciation", [])
        if _oi and _da and _oi[-1] is not None and _da[-1] is not None:
            config["core_ebitda"] = _oi[-1] + _da[-1]
            print(f"  [Auto] core_ebitda = {_oi[-1]:,.0f} (OI) + {_da[-1]:,.0f} (D&A) = {config['core_ebitda']:,.0f}")

    # Normalize ticker for yfinance: overrides may carry a bare 4-digit code
    # (e.g. "4192"), which yfinance 404s on. TSE tickers need the ".T" suffix.
    _t = str(config.get("ticker", ticker_code)).strip()
    if "." not in _t and _t[:4].isdigit():
        config["ticker"] = f"{_t[:4]}.T"

    # Step 5: Fetch live market data via yfinance (price, shares, beta)
    print(f"\n[Step 5/9] Fetching market data...")
    ticker_str = config["ticker"]
    _price, _shares, live_beta, _mkt_note = get_live_market_data(ticker_str)
    config["current_price"], config["shares_outstanding"] = _price, _shares
    config["beta"] = live_beta  # RAW; the Blume adjustment + clamp live in the template
    config["_market_data_source"] = _mkt_note

    _from_overrides = []
    # Override shares from overrides["shares"] (single source of truth)
    if _overrides and "shares" in _overrides:
        _fd = _overrides["shares"].get("fully_diluted_shares")
        if _fd:
            config["shares_outstanding"] = _fd
            _from_overrides.append("shares")
            print(f"  Shares override: {_fd:,} (from overrides.shares.fully_diluted_shares)")

    # Re-apply overrides for market data fields (if the analyst has fixed values).
    # This now runs BEFORE the D/E auto-calculation. It used to run after, so a
    # ticker with a current_price override but no de_ratio override had its D/E
    # computed from the LIVE price while the workbook showed the override price
    # — two different market caps inside one model.
    if _overrides:
        for field in ["current_price", "shares_outstanding", "beta"]:
            if field in _overrides:
                if _is_placeholder(_overrides[field]):
                    print(f"  Skipped unfilled override '{field}' (__CONFIRM__ placeholder)")
                    continue
                config[field] = _overrides[field]
                _from_overrides.append(field)
                print(f"  Override applied: {field} = {_overrides[field]}")
    if _from_overrides:
        config["_market_data_source"] = (
            f"{_mkt_note}; from overrides: {', '.join(sorted(set(_from_overrides)))}")

    # ── Market data is not optional and is never guessed ──────────────────
    # The config used to start at price=1,000 / shares=10,000,000 as
    # "placeholders overridden by yfinance", and get_live_market_data() returned
    # them unchanged on any failure. 4568 第一三共 shipped Target JPY 294,427 /
    # BUY +293% on a market cap of JPY 10,000 mn (true: 5,084,100 mn) and passed
    # validation with FAIL 0. Both fields now start as None; if neither yfinance
    # nor the overrides produced a usable number, the run stops.
    LEGACY_PLACEHOLDERS = (1000.0, 10_000_000)
    _p, _s = config.get("current_price"), config.get("shares_outstanding")
    _bad = []
    if not isinstance(_p, (int, float)) or isinstance(_p, bool) or _p <= 0:
        _bad.append(f'current_price is {_p!r} ({_mkt_note})')
    if not isinstance(_s, (int, float)) or isinstance(_s, bool) or _s <= 0:
        _bad.append(f'shares_outstanding is {_s!r} ({_mkt_note})')
    if not _bad and (float(_p), int(_s)) == LEGACY_PLACEHOLDERS:
        _bad.append("current_price 1,000 with shares 10,000,000 - the legacy "
                    "placeholder pair. If these are the real numbers, set them "
                    "explicitly in overrides so the intent is on the record")
    if _bad:
        print()
        print("ERROR: market data could not be established:")
        for _b in _bad:
            print(f"  - {_b}")
        print(f"  Set \"current_price\" and \"shares\": {{\"fully_diluted_shares\": N}} in "
              f"data/overrides/{ticker_code}_overrides.json from the 決算短信.")
        sys.exit(2)

    # Auto-calculate D/E ratio from net_debt and the FINAL market cap
    market_cap = config["current_price"] * config["shares_outstanding"] / 1_000_000  # JPY mn
    if config["net_debt"] > 0 and market_cap > 0:
        config["de_ratio"] = min(round(config["net_debt"] / market_cap, 4), 2.0)
    else:
        config["de_ratio"] = 0.0
    if _overrides and "de_ratio" in _overrides and not _is_placeholder(_overrides["de_ratio"]):
        config["de_ratio"] = _overrides["de_ratio"]
        print(f"  Override applied: de_ratio = {_overrides['de_ratio']}")

    # Beta & Size Premium are normalized inside generate_dcf_workbook (template-level)
    _rb = config.get("beta")
    print(f"  Market data: {config['_market_data_source']}")
    print(f"  Raw Beta: {'n/a' if _rb is None else f'{_rb:.2f}'}, "
          f"D/E Ratio: {config['de_ratio']:.4f}, Mkt Cap: {market_cap:,.0f} mn")

    # Step 5: Load comparable companies data
    print(f"\n[Step 6/9] Loading comparable companies...")
    # Resolve comps CSV path: --comps-csv > data/comps/<ticker>_comps.csv
    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    if args.comps_csv:
        comps_csv_path = args.comps_csv
    else:
        comps_csv_path = os.path.join(project_root, "data", "comps", f"{ticker_code}_comps.csv")

    if args.no_comps:
        print(f"  --no-comps: generating without comparable companies (explicit).")
        config["comps"] = []
    elif os.path.isfile(comps_csv_path):
        # Loading failures are fatal: silently continuing without comps produced
        # workbooks whose Comps-implied values looked valid but meant nothing.
        try:
            config["comps"] = get_comps_data(comps_csv_path)
        except Exception as e:
            print(f"\nERROR: Failed to parse comps CSV {comps_csv_path}: "
                  f"{type(e).__name__}: {e}")
            print(f"  Fix the CSV (see templates/comps_input_template.csv) or pass --no-comps.")
            sys.exit(1)
        if not config["comps"]:
            print(f"\nERROR: {comps_csv_path} parsed to 0 comps. Check the CSV "
                  f"format (see templates/comps_input_template.csv), or pass "
                  f"--no-comps to generate without comps.")
            sys.exit(1)
        print(f"  Loaded {len(config['comps'])} comps from {comps_csv_path}")
        if args.no_peer_check:
            print("  --no-peer-check: peer price freshness check skipped.")
        else:
            check_peer_freshness(config["comps"], config.get("ticker", ticker_code))
    else:
        print(f"\nERROR: No comps CSV found at: {comps_csv_path}")
        txt_sibling = os.path.splitext(comps_csv_path)[0] + ".txt"
        if os.path.isfile(txt_sibling):
            print(f"  Found {txt_sibling} - the pipeline only reads the .csv path. "
                  f"Rename it to {os.path.basename(comps_csv_path)}.")
        print(f"  Create the CSV (data/comps/{ticker_code}_comps.csv), pass --comps-csv PATH,")
        print(f"  or pass --no-comps to explicitly generate without comparable companies.")
        sys.exit(1)

    # Step 6: Generate Excel
    # output_path and the --force guard were resolved at the top of main(),
    # before any network work - see the comment there.
    print(f"\n[Step 7/9] Generating DCF workbook...")

    # Stamp the template revision into the workbook's Adjustments Log so a model
    # can be traced back to the code that produced it.
    try:
        _rev = subprocess.run(["git", "rev-parse", "--short", "HEAD"],
                              cwd=project_root, capture_output=True, text=True,
                              timeout=15)
        config["_template_rev"] = (_rev.stdout or "").strip() or datetime.now().strftime("%Y-%m-%d")
    except Exception:
        config["_template_rev"] = datetime.now().strftime("%Y-%m-%d")

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

    # Echo the effective WACC inputs and comps so the analyst can verify at a
    # glance that the workbook reflects the overrides (no silent defaults).
    print(f"\nEffective WACC inputs (as written to 'DCF Model'!C7:C12):")
    print(f"  Risk-Free:     {config['risk_free']:.3%}")
    print(f"  Beta:          {config['beta']:.2f}")
    print(f"  ERP:           {config['erp']:.3%}")
    print(f"  Size Premium:  {config['size_premium']:.3%}")
    print(f"  Cost of Debt:  {config['cost_of_debt_at']:.3%} (after-tax)")
    print(f"  D/E Ratio:     {config['de_ratio']:.4f}")
    print(f"  => WACC:       {calc_wacc(config):.2%}")
    if config["comps"]:
        comp_names = ", ".join(c["name"] for c in config["comps"])
        print(f"\nComps ({len(config['comps'])}): {comp_names}")
        _excluded = [c["name"] for c in config["comps"] if c.get("exclude_from_stats")]
        if _excluded:
            print(f"  Excluded from statistics (stale/unavailable): {', '.join(_excluded)}")
            final_warnings.append(
                f"WARNING: peers excluded from statistics: {', '.join(_excluded)}"
            )
    else:
        print(f"\nComps: NONE (--no-comps)")

    # ── Step 8: recalculate, then machine-validate the workbook ──
    # A generated workbook has formulas but no computed values, so the checks
    # that need numbers (formula errors, terminal capex ratio, PGM sanity) only
    # mean something after a recalc. Excel COM is best-effort: without it the
    # validator still runs and reports those checks as "needs recalc".
    if not args.no_recalc:
        print(f"\n[Step 8/9] Recalculating via Excel COM...")
        try:
            _rc = subprocess.run(
                [sys.executable, os.path.join(project_root, "scripts", "recalc_excel_com.py"),
                 saved_path],
                capture_output=True, text=True, timeout=600,
            )
            print((_rc.stdout or "").strip() or "  (no output)")
            if _rc.returncode != 0:
                print(f"  WARNING: recalc failed (exit {_rc.returncode}). "
                      f"{(_rc.stderr or '').strip()[:300]}")
                print(f"  Continuing - value-level checks will be reported as "
                      f"'needs recalc'.")
        except Exception as e:
            print(f"  WARNING: recalc skipped ({type(e).__name__}: {e}).")

    validation_failed = False
    if not args.no_validate:
        print(f"\n[Step 9/9] Validating output...")
        try:
            from scripts.validate_output import validate_workbook
        except ImportError:
            from validate_output import validate_workbook
        # --no-recalc leaves the value-level checks unable to run; that is a
        # deliberate partial validation, so SKIP is not turned into a FAIL
        # there. Every normal run recalcs, and there SKIP > 0 does fail.
        result = validate_workbook(saved_path, write_report=True,
                                   allow_skip=args.no_recalc)
        validation_failed = result.failed

    if final_warnings:
        print("\n" + "=" * 60)
        for w in final_warnings:
            print(w)
        print("=" * 60)

    if validation_failed:
        print(f"\nERROR: validate_output.py reported FAIL for {saved_path}.")
        print(f"  The file was NOT deleted - inspect it and the *_validation.txt report.")
        sys.exit(1)


if __name__ == "__main__":
    main()
