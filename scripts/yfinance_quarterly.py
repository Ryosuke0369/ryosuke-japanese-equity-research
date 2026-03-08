"""
yfinance_quarterly.py - Hybrid LTM fallback using yfinance quarterly data.

When EDINET lacks Q1/Q3 data (due to 2024 Financial Instruments Act amendment
making XBRL submission optional), this module fetches quarterly financials from
yfinance and constructs a hybrid LTM to keep DCF models current.

Usage (called from generate_dcf.py):
    merged_data = enrich_merged_data_with_yfinance(merged_data, ticker_str, fiscal_year_end)
"""

import re
from collections import OrderedDict
from datetime import date

import yfinance as yf

# =====================================================================
# yfinance → EDINET key mappings
# =====================================================================

# Income Statement (quarterly_financials)
IS_MAP = {
    "Total Revenue": "revenue",
    "Operating Revenue": "revenue",
    "Cost Of Revenue": "cogs",
    "Selling General And Administration": "sga",
    "Operating Income": "operating_income",
    "Net Income": "net_income",
    "Net Income Common Stockholders": "net_income",
}

# Cash Flow (quarterly_cashflow)
CF_MAP = {
    "Operating Cash Flow": "operating_cf",
    "Capital Expenditure": "capex",
    "Depreciation And Amortization": "depreciation",
}

# Balance Sheet (quarterly_balance_sheet)
BS_MAP = {
    "Cash And Cash Equivalents": "cash",
    "Cash Cash Equivalents And Short Term Investments": "cash",
    "Accounts Receivable": "accounts_receivable",
    "Receivables": "accounts_receivable",
    "Inventory": "inventories",
    "Accounts Payable": "accounts_payable",
    "Current Debt": "short_term_debt",
    "Current Debt And Capital Lease Obligation": "short_term_debt",
    "Long Term Debt": "long_term_debt",
    "Long Term Debt And Capital Lease Obligation": "long_term_debt",
}

# Flow items are summed across quarters; stock items use latest snapshot
FLOW_KEYS = {"revenue", "cogs", "sga", "operating_income", "net_income",
             "operating_cf", "capex", "depreciation"}
STOCK_KEYS = {"cash", "accounts_receivable", "inventories", "accounts_payable",
              "short_term_debt", "long_term_debt"}

# Minimum required keys for a valid LTM
REQUIRED_KEYS = {"revenue", "operating_income"}


# =====================================================================
# 1. detect_ltm_gap
# =====================================================================
def detect_ltm_gap(merged_data, fiscal_year_end_str):
    """Check if merged_data lacks an up-to-date LTM entry.

    Args:
        merged_data: OrderedDict with FY/LTM keys.
        fiscal_year_end_str: e.g. "2025-03-31" or "--03-31".

    Returns:
        dict with keys:
            needs_fallback (bool), existing_q (int|None), expected_q (int|None),
            fy_end_month (int), latest_fy_key (str|None)
    """
    # Parse FY end month
    fy_end_month = _parse_fy_end_month(fiscal_year_end_str)

    # Find existing LTM and latest FY
    ltm_keys = [k for k in merged_data if k.startswith("LTM")]
    fy_keys = [k for k in merged_data if k.startswith("FY")
               and merged_data[k].get("revenue") is not None]
    latest_fy_key = sorted(fy_keys, reverse=True)[0] if fy_keys else None

    existing_q = None
    if ltm_keys:
        # Parse quarter from label like "LTM(2Q 2025-09)"
        m = re.search(r"(\d)Q", ltm_keys[0])
        if m:
            existing_q = int(m.group(1))

    # Expected quarter based on today's date
    expected_q = _expected_quarter(date.today(), fy_end_month)

    needs_fallback = False
    if not ltm_keys:
        needs_fallback = True
    elif existing_q is not None and expected_q is not None:
        # Stale if existing quarter is older than expected
        if existing_q < expected_q:
            needs_fallback = True

    return {
        "needs_fallback": needs_fallback,
        "existing_q": existing_q,
        "expected_q": expected_q,
        "fy_end_month": fy_end_month,
        "latest_fy_key": latest_fy_key,
    }


def _parse_fy_end_month(fiscal_year_end_str):
    """Extract month from fiscal_year_end string. Defaults to 3 (March)."""
    if not fiscal_year_end_str:
        return 3
    # Handle "2025-03-31" or "--03-31"
    m = re.search(r"-(\d{2})-\d{2}$", str(fiscal_year_end_str))
    return int(m.group(1)) if m else 3


def _expected_quarter(today, fy_end_month):
    """Determine the expected latest quarter (1-3) based on a 45-day reporting lag.
    """
    from datetime import timedelta
    # 45 days ago is the effective date we have data for
    effective = today - timedelta(days=45)
    
    # Calculate months since the start of the fiscal year
    # FY start month is fy_end_month + 1 (1-indexed)
    fy_start_month = (fy_end_month % 12) + 1
    
    # How many months passed between fy_start_month and effective date's month?
    if effective.month >= fy_start_month:
        months_elapsed = effective.month - fy_start_month + 1
    else:
        months_elapsed = (12 - fy_start_month + 1) + effective.month
        
    # Example for FY ending in March (fy_start = April):
    # If today is March 8, effective is approx Jan 22.
    # months_elapsed = (12 - 4 + 1) + 1 = 10 months elapsed
    
    # Translate elapsed months to completed expected quarters.
    # 0-2 = no expected quarter (still in Q1 reporting period)
    # 3-5 = Q1 expected
    # 6-8 = Q2 expected
    # 9-11 = Q3 expected
    # 12+ = Q4 (None returned)
    
    if months_elapsed < 3:
        return None
    elif months_elapsed < 6:
        return 1
    elif months_elapsed < 9:
        return 2
    elif months_elapsed < 12:
        return 3
    else:
        return None


def _quarter_end_year(ref_year, q_end_month, fy_end_month):
    """Determine the calendar year of a quarter end within the current FY cycle."""
    if fy_end_month >= q_end_month:
        return ref_year
    return ref_year


# =====================================================================
# 2. fetch_yf_quarterly
# =====================================================================
def fetch_yf_quarterly(ticker_str):
    """Fetch quarterly financials from yfinance.

    Args:
        ticker_str: e.g. "2359.T"

    Returns:
        dict with key 'quarters': list of dicts sorted by date descending,
        each with {date, flow_items, stock_items}.
        Returns None on failure.
    """
    try:
        tkr = yf.Ticker(ticker_str)
    except Exception as e:
        print(f"  WARNING: yfinance Ticker init failed for {ticker_str}: {e}")
        return None

    quarters = {}  # date_str -> {flow_items, stock_items}

    # --- Income Statement ---
    try:
        qf = tkr.quarterly_financials
        if qf is not None and not qf.empty:
            _extract_from_df(qf, IS_MAP, quarters, is_flow=True)
    except Exception as e:
        print(f"  WARNING: yfinance quarterly_financials failed: {e}")

    # --- Cash Flow ---
    try:
        qcf = tkr.quarterly_cashflow
        if qcf is not None and not qcf.empty:
            _extract_from_df(qcf, CF_MAP, quarters, is_flow=True)
    except Exception as e:
        print(f"  WARNING: yfinance quarterly_cashflow failed: {e}")

    # --- Balance Sheet ---
    try:
        qbs = tkr.quarterly_balance_sheet
        if qbs is not None and not qbs.empty:
            _extract_from_df(qbs, BS_MAP, quarters, is_flow=False)
    except Exception as e:
        print(f"  WARNING: yfinance quarterly_balance_sheet failed: {e}")

    if not quarters:
        print(f"  WARNING: No quarterly data found on yfinance for {ticker_str}")
        return None

    # Build sorted list (newest first)
    result = []
    for date_str in sorted(quarters.keys(), reverse=True):
        q = quarters[date_str]
        result.append({
            "date": date_str,
            "flow_items": q.get("flow_items", {}),
            "stock_items": q.get("stock_items", {}),
        })

    print(f"  yfinance: Retrieved {len(result)} quarters for {ticker_str}")
    return {"quarters": result}


def _extract_from_df(df, mapping, quarters, is_flow):
    """Extract mapped values from a yfinance DataFrame into quarters dict.

    Args:
        df: DataFrame with index=item names, columns=dates.
        mapping: dict mapping yfinance row names to EDINET keys.
        quarters: dict to populate, keyed by date string.
        is_flow: True for flow items (IS/CF), False for stock items (BS).
    """
    for col in df.columns:
        date_str = col.strftime("%Y-%m-%d") if hasattr(col, "strftime") else str(col)[:10]

        if date_str not in quarters:
            quarters[date_str] = {"flow_items": {}, "stock_items": {}}

        target = "flow_items" if is_flow else "stock_items"
        seen_keys = set()

        for yf_name, edinet_key in mapping.items():
            if edinet_key in seen_keys:
                continue  # already mapped by primary key
            if yf_name in df.index:
                raw = df.loc[yf_name, col]
                if raw is not None and str(raw) != "nan":
                    val = float(raw)
                    # capex: yfinance reports as negative, we want positive
                    if edinet_key == "capex":
                        val = abs(val)
                    # Convert JPY to JPY millions
                    val = round(val / 1_000_000, 1)
                    quarters[date_str][target][edinet_key] = val
                    seen_keys.add(edinet_key)


# =====================================================================
# 3. compute_hybrid_ltm
# =====================================================================
def compute_hybrid_ltm(merged_data, yf_quarters, gap_info):
    """Compute hybrid LTM from EDINET FY data + yfinance quarterly deltas.

    LTM formula for flow items:
        LTM = EDINET_FY + yf_current_cumulative - yf_prior_cumulative

    Stock items use the latest yfinance quarter snapshot.

    Uses the actual FY end date as a boundary to classify quarters as
    "current FY" (after the FY end date) or "prior FY" (before/at FY end).

    Args:
        merged_data: OrderedDict with FY keys.
        yf_quarters: dict from fetch_yf_quarterly().
        gap_info: dict from detect_ltm_gap().

    Returns:
        (ltm_dict, ltm_label) or (None, None) if construction fails.
    """
    import calendar

    latest_fy_key = gap_info["latest_fy_key"]
    if not latest_fy_key or latest_fy_key not in merged_data:
        print("  WARNING: No FY data available for hybrid LTM base.")
        return None, None

    fy_end_month = gap_info["fy_end_month"]
    existing_q = gap_info.get("existing_q")
    fy_data = merged_data[latest_fy_key]

    # Parse FY year from key like "FY2025"
    fy_year_match = re.search(r"FY(\d{4})", latest_fy_key)
    if not fy_year_match:
        return None, None
    latest_fy_year = int(fy_year_match.group(1))

    # Compute the actual FY end date (e.g., FY2025 with March end → 2025-03-31)
    last_day = calendar.monthrange(latest_fy_year, fy_end_month)[1]
    fy_end_date = date(latest_fy_year, fy_end_month, last_day)

    # Also compute the prior FY end date for bounding prior quarters
    prior_fy_year = latest_fy_year - 1
    prior_last_day = calendar.monthrange(prior_fy_year, fy_end_month)[1]
    prior_fy_end_date = date(prior_fy_year, fy_end_month, prior_last_day)

    today = date.today()

    # Collect and classify quarters using actual FY end date as boundary
    current_quarters = []   # quarters after fy_end_date (current FY)
    prior_quarters = []     # quarters between prior_fy_end_date and fy_end_date

    for q in yf_quarters["quarters"]:
        q_date = _parse_date(q["date"])
        if q_date is None:
            continue

        if q_date > fy_end_date:
            # Current FY: skip if more than 18 months old (staleness guard)
            if (today - q_date).days > 548:
                continue
            current_quarters.append((q_date, q))
        elif q_date > prior_fy_end_date:
            # Prior FY: always include (needed for LTM delta calculation)
            prior_quarters.append((q_date, q))

    # Sort current ascending (Q1, Q2, Q3), prior ascending too
    current_quarters.sort(key=lambda x: x[0])
    prior_quarters.sort(key=lambda x: x[0])

    # Only count current quarters that have at least some flow data
    current_with_flows = [(d, q) for d, q in current_quarters if q["flow_items"]]
    current_q_count = len(current_with_flows) if current_with_flows else len(current_quarters)

    if current_q_count == 0:
        print("  WARNING: No current-FY quarters found on yfinance. Cannot build LTM.")
        return None, None

    # Quality gate: only proceed if yfinance provides MORE quarters than existing EDINET LTM
    if existing_q is not None and current_q_count <= existing_q:
        print(f"  yfinance has {current_q_count}Q but EDINET already has {existing_q}Q. "
              f"Keeping EDINET LTM.")
        return None, None

    # Check that we actually have flow data (not just balance sheet)
    has_flow_data = any(q["flow_items"] for _, q in current_quarters)
    if not has_flow_data:
        print("  WARNING: yfinance quarters have no income/cashflow data. "
              "Cannot build meaningful hybrid LTM.")
        return None, None

    # Extract flow/stock data
    current_fy_flows = [q["flow_items"] for _, q in current_quarters]
    prior_fy_flows = [q["flow_items"] for _, q in prior_quarters]

    # Track latest date and stock items
    latest_yf_date = current_quarters[-1][0]  # last (newest) current quarter
    latest_stock_items = current_quarters[-1][1]["stock_items"]

    # Compute cumulative flow sums
    current_cumulative = _sum_flow_dicts(current_fy_flows)
    prior_cumulative = _sum_flow_dicts(prior_fy_flows[:current_q_count])  # match same Q count

    # Build LTM
    ltm = {}

    # Flow items: LTM = FY + current_cumulative - prior_cumulative
    for key in FLOW_KEYS:
        fy_val = fy_data.get(key)
        cur_val = current_cumulative.get(key)
        pri_val = prior_cumulative.get(key)

        if fy_val is not None and cur_val is not None:
            pri = pri_val if pri_val is not None else 0
            ltm[key] = round(fy_val + cur_val - pri, 1)
        elif fy_val is not None:
            ltm[key] = fy_val

    # Stock items: latest yfinance snapshot, fall back to FY
    for key in STOCK_KEYS:
        if key in latest_stock_items:
            ltm[key] = latest_stock_items[key]
        elif fy_data.get(key) is not None:
            ltm[key] = fy_data[key]

    # Derived calculations
    st_debt = ltm.get("short_term_debt", 0) or 0
    lt_debt = ltm.get("long_term_debt", 0) or 0
    cash = ltm.get("cash", 0) or 0
    ltm["total_debt"] = round(st_debt + lt_debt, 1)
    ltm["net_debt"] = round(st_debt + lt_debt - cash, 1)

    # Validate minimum required keys
    if not all(ltm.get(k) for k in REQUIRED_KEYS):
        print("  WARNING: Hybrid LTM missing required keys (revenue/operating_income). Skipping.")
        return None, None

    # Generate label: "LTM(2Q 2025-09)(yf)"
    date_str = latest_yf_date.strftime("%Y-%m") if latest_yf_date else "unknown"
    ltm_label = f"LTM({current_q_count}Q {date_str})(yf)"

    return ltm, ltm_label


def _parse_date(date_str):
    """Parse YYYY-MM-DD string to date object."""
    try:
        parts = date_str.split("-")
        return date(int(parts[0]), int(parts[1]), int(parts[2]))
    except (ValueError, IndexError):
        return None


def _assign_fy_year(quarter_end_date, fy_end_month):
    """Assign a fiscal year to a quarter end date.

    Convention: FY is labeled by the year in which the FY ends.
    e.g., for March FY-end: 2025-06-30 → FY2026, 2025-03-31 → FY2025.
    """
    if fy_end_month == 12:
        return quarter_end_date.year
    if quarter_end_date.month > fy_end_month:
        return quarter_end_date.year + 1
    return quarter_end_date.year


def _sum_flow_dicts(dicts):
    """Sum flow items across multiple quarter dicts."""
    result = {}
    for d in dicts:
        for key, val in d.items():
            if key in FLOW_KEYS and val is not None:
                result[key] = round(result.get(key, 0) + val, 1)
    return result


# =====================================================================
# 4. enrich_merged_data_with_yfinance (top-level orchestrator)
# =====================================================================
def enrich_merged_data_with_yfinance(merged_data, ticker_str, fiscal_year_end):
    """Enrich merged_data with yfinance-based hybrid LTM if needed.

    This is the only function called from generate_dcf.py.
    Graceful degradation: on any failure, returns merged_data unchanged.

    Args:
        merged_data: OrderedDict from fetch_and_parse_multi_year().
        ticker_str: e.g. "2359.T"
        fiscal_year_end: e.g. "2025-03-31" or "--03-31"

    Returns:
        OrderedDict: merged_data, potentially with a new LTM key inserted.
    """
    # Step 1: Detect gap
    gap_info = detect_ltm_gap(merged_data, fiscal_year_end)

    if not gap_info["needs_fallback"]:
        existing_q = gap_info["existing_q"]
        print(f"  LTM data is current (Q{existing_q}). No yfinance fallback needed.")
        return merged_data

    print(f"  LTM gap detected (existing: Q{gap_info['existing_q']}, "
          f"expected: Q{gap_info['expected_q']}). Attempting yfinance fallback...")

    # Step 2: Fetch yfinance quarterly data
    yf_data = fetch_yf_quarterly(ticker_str)
    if yf_data is None:
        print("  WARNING: yfinance fallback failed. Proceeding without LTM update.")
        return merged_data

    # Step 3: Compute hybrid LTM
    ltm, ltm_label = compute_hybrid_ltm(merged_data, yf_data, gap_info)
    if ltm is None:
        print("  WARNING: Hybrid LTM computation failed. Proceeding without LTM update.")
        return merged_data

    # Step 4: Insert LTM into merged_data (at the front)
    # Remove any existing LTM keys first
    new_data = OrderedDict()
    new_data[ltm_label] = ltm
    for k, v in merged_data.items():
        if not k.startswith("LTM"):
            new_data[k] = v

    print(f"  Hybrid LTM inserted: {ltm_label}")
    print(f"    Revenue: {ltm.get('revenue', 'N/A'):,.1f} mn")
    print(f"    Op Income: {ltm.get('operating_income', 'N/A'):,.1f} mn")

    return new_data
