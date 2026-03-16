"""
edinet_parser.py - Parse EDINET XBRL instance files to extract financial data.

Reads a .xbrl file downloaded via edinet_fetcher.py, identifies the correct
consolidated contexts, and extracts key financial metrics (PL, BS, CF) needed
for DCF modeling.

Usage:
    python scripts/edinet_parser.py path/to/file.xbrl

Output values are in JPY millions (百万円) by default.
"""

import os
import sys
import logging
from collections import OrderedDict

from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

# =====================================================================
# CONSTANTS
# =====================================================================
# Scale factor: XBRL values are in JPY (円), our models use JPY mn (百万円)
SCALE_TO_MN = 1_000_000

# Context ID naming patterns used in EDINET XBRL
# "Clean" contexts have no scenario/dimension members — these contain
# top-level consolidated (or non-consolidated) totals.
CONTEXT_PATTERNS = {
    "current_duration":  "CurrentYearDuration",
    "current_instant":   "CurrentYearInstant",
    "prior1_duration":   "Prior1YearDuration",
    "prior1_instant":    "Prior1YearInstant",
    "prior2_duration":   "Prior2YearDuration",
    "prior2_instant":    "Prior2YearInstant",
    "prior3_duration":   "Prior3YearDuration",
    "prior3_instant":    "Prior3YearInstant",
    "prior4_duration":   "Prior4YearDuration",
    "prior4_instant":    "Prior4YearInstant",
}

# =====================================================================
# FINANCIAL ITEM DEFINITIONS
#
# Each item: (display_name, type, fallback_tags, is_sum)
#   type: "duration" (PL/CF) or "instant" (BS)
#   fallback_tags: list of XBRL element local names to try, in priority order
#   is_sum: if True, sum ALL matching tags (for multi-component items)
# =====================================================================
FINANCIAL_ITEMS = OrderedDict([
    # ── Income Statement (Duration) ──
    ("revenue", {
        "label": "Revenue (売上高)",
        "type": "duration",
        "tags": [
            "NetSales",
            "OperatingRevenue1",
            "Revenue",
            "OperatingRevenue",
            "NetSalesOfCompletedConstructionContracts",
            "OrdinaryIncomeBNK",
        ],
        "sum": False,
    }),
    ("cogs", {
        "label": "Cost of Sales (売上原価)",
        "type": "duration",
        "tags": [
            "CostOfSales",
            "CostOfProductsManufactured",
            "CostOfSalesOfCompletedConstructionContracts",
            "OperatingExpenses",
        ],
        "sum": False,
    }),
    ("sga", {
        "label": "SGA Expenses (販管費)",
        "type": "duration",
        "tags": [
            "SellingGeneralAndAdministrativeExpenses",
        ],
        "sum": False,
    }),
    ("operating_income", {
        "label": "Operating Income (営業利益)",
        "type": "duration",
        "tags": [
            "OperatingIncome",
            "OperatingProfit",
        ],
        "sum": False,
    }),
    ("net_income", {
        "label": "Net Income (当期純利益)",
        "type": "duration",
        "tags": [
            "ProfitLossAttributableToOwnersOfParent",
            "ProfitLoss",
            "NetIncome",
        ],
        "sum": False,
    }),

    # ── Balance Sheet (Instant) ──
    ("cash", {
        "label": "Cash & Deposits (現金及び預金)",
        "type": "instant",
        "tags": [
            "CashAndDeposits",
            "CashAndCashEquivalents",
        ],
        "sum": False,
    }),
    ("accounts_receivable", {
        "label": "Accounts Receivable (売上債権)",
        "type": "instant",
        "tags": [
            # Aggregate tags (excluding contract assets for cleaner DSO)
            "NotesAndAccountsReceivableTrade",
            "NotesAndOperatingAccountsReceivableTrade",
            # Fallback: sum individual trade receivable components
            # (ContractAssets excluded — it inflates DSO for progress-billing companies)
            "AccountsReceivableTrade",
            "NotesReceivableTrade",
            "ElectronicallyRecordedMonetaryClaimsOperatingCA",
            "ElectronicallyRecordedMonetaryClaimsOperatingAccounts",
        ],
        "sum": "fallback",  # try single tags first, then sum components
    }),
    ("inventories", {
        "label": "Inventories (棚卸資産)",
        "type": "instant",
        "tags": [
            "Inventories",
            "Inventory",
            # Fallback: sum individual components
            "MerchandiseAndFinishedGoods",
            "WorkInProcess",
            "WorkInProcessInventory",
            "RawMaterialsAndSupplies",
            "RawMaterials",
            "Merchandise",
            "FinishedGoods",
        ],
        "sum": "fallback",
    }),
    ("accounts_payable", {
        "label": "Accounts Payable (買掛金)",
        "type": "instant",
        "tags": [
            "NotesAndAccountsPayableTrade",
            "NotesAndOperatingAccountsPayableTrade",
            "AccountsPayableTrade",
            "NotesPayableTrade",
        ],
        "sum": "fallback",
    }),
    ("short_term_debt", {
        "label": "Short-term Debt (短期借入金)",
        "type": "instant",
        "tags": [
            "ShortTermLoansPayable",
            "ShortTermBorrowings",
            "CurrentPortionOfLongTermLoansPayable",
        ],
        "sum": True,
    }),
    ("long_term_debt", {
        "label": "Long-term Debt (長期借入金)",
        "type": "instant",
        "tags": [
            "LongTermLoansPayable",
            "LongTermDebt",
            "BondsPayable",
        ],
        "sum": True,
    }),

    # ── Cash Flow Statement (Duration) ──
    ("depreciation", {
        "label": "Depreciation & Amortization (減価償却費)",
        "type": "duration",
        "tags": [
            # J-GAAP standard
            "DepreciationAndAmortizationOpeCF",
            "DepreciationAndAmortization",
            "DepreciationSGA",
            # IFRS Taxonomy
            "DepreciationAmortisationAndImpairmentLossReversalOfImpairmentLossRecognisedInProfitOrLoss",
            "DepreciationAndAmortisationExpense",
            "DepreciationExpense",
            "AmortisationExpense",
            # Additional J-GAAP variants
            "DepreciationCOS",
            "DepreciationAndAmortizationNOPE",
            "Depreciation",
        ],
        "sum": False,
    }),
    ("operating_cf", {
        "label": "Operating Cash Flow (営業CF)",
        "type": "duration",
        "tags": [
            "NetCashProvidedByUsedInOperatingActivities",
        ],
        "sum": False,
    }),
    ("capex", {
        "label": "Capital Expenditure (設備投資)",
        "type": "duration",
        "tags": [
            # J-GAAP standard
            "PurchaseOfPropertyPlantAndEquipmentAndIntangibleAssets",
            "PurchaseOfPropertyPlantAndEquipment",
            "PurchaseOfPropertyPlantAndEquipmentInvCF",
            "PurchaseOfNoncurrentAssetsInvCF",
            # IFRS Taxonomy
            "PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities",
            "AcquisitionsOfPropertyPlantAndEquipment",
            # Additional J-GAAP variants
            "PurchaseOfPropertyPlantAndEquipmentAndOtherAssets",
            "PaymentsForPurchaseOfPropertyPlantAndEquipment",
            "PurchaseOfTangibleAndIntangibleAssets",
            "PurchaseOfFixedAssetsInvCF",
            "IncreaseInPropertyPlantAndEquipmentAndIntangibleAssets",
            # Broader fallbacks
            "PurchaseOfNonCurrentAssets",
            "PaymentsToAcquirePropertyPlantAndEquipment",
            "AdditionsToNoncurrentAssetsOtherThanFinancialInstrumentsDeferredTaxAssetsDefinedBenefitAssetsRightsArisingUnderInsuranceContractsAndRightsArisingUnderReinsuranceContracts",
        ],
        "sum": False,
        "negate": True,
    }),
])

# Company guidance / forecast items (決算短信の業績予想)
FORECAST_ITEMS = {
    "forecast_revenue": {
        "label": "Forecast Revenue (売上高予想)",
        "tags": [
            "ForecastRevenueOperatingRevenue1",
            "ForecastNetSales",
            "ForecastRevenue",
        ],
    },
    "forecast_operating_income": {
        "label": "Forecast Operating Income (営業利益予想)",
        "tags": [
            "ForecastOperatingIncome",
            "ForecastOperatingProfit",
        ],
    },
    "forecast_net_income": {
        "label": "Forecast Net Income (当期純利益予想)",
        "tags": [
            "ForecastProfitLossAttributableToOwnersOfParent",
            "ForecastNetIncome",
            "ForecastProfitLoss",
        ],
    },
}

# Components that should be summed when a single aggregate tag is not found
_SUM_COMPONENTS = {
    "accounts_receivable": [
        "AccountsReceivableTrade",
        "NotesReceivableTrade",
        "ElectronicallyRecordedMonetaryClaimsOperatingCA",
        "ElectronicallyRecordedMonetaryClaimsOperatingAccounts",
    ],
    "inventories": [
        "MerchandiseAndFinishedGoods",
        "WorkInProcess",
        "WorkInProcessInventory",
        "RawMaterialsAndSupplies",
        "RawMaterials",
        "Merchandise",
        "FinishedGoods",
    ],
    "accounts_payable": [
        "AccountsPayableTrade",
        "NotesPayableTrade",
    ],
}


# =====================================================================
# CORE FUNCTIONS
# =====================================================================
def parse_xbrl_file(xbrl_file_path):
    """Parse an XBRL instance file and return a BeautifulSoup object.

    Args:
        xbrl_file_path: Absolute or relative path to a .xbrl file.

    Returns:
        BeautifulSoup object parsed with lxml-xml parser.
    """
    logger.info("Parsing XBRL file: %s", xbrl_file_path)

    with open(xbrl_file_path, "r", encoding="utf-8") as f:
        content = f.read()

    soup = BeautifulSoup(content, "lxml-xml")
    return soup


def identify_clean_contexts(soup):
    """Identify 'clean' context IDs — those without scenario/dimension members.

    In EDINET XBRL, clean contexts (no xbrli:scenario element) represent
    top-level consolidated financial totals. Contexts with dimension members
    (e.g., equity components, segments) should be excluded.

    For non-consolidated companies (単体), all data lives under
    NonConsolidatedMember contexts. If clean contexts yield no revenue data,
    we fall back to these.

    Args:
        soup: BeautifulSoup object of parsed XBRL.

    Returns:
        dict mapping context pattern keys to their clean context IDs.
        Example: {"current_duration": "CurrentYearDuration", ...}
    """
    clean_contexts = {}

    for ctx in soup.find_all("xbrli:context"):
        ctx_id = ctx.get("id", "")
        # A context is "clean" if it has no scenario element (no dimensions)
        scenario = ctx.find("xbrli:scenario")
        if scenario is not None:
            continue

        for key, pattern in CONTEXT_PATTERNS.items():
            if ctx_id == pattern:
                clean_contexts[key] = ctx_id
                period = ctx.find("xbrli:period")
                if period:
                    instant = period.find("xbrli:instant")
                    start = period.find("xbrli:startDate")
                    end = period.find("xbrli:endDate")
                    if instant:
                        logger.info("  Context %-25s -> %s", ctx_id, instant.text)
                    elif start and end:
                        logger.info("  Context %-25s -> %s to %s", ctx_id, start.text, end.text)

    # --- Fallback for non-consolidated (単体) companies ---
    # If clean contexts found no revenue data, try NonConsolidatedMember contexts
    if clean_contexts:
        has_revenue = False
        revenue_tags = FINANCIAL_ITEMS["revenue"]["tags"]
        dur_ctx = clean_contexts.get("current_duration")
        if dur_ctx:
            for tag in revenue_tags:
                if _get_value(soup, tag, dur_ctx) is not None:
                    has_revenue = True
                    break
        if has_revenue:
            return clean_contexts

    # Look for NonConsolidatedMember contexts as fallback
    noncon_contexts = {}
    for ctx in soup.find_all("xbrli:context"):
        ctx_id = ctx.get("id", "")
        scenario = ctx.find("xbrli:scenario")
        if scenario is None:
            continue
        # Check if this is a pure NonConsolidatedMember context
        # (single dimension: ConsolidatedOrNonConsolidatedAxis:NonConsolidatedMember)
        if "NonConsolidatedMember" not in ctx_id:
            continue

        for key, pattern in CONTEXT_PATTERNS.items():
            expected_id = f"{pattern}_NonConsolidatedMember"
            if ctx_id == expected_id:
                noncon_contexts[key] = ctx_id
                period = ctx.find("xbrli:period")
                if period:
                    instant = period.find("xbrli:instant")
                    start = period.find("xbrli:startDate")
                    end = period.find("xbrli:endDate")
                    if instant:
                        logger.info("  NonCon Context %-25s -> %s", ctx_id, instant.text)
                    elif start and end:
                        logger.info("  NonCon Context %-25s -> %s to %s", ctx_id, start.text, end.text)

    if noncon_contexts:
        logger.info("Using NonConsolidatedMember contexts (単体 company fallback)")
        return noncon_contexts

    return clean_contexts


def _get_value(soup, tag_local_name, context_id):
    """Extract a numeric value for a given tag and context from the XBRL soup.

    Searches for elements with the given local name (ignoring namespace prefix)
    that have contextRef matching the specified context ID.

    Args:
        soup: BeautifulSoup object.
        tag_local_name: Local element name (e.g. "NetSales").
        context_id: The context reference string (e.g. "CurrentYearDuration").

    Returns:
        float value in JPY, or None if not found.
    """
    # Search across all namespace prefixes (jppfs_cor, jpcrp_cor, etc.)
    for el in soup.find_all(True):
        local = el.name.split(":")[-1] if ":" in el.name else el.name
        if local != tag_local_name:
            continue
        if el.get("contextRef") != context_id:
            continue
        # Skip nil values
        if el.get("{http://www.w3.org/2001/XMLSchema-instance}nil") == "true":
            continue
        if el.get("xsi:nil") == "true":
            continue
        try:
            text = el.text.strip()
            if not text:
                continue
            return float(text)
        except (ValueError, TypeError):
            continue

    return None


def extract_item(soup, item_key, item_def, context_id, scale=SCALE_TO_MN):
    """Extract a single financial item from XBRL.

    Implements the fallback strategy:
    1. For sum=False: try each tag in order, return first match.
    2. For sum=True: sum all matching tags.
    3. For sum="fallback": try single aggregate tags first, then sum components.

    Args:
        soup: BeautifulSoup object.
        item_key: Key in FINANCIAL_ITEMS (e.g. "revenue").
        item_def: Item definition dict from FINANCIAL_ITEMS.
        context_id: The XBRL context ID to search in.
        scale: Divisor for unit conversion (default: 1_000_000 for JPY mn).

    Returns:
        Scaled float value, or None if not found.
    """
    tags = item_def["tags"]
    sum_mode = item_def.get("sum", False)
    negate = item_def.get("negate", False)

    if sum_mode == "fallback":
        # Strategy: try aggregate tags first (before the component tags)
        components = _SUM_COMPONENTS.get(item_key, [])
        aggregate_tags = [t for t in tags if t not in components]

        # Try aggregate tags first
        for tag in aggregate_tags:
            val = _get_value(soup, tag, context_id)
            if val is not None:
                result = val / scale if scale else val
                return abs(result) if negate else result

        # Fallback: sum components
        total = 0.0
        found_any = False
        for tag in components:
            val = _get_value(soup, tag, context_id)
            if val is not None:
                total += val
                found_any = True

        if found_any:
            result = total / scale if scale else total
            return abs(result) if negate else result

        return None

    elif sum_mode:
        # Sum all matching tags
        total = 0.0
        found_any = False
        for tag in tags:
            val = _get_value(soup, tag, context_id)
            if val is not None:
                total += val
                found_any = True
        if found_any:
            result = total / scale if scale else total
            return abs(result) if negate else result
        return None

    else:
        # Try each tag in order, return first match
        for tag in tags:
            val = _get_value(soup, tag, context_id)
            if val is not None:
                result = val / scale if scale else val
                return abs(result) if negate else result
        return None


def extract_financial_data(soup, contexts, years=None, scale=SCALE_TO_MN):
    """Extract all financial data for specified fiscal years.

    Args:
        soup: BeautifulSoup object of parsed XBRL.
        contexts: dict from identify_clean_contexts().
        years: list of year keys to extract, e.g. ["current", "prior1"].
                Defaults to all available years.
        scale: Unit divisor (default: 1_000_000 for JPY mn).

    Returns:
        dict: {year_key: {item_key: value_or_None, ...}, ...}
        Also includes metadata under "_meta" key with period dates.
    """
    if years is None:
        # Auto-detect available years from contexts
        year_prefixes = set()
        for key in contexts:
            prefix = key.rsplit("_", 1)[0]  # "current", "prior1", etc.
            year_prefixes.add(prefix)
        years = sorted(year_prefixes, key=lambda x: (x != "current", x))

    result = {}

    for year_key in years:
        dur_ctx = contexts.get(f"{year_key}_duration")
        inst_ctx = contexts.get(f"{year_key}_instant")

        if not dur_ctx and not inst_ctx:
            logger.warning("No contexts found for year '%s', skipping.", year_key)
            continue

        year_data = {}
        for item_key, item_def in FINANCIAL_ITEMS.items():
            ctx_id = dur_ctx if item_def["type"] == "duration" else inst_ctx
            if ctx_id is None:
                year_data[item_key] = None
                continue
            year_data[item_key] = extract_item(soup, item_key, item_def, ctx_id, scale)

        # Derive computed fields
        if year_data.get("short_term_debt") is not None or year_data.get("long_term_debt") is not None:
            st = year_data.get("short_term_debt") or 0
            lt = year_data.get("long_term_debt") or 0
            year_data["total_debt"] = round(st + lt, 1)
        else:
            year_data["total_debt"] = None

        if year_data.get("total_debt") is not None and year_data.get("cash") is not None:
            year_data["net_debt"] = round(year_data["total_debt"] - year_data["cash"], 1)
        else:
            year_data["net_debt"] = None

        result[year_key] = year_data

    # Add metadata: period dates from contexts
    meta = {}
    for key, ctx_id in contexts.items():
        ctx_el = soup.find("xbrli:context", id=ctx_id)
        if ctx_el:
            period = ctx_el.find("xbrli:period")
            if period:
                instant = period.find("xbrli:instant")
                start = period.find("xbrli:startDate")
                end = period.find("xbrli:endDate")
                if instant:
                    meta[key] = {"instant": instant.text}
                elif start and end:
                    meta[key] = {"start": start.text, "end": end.text}
    result["_meta"] = meta

    return result


def extract_company_info(soup):
    """Extract company identification from XBRL DEI (Document & Entity Information).

    Args:
        soup: BeautifulSoup object.

    Returns:
        dict with company_name, edinet_code, securities_code, fiscal_year_end.
    """
    info = {}

    dei_items = {
        "company_name": ["FilerNameInJapaneseDEI"],
        "edinet_code": ["EDINETCodeDEI"],
        "securities_code": ["SecurityCodeDEI"],
        "fiscal_year_end": ["CurrentFiscalYearEndDateDEI"],
        "current_period_end": ["CurrentPeriodEndDateDEI"],
    }

    for key, tags in dei_items.items():
        for tag in tags:
            el = soup.find(True, {"name": lambda n: n and n.endswith(f":{tag}")})
            if el is None:
                # Try without namespace
                for candidate in soup.find_all(True):
                    local = candidate.name.split(":")[-1] if ":" in candidate.name else candidate.name
                    if local == tag:
                        el = candidate
                        break
            if el and el.text.strip():
                info[key] = el.text.strip()
                break

    return info


# =====================================================================
# FORECAST / GUIDANCE EXTRACTION (決算短信・業績予想)
# =====================================================================
def extract_forecast_data(soup, scale=SCALE_TO_MN):
    """Extract company guidance/forecast data from XBRL.

    Forecast data appears in contexts with ForecastMember scenario dimension,
    commonly found in 決算短信 (earnings summaries) and annual reports.

    Args:
        soup: BeautifulSoup object of parsed XBRL.
        scale: Unit divisor (default: 1_000_000 for JPY mn).

    Returns:
        dict: {item_key: value_in_jpy_mn, ...} or empty dict if no forecasts found.
              Keys: forecast_revenue, forecast_operating_income, forecast_net_income.
    """
    # Find contexts with ForecastMember in scenario dimension
    forecast_ctx_ids = []
    for ctx in soup.find_all("xbrli:context"):
        ctx_id = ctx.get("id", "")
        scenario = ctx.find("xbrli:scenario")
        if scenario is None:
            continue
        # Check for ForecastMember in scenario (covers both explicit dimension
        # members and context ID naming conventions)
        scenario_str = str(scenario)
        if "ForecastMember" in scenario_str or "ResultForecastMember" in scenario_str:
            # Prefer CurrentYearDuration forecast (next FY forecast)
            forecast_ctx_ids.append(ctx_id)

    if not forecast_ctx_ids:
        logger.info("No forecast contexts (ForecastMember) found in XBRL.")
        return {}

    # Prioritize: CurrentYearDuration forecasts first, then others
    forecast_ctx_ids.sort(key=lambda x: (
        "CurrentYear" not in x,
        "NextYear" not in x,
        x,
    ))
    logger.info("Found %d forecast context(s): %s",
                len(forecast_ctx_ids), forecast_ctx_ids[:5])

    result = {}
    for item_key, item_def in FORECAST_ITEMS.items():
        for ctx_id in forecast_ctx_ids:
            for tag in item_def["tags"]:
                val = _get_value(soup, tag, ctx_id)
                if val is not None:
                    result[item_key] = val / scale if scale else val
                    logger.info("  Forecast: %s = %.1f mn (tag=%s, ctx=%s)",
                                item_key, result[item_key], tag, ctx_id)
                    break
            if item_key in result:
                break

    return result


# =====================================================================
# QUARTERLY REPORT PARSING & LTM CALCULATION
# =====================================================================

# Quarterly context patterns (clean, no dimensions)
# Q3 example: CurrentAccumulatedQ3Duration, Prior1AccumulatedQ3Duration
# Q2 example: CurrentAccumulatedQ2Duration, Prior1AccumulatedQ2Duration
# Q1 example: CurrentAccumulatedQ1Duration, Prior1AccumulatedQ1Duration
# BS: CurrentQuarterInstant
QUARTERLY_CONTEXT_PREFIXES = {
    "current_accumulated": "CurrentAccumulatedQ",
    "prior1_accumulated": "Prior1AccumulatedQ",
    "current_quarter_instant": "CurrentQuarterInstant",
    "prior1_quarter_instant": "Prior1QuarterInstant",
}


def identify_quarterly_contexts(soup):
    """Identify clean quarterly context IDs from a quarterly XBRL file.

    Detects which quarter (Q1/Q2/Q3) by scanning for AccumulatedQ{n} contexts.

    Returns:
        dict with keys like:
            'current_accumulated_duration': context ID for current cumulative period
            'prior1_accumulated_duration': context ID for prior year same cumulative
            'current_quarter_instant': context ID for current quarter-end BS
            'quarter_number': int (1, 2, or 3)
            'period_end': quarter end date string
        Returns empty dict if no quarterly contexts found.
    """
    result = {}
    quarter_number = None

    for ctx in soup.find_all("xbrli:context"):
        ctx_id = ctx.get("id", "")
        scenario = ctx.find("xbrli:scenario")
        if scenario is not None:
            continue  # skip dimensioned contexts

        # Detect cumulative duration contexts
        for q in [3, 2, 1]:
            current_pattern = f"CurrentAccumulatedQ{q}Duration"
            prior_pattern = f"Prior1AccumulatedQ{q}Duration"

            if ctx_id == current_pattern:
                result["current_accumulated_duration"] = ctx_id
                quarter_number = q
                period = ctx.find("xbrli:period")
                if period:
                    end = period.find("xbrli:endDate")
                    if end:
                        result["period_end"] = end.text
                    start = period.find("xbrli:startDate")
                    if start and end:
                        logger.info("  Q%d Context %-35s -> %s to %s",
                                    q, ctx_id, start.text, end.text)

            if ctx_id == prior_pattern:
                result["prior1_accumulated_duration"] = ctx_id
                period = ctx.find("xbrli:period")
                if period:
                    start = period.find("xbrli:startDate")
                    end = period.find("xbrli:endDate")
                    if start and end:
                        logger.info("  Q%d Context %-35s -> %s to %s",
                                    q, ctx_id, start.text, end.text)

        # Detect quarterly instant contexts (BS)
        if ctx_id == "CurrentQuarterInstant":
            result["current_quarter_instant"] = ctx_id
            period = ctx.find("xbrli:period")
            if period:
                instant = period.find("xbrli:instant")
                if instant:
                    logger.info("  Context %-35s -> %s", ctx_id, instant.text)

        if ctx_id == "Prior1QuarterInstant":
            result["prior1_quarter_instant"] = ctx_id

        # --- New semi-annual report (docType 160) context patterns ---
        # InterimDuration = current H1 cumulative (equivalent to AccumulatedQ2)
        if ctx_id == "InterimDuration":
            result["current_accumulated_duration"] = ctx_id
            if quarter_number is None:
                quarter_number = 2  # semi-annual = Q2
            period = ctx.find("xbrli:period")
            if period:
                end = period.find("xbrli:endDate")
                if end:
                    result["period_end"] = end.text
                start = period.find("xbrli:startDate")
                if start and end:
                    logger.info("  H1 Context %-35s -> %s to %s",
                                ctx_id, start.text, end.text)

        if ctx_id == "Prior1InterimDuration":
            result["prior1_accumulated_duration"] = ctx_id
            period = ctx.find("xbrli:period")
            if period:
                start = period.find("xbrli:startDate")
                end = period.find("xbrli:endDate")
                if start and end:
                    logger.info("  H1 Context %-35s -> %s to %s",
                                ctx_id, start.text, end.text)

        if ctx_id == "InterimInstant":
            result["current_quarter_instant"] = ctx_id
            period = ctx.find("xbrli:period")
            if period:
                instant = period.find("xbrli:instant")
                if instant:
                    logger.info("  Context %-35s -> %s", ctx_id, instant.text)

        if ctx_id == "Prior1InterimInstant":
            result["prior1_quarter_instant"] = ctx_id

    if quarter_number:
        result["quarter_number"] = quarter_number

    # --- Fallback for non-consolidated (単体) companies ---
    # If clean quarterly contexts have no revenue data, try NonConsolidatedMember variants
    _need_noncon_fallback = not result.get("current_accumulated_duration")
    if not _need_noncon_fallback and result.get("current_accumulated_duration"):
        _has_q_revenue = False
        for tag in FINANCIAL_ITEMS["revenue"]["tags"]:
            if _get_value(soup, tag, result["current_accumulated_duration"]) is not None:
                _has_q_revenue = True
                break
        if not _has_q_revenue:
            _need_noncon_fallback = True
            # Clear the empty clean contexts so NonCon can replace them
            result = {k: v for k, v in result.items()
                      if k in ("quarter_number", "period_end")}

    if _need_noncon_fallback:
        for ctx in soup.find_all("xbrli:context"):
            ctx_id = ctx.get("id", "")
            if "NonConsolidatedMember" not in ctx_id:
                continue

            for q in [3, 2, 1]:
                cur_pat = f"CurrentAccumulatedQ{q}Duration_NonConsolidatedMember"
                pri_pat = f"Prior1AccumulatedQ{q}Duration_NonConsolidatedMember"

                if ctx_id == cur_pat and "current_accumulated_duration" not in result:
                    result["current_accumulated_duration"] = ctx_id
                    quarter_number = q
                    result["quarter_number"] = q
                    period = ctx.find("xbrli:period")
                    if period:
                        end = period.find("xbrli:endDate")
                        if end:
                            result["period_end"] = end.text
                    logger.info("  NonCon Q%d Context %s", q, ctx_id)

                if ctx_id == pri_pat and "prior1_accumulated_duration" not in result:
                    result["prior1_accumulated_duration"] = ctx_id
                    logger.info("  NonCon Q%d Context %s", q, ctx_id)

            if ctx_id == "CurrentQuarterInstant_NonConsolidatedMember":
                result["current_quarter_instant"] = ctx_id
            if ctx_id == "Prior1QuarterInstant_NonConsolidatedMember":
                result["prior1_quarter_instant"] = ctx_id
            if ctx_id == "InterimInstant_NonConsolidatedMember":
                result.setdefault("current_quarter_instant", ctx_id)
            if ctx_id == "Prior1InterimInstant_NonConsolidatedMember":
                result.setdefault("prior1_quarter_instant", ctx_id)
            if ctx_id == "InterimDuration_NonConsolidatedMember":
                result.setdefault("current_accumulated_duration", ctx_id)
                if not result.get("quarter_number"):
                    result["quarter_number"] = 2
                period = ctx.find("xbrli:period")
                if period:
                    end = period.find("xbrli:endDate")
                    if end:
                        result.setdefault("period_end", end.text)
            if ctx_id == "Prior1InterimDuration_NonConsolidatedMember":
                result.setdefault("prior1_accumulated_duration", ctx_id)

        if result.get("current_accumulated_duration"):
            logger.info("Using NonConsolidatedMember quarterly contexts (単体 company fallback)")

    return result


def extract_quarterly_data(soup, q_contexts, scale=SCALE_TO_MN):
    """Extract financial data from quarterly XBRL contexts.

    Returns:
        dict with keys:
            'current_cumulative': {item_key: value} for current Q cumulative
            'prior1_cumulative': {item_key: value} for prior year same Q cumulative
            'current_instant': {item_key: value} for BS at quarter end
            'quarter_number': int
            'period_end': quarter end date string
    """
    result = {
        "quarter_number": q_contexts.get("quarter_number"),
        "period_end": q_contexts.get("period_end", ""),
    }

    # Extract cumulative PL/CF data (current quarter)
    cur_dur = q_contexts.get("current_accumulated_duration")
    pri_dur = q_contexts.get("prior1_accumulated_duration")
    cur_inst = q_contexts.get("current_quarter_instant")

    for label, ctx_id in [("current_cumulative", cur_dur),
                           ("prior1_cumulative", pri_dur),
                           ("current_instant", cur_inst)]:
        if ctx_id is None:
            result[label] = {}
            continue

        data = {}
        for item_key, item_def in FINANCIAL_ITEMS.items():
            if label.endswith("_instant") and item_def["type"] != "instant":
                continue
            if label.endswith("_cumulative") and item_def["type"] != "duration":
                continue
            data[item_key] = extract_item(soup, item_key, item_def, ctx_id, scale)

        # Derive total_debt and net_debt for instant data
        if label == "current_instant":
            st = data.get("short_term_debt") or 0
            lt = data.get("long_term_debt") or 0
            if data.get("short_term_debt") is not None or data.get("long_term_debt") is not None:
                data["total_debt"] = round(st + lt, 1)
            else:
                data["total_debt"] = None
            if data.get("total_debt") is not None and data.get("cash") is not None:
                data["net_debt"] = round(data["total_debt"] - data["cash"], 1)
            else:
                data["net_debt"] = None

        result[label] = data

    return result


def calculate_ltm(latest_fy_data, quarterly_data, quarterly_period_end):
    """Calculate LTM (Last Twelve Months) data from annual + quarterly data.

    LTM formula for flow items (PL/CF):
        LTM = Latest FY (12 months) + Current Q cumulative - Prior Q cumulative

    For stock items (BS):
        Use the latest quarterly instant values directly.

    Args:
        latest_fy_data: dict of {item_key: value} from the most recent full FY.
        quarterly_data: output of extract_quarterly_data().
        quarterly_period_end: period end date of the quarterly report.

    Returns:
        tuple: (ltm_data_dict, ltm_label_string) or (None, None) if insufficient data.
    """
    q_num = quarterly_data.get("quarter_number")
    cur_cum = quarterly_data.get("current_cumulative", {})
    pri_cum = quarterly_data.get("prior1_cumulative", {})
    cur_inst = quarterly_data.get("current_instant", {})

    if not q_num or not cur_cum:
        return None, None

    ltm = {}
    for item_key, item_def in FINANCIAL_ITEMS.items():
        if item_def["type"] == "duration":
            # Flow item: LTM = FY + current_cumulative - prior_cumulative
            fy_val = latest_fy_data.get(item_key)
            cur_val = cur_cum.get(item_key)
            pri_val = pri_cum.get(item_key)

            if fy_val is not None and cur_val is not None and pri_val is not None:
                ltm[item_key] = round(fy_val + cur_val - pri_val, 1)
            elif cur_val is not None:
                # If no FY data but quarterly exists, just show cumulative
                ltm[item_key] = cur_val
            else:
                ltm[item_key] = None
        else:
            # Stock item (BS): use latest quarterly instant
            ltm[item_key] = cur_inst.get(item_key)

    # Derive total_debt and net_debt from quarterly BS
    if cur_inst.get("total_debt") is not None:
        ltm["total_debt"] = cur_inst["total_debt"]
    elif ltm.get("short_term_debt") is not None or ltm.get("long_term_debt") is not None:
        ltm["total_debt"] = round((ltm.get("short_term_debt") or 0) +
                                   (ltm.get("long_term_debt") or 0), 1)
    else:
        ltm["total_debt"] = None

    if ltm.get("total_debt") is not None and ltm.get("cash") is not None:
        ltm["net_debt"] = round(ltm["total_debt"] - ltm["cash"], 1)
    else:
        ltm["net_debt"] = None

    # Build label: "LTM Q3" or "LTM (3Q 2024/12)"
    q_end = quarterly_period_end or quarterly_data.get("period_end", "")
    if q_end and len(q_end) >= 7:
        ltm_label = f"LTM({q_num}Q {q_end[:7]})"
    else:
        ltm_label = f"LTM(Q{q_num})"

    return ltm, ltm_label


# =====================================================================
# MULTI-YEAR MERGE
# =====================================================================
def _fiscal_year_label(period_end_str):
    """Convert a period_end date string like '2025-03-31' to 'FY2025'."""
    if period_end_str and len(period_end_str) >= 4:
        return f"FY{period_end_str[:4]}"
    return period_end_str


def _resolve_year_from_context(meta, year_key):
    """Extract the fiscal year end date from context metadata."""
    dur_meta = meta.get(f"{year_key}_duration", {})
    inst_meta = meta.get(f"{year_key}_instant", {})
    return inst_meta.get("instant", dur_meta.get("end", ""))


def merge_multi_year_data(all_year_data):
    """Merge financial data from multiple XBRL files into one dataset.

    Each XBRL file typically contains CurrentYear and Prior1Year data.
    This function collects all years, using the newest file's values as
    authoritative (since the same year may appear as 'current' in one file
    and 'prior1' in a newer file).

    Args:
        all_year_data: list of (period_end, data_dict) tuples, sorted newest-first.
            Each data_dict is the output of extract_financial_data() and contains
            year keys like 'current', 'prior1', etc., plus '_meta'.

    Returns:
        OrderedDict keyed by fiscal year label (e.g. "FY2025"), sorted newest-first.
        Includes '_meta' key mapping each FY label to period date info.
    """
    # Collect all fiscal year data, keyed by the actual calendar year end date
    # Priority: newer XBRL file's data wins (since all_year_data is newest-first)
    fy_data = OrderedDict()  # date_str -> {item: value}
    fy_meta = {}  # date_str -> meta info

    for period_end, data in all_year_data:
        meta = data.get("_meta", {})
        year_keys = [k for k in data if k != "_meta"]

        for year_key in year_keys:
            # Resolve the actual calendar date for this year_key
            actual_date = _resolve_year_from_context(meta, year_key)
            if not actual_date:
                continue

            # Only fill in if we haven't seen this fiscal year yet (newer data wins)
            if actual_date not in fy_data:
                fy_data[actual_date] = data[year_key]
                # Collect meta for this year
                dur_key = f"{year_key}_duration"
                inst_key = f"{year_key}_instant"
                if dur_key in meta:
                    fy_meta[actual_date] = {"duration": meta[dur_key]}
                if inst_key in meta:
                    fy_meta.setdefault(actual_date, {})["instant"] = meta[inst_key]
            else:
                # Fill in any None values from older file's data
                existing = fy_data[actual_date]
                for item_key, val in data[year_key].items():
                    if existing.get(item_key) is None and val is not None:
                        existing[item_key] = val

    # Sort by date descending (newest first) and re-key as FY labels
    sorted_dates = sorted(fy_data.keys(), reverse=True)

    result = OrderedDict()
    result_meta = {}
    for date_str in sorted_dates:
        fy_label = _fiscal_year_label(date_str)
        result[fy_label] = fy_data[date_str]
        if date_str in fy_meta:
            fm = fy_meta[date_str]
            # Map meta keys to match print_results expectations
            result_meta[f"{fy_label}_instant"] = fm.get("instant", {"instant": date_str})
            if "duration" in fm:
                result_meta[f"{fy_label}_duration"] = fm["duration"]

    result["_meta"] = result_meta
    return result


# =====================================================================
# PRETTY PRINT
# =====================================================================
def print_results(company_info, data):
    """Pretty-print extracted financial data to console."""
    print()
    meta = data.get("_meta", {})
    year_keys = [k for k in data if k != "_meta"]

    # Determine column width based on number of columns
    num_cols = len(year_keys)
    col_width = 12
    total_width = max(70, 42 + (col_width + 2) * num_cols)

    print("=" * total_width)
    print(f"  EDINET XBRL Financial Data Extraction")
    print("=" * total_width)

    if company_info:
        print(f"  Company:         {company_info.get('company_name', 'N/A')}")
        print(f"  EDINET Code:     {company_info.get('edinet_code', 'N/A')}")
        print(f"  Securities Code: {company_info.get('securities_code', 'N/A')}")
        print(f"  FY End:          {company_info.get('fiscal_year_end', 'N/A')}")

    print()
    print("-" * total_width)

    # Build column headers
    col_headers = []
    for yk in year_keys:
        if yk.startswith("FY") or yk.startswith("LTM"):
            # Already a fiscal year or LTM label (from merge_multi_year_data)
            col_headers.append(yk)
        else:
            # Legacy format: resolve from meta
            dur_meta = meta.get(f"{yk}_duration", {})
            inst_meta = meta.get(f"{yk}_instant", {})
            period_end = inst_meta.get("instant", dur_meta.get("end", yk))
            col_headers.append(f"FY{period_end[:4]}" if len(period_end) >= 4 else yk)

    # Print header
    header = f"  {'Item':<40s}"
    for ch in col_headers:
        header += f"  {ch:>{col_width}s}"
    print(header)
    print("-" * total_width)

    # Print each financial item
    for item_key, item_def in FINANCIAL_ITEMS.items():
        label = item_def["label"]
        row = f"  {label:<40s}"
        for yk in year_keys:
            val = data[yk].get(item_key)
            if val is not None:
                row += f"  {val:>{col_width},.0f}"
            else:
                row += f"  {'N/A':>{col_width}s}"
        print(row)

    # Print derived fields
    print("-" * total_width)
    for derived in ["total_debt", "net_debt"]:
        label_map = {
            "total_debt": "Total Debt (有利子負債合計)",
            "net_debt": "Net Debt (ネットデット)",
        }
        label = label_map.get(derived, derived)
        row = f"  {label:<40s}"
        for yk in year_keys:
            val = data[yk].get(derived)
            if val is not None:
                row += f"  {val:>{col_width},.0f}"
            else:
                row += f"  {'N/A':>{col_width}s}"
        print(row)

    print("=" * total_width)
    print("  All values in JPY millions (百万円)")
    print("=" * total_width)


# =====================================================================
# CLI ENTRY POINT
# =====================================================================
def main():
    """CLI interface for edinet_parser."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    if len(sys.argv) < 2:
        print("Usage: python scripts/edinet_parser.py <path_to_xbrl_file>")
        print("Example: python scripts/edinet_parser.py tmp/edinet_data/S100XXXX/XBRL/PublicDoc/xxx.xbrl")
        sys.exit(1)

    xbrl_path = sys.argv[1]
    if not os.path.isfile(xbrl_path):
        print(f"ERROR: File not found: {xbrl_path}")
        sys.exit(1)

    # Parse
    soup = parse_xbrl_file(xbrl_path)

    # Company info
    company_info = extract_company_info(soup)

    # Identify clean contexts
    contexts = identify_clean_contexts(soup)
    if not contexts:
        print("ERROR: No clean contexts found in XBRL file (checked both consolidated and non-consolidated).")
        sys.exit(1)

    # Extract data
    data = extract_financial_data(soup, contexts)

    # Print results
    print_results(company_info, data)


if __name__ == "__main__":
    main()
