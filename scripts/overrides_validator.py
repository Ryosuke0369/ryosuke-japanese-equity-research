"""
overrides_validator.py - Strict contract validation for data/overrides/*.json.

Why this exists: the generation pipeline used to copy ANY override key into the
config dict and silently ignore keys the templates never read (nested
`wacc_inputs`, legacy scenario names like Bull/Bear, typos, etc.). The run
completed with zero errors while the model was built on default assumptions.
This validator makes that failure mode impossible: any key or structure the
templates do not actually consume is rejected BEFORE generation starts.

Contract summary (see docs/overrides_schema.md for the full spec):
  - WACC inputs are FLAT top-level keys (risk_free / beta / erp / size_premium
    / cost_of_debt_at / de_ratio). No nesting.
  - scenarios uses exactly the 5 fixed names: Base / Upside / Management /
    Downside 1 / Downside 2. Per-year arrays must match projection_years.
  - Keys starting with "_" are comments and always allowed.
  - "__CONFIRM__" placeholder values abort the run unless explicitly allowed
    (generate_dcf.py --allow-unconfirmed).
"""

import re

SCENARIO_NAMES = ("Base", "Upside", "Management", "Downside 1", "Downside 2")

LEGACY_SCENARIO_NAMES = {
    "Bull", "StrongBull", "Strong Bull", "Bear", "StrongBear", "Strong Bear",
}

# Per-year array keys inside a scenario block (length == projection_years)
SCENARIO_ARRAY_KEYS = (
    "revenue_growth", "cogs_pct", "sga_pct",
    "dso_days", "dih_days", "dpo_days", "nwc_pct",
)

# Nested containers people intuitively write but the templates never read.
REJECTED_CONTAINERS = {
    "wacc_inputs": "WACC inputs must be FLAT top-level keys: "
                   "risk_free / beta / erp / size_premium / cost_of_debt_at / de_ratio",
    "wacc": "WACC is computed from flat keys (risk_free / beta / erp / "
            "size_premium / cost_of_debt_at / de_ratio); it cannot be set directly",
    "dcf_assumptions": "DCF assumptions must be flat top-level keys "
                       "(tax_rate, terminal_growth, exit_multiple, ...)",
    "comps_input": "Comps come from data/comps/<ticker>_comps.csv, "
                   "not from the overrides JSON",
    "comps": "generate_dcf.py ALWAYS loads comps from data/comps/<ticker>_comps.csv "
             "(Step 6), overwriting any value set here. Move the data to the CSV; "
             "keep notes under a '_'-prefixed key",
}

# Common wrong spellings -> the key the code actually reads.
KEY_SUGGESTIONS = {
    "risk_free_rate": "risk_free",
    "rf": "risk_free",
    "equity_risk_premium": "erp",
    "market_risk_premium": "erp",
    "cost_of_debt": "cost_of_debt_at",
    "debt_equity_ratio": "de_ratio",
    "hist_da": "hist_depreciation",
    "hist_ordinary_income": None,   # not consumed anywhere
    "company_name_en": "company_name",
    "company_name_jp": "company_name",
    "fiscal_year_end": "fiscal_year_end_month",
    "arr_latest": None,             # not consumed anywhere
    "exit_ev_sales": "exit_sales_multiple",
    "shares_fully_diluted": "shares",
}

_NUM = (int, float)
_STR_OR_NUM = (str, int, float)  # values that may carry a __CONFIRM__ string

# Top-level whitelist: key -> allowed python types.
# Every key here is verifiably consumed by generate_dcf.py,
# templates/dcf_comps_template.py, or scripts/generate_sotp.py.
ALLOWED_KEYS = {
    # Company / meta
    "ticker": (str, int),
    "company_name": (str,),
    "exchange": (str,),
    "sector": (str,),
    "fiscal_year_end_month": (int,),
    # Market data
    "current_price": _STR_OR_NUM,
    "shares_outstanding": _STR_OR_NUM,
    "shares": (dict,),                     # {"fully_diluted_shares": N}
    "net_debt": _STR_OR_NUM,
    "beta": _NUM,
    "de_ratio": _NUM,
    # WACC inputs (flat!)
    "risk_free": _NUM,
    "erp": _NUM,
    "size_premium": _NUM,
    "cost_of_debt_at": _NUM,
    "tax_rate": _NUM,
    # Terminal value / exit
    "terminal_growth": _NUM,
    "exit_multiple": _NUM,
    "exit_sales_multiple": _NUM,
    "primary_multiple": (str,),
    # Capex / D&A
    "capex_method": (str,),
    "da_method": (str,),
    "capex_pct": _NUM,
    "da_pct": _NUM,
    "capex_direct": (dict,),
    "da_direct": (dict,),
    # NWC
    "nwc_method": (str,),
    "nwc_items": (list,),
    "base_year_ar": _NUM,
    "base_year_inv": _NUM,
    "base_year_ap": _NUM,
    "base_year_nwc": _NUM,
    "base_year_trade_receivables": _NUM,
    "base_year_trade_payables": _NUM,
    # Projection frame
    "projection_years": (int,),
    "projection_start_fy": (str,),
    "stub_fraction": _NUM,
    "stub_months_elapsed": (int,),
    "ltm_revenue": _NUM,
    "base_year_revenue": _NUM,
    "base_year_cogs": _NUM,
    "core_ebitda": _NUM,
    "core_net_income": _NUM,
    # Historical arrays
    "hist_years": (list,),
    "hist_revenue": (list,),
    "hist_operating_income": (list,),
    "hist_net_income": (list,),
    "hist_cogs": (list,),
    "hist_sga": (list,),
    "hist_ocf": (list,),
    "hist_capex": (list,),
    "hist_cash": (list,),
    "hist_debt": (list,),
    "hist_depreciation": (list,),
    "hist_nwc_pct": (list,),
    # Structured blocks
    "scenarios": (dict,),
    "segments": (list,),
    "sotp": (dict,),                       # consumed by generate_sotp.py
    "cost_structure": (dict,),             # consumed by SOTP/segment runs
    # Narrative
    "investment_thesis": (list,),
    "key_risks": (list,),
}

_FUNDAMENTALS_RE = re.compile(r"fundamentals_fy\d{4}$")


class OverridesValidationError(ValueError):
    """Raised when an overrides JSON violates the template contract."""


def _type_name(types):
    return " | ".join(t.__name__ for t in types)


def _find_confirm_placeholders(node, path=""):
    """Recursively collect JSON paths whose value contains '__CONFIRM__'."""
    found = []
    if isinstance(node, str):
        if "__CONFIRM__" in node:
            found.append(path or "<root>")
    elif isinstance(node, dict):
        for k, v in node.items():
            found.extend(_find_confirm_placeholders(v, f"{path}.{k}" if path else k))
    elif isinstance(node, list):
        for i, v in enumerate(node):
            found.extend(_find_confirm_placeholders(v, f"{path}[{i}]"))
    return found


NWC_METHODS = ("days", "revenue_pct", "itemized")
NWC_ITEM_REQUIRED = ("label", "base_value", "scenario_key", "side", "denom")


def _validate_scenarios(scenarios, projection_years, errors, extra_array_keys=()):
    allowed_array_keys = tuple(SCENARIO_ARRAY_KEYS) + tuple(extra_array_keys)
    unknown = [n for n in scenarios if n not in SCENARIO_NAMES]
    for name in unknown:
        if name in LEGACY_SCENARIO_NAMES:
            errors.append(
                f"scenarios.'{name}': legacy scenario name. The template only reads "
                f"the 5 fixed names: {', '.join(SCENARIO_NAMES)}. "
                f"Map Bull->Upside, Bear->Downside 1, StrongBear->Downside 2 etc."
            )
        else:
            errors.append(
                f"scenarios.'{name}': unknown scenario name (silently ignored by the "
                f"template). Allowed: {', '.join(SCENARIO_NAMES)}"
            )
    for name in SCENARIO_NAMES:
        block = scenarios.get(name)
        if block is None:
            continue  # partial scenario overrides merge onto auto-generated ones
        if not isinstance(block, dict):
            errors.append(f"scenarios.'{name}': must be an object")
            continue
        for k, v in block.items():
            if k.startswith("_"):
                continue
            if k not in allowed_array_keys:
                errors.append(
                    f"scenarios.'{name}'.{k}: unknown scenario key. "
                    f"Allowed: {', '.join(allowed_array_keys)} "
                    f"(itemized-NWC keys must be declared via nwc_items[*].scenario_key)"
                )
                continue
            if not isinstance(v, list):
                errors.append(f"scenarios.'{name}'.{k}: must be a per-year array")
            elif len(v) != projection_years:
                errors.append(
                    f"scenarios.'{name}'.{k}: length {len(v)} != projection_years "
                    f"({projection_years})"
                )


def validate_overrides(overrides, source_path="<overrides>", allow_unconfirmed=False):
    """Validate an overrides dict against the template contract.

    Raises OverridesValidationError listing ALL violations at once (so the
    analyst can fix the file in one pass). Returns None on success.
    """
    if not isinstance(overrides, dict):
        raise OverridesValidationError(f"{source_path}: top level must be a JSON object")

    errors = []
    projection_years = overrides.get("projection_years", 5)
    if not isinstance(projection_years, int) or projection_years < 1:
        errors.append(f"projection_years: must be a positive integer, got {projection_years!r}")
        projection_years = 5

    for key, value in overrides.items():
        if key.startswith("_"):
            continue  # comment key, always allowed
        if _FUNDAMENTALS_RE.match(key):
            if not isinstance(value, dict):
                errors.append(f"{key}: must be an object with 決算短信 actuals")
            elif value.get("revenue") is None:
                errors.append(f"{key}: 'revenue' is required (used as latest FY actual)")
            continue
        if key in REJECTED_CONTAINERS:
            errors.append(f"{key}: not read by any template. {REJECTED_CONTAINERS[key]}")
            continue
        if key not in ALLOWED_KEYS:
            if key in KEY_SUGGESTIONS:
                hint = KEY_SUGGESTIONS[key]
                fix = f"did you mean '{hint}'?" if hint else "no template consumes this; delete it or prefix with '_' to keep as a comment"
                errors.append(f"{key}: unknown key - {fix}")
            else:
                errors.append(
                    f"{key}: unknown key (would be silently ignored). If it is a "
                    f"comment, prefix it with '_'. See docs/overrides_schema.md"
                )
            continue
        allowed_types = ALLOWED_KEYS[key]
        if value is None:
            continue  # explicit null = clear the auto-derived value (guarded downstream)
        if isinstance(value, bool) or not isinstance(value, allowed_types):
            # bool is an int subclass — never a valid override value here
            errors.append(
                f"{key}: expected {_type_name(allowed_types)}, got "
                f"{type(value).__name__} ({value!r})"
            )

    if overrides.get("nwc_method") is not None and overrides["nwc_method"] not in NWC_METHODS:
        errors.append(
            f"nwc_method: {overrides['nwc_method']!r} is not one of {NWC_METHODS}"
        )

    # nwc_items (itemized NWC): each item's scenario_key names a per-year array
    # inside every scenario block, so those keys become valid scenario keys.
    nwc_item_keys = []
    if isinstance(overrides.get("nwc_items"), list):
        for i, item in enumerate(overrides["nwc_items"]):
            if not isinstance(item, dict):
                errors.append(f"nwc_items[{i}]: must be an object")
                continue
            missing = [k for k in NWC_ITEM_REQUIRED if k not in item]
            if missing:
                errors.append(f"nwc_items[{i}]: missing required key(s): {', '.join(missing)}")
            if item.get("scenario_key"):
                nwc_item_keys.append(item["scenario_key"])

    if isinstance(overrides.get("scenarios"), dict):
        _validate_scenarios(overrides["scenarios"], projection_years, errors,
                            extra_array_keys=nwc_item_keys)

    if isinstance(overrides.get("shares"), dict):
        if "fully_diluted_shares" not in overrides["shares"]:
            errors.append("shares: must contain 'fully_diluted_shares'")

    confirms = _find_confirm_placeholders(
        {k: v for k, v in overrides.items() if not k.startswith("_")}
    )
    if confirms:
        msg = (
            f"unfilled __CONFIRM__ placeholder(s): {', '.join(confirms)}. "
            f"Fill confirmed values from the earnings release (kessan tanshin), "
            f"or rerun with --allow-unconfirmed to knowingly use auto-derived fallbacks."
        )
        if allow_unconfirmed:
            print(f"  [overrides] WARNING (allowed by --allow-unconfirmed): {msg}")
        else:
            errors.append(msg)

    if errors:
        bullet = "\n  - ".join(errors)
        raise OverridesValidationError(
            f"Overrides contract violation in {source_path} "
            f"({len(errors)} issue(s)):\n  - {bullet}\n"
            f"Nothing was generated. Fix the overrides file and rerun."
        )
