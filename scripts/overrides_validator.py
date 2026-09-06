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
# 手順書 §2 の銘柄型。宣言は任意だが、型D は DCF が成立しないため
# scripts/arbitration.py がこのキーを見て DCF 脚の裁定をスキップする(追補12 §A-3)。
COMPANY_TYPES = {"A", "B", "C", "D", "E"}

ALLOWED_KEYS = {
    # Company / meta
    "ticker": (str, int),
    "company_name": (str,),
    "exchange": (str,),
    "sector": (str,),
    "fiscal_year_end_month": (int,),
    "company_type": (str,),                # A/B/C/D/E - 手順書 §2 の銘柄型
    "bank_valuation": (dict,),             # 型D 専用。scripts/ddm_ri.py が消費する
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
    # Actual cost of debt (optional): supplying interest_expense switches C11
    # from the assumed rate to interest / average interest-bearing debt.
    "interest_expense": _NUM,
    "loan_fees": _NUM,
    "debt_beginning": _NUM,
    "debt_ending": _NUM,
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
    "book_value": _NUM,                    # 自社の純資産（Comps の PBR/ROE 数式用）
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
    "reverse_dcf": (dict,),                # consumed by templates/reverse_dcf_sheet.py
    "fx_sensitivity": (dict,),             # Sensitivity Table 3 (export-exposed names)
    "normalized_net_income": (dict,) + _NUM,   # Comps 正常化純利益の参考行
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

# reverse_dcf sub-keys: everything else in that block is derived from the
# hist_* arrays. Anything unrecognised here would be silently dropped, which is
# exactly the failure mode this validator exists to prevent.
REVERSE_DCF_KEYS = {
    "enabled": (bool,),
    "op0": (int, float),
    "op0_label": (str,),
    "peak_op": (int, float),
    "peak_label": (str,),
    "peak_opm": (int, float),
    "opm_grid": (list,),
    "n_years": (list,),
    "benchmark_ticker": (str, int),
    "deal_note": (list,),
}


FX_SENSITIVITY_KEYS = {
    "enabled": (bool,),
    "assumption_rate": (int, float),
    "assumption_source": (str,),
    "usd_revenue_ratio": (int, float),
    "usd_cogs_ratio": (int, float),
    "currency_pair": (str,),
    "offsets": (list,),
    "estimated": (bool,),
    "note": (str,),
}

NORMALIZED_NI_KEYS = {
    "pretax": (int, float),
    "addbacks": (int, float),
    "value": (int, float),
    "label": (str,),
    "note": (str,),
}


def _validate_subblock(name, block, allowed, errors):
    for k, v in block.items():
        if k.startswith("_"):
            continue
        if k not in allowed:
            errors.append(
                f"{name}.{k}: unknown key (would be silently ignored). "
                f"Allowed: {', '.join(sorted(allowed))}"
            )
            continue
        if v is None:
            continue
        types = allowed[k]
        if types != (bool,) and isinstance(v, bool):
            errors.append(f"{name}.{k}: expected {_type_name(types)}, got bool")
        elif not isinstance(v, types):
            errors.append(f"{name}.{k}: expected {_type_name(types)}, got "
                          f"{type(v).__name__} ({v!r})")


def _validate_fx_sensitivity(block, errors):
    _validate_subblock("fx_sensitivity", block, FX_SENSITIVITY_KEYS, errors)
    if not block.get("enabled"):
        return
    for k in ("assumption_rate", "usd_revenue_ratio", "usd_cogs_ratio"):
        if block.get(k) is None:
            errors.append(f"fx_sensitivity.{k}: required when enabled is true")
    rate = block.get("assumption_rate")
    if isinstance(rate, (int, float)) and not isinstance(rate, bool) and rate <= 0:
        errors.append("fx_sensitivity.assumption_rate: must be > 0 "
                      "(it is the divisor of the sensitivity formula)")
    for k in ("usd_revenue_ratio", "usd_cogs_ratio"):
        v = block.get(k)
        if isinstance(v, (int, float)) and not isinstance(v, bool) and not 0 <= v <= 1:
            errors.append(f"fx_sensitivity.{k}: must be a decimal share in [0, 1], "
                          f"got {v!r}")
    if isinstance(block.get("offsets"), list):
        if not block["offsets"]:
            errors.append("fx_sensitivity.offsets: must not be empty")
        for i, o in enumerate(block["offsets"]):
            if isinstance(o, bool) or not isinstance(o, (int, float)):
                errors.append(f"fx_sensitivity.offsets[{i}]: must be a number "
                              f"(JPY offset from the assumption rate)")


def _validate_normalized_ni(block, errors):
    if not isinstance(block, dict):
        return
    _validate_subblock("normalized_net_income", block, NORMALIZED_NI_KEYS, errors)
    if block.get("pretax") is None and block.get("value") is None:
        errors.append("normalized_net_income: needs either 'pretax' (+ optional "
                      "'addbacks', taxed at the model rate) or a ready 'value'")


def _validate_reverse_dcf(block, errors):
    for k, v in block.items():
        if k.startswith("_"):
            continue
        if k not in REVERSE_DCF_KEYS:
            errors.append(
                f"reverse_dcf.{k}: unknown key (would be silently ignored). "
                f"Allowed: {', '.join(sorted(REVERSE_DCF_KEYS))}"
            )
            continue
        if v is None:
            continue
        types = REVERSE_DCF_KEYS[k]
        if types != (bool,) and isinstance(v, bool):
            errors.append(f"reverse_dcf.{k}: expected {_type_name(types)}, got bool")
        elif not isinstance(v, types):
            errors.append(
                f"reverse_dcf.{k}: expected {_type_name(types)}, got "
                f"{type(v).__name__} ({v!r})"
            )
    if isinstance(block.get("opm_grid"), list):
        for i, m in enumerate(block["opm_grid"]):
            if m is not None and (isinstance(m, bool) or not isinstance(m, (int, float))):
                errors.append(f"reverse_dcf.opm_grid[{i}]: must be a number or null "
                              f"(null = the live cycle-peak OPM cell)")
    if isinstance(block.get("n_years"), list):
        if not block["n_years"]:
            errors.append("reverse_dcf.n_years: must not be empty")
        for i, n in enumerate(block["n_years"]):
            if isinstance(n, bool) or not isinstance(n, int) or n < 1:
                errors.append(f"reverse_dcf.n_years[{i}]: must be a positive integer")
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


# Historical arrays that must line up column-for-column with hist_years.
# A short/long array used to be copied in positionally, putting one year's cash
# flow under another year's header (bug B1) — length is now part of the contract.
HIST_ARRAY_KEYS = (
    "hist_revenue", "hist_operating_income", "hist_net_income", "hist_cogs",
    "hist_sga", "hist_ocf", "hist_capex", "hist_cash", "hist_debt",
    "hist_depreciation", "hist_nwc_pct",
)

# Tokens the template can resolve inside investment_thesis / key_risks.
# Keep in sync with NARRATIVE_TOKEN_FORMATS in templates/dcf_comps_template.py.
NARRATIVE_TOKENS = ("price", "target_price", "upside_pct", "pb", "per", "wacc")
_NARRATIVE_TOKEN_RE = re.compile(r"\{([a-zA-Z_][a-zA-Z0-9_]*)\}")


def _validate_hist_lengths(overrides, errors):
    years = overrides.get("hist_years")
    if not isinstance(years, list):
        return
    n = len(years)
    for key in HIST_ARRAY_KEYS:
        val = overrides.get(key)
        if isinstance(val, list) and len(val) != n:
            errors.append(
                f"{key}: length {len(val)} != hist_years length ({n}). Historical "
                f"arrays are matched to hist_years column-by-column; a mismatched "
                f"array would put values under the wrong fiscal year."
            )


def _validate_narrative_tokens(overrides, errors):
    for key in ("investment_thesis", "key_risks"):
        lines = overrides.get(key)
        if not isinstance(lines, list):
            continue
        for i, line in enumerate(lines):
            if not isinstance(line, str):
                continue
            for m in _NARRATIVE_TOKEN_RE.finditer(line):
                name = m.group(1)
                if name not in NARRATIVE_TOKENS:
                    errors.append(
                        f"{key}[{i}]: unknown narrative token '{{{name}}}'. "
                        f"Allowed: {', '.join('{' + t + '}' for t in NARRATIVE_TOKENS)}. "
                        f"(A typo would otherwise be printed literally in the report.)"
                    )


def _warn_terminal_capex(overrides):
    """Pre-flight version of validate_output.py check #9 (warning only).

    A perpetuity growth model assumes the terminal year repeats forever. With
    g <= 1.5% and terminal capex far above D&A, the model quietly reinvests more
    than it depreciates for eternity (8267: PGM went negative). This is an
    analyst call, so it is never auto-corrected — but it should be visible
    BEFORE generation, not after.
    """
    g = overrides.get("terminal_growth")
    if not isinstance(g, (int, float)) or g > 0.015:
        return
    capex = (overrides.get("capex_direct") or {}).get("projections") or []
    da = (overrides.get("da_direct") or {}).get("projections") or []
    if not capex or not da:
        capex_pct, da_pct = overrides.get("capex_pct"), overrides.get("da_pct")
        if not (isinstance(capex_pct, (int, float)) and isinstance(da_pct, (int, float))
                and da_pct):
            return
        ratio = capex_pct / da_pct
    else:
        if not da[-1]:
            return
        ratio = capex[-1] / da[-1]
    if not (0.90 <= ratio <= 1.15):
        print(f"  [overrides] WARNING: terminal capex / D&A = {ratio:.2f}x with "
              f"terminal_growth {g:.2%} (outside [0.90, 1.15]). A perpetuity at "
              f"this reinvestment rate can drive PGM negative - confirm it is "
              f"intentional (not auto-corrected).")



def _check_type_e_contract(overrides):
    """型E (銀行/金融子会社を連結に持つ事業会社) が明示を要求する4項目。

    型E で自動値に落ちてよいものは一つもない。連結BS の預金・貸出金、連結P/L の
    経常収益が、それぞれ net_debt / de_ratio / DCF の売上に混入するため、
    「指定が無ければ従来どおり自動」では黙って誤った数字が出る。
    """
    errors = []

    if "net_debt" not in overrides or _is_placeholder_value(overrides.get("net_debt")):
        errors.append(
            "型E: net_debt の明示が必須です。連結BS からの自動抽出は銀行の預金を"
            "有利子負債に、貸出金を資産に含めてしまいます。非金融ベース"
            "(銀行預金・貸出金・コールローンを除外し、非支配株主持分を加算)の値を"
            "一次資料から入れてください")

    if "de_ratio" not in overrides or _is_placeholder_value(overrides.get("de_ratio")):
        errors.append(
            "型E: de_ratio の明示が必須です(自動計算は禁止)。自動計算は "
            "net_debt / 時価総額 で求めるため、金融部門を含む連結 net_debt を"
            "使うと WACC の資本構成が壊れます")

    segs = overrides.get("segments")
    if not segs:
        errors.append(
            "型E: segments が必須です。連結 P/L には銀行の経常収益・経常利益が"
            "含まれるため、非金融セグメントのみを Segment Analysis 経由で"
            "DCF に供給してください(Segment Analysis が DCF Revenue/EBIT の"
            "single source of truth)")

    sotp = overrides.get("sotp")
    if not isinstance(sotp, dict):
        errors.append(
            "型E: sotp ブロックが必須です。金融セグメントは DCF ではなく "
            "PBR×純資産で評価し、非金融の事業価値と SOTP で合算します")
    else:
        fin = [x for x in sotp.get("segments", [])
               if str(x.get("valuation_method", "")).strip().lower() == "pbr"]
        if not fin:
            errors.append(
                "型E: sotp.segments に valuation_method:\"pbr\" のセグメントが"
                "1つもありません。金融子会社を PBR で評価しないなら、その銘柄は"
                "型E ではありません")
        for x in fin:
            if not isinstance(x.get("net_assets_mn"), (int, float)) or isinstance(
                    x.get("net_assets_mn"), bool):
                errors.append(
                    f"型E: sotp segment {x.get('key')!r} は "
                    f"net_assets_mn (JPY mn, 金融セグメントの純資産) が必須です")
    return errors


def _is_placeholder_value(v):
    return isinstance(v, str) and "__CONFIRM__" in v


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

    bv = overrides.get("bank_valuation")
    if bv is not None:
        if str(overrides.get("company_type", "")).strip().upper() != "D":
            errors.append(
                "bank_valuation: 型D 専用のブロックです。company_type: \"D\" を宣言するか、"
                "このブロックを削除してください")
        allowed_bv = {"book_value_mn", "net_income_actual_mn", "dps", "roe",
                      "terminal_growth", "year_labels", "dps_note", "roe_note"}
        for k in bv:
            if k not in allowed_bv and not k.startswith("_"):
                errors.append(f"bank_valuation.{k}: unknown key. "
                              f"許可: {', '.join(sorted(allowed_bv))}")
        for k in ("dps", "roe"):
            v = bv.get(k)
            if v is None:
                errors.append(f"bank_valuation.{k}: required (5要素の配列)")
            elif not isinstance(v, list) or len(v) != 5:
                errors.append(f"bank_valuation.{k}: 5要素の配列が必要 "
                              f"(got {type(v).__name__} len={len(v) if isinstance(v, list) else 'n/a'})")
            elif any(not isinstance(x, (int, float)) or isinstance(x, bool) for x in v):
                errors.append(f"bank_valuation.{k}: 全要素が数値でなければならない")
        if not isinstance(bv.get("book_value_mn"), (int, float)) or isinstance(bv.get("book_value_mn"), bool):
            errors.append("bank_valuation.book_value_mn: required (JPY mn の数値)")
        _p = bv.get("roe")
        if isinstance(_p, list) and any(isinstance(x, (int, float)) and x > 1 for x in _p):
            errors.append("bank_valuation.roe: 小数で指定すること (6.1% は 0.061)")

    if str(overrides.get("company_type", "")).strip().upper() == "E":
        errors.extend(_check_type_e_contract(overrides))

    ct = overrides.get("company_type")
    if ct is not None and str(ct).strip().upper() not in COMPANY_TYPES:
        errors.append(
            f"company_type: {ct!r} is not one of {sorted(COMPANY_TYPES)}. "
            f"手順書 §2 の銘柄型 (A: 通常の事業会社 / B: シクリカル / "
            f"C: captive finance 持ち製造業 / D: 銀行 / E: 銀行を連結に持つ持株会社)"
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

    if isinstance(overrides.get("reverse_dcf"), dict):
        _validate_reverse_dcf(overrides["reverse_dcf"], errors)

    if isinstance(overrides.get("fx_sensitivity"), dict):
        _validate_fx_sensitivity(overrides["fx_sensitivity"], errors)

    if isinstance(overrides.get("normalized_net_income"), dict):
        _validate_normalized_ni(overrides["normalized_net_income"], errors)

    # interest_expense drives C11; without a debt base it silently does nothing.
    if overrides.get("interest_expense") is not None:
        _has_explicit = (overrides.get("debt_beginning") is not None
                         and overrides.get("debt_ending") is not None)
        _hd = overrides.get("hist_debt")
        _n_hd = len([d for d in _hd if isinstance(d, (int, float))]) if isinstance(_hd, list) else 0
        if not _has_explicit and _n_hd < 2:
            errors.append(
                "interest_expense: needs a debt base — set debt_beginning and "
                "debt_ending, or supply at least two numeric hist_debt years "
                "(the average of the last two is used)"
            )

    _validate_hist_lengths(overrides, errors)
    _validate_narrative_tokens(overrides, errors)
    _warn_terminal_capex(overrides)

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
