"""
validate_output.py - Machine-check a generated (or hand-edited) DCF workbook.

Usage:
    python scripts/validate_output.py models/3687_DCF_Model_20260731.xlsx

Exit code 1 when any check FAILs, 0 otherwise. Results also go to
<xlsx>_validation.txt next to the workbook.

A SKIPped check is a check that could NOT run, so SKIP > 0 is a FAIL: the
workbook is unverified, not verified-and-clean. 9503 shipped with
"FAIL 0 / SKIP 5 / VERDICT: PASS" and an empty Target Price because a recalc
had been interrupted. Pass --allow-skip (or validate_workbook(allow_skip=True))
when partial validation is deliberate - generate_dcf.py does exactly that for
--no-recalc runs, where the value-level checks legitimately cannot run.

Why: every bug this checks for shipped at least once in a workbook that looked
correct on screen — a scenario dropdown wired to the constant 1, a median that
included the subject company, operating cash flow sitting under the wrong year.
The workbook carries its own ground truth in the 'Adjustments Log' sheet's
Pipeline Metadata block, so the checks compare the sheet against what the
generator intended rather than re-deriving it.

The workbook is loaded twice: data_only=False for formulas, data_only=True for
the values Excel cached at the last recalc. Checks that need values report
"needs recalc" (WARN) instead of failing when the workbook was never recalced.
"""

import json
import os
import re
import sys

import openpyxl

# cp932 console: never let an un-encodable character abort a validation run.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from templates.dcf_comps_template import (  # noqa: E402
    SCENARIO_NAMES, R_DA, R_CAPEX, R_TV_PGM, R_EV_PGM, R_EV_EXIT, R_YR5_EBITDA,
)
from templates.reverse_dcf_sheet import SHEET_NAME as REVERSE_DCF_SHEET  # noqa: E402

FAIL, WARN, PASS, SKIP = "FAIL", "WARN", "PASS", "SKIP"

# Errors that always mean a broken model
FATAL_ERRORS = ("#REF!", "#VALUE!", "#DIV/0!", "#NAME?")
# Errors that are usually a symptom but occasionally intentional
SOFT_ERRORS = ("#NUM!", "#N/A", "#NULL!")

TOL = 0.001  # 0.1%


class Result:
    def __init__(self, path):
        self.path = path
        self.rows = []          # (n, level, title, detail)
        self.failed = False

    def add(self, n, level, title, detail=""):
        self.rows.append((n, level, title, detail))
        if level == FAIL:
            self.failed = True

    def render(self):
        width = max((len(t) for _, _, t, _ in self.rows), default=20)
        lines = [f"validate_output.py — {os.path.basename(self.path)}", "=" * 78]
        for n, level, title, detail in self.rows:
            lines.append(f"[{level:<4}] {n:>2}. {title.ljust(width)}  {detail}")
        n_fail = sum(1 for r in self.rows if r[1] == FAIL)
        n_warn = sum(1 for r in self.rows if r[1] == WARN)
        n_skip = sum(1 for r in self.rows if r[1] == SKIP)
        lines += ["=" * 78,
                  f"FAIL {n_fail} / WARN {n_warn} / SKIP {n_skip} / "
                  f"PASS {len(self.rows) - n_fail - n_warn - n_skip}",
                  "VERDICT: " + ("FAIL" if self.failed else "PASS")]
        return "\n".join(lines)


# =====================================================================
# helpers
# =====================================================================
def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _close(a, b, tol=TOL):
    if a is None or b is None:
        return False
    if b == 0:
        return abs(a) < tol
    return abs(a - b) / abs(b) <= tol


def read_metadata(wbf):
    """Read the Pipeline Metadata key/value block from 'Adjustments Log'."""
    meta = {}
    if "Adjustments Log" not in wbf.sheetnames:
        return meta
    ws = wbf["Adjustments Log"]
    in_block = False
    for row in ws.iter_rows(min_col=2, max_col=3):
        key, val = row[0].value, row[1].value
        if isinstance(key, str) and key.startswith("Pipeline Metadata"):
            in_block = True
            continue
        if not in_block or key in (None, "key"):
            continue
        meta[str(key)] = val
    return meta


def _meta_num(meta, key):
    """Metadata is written as text; coerce a numeric entry back to a float."""
    v = meta.get(key)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def _cells_in_ranges(expr):
    """Yield (col, row) pairs for every cell covered by a formula's ranges."""
    cells = set()
    for m in re.finditer(r"\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?", expr):
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
        if c2 is None:
            cells.add((c1, r1))
        else:
            for r in range(r1, int(r2) + 1):
                cells.add((c1, r))
    return cells


def _split_ref(ref):
    """('Segment Analysis', 'B70', 'B74') from "'Segment Analysis'!B70:B74"."""
    sheet = None
    if "!" in ref:
        sheet, ref = ref.rsplit("!", 1)
        sheet = sheet.strip("'")
    ref = ref.replace("$", "")
    if ":" in ref:
        a, b = ref.split(":")
    else:
        a = b = ref
    return sheet, a, b


def _proj_columns(ws, header_row=29):
    """Column indices of the projection year headers on the DCF Model sheet."""
    cols = []
    for col in range(3, 20):
        v = ws.cell(row=header_row, column=col).value
        if isinstance(v, str) and v.strip():
            cols.append(col)
        elif cols:
            break
    return cols


# =====================================================================
# checks
# =====================================================================
def check_formula_errors(res, wbv, has_values):
    if not has_values:
        res.add(1, SKIP, "No formula errors", "workbook has no cached values — needs recalc")
        return
    fatal, soft = [], []
    for ws in wbv.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    v = c.value.strip()
                    if v in FATAL_ERRORS:
                        fatal.append(f"{ws.title}!{c.coordinate}={v}")
                    elif v in SOFT_ERRORS:
                        soft.append(f"{ws.title}!{c.coordinate}={v}")
    if fatal:
        res.add(1, FAIL, "No formula errors",
                f"{len(fatal)} error cell(s): {', '.join(fatal[:8])}")
    elif soft:
        res.add(1, WARN, "No formula errors",
                f"no fatal errors; {len(soft)} soft error(s): {', '.join(soft[:8])}")
    else:
        res.add(1, PASS, "No formula errors", "0 error cells in all sheets")


def check_scenario_index(res, wbf, meta):
    """#2 scenario index is a MATCH, and #5 the range really holds the 5 names."""
    if "DCF Model" not in wbf.sheetnames:
        res.add(2, FAIL, "Scenario index is MATCH", "no 'DCF Model' sheet")
        res.add(5, SKIP, "Scenario names are the fixed 5", "no 'DCF Model' sheet")
        return
    ws = wbf["DCF Model"]
    f = ws["D27"].value
    if not (isinstance(f, str) and f.upper().startswith("=MATCH(")):
        res.add(2, FAIL, "Scenario index is MATCH",
                f"'DCF Model'!D27 = {f!r} — the dropdown does nothing")
        res.add(5, SKIP, "Scenario names are the fixed 5", "scenario index not a MATCH")
        return
    res.add(2, PASS, "Scenario index is MATCH", f"D27 = {f}")

    m = re.match(r"=MATCH\(\s*[^,]+,\s*(.+?)\s*,\s*0\s*\)\s*$", f, re.I)
    if not m:
        res.add(5, WARN, "Scenario names are the fixed 5",
                "could not parse the MATCH lookup range")
        return
    ref = m.group(1)
    if ref.startswith("{"):
        res.add(5, FAIL, "Scenario names are the fixed 5",
                "MATCH uses an inline array constant, not the sheet's own "
                "scenario-name cells — edits to the matrix would not be picked up")
        return
    sheet, a, b = _split_ref(ref)
    target = wbf[sheet] if sheet and sheet in wbf.sheetnames else ws
    col = re.match(r"([A-Z]+)", a).group(1)
    r1, r2 = int(re.search(r"(\d+)", a).group(1)), int(re.search(r"(\d+)", b).group(1))
    names = [target[f"{col}{r}"].value for r in range(r1, r2 + 1)]
    if names == SCENARIO_NAMES:
        res.add(5, PASS, "Scenario names are the fixed 5", f"{ref} = {names}")
    else:
        res.add(5, FAIL, "Scenario names are the fixed 5",
                f"{ref} holds {names!r}, expected {SCENARIO_NAMES!r}")


def _subject_row(wbf, meta):
    r = meta.get("comps_subject_row")
    try:
        return int(r)
    except (TypeError, ValueError):
        pass
    # Fallback for workbooks generated before the metadata block existed:
    # match the Executive Summary ticker against the comps table.
    if "Comps Analysis" not in wbf.sheetnames or "Executive Summary" not in wbf.sheetnames:
        return None
    exec_t = str(wbf["Executive Summary"]["C7"].value or "").split(" ")[0]
    key = exec_t.strip().upper().split(".")[0]
    ws = wbf["Comps Analysis"]
    for r in range(5, 40):
        t = ws.cell(row=r, column=3).value
        if t and str(t).strip().upper().split(".")[0] == key:
            return r
    return None


def _stat_rows(wbf, meta):
    med = meta.get("comps_stat_median_row")
    try:
        med = int(med)
        return [med - 1, med, med + 1]
    except (TypeError, ValueError):
        return [15, 16, 17]


def check_stats_exclude_subject(res, wbf, meta):
    if "Comps Analysis" not in wbf.sheetnames:
        res.add(3, SKIP, "Comps stats exclude subject row", "no 'Comps Analysis' sheet")
        return
    ws = wbf["Comps Analysis"]
    subj = _subject_row(wbf, meta)
    if subj is None:
        res.add(3, WARN, "Comps stats exclude subject row",
                "subject row could not be identified")
        return
    offenders = []
    for r in _stat_rows(wbf, meta):
        for c in range(3, 20):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and ("MEDIAN(" in v.upper() or "PERCENTILE(" in v.upper()):
                if any(row == subj for _, row in _cells_in_ranges(v)):
                    offenders.append(f"{ws.cell(row=r, column=c).coordinate}: {v}")
    if offenders:
        res.add(3, FAIL, "Comps stats exclude subject row",
                f"subject row {subj} is inside {len(offenders)} statistic(s): "
                f"{offenders[0]}")
    else:
        res.add(3, PASS, "Comps stats exclude subject row",
                f"subject row {subj} excluded from all MEDIAN/PERCENTILE ranges")


def check_subject_row_formulas(res, wbf, meta):
    if "Comps Analysis" not in wbf.sheetnames:
        res.add(4, SKIP, "Subject Mkt Cap / EV are formulas", "no 'Comps Analysis' sheet")
        return
    ws = wbf["Comps Analysis"]
    subj = _subject_row(wbf, meta)
    if subj is None:
        res.add(4, WARN, "Subject Mkt Cap / EV are formulas",
                "subject row could not be identified")
        return
    mc, ev = ws.cell(row=subj, column=4).value, ws.cell(row=subj, column=5).value
    bad = []
    if not (isinstance(mc, str) and "Executive Summary" in mc):
        bad.append(f"Mkt Cap D{subj} = {mc!r} (expected a formula off "
                   f"'Executive Summary' price x shares)")
    if not (isinstance(ev, str) and ev.startswith("=") and f"D{subj}" in ev):
        bad.append(f"EV E{subj} = {ev!r} (expected =D{subj}+net debt)")
    if bad:
        res.add(4, FAIL, "Subject Mkt Cap / EV are formulas", "; ".join(bad))
    else:
        res.add(4, PASS, "Subject Mkt Cap / EV are formulas",
                f"D{subj}={mc} / E{subj}={ev}")


def check_ltm_revenue(res, wbf, wbv, meta, has_values):
    try:
        expected = float(meta.get("ltm_revenue_c20"))
    except (TypeError, ValueError):
        expected = None
    if expected is None:
        res.add(6, SKIP, "C20 LTM Revenue matches generator", "no metadata")
        return
    actual = _num(wbf["DCF Model"]["C20"].value)
    if actual is None and has_values:
        actual = _num(wbv["DCF Model"]["C20"].value)
    if actual is None:
        res.add(6, FAIL, "C20 LTM Revenue matches generator", "C20 is empty or text")
        return
    src = meta.get("ltm_revenue_source", "?")
    if _close(actual, expected):
        res.add(6, PASS, "C20 LTM Revenue matches generator",
                f"{actual:,.1f} mn (source: {src})")
    else:
        res.add(6, FAIL, "C20 LTM Revenue matches generator",
                f"sheet {actual:,.1f} vs generator {expected:,.1f} mn")


BASIS_WORDS = {
    "explicit_override": "explicit assumption from overrides",
    "auto_hist_avg": "auto-derived from the filings (mean of the FY ratios)",
    "hist_3yr_avg": "mean of the last 3 historical years",   # pre-フェーズ2 models
    "assumption": "explicit assumption",                     # pre-フェーズ2 models
}


def check_capex_da_ratios(res, wbf, meta):
    """#7 C5 / C18 are what the generator meant, and the basis is named.

    フェーズ2 #10: the basis no longer depends on capex_method. It used to be
    replaced by the 3-year historical mean whenever the method was "direct",
    which made the displayed basis flip between "explicit assumption" and
    "historical mean" according to whether hist_capex happened to be in the
    overrides. The historical mean is still reported here as a cross-check
    against the assumption, but it is no longer the assumption.
    """
    ws = wbf["DCF Model"]
    for n, cell, key, basis_key, hist_key, label in (
        (7, "C5", "capex_pct_c5", "capex_pct_basis", "capex_pct_hist3yr", "Capex/Revenue"),
        (7, "C18", "da_pct_c18", "da_pct_basis", "da_pct_hist3yr", "D&A/Revenue"),
    ):
        try:
            expected = float(meta.get(key))
        except (TypeError, ValueError):
            res.add(n, SKIP, f"{cell} {label} basis", "no metadata")
            continue
        actual = _num(ws[cell].value)
        basis = str(meta.get(basis_key, "?"))
        hist3 = _meta_num(meta, hist_key)
        cross = f"; hist 3yr mean {hist3:.2%}" if hist3 is not None else ""
        if actual is None:
            res.add(n, FAIL, f"{cell} {label} basis", f"{cell} is empty or text")
        elif not _close(actual, expected):
            res.add(n, FAIL, f"{cell} {label} basis",
                    f"sheet {actual:.4%} vs generator {expected:.4%}")
        else:
            res.add(n, PASS, f"{cell} {label} basis",
                    f"{actual:.2%} - {BASIS_WORDS.get(basis, basis)}{cross}")


def check_fs_year_alignment(res, wbf, meta):
    ymap = meta.get("fs_year_map")
    if not ymap or "Financial Statements" not in wbf.sheetnames:
        res.add(8, SKIP, "FS OCF/Cash/Debt year alignment", "no metadata / sheet")
        return
    ws = wbf["Financial Statements"]
    headers = {}
    for col in range(3, 25):
        v = ws.cell(row=4, column=col).value
        if v:
            headers[str(v)] = col
    # Row of each series on the Financial Statements sheet
    series_rows = {"ocf": 19, "cash": 26, "debt": 27}
    problems, checked = [], 0
    for entry in str(ymap).split(";"):
        entry = entry.strip()
        if "<-" not in entry:
            continue
        label, rest = entry.split("<-", 1)
        label = label.strip()
        m = re.match(r"(.*?)\s*\[(.*)\]\s*$", rest.strip())
        if m:
            src, filled = m.group(1), [f for f in m.group(2).split("+") if f]
        else:  # metadata written before the per-series audit existed
            src, filled = rest.strip(), None
        col = headers.get(label)
        if col is None:
            problems.append(f"{label}: no column on the FS sheet")
            continue
        checked += 1
        for name, row in series_rows.items():
            present = ws.cell(row=row, column=col).value is not None
            if filled is None:
                continue
            if present and name not in filled:
                problems.append(f"{label}/{name}: filled on the sheet but not in "
                                f"the source year ({src})")
            if not present and name in filled:
                problems.append(f"{label}/{name}: sourced from {src} but blank "
                                f"on the sheet")
        if filled is None and src == "BLANK":
            if any(ws.cell(row=r, column=col).value is not None
                   for r in series_rows.values()):
                problems.append(f"{label}: unmatched year but OCF/Cash/Debt filled")
    if problems:
        res.add(8, FAIL, "FS OCF/Cash/Debt year alignment", "; ".join(problems[:4]))
    else:
        res.add(8, PASS, "FS OCF/Cash/Debt year alignment",
                f"{checked} year column(s) match the source year keys cell-for-cell")


def check_terminal_capex(res, wbf, wbv, has_values):
    if not has_values:
        res.add(9, SKIP, "Terminal-year capex / D&A ratio", "needs recalc")
        return
    wsf, wsv = wbf["DCF Model"], wbv["DCF Model"]
    g = _num(wsv["C13"].value) if _num(wsv["C13"].value) is not None else _num(wsf["C13"].value)
    cols = _proj_columns(wsf)
    if not cols or g is None:
        res.add(9, SKIP, "Terminal-year capex / D&A ratio", "no projection columns")
        return
    last = cols[-1]
    da = _num(wsv.cell(row=R_DA, column=last).value)
    capex = _num(wsv.cell(row=R_CAPEX, column=last).value)
    if not da or capex is None:
        res.add(9, SKIP, "Terminal-year capex / D&A ratio", "no cached values")
        return
    ratio = abs(capex) / abs(da)
    if g > 0.015:
        res.add(9, PASS, "Terminal-year capex / D&A ratio",
                f"g={g:.2%} > 1.5% — steady-state test not applicable "
                f"(ratio {ratio:.2f}x)")
    elif 0.90 <= ratio <= 1.15:
        res.add(9, PASS, "Terminal-year capex / D&A ratio",
                f"{ratio:.2f}x within [0.90, 1.15] at g={g:.2%}")
    else:
        res.add(9, WARN, "Terminal-year capex / D&A ratio",
                f"{ratio:.2f}x outside [0.90, 1.15] at g={g:.2%} — perpetuity "
                f"assumes this capex forever (analyst call, not auto-fixed)")


def check_pgm_negative_equity(res, wbf, wbv, has_values):
    if not has_values:
        res.add(10, SKIP, "PGM implied price sanity", "needs recalc")
        return
    wsv = wbv["DCF Model"]
    ev = _num(wsv.cell(row=R_EV_PGM, column=3).value)
    nd = _num(wsv["C16"].value)
    if ev is None or nd is None:
        res.add(10, SKIP, "PGM implied price sanity", "no cached EV / net debt")
        return
    if ev >= nd:
        res.add(10, PASS, "PGM implied price sanity",
                f"EV {ev:,.0f} >= net debt {nd:,.0f} mn")
        return
    label = wbv["Executive Summary"]["C16"].value
    if isinstance(label, str) and "INVALID" in label.upper():
        res.add(10, PASS, "PGM implied price sanity",
                f"EV {ev:,.0f} < net debt {nd:,.0f} and the method is labelled "
                f"{label!r} (text -> skipped by AVERAGE)")
    else:
        res.add(10, FAIL, "PGM implied price sanity",
                f"EV {ev:,.0f} < net debt {nd:,.0f} but Executive Summary C16 = "
                f"{label!r} — a negative-equity artefact is being averaged into "
                f"the Target Price")


def check_target_excludes_comps(res, wbf):
    """#14 The Target must be built from the valuation methods, not from comps.

    The standard (docs/DCFフォーマット標準メモ §2) is that comps are a reference
    mark, not a target input. The old formula averaged C16:C19, so every model
    silently shipped a half-comps target - a violation nobody could see on the
    sheet because the cell just shows a number.

    Two things were wrong with the original implementation and are fixed here:

    1. It resolved the Target at C10 and the comps at C18/C19 by row NUMBER.
       That holds for 型A-C but not for 型D, where the DDM and Residual Income
       rows are inserted at 18/19 and the comps move down - a row-number rule
       calls the bank's own primary methods "comps" and fails a correct model.
       Rows are now resolved by label.

    2. Its single-reference pattern was `r"\bC1[6-9]\b"`, but the file carried
       literal backspace characters where the \b escapes should have been (this
       predates フェーズ2 - it is in c9b5dd9), so that pattern never matched
       anything. Only the range pattern `C1[6-9]:C1[6-9]` was live, which means
       a formula listing comps individually - `=AVERAGE(C16,C17,C18)` - passed
       the check that exists to catch exactly that. Cell references are now
       parsed properly and ranges are expanded.
    """
    if "Executive Summary" not in wbf.sheetnames:
        res.add(14, SKIP, "Target Price averages DCF legs only", "no Executive Summary")
        return
    ws = wbf["Executive Summary"]
    r_tgt, comps_rows = None, []
    for r in range(1, 44):
        v = ws.cell(r, 2).value
        if not isinstance(v, str):
            continue
        if r_tgt is None and v.startswith("Target Price"):
            r_tgt = r
        if v.startswith("Comps -"):
            comps_rows.append(r)
    if r_tgt is None:
        res.add(14, SKIP, "Target Price averages DCF legs only",
                "no 'Target Price' row on the Executive Summary")
        return
    f = ws.cell(r_tgt, 3).value
    if not isinstance(f, str) or not f.startswith("="):
        res.add(14, FAIL, "Target Price averages DCF legs only",
                f"C{r_tgt} is not a formula ({f!r}) - a hardcoded target cannot "
                f"track a price or scenario change")
        return

    cited = set()
    for m in re.finditer(r"C(\d+)(?::C(\d+))?", f):
        a = int(m.group(1))
        b = int(m.group(2)) if m.group(2) else a
        cited.update(range(min(a, b), max(a, b) + 1))
    hit = sorted(cited & set(comps_rows))
    if hit:
        names = ", ".join(f"C{r} ({ws.cell(r, 2).value})" for r in hit)
        res.add(14, FAIL, "Target Price averages DCF legs only",
                f"C{r_tgt} = {f} references the comps rows [{names}] - comps must "
                f"stay [参考] and out of the Target average")
        return
    if "C16:C17" in f:
        res.add(14, PASS, "Target Price averages DCF legs only",
                f"C{r_tgt} averages C16:C17 (PGM + Exit); "
                f"comps rows {comps_rows or 'none'} excluded")
    else:
        # A demoted leg (追補6 §X) or a 型D DDM/RI target legitimately looks
        # different; the check has already proved no comps row is cited.
        res.add(14, WARN, "Target Price averages DCF legs only",
                f"C{r_tgt} = {f} does not average C16:C17 - expected when a leg "
                f"was demoted (追補6 §X) or the Target is a 型D DDM/RI average. "
                f"No comps row is referenced.")


def check_exit_negative_equity(res, wbf, wbv, has_values):
    """#15 The Exit leg needs the same EV < net debt guard as the PGM leg.

    Found on 5726: under Downside 2 the PGM leg correctly went INVALID while the
    Exit leg's -558 stayed a plain number and dragged the Target average down.
    """
    if not has_values:
        res.add(15, SKIP, "Exit implied price sanity", "needs recalc")
        return
    wsv = wbv["DCF Model"]
    ev = _num(wsv.cell(row=R_EV_EXIT, column=3).value)
    nd = _num(wsv["C16"].value)
    if ev is None or nd is None:
        res.add(15, SKIP, "Exit implied price sanity", "no cached EV / net debt")
        return
    if ev >= nd:
        res.add(15, PASS, "Exit implied price sanity",
                f"EV {ev:,.0f} >= net debt {nd:,.0f} mn")
        return
    label = wbv["Executive Summary"]["C17"].value
    if isinstance(label, str) and "INVALID" in label.upper():
        res.add(15, PASS, "Exit implied price sanity",
                f"EV {ev:,.0f} < net debt {nd:,.0f} and the method is labelled "
                f"{label!r} (text -> skipped by AVERAGE)")
    else:
        res.add(15, FAIL, "Exit implied price sanity",
                f"EV {ev:,.0f} < net debt {nd:,.0f} but Executive Summary C17 = "
                f"{label!r} — a negative-equity artefact is being averaged into "
                f"the Target Price")







def _company_type(path):
    """'A'..'E' from the ticker's overrides, or None."""
    try:
        from scripts.arbitration import resolve_company_type
    except ImportError:
        return None
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ov = os.path.join(root, "data", "overrides",
                      f"{os.path.basename(path)[:4]}_overrides.json")
    if not os.path.isfile(ov):
        return None
    try:
        return resolve_company_type(json.load(open(ov, encoding="utf-8")))
    except (OSError, ValueError):
        return None


def check_bank_model(res, path, wbf, wbv, has_values):
    """#24 型D (bank): the DDM/RI model is present, coherent, and drives the Target.

    手順書 §2 says a bank's DCF does not hold and its Target is the average of a
    dividend discount model and a residual income model. That is a structural
    claim about the workbook, so it is checked structurally: the sheets exist,
    Ke exceeds g (without which the Gordon terminal is undefined), the dividend
    path is not a payout above earnings, and the Target cell really does average
    the two - not the DCF legs.

    Non-型D models pass immediately: this says nothing about them.
    """
    ctype = _company_type(path)
    if ctype != "D":
        res.add(24, PASS, "型D DDM/RI model",
                f"対象外（company_type={ctype or '未宣言'}）")
        return
    missing = [s for s in ("DDM", "Residual Income") if s not in wbf.sheetnames]
    if missing:
        res.add(24, FAIL, "型D DDM/RI model",
                f"型D なのに {', '.join(missing)} シートが無い。"
                f"scripts/ddm_ri.py で生成すること")
        return
    if not has_values:
        res.add(24, SKIP, "型D DDM/RI model", "needs recalc")
        return
    ddm, ri = wbv["DDM"], wbv["Residual Income"]

    def labelled(ws, prefix, limit=60):
        for r in range(1, limit):
            v = ws.cell(r, 2).value
            if isinstance(v, str) and v.startswith(prefix):
                return _num(ws.cell(r, 3).value), r
        return None, None

    ke, _ = labelled(ddm, "Cost of Equity")
    g, _ = labelled(ddm, "Terminal Growth")
    ddm_v, _ = labelled(ddm, "Implied Value per Share")
    ri_v, _ = labelled(ri, "Implied Value per Share")
    bps0, _ = labelled(ri, "BPS_0")
    problems, notes = [], []
    if ke is None or g is None:
        problems.append("DDM の Ke / g を読めない")
    elif ke <= g:
        problems.append(f"Ke {ke:.2%} <= g {g:.2%} — ゴードン成長式が成立しない")
    else:
        notes.append(f"Ke {ke:.2%} > g {g:.2%}")
    for nm, v in (("DDM", ddm_v), ("Residual Income", ri_v)):
        if v is None or v <= 0:
            problems.append(f"{nm} の1株価値が {v!r}")
    # Payout sanity: dividends must not exceed the earnings that fund them.
    dps_row = ni_row = div_row = None
    for r in range(1, 60):
        v = ri.cell(r, 2).value
        if isinstance(v, str):
            if v.startswith("Net Income_t"):
                ni_row = r
            elif v.startswith("Dividends_t"):
                div_row = r
    if ni_row and div_row:
        bad = []
        for c in range(3, 8):
            ni, dv = _num(ri.cell(ni_row, c).value), _num(ri.cell(div_row, c).value)
            if ni and dv is not None:
                p = dv / ni
                if p < 0 or p > 1.0:
                    bad.append(f"Y{c - 2} payout {p:.0%}")
        if bad:
            problems.append("配当性向が 0〜100% の外: " + ", ".join(bad))
        else:
            notes.append("配当性向は全予測年で 0〜100% の内側")
    # Target must be the DDM/RI average, not a DCF leg.
    if "Executive Summary" in wbf.sheetnames:
        wsf = wbf["Executive Summary"]
        tgt_f = None
        for r in range(1, 44):
            v = wsf.cell(r, 2).value
            if isinstance(v, str) and v.startswith("Target Price"):
                tgt_f = str(wsf.cell(r, 3).value or "")
                break
        if tgt_f and "AVERAGE" not in tgt_f.upper():
            problems.append(f"Target が DDM/RI の平均を参照していない: {tgt_f[:60]}")
        elif tgt_f:
            notes.append("Target = AVERAGE(DDM, Residual Income)")
    if problems:
        res.add(24, FAIL, "型D DDM/RI model", "; ".join(problems))
        return
    extra = ""
    if ri_v and bps0:
        extra = f"; RI {ri_v:,.0f} vs BPS0 {bps0:,.0f} ({ri_v / bps0:.2f}x book)"
    res.add(24, PASS, "型D DDM/RI model",
            f"DDM {ddm_v:,.0f} / RI {ri_v:,.0f} → Target {(ddm_v + ri_v) / 2:,.0f}"
            f"; " + "; ".join(notes) + extra)


def check_type_f_model(res, path, wbf, wbv, has_values):
    """#26 型F: Target = コアDCF + 持分法投資価値（1株）が実際にそうなっているか。"""
    meta = {}
    if "Adjustments Log" in wbf.sheetnames:
        ws = wbf["Adjustments Log"]
        for r in range(1, ws.max_row + 1):
            k = ws.cell(r, 2).value
            if isinstance(k, str):
                meta[k.strip()] = ws.cell(r, 3).value
    declared = str(meta.get("company_type", "")).strip().upper()
    has_sheet = "Equity Method Value" in wbf.sheetnames
    if declared != "F" and not has_sheet:
        res.add(26, PASS, "型F 持分法投資価値の別途加算",
                f"対象外（company_type={declared or '未宣言'}）")
        return
    if not has_sheet:
        res.add(26, FAIL, "型F 持分法投資価値の別途加算",
                "company_type=F だが 'Equity Method Value' シートが無い — "
                "持分法投資価値が Target に入っていない")
        return

    es = wbf["Executive Summary"]
    r_tgt = r_add = r_core = None
    for r in range(1, 48):
        v = es.cell(r, 2).value
        if not isinstance(v, str):
            continue
        if r_core is None and v.startswith("コアDCF 1株値"):
            r_core = r
        if r_tgt is None and v.startswith("Target Price"):
            r_tgt = r
        # 行【ラベル】だけを見る。注記文にも同じ語が出てくるため、含有判定だと
        # 注記行を2つ目の加算行と誤認する（8058 の初回生成で実際に誤検出した）。
        if v.endswith("[1株・別途加算]"):
            if r_add is not None:
                res.add(26, FAIL, "型F 持分法投資価値の別途加算",
                        f"加算行が2つある（C{r_add} と C{r}）— 再生成で二重配線され、"
                        f"持分法投資価値を二度足している")
                return
            r_add = r
    if r_tgt is None or r_add is None:
        res.add(26, FAIL, "型F 持分法投資価値の別途加算",
                f"Executive Summary に Target 行（{r_tgt}）または加算行（{r_add}）が無い")
        return
    f = es.cell(r_tgt, 3).value
    if not isinstance(f, str) or f"C{r_add}" not in f:
        res.add(26, FAIL, "型F 持分法投資価値の別途加算",
                f"Target C{r_tgt} = {f!r} が加算行 C{r_add} を参照していない — "
                f"シートは作られているが Target に入っていない")
        return

    method = str(meta.get("equity_method_method", "")).strip()
    bal = meta.get("equity_method_balance_mn")
    mult = meta.get("equity_method_multiple")
    detail = f"Target C{r_tgt} = {f} に加算行 C{r_add} が入っている; 方式 {method or '?'}"
    if not has_values:
        res.add(26, SKIP, "型F 持分法投資価値の別途加算", "needs recalc")
        return
    emv = wbv["Equity Method Value"]
    per_share = tgt = add = None
    for r in range(1, emv.max_row + 1):
        lab = emv.cell(r, 2).value
        if isinstance(lab, str) and lab.startswith("1株あたり持分法投資価値"):
            per_share = emv.cell(r, 3).value
    esv = wbv["Executive Summary"]
    tgt, add = esv.cell(r_tgt, 3).value, esv.cell(r_add, 3).value
    if per_share is None or add is None:
        res.add(26, FAIL, "型F 持分法投資価値の別途加算",
                "1株あたり持分法投資価値が計算されていない（recalc 未実施か数式エラー）")
        return
    if abs(float(per_share) - float(add)) > 1:
        res.add(26, FAIL, "型F 持分法投資価値の別途加算",
                f"加算行 {add:,.0f} が Equity Method Value の 1株値 {per_share:,.0f} と"
                f"一致しない")
        return
    if method == "book_value" and mult is not None and abs(float(mult) - 1.0) > 1e-9:
        res.add(26, WARN, "型F 持分法投資価値の別途加算",
                detail + f"; ただし book_value 方式で倍率が {mult} — 設計は 1.0 に統一")
        return
    extra = ""
    if isinstance(tgt, (int, float)) and float(tgt) != 0:
        extra = (f"; Target {tgt:,.0f} のうち加算脚 {add:,.0f} "
                 f"({float(add)/float(tgt):.0%})")
    if isinstance(bal, (int, float)):
        extra += f"; BS残高 {bal:,.0f} mn"

    # 追補15 A-1 §3 のフロアが効いた（コア = 0）なら、Target は加算脚のみである。
    # これは「コアDCF が成立していない」という重い事実なので黙って通してはいけない。
    if r_core is not None:
        core_v = esv.cell(r_core, 3).value
        if isinstance(core_v, (int, float)) and float(core_v) == 0:
            res.add(26, WARN, "型F 持分法投資価値の別途加算",
                    detail + extra +
                    "; **コアDCF 不成立** — コア脚のエクイティ（EV − net_debt）が負または"
                    "算定不能でフロア規則により 0 とした。Target は加算脚のみで構成されており、"
                    "事業そのものの価値は含まれていない。§AF の相対評価として読むこと"
                    + ("; さらに加算脚は book_value（BS残高×1.0）で、上場持分の時価置換が"
                       "できていないため過小側" if method == "book_value" else ""))
            return
    # 追補15 A-1 §1: 上場持分先は【時価 × 比率】が原則で、「残高 × 1.0」は
    # 上場分の簿価が注記から取れないときの最終手段である。採ったときは保守的である
    # ことを表に出す（黙って簿価のままだと、含み益のある持分がゼロ評価されていることに
    # 気づけない）。
    if method == "book_value":
        res.add(26, WARN, "型F 持分法投資価値の別途加算",
                detail + extra +
                "; 方式が book_value（BS残高×1.0）— 上場持分先の【簿価】が有報の注記から"
                "取得できないため時価への置換ができていない。上場持分に含み益がある分だけ"
                "**保守的（過小）**に出る。個別簿価が取れたら listed_stakes へ切り替えること")
        return
    res.add(26, PASS, "型F 持分法投資価値の別途加算", detail + extra)


def check_beta_clamped(res, wbf, meta):
    """#25 A clamped beta is a substituted assumption — say so out loud.

    フェーズ2 #6 の規則は「生ベータ → Blume 収縮 → [0.3, 2.0] でクランプ、
    クランプしたら WARN」。テンプレートはクランプを実行して Adjustments Log に
    記録するが、その記録は誰も読んでいなかった。クランプは WACC を直接動かすので、
    レポートの表に出ないと「測定値」と「下限に張り付いた代替値」の区別がつかない。
    """
    log = {}
    if "Adjustments Log" in wbf.sheetnames:
        ws = wbf["Adjustments Log"]
        for r in range(1, ws.max_row + 1):
            k = ws.cell(r, 2).value
            if isinstance(k, str):
                log[k.strip()] = ws.cell(r, 3).value
    if "beta_clamped" not in log:
        res.add(25, PASS, "Beta clamp not silently applied",
                "Adjustments Log にベータ記録なし（旧テンプレートの生成物）")
        return
    clamped = str(log.get("beta_clamped", "")).strip().lower() in ("yes", "true", "1")
    basis = log.get("beta_basis") or ""
    adopted = log.get("beta_adopted_c8")
    if clamped:
        res.add(25, WARN, "Beta clamp not silently applied",
                f"ベータが [0.3, 2.0] にクランプされて {adopted} が採用された — "
                f"これは測定値ではなく代替値であり WACC を直接動かす。{basis} / "
                f"生ベータの出所を確認し、必要なら TOPIX 2年週次回帰の値を "
                f"overrides の beta に明示すること")
    else:
        res.add(25, PASS, "Beta clamp not silently applied",
                f"クランプなし（採用 {adopted}）。{basis}")


def check_disclosure_vintage(res, path, meta):
    """#23 Which disclosures the model is built on, and how old they are.

    追補13 §A relaxed the freshness rule for the queue remainder: a model may be
    built on the newest CONFIRMED filing even when that filing is a year old.
    The price of the relaxation is that the vintage must be unambiguous, so this
    reports the basis date and the docIDs, and WARNs - never FAILs - when the
    newest annual report is more than a year behind the analysis date.

    This is NOT the 追補6 §Z freshness gate. That one asks "is the workbook newer
    than the overrides and comps it was built from", which is about the build
    being current, and it still fails. This one is about the DATA being old,
    which is now a disclosed condition rather than a blocker.
    """
    basis = meta.get("disclosure_basis_date")
    docs = meta.get("annual_doc_ids")
    interim = meta.get("interim_doc_id")
    if not basis:
        res.add(23, PASS, "Disclosure vintage recorded",
                "no disclosure metadata (pre-追補13 model or overrides-only build)")
        return
    import datetime as _dt
    # The analysis-basis date is the filename stamp (--date), not today: a model
    # re-validated months later must not age just because the clock moved.
    stamp = re.search(r"_(\d{8})\.xlsx$", os.path.basename(path))
    try:
        asof = _dt.date(int(stamp.group(1)[:4]), int(stamp.group(1)[4:6]),
                        int(stamp.group(1)[6:])) if stamp else _dt.date.today()
        b = _dt.date.fromisoformat(str(basis)[:10])
    except (ValueError, AttributeError):
        res.add(23, PASS, "Disclosure vintage recorded", f"basis {basis}")
        return
    months = (asof.year - b.year) * 12 + (asof.month - b.month)
    detail = (f"最新の確定年次開示 {b.isoformat()}（分析基準日から {months} か月前）"
              f" / 年次 docID: {docs}"
              + (f" / 半期 docID: {interim}" if interim else ""))
    if months > 12:
        res.add(23, WARN, "Disclosure vintage recorded",
                detail + " — 追補13 §A の鮮度緩和が適用された状態。"
                         "最新期の確定開示が出たら再生成すること")
    else:
        res.add(23, PASS, "Disclosure vintage recorded", detail)


def check_arbitration_applied(res, path, wbf, meta):
    """#22 A model the machine rules say to arbitrate has actually been arbitrated.

    追補5 forbids averaging two irreconcilable DCF legs, and 追補6 §X decides
    which one to drop. That decision used to be applied by a separate script
    after the batch, so a regeneration silently reverted it and the workbook went
    back to the forbidden midpoint - 6857 アドバンテスト's Target moved 4,288 ->
    8,021 that way, and nothing failed. The generator now arbitrates before
    validating (追補12 §A-3); this check is the lock that says so.

    型D (banks) are exempt by construction: a DCF does not hold for them, their
    Target is DDM + Residual Income, and arbitrating two DCF legs there would
    arbitrate between two numbers nobody uses.
    """
    try:
        from scripts.arbitration import (arbitrate, resolve_company_type,
                                         arbitration_applies)
    except ImportError:
        res.add(22, SKIP, "追補6 §X arbitration applied", "scripts.arbitration unavailable")
        return
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    code = os.path.basename(path)[:4]
    ov = os.path.join(root, "data", "overrides", f"{code}_overrides.json")
    overrides = None
    if os.path.isfile(ov):
        try:
            overrides = json.load(open(ov, encoding="utf-8"))
        except (OSError, ValueError):
            overrides = None
    ctype = resolve_company_type(overrides)
    ok, why = arbitration_applies(ctype)
    if not ok:
        res.add(22, PASS, "追補6 §X arbitration applied", f"対象外: {why}")
        return

    r = arbitrate(path,
                  overrides_path=ov if os.path.isfile(ov) else None,
                  comps_csv=os.path.join(root, "data", "comps", f"{code}_comps.csv"),
                  cache_path=os.path.join(root, "batch", "cache", f"{code}.json"),
                  code=code)
    if r is None:
        res.add(22, SKIP, "追補6 §X arbitration applied",
                "乖離を算出できない（ターミナル価値 / Year5 EBITDA / Exit倍率のキャッシュ値なし）")
        return
    if r["treated"]:
        res.add(22, PASS, "追補6 §X arbitration applied",
                f"適用済 {r['treated'][1]}: {r['treated'][0].upper()} 脚を降格 "
                f"(乖離 {r['div']:.2f}x, {r['regime']})")
        return
    if r["demote"]:
        res.add(22, FAIL, "追補6 §X arbitration applied",
                f"機械化ルールは「{r['verdict']}」と判定しているが未適用。"
                f"乖離 {r['div']:.2f}x > 3.0x で中点平均は追補5 が禁じている。"
                f"`python batch/apply_arbitration.py --write {code}` を実行し "
                f"recalc してから再検証すること")
        return
    res.add(22, PASS, "追補6 §X arbitration applied",
            f"降格不要: {r['regime']} / {r['verdict'][:60]} (乖離 {r['div']:.2f}x)")


def check_market_data(res, wbf, meta):
    """#21 Price and share count are real numbers, not the template placeholders.

    generate_dcf.py used to seed the config with price=1,000 and
    shares=10,000,000 "to be overridden by yfinance", and the yfinance helper
    returned those unchanged whenever the lookup threw. 4568 第一三共 shipped
    Target JPY 294,427 / BUY +293% on a market cap of JPY 10,000 mn (true:
    5,084,100 mn) and validated FAIL 0. The generator now refuses to run in that
    state; this check is the second lock, and it also covers hand-edited
    workbooks the generator never saw.
    """
    if "Executive Summary" not in wbf.sheetnames or "DCF Model" not in wbf.sheetnames:
        res.add(21, SKIP, "Market data is not a placeholder", "missing sheet")
        return
    price = _num(wbf["Executive Summary"]["C9"].value)
    shares = _num(wbf["DCF Model"]["C15"].value)
    src = str(meta.get("market_data_source", "not recorded"))
    problems = []
    if price is None or price <= 0:
        problems.append(f"current price is {price!r}")
    if shares is None or shares <= 0:
        problems.append(f"fully diluted shares is {shares!r}")
    if not problems and float(price) == 1000.0 and int(shares) == 10_000_000:
        problems.append("price 1,000 x shares 10,000,000 - the exact template "
                        "placeholder pair (market cap JPY 10,000 mn)")
    if problems:
        res.add(21, FAIL, "Market data is not a placeholder",
                "; ".join(problems) + f" [source: {src}]")
    else:
        res.add(21, PASS, "Market data is not a placeholder",
                f"price {price:,.0f} x {shares:,} shares = "
                f"{price * shares / 1_000_000:,.0f} JPY mn [source: {src}]")


def check_core_ebitda(res, wbf, meta):
    """#19 The subject EBITDA behind the Comps legs is coherent with the P/L.

    core_ebitda = latest operating income + latest D&A. When EDINET failed to
    supply one of those legs the sum used to be written anyway, and the Comps
    reference price built on it went negative or N/A while the run still
    reported FAIL 0 (4502: JPY -2,761 per share). EBITDA cannot be below
    operating income - depreciation is not negative - so the two numbers test
    each other.
    """
    eb = _meta_num(meta, "core_ebitda")
    oi = _meta_num(meta, "latest_operating_income")
    excluded = str(meta.get("comps_ebitda_excluded", "")).lower() == "yes"
    if oi is None:
        res.add(19, SKIP, "Subject EBITDA vs operating income",
                "no latest_operating_income in metadata")
        return
    if eb is None:
        if oi > 0 and not excluded:
            res.add(19, FAIL, "Subject EBITDA vs operating income",
                    f"operating income is {oi:,.0f} mn but core_ebitda is absent "
                    f"and the EV/EBITDA leg was NOT excluded")
        else:
            res.add(19, PASS, "Subject EBITDA vs operating income",
                    "core_ebitda absent; the EV/EBITDA leg is excluded (N/A)")
        return
    if oi > 0 and eb <= 0:
        res.add(19, FAIL, "Subject EBITDA vs operating income",
                f"operating income {oi:,.0f} mn > 0 but core_ebitda is {eb:,.0f} "
                f"- a D&A leg is missing or the sign is wrong")
    elif oi > 0 and eb < oi * 0.999:
        res.add(19, WARN, "Subject EBITDA vs operating income",
                f"core_ebitda {eb:,.0f} < operating income {oi:,.0f} mn. "
                f"D&A is not negative, so the two are on different bases "
                f"(different period or scope) - confirm it is intended")
    else:
        res.add(19, PASS, "Subject EBITDA vs operating income",
                f"core_ebitda {eb:,.0f} >= operating income {oi:,.0f} mn "
                f"(implied D&A {eb - oi:,.0f})")


def check_comps_reference_band(res, wbf, wbv, has_values):
    """#20 The two Comps reference prices are inside a sane band.

    They are [参考] and never enter the Target, but a reference price of
    JPY -2,761 (4502) or of 100x the market price is not a reference, it is a
    broken input on display. Excluded methods are written as text (N/A /
    INVALID) and pass by construction - that is the designed way out.
    Band: 0 < value < current price x 10.
    """
    if not has_values:
        res.add(20, SKIP, "Comps reference prices in a sane band", "needs recalc")
        return
    if "Executive Summary" not in wbf.sheetnames:
        res.add(20, SKIP, "Comps reference prices in a sane band",
                "no Executive Summary")
        return
    wsv = wbv["Executive Summary"]
    price = _num(wsv["C9"].value)
    if not price or price <= 0:
        res.add(20, SKIP, "Comps reference prices in a sane band",
                "no current price to scale the band")
        return
    ceiling = price * 10
    # comps 行は【ラベル】で解決する。行番号決め打ちは、加算脚(型E/F)や DDM/RI(型D)が
    # Exit 行の下に行を挿入した瞬間に別のセルを読み始める。8058 では挿入された
    # 「コアDCF 1株値(フロア)」= 0 を EV/EBITDA の含意株価だと誤読して FAIL し、
    # 8001/8002/8031/8053 では逆に、コア行と加算行が正の数だったため
    # **comps を一度も見ずに PASS** していた（check 14 が踏んだのと同じ罠）。
    wsf = wbf["Executive Summary"]
    targets = []
    for r in range(1, 48):
        lab = wsf.cell(r, 2).value
        if not isinstance(lab, str) or not lab.startswith("Comps - "):
            continue
        nm = "EV/EBITDA" if lab.startswith("Comps - EV/") else (
             "PER" if "PER" in lab else lab[8:28])
        targets.append((wsv.cell(r, 3), nm))
    if not targets:
        res.add(20, SKIP, "Comps reference prices in a sane band",
                "Executive Summary に 'Comps - ' 行が無い")
        return
    problems, notes = [], []
    for cellobj, label in targets:
        v = cellobj.value
        n = _num(v)
        if n is None:
            notes.append(f"{label}={v!r} (text - method excluded)")
            continue
        if n <= 0:
            problems.append(f"{label} implies {n:,.0f} JPY/share (<= 0)")
        elif n > ceiling:
            notes.append(f"{label}={n:,.0f} ABOVE the band ceiling {ceiling:,.0f}")
        else:
            notes.append(f"{label}={n:,.0f} ok")
    if problems:
        res.add(20, FAIL, "Comps reference prices in a sane band",
                "; ".join(problems) + f" - band is 0 < v < {ceiling:,.0f} "
                f"(price {price:,.0f} x10). An unusable method must be written "
                f"as text (N/A / INVALID), not as a negative number")
    elif any("ABOVE" in n for n in notes):
        res.add(20, WARN, "Comps reference prices in a sane band", "; ".join(notes))
    else:
        res.add(20, PASS, "Comps reference prices in a sane band", "; ".join(notes))


def check_reverse_dcf_sheet(res, wbf, meta):
    """#16 'Reverse DCF' is a standard sheet, placed 4th, and fully live."""
    note = meta.get("reverse_dcf_sheet", "")
    if REVERSE_DCF_SHEET not in wbf.sheetnames:
        if note.startswith("skipped"):
            res.add(16, WARN, "Reverse DCF sheet present",
                    f"not generated — {note}")
        else:
            res.add(16, FAIL, "Reverse DCF sheet present",
                    "the standard 8-sheet layout requires a 'Reverse DCF' sheet "
                    "(regenerate, or record why it was skipped)")
        return
    ws = wbf[REVERSE_DCF_SHEET]
    n_f = sum(1 for row in ws.iter_rows() for c in row
              if isinstance(c.value, str) and c.value.startswith("="))
    pos = wbf.sheetnames.index(REVERSE_DCF_SHEET)
    detail = f"{n_f} live formulas, sheet position {pos + 1}"
    if n_f < 50:
        res.add(16, FAIL, "Reverse DCF sheet present",
                f"only {n_f} live formulas — the sheet must be formula-driven, "
                f"not a snapshot of values")
    elif "DCF Model" in wbf.sheetnames and pos != wbf.sheetnames.index("DCF Model") + 1:
        res.add(16, WARN, "Reverse DCF sheet present",
                f"{detail} — expected directly after 'DCF Model'")
    else:
        res.add(16, PASS, "Reverse DCF sheet present", detail)


def check_cost_of_debt(res, wbf, wbv, meta, has_values):
    """#17 C11 must equal what the generator says it derived.

    The actual-cost-of-debt module rewrites C11 from interest / average debt.
    A model whose C11 no longer matches its own recorded derivation has been
    hand-edited without a log entry — which is exactly the drift the
    Adjustments Log exists to prevent.
    """
    basis = str(meta.get("cost_of_debt_basis") or "").strip()
    if not basis:
        res.add(17, SKIP, "Cost of debt basis", "no metadata (pre-2026-08 model)")
        return
    stated = _meta_num(meta, "cost_of_debt_at_c11")
    cell = _num(wbf["DCF Model"]["C11"].value)
    if cell is None or stated is None:
        res.add(17, SKIP, "Cost of debt basis", "C11 or metadata not numeric")
        return
    if abs(cell - stated) > 1e-9:
        res.add(17, FAIL, "Cost of debt basis",
                f"C11 = {cell:.4%} but the generator recorded {stated:.4%} — "
                f"the cell was edited without updating the Adjustments Log")
        return
    if basis != "actual":
        res.add(17, PASS, "Cost of debt basis",
                f"{cell:.4%} after-tax (assumption; supply interest_expense + "
                f"debt balances to derive the actual rate)")
        return
    interest = _meta_num(meta, "cost_of_debt_interest")
    avg_debt = _meta_num(meta, "cost_of_debt_avg_debt")
    pretax = _meta_num(meta, "cost_of_debt_pretax")
    tax = _meta_num(meta, "tax_rate_c6")
    if None in (interest, avg_debt, pretax, tax) or not avg_debt:
        res.add(17, WARN, "Cost of debt basis",
                f"basis=actual but the derivation inputs are missing from the "
                f"metadata; C11 = {cell:.4%}")
        return
    if not _close(interest / avg_debt, pretax, tol=1e-6):
        res.add(17, FAIL, "Cost of debt basis",
                f"recorded pre-tax {pretax:.6f} != interest {interest:,.0f} / "
                f"avg debt {avg_debt:,.0f} = {interest / avg_debt:.6f}")
        return
    expected = round(pretax * (1 - tax), 4)
    if abs(expected - cell) > 1e-9:
        res.add(17, FAIL, "Cost of debt basis",
                f"C11 = {cell:.4%} but pre-tax {pretax:.4%} x (1-{tax:.1%}) "
                f"rounds to {expected:.4%}")
    else:
        res.add(17, PASS, "Cost of debt basis",
                f"actual: {interest:,.0f} / {avg_debt:,.0f} = {pretax:.4%} "
                f"pre-tax -> {cell:.4%} after-tax")


def check_fx_sensitivity(res, wbf, wbv, meta, has_values):
    """#18 Table 3's centre column must reproduce the model's own Year-1 OP."""
    note = str(meta.get("fx_sensitivity") or "")
    if not note:
        res.add(18, SKIP, "FX sensitivity Table 3", "no metadata (pre-2026-08 model)")
        return
    if note == "not enabled":
        res.add(18, PASS, "FX sensitivity Table 3",
                "not enabled (domestic name) — no table expected")
        return
    if note.startswith("skipped"):
        res.add(18, FAIL, "FX sensitivity Table 3",
                f"enabled but not generated — {note}")
        return
    m = re.search(r"rows (\d+)-(\d+)", note)
    if not m or "Sensitivity Analysis" not in wbf.sheetnames:
        res.add(18, WARN, "FX sensitivity Table 3", f"cannot locate the table ({note})")
        return
    top = int(m.group(1))
    wsf = wbf["Sensitivity Analysis"]
    sens_cell = wsf.cell(row=top + 7, column=3).value
    if not (isinstance(sens_cell, str) and sens_cell.startswith("=")):
        res.add(18, FAIL, "FX sensitivity Table 3",
                f"the per-1-JPY sensitivity at C{top + 7} is {sens_cell!r}, not a "
                f"formula — a hardcoded sensitivity stops tracking the scenario")
        return
    if not has_values:
        res.add(18, SKIP, "FX sensitivity Table 3", "needs recalc")
        return
    wsv = wbv["Sensitivity Analysis"]
    rate = _num(wsv.cell(row=top + 1, column=3).value)
    op_model = _num(wsv.cell(row=top + 6, column=3).value)
    sens = _num(wsv.cell(row=top + 7, column=3).value)
    # The column whose header equals the assumption rate must return the
    # unshifted operating income; anything else means the grid is off-centre.
    centre = None
    for col in range(3, 12):
        hdr = _num(wsv.cell(row=top + 9, column=col).value)
        if hdr is not None and rate is not None and abs(hdr - rate) < 1e-9:
            centre = _num(wsv.cell(row=top + 10, column=col).value)
            break
    if centre is None or op_model is None:
        res.add(18, WARN, "FX sensitivity Table 3",
                "no cached grid values to check the centre column against")
        return
    if not _close(centre, op_model):
        res.add(18, FAIL, "FX sensitivity Table 3",
                f"at the assumption rate the grid shows OP {centre:,.0f} but the "
                f"DCF Model's Year-1 OP is {op_model:,.0f}")
    else:
        res.add(18, PASS, "FX sensitivity Table 3",
                f"centre column reproduces Year-1 OP {op_model:,.0f}; "
                f"sensitivity {sens:,.0f} JPY mn per 1 JPY of rate")


def check_implied_exit_multiple(res, wbf, wbv, has_values):
    if not has_values:
        res.add(11, SKIP, "PGM-implied vs assumed exit multiple", "needs recalc")
        return
    wsv = wbv["DCF Model"]
    tv = _num(wsv.cell(row=R_TV_PGM, column=3).value)
    ebitda5 = _num(wsv.cell(row=R_YR5_EBITDA, column=3).value)
    assumed = _num(wsv["C14"].value)
    if not tv or not ebitda5 or not assumed:
        res.add(11, SKIP, "PGM-implied vs assumed exit multiple", "no cached values")
        return
    implied = tv / ebitda5
    gap = max(implied, assumed) / min(implied, assumed)
    if gap > 1.8:
        res.add(11, WARN, "PGM-implied vs assumed exit multiple",
                f"PGM implies {implied:.2f}x vs assumed {assumed:.2f}x "
                f"({gap:.2f}x apart) — the two DCF legs disagree on terminal value")
    else:
        res.add(11, PASS, "PGM-implied vs assumed exit multiple",
                f"PGM implies {implied:.2f}x vs assumed {assumed:.2f}x ({gap:.2f}x)")


def check_peer_ebitda_equals_ebit(res, wbf, wbv, meta, has_values):
    if "Comps Analysis" not in wbf.sheetnames:
        res.add(12, SKIP, "Peer EBITDA != EBIT", "no 'Comps Analysis' sheet")
        return
    ws = wbv["Comps Analysis"] if has_values else wbf["Comps Analysis"]
    subj = _subject_row(wbf, meta)
    hits = []
    for r in range(5, 40):
        name = ws.cell(row=r, column=2).value
        if not name or r == subj:
            if not name:
                break
            continue
        eb = _num(ws.cell(row=r, column=7).value)
        oi = _num(ws.cell(row=r, column=8).value)
        if eb is not None and oi is not None and eb == oi and eb > 0:
            hits.append(f"{name} (row {r})")
    if hits:
        res.add(12, WARN, "Peer EBITDA != EBIT",
                f"D&A not added back for: {', '.join(hits)}")
    else:
        res.add(12, PASS, "Peer EBITDA != EBIT",
                "no peer row where EBITDA equals operating income")


# =====================================================================
# market_analysis / sotp checks (14-20)
# =====================================================================
def _kind(wbf):
    """Classify the workbook by its sheets."""
    names = set(wbf.sheetnames)
    if {'Implied Growth Analysis', 'Market Scorecard'} & names:
        return 'market_analysis'
    if {'SOTP Valuation', 'D&A Allocation'} & names:
        return 'sotp'
    if 'DCF Model' in names:
        return 'dcf'
    return 'unknown'


def check_interp_iferror(res, wbf):
    """#14 Block 3's MATCH must be wrapped in IFERROR (price below the grid)."""
    ws = wbf['Implied Growth Analysis']
    hits, bad = 0, []
    for row in ws.iter_rows(min_col=4, max_col=4):
        v = row[0].value
        if isinstance(v, str) and 'MATCH(' in v.upper():
            hits += 1
            if 'IFERROR(' not in v.upper():
                bad.append(row[0].coordinate)
    if not hits:
        res.add(14, SKIP, "Block 3 interpolation guards #N/A", "no MATCH formula found")
    elif bad:
        res.add(14, FAIL, "Block 3 interpolation guards #N/A",
                f"{len(bad)} formula(s) without IFERROR: {', '.join(bad[:6])} — a "
                f"price below the alpha grid returns #N/A")
    else:
        res.add(14, PASS, "Block 3 interpolation guards #N/A",
                f"all {hits} interpolation formula(s) wrap MATCH in IFERROR")


def _peer_breakdown_rows(ws):
    """[(row, name)] for the peer rows under the 'N社内訳' header."""
    out = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=2).value
        if not (isinstance(v, str) and '内訳' in v):
            continue
        rr = r + 1
        while True:
            nm = ws.cell(row=rr, column=2).value
            if not isinstance(nm, str) or not nm.strip():
                break
            s = nm.strip()
            if s.startswith('（参考') or s.startswith('(参考'):
                break  # explicitly flagged as outside the sample
            out.append((rr, s))
            rr += 1
        break
    return out


def check_reverse_comps_excludes_self(res, wbf, wbv):
    """#15 the distribution inputs must not contain the subject's own multiple."""
    if 'Implied Multiple Analysis' not in wbf.sheetnames:
        res.add(15, SKIP, "逆算Comps excludes the subject", "no reverse-comps sheet")
        return
    ws = wbf['Implied Multiple Analysis']
    title = str(wbf['Implied Growth Analysis']['B2'].value or '')
    # "... - <company> (<ticker>)" -> company name
    subject = title.split(' - ')[-1].rsplit('(', 1)[0].strip() if ' - ' in title else ''
    # Only the peer-breakdown block counts — the sheet title also carries the
    # company name, and matching that would fail every workbook.
    listed = _peer_breakdown_rows(ws)
    if not subject:
        res.add(15, WARN, "逆算Comps excludes the subject",
                "could not determine the subject company name")
        return
    hit = [f"B{r}" for r, s in listed if subject and subject in s]
    if hit:
        res.add(15, FAIL, "逆算Comps excludes the subject",
                f"subject '{subject}' appears in the peer breakdown at "
                f"{', '.join(hit)} — its own multiple is inside the distribution")
    else:
        res.add(15, PASS, "逆算Comps excludes the subject",
                f"'{subject}' not present among the distribution inputs")


def check_alpha_scan_ascending(res, wbv, has_values):
    """#16 implied price row must be ascending across the alpha grid."""
    if not has_values:
        res.add(16, SKIP, "alpha-scan is ascending", "needs recalc")
        return
    ws = wbv['Implied Growth Analysis']
    vals = []
    for col in range(3, 21):
        v = _num(ws.cell(row=45, column=col).value)
        if v is None:
            break
        vals.append(v)
    if len(vals) < 3:
        res.add(16, SKIP, "alpha-scan is ascending", "no cached implied prices")
        return
    if all(b >= a for a, b in zip(vals, vals[1:])):
        res.add(16, PASS, "alpha-scan is ascending",
                f"{vals[0]:,.0f} -> {vals[-1]:,.0f} across {len(vals)} alphas")
    else:
        res.add(16, WARN, "alpha-scan is ascending",
                f"implied price is NOT ascending ({vals[0]:,.0f} -> {vals[-1]:,.0f}); "
                f"Base growth is negative, so Block 3's MATCH bracket is unreliable")


def check_sotp_da_check_ok(res, wbv, wbf, has_values):
    """#17 the D&A allocation Check cell must read OK (weights sum to 100%)."""
    ws = wbv['D&A Allocation'] if has_values else wbf['D&A Allocation']
    found = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=5).value
        if isinstance(v, str) and v.strip() in ('OK', 'CHECK'):
            found = (r, v.strip())
            break
    if found is None:
        res.add(17, SKIP if not has_values else FAIL,
                "D&A allocation sums to 100%",
                "Check cell not found / not recalculated")
        return
    r, v = found
    if v == 'OK':
        res.add(17, PASS, "D&A allocation sums to 100%", f"E{r} = OK")
    else:
        res.add(17, FAIL, "D&A allocation sums to 100%",
                f"E{r} = CHECK — the allocation percentages do not sum to 100%")


def check_sotp_cover_link(res, wbf):
    """#18 the Cover's SOTP row must be a live link, not a pasted number."""
    ws = wbf['Cover & Thesis']
    row = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=2).value
        # 'SOTP (Base Case)' in the cross-check table — not the sheet title
        # ('SOTP Valuation Model'), which also starts with "SOTP".
        if isinstance(v, str) and v.strip().upper().startswith('SOTP ('):
            row = r
            break
    if row is None:
        res.add(18, FAIL, "Cover SOTP value is a live link",
                "no 'SOTP (Base Case)' row on the Cover sheet")
        return
    v = ws.cell(row=row, column=3).value
    if isinstance(v, str) and "'SOTP Valuation'" in v:
        res.add(18, PASS, "Cover SOTP value is a live link", f"C{row} = {v}")
    else:
        res.add(18, FAIL, "Cover SOTP value is a live link",
                f"C{row} = {v!r} — expected a formula referencing 'SOTP Valuation'")


def check_sotp_fair_value_range(res, wbv, has_values):
    """#19 fair value per share in a plausible range (unit-error detector)."""
    if not has_values:
        res.add(19, SKIP, "SOTP fair value per share is plausible", "needs recalc")
        return
    ws = wbv['SOTP Valuation']
    for r in range(1, ws.max_row + 1):
        lbl = ws.cell(row=r, column=2).value
        if isinstance(lbl, str) and lbl.strip().startswith('Fair Value Per Share'):
            v = _num(ws.cell(row=r, column=3).value)
            if v is None:
                res.add(19, SKIP, "SOTP fair value per share is plausible",
                        "no cached value")
            elif 10 <= v <= 1_000_000:
                res.add(19, PASS, "SOTP fair value per share is plausible",
                        f"JPY {v:,.0f}")
            else:
                res.add(19, WARN, "SOTP fair value per share is plausible",
                        f"JPY {v:,.0f} is outside [10, 1,000,000] — check the "
                        f"shares unit (this template wants 千株, not 株)")
            return
    res.add(19, SKIP, "SOTP fair value per share is plausible", "row not found")


def check_hardcoded_company_count(res, wbf, wbv, has_values):
    """#20 a literal "N社" caption must match the number of peers actually shown."""
    ws = wbf['Implied Multiple Analysis'] if 'Implied Multiple Analysis' in wbf.sheetnames else None
    if ws is None:
        res.add(20, SKIP, "Peer-count caption matches the data", "no reverse-comps sheet")
        return
    counted = set()
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=2).value
        if not isinstance(v, str):
            continue
        s = v.strip()
        if s.startswith('（参考') or s.startswith('(参考'):
            continue
        m = re.search(r'(\d+)\s*社', s)
        if m:
            counted.add(int(m.group(1)))
    n_listed = len(_peer_breakdown_rows(ws))
    if not counted:
        res.add(20, SKIP, "Peer-count caption matches the data", "no 'N社' caption")
    elif n_listed and counted != {n_listed}:
        res.add(20, WARN, "Peer-count caption matches the data",
                f"caption(s) say {sorted(counted)}社 but {n_listed} peer row(s) "
                f"are listed")
    else:
        res.add(20, PASS, "Peer-count caption matches the data",
                f"{n_listed} peer row(s), caption says {sorted(counted)}社")


def check_adjustments_log(res, wbf):
    if "Adjustments Log" in wbf.sheetnames:
        res.add(13, PASS, "Adjustments Log sheet present", "")
    else:
        res.add(13, WARN, "Adjustments Log sheet present",
                "no ledger for manual edits — regenerate with the current template")


# =====================================================================
# driver
# =====================================================================
def validate_workbook(path, write_report=True, allow_skip=False):
    wbf = openpyxl.load_workbook(path, data_only=False)
    wbv = openpyxl.load_workbook(path, data_only=True)

    # "Has values" = Excel has cached results for at least one formula cell.
    # Probed generically (any sheet) so it works for every template, not just
    # the DCF's WACC cell.
    has_values = False
    probed = 0
    for ws in wbf.worksheets:
        wsv = wbv[ws.title]
        for row in ws.iter_rows():
            for c in row:
                if not (isinstance(c.value, str) and c.value.startswith('=')):
                    continue
                probed += 1
                if wsv[c.coordinate].value is not None:
                    has_values = True
                    break
            if has_values or probed > 200:
                break
        if has_values or probed > 200:
            break

    meta = read_metadata(wbf)
    res = Result(path)
    kind = _kind(wbf)
    print(f"  workbook kind: {kind}")

    if kind == 'dcf':
        # `has_values` probes the DCF WACC cell, which only exists on a DCF.
        check_formula_errors(res, wbv, has_values)
        check_scenario_index(res, wbf, meta)
        check_stats_exclude_subject(res, wbf, meta)
        check_subject_row_formulas(res, wbf, meta)
        check_ltm_revenue(res, wbf, wbv, meta, has_values)
        check_capex_da_ratios(res, wbf, meta)
        check_fs_year_alignment(res, wbf, meta)
        check_terminal_capex(res, wbf, wbv, has_values)
        check_pgm_negative_equity(res, wbf, wbv, has_values)
        check_implied_exit_multiple(res, wbf, wbv, has_values)
        check_peer_ebitda_equals_ebit(res, wbf, wbv, meta, has_values)
        check_adjustments_log(res, wbf)
        check_target_excludes_comps(res, wbf)
        check_cost_of_debt(res, wbf, wbv, meta, has_values)
        check_fx_sensitivity(res, wbf, wbv, meta, has_values)
        check_exit_negative_equity(res, wbf, wbv, has_values)
        check_reverse_dcf_sheet(res, wbf, meta)
        check_core_ebitda(res, wbf, meta)
        check_comps_reference_band(res, wbf, wbv, has_values)
        check_market_data(res, wbf, meta)
        check_arbitration_applied(res, path, wbf, meta)
        check_disclosure_vintage(res, path, meta)
        check_bank_model(res, path, wbf, wbv, has_values)
        check_beta_clamped(res, wbf, meta)
        check_type_f_model(res, path, wbf, wbv, has_values)
    elif kind == 'market_analysis':
        check_formula_errors(res, wbv, has_values)
        check_interp_iferror(res, wbf)
        check_reverse_comps_excludes_self(res, wbf, wbv)
        check_alpha_scan_ascending(res, wbv, has_values)
        check_hardcoded_company_count(res, wbf, wbv, has_values)
    elif kind == 'sotp':
        check_formula_errors(res, wbv, has_values)
        check_sotp_da_check_ok(res, wbv, wbf, has_values)
        check_sotp_cover_link(res, wbf)
        check_sotp_fair_value_range(res, wbv, has_values)
    else:
        res.add(0, WARN, "Workbook type recognised",
                f"sheets {wbf.sheetnames} match no known template")

    # SKIP means "this check could not run", which is not the same claim as
    # "this check passed". Counting it as neutral is how 9503 got VERDICT: PASS
    # with an empty Target Price: its recalc had been interrupted, five
    # value-level checks reported SKIP, and nothing turned that into a failure.
    n_skip = sum(1 for r in res.rows if r[1] == SKIP)
    if n_skip and not allow_skip:
        res.add(99, FAIL, "All checks executed",
                f"{n_skip} check(s) SKIPped - the workbook is NOT verified. "
                f"Recalculate it (python scripts/recalc_excel_com.py <xlsx>) and "
                f"re-validate, or pass --allow-skip if partial validation is "
                f"deliberate.")

    res.rows.sort(key=lambda r: r[0])
    text = res.render()
    print(text)
    if not has_values:
        print("\nNOTE: the workbook has no cached formula values. Run "
              "`python scripts/recalc_excel_com.py <xlsx>` and re-validate to "
              "exercise the value-level checks.")
    if write_report:
        report = os.path.splitext(path)[0] + "_validation.txt"
        with open(report, "w", encoding="utf-8") as f:
            f.write(text + "\n")
        print(f"\nReport written: {report}")
    return res


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    allow_skip = "--allow-skip" in sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(2)
    target = args[0]
    if not os.path.isfile(target):
        print(f"ERROR: not a file: {target}")
        sys.exit(2)
    res = validate_workbook(target, allow_skip=allow_skip)
    sys.exit(1 if res.failed else 0)


if __name__ == "__main__":
    main()
