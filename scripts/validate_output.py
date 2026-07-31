"""
validate_output.py - Machine-check a generated (or hand-edited) DCF workbook.

Usage:
    python scripts/validate_output.py models/3687_DCF_Model_20260731.xlsx

Exit code 1 when any check FAILs, 0 otherwise. Results also go to
<xlsx>_validation.txt next to the workbook.

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
    SCENARIO_NAMES, R_DA, R_CAPEX, R_TV_PGM, R_EV_PGM, R_YR5_EBITDA,
)

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


def check_capex_da_ratios(res, wbf, meta):
    ws = wbf["DCF Model"]
    for n, cell, key, basis_key, label in (
        (7, "C5", "capex_pct_c5", "capex_pct_basis", "Capex/Revenue"),
        (7, "C18", "da_pct_c18", "da_pct_basis", "D&A/Revenue"),
    ):
        try:
            expected = float(meta.get(key))
        except (TypeError, ValueError):
            res.add(n, SKIP, f"{cell} {label} basis", "no metadata")
            continue
        actual = _num(ws[cell].value)
        basis = str(meta.get(basis_key, "?"))
        if actual is None:
            res.add(n, FAIL, f"{cell} {label} basis", f"{cell} is empty or text")
        elif not _close(actual, expected):
            res.add(n, FAIL, f"{cell} {label} basis",
                    f"sheet {actual:.4%} vs generator {expected:.4%}")
        elif basis == "hist_3yr_avg":
            res.add(n, PASS, f"{cell} {label} basis",
                    f"{actual:.2%} = mean of last 3 historical years")
        else:
            res.add(n, PASS, f"{cell} {label} basis",
                    f"{actual:.2%} (explicit assumption, not a back-solve; "
                    f"basis={basis})")


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


def check_adjustments_log(res, wbf):
    if "Adjustments Log" in wbf.sheetnames:
        res.add(13, PASS, "Adjustments Log sheet present", "")
    else:
        res.add(13, WARN, "Adjustments Log sheet present",
                "no ledger for manual edits — regenerate with the current template")


# =====================================================================
# driver
# =====================================================================
def validate_workbook(path, write_report=True):
    wbf = openpyxl.load_workbook(path, data_only=False)
    wbv = openpyxl.load_workbook(path, data_only=True)

    # "Has values" = Excel has cached results for at least one formula cell.
    has_values = False
    if "DCF Model" in wbv.sheetnames:
        probe = wbv["DCF Model"]["C26"].value
        has_values = isinstance(probe, (int, float))

    meta = read_metadata(wbf)
    res = Result(path)

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
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    target = sys.argv[1]
    if not os.path.isfile(target):
        print(f"ERROR: not a file: {target}")
        sys.exit(2)
    res = validate_workbook(target)
    sys.exit(1 if res.failed else 0)


if __name__ == "__main__":
    main()
