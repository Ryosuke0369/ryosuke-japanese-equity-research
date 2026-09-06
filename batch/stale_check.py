"""追補6 §Z — refuse to grade a workbook that is older than the inputs it was built from.

Two separate incidents motivated this, and both produced a green report on a
file that did not reflect the current inputs:

  1. `generate_dcf.py` without `--force` prints a warning, skips the workbook,
     and still exits 0. The validator then re-reads the OLD file and prints
     VERDICT: PASS. A regeneration can therefore be "confirmed" without a
     single byte having changed.

  2. Rebuilding a comps CSV (adding one peer is enough) recomputes the subject
     row. Any workbook built before that rebuild is now inconsistent with the
     CSV it is graded against — this is how 2503 and 4183 slipped past the
     追補2 §G gate and were only caught by a full re-scan.

So every gate asks the same question first: is the workbook newer than the
overrides and the comps CSV that define it? A gate that grades a stale file is
worse than no gate, because it converts "unknown" into "verified".

Tolerance is 1 second — filesystem timestamp granularity, not a grace period.
"""
import os, re

TOL = 1.0

# 追補11 §AR — part4 は分析基準日が 2026-09-06 で、元の90件（2026-09-05）と併存する。
# ファイル名の日付は「分析基準日」であって実行日ではない（追補7 §AB）。
BASIS_DATES = ("20260905", "20260906")   # resolved newest-first (see model_path)
MODEL_GLOB = "*_DCF_Model_2026090[56].xlsx"


def model_path(code, root=None):
    """Resolve a ticker's workbook across the batch's basis dates."""
    root = root or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    # Newest basis date first: フェーズ2 regenerated every ticker at 2026-09-06
    # alongside the 2026-09-05 originals, and a gate that resolves to the older
    # file would grade the model the rerun replaced.
    for d in sorted(BASIS_DATES, reverse=True):
        p = os.path.join(root, "models", f"{code}_DCF_Model_{d}.xlsx")
        if os.path.exists(p):
            return p
    return os.path.join(root, "models", f"{code}_DCF_Model_{BASIS_DATES[0]}.xlsx")


def inputs_for(model_path):
    """The files a workbook is built from, as (label, path) pairs that exist."""
    root = os.path.dirname(os.path.dirname(os.path.abspath(model_path)))
    code = os.path.basename(model_path)[:4]
    cand = [("overrides", os.path.join(root, "data", "overrides", f"{code}_overrides.json")),
            ("comps", os.path.join(root, "data", "comps", f"{code}_comps.csv"))]
    return [(lab, p) for lab, p in cand if os.path.exists(p)]


def check(model_path):
    """(ok, [messages]) — False when any input is newer than the workbook."""
    if not os.path.exists(model_path):
        return False, ["モデルファイルが存在しない"]
    mt = os.path.getmtime(model_path)
    bad = []
    for lab, p in inputs_for(model_path):
        it = os.path.getmtime(p)
        if it > mt + TOL:
            bad.append(f"{lab} が {int(it - mt)}秒 新しい ({os.path.basename(p)})")
    return (not bad), bad


def assert_recalculated(model_path, quiet=False):
    """False when the workbook was never recalculated (validate reports SKIP > 0).

    Found the hard way: a generation killed by a timeout mid-recalc leaves a
    workbook with formulas but no cached values. validate then reports
    "FAIL 0 ... VERDICT: PASS" -- but with SKIP 5, because every formula-value
    check had to be skipped. Target / Recommendation / PGM / Exit all read None.
    A green verdict on an unusable file is the same trap as the --force silent
    skip, so the gates check SKIP too, not just FAIL.
    """
    rep = model_path.replace(".xlsx", "_validation.txt")
    if not os.path.exists(rep):
        return True                      # nothing to judge; validate itself will complain
    txt = open(rep, encoding="utf-8", errors="replace").read()
    m = re.search(r"FAIL \d+ / WARN \d+ / SKIP (\d+)", txt)
    if not m or int(m.group(1)) == 0:
        return True
    if not quiet:
        code = os.path.basename(model_path)[:4]
        print(f"*** NOT RECALCULATED *** {code}: validate が SKIP {m.group(1)} 件 "
              f"— 数式のキャッシュ値がなく Target 等が空の可能性。VERDICT が PASS でも使用不可")
        print(f"              → `sh batch/regen.sh {code}` で再生成すること（追補10 §AO で摘発された破損クラス）")
    return False


def assert_fresh(model_path, quiet=False):
    """Print a STALE line and return False when the workbook predates its inputs."""
    ok, bad = check(model_path)
    if not ok and not quiet:
        code = os.path.basename(model_path)[:4]
        print(f"*** STALE *** {code}: 入力の方が新しい — {' / '.join(bad)}")
        print(f"              → `python scripts/generate_dcf.py {code} --force` で再生成してから再判定すること（追補6 §Z）")
    return ok
