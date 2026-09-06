"""フェーズ2 §4 — the three gates plus the two freshness assertions, over one date.

The batch's completion condition is not "the generator exited 0". It is:

    validate clean (FAIL 0 and SKIP 0)  +  core_ebitda gate  +  market_data gate
    +  the workbook is newer than its overrides/comps  (追補6 §Z)
    +  the workbook was actually recalculated          (追補10 §AO)

This runs all five against the regenerated models for a single basis date, so a
stale 2026-09-05 file cannot be graded by accident, and prints one line per
ticker plus a census. Nothing here re-generates anything; it only judges.

Usage:
    python batch/gates_phase2.py                 # date 20260906, the 85 done tickers
    python batch/gates_phase2.py --date 20260905
"""
import argparse
import io
import json
import os
import re
import sys
import contextlib
import warnings

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)

from stale_check import assert_fresh, assert_recalculated       # noqa: E402
import check_core_ebitda                                        # noqa: E402
import check_market_data                                        # noqa: E402

VERDICT_RE = re.compile(r"FAIL (\d+) / WARN (\d+) / SKIP (\d+) / PASS (\d+)")


def quiet(fn, *a, **k):
    """Run a gate, swallow its console output, return (ok, captured_text)."""
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            ok = fn(*a, **k)
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"
    return bool(ok), buf.getvalue().strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default="20260906")
    ap.add_argument("--state", default=os.path.join(HERE, "batch_state.json"))
    ap.add_argument("--status", default="done")
    ap.add_argument("--out", default=os.path.join(HERE, "phase2_gates.json"))
    a = ap.parse_args()

    state = json.load(open(a.state, encoding="utf-8"))
    codes = sorted(k for k, v in state.items()
                   if isinstance(v, dict) and v.get("status") == a.status)

    rows, clean = {}, 0
    print(f"{'code':<6} {'FAIL':>4} {'WARN':>4} {'SKIP':>4} {'PASS':>4}  "
          f"{'ebitda':<7} {'market':<7} {'fresh':<6} {'recalc':<7} verdict")
    print("-" * 88)
    for code in codes:
        p = os.path.join(ROOT, "models", f"{code}_DCF_Model_{a.date}.xlsx")
        rec = {"file": os.path.basename(p)}
        if not os.path.exists(p):
            rec.update(missing=True)
            rows[code] = rec
            print(f"{code:<6}  --- workbook missing ---")
            continue
        rep = p.replace(".xlsx", "_validation.txt")
        if os.path.exists(rep):
            m = VERDICT_RE.search(open(rep, encoding="utf-8", errors="replace").read())
            if m:
                rec.update(fail=int(m.group(1)), warn=int(m.group(2)),
                           skip=int(m.group(3)), passed=int(m.group(4)))
        ok_eb, txt_eb = quiet(check_core_ebitda.check, p)
        ok_mk, txt_mk = quiet(check_market_data.check, p)
        ok_fr, txt_fr = quiet(assert_fresh, p)
        ok_rc, txt_rc = quiet(assert_recalculated, p)
        rec.update(core_ebitda=ok_eb, market_data=ok_mk, fresh=ok_fr,
                   recalculated=ok_rc,
                   notes={k: v for k, v in (("core_ebitda", txt_eb),
                                            ("market_data", txt_mk),
                                            ("fresh", txt_fr),
                                            ("recalc", txt_rc)) if v})
        all_ok = (rec.get("fail") == 0 and rec.get("skip") == 0
                  and ok_eb and ok_mk and ok_fr and ok_rc)
        rec["clean"] = all_ok
        clean += bool(all_ok)
        rows[code] = rec
        print(f"{code:<6} {rec.get('fail','-'):>4} {rec.get('warn','-'):>4} "
              f"{rec.get('skip','-'):>4} {rec.get('passed','-'):>4}  "
              f"{'ok' if ok_eb else 'NG':<7} {'ok' if ok_mk else 'NG':<7} "
              f"{'ok' if ok_fr else 'STALE':<6} {'ok' if ok_rc else 'NG':<7} "
              f"{'CLEAN' if all_ok else '*** NOT CLEAN ***'}")

    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)

    print("-" * 88)
    n = len(codes)
    print(f"{clean}/{n} clean")
    for label, key in (("FAIL > 0", "fail"), ("SKIP > 0", "skip")):
        bad = [c for c, r in rows.items() if (r.get(key) or 0) > 0]
        print(f"  {label}: {len(bad)}" + (f" -> {', '.join(bad)}" if bad else ""))
    for label, key in (("core_ebitda gate", "core_ebitda"),
                       ("market_data gate", "market_data"),
                       ("freshness (追補6 §Z)", "fresh"),
                       ("recalculated (追補10 §AO)", "recalculated")):
        bad = [c for c, r in rows.items() if r.get(key) is False]
        print(f"  {label} failures: {len(bad)}"
              + (f" -> {', '.join(bad)}" if bad else ""))
    missing = [c for c, r in rows.items() if r.get("missing")]
    if missing:
        print(f"  missing workbooks: {', '.join(missing)}")
    warns = [(c, r.get("warn")) for c, r in rows.items() if (r.get("warn") or 0) > 0]
    print(f"  workbooks with WARN > 0: {len(warns)}"
          + (f" -> {', '.join(f'{c}({w})' for c, w in warns)}" if warns else ""))
    print(f"  detail: {a.out}")


if __name__ == "__main__":
    main()
