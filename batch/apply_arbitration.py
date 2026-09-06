"""フェーズ2 §5 — re-apply the 追補6 §X leg demotions to the regenerated workbooks.

Why this is a separate step, and why it has to run again
--------------------------------------------------------
The §X arbitration is post-processing: `batch/demote_dcf_leg.py` edits a finished
workbook. `generate_dcf.py` knows nothing about it. So a regeneration necessarily
reverts every demotion the 2026-09-05 batch had applied, and the fresh workbook
goes back to averaging two legs that 追補5 forbids averaging.

That is visible in the numbers and is not a subtle effect: 6857 アドバンテスト's
Target read 4,288 (PGM alone, Exit demoted) before and 8,021 (the midpoint) after
regeneration; 6146 ディスコ 16,852 -> 37,191. Neither move is a valuation change —
it is the arbitration being absent.

This script asks batch/arbitrate_divergence.py for the current verdict on the
regenerated models (the verdict can differ from 2026-09-05: the new WACC changes
which leg is the implausible one, and 7012 川崎重工's PGM leg now fails outright
with EV < net debt) and applies exactly what it says, ticker by ticker.

Usage:
    python batch/apply_arbitration.py            # dry run, print the commands
    python batch/apply_arbitration.py --write
"""
import argparse
import os
import subprocess
import sys
import warnings

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)

from arbitrate_divergence import arbitrate, _d   # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true")
    ap.add_argument("--date", default="20260906")
    ap.add_argument("codes", nargs="*")
    a = ap.parse_args()

    import glob
    codes = a.codes or sorted({os.path.basename(p)[:4] for p in
                               glob.glob(os.path.join(ROOT, "models",
                                                      f"*_DCF_Model_{a.date}.xlsx"))})
    applied, skipped = [], []
    for code in codes:
        r = arbitrate(code)
        if not r or not r.get("demote"):
            continue
        xlsx = os.path.join(ROOT, "models", f"{code}_DCF_Model_{a.date}.xlsx")
        if not os.path.exists(xlsx):
            skipped.append((code, "no workbook for this date"))
            continue
        cmd = [sys.executable, os.path.join(HERE, "demote_dcf_leg.py"), xlsx,
               "--leg", r["demote"], "--rule", "x",
               "--div", f"{r['div']:.2f}",
               "--pgm-implied", f"{r['pgm_implied']:.2f}",
               "--assumed", f"{r['assumed']:.2f}"]
        band = r.get("band")
        if isinstance(band, (list, tuple)) and len(band) == 2:
            band = f"{band[0]:.2f}-{band[1]:.2f}"
        if band:
            cmd += ["--band", str(band)]
        print(f"{code}: demote {r['demote']}  div {r['div']:.2f}x  "
              f"regime {r.get('regime')}  -> {r.get('verdict')}")
        if not a.write:
            continue
        p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                           errors="replace")
        if p.returncode != 0:
            skipped.append((code, (p.stderr or p.stdout).strip()[-300:]))
            print(f"   FAILED: {(p.stderr or p.stdout).strip()[-200:]}")
        else:
            applied.append(code)
            print(f"   applied")

    print(f"\n{'applied' if a.write else 'would apply'}: {len(applied) if a.write else '(dry run)'}")
    if applied:
        print("  " + ", ".join(applied))
        print("\n  NOTE: the workbooks were edited by openpyxl, which does not "
              "compute formulas. Re-run scripts/recalc_excel_com.py on each and "
              "re-validate, or the Target cell has no cached value.")
    if skipped:
        print("  skipped/failed:")
        for c, why in skipped:
            print(f"    {c}: {why}")


if __name__ == "__main__":
    main()
