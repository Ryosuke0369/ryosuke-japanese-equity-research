"""Demote ONE of the two DCF legs to a reference row so Target = the other leg alone.

Two distinct triggers use this, and they demote opposite legs:

  追補4 §Q (leg=exit)  growth-premium names. The perpetuity leg implies a 5-7x exit
      multiple (g fixed at 1.0%) while the peer median sits above 25x. Demote Exit,
      keep PGM as a conservative floor.
      Trigger (all three): divergence > 3.0x AND PGM-implied < 10x AND assumed > 25x.

  追補5 §U (leg=pgm)   denominator-trough names, AFTER the 型B mid-cycle
      re-generation has been applied and the divergence still exceeds 3.0x.
      Here the PGM leg is the distorted one: for a capital-intensive company
      (capex/D&A persistently > 1.5x) the perpetuity leg implies a terminal
      EV/EBITDA far BELOW anything observable in the market, so it fails an
      empirical sanity test as a terminal value. Demote PGM, keep Exit.

In both cases the point is that 追補5 forbids the naive midpoint: a Target that is
the average of two irreconcilable worldviews belongs to neither.

追補12 §A-3 以降、書き換えの実体は `scripts/arbitration.apply_demotion()` にある
（generate_dcf.py が生成の最終段で同じ関数を呼ぶため）。本スクリプトはその CLI である。

Usage:
  python batch/demote_dcf_leg.py <xlsx> --leg exit|pgm --div D --pgm-implied P \
      --assumed A [--band p25-p75] [--rule auto|x] [--reason "extra sentence"]
"""
import argparse
import os
import sys
import warnings

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.arbitration import apply_demotion   # noqa: E402

if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("xlsx")
    p.add_argument("--leg", required=True, choices=["exit", "pgm"])
    p.add_argument("--div", required=True)
    p.add_argument("--pgm-implied", required=True)
    p.add_argument("--assumed", required=True)
    p.add_argument("--reason", default="")
    p.add_argument("--rule", choices=["auto", "x"], default="auto",
                   help="auto = the legacy §Q/§U tag; x = 追補6 §X (the unified arbitration)")
    p.add_argument("--band", default="", help="peer EV/EBITDA band 'p25-p75' for the §X sentence")
    a = p.parse_args()
    try:
        apply_demotion(a.xlsx, a.leg, a.div, a.pgm_implied, a.assumed,
                       band=a.band, rule=a.rule, reason=a.reason)
    except ValueError as e:
        raise SystemExit(f"ERROR: {e}")
