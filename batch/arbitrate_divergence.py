"""追補6 §X — single arbitration path for the PGM/Exit divergence.

追補6 replaced the old split logic. Previously §Q (追補4) and §U (追補5) each had
their own trigger, and a name whose two DCF legs disagreed badly could satisfy
NEITHER and silently keep the midpoint average — which 追補5 had already banned.
6981 村田 was the proof: divergence 4.55x, larger than either name that HAD been
arbitrated, yet its latest OPM sat 0.56pt above the window p25 so no rule fired.

追補6 §X therefore makes DIVERGENCE ITSELF the single trigger and demotes §Q/§U
to evidence about WHICH leg to drop:

    divergence > 3.0x  ->  the midpoint average is forbidden. Grade the legs:

    1. growth-premium regime (§Q: assumed > 25x AND pgm-implied < 10x)
       -> demote Exit, Target = PGM alone (conservative floor).
          The sector's observed multiples embed a premium a 1% perpetuity
          cannot express, so the floor is the honest number.

    2. trough signature (§U: latest-FY OPM <= window p25)
       -> re-generate with mid-cycle normalisation FIRST. If the divergence
          still exceeds 3.0x afterwards, fall through to 3.

    3. normal regime -> OBSERVED-MULTIPLE BAND TEST.
       Build the band from the peers' own EV/EBITDA 25th-75th percentiles and
       ask which of the two legs falls outside it.
         - exactly one leg outside -> demote that leg
         - both inside            -> keep the average, flag for human review
         - both outside           -> manual queue (the batch does not decide)

Why a band and not "anchor to the market": anchoring unconditionally to peer
multiples would turn the screener into a copy of the market and destroy its
reason to exist. The band test only fires when one leg is outside the range of
things that actually trade — it removes the impossible leg rather than adopting
the market's view.

追補12 §A-3 以降、判定ロジックそのものは scripts/arbitration.py にある
（generate_dcf.py が生成の最終段で同じ関数を呼ぶため）。本スクリプトは
「どの基準日のワークブックを見るか」を解決してバッチ全体を走査する CLI である。

Usage:
  python batch/arbitrate_divergence.py [<ticker> ...]     (default: all done)
  python batch/arbitrate_divergence.py --markdown         (report table)
"""
import sys, os, re, csv, json, glob, statistics

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from stale_check import assert_fresh, assert_recalculated, BASIS_DATES  # 追補6 §Z / 追補10 §AO
from scripts.arbitration import (            # 追補12 §A-3 — 判定の実体はこちら
    BAND_RATIO_MAX, LEVERAGE_MAX, LOW_WACC, TREATED,
    pct, peer_band, opm_series, midcycle_recorded, already_treated,
    legs_and_wacc, divergence, arbitrate as _arbitrate,
)


def _d(code):
    """Which basis date this ticker's workbook carries (追補11 §AR)."""
    # Newest basis date first: フェーズ2 regenerated every ticker at 2026-09-06
    # next to the 2026-09-05 originals, and re-scanning the superseded file
    # would grade a model that no longer exists as anyone's answer.
    for d in sorted(BASIS_DATES, reverse=True):
        if os.path.exists(os.path.join(ROOT, "models", f"{code}_DCF_Model_{d}.xlsx")):
            return d
    return sorted(BASIS_DATES)[-1]


def _paths(code):
    d = _d(code)
    return (os.path.join(ROOT, "models", f"{code}_DCF_Model_{d}.xlsx"),
            os.path.join(ROOT, "data", "overrides", f"{code}_overrides.json"),
            os.path.join(ROOT, "data", "comps", f"{code}_comps.csv"),
            os.path.join(ROOT, "batch", "cache", f"{code}.json"))


CHECK11 = re.compile(r"PGM implies ([\d.]+)x vs assumed ([\d.]+)x \(([\d.]+)x")


def check11(code):
    """(pgm_implied, assumed, divergence) from the validation report, or None.

    Kept only as the fallback for a workbook with no cached values (a model that
    was never recalculated, or one produced before フェーズ2). The live path
    computes the same three numbers from the workbook itself — see
    scripts.arbitration.divergence.
    """
    p = os.path.join(ROOT, "models", f"{code}_DCF_Model_{_d(code)}_validation.txt")
    if not os.path.exists(p):
        return None
    m = CHECK11.search(open(p, encoding="utf-8", errors="replace").read())
    if not m:
        return None
    return float(m.group(1)), float(m.group(2)), float(m.group(3))


def arbitrate(code):
    """The 追補6 §X ladder for one ticker, resolved against its basis date."""
    xlsx, ov, comps, cache = _paths(code)
    return _arbitrate(xlsx, overrides_path=ov, comps_csv=comps, cache_path=cache,
                      code=code, divergence_fallback=check11(code))


def fmt(r):
    b = "" if not r["band"] else f"[{r['band'][0]:.2f}–{r['band'][1]:.2f}] n={r['band'][2]}"
    t = "" if not r["treated"] else f"  ※適用済 {r['treated'][1]}"
    pd_ = "" if not r.get("price_div") else f" | 株価乖離 {r['price_div']:5.2f}x"
    lw = " | **低WACC**" if r.get("low_wacc") else ""
    return (f"{r['code']}  倍率乖離 {r['div']:5.2f}x{pd_} | PGM逆算 {r['pgm_implied']:6.2f}x | "
            f"仮定Exit {r['assumed']:6.2f}x | バンド {b:<20}{lw} | {r['regime']:<20} | {r['verdict']}{t}")


if __name__ == "__main__":
    md = "--markdown" in sys.argv
    codes = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not codes:
        # De-duplicate by ticker: フェーズ2 left both the 2026-09-05 original and
        # the 2026-09-06 regeneration on disk, and globbing files rather than
        # codes scanned 159 workbooks for 85 companies - each divergence counted
        # twice. _d() picks which basis date is the live one.
        codes = sorted({os.path.basename(p)[:4] for p in
                        glob.glob(os.path.join(ROOT, "models",
                                               "*_DCF_Model_2026090[56].xlsx"))})
    stale = [c for c in codes                                        # 追補6 §Z
             if not assert_fresh(os.path.join(ROOT, "models", f"{c}_DCF_Model_{_d(c)}.xlsx"), quiet=False)
             or not assert_recalculated(os.path.join(ROOT, "models", f"{c}_DCF_Model_{_d(c)}.xlsx"), quiet=False)]
    res = [r for r in (arbitrate(c) for c in codes) if r]
    over = [r for r in res if r["div"] > 3.0 or (r.get("price_div") or 1.0) > 3.0]
    if md:
        print("| 銘柄 | 乖離 | PGM逆算 | 仮定Exit | Peerバンド(p25–p75) | レジーム | 裁定 |")
        print("|---|---|---|---|---|---|---|")
        for r in sorted(over, key=lambda x: -x["div"]):
            b = "—" if not r["band"] else f"{r['band'][0]:.2f}–{r['band'][1]:.2f}x"
            print(f"| {r['code']} | **{r['div']:.2f}x** | {r['pgm_implied']:.2f}x | "
                  f"{r['assumed']:.2f}x | {b} | {r['regime']} | {r['verdict']} |")
    else:
        for r in sorted(res, key=lambda x: -x["div"]):
            print(fmt(r))
    print(f"\n走査 {len(res)}件 / 乖離 > 3.0倍 {len(over)}件")
    need = [r["code"] for r in over if r["demote"] and not r["treated"]]
    print("要降格処理（未適用）:", ", ".join(need) if need else "なし")
    q = [r["code"] for r in over if "キュー" in r["verdict"]]
    print("要手動レビューキュー:", ", ".join(q) if q else "なし")
    keep = [r["code"] for r in over if "平均維持" in r["verdict"]]
    print("平均維持+目視レビュー:", ", ".join(keep) if keep else "なし")
    unst = [f"{r['code']}(p75/p25={r['band_ratio']:.2f})" for r in over if r.get("band_unstable")]
    print("バンド不安定・要目視レビュー（追補7 §AA）:", ", ".join(unst) if unst else "なし")
    na = [f"{r['code']}(net_debt/EV={r['lev']:.2f})" for r in over if r.get("na")]
    print("Target=N/A 高レバレッジ判定不能（追補8 §AD-2）:", ", ".join(na) if na else "なし")
    lev = [f"{r['code']}(株価乖離{r['price_div']:.2f}x)" for r in over
           if r["regime"].startswith("レバレッジ増幅型") and not r.get("na")]
    print("レバレッジ増幅型・保守側採用（追補8 §AD-2）:", ", ".join(lev) if lev else "なし")
    lwv = [f"{r['code']}({r['wacc']:.2%})" for r in res if r.get("low_wacc")]
    print("低WACC要検証（追補8 §AG, WACC < 5%%）:", ", ".join(lwv) if lwv else "なし")
