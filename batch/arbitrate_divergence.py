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


def _d(code):
    """Which basis date this ticker's workbook carries (追補11 §AR)."""
    # Newest basis date first: フェーズ2 regenerated every ticker at 2026-09-06
    # next to the 2026-09-05 originals, and re-scanning the superseded file
    # would grade a model that no longer exists as anyone's answer.
    for d in sorted(BASIS_DATES, reverse=True):
        if os.path.exists(os.path.join(ROOT, "models", f"{code}_DCF_Model_{d}.xlsx")):
            return d
    return sorted(BASIS_DATES)[-1]

BAND_RATIO_MAX = 1.5   # 追補7 §AA — above this the peer band is too wide to anchor on
LEVERAGE_MAX = 0.6     # 追補8 §AD-2 — above this net_debt/EV the point estimate is meaningless
LOW_WACC = 0.05        # 追補8 §AG — below this WACC the beta clamp + leverage mix inflates Target

CHECK11 = re.compile(r"PGM implies ([\d.]+)x vs assumed ([\d.]+)x \(([\d.]+)x")
TREATED = {"exit": "追補4 §Q", "pgm": "追補5 §U", "pgm6": "追補6 §X"}


def check11(code):
    """(pgm_implied, assumed, divergence) from the validation report, or None."""
    p = os.path.join(ROOT, "models", f"{code}_DCF_Model_{_d(code)}_validation.txt")
    if not os.path.exists(p):
        return None
    m = CHECK11.search(open(p, encoding="utf-8", errors="replace").read())
    if not m:
        return None
    return float(m.group(1)), float(m.group(2)), float(m.group(3))


def peer_band(code):
    """(p25, p75, n, [values]) of PEER EV/EBITDA — the subject row is excluded.

    Same arithmetic as batch/exit_multiple.py so the band and the exit multiple
    can never disagree about what a peer multiple is.
    """
    p = os.path.join(ROOT, "data", "comps", f"{code}_comps.csv")
    if not os.path.exists(p):
        return None
    rows = list(csv.DictReader(open(p, encoding="utf-8")))
    if len(rows) < 2:
        return None

    def num(r, k):
        try:
            return float(r[k])
        except (TypeError, ValueError, KeyError):
            return None

    vals = []
    for r in rows[1:]:                       # row 0 is the subject
        e, m, n = num(r, "EBITDA"), num(r, "Market_Cap"), num(r, "Net_Debt")
        if e and m is not None and n is not None and e > 0:
            vals.append((m + n) / e)
    if len(vals) < 3:
        return None
    vals.sort()
    return (pct(vals, 0.25), pct(vals, 0.75), len(vals), vals)


def pct(sorted_vals, q):
    """Excel PERCENTILE.INC — linear interpolation between order statistics."""
    if len(sorted_vals) == 1:
        return sorted_vals[0]
    i = (len(sorted_vals) - 1) * q
    lo = int(i)
    hi = min(lo + 1, len(sorted_vals) - 1)
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (i - lo)


def opm_series(code):
    """Latest OPM and window p25, from overrides first then the yfinance cache."""
    ov = os.path.join(ROOT, "data", "overrides", f"{code}_overrides.json")
    if os.path.exists(ov):
        d = json.load(open(ov, encoding="utf-8"))
        rev, oi = d.get("hist_revenue"), d.get("hist_operating_income")
        if rev and oi and len(rev) == len(oi):
            s = [o / r for r, o in zip(rev, oi) if r]
            if len(s) >= 3:
                return s[-1], pct(sorted(s), 0.25), s
    c = os.path.join(ROOT, "batch", "cache", f"{code}.json")
    if os.path.exists(c):
        d = json.load(open(c, encoding="utf-8"))
        inc = d.get("income", {})
        ys = sorted({k for v in inc.values() for k in v})
        s = []
        for y in ys:
            r = inc.get("Total Revenue", {}).get(y)
            o = inc.get("Operating Income", {}).get(y)
            if r and o is not None:
                s.append(o / r)
        if len(s) >= 3:
            return s[-1], pct(sorted(s), 0.25), s
    return None



def _strictly_monotonic(series):
    """'increasing' / 'decreasing' / None - the same test batch/pregen_check.py uses."""
    if not series or len(series) < 3:
        return None
    if all(b > a for a, b in zip(series, series[1:])):
        return "increasing"
    if all(b < a for a, b in zip(series, series[1:])):
        return "decreasing"
    return None


def midcycle_recorded(code):
    """True when the overrides record that a mid-cycle normalisation was applied.

    追補5 §U says "re-generate with mid-cycle normalisation FIRST, then re-judge".
    `already_treated()` cannot answer that: it looks for a demoted leg, which is
    the OUTCOME of the re-judgement, not the input to it. The analyst records the
    normalisation in the overrides' `_scenario_note`, so that is what is read.
    """
    ov = os.path.join(ROOT, "data", "overrides", f"{code}_overrides.json")
    if not os.path.exists(ov):
        return False
    try:
        d = json.load(open(ov, encoding="utf-8"))
    except (OSError, ValueError):
        return False
    txt = " ".join(str(v) for k, v in d.items() if k.startswith("_"))
    return ("ミッドサイクル正常化" in txt) or ("§U" in txt)


def legs_and_wacc(code):
    """(pgm_price, exit_price, price, shares, wacc, net_debt) from the workbook."""
    p = os.path.join(ROOT, "models", f"{code}_DCF_Model_{_d(code)}.xlsx")
    if not os.path.exists(p):
        return None
    try:
        import openpyxl
        wb = openpyxl.load_workbook(p, data_only=True)
        e, d = wb["Executive Summary"], wb["DCF Model"]
    except Exception:
        return None
    pgm, ex, price = e.cell(16, 3).value, e.cell(17, 3).value, e.cell(9, 3).value
    wacc = None
    for r in range(1, 40):
        lab = d.cell(r, 2).value
        if isinstance(lab, str) and lab.startswith("WACC") and isinstance(d.cell(r, 3).value, float):
            wacc = d.cell(r, 3).value
            break
    nd = None
    ov = os.path.join(ROOT, "data", "overrides", f"{code}_overrides.json")
    if os.path.exists(ov):
        nd = json.load(open(ov, encoding="utf-8")).get("net_debt")
    sh = None
    if os.path.exists(ov):
        sh = json.load(open(ov, encoding="utf-8")).get("shares_outstanding")
    return pgm, ex, price, sh, wacc, nd


def already_treated(code):
    """Which rule (if any) has already been stamped into the Executive Summary."""
    p = os.path.join(ROOT, "models", f"{code}_DCF_Model_{_d(code)}.xlsx")
    if not os.path.exists(p):
        return None
    try:
        import openpyxl
        ws = openpyxl.load_workbook(p)["Executive Summary"]
    except Exception:
        return None
    for r in range(1, 40):
        v = ws.cell(r, 2).value
        if not isinstance(v, str):
            continue
        for leg, tag in TREATED.items():
            if tag in v and "参考" in v:
                return ("exit" if "Exit" in v else "pgm", tag)
    return None


def arbitrate(code):
    c = check11(code)
    if not c:
        return None
    pgm_i, assumed, div = c
    out = dict(code=code, pgm_implied=pgm_i, assumed=assumed, div=div,
               regime="-", verdict="-", demote=None, band=None, opm=None,
               band_ratio=None, band_unstable=False, price_div=None, pgm_price=None,
               exit_price=None, lev=None, lev_mkt=None, wacc=None, low_wacc=False, na=False,
               treated=already_treated(code))

    lw = legs_and_wacc(code)
    if lw:
        pgm_p, ex_p, price, sh, wacc, nd = lw
        out["wacc"] = wacc
        # 追補8 §AG — a very low WACC is its own review flag, independent of divergence
        if wacc is not None and wacc < LOW_WACC:
            out["low_wacc"] = True
        if all(isinstance(x, (int, float)) for x in (pgm_p, ex_p)) and min(pgm_p, ex_p) > 0:
            out["price_div"] = max(pgm_p, ex_p) / min(pgm_p, ex_p)
            out["pgm_price"], out["exit_price"] = pgm_p, ex_p
            if nd is not None and sh:
                # 追補8 §AD-2 の分母は「本モデルの株式価値」であって時価総額ではない。
                # 論点は『EV のわずかな差が株式価値で数倍に増幅される』ことなので、
                # 増幅の分母になるのはモデルが出した株式価値のほう。
                eq_model = (pgm_p + ex_p) / 2 * sh / 1e6
                ev_model = eq_model + nd
                out["lev"] = (nd / ev_model) if ev_model else None
                out["lev_mkt"] = (nd / (price * sh / 1e6 + nd)) if price else None

    pdiv = out.get("price_div") or 1.0
    # 追補8 §AD — the trigger is an OR: multiples OR prices
    if div <= 3.0 and pdiv <= 3.0:
        out["regime"] = "収斂"
        out["verdict"] = "裁定不要（倍率乖離・株価乖離とも ≤ 3.0倍）"
        return out

    if div <= 3.0 and pdiv > 3.0:
        # 追補8 §AD-2 — leverage-amplified. Neither leg is wrong; the equity value is
        # simply a small difference between two large numbers, so a point estimate is
        # not meaningful. The band test cannot separate the legs (the multiples agree).
        out["regime"] = "レバレッジ増幅型(§AD-2)"
        lev = out.get("lev")
        if lev is not None and lev > LEVERAGE_MAX:
            out["demote"] = None
            out["na"] = True
            out["verdict"] = ("**Target = N/A（高レバレッジ・判定不能）** + 要目視レビュー"
                              "（net_debt/EV = %.2f > %.1f）" % (lev, LEVERAGE_MAX))
        else:
            out["verdict"] = ("保守側の脚を採用 + EV差異原因をサマリーに分析記録"
                              "（net_debt/EV = %s ≤ %.1f）"
                              % ("%.2f" % lev if lev is not None else "算出不可", LEVERAGE_MAX))
        return out

    # 1. growth-premium regime
    if assumed > 25.0 and pgm_i < 10.0:
        out["regime"] = "成長プレミアム(§Q)"
        out["demote"] = "exit"
        out["verdict"] = "Exit降格 → Target = PGM 単独"
        return out

    # 2. trough signature
    o = opm_series(code)
    out["opm"] = o
    if o and o[0] <= o[1]:
        # §U asks "is the denominator sitting in a cyclical trough". Its signature
        # (latest <= p25) is ALSO true of every strictly monotonic decline, where
        # the latest value is the minimum by construction - and 追補6 §Y says in so
        # many words that a trend must not be treated with mean reversion. Left
        # unqualified, §U therefore captured 6273 / 6367 / 6506 (OPM 31.31 -> 25.26
        # -> 24.02 -> 22.62, 9.47 -> 8.92 -> 8.45 -> 8.27, 12.29 -> 11.50 -> 9.33
        # -> 8.73) and parked them at "re-generate with mid-cycle normalisation",
        # which is the one treatment §Y forbids for them. A trend is not a trough:
        # those names go to the band test like any other normal-regime name.
        mono = _strictly_monotonic(o[2] if len(o) > 2 else None)
        if mono:
            out["regime"] = "単調(%s)・§U非該当" % ("増加" if mono == "increasing" else "減少")
        else:
            out["regime"] = "トラフ(§U)"
            # falls through to the band test only once mid-cycle regeneration is done
            if not (out["treated"] or midcycle_recorded(code)):
                out["verdict"] = "ミッドサイクル正常化で再生成 → 再判定"
                return out
            if not out["treated"]:
                out["regime"] = "トラフ(§U)・正常化済"

    # 3. normal regime — observed-multiple band test
    b = peer_band(code)
    if not b:
        out["regime"] = out["regime"] if out["regime"] != "-" else "通常"
        out["verdict"] = "Peer倍率バンドを作れず → 要手動レビュー"
        return out
    p25, p75, n, _ = b
    out["band"] = (p25, p75, n)
    out["band_ratio"] = (p75 / p25) if p25 else None
    pgm_out = pgm_i < p25 or pgm_i > p75
    exit_out = assumed < p25 or assumed > p75
    if out["regime"] == "-":
        out["regime"] = "通常"
    if pgm_out and not exit_out:
        out["demote"] = "pgm"
        out["verdict"] = "PGM降格 → Target = Exit 単独"
    elif exit_out and not pgm_out:
        out["demote"] = "exit"
        out["verdict"] = "Exit降格 → Target = PGM 単独"
    elif not pgm_out and not exit_out:
        out["verdict"] = "両脚バンド内 → 平均維持 + 要目視レビュー"
    else:
        out["verdict"] = "両脚バンド外 → 要手動レビューキュー"

    # 追補7 §AA — the band itself can be a weak anchor. Flag ONLY the names that
    # were resolved BY the band test (a wide band is only a problem when we leant
    # on it), so the §Q growth-premium names are deliberately out of scope.
    if out["demote"] and out["band_ratio"] and out["band_ratio"] > BAND_RATIO_MAX:
        out["band_unstable"] = True
        out["verdict"] += "（**バンド不安定・要目視レビュー**: p75/p25 = %.2f > %.1f）" % (
            out["band_ratio"], BAND_RATIO_MAX)
    return out


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
