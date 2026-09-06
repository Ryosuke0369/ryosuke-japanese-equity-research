"""Detect the 追補5 §U signature: a PGM/Exit gap caused by a TROUGH DENOMINATOR.

§Q catches the numerator disease (the market pays a growth premium, so the peer
median sits above 25x while the perpetuity leg implies 5-7x). §U catches the
opposite: the Base scenario is anchored on a trough year, so PGM capitalises
depressed profit and the gap opens from below.

Trigger (all three):
  1. latest-FY OPM  <=  25th percentile of the observation window   (cycle trough)
  2. WARN 11 divergence  > 3.0x
  3. NOT §Q            (assumed exit multiple <= 25x)

The fix is NOT to demote a leg — that would leave Target = a capitalised trough,
which is worse than the midpoint. The fix is to re-generate with §4-4 mid-cycle
normalisation so the Base OPM path converges on the window median.

Usage: python batch/check_trough_artifact.py [<ticker> ...]   (default: all done)
"""
import sys, os, re, json, glob
sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

TREATED_TAG = "追補5 §U"  # stamped into the Exec Summary by batch/demote_dcf_leg.py


def already_treated(model_path):
    """True once §U has been applied to this workbook and recorded inside it.

    The §U trigger is a statement about HISTORY -- the latest FY sits at the
    bottom of the observation window -- so it keeps firing after the rework:
    the trough does not stop having happened. Without this check a later
    session would see the same hit and redo the rework in a loop. The workbook
    is the source of truth, because demote_dcf_leg.py stamps the rule name into
    the Executive Summary when the judgement is applied.
    """
    try:
        import openpyxl
        ws = openpyxl.load_workbook(model_path)["Executive Summary"]
    except Exception:
        return False
    for r in range(1, 40):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and TREATED_TAG in v:
            return True
    return False
PAT = re.compile(r"PGM implies ([\d.]+)x vs assumed ([\d.]+)x \(([\d.]+)x")


def pct(sorted_vals, p):
    """Excel PERCENTILE.INC on an already-sorted list."""
    if len(sorted_vals) == 1:
        return sorted_vals[0]
    k = (len(sorted_vals) - 1) * p
    f = int(k)
    c = min(f + 1, len(sorted_vals) - 1)
    return sorted_vals[f] + (k - f) * (sorted_vals[c] - sorted_vals[f])


def opm_series(code):
    """OPM history: overrides first, else the yfinance cache.

    The tickers generated before 追補2 §F-2 made hist_* mandatory do not carry the
    arrays in their overrides, so the scan would silently skip exactly the oldest
    models. Falling back to the cache keeps the sweep complete.
    """
    p = os.path.join(ROOT, "data", "overrides", "%s_overrides.json" % code)
    if os.path.isfile(p):
        d = json.load(open(p, encoding="utf-8"))
        rev, oi = d.get("hist_revenue"), d.get("hist_operating_income")
        if rev and oi:
            s = [o / r for o, r in zip(oi, rev)
                 if isinstance(o, (int, float)) and isinstance(r, (int, float)) and r]
            if s:
                return s, "overrides"
    c = os.path.join(ROOT, "batch", "cache", "%s.json" % code)
    if not os.path.isfile(c):
        return None
    d = json.load(open(c, encoding="utf-8"))
    inc = d.get("income", {})
    years = sorted({k for v in inc.values() for k in v})
    s = []
    for y in years:
        rev = inc.get("Total Revenue", {}).get(y) or inc.get("Operating Revenue", {}).get(y)
        oi = inc.get("Operating Income", {}).get(y) or inc.get("Total Operating Income As Reported", {}).get(y)
        if rev and oi is not None:
            s.append(oi / rev)
    return (s[-4:], "cache") if s else None


def check(code):
    vp = os.path.join(ROOT, "models", "%s_DCF_Model_20260905_validation.txt" % code)
    if not os.path.isfile(vp):
        return None
    m = PAT.search(open(vp, encoding="utf-8", errors="replace").read())
    if not m:
        return None
    pgm_i, assumed, div = (float(x) for x in m.groups())
    r = opm_series(code)
    if not r:
        return code, None, (pgm_i, assumed, div)
    s, src = r
    latest = s[-1]
    q25 = pct(sorted(s), 0.25)
    is_q = div > 3.0 and pgm_i < 10.0 and assumed > 25.0
    hit = (latest <= q25) and div > 3.0 and not (assumed > 25.0)
    return code, (latest, q25, hit, is_q, src), (pgm_i, assumed, div)


if __name__ == "__main__":
    codes = sys.argv[1:]
    if not codes:
        codes = sorted(os.path.basename(p)[:4] for p in
                       glob.glob(os.path.join(ROOT, "models", "*_DCF_Model_20260905.xlsx")))
    hits, treated = [], []
    for c in codes:
        r = check(c)
        if not r:
            continue
        code, o, (pgm_i, assumed, div) = r
        if o is None:
            print(f"        {code}: *** OPM系列を取得できず — 手動確認が必要 ***")
            continue
        latest, q25, hit, is_q, src = o
        done = hit and already_treated(os.path.join(
            ROOT, "models", f"{code}_DCF_Model_20260905.xlsx"))
        if hit and done:
            hit = False
            treated.append(code)
        tag = ("*** §U 該当 ***" if hit else
               "  (§U適用済)   " if done else
               "    (§Q該当)   " if is_q else "               ")
        print(f"{tag} {code}: 直近OPM {latest:6.2%} vs p25 {q25:6.2%} "
              f"{'≤' if latest <= q25 else '>'}  | 乖離 {div:5.2f}x | 仮定Exit {assumed:6.2f}x"
              f"{'  [cache]' if src == 'cache' else ''}")
        if hit:
            hits.append(code)
    print("\n§U 該当（分母トラフ型・要再生成）: %s" % (", ".join(hits) if hits else "なし"))
