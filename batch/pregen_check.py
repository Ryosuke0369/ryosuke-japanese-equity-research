"""追補7 — pre-generation check run on every pending ticker before it is modelled.

Three things have each caused a whole-ticker rework at least once, and all three
are decidable BEFORE generation, from the cache alone:

  §Y  OPM monotonicity. A monotonic series must use the trend-consistent rule
      (2267 precedent, Base = (latest + median)/2), NOT §4-4 median convergence.
      Applying mean-reversion to a trend is a statistical error, and 6506 安川 had
      to be regenerated twice because this was decided by eye instead of by rule.

  §F  Which fields the preventive completion must supply, and — more important —
      which fields are MISSING from yfinance so they cannot be supplied silently.
      A missing latest-FY D&A is what let 2503/4183 past the §G gate.

  §B  The financial-business prescreen result (read back from the cached scan).

It prints one block per ticker and a machine-readable summary line. It decides
nothing on its own: it tells the operator which rule applies and what is absent.
"""
import sys, os, json, glob

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def pct(sorted_vals, q):
    """Excel PERCENTILE.INC."""
    if len(sorted_vals) == 1:
        return sorted_vals[0]
    i = (len(sorted_vals) - 1) * q
    lo = int(i)
    hi = min(lo + 1, len(sorted_vals) - 1)
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (i - lo)


def monotonic(s):
    """'increasing' / 'decreasing' / None — strict, over the whole window."""
    if len(s) < 3:
        return None
    if all(b > a for a, b in zip(s, s[1:])):
        return "increasing"
    if all(b < a for a, b in zip(s, s[1:])):
        return "decreasing"
    return None


def load(code):
    p = os.path.join(ROOT, "batch", "cache", f"{code}.json")
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else None


def prescreen(code):
    p = os.path.join(ROOT, "batch", "prescreen_results.json")
    if not os.path.exists(p):
        return "未実施"
    d = json.load(open(p, encoding="utf-8")).get(code)
    if not isinstance(d, dict):
        return "未実施"
    hits = d.get("hits") or []
    return "CLEAN" if not hits else "FOUND(%d件) %s" % (
        len(hits), "; ".join(
            str(h.get("tag") or h.get("name") or h) if isinstance(h, dict) else str(h)
            for h in hits)[:160])


def check(code):
    d = load(code)
    if not d:
        print(f"{code}: キャッシュなし — batch/fetch_fundamentals.py を先に実行")
        return None
    inc, cf, bs = d.get("income", {}), d.get("cashflow", {}), d.get("balance", {})
    ys = sorted({k for v in inc.values() for k in v})
    rows = []
    for y in ys:
        r = inc.get("Total Revenue", {}).get(y)
        o = inc.get("Operating Income", {}).get(y)
        if r and o is not None:
            rows.append((y, r, o, o / r))
    if len(rows) < 3:
        print(f"{code}: OPM系列が3期未満 — 手動確認")
        return None
    s = [x[3] for x in rows]
    ss = sorted(s)
    latest, q25, med, q75 = s[-1], pct(ss, 0.25), pct(ss, 0.5), pct(ss, 0.75)
    mono = monotonic(s)
    ly = rows[-1][0]

    # §Y rule selection — 追補8 §AE-2 で方向別に確定
    if mono == "increasing":
        # 上昇トレンドに平均回帰を当てるのは誤り。かといって トレンド を外挿するのも危険なので、
        # 「直近水準の維持」を採る。上積みの外挿はしない = 保守側のガード。
        rule = "§Y トレンド整合ルール（単調増加 → 直近水準の維持）"
        base = latest
        why = "OPM系列が**単調増加** → 平均回帰(§4-4)は使わず、直近水準を維持（さらなる上積みは外挿しない）"
    elif mono == "decreasing":
        rule = "§Y トレンド整合ルール（単調減少 → 2267先例）"
        base = (latest + med) / 2
        why = "OPM系列が**単調減少** → 平均回帰(§4-4)は使わず (直近 + 中央値)/2"
    elif latest <= med:
        rule = "型A規則（中央値へ収束）"
        base = med
        why = "非単調 かつ 直近 ≤ 中央値"
    else:
        rule = "型A規則（(直近 + p75)/2 へ緩やかに正常化）"
        base = (latest + q75) / 2
        why = "非単調 かつ 直近 > 中央値"

    trough = latest <= q25            # §U signature (divergence checked post-generation)
    neg = [f"{y[:7]}" for y, r, o, m in rows if o <= 0]

    # §F — what yfinance is missing for the latest FY
    miss = []
    if cf.get("Depreciation And Amortization", {}).get(ly) is None:
        miss.append("**直近期D&A**（→ comps CSV が前期にフォールバック。EDINET一次ソース必須）")
    if cf.get("Capital Expenditure", {}).get(ly) is None:
        miss.append("**直近期capex**")
    if inc.get("Selling General And Administration", {}).get(ly) is None:
        miss.append("SGA（COGS%固定+逆算方式なので致命的ではない。hist_sga は渡さない）")
    if inc.get("Interest Expense", {}).get(ly) is None:
        miss.append("支払利息（実績Kd 算出不可 → テンプレ既定にフォールバック・要記録）")
    if bs.get("Total Debt", {}).get(ly) is None:
        miss.append("有利子負債（debt-free の可能性。net_debt を明示指定して確認）")
    # duplicated adjacent years = a yfinance artefact, not a real flat year
    dup = []
    for k in ("Depreciation And Amortization", "Capital Expenditure", "Operating Cash Flow"):
        v = cf.get(k, {})
        vals = [v.get(y) for y in ys if v.get(y) is not None]
        for a, b in zip(vals, vals[1:]):
            if a == b and a is not None:
                dup.append(k)
                break

    print(f"=== {code} {d.get('name','')}")
    print(f"  プレスクリーン(§B): {prescreen(code)}")
    print(f"  OPM: {' → '.join(f'{m:.2%}' for m in s)}")
    print(f"       min {min(s):.2%} / p25 {q25:.2%} / med {med:.2%} / p75 {q75:.2%} / max {max(s):.2%}")
    print(f"  単調性(§Y): {mono or '非単調'}")
    print(f"  → 適用ルール: {rule}  Base到達OPM = {base:.2%}   （{why}）")
    print(f"  §U トラフ兆候: {'**該当**（直近 ≤ p25。生成後に乖離>3.0倍なら §X-2）' if trough else '非該当'}")
    if neg:
        print(f"  赤字年: {', '.join(neg)} → **型B**")
    if miss:
        print("  §F 欠損:")
        for m in miss:
            print(f"     - {m}")
    if dup:
        print(f"  ⚠ 隣接年で値が重複: {', '.join(sorted(set(dup)))} → yfinance のアーティファクト疑い。一次ソース照合")
    return dict(code=code, mono=mono, rule=rule, base=base, trough=trough,
                neg=bool(neg), miss=miss, dup=sorted(set(dup)))


if __name__ == "__main__":
    codes = sys.argv[1:]
    if not codes:
        print("usage: python batch/pregen_check.py <ticker> ...")
        raise SystemExit(1)
    res = [r for r in (check(c) for c in codes) if r]
    print("\n" + "=" * 70)
    mo = [r["code"] for r in res if r["mono"]]
    tr = [r["code"] for r in res if r["trough"]]
    ng = [r["code"] for r in res if r["neg"]]
    ms = [r["code"] for r in res if r["miss"]]
    dp = [r["code"] for r in res if r["dup"]]
    print("§Y 単調（トレンド整合ルール）:", ", ".join(mo) if mo else "なし")
    print("§U トラフ兆候（生成後に要再判定）:", ", ".join(tr) if tr else "なし")
    print("赤字年あり（型B）:", ", ".join(ng) if ng else "なし")
    print("§F 欠損あり（明示指定/一次ソースが必要）:", ", ".join(ms) if ms else "なし")
    print("値の重複あり（要一次ソース照合）:", ", ".join(dp) if dp else "なし")
