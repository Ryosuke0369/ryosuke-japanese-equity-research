"""screener/report/seasonality_profile.py — 個別銘柄の進捗率の季節性と S5c(R) の内訳。

S5c（`screener/signals/progress_history.py`）を指定銘柄に当てたとき、
**R が何から出来ているか（または何故計算できないか）を年ごとに開いて見せる**ためのレポート。
定義の正本は docs/backtest_acceptance_criteria.md「シャドウI-v2」。**ここは定義を変えない。**

出すもの（銘柄ごと）:
  1. 登録済み定義での R（`progress_history.evaluate` をそのまま呼ぶ）
  2. 年度 × 四半期の進捗率（分母 = その年度の着地）。**除外条件に落ちた年も値を出して理由を付ける**
  3. 四半期ごとのばらつき（中央値・標準偏差・四分位・最小/最大。除外前の生値で）
  4. 当期の同四半期を過去と並べる（対 会社予想 と 対 着地 の両方）
  5. 四半期単独の売上・営業利益（`fins_summary` の累計差分）

S5c は 0点・表示のみ。スコアには接続しない。

    python -m screener.report.seasonality_profile --codes 6741 6742 6743 --as-of 2026-09-21 --out <md>
"""
from __future__ import annotations

import argparse
import os
import statistics
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.signals import progress_history as P

QS = ("1Q", "2Q", "3Q")
QN = {"1Q": 1, "2Q": 2, "3Q": 3, "FY": 4}


def quartiles(vals):
    """(Q1, 中央値, Q3)。statistics.quantiles の inclusive 法。2本未満は None。"""
    if len(vals) < 2:
        return None
    q = statistics.quantiles(vals, n=4, method="inclusive")
    return q[0], statistics.median(vals), q[2]


def spread(vals):
    """ばらつきの要約。値そのものを返し、判定はしない。"""
    if not vals:
        return {"n": 0}
    out = {"n": len(vals), "median": statistics.median(vals),
           "min": min(vals), "max": max(vals)}
    out["sd"] = statistics.stdev(vals) if len(vals) >= 2 else None
    out["iqr"] = None
    qs = quartiles(vals)
    if qs:
        out["q1"], _, out["q3"] = qs
        out["iqr"] = qs[2] - qs[0]
    return out


def fiscal_table(rows):
    """fy_end -> {per_type: 行}。同じ年度・四半期に複数あれば後の開示が勝つ（progress_history と同じ）。"""
    tab = {}
    for r in rows:
        if r["per_type"] in QN and r.get("fy_end") and r.get("op") is not None:
            tab.setdefault(r["fy_end"], {})[r["per_type"]] = r
    return tab


def raw_progress(tab, item):
    """(fy_end, q) -> (進捗率, 除外理由 or None)。除外条件は progress_history と同じ。"""
    out = {}
    for fy, per in tab.items():
        land = per.get("FY", {}).get(item)
        for q in QS:
            cum = per.get(q, {}).get(item)
            if cum is None:
                continue
            if not land or land <= 0:
                out[(fy, q)] = (None, "着地が無い/0以下")
                continue
            p = cum / land
            why = None
            if p < P.PROGRESS_MIN:
                why = "5%%未満（%s）" % ("累計赤字" if cum < 0 else "過小")
            elif p > P.PROGRESS_MAX:
                why = "200%超"
            out[(fy, q)] = (p, why)
    return out


def standalone(tab, item):
    """(fy_end, q) -> 四半期単独値。直前の累計が無ければ入れない（推測しない）。"""
    order = ("1Q", "2Q", "3Q", "FY")
    out = {}
    for fy, per in tab.items():
        prev = 0.0
        for i, q in enumerate(order):
            v = per.get(q, {}).get(item)
            if v is None:
                prev = None
                continue
            if i == 0:
                out[(fy, q)] = v
            elif prev is not None:
                out[(fy, q)] = v - prev
            prev = v
    return out


def _pct(x):
    return "–" if x is None else "%.1f%%" % (100 * x)


def _m(x):
    return "–" if x is None else "{:,.0f}".format(x / 1e6)


def render_code(con, code, as_of):
    rows = P.load_rows(con, code, as_of)
    name = con.execute("SELECT name FROM companies WHERE code=?", (code,)).fetchone()
    L = ["## %s %s" % (code, name[0] if name else ""), ""]
    ev = P.evaluate(con, code, as_of=as_of)
    L.append("### 1. 登録済み定義での S5c")
    L.append("")
    L.append("- available: **%s** / fired: **%s** / R: **%s**"
             % (ev["available"], ev["fired"], "–" if ev["r"] is None else "%.2f" % ev["r"]))
    L.append("- 根拠: %s" % ev["evidence"])
    tab = fiscal_table(rows)
    cur = P.current_state(rows)
    if cur:
        L.append("- 当期: %s %s（開示 %s）累計OP %s / 会社予想OP %s → 進捗 %s"
                 % (cur["fy_end"], cur["per_type"], cur["disc_date"], _m(cur["op"]),
                    _m(cur["f_op"]), _pct(cur["op"] / cur["f_op"]) if cur["f_op"] else "–"))
        s_series = P.progress_series(rows, "sales")
        s_med, s_years = P.median_progress(s_series, cur["per_type"], exclude_fy=cur["fy_end"])
        if s_med and cur.get("sales") is not None and cur.get("f_sales"):
            sp = cur["sales"] / cur["f_sales"]
            L.append("- 売上（記録のみ・R_sales）: 進捗 %s ÷ 過去中央値 %s（%d年）= **R_sales %.2f**"
                     "（記録ライン %.2f）" % (_pct(sp), _pct(s_med), len(s_years),
                                          sp / s_med, P.NOTE_R_SALES))
        else:
            L.append("- 売上（R_sales）: 計算不能（過去 %d 年）" % len(s_years))
    L.append("")
    for item, label in (("op", "営業利益"), ("sales", "売上高")):
        prog = raw_progress(tab, item)
        L.append("### 2. 進捗率（%s・分母 = その年度の着地）" % label)
        L.append("")
        L.append("| 年度末 | 1Q | 2Q | 3Q | 着地 |")
        L.append("|---|---|---|---|---|")
        for fy in sorted(tab):
            cells = []
            for q in QS:
                p, why = prog.get((fy, q), (None, "行なし"))
                if fy == (cur or {}).get("fy_end"):
                    cells.append("（当期・着地未確定）" if q in tab[fy] else "（未開示）")
                elif (fy, q) not in prog:
                    cells.append("取得不能")
                else:
                    cells.append(_pct(p) + ("（除外: %s）" % why if why else ""))
            L.append("| %s | %s | %s | %s | %s |" % (fy, *cells, _m(tab[fy].get("FY", {}).get(item))))
        L.append("")
        L.append("ばらつき（**当期を除く・除外前の生値**。S5c の分母は除外後の値だけで作る）:")
        L.append("")
        L.append("| 四半期 | n | 中央値 | 標準偏差 | 四分位(25–75%) | IQR | 最小 | 最大 |")
        L.append("|---|---|---|---|---|---|---|---|")
        for q in QS:
            vals = [p for (fy, qq), (p, _w) in prog.items()
                    if qq == q and p is not None and fy != (cur or {}).get("fy_end")]
            s = spread(vals)
            if not s["n"]:
                L.append("| %s | 0 | – | – | – | – | – | – |" % q)
                continue
            L.append("| %s | %d | %s | %s | %s | %s | %s | %s |" % (
                q, s["n"], _pct(s["median"]), _pct(s["sd"]),
                "–" if s["iqr"] is None else "%s–%s" % (_pct(s["q1"]), _pct(s["q3"])),
                _pct(s["iqr"]), _pct(s["min"]), _pct(s["max"])))
        L.append("")
    if cur:
        q = cur["per_type"]
        L.append("### 3. 当期 %s を過去の同四半期と並べる" % q)
        L.append("")
        L.append("| 年度末 | 累計売上 | 累計OP | 当時の会社予想OP | 対会予 OP進捗 | 対会予 売上進捗 | 対着地 OP進捗 |")
        L.append("|---|---|---|---|---|---|---|")
        for fy in sorted(tab):
            r = tab[fy].get(q)
            if not r:
                continue
            land = tab[fy].get("FY", {}).get("op")
            L.append("| %s | %s | %s | %s | %s | %s | %s |" % (
                fy, _m(r["sales"]), _m(r["op"]), _m(r["f_op"]),
                _pct(r["op"] / r["f_op"]) if r.get("f_op") else "–",
                _pct(r["sales"] / r["f_sales"]) if r.get("f_sales") and r.get("sales") is not None else "–",
                _pct(r["op"] / land) if land else "（未確定）"))
        L.append("")
    L.append("### 4. 四半期単独（累計の差分・百万円）")
    L.append("")
    L.append("| 年度末 | 1Q 売上 | 1Q OP | 2Q 売上 | 2Q OP | 3Q 売上 | 3Q OP | 4Q 売上 | 4Q OP |")
    L.append("|---|---|---|---|---|---|---|---|---|")
    ss, so = standalone(tab, "sales"), standalone(tab, "op")
    for fy in sorted(tab):
        cells = []
        for q in ("1Q", "2Q", "3Q", "FY"):
            cells += [_m(ss.get((fy, q))), _m(so.get((fy, q)))]
        L.append("| %s | %s |" % (fy, " | ".join(cells)))
    L.append("")
    L.append("「–」は累計が無い（決算サマリーの取得窓 2021-09-21〜 より前）か、直前の累計が無く差分が取れない期。")
    L.append("")
    return "\n".join(L)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--codes", nargs="+", required=True)
    ap.add_argument("--as-of", default=None)
    ap.add_argument("--out", default=None)
    a = ap.parse_args(argv)
    con = C.connect()
    parts = ["# 進捗率の季節性と S5c(R) の内訳", "",
             "- as_of: %s / 出典: `fins_summary`（J-Quants 決算サマリー・5年）" % (a.as_of or "全期間"),
             "- 定義: backtest_acceptance_criteria「シャドウI-v2」（R ≥ %.2f・最低 %d 年・進捗 %.0f%%〜%.0f%%）"
             % (P.FIRE_R, P.MIN_YEARS, 100 * P.PROGRESS_MIN, 100 * P.PROGRESS_MAX),
             "- S5c は 0点・表示のみ", ""]
    for code in a.codes:
        parts.append(render_code(con, C.normalise_code(code) or code, a.as_of))
    text = "\n".join(parts)
    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            f.write(text)
        C.log("wrote %s" % a.out)
    else:
        sys.stdout.write(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
