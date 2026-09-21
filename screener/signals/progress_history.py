"""screener/signals/progress_history.py — 過去の進捗率（累計 ÷ 通期着地）の系列。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウI-v2」I2-1〜I2-3。
出典は `fins_summary`（J-Quants 決算サマリー・5年）。**このモジュールは定義を実装するだけ。**

    進捗率(年度, 四半期) = その年度 Qn までの累計OP ÷ その年度の着地OP
    過去中央値進捗率     = 同じ四半期の進捗率の中央値（最大5年・最低3年）
    R = 当期の進捗率 ÷ 過去中央値進捗率 = 推定着地OP ÷ 会社予想OP

**S5c は 0点・表示のみ。** ここもスコアには接続しない。
"""
from __future__ import annotations

import os
import statistics
import sys

try:
    from screener import common as C                    # noqa: F401
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C                    # noqa: F401

# ---- I2-3 の値。結果を見て動かさない -------------------------------------
FIRE_R = 1.30               # 営業利益。開示義務ライン（+30%）と同じ
NOTE_R_SALES = 1.10         # 売上。記録のみ
MAX_YEARS = 5
MIN_YEARS = 3
PROGRESS_MIN = 0.05
PROGRESS_MAX = 2.00
QUARTERS = ("1Q", "2Q", "3Q")

TANSHIN = "FinancialStatements"      # doc_type にこれを含む行が短信（累計を持つ）


def load_rows(con, code, as_of=None):
    """その銘柄の短信行（累計を持つ行）。as_of 以前の開示だけ（PIT）。"""
    sql = ("SELECT disc_date, per_type, fy_end, op, sales, f_op, f_sales "
           "FROM fins_summary WHERE code=? AND doc_type LIKE ?"
           + (" AND disc_date<=?" if as_of else "")
           + " ORDER BY disc_date")
    args = (code, "%" + TANSHIN + "%") + ((as_of,) if as_of else ())
    return [dict(zip(("disc_date", "per_type", "fy_end", "op", "sales",
                      "f_op", "f_sales"), r)) for r in con.execute(sql, args)]


def progress_series(rows, item="op"):
    """(fy_end, per_type) -> 進捗率。同じ年度の FY 行（着地）を分母にする。

    着地が無い年度・着地 ≤ 0 の年度・範囲外の進捗率は**入れない**（推測しない）。
    """
    landing, ytd = {}, {}
    for r in rows:
        v = r.get(item)
        if v is None or not r.get("fy_end"):
            continue
        if r["per_type"] == "FY":
            landing[r["fy_end"]] = v            # 同じ年度に複数あれば後の開示が勝つ
        elif r["per_type"] in QUARTERS:
            ytd[(r["fy_end"], r["per_type"])] = v
    out = {}
    for (fy, q), cum in ytd.items():
        land = landing.get(fy)
        if not land or land <= 0:
            continue
        p = cum / land
        if PROGRESS_MIN <= p <= PROGRESS_MAX:
            out[(fy, q)] = p
    return out


def median_progress(series, quarter, exclude_fy=None, max_years=MAX_YEARS):
    """同じ四半期の進捗率の中央値（新しい年度から最大 max_years 本）。"""
    vals = [(fy, p) for (fy, q), p in series.items()
            if q == quarter and fy != exclude_fy]
    vals.sort(reverse=True)
    used = vals[:max_years]
    if len(used) < MIN_YEARS:
        return None, [fy for fy, _ in used]
    return statistics.median(p for _fy, p in used), [fy for fy, _ in used]


def current_state(rows):
    """as_of 時点の「当期」= 最後に開示された四半期の累計と、その時点の通期予想。"""
    for r in reversed(rows):
        if r["per_type"] in QUARTERS and r["op"] is not None:
            return r
    return None


def evaluate(con, code, as_of=None):
    """R（当期進捗率 ÷ 過去中央値進捗率）。**score は常に 0.0（表示のみ）。**"""
    rows = load_rows(con, code, as_of)
    if not rows:
        return _out(False, "決算サマリーの短信行が無い")
    cur = current_state(rows)
    if not cur:
        return _out(False, "当期の四半期累計が無い（直近が通期・または累計が空）")
    if not cur.get("f_op") or cur["f_op"] <= 0:
        return _out(False, "その開示時点の通期営業利益予想が0以下", period=cur["per_type"])
    series = progress_series(rows, "op")
    med, years = median_progress(series, cur["per_type"], exclude_fy=cur["fy_end"])
    if med is None:
        return _out(False, "過去の進捗率が%d年に満たない（取れたのは %d 年）"
                    % (MIN_YEARS, len(years)), period=cur["per_type"], prior_years=years)
    cur_prog = cur["op"] / cur["f_op"]
    r = cur_prog / med
    # 売上も同じ形で（記録のみ）
    r_sales = None
    if cur.get("sales") is not None and cur.get("f_sales"):
        s_series = progress_series(rows, "sales")
        s_med, _ = median_progress(s_series, cur["per_type"], exclude_fy=cur["fy_end"])
        if s_med and cur["f_sales"] > 0:
            r_sales = (cur["sales"] / cur["f_sales"]) / s_med
    return _out(True,
                "%s 累計OP %.0f / 会予 %.0f = 進捗 %.1f%% ÷ 過去中央値 %.1f%%（%d年）= R %.2f"
                % (cur["per_type"], cur["op"], cur["f_op"], 100 * cur_prog,
                   100 * med, len(years), r),
                fired=r >= FIRE_R, r=r, r_sales=r_sales,
                progress=cur_prog, median_progress=med, prior_years=years,
                period=cur["per_type"], fy_end=cur["fy_end"], disc_date=cur["disc_date"],
                cum_op=cur["op"], forecast_op=cur["f_op"],
                landing_op=cur["op"] / med,
                sales_fired=(r_sales is not None and r_sales >= NOTE_R_SALES))


def _out(available, evidence, **kw):
    out = {"score": 0.0, "available": bool(available), "evidence": evidence,
           "fired": False, "r": None, "r_sales": None, "progress": None,
           "median_progress": None, "prior_years": [], "period": None,
           "fy_end": None, "disc_date": None, "cum_op": None, "forecast_op": None,
           "landing_op": None, "sales_fired": False}
    out.update(kw)
    return out
