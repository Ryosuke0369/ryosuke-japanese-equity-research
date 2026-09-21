"""screener/signals/s5c_landing.py — S5c: 着地推定と会社予想の乖離（開示義務ライン）。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウI」（2026-09-21 登録）。
**このモジュールは定義を実装するだけで、閾値を決めない。**

なぜ S5 と別に作るのか
----------------------
S5 が測るのは「累計進捗 ÷ 経過四半期比」。一方、東証の適時開示規則で**上方修正が義務**に
なるのは「公表済み予想との差異が 売上10%以上、または営業利益等で30%以上」。**別の量**である。
3441 山王は Q3 進捗113% でも着地は会予の 93.9%（差異 −6.1%）で、修正義務は発生していない。

    推定着地OP = 当期の累計OP ÷ 前年同期の進捗率
    前年同期の進捗率 = 前年同じ四半期までの累計OP ÷ 前年の着地OP実績
    乖離 = 推定着地OP ÷ 会社予想OP − 1      （発火は乖離 ≥ +30%）

季節性を明示的にモデル化せず**前年の実績パターン**で代用する。3〜5年の四半期別構成比は
現データでは作れない（4年ぶん揃うのは4社・5年は0社。calibration_backlog §46 の実測）。

**S5c は 0点・表示のみ。** スコアには接続しない。
"""
from __future__ import annotations

import os
import sys

try:
    from screener import common as C                    # noqa: F401
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C                    # noqa: F401

from screener.signals import span_scorers as S

# ---- I-3 の値。結果を見て動かさない -------------------------------------
FIRE_OP_PCT = 0.30          # 営業利益の開示義務ライン（+30%）
NOTE_SALES_PCT = 0.10       # 売上の10%。記録のみ（主判定に使わない）
MAX_PRIOR_YEARS = 2         # 前年進捗は最大2年ぶん・中央値
PROGRESS_MIN = 0.05         # 前年同期進捗の下限（これ未満は対象外）
PROGRESS_MAX = 2.00         # 同上限
MAX_Q = 3                   # 対象は Q1〜Q3（Q4 は着地そのもの）


def _median(xs):
    xs = sorted(xs)
    n = len(xs)
    if not n:
        return None
    return xs[n // 2] if n % 2 else (xs[n // 2 - 1] + xs[n // 2]) / 2


def _full_year(con, ticker, fy, vis):
    """その年度の**着地実績**（通期の累計）。組めなければ None。"""
    rows = [r for r in con.execute(
        "SELECT period_end, period_start, span_q, sales, operating_profit "
        "FROM quarterly_standalone_all WHERE ticker=? AND fiscal_year=? AND is_valid=1",
        (ticker, fy)) if vis is None or r[0] in vis]
    # span=4（通期1本）があればそれ。無ければ span の合計が4になる組み合わせ。
    for r in rows:
        if (r[2] or 1) == 4 and r[4] is not None:
            return {"sales": r[3], "op": r[4]}
    tot_span = sum(r[2] or 1 for r in rows)
    ops = [r[4] for r in rows]
    if tot_span == 4 and rows and all(o is not None for o in ops):
        return {"sales": sum(r[3] for r in rows if r[3] is not None), "op": sum(ops)}
    return None


def _ytd_at(con, ticker, fy, n_q, vis):
    """その年度の Q1〜n_q の累計（`span_scorers._ytd` と同じ組み立て方）。"""
    y = S._ytd(con, ticker, fy, vis)
    if y and y[1] == n_q:
        return {"period_end": y[0], "n_q": y[1], "sales": y[2], "op": y[3]}
    # _ytd は「最も遠くまで覆える区間」を返すので、n_q と違えば作り直す
    by_start = {}
    for r in con.execute(
            "SELECT period_end, period_start, span_q, sales, operating_profit "
            "FROM quarterly_standalone_all WHERE ticker=? AND fiscal_year=? AND is_valid=1",
            (ticker, fy)):
        if vis is not None and r[0] not in vis:
            continue
        st = S.SM.parse_pe(r[1] or r[0])
        if st and st[0] == fy:
            by_start.setdefault(st[1], []).append(r)
    used, q = [], 1
    while q <= n_q:
        cand = [r for r in by_start.get(q, []) if q + (r[2] or 1) - 1 <= n_q]
        if not cand:
            return None
        r = min(cand, key=lambda x: x[2] or 1)
        used.append(r)
        q += (r[2] or 1)
    ops = [r[4] for r in used]
    if any(o is None for o in ops):
        return None
    return {"period_end": max(r[0] for r in used), "n_q": n_q,
            "sales": sum(r[3] for r in used if r[3] is not None), "op": sum(ops)}


def evaluate(con, ticker, as_of=None):
    """S5c。**スコアは常に 0.0**（表示のみ）。available は「計算できたか」。"""
    vis = S.visible_periods(con, ticker, as_of)
    a = S._iso(as_of)
    fc = con.execute(
        "SELECT fiscal_year, forecast_sales, forecast_op FROM company_forecasts "
        "WHERE ticker=?" + (" AND source_date<=?" if a else "")
        + " ORDER BY source_date DESC LIMIT 1",
        (ticker, a) if a else (ticker,)).fetchone()
    if not fc:
        return _out(False, "会社通期予想なし")
    fy, fc_sales, fc_op = fc[0], fc[1], fc[2]
    if not fc_op or fc_op <= 0:
        return _out(False, "会社予想の営業利益が0以下（比率が壊れるので対象外）",
                    forecast_op=fc_op)

    cur = S._ytd(con, ticker, fy, vis)
    if not cur:
        return _out(False, "当期の累計を組み立てられない")
    pe, n_q, cum_sales, cum_op = cur
    if n_q > MAX_Q:
        return _out(False, "Q4 は着地そのものなので対象外", period=pe)
    if cum_op is None:
        return _out(False, "当期の累計営業利益が取れない", period=pe)

    progresses, used_years = [], []
    for back in range(1, MAX_PRIOR_YEARS + 1):
        py = fy - back
        prior_ytd = _ytd_at(con, ticker, py, n_q, vis)
        prior_fy = _full_year(con, ticker, py, vis)
        if not prior_ytd or not prior_fy or not prior_fy["op"]:
            continue
        if prior_fy["op"] <= 0:
            continue                       # 前年が赤字だと進捗率が意味を持たない
        pr = prior_ytd["op"] / prior_fy["op"]
        if not (PROGRESS_MIN <= pr <= PROGRESS_MAX):
            continue
        progresses.append(pr)
        used_years.append(py)
    if not progresses:
        return _out(False, "前年同期の進捗率が取れない（%d年ぶん探した）" % MAX_PRIOR_YEARS,
                    period=pe, elapsed_q=n_q)

    prog = _median(progresses)
    landing_op = cum_op / prog
    gap_op = landing_op / fc_op - 1
    gap_sales = None
    if fc_sales and fc_sales > 0 and cum_sales is not None:
        ps = [p for p in progresses]       # 売上は同じ進捗率を流用しない
        prior_sales = []
        for py in used_years:
            py_ytd = _ytd_at(con, ticker, py, n_q, vis)
            py_fy = _full_year(con, ticker, py, vis)
            if py_ytd and py_fy and py_fy["sales"]:
                prior_sales.append(py_ytd["sales"] / py_fy["sales"])
        if prior_sales:
            gap_sales = (cum_sales / _median(prior_sales)) / fc_sales - 1
        del ps
    return _out(True,
                "Q%d累計OP %.0f ÷ 前年同期進捗 %.1f%% = 着地推定 %.0f / 会予 %.0f → 乖離 %+.1f%%"
                % (n_q, cum_op, 100 * prog, landing_op, fc_op, 100 * gap_op),
                fired=gap_op >= FIRE_OP_PCT, gap_op=gap_op, gap_sales=gap_sales,
                landing_op=landing_op, forecast_op=fc_op, cum_op=cum_op,
                progress_prior=prog, prior_years=used_years, elapsed_q=n_q, period=pe,
                sales_fired=(gap_sales is not None and gap_sales >= NOTE_SALES_PCT))


def _out(available, evidence, **kw):
    """S5 と同じ形の dict。**score は常に 0.0（表示のみ）。**"""
    r = {"score": 0.0, "available": bool(available), "evidence": evidence,
         "fired": False, "gap_op": None, "gap_sales": None, "landing_op": None,
         "forecast_op": None, "cum_op": None, "progress_prior": None,
         "prior_years": [], "elapsed_q": None, "period": None, "sales_fired": False}
    r.update(kw)
    return r
