"""
forecaster.py — B-3: 機械的予測（経路A+B+C）
A: 単独値時系列のトレンド外挿（前年同四半期 × 直近YoY中央値、調整後ベース）
B: BS先行指標からのPL復元（季節ナイーブ + 契約負債増分 + 建設仮勘定振替分）
C: 適時開示の「証拠」積み上げ（evidence_flag=1のみ。MOU等の約束は自動除外）
出力は「次四半期の単独売上・営業利益予測」と「定常営業利益推定」（逆算DCF比較用）
注意: earnings_calendar は「実行時点で再構築済み」の前提。過去日の再現実行では
先に module_a.earnings_calendar.build_calendar(db, as_of) で当時のカレンダーを再構築すること
"""
import statistics
from datetime import date

from .data_access import (get_pl_series, get_bs_series, get_adjustments,
                          normalized_sales, clamp)

Q_ORDER = {"1Q": 1, "2Q": 2, "3Q": 3, "FY": 4}
C_RECOGNITION = {"order": 0.6, "contract": 0.6, "facility_operation": 0.8}


def _normalized_op(pl_row, adj):
    return (pl_row["operating_profit"] or 0) - adj.get(pl_row["period_end"], 0)


def _next_quarter_info(conn, ticker):
    row = conn.execute(
        "SELECT next_earnings_date, quarter_type, fiscal_year FROM earnings_calendar WHERE ticker=?",
        (ticker,)).fetchone()
    return dict(row) if row else None


def _same_quarter_last_year(pl, quarter_type, fiscal_year):
    for r in pl:
        if r["quarter_type"] == quarter_type and r["fiscal_year"] == fiscal_year - 1:
            return r
    return None


def forecast(conn, ticker, as_of: date) -> dict:
    pl = get_pl_series(conn, ticker, as_of)      # PIT: as_of時点の公開情報のみ
    adj = get_adjustments(conn, ticker, as_of)   # 同上
    nq = _next_quarter_info(conn, ticker)
    if len(pl) < 5 or not nq:
        return {"available": False, "reason": "データ不足または決算カレンダー未登録"}

    ref = _same_quarter_last_year(pl, nq["quarter_type"], nq["fiscal_year"])
    if not ref:
        return {"available": False, "reason": "前年同四半期のデータなし"}

    norm_sales = [normalized_sales(r, adj) for r in pl]
    norm_ops = [_normalized_op(r, adj) for r in pl]

    # ---------- 経路A：トレンド外挿 ----------
    yoys = []
    for i in range(len(pl) - 1, max(3, len(pl) - 1) - 3, -1):  # 直近3四半期のYoY
        if i >= 4 and pl[i - 4]["period_end"] and norm_sales[i - 4] > 0:
            yoys.append(norm_sales[i] / norm_sales[i - 4] - 1)
    yoy_med = statistics.median(yoys) if yoys else 0.0
    yoy_med = max(-0.5, min(yoy_med, 1.0))  # 外挿の暴走防止（±50%/1年でクリップ）
    ref_sales = normalized_sales(ref, adj)
    pred_sales_a = ref_sales * (1 + yoy_med)

    opms = [o / s for o, s in zip(norm_ops[-4:], norm_sales[-4:]) if s > 0]
    opm_est = statistics.median(opms) if opms else 0.0
    pred_op_a = pred_sales_a * opm_est

    # ---------- 経路B：BS先行指標からのPL復元 ----------
    cl = dict(get_bs_series(conn, ticker, "contract_liabilities", as_of))
    d_cl = 0.0
    pes = sorted(cl.keys())
    if len(pes) >= 5:
        d_cl = cl[pes[-1]] - cl[pes[-5]]
    transfer_contrib = 0.0
    cip = dict(get_bs_series(conn, ticker, "construction_in_progress", as_of))
    mach = dict(get_bs_series(conn, ticker, "machinery", as_of))
    cpe, mpe = sorted(cip.keys()), sorted(mach.keys())
    if len(cpe) >= 2 and len(mpe) >= 2:
        d_cip = cip[cpe[-1]] - cip[cpe[-2]]
        d_mach = mach[mpe[-1]] - mach[mpe[-2]]
        if d_cip < -50 and d_mach > -d_cip * 0.5:
            transfer_contrib = min(-d_cip, d_mach) * 0.3  # 稼働率30%の保守仮定
    pred_sales_b = ref_sales + max(d_cl, 0) * 0.5 + transfer_contrib
    pred_op_b = pred_sales_b * opm_est

    # ---------- 経路C：証拠の積み上げ（約束は除外） ----------
    since = date(as_of.year - (1 if as_of.month <= 6 else 0),
                 as_of.month - 6 if as_of.month > 6 else as_of.month + 6, 1).isoformat()
    events = conn.execute(
        "SELECT event_type, amount FROM evidence_events "
        "WHERE ticker=? AND evidence_flag=1 AND event_date>=? AND event_date<=? "
        "AND amount IS NOT NULL",
        (ticker, since, as_of.isoformat())).fetchall()  # PIT: as_of以後のイベントを見ない
    c_extra = sum(e["amount"] * C_RECOGNITION.get(e["event_type"], 0.5) for e in events)

    # ---------- 統合 ----------
    pred_sales = statistics.median([pred_sales_a, pred_sales_b]) + c_extra
    pred_op = statistics.median([pred_op_a, pred_op_b]) + c_extra * 0.4  # 証拠分の限界利益率40%

    # 定常営業利益推定（逆算DCFの要求値と比較する「現在地+傾き」）
    ttm_op = sum(norm_ops[-4:])
    steady_op = ttm_op * (1 + yoy_med) + c_extra * 0.4 * 4

    return {
        "available": True,
        "pred_sales": round(pred_sales, 1),
        "pred_op": round(pred_op, 1),
        "steady_op_est": round(steady_op, 1),
        "path_A": {"sales": round(pred_sales_a, 1), "op": round(pred_op_a, 1),
                   "yoy_median": round(yoy_med, 4)},
        "path_B": {"sales": round(pred_sales_b, 1), "op": round(pred_op_b, 1),
                   "d_contract_liab": round(d_cl, 1), "transfer_contrib": round(transfer_contrib, 1)},
        "path_C": {"evidence_extra_sales": round(c_extra, 1), "n_events": len(events)},
        "opm_est": round(opm_est, 4),
    }
