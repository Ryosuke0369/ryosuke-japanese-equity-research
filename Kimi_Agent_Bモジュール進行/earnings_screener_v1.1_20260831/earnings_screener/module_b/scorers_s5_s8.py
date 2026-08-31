"""
scorers_s5_s8.py — 証拠スコア S5〜S8
S5: 通期予想に対する累計進捗率の異常検出（季節性補正つき・一時収入は調整後で評価）
S6: 在庫の両義スコア（数量/価格分解 × 売上加速度）
S7: 収益認識パターン変化（一時点移転 → 一定期間移転へのシフトはプラス）
S8: 潜在株式調整後の希薄化リスク（オーバーハングとその増加）
"""
from .data_access import (get_pl_series, get_bs_series, get_adjustments,
                          get_latest_forecast, normalized_sales, clamp)


def _result(score, available, evidence, details=None):
    return {"score": round(score, 3), "available": available,
            "evidence": evidence, "details": details or {}}


def _normalized_op(pl_row, adjustments):
    """調整後営業利益。一時収入は営業利益に全額フローすると仮定
    （注記で利益インパクトが分離できる場合は one_time_gain 等の別item_keyで上書きする運用）"""
    adj = adjustments.get(pl_row["period_end"], 0)
    return (pl_row["operating_profit"] or 0) - adj


def _seasonal_share(pl, quarter_type):
    """過去年度の「この四半期までの累計売上シェア」の平均（季節性補正用）"""
    by_fy = {}
    for r in pl:
        by_fy.setdefault(r["fiscal_year"], []).append(r)
    shares = []
    for fy, rows in by_fy.items():
        rows.sort(key=lambda x: x["period_end"])
        full = sum(x["sales"] or 0 for x in rows)
        if len(rows) < 4 or full <= 0:
            continue  # 通期まで揃った年度のみ
        q_order = {"1Q": 1, "2Q": 2, "3Q": 3, "FY": 4}
        cum = sum(x["sales"] or 0 for x in rows
                  if q_order[x["quarter_type"]] <= q_order[quarter_type])
        if quarter_type != "FY":
            shares.append(cum / full)
    return sum(shares) / len(shares) if shares else None


# ---------------------------------------------------------------- S5
def score_s5_progress_anomaly(conn, ticker, as_of=None):
    """S5: 累計進捗率の異常検出。
    例：Q3時点で113% → 暗黙のQ4が赤字という不合理 → 上方修正の先回り（プラス）。
    逆に大きく未達 → 下方修正リスク（マイナス）。"""
    fc = get_latest_forecast(conn, ticker, as_of)  # 設計原則: スコアラーはDBを直叩きしない
    if not fc:
        return _result(0.0, False, "会社通期予想なし")
    fy, fc_sales, fc_op = fc["fiscal_year"], fc["forecast_sales"], fc["forecast_op"]

    pl_all = get_pl_series(conn, ticker, as_of)
    adj = get_adjustments(conn, ticker, as_of)
    cur = [r for r in pl_all if r["fiscal_year"] == fy]
    if not cur:
        return _result(0.0, False, "当該年度の実績データなし")
    latest_q = cur[-1]["quarter_type"]
    if latest_q == "FY":
        return _result(0.0, False, "通期発表済み（進捗評価の対象外）")

    cum_sales = sum(normalized_sales(r, adj) for r in cur)
    cum_sales_rep = sum(r["sales"] or 0 for r in cur)
    cum_op = sum(_normalized_op(r, adj) for r in cur)

    expected = _seasonal_share(pl_all, latest_q) or {"1Q": 0.25, "2Q": 0.50, "3Q": 0.75}[latest_q]

    prog_s = cum_sales / fc_sales if fc_sales else 0
    ratio_s = prog_s / expected if expected > 0 else 1.0
    # OPは会社予想がプラスの場合のみ比率評価（赤字予想の進捗評価は不安定なため）
    if fc_op > 0:
        prog_o = cum_op / fc_op
        ratio_o = prog_o / expected if expected > 0 else 1.0
        ratio = 0.4 * ratio_s + 0.6 * ratio_o  # 利益進捗を重視
    else:
        prog_o, ratio, ratio_o = None, ratio_s, None

    score = clamp((ratio - 1.0) / 0.25)  # 季節性期待比+25%で満点
    rep_note = f"／報告値ベース売上進捗 {cum_sales_rep/fc_sales:.1%}" if cum_sales_rep != cum_sales else ""
    op_note = f"、調整後OP進捗 {prog_o:.1%}" if prog_o is not None else ""
    ev = (f"{latest_q}時点：調整後売上進捗 {prog_s:.1%}{op_note}（季節性期待 {expected:.0%}）{rep_note} "
          f"→ 期待比 {ratio:.2f}倍")
    return _result(score, True, ev, {"progress_sales": round(prog_s, 4),
                                     "expected": round(expected, 4), "ratio": round(ratio, 3)})


# ---------------------------------------------------------------- S6
def score_s6_inventory(conn, ticker, as_of=None):
    """S6: 在庫の両義スコア。在庫増を「仕込み」か「滞留」かに分解判定する。
    - 数量要因の売上原価増 × 売上加速 → 仕込み（加点）
    - 数量要因の増加 × 売上減速 → 滞留（減点）
    - 価格要因主導（為替・市況）の在庫増は評価増として中立"""
    inv = get_bs_series(conn, ticker, "inventory", as_of)
    pl = get_pl_series(conn, ticker, as_of)
    adj = get_adjustments(conn, ticker, as_of)
    if len(inv) < 5 or len(pl) < 6:
        return _result(0.0, False, "在庫の前年比較に必要なデータなし")

    inv_now, inv_y = inv[-1][1], inv[-5][1]
    inv_yoy = inv_now / inv_y - 1 if inv_y > 0 else 0.0
    if abs(inv_yoy) < 0.10:
        return _result(0.0, True, f"在庫変動が小さい（前年比{inv_yoy:+.1%}）→ シグナルなし",
                       {"inv_yoy": round(inv_yoy, 4)})

    # 数量/価格分解の確認
    cq = [r["cogs_quantity"] for r in pl]
    if cq[-1] is None or cq[-5] is None:
        return _result(0.0, False,
                       f"在庫は前年比{inv_yoy:+.1%}だが原価の数量/価格分解データなし → 中立")
    qty_yoy = cq[-1] / cq[-5] - 1 if cq[-5] else 0.0
    cp = [r["cogs_price"] for r in pl]
    price_yoy = (cp[-1] / cp[-5] - 1) if (cp[-1] is not None and cp[-5]) else 0.0

    # 売上加速度（前年比の変化）
    s = [normalized_sales(r, adj) for r in pl]
    yoy_now = s[-1] / s[-5] - 1 if s[-5] > 0 else 0.0
    yoy_prev = s[-2] / s[-6] - 1 if s[-6] > 0 else 0.0
    accel = yoy_now - yoy_prev

    base = clamp(abs(inv_yoy) / 0.40, 0.0, 1.0)  # 在庫変動40%で満点
    if inv_yoy > 0:
        if qty_yoy <= 0.02 and price_yoy > 0.05:
            return _result(0.0, True,
                           f"在庫{inv_yoy:+.1%}は価格要因主導（数量{qty_yoy:+.1%}・価格{price_yoy:+.1%}）"
                           f"→ 評価増とみなし中立", {"inv_yoy": round(inv_yoy, 4)})
        if accel >= 0:
            verdict, score = "仕込み（数量増×売上加速）", base * clamp(accel / 0.10, 0.15, 1.0)
        else:
            verdict, score = "滞留（数量増×売上減速）", -base * clamp(-accel / 0.10, 0.15, 1.0)
    else:
        if accel > 0:
            verdict, score = "売れ行きによる在庫減（軽微に加点）", 0.3 * base
        else:
            verdict, score = "販売不振による在庫調整（軽微に減点）", -0.3 * base

    ev = (f"在庫 {inv_y:.0f}→{inv_now:.0f}百万円（{inv_yoy:+.1%}）／数量要因{qty_yoy:+.1%}・"
          f"売上YoY {yoy_prev:+.1%}→{yoy_now:+.1%}（加速度{accel:+.1%}pt）→ {verdict}")
    return _result(score, True, ev, {"inv_yoy": round(inv_yoy, 4), "qty_yoy": round(qty_yoy, 4),
                                     "accel": round(accel, 4), "verdict": verdict})


# ---------------------------------------------------------------- S7
def score_s7_revenue_recognition(conn, ticker, as_of=None):
    """S7: 収益認識パターン変化。一時点移転→一定期間移転へのシフトは
    収益の見通し性向上（リカーリング化・長期契約化の証拠）として加点"""
    ot = dict(get_bs_series(conn, ticker, "rev_over_time", as_of))
    pt = dict(get_bs_series(conn, ticker, "rev_point_in_time", as_of))
    if len(ot) < 5:
        return _result(0.0, False, "収益認識内訳の開示なし（四半期注記非開示）")

    pes = sorted(ot.keys())
    def share(pe):
        total = ot[pe] + pt.get(pe, 0)
        return ot[pe] / total if total > 0 else None
    sh_now, sh_prev = share(pes[-1]), share(pes[-5])
    if sh_now is None or sh_prev is None:
        return _result(0.0, False, "収益認識内訳が不完全")

    change = sh_now - sh_prev
    score = clamp(change / 0.15)  # 一定期間比率+15ptで満点
    ev = (f"一定期間移転の比率 {sh_prev:.0%}→{sh_now:.0%}（{change:+.0%}pt）"
          + ("→ 収益の見通し性向上" if change > 0.03 else
             "→ 一時点移転化（収益の不安定化）" if change < -0.03 else "→ 大きな変化なし"))
    return _result(score, True, ev, {"share_now": round(sh_now, 3), "change": round(change, 4)})


# ---------------------------------------------------------------- S8
def score_s8_dilution(conn, ticker, as_of=None):
    """S8: 潜在株式調整後EPSの希薄化リスク。
    オーバーハング（潜在株式/発行済株式）の水準と増加の両方で減点"""
    out = dict(get_bs_series(conn, ticker, "shares_outstanding", as_of))
    dil = dict(get_bs_series(conn, ticker, "diluted_shares", as_of))
    if len(out) < 2 or not dil:
        return _result(0.0, False, "株式数データなし")

    pes = sorted(out.keys())
    over_now = dil[pes[-1]] / out[pes[-1]] - 1
    idx_prev = max(0, len(pes) - 5)
    over_prev = dil[pes[idx_prev]] / out[pes[idx_prev]] - 1

    if over_now <= 0.03:
        return _result(0.0, True, f"希薄化オーバーハング {over_now:.1%}（無視できる水準）")

    level = -clamp((over_now - 0.03) / 0.12, 0.0, 1.0)      # 15%で満点減点
    delta = over_now - over_prev
    growth = -clamp((delta - 0.02) / 0.08, 0.0, 1.0) if delta > 0.02 else 0.0
    score = clamp(level + growth)
    ev = (f"希薄化オーバーハング {over_prev:.1%}→{over_now:.1%}（{delta:+.1%}pt）"
          f"→ 潜在株式 {dil[pes[-1]] - out[pes[-1]]:,.0f}株分のEPS希薄化リスク")
    return _result(score, True, ev, {"overhang_now": round(over_now, 4),
                                     "overhang_change": round(delta, 4)})


SCORERS_S5_S8 = {
    "S5": score_s5_progress_anomaly,
    "S6": score_s6_inventory,
    "S7": score_s7_revenue_recognition,
    "S8": score_s8_dilution,
}

SIGNAL_LABELS = {
    "S5": "進捗率異常", "S6": "在庫両義", "S7": "収益認識変化", "S8": "希薄化リスク",
}
