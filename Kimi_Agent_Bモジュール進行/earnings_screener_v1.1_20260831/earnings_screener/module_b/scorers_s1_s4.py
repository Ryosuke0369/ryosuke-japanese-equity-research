"""
scorers_s1_s4.py — 証拠スコア S1〜S4
全スコアは [-1, +1]。0は中立／データなし。available=False はスコア計算不能（集計で除外）。
設計原則：四半期単独値ベース・一時収入は調整後売上で評価・「約束」は一切見ない。
"""
from .data_access import (connect, get_pl_series, get_bs_series,
                          get_adjustments, normalized_sales, clamp, DAYS_PER_Q)


def _result(score, available, evidence, details=None):
    return {"score": round(score, 3), "available": available,
            "evidence": evidence, "details": details or {}}


# ---------------------------------------------------------------- S1
def score_s1_dso(conn, ticker, as_of=None):
    """S1: 売掛金回転期間（DSO）の改善 — 単独値・調整後売上ベースで前年同期比較"""
    pl = get_pl_series(conn, ticker, as_of)
    adj = get_adjustments(conn, ticker, as_of)
    ar = dict(get_bs_series(conn, ticker, "accounts_receivable", as_of))
    pl = [r for r in pl if r["period_end"] in ar and (r["sales"] or 0) > 0]
    if len(pl) < 5:
        return _result(0.0, False, "DSO比較に必要な5四半期分のデータなし")

    now, prev_yr = pl[-1], pl[-5]
    s_now, s_prev = normalized_sales(now, adj), normalized_sales(prev_yr, adj)
    if s_now <= 0 or s_prev <= 0:
        return _result(0.0, False, "調整後売上がゼロ以下（一時収入のみの期）")

    dso_now = ar[now["period_end"]] / s_now * DAYS_PER_Q
    dso_prev = ar[prev_yr["period_end"]] / s_prev * DAYS_PER_Q
    change = dso_now / dso_prev - 1

    score = clamp(-change / 0.20)  # DSO 20%改善で+1.0
    confirm = ""
    if s_now < s_prev:  # 売上減少下での売掛金減は「縮小」であり改善ではない
        score *= 0.5
        confirm = "（売上減少下のため半減）"

    ev = (f"DSO {dso_prev:.0f}日→{dso_now:.0f}日（{change:+.1%}）"
          f"／調整後売上 {s_prev:.0f}→{s_now:.0f}百万円{confirm}")
    return _result(score, True, ev,
                   {"dso_now": round(dso_now, 1), "dso_prev_year": round(dso_prev, 1),
                    "change_pct": round(change, 4)})


# ---------------------------------------------------------------- S2
def score_s2_contract_liabilities(conn, ticker, as_of=None):
    """S2: 前受金・契約負債の増加 — 将来売上の先行指標（履行義務＝証拠）"""
    cl = get_bs_series(conn, ticker, "contract_liabilities", as_of)
    pl = get_pl_series(conn, ticker, as_of)
    adj = get_adjustments(conn, ticker, as_of)
    if len(cl) < 5 or not pl:
        return _result(0.0, False, "契約負債の5四半期分のデータなし")

    pe_now, v_now = cl[-1]
    pe_q, v_q = cl[-2]
    pe_y, v_y = cl[-5]

    MIN_BASE = 30  # 百万円未満はノイズとして扱わない
    if v_now < MIN_BASE and v_y < MIN_BASE:
        return _result(0.0, False, f"契約負債が少額（{v_now:.0f}百万円）で有意でない")

    yoy = (v_now / v_y - 1) if v_y >= MIN_BASE else (1.0 if v_now >= MIN_BASE * 3 else 0.0)
    qoq = (v_now / v_q - 1) if v_q > 0 else 0.0

    # 売上規模に対する比率でも確認（金額だけの増減で踊らされない）
    pl_map = {r["period_end"]: r for r in pl}
    ratio_term = 0.0
    if pe_now in pl_map and pe_y in pl_map:
        s_now = normalized_sales(pl_map[pe_now], adj)
        s_y = normalized_sales(pl_map[pe_y], adj)
        if s_now > 0 and s_y > 0:
            ratio_change = (v_now / s_now) / (v_y / s_y) - 1
            ratio_term = clamp(ratio_change / 0.30)

    score = clamp(0.5 * clamp(yoy / 0.30) + 0.2 * clamp(qoq / 0.20) + 0.3 * ratio_term)
    ev = (f"契約負債 {v_y:.0f}→{v_now:.0f}百万円（前年比{yoy:+.1%}・前期比{qoq:+.1%}）")
    return _result(score, True, ev, {"v_now": v_now, "yoy": round(yoy, 4), "qoq": round(qoq, 4)})


# ---------------------------------------------------------------- S3
def score_s3_cip_transfer(conn, ticker, as_of=None):
    """S3: 建設仮勘定→機械装置への振替（完成・償却開始＝収益化開始の証拠）"""
    cip = get_bs_series(conn, ticker, "construction_in_progress", as_of)
    mach = get_bs_series(conn, ticker, "machinery", as_of)
    pl = get_pl_series(conn, ticker, as_of)
    adj = get_adjustments(conn, ticker, as_of)
    if len(cip) < 2 or len(mach) < 2:
        return _result(0.0, False, "建設仮勘定/機械装置のデータなし")

    c_now, c_prev = cip[-1][1], cip[-2][1]
    m_now, m_prev = mach[-1][1], mach[-2][1]
    d_cip, d_mach = c_now - c_prev, m_now - m_prev

    if c_prev < 50 and c_now < 50:
        return _result(0.0, False, "建設仮勘定なし（設備投資フェーズ外）")

    s_now = normalized_sales(pl[-1], adj) if pl else 0
    if d_cip < -50 and d_mach >= -d_cip * 0.5:
        # 振替イベント：CIP減少分の半分以上が機械装置増加で説明できる
        transfer = min(-d_cip, d_mach)
        scale = transfer / s_now if s_now > 0 else 0
        score = clamp(scale / 1.0, 0.0, 1.0)  # 四半期売上相当の振替で満点
        ev = (f"建設仮勘定 {c_prev:.0f}→{c_now:.0f}百万円（{d_cip:+.0f}）"
              f"／機械装置 {m_prev:.0f}→{m_now:.0f}百万円（{d_mach:+.0f}）→ 完成・稼働開始")
        return _result(score, True, ev, {"transfer": transfer, "scale_vs_sales": round(scale, 3)})

    if d_cip > 0:
        return _result(0.0, True,
                       f"建設仮勘定が増加中（{c_prev:.0f}→{c_now:.0f}百万円）：まだ建設中。"
                       f"完成振替は将来シグナル（現時点では加点しない）",
                       {"status": "building"})
    return _result(0.0, True, "建設仮勘定に大きな変動なし", {"status": "flat"})


# ---------------------------------------------------------------- S4
def score_s4_cash_quality(conn, ticker, as_of=None):
    """S4: 営業CF vs 営業利益（キャッシュ品質）。直近4半期のCF開示分で評価"""
    pl = get_pl_series(conn, ticker, as_of)
    cf_rows = [r for r in pl if r["operating_cf"] is not None]
    if len(cf_rows) < 2:
        return _result(0.0, False, "営業CFの開示期が不足（1Q/3Qは短信にCF記載なし）")

    recent = cf_rows[-2:]  # 直近約1年分（2Q=上半期単独、FY=下半期単独）
    sum_cf = sum(r["operating_cf"] for r in recent)
    # 対応期間の営業利益：CF開示期の直前Qを含めて半期単位で集計
    sum_op = 0.0
    for r in recent:
        idx = next(i for i, x in enumerate(pl) if x["period_end"] == r["period_end"])
        half = pl[max(0, idx - 1):idx + 1]
        sum_op += sum(x["operating_profit"] or 0 for x in half)
    if sum_op <= 0:
        return _result(0.0, False, f"対象期間の営業利益が{sum_op:.0f}百万円（赤字期は品質評価しない）")

    ratio = sum_cf / sum_op
    score = clamp((ratio - 1.0) / 0.5)  # CF/利益=1.0で中立、1.5で+1、0.5で-1
    ev = f"直近1年 営業CF {sum_cf:.0f}百万円 / 営業利益 {sum_op:.0f}百万円 = {ratio:.2f}倍"
    return _result(score, True, ev, {"cf_op_ratio": round(ratio, 3)})


SCORERS_S1_S4 = {
    "S1": score_s1_dso,
    "S2": score_s2_contract_liabilities,
    "S3": score_s3_cip_transfer,
    "S4": score_s4_cash_quality,
}

# 表示名（IDは永続・ラベルは自由に変更してよい。統合時のID体系差替えに対応）
SIGNAL_LABELS = {
    "S1": "DSO改善", "S2": "契約負債増", "S3": "建設仮勘定振替", "S4": "CF品質",
}
