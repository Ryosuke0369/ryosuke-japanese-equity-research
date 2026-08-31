"""
exit_engine.py — Module C: 出口ルール自動化
決算実績と機械予測を比較し、4分岐で売却判断を下す。

分岐（引継ぎ資料§4-C）：
  1. 売上ミス（<予想95%）            → 即損切り（全量・翌日寄り指値）
  2. 売上・利益とも予想超過＋上方修正 → 50%利確、残りT+3追跡
  3. 予想通り（95〜105%レンジ）       → 全利確
  4. 方向性なし                      → 撤退

執行：PTS出来高が直近日次平均の30%以上 → PTS売却、それ以外 → 翌日寄り指値
"""
import csv
import sqlite3
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

DB = Path(__file__).resolve().parents[1] / "data" / "screener.db"
PTS_THRESHOLD = 0.30
MISS_LINE = 0.95
INLINE_BAND = 1.05

ACTIONS = {
    "STOP_LOSS": "即損切り（全量）",
    "TP50_TRAIL": "50%利確・残りT+3追跡",
    "TP_ALL": "全利確",
    "EXIT": "撤退（全量）",
}


def decide_exit(pred_sales: float, pred_op: float,
                actual_sales: float, actual_op: float,
                prev_guidance_op: float | None, new_guidance_op: float | None,
                pts_ratio: float | None) -> dict:
    """純粋関数：予測・実績・ガイダンス・PTS出来高から売却判断を返す"""
    sales_ratio = actual_sales / pred_sales if pred_sales > 0 else 1.0
    if pred_op > 0:
        op_ratio = actual_op / pred_op
        op_beat = op_ratio >= 1.0
        op_inline = 0.95 <= op_ratio <= 1.05
    else:  # 赤字予想時は比率が不安定なので絶対額で判定
        op_ratio = None
        op_beat = actual_op > pred_op
        op_inline = abs(actual_op - pred_op) <= max(abs(pred_op) * 0.05, pred_sales * 0.01)

    guidance_raise = (prev_guidance_op is not None and new_guidance_op is not None
                      and new_guidance_op > prev_guidance_op)

    # --- 4分岐（順序固定：ミス判定を最優先） ---
    if sales_ratio < MISS_LINE:
        action, fraction = "STOP_LOSS", 1.0
        reason = f"売上ミス（予想比{sales_ratio:.1%} < 95%）"
    elif sales_ratio >= 1.0 and op_beat and guidance_raise:
        action, fraction = "TP50_TRAIL", 0.5
        reason = (f"売上{sales_ratio:.1%}・利益とも予想超過＋ガイダンス上方修正"
                  f"（{prev_guidance_op:,.0f}→{new_guidance_op:,.0f}百万円）")
    elif sales_ratio <= INLINE_BAND and op_inline:
        action, fraction = "TP_ALL", 1.0
        reason = f"予想通り（売上{sales_ratio:.1%}）"
    else:
        action, fraction = "EXIT", 1.0
        reason = f"方向性なし（売上{sales_ratio:.1%}・上方修正{'あり' if guidance_raise else 'なし'}）"

    # --- 執行会場 ---
    if pts_ratio is not None and pts_ratio >= PTS_THRESHOLD:
        venue = f"PTS売却（PTS出来高/日次平均={pts_ratio:.0%} ≥ 30%）"
    else:
        venue = (f"翌日寄り指値（PTS出来高/日次平均={pts_ratio:.0%} < 30%）"
                 if pts_ratio is not None else "翌日寄り指値（PTSデータなし）")

    return {"action": action, "action_label": ACTIONS[action],
            "sell_fraction": fraction, "venue": venue, "reason": reason,
            "sales_ratio": round(sales_ratio, 4),
            "op_ratio": round(op_ratio, 4) if op_ratio is not None else None,
            "guidance_raise": guidance_raise}


def run_exit_batch(db_path=None, announce_date: date | None = None) -> tuple[Path, list]:
    """着弾した決算（earnings_actuals）について一括で出口判断を生成"""
    conn = sqlite3.connect(str(db_path or DB))
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    actuals = cur.execute(
        "SELECT * FROM earnings_actuals ORDER BY announce_date, ticker").fetchall()
    decisions = []
    for a in actuals:
        t = a["ticker"]
        snap = cur.execute(
            "SELECT * FROM forecast_snapshots WHERE ticker=? AND fiscal_year=? AND quarter_type=? "
            "ORDER BY as_of_date DESC LIMIT 1",
            (t, a["fiscal_year"], a["quarter_type"])).fetchone()
        if not snap:
            continue  # ポジション対象外（予測なし＝エントリーしていない）

        prev_fc = cur.execute(
            "SELECT forecast_op FROM company_forecasts WHERE ticker=? AND fiscal_year=? "
            "AND source_date<? ORDER BY source_date DESC LIMIT 1",
            (t, a["fiscal_year"], a["announce_date"])).fetchone()
        pts = cur.execute(
            "SELECT pts_volume, daily_avg_volume FROM pts_observations WHERE ticker=? AND event_date=?",
            (t, a["announce_date"])).fetchone()
        pts_ratio = (pts["pts_volume"] / pts["daily_avg_volume"]
                     if pts and pts["daily_avg_volume"] else None)

        d = decide_exit(snap["pred_sales"], snap["pred_op"],
                        a["actual_sales"], a["actual_op"],
                        prev_fc["forecast_op"] if prev_fc else None,
                        a["guidance_op"], pts_ratio)
        decisions.append({
            "ticker": t, "announce_date": a["announce_date"],
            "quarter_type": a["quarter_type"], "fiscal_year": a["fiscal_year"],
            "pred_sales": snap["pred_sales"], "actual_sales": a["actual_sales"],
            "pred_op": snap["pred_op"], "actual_op": a["actual_op"],
            **d,
        })
    conn.close()

    out = Path(db_path or DB).parent / f"exit_report_{(announce_date or date.today()).isoformat()}.csv"
    if decisions:
        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(decisions[0].keys()))
            w.writeheader()
            w.writerows(decisions)
    return out, decisions


if __name__ == "__main__":
    out, decisions = run_exit_batch(announce_date=date(2026, 9, 15))
    print(f"exit report -> {out}")
    from collections import Counter
    print(Counter(d["action"] for d in decisions))
