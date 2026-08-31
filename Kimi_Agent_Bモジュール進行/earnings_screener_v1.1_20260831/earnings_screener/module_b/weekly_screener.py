"""
weekly_screener.py — B-4: 週次選定エンジン（統合パイプライン）
総合スコア = 証拠スコア(S1-S8平均) × 織り込み度ギャップ係数 × 流動性フィルタ
出力：上位20社の週次レポートCSV（ticker/予想売上/予想営業利益/証拠根拠/推奨エントリー日/損切り条件）
"""
import csv
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.jp_calendar import add_business_days, business_days_between
from module_b.data_access import connect, DB
from module_b.forecaster import forecast
from module_b.run_scorers import score_ticker, SCORERS_ALL
from module_a.weekly_batch import extract_candidates

TOP_N = 20


def gap_factor(steady_op_est, required):
    """織り込み度：機械予測の定常OPが市場要求値を上回るほど割安（係数>1）"""
    if not required or required <= 0:
        return 1.0, None
    gap = (steady_op_est - required) / abs(required)
    factor = max(0.2, min(1.0 + gap, 2.0))
    return round(factor, 3), round(gap, 3)


def stop_loss_conditions(scores):
    """イベントベース損切り条件（§8-5：価格ベースにしない）"""
    conds = ["決算で売上が機械予想の95%未満 → 即損切り（翌日寄り指値）",
             "ガイダンス下方修正 → 翌日寄りで撤退"]
    weak = [name for name, v in scores.items()  # nameは安定ID（"S1".."S8"）
            if name != "evidence_score" and v["available"] and v["score"] <= -0.5]
    if weak:
        conds.append(f"弱い証拠（{', '.join(weak)}）が決算で解消しなかった場合 → 撤退")
    return "／".join(conds)


def run_weekly_screen(as_of: date, db_path=None, out_dir=None) -> Path:
    conn = connect(db_path)
    cands = extract_candidates(db_path or DB, as_of)
    rows = []
    for c in cands:
        t = c["ticker"]
        scores = score_ticker(conn, t, SCORERS_ALL, as_of=as_of)  # PIT: 過去日再現でも未来を見ない
        if scores["evidence_score"] is None:
            continue
        fc = forecast(conn, t, as_of)
        if not fc["available"]:
            continue
        req = conn.execute(
            "SELECT required_steady_op_profit FROM reverse_dcf_requirements WHERE ticker=?", (t,)
        ).fetchone()
        gf, gap = gap_factor(fc["steady_op_est"], req[0] if req else None)

        total = scores["evidence_score"] * gf  # 流動性フィルタはuniverse選定時に適用済み（×1.0）
        entry = add_business_days(date.fromisoformat(c["next_earnings_date"]), -15)
        if entry <= as_of:
            entry = add_business_days(as_of, 1)

        evidence = [f"{name}: {v['evidence']}" for name, v in scores.items()
                    if name != "evidence_score" and v["available"] and abs(v["score"]) >= 0.3]
        rows.append({
            "ticker": t,
            "company_name": c["company_name"],
            "next_earnings_date": c["next_earnings_date"],
            "quarter_type": c["quarter_type"],
            "business_days_to_earnings": c["business_days_to_earnings"],
            "pred_sales": fc["pred_sales"],
            "pred_op": fc["pred_op"],
            "evidence_score": scores["evidence_score"],
            "gap_factor": gf,
            "gap": gap,
            "total_score": round(total, 3),
            "entry_date": entry.isoformat(),
            "stop_loss": stop_loss_conditions(scores),
            "evidence_basis": " ｜ ".join(evidence) if evidence else "（強い証拠なし）",
        })
        # Module Cへの受け渡し：機械予測スナップショットを保存
        conn.execute(
            "INSERT OR REPLACE INTO forecast_snapshots "
            "(ticker, as_of_date, fiscal_year, quarter_type, pred_sales, pred_op, "
            "steady_op_est, evidence_score, total_score) VALUES (?,?,?,?,?,?,?,?,?)",
            (t, as_of.isoformat(), c["fiscal_year"], c["quarter_type"],
             fc["pred_sales"], fc["pred_op"], fc["steady_op_est"],
             scores["evidence_score"], round(total, 3)))

    rows.sort(key=lambda x: x["total_score"], reverse=True)
    top = rows[:TOP_N]
    conn.commit()  # forecast_snapshots を確定

    out_dir = Path(out_dir or (Path(__file__).resolve().parents[1] / "data"))
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"weekly_report_{as_of.isoformat()}.csv"
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(top[0].keys()) if top else
                           ["ticker", "company_name", "next_earnings_date", "quarter_type",
                            "business_days_to_earnings", "pred_sales", "pred_op",
                            "evidence_score", "gap_factor", "gap", "total_score",
                            "entry_date", "stop_loss", "evidence_basis"])
        w.writeheader()
        w.writerows(top)
    conn.close()
    return out, top


if __name__ == "__main__":
    out, top = run_weekly_screen(date(2026, 8, 31))
    print(f"weekly report -> {out}")
    for r in top[:10]:
        print(f"{r['ticker']:>5} {r['company_name'][:8]:<8} total={r['total_score']:+.3f} "
              f"(証拠{r['evidence_score']:+.3f} × ギャップ{r['gap_factor']:.2f}) "
              f"予想売上{r['pred_sales']:,.0f} / 予想OP {r['pred_op']:,.0f}")
