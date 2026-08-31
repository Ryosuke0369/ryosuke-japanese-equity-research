"""
generate_mock_actuals.py — Module C検証用：決算実績の模擬着弾
T-15候補（=予測スナップショット保有銘柄）の決算実績・ガイダンス修正・PTS出来高を生成。
シナリオ：予想超過+上方修正 / 予想通り / 売上ミス / 方向性なし をhashで割付
"""
import hashlib
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from mock.generate_mock_data import quarter_period_end

DB = Path(__file__).resolve().parents[1] / "data" / "screener.db"


def h(seed: str) -> float:
    return int(hashlib.md5(seed.encode()).hexdigest(), 16) % 10000 / 10000.0


def main():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    snaps = cur.execute(
        "SELECT s.*, c.next_earnings_date, u.fiscal_year_end FROM forecast_snapshots s "
        "JOIN earnings_calendar c ON c.ticker=s.ticker "
        "JOIN universe u ON u.ticker=s.ticker").fetchall()

    n = 0
    for s in snaps:
        t = s["ticker"]
        scenario_r = h(t + "scn")
        if scenario_r < 0.25:
            scenario = "beat_raise"
            s_f, o_f = 1.05 + h(t + "a") * 0.10, 1.05 + h(t + "b") * 0.25
        elif scenario_r < 0.60:
            scenario = "inline"
            s_f, o_f = 0.97 + h(t + "a") * 0.06, 0.97 + h(t + "b") * 0.06
        elif scenario_r < 0.80:
            scenario = "miss"
            s_f, o_f = 0.80 + h(t + "a") * 0.14, 0.70 + h(t + "b") * 0.24
        else:
            scenario = "nodirection"  # 売上は並〜超過だが利益が振るわず修正なし
            s_f, o_f = 1.00 + h(t + "a") * 0.08, 0.85 + h(t + "b") * 0.10

        actual_sales = round(s["pred_sales"] * s_f, 1)
        actual_op = round(s["pred_op"] * o_f, 1) if s["pred_op"] > 0 else round(s["pred_op"] - 50 * (1 - o_f), 1)

        # ガイダンス：beat_raiseのみ上方修正行を追加
        fy = s["fiscal_year"]
        prev = cur.execute(
            "SELECT forecast_sales, forecast_op FROM company_forecasts WHERE ticker=? AND fiscal_year=? "
            "ORDER BY source_date DESC LIMIT 1", (t, fy)).fetchone()
        if prev is None:
            # モックの補完：当該年度の期初予想が未生成の場合は実績から推定して登録
            prev_sales = round(actual_sales * 4 / max(s_f, 0.5) * 0.98, 1)
            prev_op = round(max(actual_op, 1) * 4 / max(o_f, 0.5) * 0.95, 1)
            cur.execute(
                "INSERT OR REPLACE INTO company_forecasts (ticker, fiscal_year, forecast_sales, forecast_op, source_date) "
                "VALUES (?,?,?,?,?)", (t, fy, prev_sales, prev_op, "2026-05-15"))
            prev = {"forecast_sales": prev_sales, "forecast_op": prev_op}
        if prev and scenario == "beat_raise":
            g_sales = round(prev["forecast_sales"] * (1.03 + h(t + "g") * 0.10), 1)
            g_op = round(prev["forecast_op"] * (1.05 + h(t + "g") * 0.15), 1)
            cur.execute(
                "INSERT OR REPLACE INTO company_forecasts (ticker, fiscal_year, forecast_sales, forecast_op, source_date) "
                "VALUES (?,?,?,?,?)", (t, fy, g_sales, g_op, s["next_earnings_date"]))
        elif prev:
            g_sales, g_op = prev["forecast_sales"], prev["forecast_op"]
        else:
            g_sales = g_op = None

        q_num = {"1Q": 1, "2Q": 2, "3Q": 3, "FY": 4}[s["quarter_type"]]
        pe = quarter_period_end(s["fiscal_year_end"], fy, q_num).isoformat()
        cur.execute(
            "INSERT OR REPLACE INTO earnings_actuals "
            "(ticker, fiscal_year, quarter_type, period_end, announce_date, actual_sales, actual_op, "
            "guidance_sales, guidance_op, source) VALUES (?,?,?,?,?,?,?,?,?, 'mock')",
            (t, fy, s["quarter_type"], pe, s["next_earnings_date"],
             actual_sales, actual_op, g_sales, g_op))

        pts_ratio = 0.10 + h(t + "pts") * 0.45  # 10〜55%
        cur.execute(
            "INSERT OR REPLACE INTO pts_observations (ticker, event_date, pts_volume, daily_avg_volume) "
            "VALUES (?,?,?,?)",
            (t, s["next_earnings_date"], round(100_000 * pts_ratio), 100_000))
        n += 1

    conn.commit()
    print(f"mock actuals: {n}件着弾")
    conn.close()


if __name__ == "__main__":
    main()
