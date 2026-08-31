"""
generate_mock_rdcf.py — 逆算DCF要求値のモック（既存バリュエーション層の出力代替）
実運用では reverse_dcf_requirements テーブルに既存パイプラインが書き込む。
"""
import hashlib
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
DB = Path(__file__).resolve().parents[1] / "data" / "screener.db"
AS_OF = "2026-08-31"


def h(seed: str) -> float:
    return int(hashlib.md5(seed.encode()).hexdigest(), 16) % 10000 / 10000.0


def main():
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    tickers = [r[0] for r in cur.execute(
        "SELECT DISTINCT ticker FROM quarterly_standalone").fetchall()]
    for t in tickers:
        ttm_op = cur.execute(
            "SELECT SUM(operating_profit) FROM (SELECT operating_profit FROM quarterly_standalone "
            "WHERE ticker=? AND is_valid=1 ORDER BY period_end DESC LIMIT 4)", (t,)).fetchone()[0]
        if ttm_op is None:
            continue
        # 市場の要求水準：TTM営業利益の0.6〜1.5倍（<1なら割安＝織り込み不足）
        factor = 0.6 + h(t + "rdcf") * 0.9
        required = max(ttm_op * factor, ttm_op * 0.3)  # 赤字企業でも下限あり
        cur.execute(
            "INSERT OR REPLACE INTO reverse_dcf_requirements "
            "(ticker, as_of_date, required_steady_op_profit, target_year, source_run_id) "
            "VALUES (?,?,?,?,?)",
            (t, AS_OF, round(required, 1), 2029, "mock_run_001"))
    conn.commit()
    print("reverse_dcf_requirements:", cur.execute("SELECT COUNT(*) FROM reverse_dcf_requirements").fetchone()[0], "社")
    conn.close()


if __name__ == "__main__":
    main()
