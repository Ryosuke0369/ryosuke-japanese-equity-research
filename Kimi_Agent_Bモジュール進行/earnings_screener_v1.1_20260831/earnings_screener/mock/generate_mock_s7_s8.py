"""
generate_mock_s7_s8.py — S7/S8用の補助モックデータ
- S7: 収益認識内訳（一時点移転 vs 一定期間）※四半期注記の代替として汎用itemストアに格納
- S8: 株式数（発行済・潜在株式調整後）による希薄化オーバーハング
"""
import hashlib
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
DB = Path(__file__).resolve().parents[1] / "data" / "screener.db"


def h(seed: str) -> float:
    return int(hashlib.md5(seed.encode()).hexdigest(), 16) % 10000 / 10000.0


def put(conn, ticker, period_end, key, value):
    conn.execute(
        "INSERT OR REPLACE INTO balance_sheet_items (ticker, period_end, item_key, value) "
        "VALUES (?,?,?,?)", (ticker, period_end, key, value))


def periods(conn, ticker):
    return [r[0] for r in conn.execute(
        "SELECT period_end FROM quarterly_standalone WHERE ticker=? ORDER BY period_end",
        (ticker,)).fetchall()]


def main():
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    tickers = [r[0] for r in cur.execute(
        "SELECT DISTINCT ticker FROM quarterly_standalone").fetchall()]

    for t in tickers:
        ps = periods(conn, t)
        if not ps:
            continue
        sales = {r[0]: r[1] for r in cur.execute(
            "SELECT period_end, sales FROM quarterly_standalone WHERE ticker=?", (t,)).fetchall()}

        # --- S8: 株式数（デフォルトは希薄化なし。一部銘柄にCBオーバーハング） ---
        out0 = 10_000_000 + int(h(t + "sh") * 40_000_000)
        cb = h(t + "cb")
        for i, pe in enumerate(ps):
            outstanding = out0
            if t == "278A":  # 成長資金のCB発行で希薄化が進行するケース
                schedule = [0.05] * max(0, len(ps) - 2) + [0.09, 0.13]
                overhang = schedule[i] if i < len(schedule) else 0.13
                diluted = int(outstanding * (1 + overhang))
            elif cb > 0.75:  # 約25%の銘柄に潜在株式
                overhang = 0.04 + (cb - 0.75) * 0.6  # 4〜19%
                diluted = int(outstanding * (1 + overhang))
            else:
                diluted = outstanding
            put(conn, t, pe, "shares_outstanding", outstanding)
            put(conn, t, pe, "diluted_shares", diluted)

        # --- S7: 収益認識内訳（開示がある銘柄のみ。全体の4割に付与） ---
        if h(t + "s7") > 0.6 or t in ("278A", "6501"):
            for i, pe in enumerate(ps):
                s = sales.get(pe) or 0
                if t == "278A":
                    ot_share = min(0.10 + 0.06 * i, 0.60)  # 一定期間移転の比率が上昇（長期契約化）
                elif t == "6501":
                    ot_share = 0.30
                else:
                    ot_share = 0.20 + (h(t + "s7t") - 0.5) * 0.1 * i
                    ot_share = max(0.0, min(0.8, ot_share))
                put(conn, t, pe, "rev_over_time", round(s * ot_share, 1))
                put(conn, t, pe, "rev_point_in_time", round(s * (1 - ot_share), 1))

    conn.commit()
    n7 = cur.execute("SELECT COUNT(DISTINCT ticker) FROM balance_sheet_items WHERE item_key='rev_over_time'").fetchone()[0]
    n8 = cur.execute("SELECT COUNT(DISTINCT ticker) FROM balance_sheet_items WHERE item_key='diluted_shares'").fetchone()[0]
    print(f"S7収益認識内訳: {n7}社 / S8株式数: {n8}社")
    # 278Aの確認
    for r in cur.execute(
        "SELECT period_end, item_key, value FROM balance_sheet_items "
        "WHERE ticker='278A' AND item_key IN ('rev_over_time','diluted_shares','shares_outstanding') "
        "ORDER BY period_end, item_key").fetchall()[-6:]:
        print("278A", r)
    conn.close()


if __name__ == "__main__":
    main()
