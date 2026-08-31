"""
generate_mock_prices.py — バックテスト用モック株価・市場指数
- 2024-07〜2026-09の営業日分の日次株価（幾何ブラウン運動＋決算イベントジャンプ）
- イベントジャンプの期待値は財務モックのパターン品質と相関させる
  （証拠スコアに予見力がある世界を模擬し、バックテストがエッジを計測できるか検証する）
- 発表は引け後とし、ジャンプは翌営業日に適用
"""
import hashlib
import math
import random
import sqlite3
import sys
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.jp_calendar import is_business_day, add_business_days

DB = Path(__file__).resolve().parents[1] / "data" / "screener.db"
START, END = date(2024, 7, 1), date(2026, 9, 30)


def h(seed: str) -> float:
    return int(hashlib.md5(seed.encode()).hexdigest(), 16) % 10000 / 10000.0


def pattern_quality(ticker: str) -> float:
    """generate_mock_financials と同じシードからパターン品質(0-1)を再構成"""
    dso_good = 1.0 if int(h(ticker + "dso") * 3) == 0 else 0.0   # improving
    cl_good = 1.0 if int(h(ticker + "cl") * 3) == 0 else 0.0     # growing
    cf_good = [1.2, 0.8, 0.3][int(h(ticker + "cf") * 3)] / 1.2
    return (dso_good + cl_good + cf_good) / 3


def business_days(start: date, end: date):
    d, out = start, []
    while d <= end:
        if is_business_day(d):
            out.append(d)
        d += timedelta(days=1)
    return out


def main():
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    days = business_days(START, END)
    rng = random.Random(7)

    # --- 市場指数（日経平均の代替） ---
    idx = 38000.0
    idx_rets = []
    for i, d in enumerate(days):
        idx *= 1 + 0.00025 + rng.gauss(0, 0.011)
        cur.execute("INSERT OR REPLACE INTO market_index (date, close) VALUES (?,?)",
                    (d.isoformat(), round(idx, 2)))
        idx_rets.append(idx)
    conn.commit()

    # --- 銘柄別日次株価 ---
    tickers = [r[0] for r in cur.execute("SELECT ticker FROM universe").fetchall()]
    filings = {}
    for t, fd in cur.execute(
            "SELECT ticker, filing_date FROM filings WHERE generation=1").fetchall():
        filings.setdefault(t, set()).add(fd)

    for t in tickers:
        rng_t = random.Random(t)
        beta = 0.8 + h(t + "beta") * 0.5
        ivol = 0.015 + h(t + "ivol") * 0.010
        quality = pattern_quality(t)
        price = 500 + h(t + "p0") * 2500
        base_vol = 50_000 + h(t + "vol") * 450_000

        event_days = {}
        for fd in filings.get(t, set()):
            d_fd = date.fromisoformat(fd)
            if START <= d_fd <= END:
                nxt = add_business_days(d_fd, 1)  # 引け後発表→翌営業日に反応
                jump = (-0.06 + 0.14 * quality) + rng_t.gauss(0, 0.03)
                event_days[nxt.isoformat()] = jump

        prev_idx = None
        rows = []
        for i, d in enumerate(days):
            mkt_ret = (idx_rets[i] / idx_rets[i - 1] - 1) if i > 0 else 0.0
            ret = beta * mkt_ret + rng_t.gauss(0, ivol)
            ds = d.isoformat()
            vol_mult = 1.0
            if ds in event_days:
                ret += event_days[ds]
                vol_mult = 3.0
            price *= 1 + ret
            rows.append((t, ds, round(max(price, 10), 1),
                         round(base_vol * vol_mult * (0.7 + h(t + ds) * 0.6))))
        cur.executemany(
            "INSERT OR REPLACE INTO daily_prices (ticker, date, close, volume) VALUES (?,?,?,?)", rows)
    conn.commit()
    print("daily_prices:", cur.execute("SELECT COUNT(*) FROM daily_prices").fetchone()[0], "行 /",
          cur.execute("SELECT COUNT(DISTINCT ticker) FROM daily_prices").fetchone()[0], "社")
    conn.close()


if __name__ == "__main__":
    main()
