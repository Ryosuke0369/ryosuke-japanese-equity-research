"""
generate_mock_data.py — モックデータ生成（実データ接続前の骨格検証用）
- ユニバース：テスト銘柄4社＋任意2社＋合成300社
- 開示履歴：各社3年分の四半期決算発表日（TDnetアーカイブの代替）
- 3905は資料§6の実績日付（2026-08-14に1Q発表）を正として埋め込む
"""
import random
import sys
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.jp_calendar import snap_to_business_day, is_business_day
from common.schema import init_db

TODAY = date(2026, 8, 31)  # システム基準日
rng = random.Random(42)

# ---------- テスト銘柄（引継ぎ資料§6） ----------
TEST_TICKERS = {
    "3905": {"name": "データセクション", "market": "growth",   "fy_end": 3},
    "3441": {"name": "山王",             "market": "standard", "fy_end": 3},
    "278A": {"name": "防衛ドローン銘柄",  "market": "growth",   "fy_end": 12},
    "5726": {"name": "半導体銘柄",       "market": "standard", "fy_end": 3},
    "6501": {"name": "一般銘柄A",        "market": "standard", "fy_end": 3},
    "7203": {"name": "一般銘柄B",        "market": "standard", "fy_end": 3},
}

# 3905の実績発表日（§6：2027年3月期1Qを2026-08-14に発表）
FIXED_FILINGS_3905 = [
    ("2024-05-15", "2024-03-31", "FY", 2024),
    ("2024-08-14", "2024-06-30", "1Q", 2025),
    ("2024-11-14", "2024-09-30", "2Q", 2025),
    ("2025-02-14", "2024-12-31", "3Q", 2025),
    ("2025-05-15", "2025-03-31", "FY", 2025),
    ("2025-08-14", "2025-06-30", "1Q", 2026),
    ("2025-11-14", "2025-09-30", "2Q", 2026),
    ("2026-02-13", "2025-12-31", "3Q", 2026),
    ("2026-05-15", "2026-03-31", "FY", 2026),
    ("2026-08-14", "2026-06-30", "1Q", 2027),  # ← 直近（§6の数値の発表日）
]


def quarter_period_end(fy_end_month: int, fiscal_year: int, q: int) -> date:
    """決算期末月・年度・四半期番号(1-4)から期末日を算出。q=4は通期"""
    m = (fy_end_month + 3 * q - 1) % 12 + 1
    y = fiscal_year - 1 + (1 if fy_end_month + 3 * q > 12 else 0)
    if m == 2:
        day = 29 if (y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)) else 28
    elif m in (4, 6, 9, 11):
        day = 30
    else:
        day = 31
    return date(y, m, day)


def gen_announce_date(period_end: date, jitter: int = 4) -> date:
    """期末+約45日±jitter → 営業日に寄せる"""
    base = period_end + timedelta(days=45 + rng.randint(-jitter, jitter))
    return snap_to_business_day(base)


def gen_universe(n_synthetic: int = 300):
    rows = []
    for t, info in TEST_TICKERS.items():
        rows.append((t, info["name"], info["market"], info["fy_end"]))
    # 合成銘柄：7月決算を少数混入（デモ用にT-15内候補を作る）
    fy_choices = [3]*60 + [12]*12 + [9]*6 + [2]*4 + [6]*3 + [7]*8 + [1, 4, 5, 8, 10, 11]
    for i in range(n_synthetic):
        t = str(1400 + i * 7)[:4]
        if t in TEST_TICKERS:
            t = str(int(t) + 1)
        fy = 7 if i % 25 == 0 else rng.choice(fy_choices)
        rows.append((t, f"合成銘柄{i:03d}", rng.choice(["growth", "standard"]), fy))
    return rows


def gen_filings(universe_rows):
    """2024〜2026年度の完了済み発表＋進行中年度の既発表分を生成"""
    records = []
    for t, name, market, fy_end in universe_rows:
        if t == "3905":
            records.extend([(t,) + r for r in FIXED_FILINGS_3905])
            continue
        for fy in (2024, 2025, 2026):
            for q in (1, 2, 3, 4):
                pe = quarter_period_end(fy_end, fy, q)
                ad = gen_announce_date(pe)
                if ad > TODAY:
                    continue  # 未来の発表は履歴に入れない
                qt = {1: "1Q", 2: "2Q", 3: "3Q", 4: "FY"}[q]
                records.append((t, ad.isoformat(), pe.isoformat(), qt, fy))
        # 2027年度の進行中分（発表済みのみ）
        for q in (1, 2, 3, 4):
            pe = quarter_period_end(fy_end, 2027, q)
            ad = gen_announce_date(pe)
            if ad <= TODAY:
                qt = {1: "1Q", 2: "2Q", 3: "3Q", 4: "FY"}[q]
                records.append((t, ad.isoformat(), pe.isoformat(), qt, 2027))
    return records


def main():
    db_path = Path(__file__).resolve().parents[1] / "data" / "screener.db"
    conn = init_db(db_path)
    cur = conn.cursor()

    universe4 = gen_universe()
    filings = gen_filings(universe4)
    sectors = ["情報・通信", "製造業", "サービス", "小売", "建設", "医薬品"]
    universe = [(t, n, m, sectors[int(t[0], 36) % 6], fy) for (t, n, m, fy) in universe4]
    cur.executemany(
        "INSERT OR REPLACE INTO universe (ticker, company_name, market, sector, fiscal_year_end) VALUES (?,?,?,?,?)",
        universe,
    )
    cur.executemany(
        "INSERT OR REPLACE INTO filings (ticker, filing_date, period_end, quarter_type, fiscal_year, source) "
        "VALUES (?,?,?,?,?, 'mock')",
        filings,
    )
    conn.commit()
    n_u = cur.execute("SELECT COUNT(*) FROM universe").fetchone()[0]
    n_f = cur.execute("SELECT COUNT(*) FROM filings").fetchone()[0]
    print(f"universe: {n_u}社 / filings: {n_f}件")
    for t in ("3905", "278A"):
        for row in cur.execute(
            "SELECT filing_date, quarter_type, fiscal_year FROM filings WHERE ticker=? ORDER BY filing_date", (t,)
        ):
            print(t, row)
    conn.close()


if __name__ == "__main__":
    main()
