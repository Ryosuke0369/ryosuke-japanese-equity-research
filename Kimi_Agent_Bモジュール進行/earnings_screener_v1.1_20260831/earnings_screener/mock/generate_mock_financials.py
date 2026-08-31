"""
generate_mock_financials.py — B-1: 財務モックデータ生成
対象：テスト銘柄6社＋T-15候補（Module A出力）の全社
- quarterly_standalone / balance_sheet_items / pl_adjustments / company_forecasts / evidence_events
- 単位は全て百万円。3905は引継ぎ資料§6の実数値を再現
"""
import csv
import hashlib
import random
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

BASE = Path(__file__).resolve().parents[1]
DB = BASE / "data" / "screener.db"
DAYS_PER_Q = 91


def h(seed: str) -> float:
    return int(hashlib.md5(seed.encode()).hexdigest(), 16) % 10000 / 10000.0


def periods_of(conn, ticker):
    return conn.execute(
        "SELECT period_end, quarter_type, fiscal_year FROM filings WHERE ticker=? ORDER BY period_end",
        (ticker,),
    ).fetchall()


# ========== テスト銘柄の手組みデータ ==========
def craft_3905(conn):
    """データセクション：§6再現。一時収入5580をpl_adjustmentsで分離"""
    rows = [
        # period_end, q, fy, sales, op, gp, cf(H1/H2のみ)
        ("2024-06-30", "1Q", 2025, 500, -400, 200, None),
        ("2024-09-30", "2Q", 2025, 550, -350, 220, -700),
        ("2024-12-31", "3Q", 2025, 600, -300, 240, None),
        ("2025-03-31", "FY", 2025, 650, -250, 260, -300),
        ("2025-06-30", "1Q", 2026, 668, -340, 267, None),
        ("2025-09-30", "2Q", 2026, 1200, -100, 480, -500),
        ("2025-12-31", "3Q", 2026, 1400, 100, 560, None),
        ("2026-03-31", "FY", 2026, 1500, 200, 600, 800),
        ("2026-06-30", "1Q", 2027, 6330, 4840, 5880, None),  # 報告値（一時収入込み）
    ]
    bs = {
        "accounts_receivable":      [800, 900, 1000, 1100, 4000, 7000, 9500, 11180, 8790],
        "construction_in_progress": [100, 200, 500, 800, 1500, 2500, 3500, 4490, 6710],
        "machinery":                [200, 200, 200, 200, 200, 250, 280, 300, 350],
        "deposits_received":        [300, 400, 800, 1500, 4000, 4800, 5000, 5230, 2190],
        "advances_paid":            [100, 150, 300, 600, 1200, 1800, 2100, 2380, 4240],
        "contract_liabilities":     [80, 90, 100, 110, 120, 130, 140, 150, 160],
        "inventory":                [50] * 9,
    }
    return rows, bs

ADJ_3905 = [("3905", "2026-06-30", "one_time_revenue", 5580,
             "追加受注案件解約の対価（手数料収入）— 継続収入として扱わない")]
FC_3905 = [("3905", 2027, 162190, 24820, "2026-05-15")]


def craft_3441(conn):
    """山王：累計GM改善 vs 単独GM低下の再現（FY2026: 32→31→29→27%、FY2025は25%平坦）"""
    rows = [
        ("2024-06-30", "1Q", 2025, 1800, 90, 450, None),
        ("2024-09-30", "2Q", 2025, 1850, 95, 463, 200),
        ("2024-12-31", "3Q", 2025, 1900, 100, 475, None),
        ("2025-03-31", "FY", 2025, 1950, 110, 488, 260),
        ("2025-06-30", "1Q", 2026, 2000, 160, 640, None),
        ("2025-09-30", "2Q", 2026, 2100, 170, 651, 350),
        ("2025-12-31", "3Q", 2026, 2200, 150, 638, None),
        ("2026-03-31", "FY", 2026, 2300, 140, 621, 320),
        ("2026-06-30", "1Q", 2027, 2400, 130, 624, None),  # GM 26.0%（単独低下継続）
    ]
    ar = [round(s * 55 / DAYS_PER_Q) for (_, _, _, s, _, _, _) in rows]
    cl = [200] * 9
    bs = {"accounts_receivable": ar, "contract_liabilities": cl, "inventory": [600] * 9}
    return rows, bs

FC_3441 = [("3441", 2027, 9800, 700, "2026-05-12")]


def craft_278a(conn):
    """防衛ドローン銘柄（12月決算）：MOU（約束）と装備庁受注（証拠）の分離"""
    rows = [
        ("2024-03-31", "1Q", 2024, 280, -60, 112, None),
        ("2024-06-30", "2Q", 2024, 300, -50, 120, -100),
        ("2024-09-30", "3Q", 2024, 310, -40, 124, None),
        ("2024-12-31", "FY", 2024, 330, -30, 132, -80),
        ("2025-03-31", "1Q", 2025, 320, -45, 128, None),
        ("2025-06-30", "2Q", 2025, 350, -30, 140, -60),
        ("2025-09-30", "3Q", 2025, 380, -10, 152, None),
        ("2025-12-31", "FY", 2025, 420, 10, 168, 40),
        ("2026-03-31", "1Q", 2026, 460, 30, 184, None),
        ("2026-06-30", "2Q", 2026, 520, 55, 208, 120),
    ]
    ar = [round(s * 60 / DAYS_PER_Q) for (_, _, _, s, _, _, _) in rows]
    cl = [20, 25, 25, 30, 30, 35, 40, 45, 180, 380]  # 受注に伴う前受金が急増（証拠）
    bs = {"accounts_receivable": ar, "contract_liabilities": cl, "inventory": [150] * 10}
    return rows, bs

FC_278A = [("278A", 2026, 1800, 80, "2026-02-10")]
EVENTS_278A = [
    ("278A", "2026-03-15", "mou", None, 0, "mock", "海外メーカーとMOU締結（約束→スコア対象外）"),
    ("278A", "2026-05-20", "order", 450, 1, "mock", "防衛装備庁 初受注（履行義務のある実契約）"),
]


def craft_5726(conn):
    """半導体銘柄：在庫両義分解用（S6で本領）。数量主導の原価増×売上加速パターン"""
    base_rows = [
        ("2024-06-30", "1Q", 2025, 3000, 300, 900, None),
        ("2024-09-30", "2Q", 2025, 3100, 320, 930, 700),
        ("2024-12-31", "3Q", 2025, 3050, 290, 915, None),
        ("2025-03-31", "FY", 2025, 3150, 310, 945, 680),
        ("2025-06-30", "1Q", 2026, 3200, 330, 960, None),
        ("2025-09-30", "2Q", 2026, 3400, 370, 1020, 780),
        ("2025-12-31", "3Q", 2026, 3700, 430, 1110, None),
        ("2026-03-31", "FY", 2026, 4100, 500, 1230, 1000),
        ("2026-06-30", "1Q", 2027, 4600, 580, 1380, None),
    ]
    rows = []
    for i, (pe, qt, fy, s, op, gp, cf) in enumerate(base_rows):
        cogs = s - gp
        cq = cogs * 0.70 * (1 + 0.04 * i)   # 数量要因：四半期ごとに+4%（数量主導）
        cp = cogs - cq                       # 価格要因：残差（ほぼ横ばい〜逓減）
        rows.append((pe, qt, fy, s, op, gp, round(cq, 1), round(cp, 1), cf))
    ar = [round(r[3] * 70 / DAYS_PER_Q) for r in rows]
    inv = [900, 920, 950, 980, 1050, 1150, 1300, 1480, 1700]  # 在庫増（仕込みか滞留か）
    bs = {"accounts_receivable": ar, "inventory": inv, "contract_liabilities": [100] * 9}
    return rows, bs

FC_5726 = [("5726", 2027, 18000, 2300, "2026-05-14")]


def craft_generic(ticker, conn):
    """一般銘柄A/B：ごく普通のパターン"""
    periods = periods_of(conn, ticker)
    rng = random.Random(ticker)
    base = rng.uniform(1500, 2500)
    rows, ar_l, cl_l, inv_l = [], [], [], []
    for i, (pe, qt, fy) in enumerate(periods):
        s = base * (1 + 0.02 * i) * (1 + 0.05 * (1 if qt in ("3Q", "FY") else -1))
        gp = s * 0.30
        op = gp - s * 0.24
        cf = round(op * 2 * 0.9, 1) if qt in ("2Q", "FY") else None
        rows.append((pe, qt, fy, round(s, 1), round(op, 1), round(gp, 1), None, None, cf))
        ar_l.append(round(s * 60 / DAYS_PER_Q))
        cl_l.append(round(s * 0.05))
        inv_l.append(round(s * 0.25))
    bs = {"accounts_receivable": ar_l, "contract_liabilities": cl_l, "inventory": inv_l}
    fc = [(ticker, 2027, round(base * 4.2, 1), round(base * 0.28, 1), "2026-05-15")]
    return rows, bs, fc


# ========== 合成銘柄（パターン割当） ==========
def synth_ticker(ticker, conn):
    periods = periods_of(conn, ticker)
    if not periods:
        return None
    dso_pat = ["improving", "flat", "worsening"][int(h(ticker + "dso") * 3)]
    cl_pat = ["growing", "flat", "declining"][int(h(ticker + "cl") * 3)]
    cip_pat = ["transfer", "building", "none", "none", "none"][int(h(ticker + "cip") * 5)]
    cf_q = [1.2, 0.8, 0.3][int(h(ticker + "cf") * 3)]
    base = 500 + h(ticker + "base") * 4500
    growth = -0.05 + h(ticker + "g") * 0.20
    gm = 0.20 + h(ticker + "gm") * 0.20
    dso0 = 45 + h(ticker + "dso0") * 45
    cl0 = 0.02 + h(ticker + "cl0") * 0.13

    rows = []
    bs = {"accounts_receivable": [], "contract_liabilities": [], "inventory": [],
          "construction_in_progress": [], "machinery": []}
    n = len(periods)
    for i, (pe, qt, fy) in enumerate(periods):
        yr = i / 4
        s = base * (1 + growth) ** yr * (1 + 0.08 * (1 if qt in ("3Q", "FY") else -1))
        gp = s * gm
        op = gp - s * (gm - 0.03 - h(ticker + "opm") * 0.07)
        cf = round(op * 2 * cf_q, 1) if qt in ("2Q", "FY") else None
        cogs_q = (s - gp) * 0.7
        cogs_p = (s - gp) * 0.3
        rows.append((pe, qt, fy, round(s, 1), round(op, 1), round(gp, 1),
                     round(cogs_q, 1), round(cogs_p, 1), cf))
        dso_trend = {"improving": -0.08, "flat": 0.0, "worsening": 0.08}[dso_pat]
        dso = dso0 * (1 + dso_trend) ** yr
        bs["accounts_receivable"].append(round(s * dso / DAYS_PER_Q, 1))
        cl_g = {"growing": 0.30, "flat": 0.0, "declining": -0.15}[cl_pat]
        bs["contract_liabilities"].append(round(s * cl0 * (1 + cl_g) ** yr, 1))
        bs["inventory"].append(round(s * 0.28 * (1 + 0.05) ** yr, 1))
        if cip_pat == "none":
            cip, mach = 0.0, 500.0
        elif cip_pat == "building":
            cip = 50 + (800 - 50) * (i / max(n - 1, 1))
            mach = 500.0
        else:  # transfer：最新期でCIP 800→150、機械装置+650
            cip = 800.0 if i < n - 1 else 150.0
            mach = 500.0 if i < n - 1 else 1150.0
        bs["construction_in_progress"].append(round(cip, 1))
        bs["machinery"].append(round(mach, 1))
    ttm = sum(r[3] for r in rows[-4:])
    fc = [(ticker, 2027, round(ttm * (1 + growth / 2), 1), round(ttm * 0.06, 1), "2026-05-15")]
    return rows, bs, fc


def insert(conn, ticker, rows, bs):
    cur = conn.cursor()
    for row in rows:
        pe, qt, fy, s, op, gp = row[:6]
        rest = list(row[6:])
        if len(rest) == 1:      # 7要素形式: (…, gp, cf)
            cq, cp, cf = None, None, rest[0]
        else:                   # 9要素形式: (…, gp, cq, cp, cf)
            cq, cp, cf = (rest + [None, None, None])[:3]
        cur.execute(
            "INSERT OR REPLACE INTO quarterly_standalone "
            "(ticker, fiscal_year, quarter_type, period_end, sales, operating_profit, "
            "gross_profit, cogs_quantity, cogs_price, operating_cf) VALUES (?,?,?,?,?,?,?,?,?,?)",
            (ticker, fy, qt, pe, s, op, gp, cq, cp, cf),
        )
    for key, vals in bs.items():
        for row, v in zip(rows, vals):
            cur.execute(
                "INSERT OR REPLACE INTO balance_sheet_items (ticker, period_end, item_key, value) "
                "VALUES (?,?,?,?)", (ticker, row[0], key, v),
            )
    conn.commit()


def main():
    conn = sqlite3.connect(DB)
    cur = conn.cursor()

    # テスト銘柄（手組み）
    for t, crafter in [("3905", craft_3905), ("3441", craft_3441),
                       ("278A", craft_278a), ("5726", craft_5726)]:
        rows, bs = crafter(conn)
        insert(conn, t, rows, bs)
    for t in ("6501", "7203"):
        rows, bs, fc = craft_generic(t, conn)
        insert(conn, t, rows, bs)
        cur.executemany(
            "INSERT OR REPLACE INTO company_forecasts (ticker, fiscal_year, forecast_sales, forecast_op, source_date) VALUES (?,?,?,?,?)", fc)

    # --- 訂正短信デモ：3905の3Q FY2026（2025-12-31期）を generation=2 で訂正 ---
    # 初回: 2026-02-13発表 sales=1400 → 訂正: 2026-03-10発表 sales=1350（遡及修正）
    cur.execute(
        "INSERT OR REPLACE INTO filings (ticker, filing_date, period_end, quarter_type, fiscal_year, source, generation) "
        "VALUES ('3905','2026-03-10','2025-12-31','3Q',2026,'mock',2)")
    cur.execute(
        "INSERT OR REPLACE INTO quarterly_standalone "
        "(ticker, fiscal_year, quarter_type, period_end, sales, operating_profit, gross_profit, generation) "
        "VALUES ('3905',2026,'3Q','2025-12-31',1350,60,540,2)")
    cur.execute(
        "INSERT OR REPLACE INTO balance_sheet_items (ticker, period_end, item_key, value, generation) "
        "VALUES ('3905','2025-12-31','accounts_receivable',9600,2)")

    cur.executemany(
        "INSERT OR REPLACE INTO pl_adjustments (ticker, period_end, item_key, amount, note) VALUES (?,?,?,?,?)", ADJ_3905)
    for fc in (FC_3905, FC_3441, FC_278A, FC_5726):
        cur.executemany(
            "INSERT OR REPLACE INTO company_forecasts (ticker, fiscal_year, forecast_sales, forecast_op, source_date) VALUES (?,?,?,?,?)", fc)
    cur.executemany(
        "INSERT OR REPLACE INTO evidence_events (ticker, event_date, event_type, amount, evidence_flag, source_doc, note) VALUES (?,?,?,?,?,?,?)", EVENTS_278A)
    conn.commit()

    # ユニバース全社の合成銘柄（バックテスト用に全社分を生成）
    all_tickers = [r[0] for r in cur.execute("SELECT ticker FROM universe").fetchall()]
    n_syn = 0
    for t in all_tickers:
        if t in ("3905", "3441", "278A", "5726", "6501", "7203"):
            continue
        res = synth_ticker(t, conn)
        if res:
            rows, bs, fc = res
            insert(conn, t, rows, bs)
            cur.executemany(
                "INSERT OR REPLACE INTO company_forecasts (ticker, fiscal_year, forecast_sales, forecast_op, source_date) VALUES (?,?,?,?,?)", fc)
            n_syn += 1
    conn.commit()
    print(f"financials generated: テスト6社 + 合成{n_syn}社")
    print("quarterly_standalone:", cur.execute("SELECT COUNT(*) FROM quarterly_standalone").fetchone()[0], "行")
    print("balance_sheet_items:", cur.execute("SELECT COUNT(*) FROM balance_sheet_items").fetchone()[0], "行")
    conn.close()


if __name__ == "__main__":
    main()
