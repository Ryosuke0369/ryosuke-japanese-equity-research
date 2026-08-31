"""
earnings_calendar.py — Module A: 決算カレンダーDB構築
過去の決算発表日から「前年同日±3営業日」で次回決算予定日を推定する。

設計上の前提（引継ぎ資料§4-A）：
- 推定は「同じ四半期種別の直近の発表日 + 1年」を基本形とする
- 四半期種別は期末月ベースで自動判定済み（filings.quarter_type）のものを利用
- confidence_level: 過去の同四半期発表日のばらつき（営業日換算）で HIGH/MEDIUM/LOW
"""
import sqlite3
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.jp_calendar import snap_to_business_day, business_days_between

Q_ORDER = ["1Q", "2Q", "3Q", "FY"]


def next_quarter(q_type: str, fiscal_year: int) -> tuple[str, int]:
    """次の四半期種別と年度。FY(通期)発表の次は翌年度の1Q"""
    i = Q_ORDER.index(q_type)
    if i == 3:
        return "1Q", fiscal_year + 1
    return Q_ORDER[i + 1], fiscal_year


def estimate_next_date(history_same_q: list[str]) -> tuple[str, str]:
    """同四半期の過去発表日リスト（昇順）から次回を推定。
    戻り値: (推定日ISO, 根拠日ISO)。直近年 + 1年 → 営業日寄せ。"""
    base = date.fromisoformat(history_same_q[-1])
    try:
        cand = base.replace(year=base.year + 1)
    except ValueError:  # 2/29
        cand = base.replace(year=base.year + 1, day=28)
    return snap_to_business_day(cand).isoformat(), base.isoformat()


def calc_confidence(history_same_q: list[str]) -> str:
    """同四半期発表日の年またぎばらつき（営業日差）で判定"""
    if len(history_same_q) < 2:
        return "LOW"
    gaps = []
    for a, b in zip(history_same_q[:-1], history_same_q[1:]):
        da, db = date.fromisoformat(a), date.fromisoformat(b)
        # 1年ずらして比較（曜日ずれを吸収するため営業日差で評価）
        try:
            da_shift = da.replace(year=da.year + 1)
        except ValueError:
            da_shift = da.replace(year=da.year + 1, day=28)
        gaps.append(abs(business_days_between(da_shift, db)))
    spread = max(gaps)
    if len(history_same_q) >= 3 and spread <= 5:
        return "HIGH"
    if spread <= 10:
        return "MEDIUM"
    return "LOW"


def build_calendar(db_path: str | Path, as_of: date) -> list[dict]:
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("DELETE FROM earnings_calendar")

    results = []
    tickers = cur.execute(
        "SELECT ticker, fiscal_year_end FROM universe WHERE market_cap_ok=1 AND liquidity_ok=1"
    ).fetchall()

    for row in tickers:
        t, fy_end = row["ticker"], row["fiscal_year_end"]
        filings = cur.execute(
            "SELECT filing_date, quarter_type, fiscal_year FROM filings "
            "WHERE ticker=? AND filing_date<=? AND generation=1 ORDER BY filing_date",
            (t, as_of.isoformat()),  # 初回発表のみ使用（訂正短信は発表日カデンスをずらさない）
        ).fetchall()
        if not filings:
            continue

        # 直近発表の次の四半期から開始し、推定日が未来になるまで繰り上げ
        last = filings[-1]
        q_next, fy_next = next_quarter(last["quarter_type"], last["fiscal_year"])

        for _ in range(5):  # 最大5四半期先まで繰り上げ
            same_q = [f["filing_date"] for f in filings if f["quarter_type"] == q_next]
            if not same_q:
                break
            est, est_from = estimate_next_date(same_q)
            if date.fromisoformat(est) > as_of:
                conf = calc_confidence(same_q)
                results.append({
                    "ticker": t,
                    "next_earnings_date": est,
                    "quarter_type": q_next,
                    "fiscal_year": fy_next,
                    "fiscal_year_end": fy_end,
                    "confidence_level": conf,
                    "estimated_from": est_from,
                    "updated_at": as_of.isoformat(),
                })
                break
            q_next, fy_next = next_quarter(q_next, fy_next)

    cur.executemany(
        "INSERT INTO earnings_calendar VALUES (:ticker, :next_earnings_date, :quarter_type, "
        ":fiscal_year, :fiscal_year_end, :confidence_level, :estimated_from, :updated_at)",
        results,
    )
    conn.commit()
    conn.close()
    return results


if __name__ == "__main__":
    db = Path(__file__).resolve().parents[1] / "data" / "screener.db"
    res = build_calendar(db, date(2026, 8, 31))
    print(f"calendar built: {len(res)}社")
