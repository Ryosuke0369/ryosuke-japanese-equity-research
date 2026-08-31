"""
weekly_batch.py — Module A: 週次更新バッチ（毎週月曜09:00実行想定）
1) 決算カレンダーDBを再構築
2) T-15営業日以内に決算を控える銘柄を抽出してCSV出力（Module Bへの入力）
"""
import csv
import sqlite3
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.jp_calendar import business_days_between
from module_a.earnings_calendar import build_calendar

T_MINUS = 15  # エントリー閾値（営業日）


def extract_candidates(db_path: str | Path, as_of: date, t_minus: int = T_MINUS) -> list[dict]:
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        "SELECT c.*, u.company_name, u.market FROM earnings_calendar c "
        "JOIN universe u ON u.ticker = c.ticker"
    ).fetchall()
    conn.close()

    cands = []
    for r in rows:
        d = date.fromisoformat(r["next_earnings_date"])
        bd = business_days_between(as_of, d)
        if 0 < bd <= t_minus:
            cands.append({
                "ticker": r["ticker"],
                "company_name": r["company_name"],
                "market": r["market"],
                "next_earnings_date": r["next_earnings_date"],
                "quarter_type": r["quarter_type"],
                "fiscal_year": r["fiscal_year"],
                "business_days_to_earnings": bd,
                "confidence_level": r["confidence_level"],
            })
    cands.sort(key=lambda x: (x["business_days_to_earnings"], x["ticker"]))
    return cands


def run_batch(as_of: date, db_path: str | Path, out_dir: str | Path) -> Path:
    build_calendar(db_path, as_of)
    cands = extract_candidates(db_path, as_of)

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = out_dir / f"candidates_{as_of.isoformat()}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(cands[0].keys()) if cands else
                           ["ticker", "company_name", "market", "next_earnings_date",
                            "quarter_type", "fiscal_year", "business_days_to_earnings",
                            "confidence_level"])
        w.writeheader()
        w.writerows(cands)
    return csv_path


if __name__ == "__main__":
    base = Path(__file__).resolve().parents[1]
    as_of = date(2026, 8, 31)  # 実運用では date.today()
    csv_path = run_batch(as_of, base / "data" / "screener.db", base / "data")
    print(f"batch done -> {csv_path}")
