"""
data_access.py — Module B 共通データアクセス層（B-1接続層）
実データ接続時はこの層だけ差し替える（XBRLパーサー／単独値ビルダーの出力を読む）。

PIT（Point-in-Time）規則:
1. 判定基準は period_end ではなく filings.filing_date（期末の値は発表日に初めて公知になる）
2. 訂正世代: 同一 period_end に複数 generation がある場合、as_of 時点で
   「発表済みの最新世代」のみを返す。過去日を指定すれば訂正前の値が再現される
3. スコアラー・予測器はDBを直接叩かず、必ずこの層の関数を経由する
"""
import sqlite3
from pathlib import Path

DB = Path(__file__).resolve().parents[1] / "data" / "screener.db"
DAYS_PER_Q = 91


def connect(db_path=None):
    conn = sqlite3.connect(str(db_path or DB))
    conn.row_factory = sqlite3.Row
    return conn


def _iso(as_of):
    return as_of.isoformat() if hasattr(as_of, "isoformat") else as_of


def visible_generations(conn, ticker, as_of=None) -> dict:
    """period_end → as_of時点で見える最新generation のマップを返す。
    as_of=None のときは全履歴の最新世代（=通常運用）。"""
    if as_of is None:
        rows = conn.execute(
            "SELECT period_end, MAX(generation) g FROM filings "
            "WHERE ticker=? GROUP BY period_end", (ticker,)).fetchall()
    else:
        rows = conn.execute(
            "SELECT period_end, MAX(generation) g FROM filings "
            "WHERE ticker=? AND filing_date<=? GROUP BY period_end",
            (ticker, _iso(as_of))).fetchall()
    return {r["period_end"]: r["g"] for r in rows}


def get_pl_series(conn, ticker, as_of=None):
    """単独値PL時系列（is_valid=1・可視世代のみ）。古い順。"""
    vis = visible_generations(conn, ticker, as_of)
    rows = conn.execute(
        "SELECT period_end, quarter_type, fiscal_year, sales, operating_profit, gross_profit, "
        "cogs_quantity, cogs_price, operating_cf, generation FROM quarterly_standalone "
        "WHERE ticker=? AND is_valid=1 ORDER BY period_end", (ticker,),
    ).fetchall()
    return [dict(r) for r in rows
            if r["period_end"] in vis and r["generation"] == vis[r["period_end"]]]


def get_bs_series(conn, ticker, item_key, as_of=None):
    vis = visible_generations(conn, ticker, as_of)
    rows = conn.execute(
        "SELECT period_end, value, generation FROM balance_sheet_items "
        "WHERE ticker=? AND item_key=? ORDER BY period_end", (ticker, item_key),
    ).fetchall()
    return [(r["period_end"], r["value"]) for r in rows
            if r["period_end"] in vis and r["generation"] == vis[r["period_end"]]]


def get_adjustments(conn, ticker, as_of=None):
    """一時収入などの調整（period_end → 控除合計額）。可視世代のみ。"""
    vis = visible_generations(conn, ticker, as_of)
    rows = conn.execute(
        "SELECT period_end, amount, generation FROM pl_adjustments WHERE ticker=?",
        (ticker,),
    ).fetchall()
    out = {}
    for r in rows:
        if r["period_end"] in vis and r["generation"] == vis[r["period_end"]]:
            out[r["period_end"]] = out.get(r["period_end"], 0) + r["amount"]
    return out


def get_latest_forecast(conn, ticker, as_of=None):
    """会社通期予想（最新版）。as_of指定で当時の版を再現。RowまたはNone。"""
    if as_of is None:
        return conn.execute(
            "SELECT fiscal_year, forecast_sales, forecast_op, source_date FROM company_forecasts "
            "WHERE ticker=? ORDER BY source_date DESC LIMIT 1", (ticker,)).fetchone()
    return conn.execute(
        "SELECT fiscal_year, forecast_sales, forecast_op, source_date FROM company_forecasts "
        "WHERE ticker=? AND source_date<=? ORDER BY source_date DESC LIMIT 1",
        (ticker, _iso(as_of))).fetchone()


def normalized_sales(pl_row, adjustments):
    """調整後売上（一時収入を控除）。§8-3：継続収入ベースライン"""
    adj = adjustments.get(pl_row["period_end"], 0)
    return (pl_row["sales"] or 0) - adj


def clamp(x, lo=-1.0, hi=1.0):
    return max(lo, min(hi, x))
