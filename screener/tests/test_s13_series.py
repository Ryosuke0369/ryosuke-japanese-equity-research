"""tests/test_s13_series.py — S13 受注の四半期単独推移（§34）。

守る不変条件:
  - 受注高は**フロー**なので年度内で累計を差分する。受注残高は**ストック**なので差分しない
  - B/B は同じ期間同士でのみ割る（1四半期あたりに正規化してから割る）
  - 年度をまたいだら差分しない（Q1 は期首からの累計そのもの）
  - 同じ期を語る書類が複数あれば後から出たほうを採る
"""
import os
import sqlite3
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from screener.projection import materialize as MZ          # noqa: E402
from screener.report import s13_series as S13S             # noqa: E402

MAIN_DDL = """
CREATE TABLE filings (id INTEGER PRIMARY KEY, code TEXT, date TEXT, title TEXT);
CREATE TABLE financials_cum (filing_id INTEGER, code TEXT, period TEXT, q_no INTEGER);
CREATE TABLE s13_orders (filing_id INTEGER PRIMARY KEY, code TEXT, doc_id TEXT,
  doc_date TEXT, doc_source TEXT, available INTEGER, orders_amount REAL,
  closing_backlog REAL, opening_backlog REAL, completed_amount REAL,
  backlog_yoy_pct REAL, orders_yoy_pct REAL);
"""


def _main(rows):
    """rows: (filing_id, fy, q, doc_date, cum_orders, closing_backlog)。"""
    con = sqlite3.connect(":memory:")
    con.row_factory = sqlite3.Row
    con.executescript(MAIN_DDL)
    for fid, fy, q, dd, orders, backlog in rows:
        con.execute("INSERT INTO filings (id, code, date, title) VALUES (?,?,?,?)",
                    (fid, "T", dd, "第%d四半期" % q))
        con.execute("INSERT INTO financials_cum (filing_id, code, period, q_no)"
                    " VALUES (?,?,?,?)", (fid, "T", "FY%d" % fy, q))
        con.execute(
            "INSERT INTO s13_orders (filing_id, code, doc_id, doc_date, doc_source,"
            " available, orders_amount, closing_backlog) VALUES (?,?,?,?,?,1,?,?)",
            (fid, "T", "D%d" % fid, dd, "tdnet", orders, backlog))
    con.commit()
    return con


def _proj(rows):
    """rows: (period_end, period_start, span_q, sales)。"""
    con = sqlite3.connect(":memory:")
    con.row_factory = sqlite3.Row
    con.executescript(MZ.DDL)
    for pe, ps, sp, sales in rows:
        con.execute(
            "INSERT INTO quarterly_standalone_all (ticker, fiscal_year, quarter_type,"
            " period_end, span_q, period_start, sales, is_valid)"
            " VALUES ('T', ?, ?, ?, ?, ?, ?, 1)",
            (int(pe[2:6]), "%dQ" % int(pe[-1]), pe, sp, ps, sales))
    con.commit()
    return con


class TestQuarterlySeries(unittest.TestCase):
    def test_受注高は年度内で累計を差分する(self):
        m = _main([(1, 2026, 1, "2026-03-10", 1354.0, 5000.0),
                   (2, 2026, 2, "2026-06-10", 4298.0, 6200.0),
                   (3, 2026, 3, "2026-09-14", 5748.0, 6100.0)])
        p = _proj([("FY2026-Q1", "FY2026-Q1", 1, 2051.0),
                   ("FY2026-Q2", "FY2026-Q2", 1, 1691.9),
                   ("FY2026-Q3", "FY2026-Q3", 1, 1559.2)])
        r = S13S.quarterly_series(m, p, "T", "2026-09-18")
        got = {x["period_end"]: x for x in r["rows"]}
        self.assertAlmostEqual(got["FY2026-Q1"]["orders"], 1354.0)
        self.assertAlmostEqual(got["FY2026-Q2"]["orders"], 2944.0)   # 4298-1354
        self.assertAlmostEqual(got["FY2026-Q3"]["orders"], 1450.0)   # 5748-4298
        # 受注残高はストック。差分しない。
        self.assertAlmostEqual(got["FY2026-Q3"]["closing_backlog"], 6100.0)

    def test_BB比(self):
        m = _main([(1, 2026, 1, "2026-03-10", 1354.0, 5000.0),
                   (2, 2026, 2, "2026-06-10", 4298.0, 6200.0),
                   (3, 2026, 3, "2026-09-14", 5748.0, 6100.0)])
        p = _proj([("FY2026-Q1", "FY2026-Q1", 1, 2051.0),
                   ("FY2026-Q2", "FY2026-Q2", 1, 1691.9),
                   ("FY2026-Q3", "FY2026-Q3", 1, 1559.2)])
        got = {x["period_end"]: x
               for x in S13S.quarterly_series(m, p, "T", "2026-09-18")["rows"]}
        # 6838 の人手計算と同じ形: 0.66 → 1.74 → 0.93
        self.assertAlmostEqual(got["FY2026-Q1"]["bb"], 0.66, places=2)
        self.assertAlmostEqual(got["FY2026-Q2"]["bb"], 1.74, places=2)
        self.assertAlmostEqual(got["FY2026-Q3"]["bb"], 0.93, places=2)

    def test_spanが違うものを素のまま割らない(self):
        # 受注は H1 累計（span=2）、売上も H1（span=2）。どちらも1四半期あたりに
        # 直してから割るので、比は 2944/1871 ではなく (2944/2)/(3743/2)。
        m = _main([(1, 2026, 2, "2026-06-10", 3742.0, 6200.0)])
        p = _proj([("FY2026-Q2", "FY2026-Q1", 2, 3742.889)])
        row = S13S.quarterly_series(m, p, "T", "2026-09-18")["rows"][0]
        self.assertEqual(row["span_q"], 2)
        self.assertAlmostEqual(row["orders_per_q"], 1871.0, places=0)
        self.assertAlmostEqual(row["bb"], 1.0, places=2)

    def test_年度をまたいだら差分しない(self):
        m = _main([(1, 2025, 4, "2025-12-10", 8000.0, 4000.0),
                   (2, 2026, 1, "2026-03-10", 1354.0, 5000.0)])
        p = _proj([("FY2026-Q1", "FY2026-Q1", 1, 2051.0)])
        got = {x["period_end"]: x
               for x in S13S.quarterly_series(m, p, "T", "2026-09-18")["rows"]}
        self.assertAlmostEqual(got["FY2026-Q1"]["orders"], 1354.0)
        self.assertIn("期首からの累計", got["FY2026-Q1"]["basis"])

    def test_受注の節が無ければ空で理由が出る(self):
        m = _main([])
        p = _proj([])
        r = S13S.quarterly_series(m, p, "T", "2026-09-18")
        self.assertEqual(r["rows"], [])
        self.assertIn("受注実績", r["note"])

    def test_売上が取れなければBBは未算出で理由が出る(self):
        m = _main([(1, 2026, 1, "2026-03-10", 1354.0, 5000.0)])
        p = _proj([])
        r = S13S.quarterly_series(m, p, "T", "2026-09-18")
        self.assertIsNone(r["rows"][0]["bb"])
        self.assertIn("B/B", r["note"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
