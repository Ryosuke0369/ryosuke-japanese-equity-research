"""「取得済み・解析済み」は存在ではなく完全性で持つ（calibration_backlog §40）。

日次株価の「その日の行が1つでもあれば取得済み」（9月欠落の原因）と同型の
判定が、xbrl_parser --resume と presentation_text にも残っていた。
"""
import os
import sqlite3
import tempfile
import unittest

from screener.extract import presentation_text as P
from screener.extract import xbrl_parser as X


class TestResumeWhere(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(":memory:")
        self.con.executescript(
            "CREATE TABLE filings (id INTEGER PRIMARY KEY, xbrl_path TEXT);"
            "CREATE TABLE financials_cum (filing_id INTEGER, item TEXT);"
            "CREATE TABLE guidance (filing_id INTEGER, item TEXT);")
        self.con.executemany("INSERT INTO filings VALUES (?,?)",
                             [(i, "z%d.zip" % i) for i in range(1, 6)])
        self.con.executemany("INSERT INTO financials_cum VALUES (?,?)", [
            (1, "revenue"), (1, "operating_income"), (1, "net_assets"),
            (1, "cash"), (1, "total_assets"),          # 5項目 = 解析済み
            (2, "revenue"), (2, "operating_income"),   # 2項目 = 部分解析
        ])
        self.con.execute("INSERT INTO guidance VALUES (3, 'operating_income')")

    def _todo(self):
        sql = "SELECT id FROM filings WHERE xbrl_path IS NOT NULL" + X.RESUME_WHERE
        return sorted(r[0] for r in self.con.execute(sql))

    def test_partial_parse_is_retried(self):
        # 2 は2項目しか無い（旧実装では「解析済み」として二度と読まれなかった）
        self.assertIn(2, self._todo())

    def test_complete_parse_is_skipped(self):
        self.assertNotIn(1, self._todo())

    def test_guidance_only_filing_is_done(self):
        # 予想の修正開示は財務諸表を持たない。guidance に入っていれば解析済み
        self.assertNotIn(3, self._todo())

    def test_unparsed_filings_remain(self):
        self.assertEqual(self._todo(), [2, 4, 5])


class TestPresentationMinSize(unittest.TestCase):
    def test_empty_output_is_not_treated_as_done(self):
        self.assertGreater(P.MIN_TEXT_BYTES, 0)
        with tempfile.TemporaryDirectory() as d:
            empty = os.path.join(d, "a.txt")
            open(empty, "w").close()
            full = os.path.join(d, "b.txt")
            with open(full, "w", encoding="utf-8") as fh:
                fh.write("x" * (P.MIN_TEXT_BYTES + 1))
            self.assertLess(os.path.getsize(empty), P.MIN_TEXT_BYTES)
            self.assertGreaterEqual(os.path.getsize(full), P.MIN_TEXT_BYTES)


if __name__ == "__main__":
    unittest.main()
