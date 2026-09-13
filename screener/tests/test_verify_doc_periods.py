"""tests/test_verify_doc_periods.py — 根拠期照合が TDnet リンクも解決できること。

2026-09-13: 全銘柄スキャンで TDnet 根拠の 1,089 件が「doc_id が本体に無い」になった。
リンクから doc_id を取る処理が EDINET の `?S100...` 形式しか読んでいなかった。
"""
import os
import sqlite3
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from screener.report import weekly_screen as WS   # noqa: E402


class TestDocIdFromUrl(unittest.TestCase):
    def test_edinet(self):
        self.assertEqual(WS.doc_id_from_url(
            "https://disclosure2.edinet-fsa.go.jp/WZEK0040.aspx?S100XAJ0"), "S100XAJ0")

    def test_tdnet(self):
        self.assertEqual(WS.doc_id_from_url(
            "https://www.release.tdnet.info/inbs/140120260813519912.pdf"), "140120260813519912")


class TestVerifyDocPeriods(unittest.TestCase):
    def setUp(self):
        c = sqlite3.connect(":memory:")
        c.row_factory = sqlite3.Row
        c.executescript(
            "CREATE TABLE filings (id INTEGER PRIMARY KEY, code TEXT, doc_id TEXT, date TEXT, title TEXT);"
            "CREATE TABLE financials_cum (filing_id INTEGER, period TEXT, q_no INTEGER);")
        c.execute("INSERT INTO filings VALUES (1,'2962','140120260813519912','2026-08-13','決算短信')")
        c.execute("INSERT INTO financials_cum VALUES (1,'FY2026',4)")
        c.execute("INSERT INTO filings VALUES (2,'7050','S100XAJ0','2025-12-15','半期報告書')")
        c.execute("INSERT INTO financials_cum VALUES (2,'FY2026',2)")
        self.c = c

    def test_TDnetとEDINETの両方が一致する(self):
        rows = [
            {"code": "2962", "evidence_docs":
             "S1:FY2026-Q4 https://www.release.tdnet.info/inbs/140120260813519912.pdf"},
            {"code": "7050", "evidence_docs":
             "S1:FY2026-Q2 https://disclosure2.edinet-fsa.go.jp/WZEK0040.aspx?S100XAJ0"},
        ]
        bad, ok, nolink = WS.verify_doc_periods(rows, self.c)
        self.assertEqual((bad, ok, nolink), ([], 2, 0))

    def test_期が違えば不一致(self):
        rows = [{"code": "2962", "evidence_docs":
                 "S1:FY2026-Q2 https://www.release.tdnet.info/inbs/140120260813519912.pdf"}]
        bad, ok, _ = WS.verify_doc_periods(rows, self.c)
        self.assertEqual(ok, 0)
        self.assertEqual(len(bad), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
