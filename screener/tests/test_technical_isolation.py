"""テクニカル層はスコアに加点できない構造であること（backtest_acceptance_criteria.md シャドウE E-0）。

テクニカルは証拠で選ばれた銘柄への veto・タイミング・撤退だけに使う。
候補の追加・加点を「気をつける」ではなく import と書き込みの面で塞ぐ。
"""
import os
import re
import unittest

PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # screener/
TECH = os.path.join(PKG, "technical")

# スコア・候補・建玉を作る側
SCORING_SIDE = ["signals", "projection", "extract",
                os.path.join("report", "weekly_screen.py"),
                os.path.join("report", "paper_weekly.py"),
                os.path.join("report", "backtest_v2.py"),
                os.path.join("report", "backtest_eval.py"),
                os.path.join("report", "earnings_window.py")]

IMPORTS_TECH = re.compile(r"screener\.technical|from\s+screener\s+import\s+[^\n]*\btechnical\b"
                          r"|from\s+\.+technical|import\s+technical\b")
SQL_WRITE = re.compile(r"\b(INSERT|UPDATE|DELETE|REPLACE|CREATE|DROP|ALTER)\b")
SCORE_TABLES = ("scores", "signals", "forecast_snapshots", "paper_trades",
                "shadow_snapshots", "shadow_trades")


def _py_files(path):
    if path.endswith(".py"):
        yield path
        return
    for root, _, files in os.walk(path):
        for f in files:
            if f.endswith(".py"):
                yield os.path.join(root, f)


def _read(p):
    with open(p, encoding="utf-8") as fh:
        return fh.read()


class TestTechnicalIsolation(unittest.TestCase):
    def test_scoring_side_does_not_import_technical(self):
        for rel in SCORING_SIDE:
            for p in _py_files(os.path.join(PKG, rel)):
                with self.subTest(file=os.path.relpath(p, PKG)):
                    self.assertIsNone(IMPORTS_TECH.search(_read(p)))

    def test_technical_does_not_write_or_call_scorers(self):
        files = list(_py_files(TECH))
        self.assertTrue(files)
        for p in files:
            src = _read(p)
            with self.subTest(file=os.path.relpath(p, PKG)):
                self.assertIsNone(SQL_WRITE.search(src), "SQL の書き込み文がある")
                self.assertNotIn("init_db(", src)
                self.assertNotIn("C.connect(", src)
                self.assertIsNone(re.search(r"screener\.signals|from\s+screener\s+import\s+signals", src))
                for t in SCORE_TABLES:
                    self.assertIsNone(re.search(r"(INTO|UPDATE|FROM)\s+%s\b" % t, src, re.I),
                                      "スコア・建玉のテーブル %s に触れている" % t)

    def test_db_opened_read_only(self):
        src = _read(os.path.join(TECH, "event_response.py"))
        self.assertIn("mode=ro", src)
        self.assertEqual(len(re.findall(r"sqlite3\.connect\(", src)), 1)


if __name__ == "__main__":
    unittest.main()
