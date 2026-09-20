"""外部リスト（株探ウォッチ）の取り込み。**記録専用**であることを構造で守る。"""
import os
import re
import sqlite3
import tempfile
import unittest

from screener.extract import kabutan_watch as K

PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# スコア・候補・採否を作る側。ここが kabutan_watch を読んではいけない
SCORING_SIDE = ["signals", "projection",
                os.path.join("report", "weekly_screen.py"),
                os.path.join("report", "paper_weekly.py"),
                os.path.join("report", "backtest_v2.py"),
                os.path.join("report", "backtest_eval.py"),
                os.path.join("report", "entry_sim.py"),
                os.path.join("report", "exit_sim.py"),
                os.path.join("report", "earnings_window.py")]


def _py_files(path):
    if path.endswith(".py"):
        yield path
        return
    for root, _dirs, files in os.walk(path):
        for f in files:
            if f.endswith(".py"):
                yield os.path.join(root, f)


class TestRecordOnly(unittest.TestCase):
    def test_scoring_side_does_not_read_the_table(self):
        pat = re.compile(r"kabutan")
        for rel in SCORING_SIDE:
            for p in _py_files(os.path.join(PKG, rel)):
                with open(p, encoding="utf-8") as fh:
                    with self.subTest(file=os.path.relpath(p, PKG)):
                        self.assertIsNone(pat.search(fh.read()))

    def test_module_does_not_fetch(self):
        with open(os.path.join(PKG, "extract", "kabutan_watch.py"), encoding="utf-8") as fh:
            src = fh.read()
        for word in ("requests", "urllib", "http://", "Fetcher"):
            self.assertNotIn(word, src.replace("http://", "", 0) if False else src)


class TestParse(unittest.TestCase):
    def test_code_and_date_normalisation(self):
        rows = [{"日付": "2026/9/20", "リスト": "上方修正有望", "コード": "６７５８",
                 "銘柄名": "ソニーグループ", "進捗率": "72.5%", "乖離率": "＋18.3"},
                {"日付": "2026-09-20", "リスト": "上方修正有望", "コード": "13820",
                 "銘柄名": "ホーブ", "進捗率": "n/a"}]
        out, skipped = K.parse_rows(rows, "x.csv")
        self.assertEqual([r["code"] for r in out], ["6758", "1382"])
        self.assertEqual(out[0]["date"], "2026-09-20")
        self.assertIn('"進捗率": 72.5', out[0]["metrics_json"])
        self.assertIn('"乖離率": 18.3', out[0]["metrics_json"])
        self.assertIsNone(out[1]["metrics_json"])       # 数値に読めない列は入れない
        self.assertEqual(skipped, [])

    def test_missing_columns_use_defaults_or_are_skipped(self):
        out, skipped = K.parse_rows([{"code": "6758"}], "x.csv",
                                    default_date="2026-09-20", default_type="上方修正有望")
        self.assertEqual(len(out), 1)
        out2, skipped2 = K.parse_rows([{"code": "6758"}], "x.csv")
        self.assertEqual(out2, [])
        self.assertIn("日付", skipped2[0]["reason"])

    def test_unreadable_code_is_skipped_not_guessed(self):
        out, skipped = K.parse_rows([{"date": "2026-09-20", "list_type": "x",
                                      "code": "----", "name": "?"}], "x.csv")
        self.assertEqual(out, [])
        self.assertIn("コード", skipped[0]["reason"])


class TestImport(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(":memory:")
        self.con.execute("CREATE TABLE companies (code TEXT, universe_flag INTEGER)")
        self.con.executemany("INSERT INTO companies VALUES (?,?)",
                             [("6758", 0), ("1382", 1)])
        self.rows, _ = K.parse_rows(
            [{"date": "2026-09-20", "list_type": "上方修正有望", "code": "6758", "name": "A"},
             {"date": "2026-09-20", "list_type": "上方修正有望", "code": "1382", "name": "B"},
             {"date": "2026-09-20", "list_type": "上方修正有望", "code": "9999", "name": "C"}],
            "w.csv")

    def test_universe_flag_is_frozen_at_import(self):
        st = K.import_rows(self.con, self.rows)
        self.assertEqual(st["new"], 3)
        self.assertEqual(st["in_universe"], 1)
        self.assertEqual(st["unknown_code"], ["9999"])
        got = dict(self.con.execute("SELECT code, in_universe FROM kabutan_watch"))
        self.assertEqual(got, {"6758": 0, "1382": 1, "9999": 0})

    def test_idempotent(self):
        K.import_rows(self.con, self.rows)
        st = K.import_rows(self.con, self.rows)
        self.assertEqual((st["new"], st["updated"]), (0, 3))
        self.assertEqual(self.con.execute(
            "SELECT COUNT(*) FROM kabutan_watch").fetchone()[0], 3)


class TestReadCsv(unittest.TestCase):
    def test_cp932_file(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "a.csv")
            with open(p, "w", encoding="cp932", newline="") as fh:
                fh.write("日付,リスト,コード,銘柄名\n2026-09-20,上方修正有望,6758,ソニー\n")
            rows = K.read_csv(p)
            self.assertEqual(rows[0]["銘柄名"], "ソニー")


if __name__ == "__main__":
    unittest.main()
