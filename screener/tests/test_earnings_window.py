"""tests/test_earnings_window.py — 決算窓フィルタの不変条件。

1. 期ラベル FY{Y}-Q{q} は期末月で暦に写す（ラベル文字列を暦日と比べない）。
2. 休場日（祝日・国民の休日）の推定日は翌営業日に寄せ、寄せたことを残す。
3. 並べ替えの第一基準は「直前四半期が本体にあるか」。stale_flag ではない。
"""
import os
import sys
import unittest
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from screener.report import earnings_window as EW   # noqa: E402


class TestFiscalLabel(unittest.TestCase):
    def test_7月期(self):
        # 3441 山王: FY2025-Q4 = 2025年7月期末（有報 第67期 2024/08-2025/07）
        self.assertEqual(EW.fiscal_label_ym("FY2025", 4, 7), (2025, 7))
        self.assertEqual(EW.fiscal_label_ym("FY2026", 3, 7), (2026, 4))
        self.assertEqual(EW.fiscal_label_ym("FY2026", 1, 7), (2025, 10))

    def test_3月期と1月期(self):
        self.assertEqual(EW.fiscal_label_ym("FY2026", 1, 3), (2025, 6))
        self.assertEqual(EW.fiscal_label_ym("FY2027", 1, 1), (2026, 4))

    def test_読めないラベルはNone(self):
        self.assertIsNone(EW.fiscal_label_ym("2026-Q1", 1, 3))
        self.assertIsNone(EW.fiscal_label_ym("FY2026", 1, None))

    def test_ym_add(self):
        self.assertEqual(EW.ym_add((2026, 1), -3), (2025, 10))
        self.assertEqual(EW.ym_add((2025, 11), 3), (2026, 2))


class TestHoliday(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            cls.cal = EW._jp_calendar()
        except SystemExit:
            raise unittest.SkipTest("別枠 jp_calendar が無い環境")

    def test_敬老の日と国民の休日と秋分の日(self):
        for d in (date(2026, 9, 21), date(2026, 9, 22), date(2026, 9, 23)):
            self.assertFalse(self.cal.is_business_day(d), d)
        rolled, ok = EW.roll_business_day(date(2026, 9, 21), self.cal)
        self.assertEqual((rolled, ok), (date(2026, 9, 24), False))

    def test_営業日はそのまま(self):
        self.assertEqual(EW.roll_business_day(date(2026, 9, 18), self.cal),
                         (date(2026, 9, 18), True))


class TestSort(unittest.TestCase):
    def test_直前四半期ありが先_stale_flagは見ない(self):
        rows = [
            {"code": "A", "prev_quarter_in_db": 0, "est_date": "2026-09-14", "score": "0.9", "stale_flag": "0"},
            {"code": "B", "prev_quarter_in_db": 1, "est_date": "2026-09-30", "score": "0.1", "stale_flag": "1"},
            {"code": "C", "prev_quarter_in_db": 1, "est_date": "2026-09-14", "score": "", "stale_flag": "0"},
            {"code": "D", "prev_quarter_in_db": "", "est_date": "2026-09-14", "score": "0.5"},
        ]
        self.assertEqual([r["code"] for r in sorted(rows, key=EW.sort_key)], ["C", "B", "A", "D"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
