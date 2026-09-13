"""tests/test_fiscal_year_end.py — 書類タイトルからの決算期末月の導出。

2026-09-13: 半期報告書のタイトルが半期の期間だけを書く会社（4396 / 4495）で、
6月期が12月期と判定されていた。直前四半期の判定（evidence_strict）と発表日カレンダーの
期末月がこれに依存する。
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from screener.projection import materialize as MZ   # noqa: E402


class TestFiscalYearEndMonth(unittest.TestCase):
    def test_有報は1年の期間なので期末月を返す(self):
        self.assertEqual(MZ.fiscal_year_end_month(
            "有価証券報告書－第46期(2024/07/01－2025/06/30)"), 6)

    def test_半期の期間しか書いていない半期報告書は使わない(self):
        # 4396 システムサポートHD の実タイトル
        self.assertIsNone(MZ.fiscal_year_end_month(
            "半期報告書－第47期(2025/07/01－2025/12/31)"))

    def test_年度の期間を書いた半期報告書は使ってよい(self):
        # 3441 山王の実タイトル
        self.assertEqual(MZ.fiscal_year_end_month(
            "半期報告書－第68期(2025/08/01－2026/07/31)"), 7)

    def test_全角数字でも読む(self):
        self.assertEqual(MZ.fiscal_year_end_month(
            "有価証券報告書－第９期（２０２４／０８／０１－２０２５／０７／３１）"), 7)

    def test_読めなければNone(self):
        self.assertIsNone(MZ.fiscal_year_end_month("訂正報告書"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
