"""seasonality_profile の算術（除外理由の付与・単独値の差分・ばらつき）。"""
import unittest

from screener.report import seasonality_profile as S


def _row(fy, q, op, sales=None):
    return {"fy_end": fy, "per_type": q, "op": op, "sales": sales,
            "f_op": None, "f_sales": None, "disc_date": fy}


class TestSeasonalityProfile(unittest.TestCase):
    def setUp(self):
        self.tab = S.fiscal_table([
            _row("2025-03-31", "1Q", -200, 100), _row("2025-03-31", "2Q", 300, 250),
            _row("2025-03-31", "3Q", 500, 400), _row("2025-03-31", "FY", 1000, 1000),
            _row("2026-03-31", "2Q", 100, 300), _row("2026-03-31", "FY", 800, 900),
        ])

    def test_negative_progress_is_shown_with_reason(self):
        prog = S.raw_progress(self.tab, "op")
        p, why = prog[("2025-03-31", "1Q")]
        self.assertAlmostEqual(p, -0.2)
        self.assertIn("累計赤字", why)
        self.assertIsNone(prog[("2025-03-31", "2Q")][1])

    def test_standalone_needs_previous_cumulative(self):
        st = S.standalone(self.tab, "op")
        self.assertEqual(st[("2025-03-31", "1Q")], -200)
        self.assertEqual(st[("2025-03-31", "2Q")], 500)
        self.assertEqual(st[("2025-03-31", "FY")], 500)
        # 1Q が無い年度の 2Q は単独にしない（推測しない）
        self.assertNotIn(("2026-03-31", "2Q"), st)
        self.assertNotIn(("2026-03-31", "FY"), st)   # 3Q が無いので 4Q 単独も出さない

    def test_spread(self):
        s = S.spread([0.10, 0.20, 0.30, 0.40])
        self.assertEqual(s["n"], 4)
        self.assertAlmostEqual(s["median"], 0.25)
        self.assertAlmostEqual(s["iqr"], 0.15)
        self.assertEqual(S.spread([])["n"], 0)
        self.assertIsNone(S.spread([0.1])["sd"])


if __name__ == "__main__":
    unittest.main()
