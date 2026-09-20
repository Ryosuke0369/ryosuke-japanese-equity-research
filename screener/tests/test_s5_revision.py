"""シャドウG（S5 → 上方修正）の窓・方向・ベースレートの取り方を固定する。"""
import unittest
from datetime import date

from screener.report import s5_revision as G


class TestJudgmentDates(unittest.TestCase):
    def test_mondays_rolled_to_business_days(self):
        cal = ["2026-07-24", "2026-07-27", "2026-08-03", "2026-08-11", "2026-09-14"]
        got = G.judgment_dates(cal)
        self.assertIn("2026-07-27", got)          # 最初の月曜
        self.assertIn("2026-08-11", got)          # 8/10 が休場なら翌営業日
        self.assertTrue(all(d <= G.LAST_DATE.isoformat() for d in got))


class TestOutcome(unittest.TestCase):
    def setUp(self):
        self.revs = {"1111": [("2026-08-01", "up"), ("2026-09-01", "down")],
                     "2222": [("2026-07-01", "up")],
                     "3333": [("2026-08-20", "unknown")]}

    def test_first_revision_in_window_wins(self):
        self.assertEqual(G.outcome_for(self.revs, "1111", "2026-07-27", "2026-10-25"), "up")

    def test_revision_before_judgment_is_not_counted(self):
        self.assertEqual(G.outcome_for(self.revs, "2222", "2026-07-27", "2026-10-25"), "none")

    def test_outside_window_is_none(self):
        self.assertEqual(G.outcome_for(self.revs, "1111", "2026-07-27", "2026-07-31"), "none")

    def test_unknown_direction_is_kept_separate(self):
        self.assertEqual(G.outcome_for(self.revs, "3333", "2026-07-27", "2026-10-25"), "unknown")


class TestS5State(unittest.TestCase):
    def test_states(self):
        self.assertEqual(G.s5_state({"S5": {"available": True, "score": 0.4,
                                            "details": {}}})[0], "S5発火(進捗超過)")
        self.assertEqual(G.s5_state({"S5": {"available": True, "score": 0.0,
                                            "details": {}}})[0], "S5非発火")
        self.assertEqual(G.s5_state({"S5": {"available": False}})[0], "S5評価不能")
        self.assertEqual(G.s5_state({})[0], "S5評価不能")

    def test_guidance_dead_flag(self):
        self.assertTrue(G.s5_state({"S5": {"available": True, "score": 0.4,
                                           "details": {"guidance_dead": True}}})[2])

    def test_guidance_dead_survives_strict_rule6(self):
        # evidence_strict は §36 の規則6 で加点しない。available=False に落ちるが
        # フラグは strict_flags に残る。ここを拾わないと 0件になる（2026-09-20 に踏んだ）
        r = {"S5": {"available": False, "score": 0.0, "details": {"guidance_dead": True},
                    "strict_flags": ["guidance_dead_unscored"]}}
        group, score, dead = G.s5_state(r)
        self.assertEqual(group, "S5評価不能")
        self.assertTrue(dead)


class TestDiffCI(unittest.TestCase):
    def test_zero_difference_has_ci_around_zero(self):
        d, lo, hi = G.diff_ci(0.10, 500, 0.10, 500)
        self.assertAlmostEqual(d, 0.0)
        self.assertLess(lo, 0)
        self.assertGreater(hi, 0)

    def test_missing_inputs(self):
        self.assertEqual(G.diff_ci(None, 10, 0.1, 10), (None, None, None))


if __name__ == "__main__":
    unittest.main()
