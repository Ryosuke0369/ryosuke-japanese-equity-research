"""取得後に「頼んだ範囲が実際に埋まったか」を突き合わせる（calibration_backlog §53）。

同じ型の事故が3件あった: 待機ループの空転 / 部分取得を ok と記録 / 別の期間を埋めた。
どれも「実行は成功」だが「意図した範囲は埋まっていない」。ここを機械が数える。
"""
import unittest
from datetime import date

from screener import common as C


class TestVerifyRange(unittest.TestCase):
    def test_all_present(self):
        days = ["2026-09-14", "2026-09-15", "2026-09-16", "2026-09-17", "2026-09-18"]
        v = C.verify_range("t", date(2026, 9, 14), date(2026, 9, 18), days)
        self.assertTrue(v["ok"])
        self.assertEqual((v["requested"], v["present"]), (5, 5))

    def test_missing_days_are_counted_and_listed(self):
        v = C.verify_range("t", date(2026, 9, 14), date(2026, 9, 18),
                           ["2026-09-14", "2026-09-18"])
        self.assertFalse(v["ok"])
        self.assertEqual(v["missing"], ["2026-09-15", "2026-09-16", "2026-09-17"])

    def test_weekends_are_not_requested(self):
        # 9/19(土) 9/20(日) は要求に入らない
        v = C.verify_range("t", date(2026, 9, 18), date(2026, 9, 21),
                           ["2026-09-18", "2026-09-21"])
        self.assertTrue(v["ok"])
        self.assertEqual(v["requested"], 2)

    def test_wrong_range_is_caught(self):
        """3件目の事故の形: 頼んだのは2月〜7月、埋めたのは前年の別期間。"""
        filled = ["2024-12-3%d" % i for i in (0, 1)]
        v = C.verify_range("t", date(2026, 2, 2), date(2026, 2, 6), filled)
        self.assertFalse(v["ok"])
        self.assertEqual(v["present"], 0)
        self.assertEqual(len(v["missing"]), 5)

    def test_accepts_date_objects(self):
        v = C.verify_range("t", date(2026, 9, 14), date(2026, 9, 15),
                           [date(2026, 9, 14), date(2026, 9, 15)])
        self.assertTrue(v["ok"])


if __name__ == "__main__":
    unittest.main()
