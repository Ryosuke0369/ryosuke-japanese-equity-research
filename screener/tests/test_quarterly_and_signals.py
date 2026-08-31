"""screener/tests/test_quarterly_and_signals.py — P2 (仕様書 §3-2 / §4)。

このファイルが守る規則: **累計は平均で嘘をつく。単独値だけが傾きを語る。**
累計から作った単独値が手計算と一致すること、そして「無効」と「まだデータが
無い」が絶対に混ざらないことを固定する。
"""
from __future__ import annotations

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))

from screener import common as C
from screener.extract import quarterly_builder as Q
from screener.signals import signal_defs as S


class _Base(unittest.TestCase):
    def setUp(self):
        fd, self.db = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.con = C.init_db(self.db)
        self._fid = 0

    def tearDown(self):
        self.con.close()
        os.unlink(self.db)

    def cum(self, code, period, q_no, item, value, *, date=None, source="tdnet",
            subs=None):
        """累計値を1行入れる。書類は四半期ごとに1つ作る（実際の短信と同じ形）。"""
        key = (code, period, q_no, date or f"2026-{q_no:02d}-01", source)
        fid = getattr(self, "_ids", {}).get(key)
        if fid is None:
            cur = self.con.execute(
                "INSERT INTO filings (code, date, type, source, doc_id, xbrl_ok) "
                "VALUES (?,?,?,?,?,1)",
                (code, key[3], "短信", source, f"D{self._fid}"))
            self._fid += 1
            fid = cur.lastrowid
            if not hasattr(self, "_ids"):
                self._ids = {}
            self._ids[key] = fid
        self.con.execute(
            "INSERT OR REPLACE INTO financials_cum (filing_id, code, period, q_no, "
            " item, value, context_ref, source_tag) VALUES (?,?,?,?,?,?,?,?)",
            (fid, code, period, q_no, item, value, f"ctx{q_no}_{item}", "t"))
        if subs is not None:
            self.con.execute(
                "INSERT OR REPLACE INTO financials_cum (filing_id, code, period, "
                " q_no, item, value, context_ref, source_tag) VALUES (?,?,?,?,?,?,?,?)",
                (fid, code, period, q_no, "consolidated_subsidiaries", subs,
                 f"ctx{q_no}_subs", "t"))
        self.con.commit()

    def q(self, code, period, q_no, item):
        r = self.con.execute(
            "SELECT value, valid_flag, invalid_reason, span_q FROM financials_q "
            "WHERE code=? AND period=? AND q_no=? AND item=?",
            (code, period, q_no, item)).fetchone()
        return r


class TestQuarterlyBuilder(_Base):
    def test_single_quarter_is_cumulative_minus_previous(self):
        """§3-2 の核心。Q単独 = 当期累計 − 前四半期累計。手計算と一致すること。

        累計 Q1 1,000 / Q2 2,500 / Q3 4,000 なら
        単独 Q1 1,000 / Q2 1,500 / Q3 1,500。
        """
        for qn, v in ((1, 1000.0), (2, 2500.0), (3, 4000.0)):
            self.cum("9999", "FY2026", qn, "revenue", v)
        Q.build(self.con)
        self.assertEqual(self.q("9999", "FY2026", 1, "revenue")["value"], 1000.0)
        self.assertEqual(self.q("9999", "FY2026", 2, "revenue")["value"], 1500.0)
        self.assertEqual(self.q("9999", "FY2026", 3, "revenue")["value"], 1500.0)
        for qn in (1, 2, 3):
            self.assertEqual(self.q("9999", "FY2026", qn, "revenue")["valid_flag"], 1)

    def test_cumulative_hides_the_slope(self):
        """3441 で実証された現象そのもの。累計の粗利率は高いまま、単独は落ちる。

        累計: Q2 粗利率 25.7%、Q3 24.2%
        単独: Q3 は 21.9%  ← 累計だけ見ていると気づけない
        """
        # 累計売上 Q2 5,000 / Q3 8,000、累計粗利 Q2 1,285 / Q3 1,936
        self.cum("3441", "FY2026", 2, "revenue", 5000.0)
        self.cum("3441", "FY2026", 2, "gross_profit", 1285.0)
        self.cum("3441", "FY2026", 3, "revenue", 8000.0)
        self.cum("3441", "FY2026", 3, "gross_profit", 1936.0)
        Q.build(self.con)
        rev = self.q("3441", "FY2026", 3, "revenue")["value"]
        gp = self.q("3441", "FY2026", 3, "gross_profit")["value"]
        self.assertEqual(rev, 3000.0)                       # 8000 - 5000
        self.assertEqual(gp, 651.0)                         # 1936 - 1285
        self.assertAlmostEqual(gp / rev * 100, 21.7, places=1)
        self.assertAlmostEqual(1936 / 8000 * 100, 24.2, places=1)   # 累計は 24.2%

    def test_stock_items_are_not_differenced(self):
        """BSは期末残高。引き算してはいけない —— 総資産の『単独値』に意味は無い。"""
        self.cum("9999", "FY2026", 2, "total_assets", 50000.0)
        self.cum("9999", "FY2026", 3, "total_assets", 52000.0)
        Q.build(self.con)
        self.assertEqual(self.q("9999", "FY2026", 3, "total_assets")["value"], 52000.0)

    def test_no_prior_cumulative_is_a_period_start_single_with_span(self):
        """手前の累計が1つも無い期は、累計そのものが『期首からの単独値』。
        0 を書いたり無効にしたりせず、span_q に何四半期ぶんかを残す ——
        3ヶ月と9ヶ月を同じ土俵に載せないための列。"""
        self.cum("9999", "FY2026", 3, "revenue", 4000.0)     # Q1もQ2も無い
        Q.build(self.con)
        row = self.q("9999", "FY2026", 3, "revenue")
        self.assertEqual(row["value"], 4000.0)
        self.assertEqual(row["span_q"], 3, "9ヶ月ぶんを1四半期と名乗っている")
        self.assertEqual(row["valid_flag"], 1)

    def test_half_year_single_is_marked_span_2(self):
        """半期報告書しか無い会社。q2の累計はH1そのもの、q4−q2は下期6ヶ月。
        どちらも四半期ではないので span_q=2 で区別できること。"""
        self.cum("3441", "FY2026", 2, "revenue", 6747128.0)
        self.cum("3441", "FY2026", 4, "revenue", 14000000.0, date="2026-10-01")
        Q.build(self.con)
        h1 = self.q("3441", "FY2026", 2, "revenue")
        h2 = self.q("3441", "FY2026", 4, "revenue")
        self.assertEqual((h1["value"], h1["span_q"]), (6747128.0, 2))
        self.assertEqual((h2["value"], h2["span_q"]), (14000000.0 - 6747128.0, 2))

    def test_consolidated_and_nonconsolidated_are_not_a_restatement(self):
        """有報は同じ期・同じ項目を連結と単体の両方で載せる。分けずに1つの
        キーへ詰めると、その差を『後の書類が言い直した』= 遡及修正と誤検出する
        (3441 で実際に65件の偽陽性が出た)。連結を正とする。"""
        cur = self.con.execute(
            "INSERT INTO filings (code, date, type, source, doc_id, xbrl_ok) "
            "VALUES ('3441','2025-10-27','有報','edinet','X1',1)")
        fid = cur.lastrowid
        for ctx, val in (("CurrentYearDuration", 10830372.0),
                         ("CurrentYearDuration_NonConsolidatedMember", 7580043.0)):
            self.con.execute(
                "INSERT INTO financials_cum (filing_id, code, period, q_no, item, "
                " value, context_ref, source_tag) VALUES (?,?,?,?,?,?,?,?)",
                (fid, "3441", "FY2025", 4, "revenue", val, ctx, "t"))
        self.con.commit()
        Q.build(self.con)
        row = self.q("3441", "FY2025", 4, "revenue")
        self.assertEqual(row["valid_flag"], 1, "連結/単体の差を遡及修正としている")
        self.assertEqual(row["value"], 10830372.0, "連結ではなく単体を採っている")

    def test_restatement_invalidates(self):
        """後の書類が同じ四半期を違う値で言い直したら遡及修正。単独値は無効。"""
        self.cum("9999", "FY2026", 1, "revenue", 1000.0)
        self.cum("9999", "FY2026", 2, "revenue", 2500.0)
        # Q3短信が Q2累計を 2,600 に言い直した（別書類）
        self.cum("9999", "FY2026", 2, "revenue", 2600.0, date="2026-09-01")
        self.cum("9999", "FY2026", 3, "revenue", 4000.0, date="2026-09-01")
        Q.build(self.con)
        row = self.q("9999", "FY2026", 2, "revenue")
        self.assertEqual(row["valid_flag"], 0)
        self.assertIn("遡及修正", row["invalid_reason"])

    def test_consolidation_scope_change_invalidates(self):
        """連結子会社数が変わったら、前四半期と同じ会社を見ていない。"""
        self.cum("9999", "FY2026", 1, "revenue", 1000.0, subs=5)
        self.cum("9999", "FY2026", 2, "revenue", 2500.0, subs=7)
        Q.build(self.con)
        row = self.q("9999", "FY2026", 2, "revenue")
        self.assertEqual(row["valid_flag"], 0)
        self.assertIn("連結範囲変更", row["invalid_reason"])

    def test_invalid_rows_are_kept(self):
        """無効行は消さない。消すと『無効』と『まだ取得していない』が
        区別できなくなる —— このリポジトリが一貫して守っている規則。"""
        self.cum("9999", "FY2026", 3, "revenue", 4000.0)
        Q.build(self.con)
        self.assertEqual(self.con.execute(
            "SELECT COUNT(*) c FROM financials_q WHERE code='9999'").fetchone()["c"], 1)


class TestSignals(_Base):
    def setUp(self):
        super().setUp()
        self.th = S.load_thresholds()

    def _series(self, code):
        return S.Series(self.con, code)

    def qrow(self, code, period, q_no, item, value, flag=1):
        self.con.execute(
            "INSERT OR REPLACE INTO financials_q (code, period, q_no, item, value, "
            " valid_flag) VALUES (?,?,?,?,?,?)", (code, period, q_no, item, value, flag))
        self.con.commit()

    def test_s1_fires_on_rising_gross_margin(self):
        """Q単独粗利率が前Q比 +1.5pt 以上で発火。手計算: 20.0% → 23.0% = +3.0pt。"""
        self.qrow("A", "FY2026", 2, "revenue", 1000.0)
        self.qrow("A", "FY2026", 2, "gross_profit", 200.0)      # 20.0%
        self.qrow("A", "FY2026", 3, "revenue", 1000.0)
        self.qrow("A", "FY2026", 3, "gross_profit", 230.0)      # 23.0%
        sig = S.s1_gross_margin_slope(self._series("A"), self.th)
        self.assertTrue(sig.fired)
        self.assertAlmostEqual(sig.value, 3.0, places=6)
        self.assertEqual(sig.direction, 1)
        self.assertIn("23.0%", sig.evidence)

    def test_s1_penalises_falling_gross_margin(self):
        """低下側は逆シグナルとして減点。3441 の Q3(25.7% → 21.9%)がこの形。"""
        self.qrow("A", "FY2026", 2, "revenue", 1000.0)
        self.qrow("A", "FY2026", 2, "gross_profit", 257.0)      # 25.7%
        self.qrow("A", "FY2026", 3, "revenue", 1000.0)
        self.qrow("A", "FY2026", 3, "gross_profit", 219.0)      # 21.9%
        sig = S.s1_gross_margin_slope(self._series("A"), self.th)
        self.assertTrue(sig.fired)
        self.assertEqual(sig.direction, -1, "低下を加点している")
        self.assertAlmostEqual(sig.value, -3.8, places=6)

    def test_s1_ignores_invalid_rows(self):
        """無効行を使うと決算期変更や遡及修正を『傾き』として読んでしまう。"""
        self.qrow("A", "FY2026", 2, "revenue", 1000.0, flag=0)
        self.qrow("A", "FY2026", 2, "gross_profit", 200.0, flag=0)
        self.qrow("A", "FY2026", 3, "revenue", 1000.0)
        self.qrow("A", "FY2026", 3, "gross_profit", 230.0)
        sig = S.s1_gross_margin_slope(self._series("A"), self.th)
        self.assertFalse(sig.fired, "無効行を比較対象にしている")

    def test_s2_inventory_build_vs_stagnation(self):
        """**在庫増は両義的**。同じ在庫増でも売上加速なら仕込み、減速なら滞留。"""
        def setup(code, rev_now):
            for (y, q), rev in (((2025, 2), 900.0), ((2025, 3), 900.0),
                                ((2026, 2), 1000.0)):
                self.qrow(code, f"FY{y}", q, "revenue", rev)
            self.qrow(code, "FY2026", 3, "revenue", rev_now)
            self.qrow(code, "FY2026", 2, "inventories_total", 100.0)
            self.qrow(code, "FY2026", 3, "inventories_total", 130.0)   # +30%
            self.qrow(code, "FY2026", 2, "cogs", 700.0)
            self.qrow(code, "FY2026", 3, "cogs", 707.0)                # +1%
        setup("ACC", 1200.0)       # 売上YoY 11.1% → 33.3% = 加速
        setup("DEC", 900.0)        # 売上YoY 11.1% → 0.0%  = 減速
        a = S.s2_inventory_split(self._series("ACC"), self.th)
        d = S.s2_inventory_split(self._series("DEC"), self.th)
        self.assertTrue(a.fired and a.direction == 1, "仕込みを加点していない")
        self.assertIn("仕込み", a.evidence)
        self.assertTrue(d.fired and d.direction == -1, "滞留を減点していない")
        self.assertIn("滞留", d.evidence)

    def test_s3b_detects_transfer_to_machinery(self):
        """建仮減 × 機械装置増 = 振替。稼働開始の証拠(3905のQ2チェックと同型)。"""
        self.qrow("A", "FY2026", 2, "construction_in_progress", 1000.0)
        self.qrow("A", "FY2026", 3, "construction_in_progress", 700.0)   # -30%
        self.qrow("A", "FY2026", 2, "machinery_and_equipment", 5000.0)
        self.qrow("A", "FY2026", 3, "machinery_and_equipment", 5400.0)   # +8%
        s3, s3b = S.s3_construction_in_progress(self._series("A"), self.th)
        self.assertFalse(s3.fired, "建仮は減っているのに稼働前投資で発火している")
        self.assertTrue(s3b.fired)
        self.assertIn("振替", s3b.evidence)

    def test_s4_needs_to_outpace_revenue(self):
        """契約負債が増えても、売上が同じだけ増えていれば顧客コミットではない。"""
        self.qrow("A", "FY2026", 2, "contract_liabilities", 100.0)
        self.qrow("A", "FY2026", 3, "contract_liabilities", 130.0)   # +30%
        self.qrow("A", "FY2026", 2, "revenue", 1000.0)
        self.qrow("A", "FY2026", 3, "revenue", 1300.0)               # +30%
        self.assertFalse(S.s4_contract_liabilities(self._series("A"), self.th).fired)
        self.qrow("A", "FY2026", 3, "revenue", 1050.0)               # +5% だけ
        sig = S.s4_contract_liabilities(self._series("A"), self.th)
        self.assertTrue(sig.fired)
        self.assertAlmostEqual(sig.value, 25.0, places=6)

    def test_s5_turnaround_and_dead_guidance(self):
        """①赤字→黒字転換 ②累計進捗が四半期数比を大幅超過(3441の113%型)。"""
        self.qrow("A", "FY2026", 1, "operating_income", 30.0)
        self.qrow("A", "FY2026", 2, "operating_income", -10.0)
        self.qrow("A", "FY2026", 3, "operating_income", 93.0)
        self.con.execute("INSERT INTO guidance (code, date, fy, item, value) "
                         "VALUES ('A','2026-06-01','FY2026','operating_income',100.0)")
        self.con.commit()
        sig = S.s5_turnaround_and_progress(self.con, self._series("A"), self.th)
        self.assertTrue(sig.fired)
        self.assertIn("赤字→黒字転換", sig.evidence)
        # 累計 30 - 10 + 93 = 113 / 予想 100 = 113%、四半期数比 75% → 超過 +38pt
        self.assertIn("113%", sig.evidence)
        self.assertIn("死んだガイダンス", sig.evidence)


if __name__ == "__main__":
    unittest.main(verbosity=2)
