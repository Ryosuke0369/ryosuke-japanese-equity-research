"""screener/tests/test_projection.py — 投影層の不変条件 (統合引継ぎ書 v1.1 §9-4)。

別枠の tests/test_pit.py と同じ役割を本体側で果たす。ここが守るのは1つ:
**as_of より後に公知になった値は、絶対に返らない。**
バックテストが未来を見ていたら、出てくる数字は全部嘘になる。
"""
from __future__ import annotations

import ast
import os
import sys
import tempfile
import unittest
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))

from screener import common as C
from screener.projection import pit, adapter
from screener.extract import adjustment_recorder as AR


class _Base(unittest.TestCase):
    def setUp(self):
        fd, self.db = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.con = C.init_db(self.db)

    def tearDown(self):
        self.con.close()
        os.unlink(self.db)

    def filing(self, code, d, title="短信", source="tdnet"):
        return self.con.execute(
            "INSERT INTO filings (code, date, type, source, doc_id, xbrl_ok) "
            "VALUES (?,?,?,?,?,1)",
            (code, d, "短信", source, f"D{d}{code}{title}")).lastrowid

    def cum(self, fid, code, period, q_no, item, value, ctx="c1"):
        self.con.execute(
            "INSERT OR REPLACE INTO financials_cum (filing_id, code, period, q_no, "
            " item, value, context_ref, source_tag) VALUES (?,?,?,?,?,?,?,?)",
            (fid, code, period, q_no, item, value, ctx, "t"))
        self.con.commit()

    def q(self, code, period, q_no, item, value, span=1, flag=1):
        self.con.execute(
            "INSERT OR REPLACE INTO financials_q (code, period, q_no, item, value, "
            " valid_flag, span_q) VALUES (?,?,?,?,?,?,?)",
            (code, period, q_no, item, value, flag, span))
        self.con.commit()


class TestPitInvariants(_Base):
    def test_p1_restatement_switches_at_the_correction_date(self):
        """P1: 訂正短信を跨いで as_of を動かすと返る値が変わる。

        本体は generation 列を持たない。financials_cum が書類ごとに
        ファクトを保持しているので、訂正は自然に別行になり
        filings.date の降順で最新版が決まる。
        """
        f1 = self.filing("9999", "2026-05-10")
        f2 = self.filing("9999", "2026-08-20")          # 訂正短信
        self.cum(f1, "9999", "FY2026", 1, "revenue", 1400.0)
        self.cum(f2, "9999", "FY2026", 1, "revenue", 1350.0)
        before = pit.visible_cum(self.con, "9999", date(2026, 8, 19))
        after = pit.visible_cum(self.con, "9999", date(2026, 8, 21))
        self.assertEqual([r["value"] for r in before], [1400.0])
        self.assertEqual([r["value"] for r in after], [1350.0],
                         "訂正後の値に切り替わっていない")

    def test_p2_as_of_result_is_a_subset(self):
        """P2: as_of を指定した結果は、指定しない結果の部分集合である。
        PITフィルタが「絞る」以外のことをしていないことの確認。"""
        f1 = self.filing("9999", "2026-05-10")
        f2 = self.filing("9999", "2026-08-20")
        self.cum(f1, "9999", "FY2026", 1, "revenue", 100.0, ctx="a")
        self.cum(f2, "9999", "FY2026", 2, "revenue", 200.0, ctx="b")
        all_rows = {(r["period"], r["q_no"]) for r in
                    pit.visible_cum(self.con, "9999", None)}
        sub = {(r["period"], r["q_no"]) for r in
               pit.visible_cum(self.con, "9999", date(2026, 6, 1))}
        self.assertTrue(sub <= all_rows)
        self.assertEqual(sub, {("FY2026", 1)})

    def test_p3_same_day_disclosure_is_hidden_in_strict_mode(self):
        """P3: 発表日当日は strict で見えず live で見える。

        TDnet は引け後発表が多い。当日の値を当日の判定に使うのは
        ルックアヘッド。危険な側を既定にしないため strict を既定にする。
        """
        f = self.filing("9999", "2026-08-20")
        self.cum(f, "9999", "FY2026", 1, "revenue", 100.0)
        d = date(2026, 8, 20)
        self.assertEqual(len(pit.visible_cum(self.con, "9999", d, "strict")), 0)
        self.assertEqual(len(pit.visible_cum(self.con, "9999", d, "live")), 1)

    def test_p4_scorers_never_touch_the_db_directly(self):
        """P4: スコアラーが投影層を経由せずDBを直接叩いていないこと。

        v1.1 §3-2 の「構造的に強制する」を機械化する。別枠では S5 だけが
        この原則を破っていた（v1.1 で修正済み）。本体で同じことを起こさない。
        """
        src = os.path.join(os.path.dirname(os.path.dirname(
            os.path.abspath(__file__))), "signals", "signal_defs.py")
        with open(src, encoding="utf-8") as fh:
            tree = ast.parse(fh.read())
        # 禁じるのは「ファクトの読み出し」。結果の書き込み(store の INSERT)は
        # 投影層の対象外なので、SELECT を含む文字列リテラルだけを探す。
        bad = [n.lineno for n in ast.walk(tree)
               if isinstance(n, ast.Constant) and isinstance(n.value, str)
               and "select " in n.value.lower()]
        self.assertEqual(bad, [],
                         f"signal_defs.py が投影層を経由せずDBを読んでいる(行 {bad})")


class TestAdapter(_Base):
    def test_half_year_values_are_not_passed_as_quarters(self):
        """span_q=2（半期粒度）は既定で渡さない。別枠に対応概念が無く、
        渡すと6ヶ月の値を四半期として扱う。3441 がこれに該当する。"""
        f = self.filing("3441", "2026-03-13", source="edinet")
        self.cum(f, "3441", "FY2026", 2, "revenue", 6747128000.0)
        self.q("3441", "FY2026", 2, "revenue", 6747128000.0, span=2)
        self.q("3441", "FY2026", 1, "revenue", 3000000000.0, span=1)
        rows = adapter.get_pl_series(self.con, "3441")
        self.assertEqual([r["q_no"] for r in rows], [1], "半期値を渡している")
        wide = adapter.get_pl_series(self.con, "3441", allow_span=(1, 2))
        half = [r for r in wide if r["q_no"] == 2][0]
        self.assertTrue(half["is_half"], "is_half が立っていない")
        self.assertEqual(half["span_q"], 2)

    def test_unverified_item_keys_return_empty_not_wrong_numbers(self):
        """写像が未検証の item_key は空を返す。名前が近いだけで意味がずれる
        可能性がある以上、推測で繋ぐと間違った数字が静かに流れる。"""
        f = self.filing("9999", "2026-05-10")
        self.cum(f, "9999", "FY2026", 1, "advances_received", 500.0)
        self.q("9999", "FY2026", 1, "advances_received", 500.0)
        self.assertEqual(adapter.get_bs_series(self.con, "9999", "deposits_received"), [])
        self.assertEqual(adapter.get_bs_series(self.con, "9999", "diluted_shares"), [])
        # verified のものは通る
        self.q("9999", "FY2026", 1, "trade_receivables", 800.0)
        self.assertEqual(len(adapter.get_bs_series(self.con, "9999",
                                                   "accounts_receivable")), 1)

    def test_unit_is_converted_to_millions(self):
        """本体は円、別枠は百万円。投影層で割る。"""
        f = self.filing("9999", "2026-05-10")
        self.cum(f, "9999", "FY2026", 1, "revenue", 6327633000.0)
        self.q("9999", "FY2026", 1, "revenue", 6327633000.0)
        r = adapter.get_pl_series(self.con, "9999")[0]
        self.assertAlmostEqual(r["sales"], 6327.633, places=3)


class TestAdjustmentLayer(_Base):
    def test_adjustment_without_evidence_is_rejected(self):
        """根拠なき調整額を登録できない構造であること（設計 A-2）。
        アプリ層だけでなくスキーマの CHECK でも弾ける。"""
        f = self.filing("3905", "2026-08-14")
        with self.assertRaises(SystemExit):
            AR.record(self.con, code="3905", period="FY2027", q_no=1,
                      item_key="one_time_revenue", amount=5580000000.0,
                      filing_id=f, locator="(セグメント情報等)", note="短い", by="me")
        # スキーマ側でも弾けること（アプリ層を迂回した経路の防御）
        import sqlite3
        with self.assertRaises(sqlite3.IntegrityError):
            self.con.execute(
                "INSERT INTO pl_adjustments (code, period, q_no, item_key, amount, "
                " source_note, source_filing_id, source_locator, confirmed_by, "
                " confirmed_at) VALUES (?,?,?,?,?,?,?,?,?,?)",
                ("3905", "FY2027", 1, "one_time_revenue", 1.0, "みじかい", f,
                 "x", "me", "now"))

    def test_adjustment_flows_into_normalized_sales(self):
        """調整後売上が投影層を通って別枠の計算に届くこと。"""
        f = self.filing("3905", "2026-08-14")
        self.cum(f, "3905", "FY2027", 1, "revenue", 6327633000.0)
        self.q("3905", "FY2027", 1, "revenue", 6327633000.0)
        AR.record(self.con, code="3905", period="FY2027", q_no=1,
                  item_key="one_time_revenue", amount=5580000000.0, filing_id=f,
                  locator="(セグメント情報等) 3.",
                  note="当第1四半期連結累計期間において手数料収入5,580百万円を計上している",
                  by="test")
        pl = adapter.get_pl_series(self.con, "3905")[0]
        adj = adapter.get_adjustments(self.con, "3905")
        self.assertAlmostEqual(pl["sales"], 6327.633, places=3)
        self.assertAlmostEqual(adapter.normalized_sales(pl, adj), 747.633, places=3)

    def test_adjustment_visibility_follows_the_source_filing(self):
        """調整は「引用元の書類が公知になった時点」から見える。登録日を
        基準にすると、同じ as_of でも登録作業の進み具合で結果が変わる。"""
        f = self.filing("3905", "2026-08-14")
        AR.record(self.con, code="3905", period="FY2027", q_no=1,
                  item_key="one_time_revenue", amount=5580000000.0, filing_id=f,
                  locator="(セグメント情報等) 3.",
                  note="当第1四半期連結累計期間において手数料収入5,580百万円を計上している",
                  by="test")
        self.assertEqual(adapter.get_adjustments(self.con, "3905", date(2026, 8, 13)), {})
        self.assertNotEqual(adapter.get_adjustments(self.con, "3905", date(2026, 8, 15)), {})


if __name__ == "__main__":
    unittest.main(verbosity=2)
