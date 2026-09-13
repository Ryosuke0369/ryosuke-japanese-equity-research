"""screener/tests/test_materialize.py — 投影DBの突合（統合引継ぎ書 v1.1 §3-1）。

投影DBは生成物であって正本ではない。だから守るべきは1つ:
**本体を読んだ結果と、投影DBを読んだ結果が食い違わないこと。**
食い違ったまま気づかないと、バックテストは「本体には無い数字」で
評価されることになり、出てくる勝率も期待値も全部意味を失う。

ここで固定するのは3種類:
  1. 行数      —— 投影で行が増えたり減ったりしていないか
  2. 代表値    —— 代表銘柄の代表値が本体と一致するか（単位換算込み）
  3. PIT の生存 —— as_of を動かすと見える範囲が変わるか
"""
from __future__ import annotations

import os
import sqlite3
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))

from screener import common as C
from screener.projection import materialize as M


def _visible_generations(conn, ticker, as_of=None):
    """別枠 data_access.visible_generations と同じSQL。投影DBが別枠の
    PIT解決に耐えるかを、別枠のコードを読み込まずに確かめる。"""
    if as_of is None:
        rows = conn.execute("SELECT period_end, MAX(generation) g FROM filings "
                            "WHERE ticker=? GROUP BY period_end", (ticker,))
    else:
        rows = conn.execute("SELECT period_end, MAX(generation) g FROM filings "
                            "WHERE ticker=? AND filing_date<=? GROUP BY period_end",
                            (ticker, as_of))
    return {r[0]: r[1] for r in rows}


class MaterializeTest(unittest.TestCase):
    def setUp(self):
        fd, self.db = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.con = C.init_db(self.db)
        self.out = self.db + ".projection.db"
        self._seed()

    def tearDown(self):
        self.con.close()
        for p in (self.db, self.db + "-wal", self.db + "-shm", self.out):
            try:
                os.remove(p)
            except OSError:
                pass

    def _seed(self):
        c = self.con
        c.execute("INSERT INTO companies (code, name, market, sector) "
                  "VALUES ('1301','テスト水産','プライム','水産')")
        # universe は決算期末月をタイトルから取る。読める書類を1本置く。
        c.execute("INSERT INTO filings (id, code, date, type, source, subtype, "
                  " doc_id, title) VALUES "
                  "(1,'1301','2024-06-20','有報','edinet','120','S1',"
                  " '有価証券報告書－第10期(2023/04/01－2024/03/31)')")
        # 2つの四半期を、別々の日に開示した2本の書類として置く
        c.execute("INSERT INTO filings (id, code, date, type, source, subtype, "
                  " doc_id, title) VALUES "
                  "(2,'1301','2023-08-10','四半期','edinet','140','S2','q1'),"
                  "(3,'1301','2023-11-10','四半期','edinet','140','S3','q2')")
        for fid, period, q_no, item, val in (
                (2, 'FY2024', 1, 'revenue', 20_000_000_000.0),
                (2, 'FY2024', 1, 'operating_income', 1_000_000_000.0),
                (3, 'FY2024', 2, 'revenue', 21_000_000_000.0),
                (3, 'FY2024', 2, 'operating_income', 1_500_000_000.0)):
            c.execute("INSERT INTO financials_cum (filing_id, code, period, q_no, "
                      " item, value, context_ref) VALUES (?,?,?,?,?,?,'x')",
                      (fid, '1301', period, q_no, item, val))
            c.execute("INSERT INTO financials_q (code, period, q_no, item, value, "
                      " valid_flag, span_q) VALUES (?,?,?,?,?,1,1)",
                      ('1301', period, q_no, item, val))
        # 半期粒度の行は投影に渡らないこと（別枠に対応概念が無い）
        c.execute("INSERT INTO financials_q (code, period, q_no, item, value, "
                  " valid_flag, span_q) VALUES "
                  "('1301','FY2024',4,'revenue',40000000000.0,1,2)")
        for d, close in (("2023-08-09", 100.0), ("2023-08-10", 101.0),
                         ("2023-08-11", 102.0), ("2023-11-10", 110.0)):
            c.execute("INSERT INTO prices (code, date, close, volume, adj_close, "
                      " adj_volume) VALUES (?,?,?,1000,?,1000)", ('1301', d, close, close))
            c.execute("INSERT INTO market_index (date, close, name) "
                      "VALUES (?,?,'TOPIX')", (d, 2000.0))
        c.commit()

    # ---------------------------------------------------------------- 1. 行数
    def test_row_counts_match_the_source(self):
        st = M.build(self.out, src_con=self.con)
        dst = sqlite3.connect(self.out)
        n_q = dst.execute("SELECT COUNT(*) FROM quarterly_standalone").fetchone()[0]
        n_f = dst.execute("SELECT COUNT(*) FROM filings").fetchone()[0]
        n_p = dst.execute("SELECT COUNT(*) FROM daily_prices").fetchone()[0]
        # 本体の (期,四半期) は span_q=1 が2つ、span_q=2 が1つ。渡るのは2つ。
        self.assertEqual(n_q, 2, "半期粒度の行が四半期として混ざっている")
        self.assertEqual(n_f, 2, "filings と quarterly_standalone の数が揃わない")
        self.assertEqual(n_p, 4)
        self.assertEqual(st["quarterly_standalone"], n_q)
        src_px = self.con.execute("SELECT COUNT(*) FROM daily_prices").fetchone()[0]
        self.assertEqual(n_p, src_px, "価格の行数が本体と違う")
        dst.close()

    # -------------------------------------------------------------- 2. 代表値
    def test_representative_values_match_the_source(self):
        M.build(self.out, src_con=self.con)
        dst = sqlite3.connect(self.out)
        row = dst.execute("SELECT sales, operating_profit FROM quarterly_standalone "
                          "WHERE ticker='1301' AND period_end='FY2024-Q1'").fetchone()
        # 本体は円、別枠は百万円
        self.assertAlmostEqual(row[0], 20_000.0, places=6)
        self.assertAlmostEqual(row[1], 1_000.0, places=6)
        px = dst.execute("SELECT close FROM daily_prices "
                         "WHERE ticker='1301' AND date='2023-08-10'").fetchone()[0]
        src = self.con.execute("SELECT adj_close FROM prices "
                               "WHERE code='1301' AND date='2023-08-10'").fetchone()[0]
        self.assertEqual(px, src, "投影した終値が本体の調整後終値と違う")
        fy = dst.execute("SELECT fiscal_year_end FROM universe "
                         "WHERE ticker='1301'").fetchone()[0]
        self.assertEqual(fy, 3, "決算期末月がタイトルから正しく取れていない")
        dst.close()

    # ------------------------------------------------------------- 3. PIT生存
    def test_as_of_changes_what_is_visible(self):
        M.build(self.out, src_con=self.con)
        dst = sqlite3.connect(self.out)
        before = _visible_generations(dst, '1301', '2023-09-01')
        after = _visible_generations(dst, '1301', '2023-12-01')
        allv = _visible_generations(dst, '1301')
        self.assertEqual(set(before), {'FY2024-Q1'},
                         "開示前の四半期が as_of で見えている（前方視）")
        self.assertEqual(set(after), {'FY2024-Q1', 'FY2024-Q2'})
        self.assertEqual(set(allv), {'FY2024-Q1', 'FY2024-Q2'})
        self.assertTrue(set(before) < set(after), "as_of を進めても増えていない")
        dst.close()

    # ------------------------------------------------- 4. 感度分析用のずらし
    def test_shift_days_moves_the_filing_date_by_business_days(self):
        M.build(self.out, shift_days=1, src_con=self.con)
        dst = sqlite3.connect(self.out)
        d = dst.execute("SELECT filing_date FROM filings "
                        "WHERE period_end='FY2024-Q1'").fetchone()[0]
        # 営業日は 08-09, 08-10, 08-11, 11-10。08-10 の +1 営業日は 08-11
        self.assertEqual(d, "2023-08-11")
        dst.close()
        M.build(self.out, shift_days=-1, src_con=self.con)
        dst = sqlite3.connect(self.out)
        d = dst.execute("SELECT filing_date FROM filings "
                        "WHERE period_end='FY2024-Q1'").fetchone()[0]
        self.assertEqual(d, "2023-08-09")
        dst.close()

    def test_invalid_rows_are_projected_as_invalid_not_dropped(self):
        """無効化された四半期は「消す」のではなく is_valid=0 で渡す。
        別枠が「データが無い」と「無効と判定された」を区別できるように。"""
        self.con.execute("UPDATE financials_q SET valid_flag=0, "
                         "invalid_reason='連結範囲変更' WHERE q_no=2")
        self.con.commit()
        M.build(self.out, src_con=self.con)
        dst = sqlite3.connect(self.out)
        r = dst.execute("SELECT is_valid, invalid_reason FROM quarterly_standalone "
                        "WHERE period_end='FY2024-Q2'").fetchone()
        self.assertEqual(r[0], 0)
        self.assertEqual(r[1], '連結範囲変更')
        dst.close()


if __name__ == "__main__":
    unittest.main()
