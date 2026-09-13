r"""screener/tests/test_s12.py — S12 の fixture テスト。

fixture は実データ 5ペア（`$DATA_ROOT\s12_narrative_pairs_20260902.txt`、
EDINET 四半期報告書 1301×3 / 1375×2）。**実テキストで確認した期待値のみ**を
固定する。仕様と実データが食い違ったまま「テストが通った」と言わないため、
期待値は 2026-09-02 に実文面を読んで確定させた。

  ペア1-3 (1301 Q1/Q2/Q3): segment_changed=1
  ペア4-5 の当期 (1375 Q1/Q2): new_product_mention=1 かつ **加点0**
  ペア5 (1375 Q2): forecast_revision_mentioned=1 / direction は本文から判定しない
"""
from __future__ import annotations

import os
import re
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))

from screener import common as C
from screener.extract import s12_narrative as S12

FIXTURE = os.path.join(C.DATA_DIR, "s12_narrative_pairs_20260902.txt")


def load_pairs():
    """fixture ファイルを (前期本文, 当期本文) に割る。"""
    if not os.path.exists(FIXTURE):
        return []
    with open(FIXTURE, encoding="utf-8") as fh:
        txt = fh.read()
    out = []
    for block in re.split(r"#{78}\n# ペア", txt)[1:]:
        head = block.split("\n")[0].strip()
        if "【前期】" not in block or "【当期】" not in block:
            continue
        pri = block.split("【前期】")[1].split("【当期】")[0]
        cur = block.split("【当期】")[1]
        out.append((head, pri, cur))
    return out


@unittest.skipUnless(os.path.exists(FIXTURE), "fixture が無い環境ではスキップ")
class S12FixtureTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pairs = load_pairs()

    def setUp(self):
        fd, self.db = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.con = C.init_db(self.db)

    def tearDown(self):
        self.con.close()
        for s in ("", "-wal", "-shm"):
            try:
                os.remove(self.db + s)
            except OSError:
                pass

    def test_fixture_has_five_pairs(self):
        self.assertEqual(len(self.pairs), 5, "fixture のペア数が想定と違う")

    def test_segment_change_detected_in_pairs_1_to_3(self):
        """1301 の3ペアは報告セグメントの変更を含む。"""
        for i in (0, 1, 2):
            _head, _pri, cur = self.pairs[i]
            flags = S12.detect_flags(S12.sentences(cur))
            self.assertTrue(flags.get("segment_changed"),
                            "ペア%d で segment_changed が立たない" % (i + 1))

    def test_segment_change_disables_tier_b(self):
        """区分が変わった期は Tier B（セグメント方向）を評価しない。"""
        _h, pri, cur = self.pairs[0]
        _a, b_on, _c, _ev = S12.score_pair(cur, pri, segment_changed=False)
        _a2, b_off, _c2, _e2 = S12.score_pair(cur, pri, segment_changed=True)
        self.assertEqual(b_off, 0.0, "segment_changed でも Tier B が効いている")
        self.assertIsInstance(b_on, float)

    def test_new_product_mention_in_pairs_4_and_5(self):
        """ペア4・5の当期に新製品言及がある（「代替肉の開発に成功」）。"""
        for i in (3, 4):
            _head, _pri, cur = self.pairs[i]
            flags = S12.detect_flags(S12.sentences(cur))
            self.assertTrue(flags.get("new_product_mention"),
                            "ペア%d で new_product_mention が立たない" % (i + 1))

    def test_new_product_mention_never_scores(self):
        """新製品言及は **加点しない**。フラグのみ。"""
        for i in (3, 4):
            _head, pri, cur = self.pairs[i]
            _a, _b, _c, ev = S12.score_pair(cur, pri)
            keys = [k for _t, k, _s, _sc, _c in ev]
            self.assertNotIn("new_product_mention", keys,
                             "新製品言及が加点対象に入っている")
            for tier, _k, _s, _sc, _c in ev:
                self.assertIn(tier, ("A", "B", "C"), "Tier D が加点されている")

    def test_forecast_revision_detected_in_pair5(self):
        """ペア5に業績予想の修正言及がある。方向は本文から判定しない。"""
        _head, _pri, cur = self.pairs[4]
        flags = S12.detect_flags(S12.sentences(cur))
        self.assertTrue(flags.get("forecast_revision_mentioned"),
                        "ペア5 で forecast_revision_mentioned が立たない")
        d = S12.revision_direction(self.con, "1375", "2023-11-10")
        self.assertEqual(d, "unknown",
                         "guidance が無いのに方向を決めている（推測している）")

    def test_forecast_revision_dedup(self):
        """修正言及は TDnet 修正開示が既にあるなら 0 点にする（両パターン）。"""
        ev = [("A", "forecast_revision", "業績予想を修正しております。", 0.10)]
        # dedup 対象が無い場合: 点はそのまま
        out = S12.apply_dedup(self.con, "1375", "FY2024-2Q", ev)
        self.assertEqual(out[0][3], 0.10)
        self.assertIsNone(out[0][4])
        # TDnet の業績予想修正がある場合: 0点 + 理由が残る
        self.con.execute(
            "INSERT INTO filings (code, date, type, source, subtype, doc_id) "
            "VALUES ('1375','2023-11-09','修正','tdnet','業績予想修正','X1')")
        self.con.commit()
        out = S12.apply_dedup(self.con, "1375", "FY2024-2Q", ev)
        self.assertEqual(out[0][3], 0.0, "dedup が効いていない")
        self.assertIn("tdnet_revision", out[0][4])
        self.assertEqual(out[0][2], "業績予想を修正しております。",
                         "根拠文まで消してはいけない（点だけ落とす）")

    def test_no_credit_list_blocks_policy_sentences(self):
        """方針表明・中計引用は加点しない。"""
        self.assertTrue(S12._no_credit(
            "中期経営計画に基づき、収益基盤の強化に取り組んでおります。"))
        self.assertTrue(S12._no_credit(
            "当連結会計年度中に最初の製品を発売することを目標に準備を進めております。"))
        self.assertFalse(S12._no_credit(
            "受注が増加し、売上高は前年同期比で伸長しました。"))

    def test_number_fragments_are_joined(self):
        """行分割された数字断片を結合してから正規表現を当てる。"""
        raw = "売上高は\n652\n億\n82\n百万円\nとなりました。"
        joined = S12.join_number_fragments(raw)
        self.assertIn("65282", joined.replace("億", "").replace("百万円", ""))
        self.assertLess(joined.count("\n"), raw.count("\n"))

    def test_composite_is_clamped(self):
        """合成は ±0.30 で clamp される。"""
        cap = S12.cfg()["caps"]["composite"]
        self.assertEqual(cap, 0.30)
        for _h, pri, cur in self.pairs:
            a, b, c, _ev = S12.score_pair(cur, pri)
            self.assertLessEqual(abs(a), S12.cfg()["caps"]["tier_a"] + 1e-9)
            self.assertLessEqual(abs(b), S12.cfg()["caps"]["tier_b"] + 1e-9)
            self.assertLessEqual(abs(c), S12.cfg()["caps"]["tier_c"] + 1e-9)

    def test_every_score_has_evidence(self):
        """根拠文の無い点数は存在してはならない。"""
        for _h, pri, cur in self.pairs:
            _a, _b, _c, ev = S12.score_pair(cur, pri)
            for tier, key, text, sc, _cls in ev:
                self.assertTrue(text and text.strip(),
                                "点が付いているのに根拠文が空: %s/%s" % (tier, key))


if __name__ == "__main__":
    unittest.main()
