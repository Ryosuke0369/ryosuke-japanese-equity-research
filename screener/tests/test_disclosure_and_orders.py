"""tests — 開示信頼性・会計処理変更・受注残（タスク2/3/4/5）の fixture。

すべて 2026-09-02 に人間が原文精読して確定した実測値に基づく。
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))

from screener.extract.disclosure_flags import (          # noqa: E402
    detect_accounting_change, detect_going_concern)
from screener.extract.order_backlog import parse_orders, yoy   # noqa: E402
from screener.extract.s12_narrative import classify_paragraph  # noqa: E402


class TestGoingConcern(unittest.TestCase):
    def test_該当事項はありませんは0(self):
        t = "（継続企業の前提に関する注記）\n該当事項はありません。"
        self.assertEqual(detect_going_concern(t)[0], 0)

    def test_重要な疑義は1(self):
        t = ("（継続企業の前提に関する注記）\n"
             "当社グループには、継続企業の前提に関する重要な疑義が存在します。")
        f, m = detect_going_concern(t)
        self.assertEqual(f, 1)
        self.assertIn("重要な疑義", m)

    def test_不確実性は認められないは0(self):
        """**否定文を陽性にしない。** 4813 の但し書きがこの形。"""
        t = ("継続企業の前提に関する注記\n"
             "重要な不確実性は認められないと判断しております。")
        self.assertEqual(detect_going_concern(t)[0], 0)

    def test_節が無ければ0(self):
        self.assertEqual(detect_going_concern("普通の業績説明のみ。")[0], 0)


class TestAccountingChange(unittest.TestCase):
    NOTES_ALL_NONE = ("(4) 会計方針の変更・会計上の見積りの変更・修正再表示\n"
                      "① 会計基準等の改正に伴う会計方針の変更：無\n"
                      "② 会計方針の変更：無\n③ 会計上の見積りの変更：無\n"
                      "④ 修正再表示：無\n")

    def test_アセンテックQ1の構造(self):
        """注記は全て「無」。変更は定性情報の説明文にしか出ていない。
        **注記だけを見る実装では捕まらない**ことを固定する。"""
        t = (self.NOTES_ALL_NONE + "\n1.経営成績に関する説明\n"
             "一部取引について代理人と判断し、純額処理へ変更したことにより"
             "売上高が677百万円減少しております。")
        f, notes, nar, matched = detect_accounting_change(t)
        self.assertEqual((f, notes, nar), (1, 0, 1))
        self.assertIn("純額処理", matched)

    def test_見出し行を項目と誤認しない(self):
        """「(4) 会計方針の変更・会計上の見積りの変更・修正再表示」は
        項目名を並べた見出しで、値を持たない。"""
        f, notes, nar, _ = detect_accounting_change(self.NOTES_ALL_NONE)
        self.assertEqual((f, notes, nar), (0, 0, 0))

    def test_注記が有なら注記由来で立つ(self):
        t = self.NOTES_ALL_NONE.replace("② 会計方針の変更：無",
                                        "② 会計方針の変更：有")
        self.assertEqual(detect_accounting_change(t)[1], 1)

    def test_定型文だけでは立てない(self):
        """「収益認識に関する会計基準の適用」は2021年以降ほぼ全社に載る。
        語があるだけで拾うと 42% が陽性になり印にならなかった。"""
        t = "当社は収益認識に関する会計基準を適用しております。表示方法の変更"
        self.assertEqual(detect_accounting_change(t)[0], 0)


class TestMacroParagraph(unittest.TestCase):
    """S12 マクロ定型文の除外（タスク3）。"""

    BESTELLA = ("当第１四半期連結累計期間におけるわが国経済は、雇用・所得環境の"
                "改善や各種政策の効果により緩やかな回復基調を維持しております。")

    def test_マクロ定型文はmacro(self):
        self.assertEqual(classify_paragraph(self.BESTELLA), "macro")

    def test_同じ語でも会社事業段落なら採る(self):
        t = "当社の主力製品の受注が回復し、売上高は前年同期を上回りました。"
        self.assertEqual(classify_paragraph(t), "company")

    def test_業界環境段落は残す(self):
        """**マクロと業界の区別がこの修正の核心。** 業界まで落とすと
        シグナルの中身が痩せるだけになる。"""
        t = ("当社グループの属する解体・メンテナンス業界では、"
             "老朽化設備の更新需要が堅調に推移しております。")
        self.assertEqual(classify_paragraph(t), "industry")

    def test_マクロの2文目も引き継ぐ(self):
        t = "一方、米国の保護主義的な通商政策の再強化は、逆風となる可能性があります。"
        self.assertEqual(classify_paragraph(t), "macro")


class TestOrderBacklog(unittest.TestCase):
    """S13 受注残（タスク5）。1433 ベステラ FY2027-Q1 短信の実数値。"""

    SHINSAI = ("ａ　受注実績\n項目\n金額(千円)\n金額(千円)\n"
               "前期繰越工事高\n7,197,382\n8,512,120\n"
               "当期受注工事高\n1,339,745\n3,865,046\n"
               "当期完成工事高\n2,432,692\n3,235,533\n"
               "次期繰越工事高\n6,104,435\n9,141,633\n")

    def test_4数値のパース(self):
        d = parse_orders(self.SHINSAI)
        self.assertIsNotNone(d)
        # 千円 → 百万円に揃える
        self.assertAlmostEqual(d["opening_backlog"]["current"], 8512.120, places=2)
        self.assertAlmostEqual(d["orders"]["current"], 3865.046, places=2)
        self.assertAlmostEqual(d["completed"]["current"], 3235.533, places=2)
        self.assertAlmostEqual(d["closing_backlog"]["current"], 9141.633, places=2)

    def test_受注残YoYの再計算(self):
        """+49.8% が再現できること（受入条件）。"""
        d = parse_orders(self.SHINSAI)
        self.assertEqual(yoy(d, "closing_backlog"), 49.8)
        self.assertEqual(yoy(d, "orders"), 188.5)
        self.assertEqual(yoy(d, "completed"), 33.0)
        self.assertEqual(yoy(d, "opening_backlog"), 18.3)

    def test_セグメント表は合計行を採る(self):
        """1行目はセグメント1つ分。会社全体として扱うと桁も意味も違う。"""
        t = ("ｂ．受注実績\nセグメントの名称\n受注高（百万円）\n前年同期比（％）\n"
             "受注残高（百万円）\n前年同期比（％）\n"
             "電子機器部品製造装置\n4,219\n96.9\n2,556\n79.5\n"
             "合計\n15,011\n106.1\n3,538\n84.7\n")
        d = parse_orders(t)
        self.assertAlmostEqual(d["orders"]["current"], 15011.0)
        self.assertAlmostEqual(d["closing_backlog"]["current"], 3538.0)
        self.assertEqual(yoy(d, "closing_backlog"), -15.3)

    def test_節が無ければNone(self):
        """**n/a であって 0 ではない。**"""
        self.assertIsNone(parse_orders("普通の経営成績の説明のみ。"))

    def test_年号を金額と読まない(self):
        t = "ａ　受注実績\n自 2025年２月１日\n至 2026年１月31日\n受注残高\n2,556\n79.5\n"
        d = parse_orders(t)
        self.assertNotIn(d and d.get("closing_backlog", {}).get("current"),
                         (2025.0, 2026.0))


class TestExclusionScope(unittest.TestCase):
    """除外は「数字の信頼性が否定された種別」だけ（タスクC）。"""

    def test_特別注意は除外する(self):
        from screener.report.weekly_screen import excludes
        self.assertTrue(excludes([("special_alert", "2025-08-27", "")]))

    def test_上場維持基準の未適合は除外しない(self):
        """上場継続性の話であって、計上済み数字が疑わしいという話ではない。
        除外の範囲を広げると本来見るべき候補が静かに消える。"""
        from screener.report.weekly_screen import excludes
        self.assertFalse(excludes([("listing_maintenance", "2026-04-30", "")]))

    def test_両方あれば除外する(self):
        from screener.report.weekly_screen import excludes
        self.assertTrue(excludes([("listing_maintenance", "2026-04-30", ""),
                                  ("special_alert", "2025-08-27", "")]))

    def test_フラグ無しは除外しない(self):
        from screener.report.weekly_screen import excludes
        self.assertFalse(excludes(None))
        self.assertFalse(excludes([]))
