"""009 — 開示信頼性の2表を足す（追記のみ。既存の006/007は触らない）。

背景（2026-09-02、人間が原文精読して確定）
------------------------------------------
4813 ACCESS は過年度の不適切会計（海外子会社の売上過大計上・先行計上）により
現在も東証の特別注意銘柄。**「計上済み数字＝証拠」という本システムの根本前提が、
取引所によって否定されている会社**である。スコアがいくら高くても、
前提が成り立たない銘柄を主出力に載せてはいけない。

2層にする理由
-------------
`security_flags` は**人が入れる**。取引所の指定は機械判定できる事実ではないし、
誤って外すと危ないほうに倒れる。`disclosure_flags` は**本文から自動検知**する。
継続企業の前提の注記は文言が定型なので拾える。前者は除外、後者は表示のみ
—— 自動検知は誤検知しうるので、候補を消す権限は与えない。
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))

from screener import common as C

DDL = """
-- 人が管理する。取引所の指定など、機械判定に任せてはいけない事実。
CREATE TABLE IF NOT EXISTS security_flags (
    code        TEXT NOT NULL,
    flag_type   TEXT NOT NULL,      -- special_alert / supervision / 等
    since_date  TEXT,               -- 指定日。**不明なら NULL。推測で埋めない**
    until_date  TEXT,               -- 解除日。現行なら NULL
    source_url  TEXT,               -- 一次ソース。**不明なら NULL**
    note        TEXT,
    added_at    TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (code, flag_type, since_date)
);

-- 書類本文から自動検知したフラグ。**表示のみ。除外しない。**
CREATE TABLE IF NOT EXISTS disclosure_flags (
    filing_id          INTEGER PRIMARY KEY,
    code               TEXT NOT NULL,
    period_label       TEXT,
    doc_id             TEXT,
    doc_date           TEXT,
    going_concern      INTEGER DEFAULT 0,
    gc_matched         TEXT,
    accounting_change  INTEGER DEFAULT 0,
    ac_from_notes      INTEGER DEFAULT 0,   -- 注記事項①〜④が「無」以外
    ac_from_narrative  INTEGER DEFAULT 0,   -- 定性情報本文のキーワード
    ac_matched         TEXT,
    computed_at        TEXT DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_disclosure_flags_code
    ON disclosure_flags(code, doc_date);
"""

# 初期データ。**指定日と一次ソースURLは埋めていない。**
# JPX の該当ページ（/listing/market-alerts/ 配下）を 2026-09-02 に取得したが
# 特別注意銘柄の一覧に到達できず（404 / 一覧に 4813 の記載なし）、
# 指定日を一次ソースで確認できなかった。**確認できない日付を書かない。**
# 自社データで確認できた傍証は note に残す。
SEED = [
    ("4813", "special_alert", None, None, None,
     "過年度の不適切会計（海外子会社の売上過大計上・先行計上）による東証の"
     "特別注意銘柄。2026-09-02 に人間が原文精読して確認。"
     "指定日・一次ソースURLは JPX ページに到達できず未確認（要手入力）。"
     "自社データ側の傍証: 2025-06-30 に第39〜41期の四半期報告書を一斉訂正、"
     "2025-07-29 に訂正有価証券報告書。"),
]


def migrate(con):
    con.executescript(DDL)
    n = 0
    for row in SEED:
        cur = con.execute(
            "SELECT COUNT(*) FROM security_flags WHERE code=? AND flag_type=?",
            (row[0], row[1])).fetchone()[0]
        if cur:
            continue
        con.execute(
            "INSERT INTO security_flags (code, flag_type, since_date, until_date,"
            " source_url, note) VALUES (?,?,?,?,?,?)", row)
        n += 1
    con.commit()
    return {"security_flags_seeded": n}


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    con = C.init_db()
    st = migrate(con)
    for k, v in st.items():
        C.log("  %-28s %s" % (k, v))
    C.log("  security_flags 現在: %d 件"
          % con.execute("SELECT COUNT(*) FROM security_flags").fetchone()[0])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
