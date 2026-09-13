"""screener/fetch/tdnet_titles.py — 非決算の適時開示を**タイトルだけ**記録する。

なぜタイトルだけか
------------------
2026-09-02 の実測（9営業日、TDnet の一覧ページ）:

    全開示 1,998件 / 1日平均 222件
      決算系  308 (15.4%)  ← 既存の tdnet_archiver が本文まで取っている
      非決算 1,690 (84.6%) ← 年換算 約46,000件
    リスク語ヒット 4件（全体の 0.20%）→ 年換算 約108件/年

**46,000件のPDFを落とす必要はない。** テールリスクの検知に要るのは
タイトルだけで、タイトルは日次の一覧ページに全部載っている。
本文を取りに行くと保存量も失敗率も跳ね上がるうえ、拾いたい情報は
1件も増えない。

なぜ既存の archiver を触らないか
--------------------------------
`tdnet_archiver` は「本文を取って filings に入れる」責務で、`filings` は
財務パースの入口でもある。非決算の開示をそこに混ぜると、財務パースの
母集団が汚れる（`xbrl_path` が無い行が大量に増える）。別表に分ける。

保持の制約
----------
TDnet の一覧は**約1か月しか遡れない**（2026-09-02 時点で 2026-07-23 は
すでに 404）。過去に遡って埋めることはできないので、**日次で拾い続ける
ことだけが唯一の手段**。取りこぼした日は永久に空く。
"""
from __future__ import annotations

import argparse
import os
import sqlite3
import sys
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.fetch import tdnet_archiver as T

DDL = """
CREATE TABLE IF NOT EXISTS disclosure_titles (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    code        TEXT,
    date        TEXT NOT NULL,
    time        TEXT,
    title       TEXT NOT NULL,
    doc_name    TEXT,          -- TDnet の PDF ファイル名（本文は落とさない）
    subtype     TEXT,          -- 決算系ならその種別、それ以外は NULL
    risk_hit    INTEGER DEFAULT 0,
    fetched_at  TEXT DEFAULT (datetime('now')),
    UNIQUE(date, code, title)
);
CREATE INDEX IF NOT EXISTS idx_disc_titles_code ON disclosure_titles(code, date);
CREATE INDEX IF NOT EXISTS idx_disc_titles_risk ON disclosure_titles(risk_hit, date);
"""


def _risk_re():
    from screener.report.weekly_screen import RISK_TITLE
    return RISK_TITLE


def archive_titles(con, fetcher, d: date) -> dict:
    """1日ぶんの一覧を読み、全開示のタイトルを記録する。本文は落とさない。"""
    rx = _risk_re()
    rows, total = T.fetch_day_index(fetcher, d)
    n_new = n_risk = 0
    for r in rows:
        st, _ = T.classify(r["title"])
        hit = 1 if (st is None and rx.search(r["title"] or "")) else 0
        # **一覧の行が持つキーは code_raw / pdf / name。** 2026-09-13 まで
        # r.get("code") を読んでいたため全行 code=NULL、doc_name には会社名が
        # 入っていた（キーが無くても例外にならない型の故障）。
        cur = con.execute(
            "INSERT OR IGNORE INTO disclosure_titles (code, date, time, title,"
            " doc_name, subtype, risk_hit) VALUES (?,?,?,?,?,?,?)",
            (C.normalise_code(r.get("code_raw")), d.isoformat(), r.get("time"),
             r["title"], r.get("pdf") or r.get("zip"), st, hit))
        if cur.rowcount:
            n_new += 1
            n_risk += hit
    con.commit()
    return {"date": d.isoformat(), "listed": total, "stored": n_new,
            "risk": n_risk}


def rederive_codes(con, fetcher, dates=None) -> dict:
    """code=NULL で保存された既存行の code / doc_name を一覧から引き直す。

    照合キーは (date, time, title, 会社名)。旧実装は doc_name に会社名を
    入れていたので、それを会社名として使える。同じ時刻・同じタイトルの
    開示は複数社にまたがる（「業績予想の修正に関するお知らせ」等）ので、
    会社名まで合わないものは**埋めない**（推測で code を振らない）。
    一覧が遡れない日（約1ヶ月超）は unreachable として数える。
    """
    if dates is None:
        dates = [r[0] for r in con.execute(
            "SELECT DISTINCT date FROM disclosure_titles WHERE code IS NULL ORDER BY date")]
    st = {"dates": 0, "updated": 0, "unmatched": 0, "ambiguous": 0,
          "no_code_listing": 0, "unreachable": 0}
    for ds in dates:
        d = date.fromisoformat(ds)
        rows, total = T.fetch_day_index(fetcher, d)
        if not rows:
            st["unreachable"] += 1
            continue
        st["dates"] += 1
        idx = {}
        for r in rows:
            idx.setdefault((r.get("time"), r["title"], r.get("name")), []).append(r)
        for x in con.execute("SELECT id, time, title, doc_name FROM disclosure_titles "
                             "WHERE date=? AND code IS NULL", (ds,)).fetchall():
            hits = idx.get((x["time"], x["title"], x["doc_name"]), [])
            if not hits:
                st["unmatched"] += 1
                continue
            codes = {C.normalise_code(h.get("code_raw")) for h in hits}
            if len(codes) != 1:
                st["ambiguous"] += 1
                continue
            code = codes.pop()
            if not code:
                st["no_code_listing"] += 1
                continue
            try:
                con.execute("UPDATE disclosure_titles SET code=?, doc_name=? WHERE id=?",
                            (code, hits[0].get("pdf") or hits[0].get("zip"), x["id"]))
                st["updated"] += 1
            except sqlite3.IntegrityError:
                # 修正後の実装で同じ (date, code, title) が既に入っている → 旧行は重複
                con.execute("DELETE FROM disclosure_titles WHERE id=?", (x["id"],))
        con.commit()
    return st


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--rederive-codes", action="store_true",
                   help="code=NULL の既存行を一覧から引き直す（2026-09-13 の不具合修復）")
    p.add_argument("--days", type=int, default=1, help="今日から遡る営業日数")
    p.add_argument("--from", dest="frm", help="開始日 YYYY-MM-DD")
    p.add_argument("--to", dest="to", help="終了日 YYYY-MM-DD")
    p.add_argument("--report", action="store_true")
    p.add_argument("--lock-wait", type=float, default=3600.0)
    a = p.parse_args(argv)

    con = C.init_db()
    con.executescript(DDL)
    if a.report:
        n = con.execute("SELECT COUNT(*) FROM disclosure_titles").fetchone()[0]
        rng = con.execute("SELECT MIN(date), MAX(date) FROM disclosure_titles").fetchone()
        C.log("=== 適時開示タイトル ===")
        C.log("  %d件  %s 〜 %s" % (n, rng[0], rng[1]))
        C.log("  リスク語ヒット %d件"
              % con.execute("SELECT COUNT(*) FROM disclosure_titles "
                            "WHERE risk_hit=1").fetchone()[0])
        for r in con.execute("SELECT date, code, title FROM disclosure_titles "
                             "WHERE risk_hit=1 ORDER BY date DESC LIMIT 10"):
            C.log("    %s %s %s" % (r[0], r[1] or "-", (r[2] or "")[:60]))
        return 0

    if a.rederive_codes:
        fetcher = C.Fetcher()
        try:
            with C.writer_lock("tdnet_titles", wait_seconds=a.lock_wait):
                st = rederive_codes(con, fetcher)
        except C.WriterBusy as e:
            C.log("  ! 他の書き込みジョブが実行中: %s" % e)
            return 2
        C.log("  code 再導出: %s" % st)
        C.log("  残る code=NULL: %d 行" % con.execute(
            "SELECT COUNT(*) FROM disclosure_titles WHERE code IS NULL").fetchone()[0])
        return 0

    if a.frm:
        start = date.fromisoformat(a.frm)
        end = date.fromisoformat(a.to) if a.to else date.today()
        days = [start + timedelta(days=i) for i in range((end - start).days + 1)]
    else:
        days = [date.today() - timedelta(days=i) for i in range(a.days)]

    fetcher = C.Fetcher() if hasattr(C, "Fetcher") else None
    st = {"listed": 0, "stored": 0, "risk": 0, "days": 0, "failed": 0}
    try:
        with C.writer_lock("tdnet_titles", wait_seconds=a.lock_wait):
            for d in sorted(days):
                try:
                    r = archive_titles(con, fetcher, d)
                except Exception as e:
                    st["failed"] += 1
                    C.log("  %s 取得失敗: %s: %s" % (d, type(e).__name__, e))
                    continue
                st["days"] += 1
                for k in ("listed", "stored", "risk"):
                    st[k] += r[k]
                C.log("  %s 一覧%d件 / 新規%d件 / リスク語%d件"
                      % (r["date"], r["listed"], r["stored"], r["risk"]))
    except C.WriterBusy as e:
        C.log("  ! 他の書き込みジョブが実行中: %s" % e)
        return 2
    C.log("  合計: %d営業日 / 一覧%d / 保存%d / リスク語%d / 失敗%d"
          % (st["days"], st["listed"], st["stored"], st["risk"], st["failed"]))
    # **静黙縮退の検出。** 一覧が読めているのに1件も保存できないなら壊れている。
    if st["days"] and st["listed"] and st["stored"] == 0:
        C.log("  ! 一覧は読めたが1件も保存できていない")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
