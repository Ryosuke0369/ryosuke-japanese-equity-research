"""screener/projection/pit.py — Point-in-Time 可視性の解決 (統合引継ぎ書 v1.1 §3-2)。

「その日時点で公知だった値だけを返す」を1か所で実装する。バックテストが
未来を見ていたら、出てくる数字は全部嘘になる。ここが生命線。

## なぜ generation 列を作らないか

別枠(Kimi側)は訂正世代を保持するために全ファクト表へ generation 列を足した。
本体は不要 —— financials_cum の主キーが (filing_id, item, context_ref) で
**書類ごとにファクトを保持している**ので、訂正短信は自動的に別行になる。
「as_of 時点で最新だった版」は filings.date の降順で1件選べば決まる。
v1.1 §0-1 自身が「実運用のファクト層は本体構造を正とする」としている。

## 引け後発表

TDnet は引け後発表が多い。発表日「当日」の判定に当日発表の値を使うと
ルックアヘッドになる。既定は strict（当日発表を見ない）。危険な側を既定に
しないため、live は明示的に選ばせる。

    strict : filings.date <  as_of   バックテスト既定
    live   : filings.date <= as_of   実運用（寄りまでに発表済みを見る）
"""
from __future__ import annotations

import os
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

MODES = ("strict", "live")


def _as_of_str(as_of) -> str | None:
    if as_of is None:
        return None
    return as_of.isoformat() if hasattr(as_of, "isoformat") else str(as_of)


def cutoff_op(mode: str) -> str:
    if mode not in MODES:
        raise ValueError(f"mode は {MODES} のいずれか: {mode!r}")
    return "<" if mode == "strict" else "<="


def visible_cum(con, code: str, as_of=None, mode: str = "strict"):
    """financials_cum のうち as_of 時点で見えていた行。訂正は最新版が勝つ。

    ROW_NUMBER() で (code, period, q_no, item, context_ref) ごとに
    filings.date の降順1件を採る。同日に複数書類があるときは id の降順
    （後に登録された方＝訂正）を採る。
    """
    ad = _as_of_str(as_of)
    where = "" if ad is None else f" AND f.date {cutoff_op(mode)} ?"
    args = (code,) if ad is None else (code, ad)
    return con.execute(
        "WITH visible AS ("
        "  SELECT fc.code, fc.period, fc.q_no, fc.item, fc.context_ref, fc.value,"
        "         fc.source_tag, f.date AS filing_date, f.id AS filing_id,"
        "         ROW_NUMBER() OVER ("
        "           PARTITION BY fc.code, fc.period, fc.q_no, fc.item, fc.context_ref"
        "           ORDER BY f.date DESC, f.id DESC) AS rn"
        "  FROM financials_cum fc JOIN filings f ON f.id = fc.filing_id"
        f"  WHERE fc.code = ?{where}"
        ") SELECT * FROM visible WHERE rn = 1 ORDER BY period, q_no, item", args
    ).fetchall()


def visible_q(con, code: str, as_of=None, mode: str = "strict",
              allow_span=(1,), valid_only: bool = True):
    """financials_q（単独値）のうち as_of 時点で見えていた行。

    financials_q は書類ではなく「銘柄×期×四半期×項目」の粒度なので、
    どの書類までを使って作られたかを行自身は持っていない。そこで
    「その (period, q_no) を語る書類が as_of までに出ていたか」を
    financials_cum 側で確認して絞る —— 単独値は累計から作られるので、
    累計が公知でなければ単独値も公知ではない。
    """
    ad = _as_of_str(as_of)
    rows = con.execute(
        "SELECT code, period, q_no, item, value, valid_flag, invalid_reason, span_q "
        "FROM financials_q WHERE code=?" + (" AND valid_flag=1" if valid_only else ""),
        (code,)).fetchall()
    rows = [r for r in rows if (r["span_q"] or 1) in allow_span]
    if ad is None:
        return rows
    seen = {(r["period"], r["q_no"]) for r in con.execute(
        "SELECT DISTINCT fc.period, fc.q_no FROM financials_cum fc "
        "JOIN filings f ON f.id = fc.filing_id "
        f"WHERE fc.code=? AND f.date {cutoff_op(mode)} ?", (code, ad))}
    return [r for r in rows if (r["period"], r["q_no"]) in seen]


def visible_dim(con, code: str, item: str, as_of=None, mode: str = "strict",
                axis: str = "segment"):
    """financials_dim（次元付き）のうち as_of 時点で見えていた行。"""
    ad = _as_of_str(as_of)
    where = "" if ad is None else f" AND f.date {cutoff_op(mode)} ?"
    args = (code, item, axis) if ad is None else (code, item, axis, ad)
    return con.execute(
        "WITH visible AS ("
        "  SELECT fd.code, fd.period, fd.q_no, fd.item, fd.member, fd.value,"
        "         fd.valid_flag, f.date AS filing_date,"
        "         ROW_NUMBER() OVER ("
        "           PARTITION BY fd.code, fd.period, fd.q_no, fd.item, fd.member"
        "           ORDER BY f.date DESC, f.id DESC) AS rn"
        "  FROM financials_dim fd JOIN filings f ON f.id = fd.filing_id"
        f"  WHERE fd.code=? AND fd.item=? AND fd.axis=?{where}"
        ") SELECT * FROM visible WHERE rn = 1 ORDER BY period, q_no, member", args
    ).fetchall()


def visible_guidance(con, code: str, item: str = "operating_income",
                     as_of=None, mode: str = "strict"):
    """会社予想の最新版。as_of 指定でその時点の版を再現する。"""
    ad = _as_of_str(as_of)
    where = "" if ad is None else f" AND date {cutoff_op(mode)} ?"
    args = (code, item) if ad is None else (code, item, ad)
    return con.execute(
        f"SELECT * FROM guidance WHERE code=? AND item=?{where} "
        "ORDER BY date DESC LIMIT 1", args).fetchone()


def visible_adjustments(con, code: str, as_of=None, mode: str = "strict"):
    """一時収入の調整。(period, q_no) → 控除合計。

    調整は「引用元の書類が公知になった時点」から見える。注記を読んで
    登録するのは後日だが、**情報自体は発表日に公知**なので基準は
    source_filing_id の書類日付にする。登録日を基準にすると、同じ
    as_of でも登録作業の進み具合で結果が変わってしまう。
    """
    ad = _as_of_str(as_of)
    where = "" if ad is None else f" AND f.date {cutoff_op(mode)} ?"
    args = (code,) if ad is None else (code, ad)
    out: dict = {}
    for r in con.execute(
            "SELECT a.period, a.q_no, a.item_key, a.amount FROM pl_adjustments a "
            "JOIN filings f ON f.id = a.source_filing_id "
            f"WHERE a.code=?{where}", args):
        out.setdefault((r["period"], r["q_no"]), {})[r["item_key"]] = r["amount"]
    return out
