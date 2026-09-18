"""screener/report/s13_series.py — S13 受注の四半期単独推移（2026-09-18）。

なぜ要るのか
------------
S13 は §22 / §29 で「可用率が低い（29.1%）ので合成スコアには接続しない」と
決めた。**接続しないことと、見せないことは別である。** 現状は 0点扱いの
うえ週次レポートに残高 YoY を1行出すだけで、取得できている銘柄でも
四半期ごとの動きが人間の目に入らない。

実例（6838 多摩川HD・2026年10月期）: B/B 比は Q1 0.66 → Q2 1.74 → Q3 0.93。
Q2 の 1.74 を見て「売上減速は期ズレ」と読めるが、Q3 で 0.93 に戻ることで
「単発だった」と分かる。**この推移が判断を決めた。**1点の YoY だけでは
この読み替えができない以上、非表示は情報の損失である。

何を組み立てるか
----------------
`s13_orders` は書類1本につき1行で、**その書類が語る期間の累計**（受注高）と
**期末残高**（受注残高・ストック）を持つ。四半期単独を出すには、同じ年度の
中で累計を差分する —— `quarterly_builder` が売上に対してやっているのと
同じ操作を受注高に対して行う。

  受注高(単独)   = 当期までの累計 − 直前開示までの累計   （フロー・差分する）
  受注残高       = 期末残高そのまま                      （ストック・差分しない）
  Book-to-Bill   = 受注高(単独) ÷ 売上(単独)             （同じ期間同士でのみ）

**span が違うものを割らない。** 6ヶ月ぶんの受注高を3ヶ月ぶんの売上で割ると
比が倍になる。どちらも「1四半期あたり」に正規化してから割り、span は
必ず表示に出す。
"""
from __future__ import annotations

import os
import re
import sqlite3
import sys

try:
    from screener import common as C            # noqa: F401
except ImportError:                             # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C            # noqa: F401

from screener.signals import sales_direction as SD

TREND_N = 4                     # 何四半期ぶん並べるか


def period_of_filing(mcon, filing_id, title=None):
    """その書類が語っている期 (fiscal_year, q_no)。分からなければ None。

    `financials_cum` は同じ書類から前年同期の比較値も取り込むので、
    **最大の期**がその書類の報告期になる。取り込めていない書類は
    タイトルの「第N四半期」から拾う（有報・半期報告書は Q4 / Q2）。
    """
    try:
        r = mcon.execute(
            "SELECT period, q_no FROM financials_cum WHERE filing_id=? "
            "AND period IS NOT NULL ORDER BY period DESC, q_no DESC LIMIT 1",
            (filing_id,)).fetchone()
    except sqlite3.OperationalError:
        r = None
    if r and r[0]:
        m = re.match(r"FY(\d{4})", r[0])
        if m:
            return (int(m.group(1)), int(r[1] or 4))
    t = title or ""
    m = re.search(r"第(\d)四半期", t)
    y = re.search(r"(\d{4})/\d\d/\d\d\s*[-－~〜]\s*(\d{4})/(\d\d)/\d\d", t)
    if m and y:
        return (int(y.group(2)), int(m.group(1)))
    if y:
        return (int(y.group(2)), 2 if "半期報告書" in t else 4)
    return None


def _orders_rows(mcon, code, as_of):
    """s13_orders の available 行に報告期を付けて返す。昇順。"""
    try:
        rows = mcon.execute(
            "SELECT o.filing_id, o.doc_id, o.doc_date, o.doc_source, o.orders_amount,"
            " o.closing_backlog, o.opening_backlog, o.completed_amount,"
            " o.backlog_yoy_pct, o.orders_yoy_pct, f.title "
            "FROM s13_orders o LEFT JOIN filings f ON f.id=o.filing_id "
            "WHERE o.code=? AND o.available=1 AND o.doc_date<=? "
            "ORDER BY o.doc_date", (code, as_of)).fetchall()
    except sqlite3.OperationalError:
        return []
    out = []
    for r in rows:
        pq = period_of_filing(mcon, r[0], r[10])
        if not pq:
            continue
        out.append({"fy": pq[0], "q": pq[1], "period_end": "FY%d-Q%d" % pq,
                    "doc_id": r[1], "doc_date": r[2], "doc_source": r[3],
                    "cum_orders": r[4], "closing_backlog": r[5],
                    "backlog_yoy_pct": r[8], "orders_yoy_pct": r[9]})
    out.sort(key=lambda x: (x["fy"], x["q"], x["doc_date"]))
    # 同じ期を語る書類が複数（短信→有報）あれば後から出たほうを採る。
    dedup = {}
    for r in out:
        dedup[(r["fy"], r["q"])] = r
    return [dedup[k] for k in sorted(dedup)]


def quarterly_series(mcon, pcon, code, as_of, n=TREND_N):
    """受注高(単独)・受注残高・B/B の四半期推移。

    戻り値: {"rows": [...], "text": "...", "note": "..."}。
    rows の各要素:
      period_end / span_q / orders (その span の受注高) / orders_per_q /
      closing_backlog / sales / sales_per_q / bb (Book-to-Bill) / basis
    """
    src = _orders_rows(mcon, code, as_of)
    if not src:
        return {"rows": [], "text": "", "note": "受注実績の節を持つ開示が無い"}

    # 売上タイル（期 → 1四半期あたり）。span を合わせて割るために使う。
    sales = {}
    try:
        for t in SD.tiles(pcon, code):
            sales[t["period_end"]] = t
    except sqlite3.OperationalError:
        pass

    rows, prev = [], {}
    for r in src:
        fy, q = r["fy"], r["q"]
        cum = r["cum_orders"]
        p = prev.get(fy)
        orders = span = None
        if cum is not None:
            if p and p["cum"] is not None and p["q"] < q:
                orders, span = cum - p["cum"], q - p["q"]
                basis = "累計差分(%s→%s)" % (p["period_end"], r["period_end"])
            else:
                orders, span = cum, q
                basis = "期首からの累計(Q1〜Q%d)" % q
            prev[fy] = {"cum": cum, "q": q, "period_end": r["period_end"]}
        else:
            basis = "受注高の記載なし（残高のみ）"
        st = sales.get(r["period_end"])
        s_per_q = st["run_rate"] if st else None
        o_per_q = (orders / span) if (orders is not None and span) else None
        bb = None
        if o_per_q is not None and s_per_q and s_per_q > 0:
            bb = round(o_per_q / s_per_q, 2)
        rows.append({"period_end": r["period_end"], "span_q": span,
                     "orders": None if orders is None else round(orders, 1),
                     "orders_per_q": None if o_per_q is None else round(o_per_q, 1),
                     "closing_backlog": (None if r["closing_backlog"] is None
                                         else round(r["closing_backlog"], 1)),
                     "sales": None if not st else round(st["value"], 1),
                     "sales_span_q": None if not st else st["span_q"],
                     "sales_per_q": None if s_per_q is None else round(s_per_q, 1),
                     "bb": bb, "basis": basis,
                     "doc_date": r["doc_date"], "doc_id": r["doc_id"],
                     "backlog_yoy_pct": r["backlog_yoy_pct"],
                     "orders_yoy_pct": r["orders_yoy_pct"]})
    rows = rows[-n:]
    parts = []
    for r in rows:
        seg = r["period_end"]
        if r["span_q"] and r["span_q"] != 1:
            seg += "(span=%d)" % r["span_q"]
        seg += " 受注 %s / 残高 %s / B-B %s" % (
            "-" if r["orders"] is None else "%.0f" % r["orders"],
            "-" if r["closing_backlog"] is None else "%.0f" % r["closing_backlog"],
            "-" if r["bb"] is None else "%.2f" % r["bb"])
        parts.append(seg)
    note = ""
    if all(r["bb"] is None for r in rows):
        note = "B/B は売上の同期間タイルが取れないため未算出"
    return {"rows": rows, "text": " ｜ ".join(parts), "note": note}
