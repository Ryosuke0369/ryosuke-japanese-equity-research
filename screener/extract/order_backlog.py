"""screener/extract/order_backlog.py — S13 受注残高（シャドウ・0点）。

なぜ受注残なのか
----------------
本システムの哲学は「証拠＝計上済み数字・**履行義務のある実契約**のみ」。
受注残高はまさに履行義務のある実契約の残高で、**哲学上の完全な証拠**
でありながらスコアの外にあった。2026-09-02 に人間が原文精読して発見。

    1433 ベステラ FY2027-Q1「3.その他 生産、受注及び販売の状況」
      次期繰越工事高 6,104,435 → 9,141,633千円 (+49.8%)
    受注残 91.4億 = 通期売上計画 130億の 70% を Q1末に保有。

完成工事基準の業種（建設・プラント・受注生産型製造）では契約負債(S2)が
ほとんど立たない。**S2 と S13 は業種でカバレッジを補完する関係**にある。

シャドウである
--------------
計算・記録・表示だけ。composite には**繋がない**。繋ぐなら評価してから
新規に事前登録する（規則Bの手続きと同じ）。
"""
from __future__ import annotations

import os
import re
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

# セクションの入口。短信は「生産、受注及び販売の状況」、有報は「〜の実績」。
_SECTION = re.compile(
    r"生産[、,．・]?\s*受注(及び|および)?\s*販売の(状況|実績)"
    r"|受注(実績|状況|の状況)"
    r"|受注(工事)?高(及び|および)受注残高")

# 表記ゆれ。工事高／受注高／生産高／販売高、受注残高／繰越高。
# **「次期繰越」と「当期末残高」は同じものを指す**ので同じキーに寄せる。
_LABELS = [
    ("opening_backlog", r"(前期|期首)繰越(工事|受注)?高|期首(受注)?残高"),
    ("orders",          r"当(期|四半期|中間)受注(工事|)高|受注(工事)?高(?!.*残)"),
    ("completed",       r"当(期|四半期|中間)完成(工事)?高|完成工事高|売上高"),
    ("closing_backlog", r"(次期|翌期)繰越(工事|受注)?高|(当期末|期末)(受注)?残高"
                        r"|受注残高"),
]
_LABEL_RE = [(k, re.compile(p)) for k, p in _LABELS]

_NUM = re.compile(r"[0-9][0-9,，]*(?:\.[0-9]+)?")


def _to_f(tok):
    try:
        return float(tok.replace(",", "").replace("，", ""))
    except ValueError:
        return None


# 「生産、受注及び販売」の節は**3つの小節**（生産実績・受注実績・販売実績）を
# 並べて書く。節ごと切り出すと最初の「合計」が生産実績のものになり、
# 生産高を受注高として拾ってしまう（1433 で販売合計、6336 で生産合計を
# 拾っていた。2026-09-02）。**受注の小節だけに絞る。**
_ORDER_HEAD = re.compile(
    r"[ａ-ｚa-zイロハ①-⑨]?\s*[．.、　]?\s*受注(実績|状況|の状況)"
    r"|受注(工事)?高(及び|および)受注残高")
# 次の小節の見出し。行頭アンカー(?m)で表す —— パターン中に改行文字を
# 直接書くと、パッチ生成の過程で実際の改行に化けて壊れることがある。
_NEXT_HEAD = re.compile(
    r"^[ａ-ｚa-zイロハ①-⑨]\s*[．.、　]\s*(生産|販売)実績"
    r"|^\s*(生産|販売)実績\s*$"
    r"|^\s*[(（]\d+[)）]\s"
    r"|経営者の視点による", re.M)


def find_section(text, window=2500):
    """**受注実績の小節だけ**を切り出す。無ければ None。"""
    if not text:
        return None
    m = _ORDER_HEAD.search(text)
    if not m:
        m = _SECTION.search(text)
        if not m:
            return None
    seg = text[m.start():m.start() + window]
    nxt = _NEXT_HEAD.search(seg, 1)
    return seg[:nxt.start()] if nxt else seg


def _is_year(tok):
    """『2026』のような年号を金額と読まないための門番。

    表の見出しに「自 2025年２月１日」等が混ざるので、カンマの無い
    1900〜2100 の整数は金額ではなく年として捨てる。6619 で受注残高が
    「当期 2026 / 前年 2025」になっていた（2026-09-02）。
    """
    if "," in tok or "，" in tok or "." in tok:
        return False
    v = _to_f(tok)
    return v is not None and 1900 <= v <= 2100 and v == int(v)


def _numbers_after(seg, pos, take=3):
    """ラベル直後の数値トークンを拾う。次のラベルに当たったら止める。"""
    tail = seg[pos:pos + 400]
    stop = len(tail)
    for _k, rx in _LABEL_RE:
        n = rx.search(tail)
        if n and n.start() > 0:
            stop = min(stop, n.start())
    return [t.group(0) for t in _NUM.finditer(tail[:stop])
            if not _is_year(t.group(0))][:take]


def _unit_scale(seg):
    """百万円に揃える倍率。単位が読めなければ 1.0（＝そのまま）。

    建設は千円、装置メーカーは百万円で書く。**揃えずに保存すると
    金額が1000倍ずれた行が静かに混ざる。**
    """
    if re.search(r"[（(]千円[）)]|金額\(千円\)|金額（千円）", seg):
        return 1.0 / 1000.0
    if re.search(r"[（(]百万円[）)]", seg):
        return 1.0
    return 1.0


def _parse_rollforward(seg):
    """建設型: 前期繰越 / 当期受注 / 当期完成 / 次期繰越 の繰り越し表。"""
    out = {}
    for key, rx in _LABEL_RE:
        m = rx.search(seg)
        if not m:
            continue
        nums = [_to_f(t) for t in _numbers_after(seg, m.end())]
        nums = [n for n in nums if n is not None]
        if not nums:
            continue
        if len(nums) == 1:
            out[key] = {"current": nums[0], "prior": None, "yoy_pct": None}
            continue
        a, b = nums[0], nums[1]
        # 2つ目が比率か金額か。**小数点つきで1000未満なら比率**とみなす。
        # 金額は千円/百万円単位で桁が大きいので実務上この区別で足りる。
        is_ratio = b < 1000 and (b != int(b) or a > 10000)
        if is_ratio:
            out[key] = {"current": a, "prior": (a / (b / 100.0)) if b else None,
                        "yoy_pct": (b - 100.0) if b else None}
        else:
            out[key] = {"current": b, "prior": a,
                        "yoy_pct": ((b / a - 1) * 100.0) if a else None}
    return out


_TOTAL_LINE = re.compile(r"^(合計|計)$", re.M)


def _parse_segment_table(seg):
    """装置メーカー型: 行がセグメント、列が受注高／受注残高の表。

    **合計行を採る。** 最初に出てくる行はセグメント1つ分でしかなく、
    それを会社全体の受注として扱うと桁も意味も違う数字になる
    （6336 で受注高 15,011百万円のところ 4,219 を拾っていた）。
    """
    m = _TOTAL_LINE.search(seg)
    if not m:
        return {}
    nums = [_to_f(t.group(0)) for t in _NUM.finditer(seg[m.end():m.end() + 200])
            if not _is_year(t.group(0))]
    nums = [n for n in nums if n is not None]
    header = seg[:m.start()]
    has_backlog = re.search(r"受注残高", header) is not None
    out = {}
    # 列の並びは「受注高, 前年同期比, [受注残高, 前年同期比]」。
    if len(nums) >= 2:
        out["orders"] = {"current": nums[0], "yoy_pct": nums[1] - 100.0,
                         "prior": nums[0] / (nums[1] / 100.0) if nums[1] else None}
    if has_backlog and len(nums) >= 4:
        out["closing_backlog"] = {
            "current": nums[2], "yoy_pct": nums[3] - 100.0,
            "prior": nums[2] / (nums[3] / 100.0) if nums[3] else None}
    return out


def _consistent(d):
    """繰り越し表の整合: 期首 + 受注 - 完成 ≒ 期末。

    合わなければ**採らない**。表のどこかを取り違えた可能性が高く、
    もっともらしい数字を出すほうが黙って間違えるぶん危ない。
    """
    need = ("opening_backlog", "orders", "completed", "closing_backlog")
    if not all(k in d and d[k].get("current") is not None for k in need):
        return True                      # 判定材料が無いときは通す
    o, r, c, e = (d[k]["current"] for k in need)
    calc = o + r - c
    scale = max(abs(e), abs(calc), 1.0)
    return abs(calc - e) / scale <= 0.05


def parse_orders(text):
    """受注実績を dict で返す。取れないキーは入れない（0 で埋めない）。

    金額は**百万円に揃える**。単位が混ざったまま保存すると、
    1000倍ずれた行が静かに混ざる。
    """
    seg = find_section(text)
    if not seg:
        return None
    # **レイアウトを先に決める。** 後から補う形にすると、繰り越し表の
    # パターンがセグメント表の1行目を拾ったところで確定してしまい、
    # 会社全体ではなく1セグメントの数字が入る（6336 で実際に起きた）。
    if re.search(r"セグメントの名称|セグメントごと", seg):
        out = _parse_segment_table(seg)
    else:
        out = _parse_rollforward(seg)
        if not _consistent(out):
            out = {}
    if not out:
        return None
    k = _unit_scale(seg)
    if k != 1.0:
        for v in out.values():
            for f in ("current", "prior"):
                if v.get(f) is not None:
                    v[f] = v[f] * k
    return out


def yoy(d, key):
    """そのキーの前年比(%)。取れなければ None。**0 を返さない。**"""
    v = (d or {}).get(key)
    if not v:
        return None
    if v.get("yoy_pct") is not None:
        return round(v["yoy_pct"], 1)
    if v.get("prior"):
        return round((v["current"] / v["prior"] - 1) * 100.0, 1)
    return None


# ---------------------------------------------------------------- 保存
DDL = """
CREATE TABLE IF NOT EXISTS s13_orders (
    filing_id        INTEGER PRIMARY KEY,
    code             TEXT NOT NULL,
    period_label     TEXT,
    doc_id           TEXT,
    doc_date         TEXT,
    doc_source       TEXT,
    period_note      TEXT,
    available        INTEGER DEFAULT 0,
    unavailable_reason TEXT,
    opening_backlog  REAL, orders_amount REAL,
    completed_amount REAL, closing_backlog REAL,
    backlog_yoy_pct  REAL, orders_yoy_pct REAL,
    computed_at      TEXT DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_s13_code ON s13_orders(code, doc_date);
"""


def store(con, filing, d, reason=None):
    def g(k, f="current"):
        return ((d or {}).get(k) or {}).get(f)
    con.execute(
        "INSERT OR REPLACE INTO s13_orders (filing_id, code, period_label,"
        " doc_id, doc_date, doc_source, period_note, available,"
        " unavailable_reason, opening_backlog, orders_amount, completed_amount,"
        " closing_backlog, backlog_yoy_pct, orders_yoy_pct, computed_at)"
        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'))",
        (filing["id"], filing["code"],
         "%s-%s" % (filing["date"][:4], filing["subtype"]),
         filing["doc_id"], filing["date"], filing["source"],
         "受注実績テーブル（%s）" % (filing["subtype"] or ""),
         1 if d else 0, reason,
         g("opening_backlog"), g("orders"), g("completed"), g("closing_backlog"),
         yoy(d, "closing_backlog"), yoy(d, "orders")))


def run(con, limit=None, codes=None, latest_per_code=None):
    from screener.extract.disclosure_flags import full_body
    con.executescript(DDL)
    where, args = "WHERE xbrl_path IS NOT NULL", []
    if codes:
        where += " AND code IN (%s)" % ",".join("?" * len(codes))
        args += list(codes)
    sql = ("SELECT id, code, date, subtype, doc_id, source, xbrl_path FROM filings "
           + where + " ORDER BY code, date DESC")

    # **銘柄ごとに最新 N 件だけ見る。** 全書類を舐めると zip 展開が
    # 数万回になり、途中で落ちたときに何も残らない。表示用のフラグは
    # 「その銘柄の直近の開示」で足りる。
    if latest_per_code:
        seen, keep = {}, []
        for r in con.execute(sql, args).fetchall():
            n = seen.get(r["code"], 0)
            if n < latest_per_code:
                keep.append(r)
                seen[r["code"]] = n + 1
        rows_iter = keep
    else:
        rows_iter = con.execute(sql, args).fetchall()
    st = {"scanned": 0, "no_body": 0, "no_section": 0, "parsed": 0,
          "backlog_yoy": 0}
    for i, f in enumerate(rows_iter, 1):
        body = full_body(con, f)
        st["scanned"] += 1
        if not body:
            st["no_body"] += 1
            continue
        d = parse_orders(body)
        if not d:
            st["no_section"] += 1
            store(con, f, None, "受注実績の節が無い")
            continue
        st["parsed"] += 1
        if yoy(d, "closing_backlog") is not None:
            st["backlog_yoy"] += 1
        store(con, f, d)
        if i % 200 == 0:
            con.commit()
            C.log("  %d 件処理" % i)
    con.commit()
    return st


def main(argv=None):
    import argparse
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--limit", type=int)
    p.add_argument("--latest-per-code", type=int, default=2,
                   help="銘柄ごとに見る最新書類の件数（既定2）")
    p.add_argument("--lock-wait", type=float, default=3600.0,
                   help="他の書き込みジョブを待つ秒数（既定1時間）")
    p.add_argument("--codes", nargs="*")
    p.add_argument("--report", action="store_true")
    a = p.parse_args(argv)
    con = C.init_db()
    if a.report:
        con.executescript(DDL)
        n = con.execute("SELECT COUNT(*) FROM s13_orders").fetchone()[0]
        av = con.execute("SELECT COUNT(*) FROM s13_orders WHERE available=1").fetchone()[0]
        C.log("=== S13 受注残（シャドウ・0点） ===")
        C.log("  評価書類 %d / 節あり %d (%.1f%%)" % (n, av, av * 100.0 / max(n, 1)))
        C.log("  業種別の節の存在率:")
        for r in con.execute(
                "SELECT c.sector17 s, COUNT(*) n, SUM(o.available) a "
                "FROM s13_orders o JOIN companies c ON c.code=o.code "
                "GROUP BY s HAVING n>=5 ORDER BY (1.0*SUM(o.available)/COUNT(*)) DESC"):
            C.log("    %-18s %4d件中 %4d (%.0f%%)"
                  % (r["s"] or "-", r["n"], r["a"] or 0,
                     (r["a"] or 0) * 100.0 / max(r["n"], 1)))
        return 0
    # **既存の作法どおり writer_lock で直列化する。**
    # 日次の ScreenerTdnetArchiver(19:00) が同じDBに数時間書くので、
    # 待たずに書くと database is locked で落ちる（2026-09-02 に踏んだ）。
    try:
        with C.writer_lock("s13_orders", wait_seconds=a.lock_wait):
            st = run(con, a.limit, a.codes, a.latest_per_code)
    except C.WriterBusy as e:
        C.log("  ! 他の書き込みジョブが実行中: %s" % e)
        return 2
    for k, v in st.items():
        C.log("  %-16s %s" % (k, format(v, ",")))
    if st["scanned"] and st["parsed"] == 0:
        C.log("  ! 1件も解析できていない。節の見出しパターンを疑う")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
