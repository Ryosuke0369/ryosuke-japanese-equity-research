"""screener/extract/disclosure_flags.py — 本文から拾う開示フラグ。**表示のみ。**

2つ検知する。どちらも 2026-09-02 に人間が原文精読して見つけた欠落。

継続企業の前提（タスク2b）
--------------------------
「計上済み数字＝証拠」という前提が揺らいでいる会社を、候補から消さずに
**印を付ける**。消さないのは、自動検知は誤検知しうるから ——
候補を消す権限は人が管理する `security_flags` 側にだけ持たせる。

会計処理変更（タスク4）
-----------------------
3565 アセンテック FY2027-Q1 は純額処理への変更で売上 -677百万円。
売上 YoY -45.5% のうち約11ptが会計処理変更によるもの。

**注記事項の①〜④はすべて「無」だった。** 変更は定性情報の説明文にしか
出ていない。だから注記だけを見る実装では捕まらない —— 本文キーワードとの
2系統が要る。片方だけでは静かに取り逃がす。
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

# ---------------------------------------------------------------- 継続企業
_GC_SECTION = re.compile(r"継続企業の前提に関する(注記|重要事象)")
# 「該当事項はありません」型は 0。疑義・不確実性の語があれば 1。
_GC_NONE = re.compile(r"該当事項(は)?(ありません|ございません|無し|なし)")
_GC_HIT = re.compile(r"継続企業の前提に関する重要な疑義|重要な疑義を生じさせる|"
                     r"継続企業の前提に(関する)?重要な不確実性|重要な不確実性が認められ")
# 「重要な不確実性は認められない」は打ち消し。**否定文を陽性にしない。**
_GC_NEG = re.compile(r"重要な不確実性は認められない|"
                     r"重要な不確実性は認められません|疑義は解消")


def detect_going_concern(text):
    """(flag, 根拠文) を返す。判定材料が無ければ (0, None)。

    節が無い会社が大多数なので、**節が無い＝0** とする（未評価と分けない）。
    継続企業の前提の注記は「あれば必ず書く」性質のものなので、
    無いことをもって「該当なし」と読んでよい数少ない項目。
    """
    if not text:
        return 0, None
    for m in _GC_SECTION.finditer(text):
        seg = text[m.start():m.start() + 700]
        if _GC_NONE.search(seg):
            continue
        if _GC_NEG.search(seg) and not _GC_HIT.search(seg):
            continue
        if _GC_HIT.search(seg):
            hit = _GC_HIT.search(seg)
            s = max(0, hit.start() - 60)
            return 1, re.sub(r"\s+", " ", seg[s:hit.end() + 90])
    return 0, None


# ---------------------------------------------------------------- 会計処理変更
# (a) 注記事項の①〜④。「無」以外なら変更あり。
_NOTE_ITEMS = ("会計基準等の改正に伴う会計方針の変更", "会計方針の変更",
               "会計上の見積りの変更", "修正再表示")
_NOTE_NONE = re.compile(r"[：:\s]*(無|なし|ありません|該当なし)")

# (b) 定性情報本文のキーワード。**注記が全部「無」でも本文には出る。**
#
# **語があるだけでは採らない。** 素朴に語だけで拾うと 60書類中 25件(42%)が
# 陽性になり、印として役に立たなかった（2026-09-02 実測）。内訳は
# 「収益認識に関する会計基準の適用」28% と「表示方法の変更」25% ——
# どちらも 2021年以降のほぼ全有報に載る定型文である。
# そこで「**変更したと書いてあること**」を条件に加える: キーワードの近傍に
# 変更を述べる動詞があること。表示方法の変更はさらに金額影響を要求する
# （見出しだけで中身が「該当なし」のものを落とすため）。
_CHANGED = r"(変更(し|いた)(て|まし)|変更しております|変更いたしました|"            r"変更したこと|見直し(た|まし)|組替え)"
_AMOUNT = r"[0-9０-９,，]+\s*(百万円|千円|円|％|%)"

_NARRATIVE = [
    # 純額/総額処理は「代理人か本人か」の判断変更で、売上が大きく動く。
    ("純額処理", r"純額(で|の)?(表示|処理|認識)|純額処理", True, False),
    ("総額処理", r"総額(で|の)?(表示|処理|認識)|総額処理", True, False),
    ("代理人と判断", r"代理人と(して)?(判断|認識)", False, False),
    ("本人と判断", r"本人と(して)?(判断|認識)", False, False),
    ("収益認識方法の変更", r"収益(の)?認識(方法|基準)(の)?変更", False, False),
    # 定型文になりやすいので、変更の記述と金額影響の両方を要求する。
    ("表示方法の変更", r"表示方法の変更", True, True),
]
_NARRATIVE_RE = [(k, re.compile(p), need_ch, need_amt)
                 for k, p, need_ch, need_amt in _NARRATIVE]
_CHANGED_RE = re.compile(_CHANGED)
_AMOUNT_RE = re.compile(_AMOUNT)


def detect_accounting_change(text):
    """(統合flag, 注記由来flag, 本文由来flag, 根拠文) を返す。"""
    if not text:
        return 0, 0, 0, None
    from_notes = 0
    for item in _NOTE_ITEMS:
        for m in re.finditer(re.escape(item), text):
            tail = text[m.end():m.end() + 40]
            # **コロンが続くものだけを「項目」とみなす。**
            # 「(4) 会計方針の変更・会計上の見積りの変更・修正再表示」のような
            # 見出し行は項目名を並べただけで値を持たない。これを項目と
            # 誤認すると、注記が全部「無」の書類が変更ありになる
            # （2026-09-02 の fixture で検出）。
            v = re.match(r"[ 	　]*[：:]\s*(\S+)", tail)
            if not v:
                continue
            if not _NOTE_NONE.match("：" + v.group(1)):
                from_notes = 1
                break
        if from_notes:
            break
    hits = []
    for key, rx, need_ch, need_amt in _NARRATIVE_RE:
        for m in rx.finditer(text):
            near = text[m.start():m.end() + 220]
            if need_ch and not _CHANGED_RE.search(near):
                continue
            if need_amt and not _AMOUNT_RE.search(near):
                continue
            # 根拠は**一致箇所から後ろ**を採る。前を含めると直前の別項目
            # （「修正再表示：無」等）を巻き込んで読み手を誤らせる。
            hits.append("%s: %s" % (key, re.sub(r"\s+", " ",
                                                text[m.start():m.end() + 90])))
            break
    from_nar = 1 if hits else 0
    matched = " ｜ ".join(hits[:3]) or None
    return (1 if (from_notes or from_nar) else 0), from_notes, from_nar, matched


# ---------------------------------------------------------------- 保存
def store(con, filing, going, gc_text, acc, ac_notes, ac_nar, ac_text):
    con.execute(
        "INSERT OR REPLACE INTO disclosure_flags (filing_id, code, period_label,"
        " doc_id, doc_date, going_concern, gc_matched, accounting_change,"
        " ac_from_notes, ac_from_narrative, ac_matched, computed_at)"
        " VALUES (?,?,?,?,?,?,?,?,?,?,?,datetime('now'))",
        (filing["id"], filing["code"],
         "%s-%s" % (filing["date"][:4], filing["subtype"]),
         filing["doc_id"], filing["date"], going, gc_text, acc,
         ac_notes, ac_nar, ac_text))


def full_body(con, filing_row):
    """PublicDoc の htm を**全部**つないで返す。

    S12 の `body_of` は `0102010_honbun`（企業の概況）1本しか読まない。
    経営成績の節を取るにはそれで足りるが、**会計方針・継続企業の注記は
    別ファイル（経理の状況）にある**。狭い本文で探すと、検知は例外も
    出さずに 0 件になる —— 実際 120書類で会計処理変更 0 件、3565 の
    有報でも「純額」が 0 回だった（2026-09-02）。
    """
    import zipfile
    from bs4 import BeautifulSoup
    p = C.full_path(filing_row["xbrl_path"])
    if not p or not os.path.exists(p):
        return None
    out = []
    try:
        with zipfile.ZipFile(p) as z:
            names = sorted(n for n in z.namelist()
                           if "PublicDoc" in n and n.endswith(".htm"))
            for n in names:
                soup = BeautifulSoup(z.read(n).decode("utf-8", "replace"), "lxml")
                for t in soup(["style", "script"]):
                    t.decompose()
                out.append(soup.get_text(chr(10), strip=True))
    except Exception:
        return None
    joined = chr(10).join(out)
    return re.sub(chr(10) + "{3,}", chr(10) * 2, joined) if out else None


def run(con, limit=None, codes=None, latest_per_code=None):
    """本文が取れる書類を舐めてフラグを立てる。事後条件つき。"""
    body_of = full_body
    where = "WHERE xbrl_path IS NOT NULL"
    args = []
    if codes:
        where += " AND code IN (%s)" % ",".join("?" * len(codes))
        args += list(codes)
    sql = ("SELECT id, code, date, subtype, doc_id, xbrl_path FROM filings "
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
    st = {"scanned": 0, "no_body": 0, "going_concern": 0,
          "accounting_change": 0, "ac_notes_only": 0, "ac_narrative_only": 0}
    for i, f in enumerate(rows_iter, 1):
        body = body_of(con, f)
        st["scanned"] += 1
        if not body:
            st["no_body"] += 1
            continue
        g, gt = detect_going_concern(body)
        a, an, ar, at = detect_accounting_change(body)
        st["going_concern"] += g
        st["accounting_change"] += a
        if an and not ar:
            st["ac_notes_only"] += 1
        if ar and not an:
            st["ac_narrative_only"] += 1
        store(con, f, g, gt, a, an, ar, at)
        if i % 200 == 0:
            con.commit()            # 途中で落ちてもそこまでは残す
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
    a = p.parse_args(argv)
    con = C.init_db()
    # **既存の作法どおり writer_lock で直列化する。**
    # 日次の ScreenerTdnetArchiver(19:00) が同じDBに数時間書くので、
    # 待たずに書くと database is locked で落ちる（2026-09-02 に踏んだ）。
    try:
        with C.writer_lock("disclosure_flags", wait_seconds=a.lock_wait):
            st = run(con, a.limit, a.codes, a.latest_per_code)
    except C.WriterBusy as e:
        C.log("  ! 他の書き込みジョブが実行中: %s" % e)
        return 2
    for k, v in st.items():
        C.log("  %-22s %s" % (k, format(v, ",")))
    # **静黙縮退の検出。** 0件でも例外は出ないので、ここで落とす。
    if st["scanned"] and st["scanned"] == st["no_body"]:
        C.log("  ! 本文が1件も取れていない。zip の場所か抽出経路を疑う")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
