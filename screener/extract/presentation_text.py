"""screener/extract/presentation_text.py — 決算説明会資料のPDFからテキストを抽出する。

方針（2026-09-02 承認）
----------------------
- **PDF原本は無加工で保存**する。抽出は別ファイルに書き、原本を上書きしない
- **スコアリングには使わない。** 前年ペアが揃うのは 2027年秋以降
  （TDnet の保持が約6週間なので遡れない）。それまでは収集と
  **抽出率のテレメトリ**だけを回す
- 抽出の成功/失敗を1件ずつログに残す。「抽出できなかった」ことが
  分からないまま母集団が痩せるのを防ぐ

なぜ抽出率を測るのか
--------------------
短信の定性セクションは HTML(`qualitative.htm`)で 96% 抽出できたが、
説明会資料は PDF で、図表中心の資料はテキストが取れないことがある。
**取れないものがどれくらいあるか**を先に知っておかないと、
2027年に差分を取ろうとした時点で母集団が想定の半分だった、という事故になる。

    python -m screener.extract.presentation_text --all
    python -m screener.extract.presentation_text --report
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import unicodedata

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

SUBTYPE = "決算説明資料"
OUT_DIRNAME = "presentation_text"
MIN_CHARS = 200          # これ未満は「図表のみ等で実質取れていない」とみなす


def _extract(path):
    """PDF からテキストを取る。使えるライブラリを順に試す。

    戻り値: (text, engine) / 取れなければ (None, 理由)
    """
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader        # 古い環境向け
        except ImportError:
            return None, "no_pdf_library"
    try:
        reader = PdfReader(path)
        parts = []
        for page in reader.pages:
            t = page.extract_text() or ""
            if t.strip():
                parts.append(t)
        txt = re.sub(r"\n{3,}", "\n\n", "\n".join(parts)).strip()
        return (txt, "pypdf") if txt else (None, "empty_text")
    except Exception as e:
        return None, "%s" % type(e).__name__



# ---- タイトルから対象期を読む ----------------------------------------------
# 収集時点で確定させる。2027年にペアを組むときタイトルからしか対象期が
# 分からないので、その時に全件再パースするのを避ける。
_RE_FY = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月期")
_RE_NENDO = re.compile(r"(\d{4})\s*年度")
_RE_Q = re.compile(r"第\s*([1-4])\s*四半期")
_RE_RANGE = re.compile(r"[~〜～]")          # 複数期をまたぐ訂正


def parse_period(title):
    """タイトル -> (period_label, fiscal_year, fy_end_month, quarter_type,
                   is_correction, ambiguous)

    決められないときは推測しない。period_label=None を返し、
    ペアの構築対象から外す（0点ではなく欠損）。
    """
    t = unicodedata.normalize("NFKC", title or "")
    is_corr = 1 if ("訂正" in t) else 0
    # 「(2024年3月期第1四半期~2026年3月期)の一部訂正」のような複数期まとめは
    # どの期のものか一意に決まらない
    if _RE_RANGE.search(t) and len(_RE_FY.findall(t)) > 1:
        return None, None, None, None, is_corr, 1

    m = _RE_FY.search(t)
    fy = fy_end = None
    if m:
        fy, fy_end = int(m.group(1)), int(m.group(2))
    else:
        m2 = _RE_NENDO.search(t)
        if m2:
            fy = int(m2.group(1))            # 「2025年度」形式。期末月は不明
    if fy is None:
        return None, None, None, None, is_corr, 0

    qm = _RE_Q.search(t)
    if qm:
        q = "%dQ" % int(qm.group(1))
    elif ("半期" in t) or ("中間" in t):
        q = "2Q"                             # 半期 = 第2四半期累計
    else:
        q = "FY"                             # 通期
    return "FY%d-%s" % (fy, q), fy, fy_end, q, is_corr, 0



# ---- 資料の種類 ------------------------------------------------------------
# 同じ期に本編・補足・書き起こし・サマリ・質疑応答が並行して出る。世代を
# 種類ごとに分けないと「最新世代」が訂正版ではなく別文書を指してしまう。
# 告知(動画公開のお知らせ等)は内容を持たないので差分母集団から外す。
_RE_QUOTED = re.compile(r"[「『]([^」』]{4,})[」』]")
_KIND_RULES = (
    ("質疑応答", re.compile(r"質疑応答|Q&A|QandA")),
    ("書き起こし", re.compile(r"書き起こし|書起こし|文字起こし|全文")),
    ("サマリ", re.compile(r"サマリ|要約|エグゼクティブ|概要")),
    ("補足", re.compile(r"補足")),
    ("本編", re.compile(r"説明資料|説明会資料|プレゼンテーション|ファクトブック|説明会")),
)
# 告知の判定は「公開されるモノがこのPDFの外にあるか」で決める。
# 「動画公開のお知らせ」= 中身は動画で、このPDFには無い -> 告知
# 「質疑応答概要の公開について」= このPDFが質疑応答そのもの -> 質疑応答
# 単に「〜について/お知らせ」で切ると、後者まで告知になって内容を捨てる。
_RE_EXTERNAL = re.compile(r"動画|ムービー|映像|ウェブ|Web|サイト|URL|配信|開催")
_RE_NOTICE = re.compile(r"(公開|公表|配信|開催).{0,12}(お知らせ|について|ご案内)"
                        r"|のお知らせ$|について$")


def classify_doc_kind(title):
    """タイトル -> 本編/補足/書き起こし/サマリ/質疑応答/告知/unknown

    分類できないものは 'unknown'。**推測で埋めない。** unknown 率が高ければ
    分類器のほうを見直す（テレメトリに出す）。
    """
    t = unicodedata.normalize("NFKC", title or "")
    # 訂正版は「元のタイトル」を引用しているので、その中身で分類する。
    # 訂正のラッパーごと見ると全部「〜について」= 告知になってしまう。
    if "訂正" in t:
        m = _RE_QUOTED.search(t)
        if m:
            t = m.group(1)
    else:
        # 内容そのものではなく「出しました」の告知。中身が無いので除外対象。
        # ただし公開対象がこのPDF自身(質疑応答概要など)なら内容として扱う。
        if (_RE_NOTICE.search(t) and not _RE_QUOTED.search(t)
                and _RE_EXTERNAL.search(t)):
            return "告知"
    for kind, rx in _KIND_RULES:
        if rx.search(t):
            return kind
    return "unknown"


def record_metadata(con, row, text_path, chars, status):
    """対象期と世代を presentation_materials に確定させる。

    （訂正）資料は元と同じ (code, period_label) の別世代として入れる。
    訂正前も残す —— 「当時はそう書かれていた」が事実だから。
    """
    label, fy, fy_end, q, is_corr, amb = parse_period(row["title"])
    kind = classify_doc_kind(row["title"])
    con.execute(
        "INSERT OR REPLACE INTO presentation_materials (filing_id, code, doc_id, "
        " disclosed_date, title, period_label, fiscal_year, fy_end_month, "
        " quarter_type, is_correction, ambiguous, doc_kind, generation, "
        " text_path, text_chars, extract_status, parsed_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,1,?,?,?,?)",
        (row["id"], row["code"], row["doc_id"], row["date"], row["title"],
         label, fy, fy_end, q, is_corr, amb, kind,
         C.store_path(text_path) if text_path else None, chars, status,
         C.utcnow()))
    # 世代は (code, period_label, doc_kind) 内で disclosed_date 昇順に振り直す。
    # 同日は訂正を後置。挿入順で振ると訂正が原本より前に来る。
    if label:
        items = con.execute(
            "SELECT filing_id FROM presentation_materials "
            "WHERE code=? AND period_label=? AND doc_kind IS ? "
            "ORDER BY disclosed_date, is_correction, filing_id",
            (row["code"], label, kind)).fetchall()
        for i, it in enumerate(items, 1):
            con.execute("UPDATE presentation_materials SET generation=? "
                        "WHERE filing_id=?", (i, it[0]))
    return label, kind


def latest_generation(con, code, period_label, as_of, doc_kind="本編"):
    """as_of 時点で見えている最新世代。差分は常にこれを使う。

    doc_kind を指定するのは、同じ期に別種類の資料が並ぶため。既定は本編。
    差分は「本編 vs 前年の本編」で取る。今年の本編と前年のサマリ版を
    比べても意味がない。
    """
    r = con.execute(
        "SELECT filing_id, text_path, generation, is_correction, doc_kind "
        "FROM presentation_materials WHERE code=? AND period_label=? "
        "AND doc_kind=? AND disclosed_date<=? "
        "ORDER BY generation DESC LIMIT 1",
        (code, period_label, doc_kind, as_of)).fetchone()
    return r


def out_path(pdf_rel):
    base = os.path.splitext(os.path.basename(pdf_rel))[0]
    return os.path.join(C.DATA_DIR, OUT_DIRNAME, base + ".txt")


def run(con, limit=None, force=False):
    rows = con.execute(
        "SELECT id, code, date, pdf_path, title, doc_id FROM filings "
        "WHERE source='tdnet' AND subtype=? AND pdf_path IS NOT NULL "
        "ORDER BY date DESC", (SUBTYPE,)).fetchall()
    if limit:
        rows = rows[:int(limit)]
    C.log("決算説明資料: %d 件" % len(rows))
    os.makedirs(os.path.join(C.DATA_DIR, OUT_DIRNAME), exist_ok=True)

    stats = {"ok": 0, "short": 0, "failed": 0, "skipped": 0,
             "missing_pdf": 0, "no_period": 0}
    reasons = {}
    for r in rows:
        dst = out_path(r["pdf_path"])
        if os.path.exists(dst) and not force:
            stats["skipped"] += 1
            continue
        src = C.full_path(r["pdf_path"])
        if not os.path.exists(src):
            stats["missing_pdf"] += 1
            record_metadata(con, r, None, 0, "missing_pdf")
            C.log("  ! PDF が無い filing %s: %s" % (r["id"], r["pdf_path"]))
            continue
        txt, engine = _extract(src)
        if txt is None:
            stats["failed"] += 1
            reasons[engine] = reasons.get(engine, 0) + 1
            record_metadata(con, r, None, 0, "failed")
            C.log("  ! 抽出失敗 %s %s (%s): %s"
                  % (r["code"], r["date"], engine, (r["title"] or "")[:40]))
            continue
        with open(dst, "w", encoding="utf-8") as fh:
            fh.write(txt)
        status = "short" if len(txt) < MIN_CHARS else "ok"
        label, kind = record_metadata(con, r, dst, len(txt), status)
        if status == "short":
            stats["short"] += 1
            C.log("  ~ 実質空 %s %s (%d字): %s"
                  % (r["code"], r["date"], len(txt), (r["title"] or "")[:40]))
        else:
            stats["ok"] += 1
        if label is None:
            stats["no_period"] = stats.get("no_period", 0) + 1
    C.log("抽出: 成功 %d / 実質空 %d / 失敗 %d / 既存スキップ %d / PDF欠 %d"
          % (stats["ok"], stats["short"], stats["failed"],
             stats["skipped"], stats["missing_pdf"]))
    for k, v in sorted(reasons.items()):
        C.log("  失敗理由 %s: %d" % (k, v))
    C.log("  対象期を決められなかった資料: %d 件（複数期の訂正など。推測しない）"
          % stats.get("no_period", 0))
    con.commit()
    done = stats["ok"] + stats["short"] + stats["failed"]
    if done:
        C.log("抽出率（実用に足る本文が取れた割合）: %.1f%%"
              % (stats["ok"] / done * 100))
    return stats


def report(con):
    n = con.execute("SELECT COUNT(*) c FROM filings WHERE source='tdnet' "
                    "AND subtype=?", (SUBTYPE,)).fetchone()["c"]
    d = os.path.join(C.DATA_DIR, OUT_DIRNAME)
    files = [f for f in os.listdir(d)] if os.path.isdir(d) else []
    sizes = []
    for f in files:
        try:
            sizes.append(os.path.getsize(os.path.join(d, f)))
        except OSError:
            pass
    C.log("=== 決算説明資料 テレメトリ ===")
    C.log("  収集した資料: %d 件" % n)
    C.log("  抽出テキスト: %d 件" % len(files))
    if sizes:
        sizes.sort()
        C.log("  本文サイズ: 中央値 %d バイト / 最小 %d / 最大 %d"
              % (sizes[len(sizes) // 2], sizes[0], sizes[-1]))
    C.log("  ※ スコアには使わない。前年ペアが揃うのは 2027年秋以降。")
    return 0


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--all", action="store_true", help="未抽出をすべて処理")
    p.add_argument("--limit", type=int)
    p.add_argument("--force", action="store_true", help="既存テキストも作り直す")
    p.add_argument("--report", action="store_true")
    a = p.parse_args(argv)
    con = C.init_db()
    if a.report:
        return report(con)
    if not (a.all or a.limit):
        p.error("--all か --limit か --report を指定する")
    run(con, a.limit, a.force)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
