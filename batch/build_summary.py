# -*- coding: utf-8 -*-
"""build_summary.py — 105銘柄の3シートサマリーを作る。

    python batch/build_summary.py [--date 20260906] [--out models/summary_<date>.xlsx]

【仕様の出所についての注記】
追補15 §5-2 は「105銘柄サマリー生成プロンプト（既交付）どおり」と指示しているが、
そのプロンプトは本セッションのコンテキストにもリポジトリにも存在しなかった
(`batch/build_summary.py` も未作成だった)。そこで、指示から確実に読み取れる要件

  * 3シート構成
  * 8410 セブン銀行は 105 件のカウントに含めない（型D のテスト生成でユニバース外）

だけを固定し、残りは既存の成果物（batch_state.json と各 *_validation.txt）から
機械的に作れる範囲で構成した。仕様が届いたら本ファイルを差し替えること。

シート構成:
  1. 全銘柄        105件の一覧（ticker / 社名 / 型 / 株価 / Target / レーティング /
                   アップサイド / ゲート / データ基準日 / 備考）
  2. 型別集計      型ごとの件数・レーティング分布・アップサイドの分布
  3. 要確認        WARN が立っている銘柄と、その WARN 本文（暫定モデル・按分仮定・
                   簿価加算・鮮度緩和などが一覧で読める）
"""
import argparse
import glob
import io
import json
import os
import re

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCLUDE_FROM_UNIVERSE = {"8410"}   # 型D のテスト生成。ユニバース外（追補15 §5-2）

HDR_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", start_color="000080", end_color="000080")
BODY = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
GREY = Font(name="Arial", size=9, italic=True, color="808080")
WARN_FILL = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
BAD_FILL = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
_T = Side(style="thin", color="B0B0B0")
BOX = Border(left=_T, right=_T, top=_T, bottom=_T)


def _hdr(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, start=1):
        c = ws.cell(row, i, lab)
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


def read_validation(code, date):
    """*_validation.txt から FAIL/WARN/SKIP/PASS と WARN 本文を読む。"""
    p = os.path.join(ROOT, "models", f"{code}_DCF_Model_{date}_validation.txt")
    if not os.path.exists(p):
        return None, []
    t = io.open(p, encoding="utf-8", errors="replace").read()
    m = re.search(r"FAIL (\d+) / WARN (\d+) / SKIP (\d+) / PASS (\d+)", t)
    gate = m.group(0) if m else None
    warns = [re.sub(r"\s+", " ", x).strip()
             for x in re.findall(r"^\[WARN\]\s+(.*)$", t, re.M)]
    return gate, warns


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default="20260906")
    ap.add_argument("--out")
    a = ap.parse_args()
    out = a.out or os.path.join(ROOT, "models", f"summary_{a.date}.xlsx")

    st = json.load(io.open(os.path.join(ROOT, "batch", "batch_state.json"),
                           encoding="utf-8"))
    codes = sorted(k for k in st if not k.startswith("_"))
    rows = []
    for c in codes:
        e = st[c]
        gate, warns = read_validation(c, a.date)
        rows.append(dict(
            code=c, name=e.get("name", ""), typ=str(e.get("type", "")),
            status=e.get("status", ""), price=e.get("price"), target=e.get("target"),
            verdict=e.get("verdict", ""), upside=e.get("upside"),
            wacc=e.get("wacc"), gate=gate, warns=warns,
            basis=e.get("regenerated", ""), reason=e.get("queue_reason", ""),
        ))

    wb = openpyxl.Workbook()

    # ── Sheet 1: 全銘柄 ──
    ws = wb.active
    ws.title = "全銘柄"
    ws.cell(1, 1, f"105銘柄サマリー（基準日 {a.date}／市場データ TARGET_DATE 基準）").font = \
        Font(name="Arial", size=13, bold=True)
    ws.cell(2, 1, f"8410 セブン銀行は型D のテスト生成でユニバース外のため 105 件に含めない。"
                  f"本シートの件数 = {len(rows)}。").font = GREY
    _hdr(ws, 4, ["ticker", "会社名", "型", "状態", "株価", "Target",
                 "レーティング", "アップサイド", "WACC", "ゲート", "基準/再生成", "キュー理由"],
         [9, 24, 8, 8, 11, 11, 11, 11, 8, 30, 26, 60])
    r = 5
    for d in rows:
        ws.cell(r, 1, d["code"]).font = BODY
        ws.cell(r, 2, d["name"]).font = BODY
        ws.cell(r, 3, d["typ"]).font = BODY
        ws.cell(r, 4, d["status"]).font = BODY
        for col, key, fmt in ((5, "price", '#,##0.0'), (6, "target", '#,##0'),
                              (8, "upside", '0.0%'), (9, "wacc", '0.00%')):
            v = d[key]
            cc = ws.cell(r, col, v)
            cc.font = BODY
            if isinstance(v, (int, float)):
                cc.number_format = fmt
        ws.cell(r, 7, d["verdict"]).font = BODY
        ws.cell(r, 10, d["gate"] or "").font = BODY
        ws.cell(r, 11, d["basis"]).font = BODY
        ws.cell(r, 12, (d["reason"] or "")[:400]).font = GREY
        if d["status"] != "done":
            for col in range(1, 13):
                ws.cell(r, col).fill = BAD_FILL
        elif d["warns"]:
            for col in range(1, 13):
                ws.cell(r, col).fill = WARN_FILL
        for col in range(1, 13):
            ws.cell(r, col).border = BOX
        r += 1
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:L{r-1}"

    # ── Sheet 2: 型別集計 ──
    ws2 = wb.create_sheet("型別集計")
    ws2.cell(1, 1, "型別集計").font = Font(name="Arial", size=13, bold=True)
    _hdr(ws2, 3, ["型", "件数", "done", "queued", "BUY", "HOLD", "SELL",
                  "アップサイド中央値"], [26, 8, 8, 9, 8, 8, 8, 16])
    from collections import defaultdict
    import statistics
    g = defaultdict(list)
    for d in rows:
        key = (d["typ"] or "未宣言").split("(")[0].strip() or "未宣言"
        g[key].append(d)
    r = 4
    for k in sorted(g):
        v = g[k]
        ups = [x["upside"] for x in v if isinstance(x["upside"], (int, float))]
        vals = [k, len(v),
                sum(1 for x in v if x["status"] == "done"),
                sum(1 for x in v if x["status"] != "done"),
                sum(1 for x in v if x["verdict"] == "BUY"),
                sum(1 for x in v if x["verdict"] == "HOLD"),
                sum(1 for x in v if x["verdict"] == "SELL"),
                statistics.median(ups) if ups else None]
        for i, x in enumerate(vals, start=1):
            c = ws2.cell(r, i, x)
            c.font = BODY
            c.border = BOX
            if i == 8 and isinstance(x, float):
                c.number_format = '0.0%'
        r += 1
    ws2.cell(r + 1, 1, "レーティングは Target/株価 の比率から機械的に付いたもので、"
                       "投資判断そのものではない。").font = GREY

    # ── Sheet 3: 要確認 ──
    ws3 = wb.create_sheet("要確認")
    ws3.cell(1, 1, "要確認（WARN が立っている銘柄・キュー維持の銘柄）").font = \
        Font(name="Arial", size=13, bold=True)
    ws3.cell(2, 1, "暫定モデル・按分仮定・簿価加算・鮮度緩和などがここに集まる。"
                   "朝のレビューはこのシートから読むのが速い。").font = GREY
    _hdr(ws3, 4, ["ticker", "会社名", "型", "状態", "WARN / キュー理由"],
         [9, 24, 10, 8, 130])
    r = 5
    for d in rows:
        items = list(d["warns"])
        if d["status"] != "done" and d["reason"]:
            items = [f"[キュー維持] {d['reason']}"] + items
        for it in items:
            ws3.cell(r, 1, d["code"]).font = BODY
            ws3.cell(r, 2, d["name"]).font = BODY
            ws3.cell(r, 3, d["typ"]).font = BODY
            ws3.cell(r, 4, d["status"]).font = BODY
            c = ws3.cell(r, 5, it[:900])
            c.font = BODY
            c.alignment = Alignment(wrap_text=True, vertical="top")
            for col in range(1, 6):
                ws3.cell(r, col).border = BOX
                ws3.cell(r, col).fill = BAD_FILL if d["status"] != "done" else WARN_FILL
            r += 1
    ws3.freeze_panes = "A5"

    wb.save(out)
    done = sum(1 for d in rows if d["status"] == "done")
    print(f"wrote {out}")
    print(f"  ユニバース {len(rows)} 件（8410 は除外）: done {done} / queued {len(rows)-done}")
    print(f"  WARN のある銘柄: {sum(1 for d in rows if d['warns'])} 件")


if __name__ == "__main__":
    main()
