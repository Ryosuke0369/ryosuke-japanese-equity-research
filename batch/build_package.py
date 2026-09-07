# -*- coding: utf-8 -*-
"""build_package.py — 105銘柄の成果物を配布パッケージにまとめる。

    python batch/build_package.py [--date 20260906] [--out dist/]

【仕様の出所についての注記】
追補15 §5-3 は「パッケージ作成手順.md（既交付）どおり」と指示しているが、その手順書は
本セッションのコンテキストにもリポジトリにも `Downloads/` にも存在しなかった。
手順が無い状態で同梱物・構造・命名を推測すると受け取り側の期待と食い違うため、
**下記の構成は自作であることを README にも明記**したうえで、後から差し替えやすいように
1ファイルにまとめてある。手順書が届いたら本ファイルを差し替えること。

構成:
    <out>/package_<date>/
        README.md                    パッケージの説明・基準日・免責・自作である旨
        MANIFEST.txt                 全ファイルの相対パス・サイズ・SHA256
        summary_<date>.xlsx          3シートサマリー
        models/                      105ユニバースの DCF ワークブック
        models/_outside_universe/    8410（型D テスト）・5726（回帰基準）— 105 に不算入
        validation/                  各銘柄の validate レポート
        reports/                     最終レポート・セッション別レポート・キュー台帳
        data/overrides/              前提の正本（JSON）
        data/comps/                  ピア入力（CSV）
        data/adjustments/            Adjustments Log の入力（あるもののみ）
"""
import argparse
import glob
import hashlib
import io
import json
import os
import shutil
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTSIDE_UNIVERSE = {"8410", "5726"}   # 追補15 §5-2: 105 のカウントに含めない


def sha256(path, buf=1 << 20):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            b = f.read(buf)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def copy(src, dst_dir, name=None):
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, name or os.path.basename(src))
    shutil.copy2(src, dst)
    return dst


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default="20260906")
    ap.add_argument("--out", default=os.path.join(ROOT, "dist"))
    a = ap.parse_args()
    pkg = os.path.join(a.out, f"package_{a.date}")
    if os.path.isdir(pkg):
        shutil.rmtree(pkg)
    os.makedirs(pkg, exist_ok=True)

    st = json.load(io.open(os.path.join(ROOT, "batch", "batch_state.json"), encoding="utf-8"))
    universe = {k for k in st if not k.startswith("_")}

    counts = {"models": 0, "outside": 0, "validation": 0,
              "overrides": 0, "comps": 0, "adjustments": 0, "reports": 0}

    # ── モデルと validate レポート ──
    for f in sorted(glob.glob(os.path.join(ROOT, "models", f"*_DCF_Model_{a.date}.xlsx"))):
        code = os.path.basename(f).split("_")[0]
        if code in OUTSIDE_UNIVERSE or code not in universe:
            copy(f, os.path.join(pkg, "models", "_outside_universe"))
            counts["outside"] += 1
        else:
            copy(f, os.path.join(pkg, "models"))
            counts["models"] += 1
    for f in sorted(glob.glob(os.path.join(ROOT, "models",
                                           f"*_DCF_Model_{a.date}_validation.txt"))):
        copy(f, os.path.join(pkg, "validation"))
        counts["validation"] += 1

    # ── サマリー ──
    summ = os.path.join(ROOT, "models", f"summary_{a.date}.xlsx")
    if os.path.exists(summ):
        copy(summ, pkg)

    # ── 入力（前提の正本） ──
    for sub, pat, key in (("overrides", "*_overrides.json", "overrides"),
                          ("comps", "*_comps.csv", "comps"),
                          ("adjustments", "*_adjustments.json", "adjustments")):
        for f in sorted(glob.glob(os.path.join(ROOT, "data", sub, pat))):
            code = os.path.basename(f).split("_")[0]
            if code in universe or code in OUTSIDE_UNIVERSE:
                copy(f, os.path.join(pkg, "data", sub))
                counts[key] += 1

    # ── レポート ──
    for name in (f"final_report_{a.date[:4]}-{a.date[4:6]}-{a.date[6:]}.md",
                 "final_report_20260907.md", "sessionA_report_20260907.md",
                 "sessionB_report_20260907.md", "overnight_report_20260907.md",
                 "queue_ledger_20260907.md", "task3_typeDE_report_20260907.md",
                 "batch_state.json"):
        p = os.path.join(ROOT, "batch", name)
        if os.path.exists(p):
            copy(p, os.path.join(pkg, "reports"))
            counts["reports"] += 1

    # ── README ──
    done = sum(1 for c in universe if st[c].get("status") == "done")
    io.open(os.path.join(pkg, "README.md"), "w", encoding="utf-8", newline="").write(f"""# 日本株 DCF モデル パッケージ（基準日 {a.date}）

## 基準

- **市場データ**: TARGET_DATE = {a.date}（直前営業日の終値）
- **開示データ**: 各銘柄の最新の確定開示。追補13 §A の鮮度緩和を適用しており、
  使用した開示の基準日と docID は各ワークブックの Adjustments Log に記録されている
- ユニバース **{len(universe)} 銘柄**のうち **{done} 件**を生成済み

## 中身

| パス | 内容 |
|---|---|
| `summary_{a.date}.xlsx` | 3シートサマリー（全銘柄 / 型別集計 / 要確認） |
| `models/` | {counts['models']} 件の DCF ワークブック |
| `models/_outside_universe/` | {counts['outside']} 件。8410（型D パスのテスト生成）と 5726（回帰リファレンス）で、**{len(universe)} のカウントには含まない** |
| `validation/` | {counts['validation']} 件の validate レポート（FAIL / WARN / SKIP / PASS の内訳） |
| `data/overrides/` | {counts['overrides']} 件。**各モデルの前提の正本**。数値の出所・導出・限界はここの `_` 接頭辞キーに書いてある |
| `data/comps/` | {counts['comps']} 件のピア入力 |
| `data/adjustments/` | {counts['adjustments']} 件 |
| `reports/` | 最終レポート・セッション別レポート・キュー台帳・`batch_state.json` |
| `MANIFEST.txt` | 全ファイルのサイズと SHA256 |

## 読む順番

1. `reports/final_report_20260907.md` — 何をやったか、何が残っているか、裁定が必要な事項
2. `summary_{a.date}.xlsx` の **「要確認」シート** — WARN が立っている銘柄とその理由が一覧で読める
   （暫定モデル・按分仮定・簿価加算・鮮度緩和はすべてここに集まる）
3. 個別のモデル — Executive Summary → Adjustments Log の順に見ると前提が追える

## 免責・取扱い

- **すべて DRAFT である。** 投資判断に用いる前に、前提（`data/overrides/`）と
  WARN（`validation/`）を必ず確認すること
- レーティングは Target と株価の比率から機械的に付いたもので、投資判断そのものではない
- 一部の銘柄には**按分仮定**（6301）、**簿価による加算**（型F 各社）、**会社計画に基づく
  projection**（3401 / 7731 / 6326）、**暫定モデル**（6594）が含まれる。
  該当箇所は WARN と overrides の注記で明示している

## このパッケージ構成について

追補15 §5-3 が参照する「パッケージ作成手順.md」が見つからなかったため、
**この構成は自作**である（`batch/build_package.py`）。手順書が届いたら差し替えること。
""")

    # ── MANIFEST ──
    lines = ["# MANIFEST — 相対パス\tサイズ(bytes)\tSHA256", ""]
    total = 0
    for dirpath, _, files in os.walk(pkg):
        for fn in sorted(files):
            if fn == "MANIFEST.txt":
                continue
            p = os.path.join(dirpath, fn)
            rel = os.path.relpath(p, pkg).replace("\\", "/")
            sz = os.path.getsize(p)
            total += sz
            lines.append(f"{rel}\t{sz}\t{sha256(p)}")
    lines.append("")
    lines.append(f"# ファイル数 {len(lines)-3} / 合計 {total:,} bytes")
    io.open(os.path.join(pkg, "MANIFEST.txt"), "w", encoding="utf-8",
            newline="").write("\n".join(lines) + "\n")

    # ── zip ──
    zpath = os.path.join(a.out, f"package_{a.date}.zip")
    if os.path.exists(zpath):
        os.remove(zpath)
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for dirpath, _, files in os.walk(pkg):
            for fn in files:
                p = os.path.join(dirpath, fn)
                z.write(p, os.path.relpath(p, a.out))

    print(f"package : {pkg}")
    print(f"zip     : {zpath}  ({os.path.getsize(zpath):,} bytes)")
    for k, v in counts.items():
        print(f"  {k:<12} {v}")


if __name__ == "__main__":
    main()
