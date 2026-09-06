# フェーズ2 パイプライン修正 — 仕様書兼実施記録 (2026-09-06)

作成: 2026-09-06
対象リポジトリ: `<HOME>\ryosuke-japanese-equity-research`
ブランチ: `phase2-pipeline-fixes-20260906`
上位文書: 「フェーズ2_パイプライン修正と全件再生成プロンプト」§1 修正11項目 / §2 新βルール
先行文書: `batch/batch_report_20260905.md` §F-6（修正候補の集約表）、`docs/DCFパイプライン標準運用手順書.md` §6-6

## 0. 実施の原則

- **1修正 = 1検証 = 1commit**。修正をまとめて入れない。
- 各修正の後に **5726 回帰テスト**を実行する。
  ```
  python scripts/generate_dcf.py 5726 --force
  python scripts/diff_models.py models/5726_DCF_Model_20260826.xlsx \
                                models/5726_DCF_Model_20260906.xlsx
  ```
  基準ワークブック `models/5726_DCF_Model_20260826.xlsx` は recalc 済み。
  比較対象は `scripts/diff_models.py` の `KEY_CELLS` **46セル**。
- **ベースライン（修正前）の実測**: 2026-09-06、未修正コードで 5726 を再生成し
  基準ワークブックと突合 → **differences: 0 / 46**。回帰ハーネスが決定論的に動くことを
  先に確認してから修正に着手した（5726 は株価・株数・β・hist 全系列を overrides で
  静的に固定しているため、ライブ取得の揺らぎが入らない）。
- 「差分ゼロ」が常に正解ではない。**意図した差分は許容し、その根拠を各節に記録する**。
  意図しない差分を残したまま次の修正に進まない。

### ベースライン性能（修正前・5726 実測）

| フェーズ | 所要 |
|---|---:|
| Step 1 EDINET 取得 | 129.1s |
| Step 2 LTM/yfinance | 2.0s |
| **Step 3 guidance 抽出** | **251.4s**（キャッシュ XBRL 全走査 200.6s + `fetch_tanshin` 50.7s、結果は取得ゼロ） |
| Step 4〜6 config/市場データ/comps | 4.5s |
| Step 7 生成 | 1.1s |
| Step 8 recalc (Excel COM) | 3.4s |
| Step 9 validate | 0.3s |
| **合計** | **393s / 銘柄** |

85銘柄の全件再生成は逐次で約9時間。**Step 3 が全体の 64%** を占め、しかも1件も
取得できていない（#9）。#9 の修正は正確性の修正であると同時に最大の高速化でもある。

---

## #3 `--force` なしのスキップが exit 0 を返す

### 症状（先行報告 §15-6 / §F-6-3）

> overrides を書き換えて `python scripts/generate_dcf.py 6752` を実行したところ、Step 7 で
> `WARNING: ... already exists` と出てワークブックを生成せずに終了したが、exit code は 0、
> さらに続く validate は**古いファイル**に対して `VERDICT: PASS` を出した。

### 実測による切り分け（推測しない）

修正前のコードで再現を試みたところ、**`generate_dcf.py` 単体の終了コードは 1 であり、
validate にも進んでいなかった**。先行報告の「exit 0」「validate が古いファイルを PASS」は
`generate_dcf.py` の挙動ではなく、**`batch/regen.sh` の構造**に由来していた。

```sh
# 修正前の regen.sh
python -u scripts/generate_dcf.py "$t" --force "$@" 2>&1 | grep -E 'VERDICT|...'
#                                                        ^^^^ 終了コードが grep のものに化ける
TARGET="${TARGET_DATE:-20260905}"
for d in ...; do   # ← 生成の成否と無関係に実行され、validate が走る
```

パイプで `grep` に渡した時点で生成側の終了コードは失われ、後続の rename + validate は
無条件に実行される。**「保護のつもりが silent no-op」の実体はラッパー側にあった。**

残る実害が2つあった:

1. 既存ファイルの検査が **Step 7**（EDINET / yfinance の通信をすべて終えた後）にあり、
   何も生成しない実行に **372 秒**かかっていた（実測）。
2. 拒否の告知が `WARNING:` であり、`grep -E '...|ERROR|...'` を通すラッパーの目に
   留まらなかった。

### 期待する動作

- 既存ファイルがあり `--force` が無いなら、**通信の前に**即座に非ゼロ終了する。
- 告知は `ERROR:` とし、「何も生成していない」ことを明示する。
- 「何も生成しなかった」(3) と「生成したが validate FAIL」(1) を終了コードで区別する。
- ラッパー（`batch/regen.sh`）は生成の終了コードを見て、失敗時は validate に進まない。

### 変更内容

**`scripts/generate_dcf.py`**

- 終了コードの規約をモジュール冒頭にコメントで明記し、`EXIT_SKIPPED = 3` を定義。
  - `0` 生成 + validate 通過 / `1` 生成したが validate FAIL / `2` 不正な起動・契約違反 /
    `3` 何も生成せず（出力先が既存 + `--force` なし）
- 出力パス解決と `--force` ガードを `main()` の**先頭**（`num_years` 決定の直後、
  overrides 読込より前）へ移動。Step 7 側の重複ガードは削除し、参照コメントを残した。
- メッセージを `ERROR:` に変更し、`exit 3` の意味を本文に書いた。

**`batch/regen.sh`**

- 生成の出力を一時ファイルに落として終了コードを捕捉し、その後 `grep` で表示する
  （`/bin/sh` に `PIPESTATUS` は無いため、パイプ自体をやめる方式を採った）。
- 生成が非ゼロで終わったら **rename も validate も行わずに同じコードで終了**する。

### 検証

| 項目 | 結果 |
|---|---|
| 既存ファイルあり・`--force` なし | `ERROR: ... already exists - nothing was generated.` / **exit 3** / **2 秒**（修正前 372 秒） |
| 5726 回帰（46セル） | **differences: 0** |

---
