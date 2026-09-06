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

## #4 `--date YYYYMMDD` のネイティブ対応

### 症状（先行報告 §F-6-4）

ファイル名の日付が `datetime.now()` 由来だったため、暦日を跨ぐバッチは
`<ticker>_DCF_Model_20260906.xlsx` と `..._20260907.xlsx` に分裂し、先に作った方が
stale な双子として残った。part4 は `TARGET_DATE` 環境変数と `batch/regen.sh` の
rename ステップで回避していた（回避策であって修正ではない）。

### 期待する動作

ファイル名の日付は**実行日ではなく分析基準日**である。基準日を引数で受け取る。

### 変更内容

**`scripts/generate_dcf.py`**

- `--date YYYYMMDD` を追加（`--output-dir` の直後）。
- `_resolve_date_stamp(raw)` を新設。`None` なら今日、書式違反なら **exit 2** で停止する
  （黙って今日にフォールバックしない。基準日の取り違えは全ファイル名に波及する）。
- `date_str` の算出を `_resolve_date_stamp(args.date)` に差し替え。出力パスは #3 で
  `main()` 先頭に移してあるので、`--force` ガードも `--date` の指すファイルを見る。

**`batch/regen.sh`**

- `--date "$TARGET"` を渡すようにし、**rename ステップを削除**した。
  `TARGET_DATE` 環境変数は既定値 `20260906` として残す（呼び出し側の互換）。
- `generate_dcf.py` が Step 9 で validate を回すので、ラッパー側の重複 validate も削除。

### 検証

| 項目 | 結果 |
|---|---|
| `--date 2026-09-06`（ハイフン付き＝書式違反） | `ERROR: --date must be YYYYMMDD` / **exit 2** |
| `--date 20260906`（既存ファイルあり・`--force` なし） | exit 3（#3 のガードが `--date` の指すパスを見ている） |
| `--date 20261231`（新規） | `models/5726_DCF_Model_20261231.xlsx` を生成（検証後に削除） |
| 5726 回帰（`--force --date 20260906`、46セル） | **differences: 0** / validate FAIL 0 WARN 0 SKIP 0 PASS 19 |

---

## #9 会社予想(guidance)が全銘柄で取得できない

### 症状（先行報告 §F-6-9）

85件**すべて**で会社予想が取得できず、Management シナリオが常に Base と同一だった。
5本のシナリオのうち1本が実質機能していない状態。

### 原因（2つ。いずれも構造的で、リトライでは直らない）

**(1) 決算短信は TDnet の文書であって EDINET の文書ではない。**
`fetch_tanshin()` は EDINET を `docTypeCode=140`（四半期報告書）で探していた。
四半期報告書は 2024年4月に制度廃止されており、そもそも業績予想を載せる書類でもない。
**成功しえない探索に1銘柄あたり約50秒**を使っていた。

**(2) 有報スキャンが銘柄でスコープされていなかった。**
Step 3 は `tmp/edinet_data/**/*.xbrl` を再帰 glob していた。これは
**過去にダウンロードした全社の全書類**（2026-09-06 時点で510ディレクトリ）で、
最初に forecast が取れたファイルを採用する実装だった。
有価証券報告書に業績予想は載らないため実際には常に空振りし、**1銘柄あたり約200秒**を消費した。
—— が、もし1件でもヒットしていたら**他社の業績予想を当該モデルに適用していた**。
空振りが幸いしていただけで、これは汚染バグである。

### 数値が実際にある場所

`screener/fetch/tdnet_archiver.py` が TDnet 決算短信を毎営業日アーカイブし、
`screener/extract/xbrl_parser.py` が Summary の inline-XBRL から業績予想を
`guidance` テーブルへ抽出済みだった。**同一リポジトリ内に一次ソースがあった。**
2026-09-06 時点で 2,984 コード収録。

### 変更内容

**`scripts/guidance_fetcher.py`（新規）** — 取得ラダーを1か所にまとめた。

| 段 | ソース | 備考 |
|---|---|---|
| 1 | screener `guidance` テーブル（TDnet 決算短信由来） | 最新の開示日 → その日の最新FY の1組だけを採る（同一短信が中間期予想と通期予想の両方を載せるため、混ざらないようにする） |
| 2 | **当該銘柄の**キャッシュ済み EDINET XBRL | パスは呼び出し側が渡す。全社 glob は廃止 |
| 3 | EDINET 短信探索（レガシー） | **既定オフ**。`--tanshin-fallback` で明示的に有効化 |
| — | 取得できず | Management は CAGR にフォールバック（現行維持）。**理由を記録**する |

- 単位変換: `guidance` テーブルは円、DCF config は百万円（`/1e6`）。
- **FY ガード**: 予想の FY が最新実績 FY 以下なら「それは履歴であって予想ではない」として
  無視し、理由を記録する。過去を未来としてモデル化することを構造的に防ぐ。
- 対応する項目は `revenue` / `operating_income` / `net_income_parent` の3つのみ。
  下流が読まないキー（`ordinary_income` / `eps`）は意図的にマップしない。

**`scripts/edinet_fetcher.py`**
- `fetch_and_parse_multi_year()` が、当該銘柄の XBRL パスを
  `merged["_meta"]["xbrl_paths"]` に公開するようにした（四半期分を先頭に）。

**`scripts/generate_dcf.py`**
- Step 3 を全面置換。`get_guidance()` を呼び、`min_fy_year` に最新実績FYを渡す。
- `--tanshin-fallback` を追加（既定オフ）。
- 取得元 / 取得できなかった理由を `config["_guidance_source"]` / `["_guidance_note"]` に載せる。
- 取得できなかった場合は `final_warnings` に1行出す（Management が会社予想でない旨）。

**`templates/dcf_comps_template.py`**
- Pipeline Metadata の受け渡しリストに `_guidance_source` / `_guidance_note` を追加。
  ワークブック単体で「その Management が会社予想由来かどうか」が判る。

### 検証

| 項目 | 結果 |
|---|---|
| 85銘柄の取得率 | **0/85 → 78/85**（DB照会のみ、生成なしで実測） |
| 5726 実取得 | `FY2027 disclosed 2026-08-25 / revenue 48,000mn` → Year 1 growth **+2.2%** |
| 5726 所要時間 | **393s → 138s**（Step 3 が 251s → 0.0s） |
| 5726 回帰（46セル） | **differences: 0** |
| ワークブック記録 | `guidance_source=tdnet_screener_db` / `guidance_note=...` が Pipeline Metadata に出る |

`--all-sheets` の差分（Executive Summary 3 / Reverse DCF 9 / Comps 1）は本修正とは無関係の
**既存差**である。基準ワークブックは 2026-08-26 11:44 の生成物、テンプレートはその後
同日 12:56 に改稿されており（ラベル文言の変更）、その分の文字列差が出ているだけで、
数値セルの差は 0 である。

### 残る 7 件（取得できない理由を個別に確認済み。バグではない）

| ticker | 理由 |
|---|---|
| 6861 キーエンス | **業績予想を開示しない会社**。短信XBRLはアーカイブ済みだが予想ファクトが存在しない |
| 4901 富士フイルム / 7751 キヤノン | **米国基準の短信**。screener 側のタグマッピングが未対応（短信XBRLはアーカイブ済み） |
| 6506 安川電機（2月期）/ 9983 ファストリ（8月期）/ 3086 J.フロント（2月期） | 直近の短信がアーカイブの保持期間（2026-07-23〜）の外 |
| 3863 日本製紙 | 業績予想の修正を**PDFのみ**で開示（XBRLなし） |

いずれも「取得経路の不備」ではなく「その経路に数値が無い」ケースである。
4901 / 7751 は screener 側の米国基準マッピングを拡張すれば取れる見込みで、
**別課題として `docs/calibration_backlog.md` に登録する**（本フェーズの範囲外）。

---
