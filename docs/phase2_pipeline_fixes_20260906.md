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

## #8 validate の `SKIP > 0` が `VERDICT: PASS` を通す

### 症状（先行報告 §F-6-8）

9503 で再計算がタイムアウトで中断したまま validate が走り、
`FAIL 0 / SKIP 5 / VERDICT: PASS` を出した。Target Price 等は空欄だった。

SKIP は「そのチェックが**実行できなかった**」という意味であって、
「実行して通った」ではない。中立扱いにすると「未検証」と「検証済みクリーン」が
同じ結論に潰れる。

### 期待する動作

`SKIP > 0` は PASS にしない。ただし `--no-recalc` のように**意図的な部分検証**の
経路は残す。

### 変更内容

**`scripts/validate_output.py`**

- 全チェック実行後に SKIP を数え、`allow_skip` でなければ
  **チェック99「All checks executed」を FAIL として追加**する。
  合成行にしたのは、既存の `Result` の集計・レンダリング・終了コードが
  そのまま働き、レポートにも理由が1行で残るため（VERDICT の分岐を増やさない）。
- `validate_workbook(path, write_report=True, allow_skip=False)` に引数追加。
- CLI に `--allow-skip` を追加（引数解析をフラグ対応に）。
- モジュール docstring に契約を明記。

**`scripts/generate_dcf.py`**

- Step 9 で `allow_skip=args.no_recalc` を渡す。`--no-recalc` では値レベルの
  チェックが原理的に走れないため、そこだけは部分検証を許す。通常ランは必ず recalc する。

### 検証

| 入力 | 結果 |
|---|---|
| recalc 済み 5726 | `FAIL 0 / WARN 0 / SKIP 0 / PASS 19` / **PASS** / exit 0 |
| 同じブックの**キャッシュ値なし**コピー | `FAIL 1 / SKIP 6 / PASS 13` / **FAIL** / **exit 1** |
| 同上 + `--allow-skip` | **PASS** / exit 0 |
| 5726 回帰（46セル） | **differences: 0** |

---

## #1 `_val(..., default=0)` — 欠損を 0 に落とす

### 症状（先行報告 §F-6-1）

`generate_dcf.py` の `_val()` が「キーが無い」を **0** として返していた。
この既定値ひとつから、性質の異なる4種類の誤りが出ていた。

| 症状 | 実例 |
|---|---|
| 古い年度の営業利益が 0 → 「OI成長率 = (0−0)÷0」で `#DIV/0!` | 生成19件中 **8件**（2802 / 2502 / 2503 / 2801 / 2897 / 4005 / 4183 / 4188） |
| `net_debt` が 0 | **2801 キッコーマン**（実際はネットキャッシュ −50,341百万円 ≒ 1株54円）/ **2897**（実際 77,200 が 4,418） |
| `core_ebitda = OI + D&A` を片脚欠損のまま算出 | 6件。**4502 の Comps 参考株価 −2,761円** |
| Reverse DCF の導出が上記をそのまま継承 | — |

`base_year_revenue` の既定値は **1（百万円）**だった。これは比率をすべて
4桁ずれさせるが、エラーにはならない。

### 期待する動作

欠損は 0 ではなく **None**。呼出側で空欄 + 警告。
`hist_*` / `net_debt` / `core_ebitda` / Reverse DCF 導出への 0 混入を断つ。

### 変更内容

**`scripts/generate_dcf.py`**

- `_val(d, key)` は**素の値（欠損は None）**を返す。数値既定値が本当に欲しい場所には
  `_num(d, key, default)` を明示的に使う。これで**この関数に残る 0 はすべて誰かが選んだ 0** になる。
  - `_num` を使う場所: 比率の分母、`or` 連鎖のフォールバック（NWC 基準年 AR/Inv/AP、
    trade receivables/payables、`hist_nwc_pct`、capex/da 比率の平均、latest FY 比率）。
    これらは症状の報告が無く、かつ falsy 判定に依存しているため挙動を変えない。
    ただし**欠損した場合は警告を出す**ようにした。
- P/L 5系列（revenue / cogs / sga / operating_income / net_income）は None を保持し、
  **どの年度のどの項目が欠損したかを1行ずつ警告**する。
- **COGS の逆算**は「revenue・OI・SGA が3つとも揃っているときだけ」に限定した。
  欠損 OI を 0 として逆算すると、もっともらしい COGS が書けてしまう
  （2802 では COGS が売上と同値で出ていた）。再構成できない年度は空欄のまま。
- **net_debt**: EDINET に net_debt 行が無い場合、`total_debt − cash` から導出し、
  導出根拠を `config["_net_debt_source"]` に記録する。どちらも無ければ **None**。
- **core_ebitda**: OI と D&A の**両方**が揃わなければ None。片脚だけの和を EBITDA と
  呼ばない。`primary_multiple` は `(core_ebitda or 0) > 0` で判定する。
- **Step 4.55（新設・overrides 適用後）**: `net_debt` と `base_year_revenue` が
  それでも未確定なら **exit 2 で停止**し、埋めるべき overrides キーを名指しする。
  この2つは推測できないうえ、DCF の1株価値に直接効く。

**`templates/dcf_comps_template.py`**

- Financial Statements の P/L 行（6〜16）を **None 安全**にした。
  キャッシュフロー行は bug B1 の修正時に同じ規則を入れてあったが、P/L には無く、
  空白セルに対して `=(C11-B11)/B11` を書いていた。これが `#DIV/0!` の実体。
  - 派生行（粗利・粗利率・営業利益率・純利益率）は、分子と**0でない分母**が
    そろっているときだけ数式を書く。
  - YoY 成長率は、前年が空欄または 0 なら `n/a`（0 は正当な値だが正当な分母ではない）。

### 検証

**(a) 合成データによる単体検証**（EDINET 相当の入力を直接与えた）

| 入力 | 結果 |
|---|---|
| FY2022/FY2023 の OI・NI・COGS・SGA が欠損 | `hist_operating_income = [None, None, 9000, 10000]`、欠損年度を名指しする警告4本 |
| `net_debt` 行なし・`total_debt` 20,000 / `cash` 5,000 | `net_debt = 15,000`、`_net_debt_source = "derived: total_debt 20,000 - cash 5,000"` |
| `net_debt` 行なし・`total_debt`/`cash` も欠損 | `net_debt = None` / `UNAVAILABLE` → Step 4.55 で exit 2 |
| 最新期 OI・D&A が欠損 | `core_ebitda = None` + 警告 |

**(b) 生成ワークブックの検証**（欠損2年を含む合成銘柄）

```
COGS         [None, None, 84000, 91000]
GrossProfit  [None, None, '=E6-E7', '=F6-F7']
OpMargin     [None, None, '=E11/E6', '=F11/F6']
OpIncYoY     ['n/a', 'n/a', 'n/a', '=(F11-E11)/E11']
```
欠損年度に数式が**書かれない**。`recalc.py` の検証は `ERRORS: None found`。

**(c) 5726 回帰**: 主要46セル **differences: 0**、`--all-sheets` の
Financial Statements も **0 differing cell**（5726 は全系列が overrides で埋まっており、
None 経路を通らないため挙動不変であることが確認できた）。

---

## #5 validate に `core_ebitda > 0` チェックが無い

### 症状（先行報告 §F-6-5）

`core_ebitda` は EDINET の「最新期営業利益 + 最新期D&A」から作られる。
D&A が取れないと壊れた値のまま Comps に流れ、**4502 で参考株価 −2,761円**が出たが
`FAIL 0 / WARN 0` で通過した。バッチ側の `batch/check_core_ebitda.py` が別ゲートとして
これを拾っていたが、**validate 単体では検出できない**状態だった。

### 期待する動作

- 営業利益 > 0 なら `core_ebitda > 0` を検査する。
- Comps 参考株価に sanity band `0 < 値 < 現値 × 10` を課す。

### 変更内容

**`templates/dcf_comps_template.py`** — 判定材料を Pipeline Metadata に出す
（ラベル文字列の走査ではなく、生成器の意図と突き合わせるという既存方針に合わせた）。

```
core_ebitda / core_net_income / comps_ebitda_excluded / comps_per_excluded /
latest_operating_income
```

**`scripts/validate_output.py`** — チェック2本を追加。

- **19. Subject EBITDA vs operating income**
  D&A は負にならないので `EBITDA >= 営業利益` が成り立つ。両者が互いの検算になる。
  - 営業利益 > 0 かつ `core_ebitda <= 0`（または欠損なのに EV/EBITDA を除外していない） → **FAIL**
  - `core_ebitda < 営業利益` → **WARN**（期間・スコープが違う基準混在の疑い）
  - `core_ebitda` 欠損だが EV/EBITDA が N/A として除外済み → **PASS**（設計どおりの逃げ道）
- **20. Comps reference prices in a sane band**
  Executive Summary C18 / C19 を現値（C9）で規格化する。
  - テキスト（`N/A` / `INVALID`）→ **PASS**（除外は設計どおり）
  - 数値 ≤ 0 → **FAIL**（使えない手法は数値でなくテキストで書かれるべき）
  - 数値 > 現値 × 10 → **WARN**

### 検証

**(a) チェック関数の単体検証**

| 入力 | 判定 |
|---|---|
| C18 = −2,761（4502 と同型）| **FAIL** |
| C18 = `INVALID (n<3)` / C19 = `N/A` | PASS |
| C18 = 45,000（現値1,500の30倍）| WARN |
| C18 = 1,582 / C19 = 2,247 | PASS |
| 営業利益 1,200 / `core_ebitda` −500・除外なし | **FAIL** |
| 営業利益 1,200 / `core_ebitda` 欠損・EV/EBITDA 除外済 | PASS |
| 営業利益 1,200 / `core_ebitda` 欠損・除外なし | **FAIL** |
| `core_ebitda` 900 < 営業利益 1,200 | WARN |

**(b) 5726 実ラン**: `[PASS] 19 core_ebitda 8,575 >= operating income 5,524 (implied D&A 3,051)` /
`[PASS] 20 EV/EBITDA=1,582 ok; PER=2,247 ok`。`FAIL 0 / WARN 0 / SKIP 0 / PASS 21`。

**(c) 5726 回帰（46セル）**: **differences: 0**

なお、既存の `models/4502_DCF_Model_20260905.xlsx` は現時点で C18 = 5,684（正値）であり、
報告された −2,761 は既に rework 済みだった。旧テンプレ生成のためチェック19は
`latest_operating_income` が無く SKIP になり、#8 の規則で当該ファイルは FAIL 判定になる
（**全件再生成でメタデータごと更新される**）。

---

## #2 市場データのサイレント・プレースホルダ

### 症状（先行報告 §F-6-2）

`generate_dcf.py` の config は `current_price = 1000` / `shares_outstanding = 10,000,000`
を「yfinance が上書きする placeholder」として持ち、`get_live_market_data()` は
**例外を握り潰してその placeholder をそのまま返して**いた。

**4568 第一三共**は時価総額 10,000百万円（実際は 5,084,100百万円）で評価され、
**Target 294,427円 / BUY +293%** を `FAIL 0` のまま出力した。

### 期待する動作

取得失敗はハードエラー。validate はプレースホルダ値を検出して FAIL。

### 変更内容

**`templates/dcf_comps_template.py` — `get_live_market_data()`**

- 戻り値を `(price, shares, beta, note)` の4つに変更。
  **取れなかった項目は None を返す**（呼出側が overrides で埋まるかを判断する）。
- 例外時も None + 失敗理由の `note` を返す。もっともらしい定数を返さない。
- 返す β は**生値**。Blume 調整とクランプはテンプレ側に集約する（#6 で使う）。
- 旧2値展開だったテンプレ内デモブロックを4値に追随。
  `scripts/run_215A_dcf.py` は添字アクセスのため無変更で動作する。

**`scripts/generate_dcf.py`**

- config の初期値を `None` に変更（placeholder を廃止）。
- **市場データ override の適用を D/E 自動計算より前に移動**。
  これまでは後だったため、`current_price` を override しつつ `de_ratio` を
  override しない銘柄では、**D/E だけがライブ株価の時価総額で計算され**、
  ワークブックが表示する株価と食い違っていた（1モデル内に2つの時価総額があった）。
  → **意図した挙動変更**。5726 は `de_ratio` を明示指定しているため回帰に差分は出ない。
- ハードガード: 価格・株数が数値かつ正でなければ **exit 2**。
  加えて `(1000, 10,000,000)` の**完全一致ペア**も拒否する
  （本当にその値なら overrides に明示させ、意図を記録に残す）。
- 取得元を `config["_market_data_source"]` に記録し、Pipeline Metadata に出す。
  `_net_debt_source`（#1）も同時に出すようにした。

**`scripts/validate_output.py`** — チェック21「Market data is not a placeholder」。
生成器が止めるのは第一の鍵、これは**第二の鍵**であり、生成器を通っていない
手編集ワークブックもカバーする。

### 検証

**(a) 生成器の停止（end-to-end）**: yfinance を必ず失敗させるスタブを `PYTHONPATH` に置き、
`current_price` / `shares` を外した overrides で 5726 を実行。

```
  yfinance lookup FAILED for 5726.T: simulated yfinance outage
ERROR: market data could not be established:
  - current_price is None (yfinance failed: simulated yfinance outage)
  - shares_outstanding is None (yfinance failed: simulated yfinance outage)
  Set "current_price" and "shares": {"fully_diluted_shares": N} in data/overrides/...
EXIT=2
```
修正前ならここで price 1,000 / shares 10,000,000 のワークブックが出来ていた。

**(b) チェック21の単体検証**

| 入力 | 判定 |
|---|---|
| price 1,000 × shares 10,000,000 | **FAIL**（プレースホルダ完全一致） |
| price 0 | **FAIL** |
| shares None | **FAIL** |
| price 2,727 × 36,800,000 | PASS |

**(c) 5726 実ラン**: `[PASS] 21 price 2,727 x 36,800,000 shares = 100,354 JPY mn
[source: yfinance; from overrides: beta, current_price, shares, shares_outstanding]`。
`FAIL 0 / WARN 0 / SKIP 0 / PASS 22`。

**(d) 5726 回帰（46セル）**: **differences: 0**

**副次的な観測**: 5726 の yfinance 生βは **0.484**。旧クランプ域 [0.6, 1.75] の外なので
**無言で 1.0 に置換される**値だった（overrides が実測回帰値 1.55 を明示していたため
本銘柄では発現していない）。#6 の前提を実測で確認した形になる。

---

## #6 β のクランプ域とサイレント置換 —— 新βルール

### 症状（先行報告 §F-6-6 / プロンプト §2）

旧ルールは `[0.6, 1.75]` のクランプで、**範囲外は無言で 1.0 に置換**していた。

- 2026-09-05 バッチの **85件中 57件（67%）が下限 0.60 に張り付いた**。
  日本のディフェンシブ銘柄で WACC が 2〜4pt 過大になっていた。
- 上側では太陽誘電の raw 1.561 がそのまま通り、WACC 13.35% を生んだ。
- 置換は**コンソールにも Adjustments Log にも残らなかった**。

実測で確認した例: 5726 の yfinance 生βは **0.484**。旧ルールでは無言で 1.0 に
なる値だった（この銘柄は overrides が実測回帰値 1.55 を明示していたため発現していない）。

### 新ルール（プロンプト §2 の実装）

1. **Blume 調整を既定化**: `β_adj = 0.67 × β_raw + 0.33 × 1.00`。
   回帰βは平均回帰するという標準的な扱い。両裾の歪みがバッチで実証されている。
2. **クランプ域を `[0.3, 2.0]` に拡大**し、**範囲外は無言置換せず WARN**を出す。
   Blume 調整後にこの域を外れるには raw が概ね `[-0.05, 2.49]` の外である必要があり、
   クランプは日常経路ではなく稀なガードになる。
3. **記録**: Adjustments Log の `DCF Model!C8` 行と Pipeline Metadata に
   **raw / adjusted / 採用値の3点**を必ず残す。

`C["beta"]` は**どこから来ても raw**として扱う。yfinance 由来でも overrides 由来でも
同じ処理を通す（#2 で `get_live_market_data` が生βを返すようにしたのはこのため）。

### 変更内容

- **`templates/dcf_comps_template.py`**: WACC 正規化ブロックの β 処理を全面置換。
  `_beta_record`（raw / blume / adopted / basis / clamped）を作り、
  `_meta` 5キー（`beta_raw` / `beta_blume_adjusted` / `beta_adopted_c8` /
  `beta_basis` / `beta_clamped`）と Adjustments Log の自動記録行に流す。
  クランプ発動時の記録ステータスは **「クランプ発動・要確認」**。
- **`docs/overrides_schema.md`**: `beta` は**実測（回帰）βの生値**を入れる契約であることを明記。
  移行時に「旧ルールでクランプ後の値が書かれている overrides を raw として渡すと
  二重に縮小される」ことを警告として追記。

### 検証

**(a) 各経路の単体検証**（合成銘柄で実際にワークブックを生成し C8 を読んだ）

| raw β | Blume 調整後 | 採用値（C8） | 挙動 |
|---:|---:|---:|---|
| 3.000 | 2.340 | **2.000** | クランプ + WARN |
| −0.500 | −0.005 | **0.300** | クランプ + WARN |
| 0.484 | 0.654 | 0.654 | 置換なし（**旧ルールなら無言で 1.0**） |
| 1.550 | 1.369 | 1.369 | 置換なし |
| なし | 1.000 | 1.000 | 市場β採用 + WARN |

**(b) 5726 回帰（46セル）— 意図した差分 20 件**

| 項目 | 修正前 | 修正後 |
|---|---:|---:|
| Beta | 1.5500 | **1.3685** |
| Cost of Equity | 12.93% | **11.93%** |
| WACC | 9.06% | **8.37%** |
| DCF PGM 株価 | 200 | **372** |
| DCF Exit 株価 | 558 | **613** |
| **Target Mid** | **379** | **493** |
| 判定 | SELL | SELL（不変） |

検算: `0.67 × 1.55 + 0.33 = 1.3685` / `Ke = 2.90% + 1.3685 × 5.50% + 1.50% = 11.93%` /
`WACC = 11.93% × (1/1.4568) + 0.59% × (0.4568/1.4568) = 8.37%`。

**差分は β 由来の連鎖 20 セルのみ**で、株価・株数・net debt・税率・RF・ERP・
size premium・Kd・D/E・Comps 統計はすべて 0 差分。**副作用が無いことを確認した**。

`validate`: `FAIL 0 / WARN 0 / SKIP 0 / PASS 22`。

### 全85件への波及（§2-4 として別途実施）

既存 overrides には**旧ルールでクランプ後の値**（0.60 等）が明示値として書かれている。
テンプレを直しても overrides が旧値のままでは反映されないため、
**全85件について β を再導出して overrides を更新する**（§4 の前段で実施）。

---

## #10 `hist_capex` を渡すと C5 の根拠が変わる

### 症状（先行報告 §15-7 / §F-6-10）

`align_hist_series_to_years()` の対象5系列のうち `hist_capex` / `hist_depreciation` を
overrides に入れると、**validate チェック7 の C5 の根拠が「明示前提 5.25%」から
「実績3期平均 8.82%」に切り替わり**、`capex_direct` の投影と不整合になっていた。

### 原因

`capex_method == "direct"` のときだけ、C5/C18 を**実績3期平均で置換**していた。
そのため表示される根拠が、前提とは無関係な `hist_capex` というキーが overrides に
あるかどうかで変わっていた。さらに `direct` 方式で**シートが実際に使う**フォールバック
（`=Revenue*C5`。投影配列が届かない年度に効く）と、CLAUDE.md が
「direct 方式でも `capex_pct` / `da_pct` はフォールバック用に必ず残す」と定めた値は
`capex_pct` であって実績平均ではない。**表示と実算出が食い違っていた**。

### 期待する動作

根拠の優先順位を統一し、表示と実算出を一致させる。

### 変更内容

**`templates/dcf_comps_template.py`**

- 根拠のラダーを **`capex_method` を一切参照しない**1本に統一した。
  1. overrides に `capex_pct` / `da_pct` の明示指定がある → `explicit_override`
  2. 無ければ生成器が有報から自動導出した値 → `auto_hist_avg`
- 実績3期平均は**引き続き算出して metadata に残す**（`capex_pct_hist3yr` /
  `da_pct_hist3yr`）が、**黙って前提に化けることはなくなった**。
- `direct` 方式なのに `capex_pct` / `da_pct` の明示指定が無い場合は WARN を出す
  （CLAUDE.md の契約違反であり、フォールバック値が自動導出値になる旨を告げる）。
- ラベルは方式で変わる（`direct` → `(fallback)`）が、**値の根拠は方式で変わらない**。

**`scripts/validate_output.py`** — チェック7 を新語彙に対応させ、
**実績3期平均を併記**するようにした（前提と実績の乖離が一目で分かる）。

### 検証

**(a) 症状の消滅**（合成銘柄・`revenue_pct`）

| ケース | C5 |
|---|---:|
| override なし | 0.06（自動導出値） |
| **`hist_capex` を渡す** | **0.06（不変）** ← 修正前はここが実績3期平均に化けた |
| `capex_pct: 0.09` を明示 | 0.09 |

**(b) 5726（`capex_method: "direct"` かつ `hist_capex` あり）**

```
[PASS] 7. C5  Capex/Revenue basis   7.80% - explicit assumption from overrides; hist 3yr mean 9.24%
[PASS] 7. C18 D&A/Revenue basis     7.90% - explicit assumption from overrides; hist 3yr mean 5.75%
```
C5 は overrides の `capex_pct: 0.078`（＝契約上のフォールバック値）を表示するようになり、
実績3期平均 9.24% は参考として validate に併記される。

**(c) 回帰**: 基準は **#6 適用後の 5726**（`scratchpad/5726_ref_after_beta.xlsx`）。
主要46セル **differences: 0**。`--all-sheets` の DCF Model 差分は
**C5 / B5 / C18 / B18 の4セルのみ**（＝意図した表示の変更）で、EV・Target は不変
（5726 は5年とも `capex_direct.projections` が埋まっており C5 が計算に効かないため）。

---

## #11 非3月期の docID 検出

### 症状（先行報告 §F-6-11）

**3086 J.フロント リテイリング（2月期）**が既定の探索窓（3/12/6/9月期）に掛からず
`EdinetDocumentNotFound`。`fiscal_year_end_month=2` を**手で明示指定**して解決していた。

### 原因

有報の提出期は「期末月 + 3」であり、`get_document_ids()` はその式で動的探索窓を作る
仕組みを既に持っていた。しかし **`fiscal_year_end_month` は呼び出し側の overrides から
しか来なかった**。非3月期の銘柄でそのキーを書き忘れると、ハードコードされた4シーズンを
素通りして「見つからない」になる。**作業者が事前に決算期を知っている前提**の設計だった。

### 期待する動作

非3月期の探索窓（期末月+3）が、明示指定なしでも正しく適用される。

### 変更内容

**`scripts/edinet_fetcher.py`**

- `infer_fiscal_year_end_month(ticker_code)` を新設。yfinance の
  `lastFiscalYearEnd` / `nextFiscalYearEnd`（UNIX 時刻）から期末月を取る。
  yfinance が無い / 404 / 値なし はすべて `None` を返して従来動作に落ちる。
- `get_document_ids()` は `fiscal_year_end_month` が未指定のときだけ推定を呼ぶ。
  **明示指定は常に優先**する。推定したことはログに残す。

`batch/prescreen.py` は既に `tickers.csv` の `fiscal_year_end_month` を渡しているので
変更不要。渡されなかった銘柄も推定で救われるようになった。

### 検証

**(a) 推定単体**

| ticker | 推定 | 実際 |
|---|---:|---|
| 3086 J.フロント | **2** | 2月期 |
| 6506 安川電機 | **2** | 2月期 |
| 9983 ファストリ | **8** | 8月期 |
| 5726 大阪チタ | 3 | 3月期 |
| 2502 アサヒGHD | 12 | 12月期 |
| 存在しないコード | None | （従来動作にフォールバック） |

**(b) 3086 を `fiscal_year_end_month` **なし**で実行（end-to-end）**

```
[Step 1/9] Fetching 5 years of financial data from EDINET...
Found 5 annual report(s):
  [1] docID=S100Y6K0  period=2026-02-28  filer=Ｊ．フロント　リテイリング株式会社
  [2] docID=S100VUV0  period=2025-02-28
  [3] docID=S100TIM5  period=2024-02-29
  [4] docID=S100QU70  period=2023-02-28
  [5] docID=S100O4WX  period=2022-02-28
```
part4 で `EdinetDocumentNotFound` になった条件で、5期すべてを取得できた。

**(c) 5726 回帰（基準 = #6 適用後）**: **differences: 0**
（5726 は `fiscal_year_end_month: 3` を明示しており推定経路を通らない）

---

## #7 EDINET 探索窓が最新の有価証券報告書を取りこぼす

### 症状（先行報告 §6-4 / §F-6-7）

| ticker | パイプラインが取得した最新本決算 | 実在する最新期 |
|---|---|---|
| 2802 味の素 | FY2025/3（S100VXJA） | **FY2026/3** |
| 2502 アサヒGHD | FY2024/12（S100VHC1） | **FY2025/12** |
| 7741 HOYA / 7752 リコー / 9021 JR西日本 | FY2024〜2025 止まり | FY2026/3 |

9021 は capex に**2期古い値**を使わざるを得なかった。

### 原因（screener の EDINET アーカイブと突合して確定）

`get_document_ids()` は EDINET の**日付別インデックスを1日ずつ問い合わせて**書類を
発見する。したがって「会社の実際の提出日がハードコードされた窓に入っているか」で
結果が変わる。3月期の窓は **6/18〜7/2** だったが、実際の提出日は次のとおり:

| ticker | FY2026/3 提出日 | 旧窓 6/18〜7/2 |
|---|---|---|
| 7741 HOYA | **6/05** | 外れ |
| 2802 味の素 | **6/12** | 外れ |
| 9021 JR西日本 | **6/16** | 外れ |
| 7752 リコー | **6/16** | 外れ |
| 5726 大阪チタ | 6/23 | 入る |
| （最も遅い例）7752 の訂正 | 7/15 | 外れ |

`adaptive_search` の ±5日も年次のドリフトより狭かった。7741 は 6/20 → 6/05 と
**15日**動いており、±5 では次年度が捕捉できない。

### 期待する動作

最新有報を確実に捕捉する（窓ロジックの修正 + 再試行）。

### 変更内容

**Tier 1（新設）— screener の EDINET アーカイブから docID を直接引く**

`screener/fetch/edinet_bulk.py` が提出インデックスを日次で保存しており、
`filings` テーブルに **docID がそのまま入っている**（4,208社 / 21,539件 /
2022-06-01〜2026-08-31）。窓もヒューリスティクスも要らず、**API コール 0 回**で正確。

- `scripts/screener_link.py`（新規）: screener.db の場所解決と読み取り専用接続を
  1か所にまとめた。screener 未セットアップの環境では None を返して従来動作に落ちる。
  `guidance_fetcher.py`（#9）もこちらへ寄せた。
- `scripts/edinet_fetcher.docs_from_screener_archive()`（新規）:
  `type='有報'` かつ `訂正` を除いた行から docID を取り、`get_document_ids()` と
  同じ形の dict を返す。**期末日**は「提出日 + 期末月」から導出し
  （提出は期末後3か月以内なので、提出日直前の期末が対象期）、
  期末月が不明な場合はタイトル解析（西暦・和暦の両方に対応）にフォールバックする。
  タイトルが `有価証券報告書` だけの行が 507件中 10件あるため、この二段構えが要る。
- `get_document_ids()` はまずアーカイブで `found_docs` を**シード**し、
  `num_years` 揃えば **API を1回も叩かずに返す**。足りなければ従来のスキャンが
  差分を埋める（やり直しではなく top-up）。

**Tier 2 — 日付スキャン側も直した**（アーカイブに無い銘柄のために）

- **窓を提出月まるごと + 15日に拡大**（例: 3月期 → 6/01〜7/15）。
  法定期限（期末後3か月）と、その後の遅延提出の両方を覆う。
- **窓の走査順を「期限日から遡る」に変更**。提出は月末に集中するため、
  日付順に走るより少ないコールで典型的な提出日に届く。**探索する日付の集合は同じで、
  順序だけが変わる**（6/23 提出なら 6 コール、6/12 提出なら 13 コール）。
- `adaptive_search` の探索幅を **±5日 → ±15日**に拡大。

### 検証

**(a) アーカイブ解決の実測**（DB のみ、ネットワーク不使用）

| ticker | 解決した最新 docID | 期末 | 旧スキャンの結果 |
|---|---|---|---|
| 2802 | **S100Y992** | 2026-03-31 | FY2025/3 止まり |
| 7741 | **S100Y90T** | 2026-03-31 | FY2024〜2025 止まり |
| 9021 | **S100YCFK** | 2026-03-31 | 同上 |
| 7752 | **S100YBFF** | 2026-03-31 | 同上 |
| 2502 | **S100YSGA** | 2025-12-31 | FY2024/12 止まり |

105銘柄（done 85 + queue 20）で、期末月を与えれば **103銘柄が5期、6594 が4期**。
**6594 ニデックの FY2026/3 有報はアーカイブにも存在しない**（最新は 2025-09-26 提出の
FY2025/3）。これは探索の不具合ではなく**実在しない**ケースで、先行報告の
「6594 は基準年が17か月古い」という観測と一致する。

**(b) 拡大した窓の単体検証**（`_search_single_date` をスタブ化し、2802 の 6/12 提出を再現）

```
found: [('2026-03-31', 'SXXX')]   api calls: 13
order: 06-30, 06-29, 06-26, 06-25, 06-24, 06-23, 06-22, 06-19, 06-18, 06-17, 06-16, 06-15, 06-12
```
旧窓（6/18〜7/2）では**永久に見つからない**日付を、13コールで捕捉した。
期限日から遡る順序になっていることも確認できる。

**(c) 2802 の end-to-end**

```
Found 5 annual report(s):
  [1] docID=S100Y992  period=2026-03-31  filer=味の素株式会社
  [2] docID=S100VXJA  period=2025-03-31
  ...
```
`fundamentals_fy2026` による回避策なしでも EDINET から FY2026/3 が取れるようになった
（当該 override は一次情報の短信由来なのでそのまま残す）。

**(d) 5726 回帰（基準 = #6 適用後）**: **differences: 0**。
docID は 5726 の overrides `_hist_note` が記録している
S100OD4X / S100R2LJ / S100TSNG / S100W7ZO と一致し、最新期は S100YHFZ。

**(e) 所要時間**: 5726 の1件あたり **393s → 81s**（4.8倍）。
Step 1 が探索から取得のみになったため。

---

## §2-4 全85件の β 再導出と overrides の更新

### なぜ必要か

#6 で `beta` の契約が「実測（raw）β を入れる。Blume 調整とクランプはテンプレが行う」に
変わった。既存 overrides は**旧ルール下で作業者が手でクランプ済みの値**であり、
**85件中57件が 0.60（旧クランプ下限そのもの）**だった。これを raw として渡すと
**縮小済みの数値をもう一度縮小する**ことになる。

### 指示からの逸脱（明示） — raw の出所を yfinance から TOPIX 回帰に変更した

指示は「クランプ発動時に raw 値を Adjustments Log に記録済みのものはそれを使用、
未記録は yfinance から再取得」。両者は同じ数値、すなわち yfinance の
`info["beta"]` フィールドに帰着する。**このフィールドは日本株では検証に耐えなかった。**

| ticker | yfinance フィールド |
|---|---:|
| 9432 NTT | **−0.165** |
| 9532 大阪ガス | **−0.201** |
| 9531 東京ガス | **−0.148** |
| 7550 ゼンショー | **−0.078** |
| 4205 | **ちょうど 0.000** |
| 4568 第一三共 / 7974 任天堂 | **null（取得不能）** |

83銘柄中 **7銘柄が負**、平均 +0.443、**85件中57件が 0.6 未満**。
ガス2社と NTT の負のβ、ぴったり 0.000、大型2社の null —— これらは
「日本株の株式リスクの測定値」ではない。Yahoo は**この市場を説明しないベンチマーク**で
当該フィールドを算出している。

そこで **TOPIX（1306.T）に対する2年週次 OLS 回帰**で β を実測し直した。
これは新手法ではなく、**5726 の overrides 自身が文書化している当リポジトリの手法**である
（`_beta_note`: 「yfinanceで 5726.T の週次2年リターンを 1306.T に対して回帰: beta = 1.553」）。
本スクリプトは同じ計算で **1.547** を得る（2週間新しいデータでの再現）。

同日・同一の価格ソースでの比較:

| ticker | yfinance フィールド | TOPIX 回帰 |
|---|---:|---:|
| 9432 NTT | −0.165 | **0.218** |
| 9532 大阪ガス | −0.201 | **0.499** |
| 4205 | 0.000 | **0.852** |
| 4568 第一三共 | null | **0.607** |
| 7974 任天堂 | null | **0.565** |
| 1801 大成建設 | 0.544 | **1.091** |
| 5726 大阪チタ（検証用） | 0.484 | **1.547**（作業者の手回帰 1.553） |

**これは指示の文言からの意図的な逸脱である。** 根拠は指示自身の原則
「推測埋め・捏造・サイレントフォールバックの禁止」で、
ガス会社に −0.20 のβを与えて WACC を組み、それを「実測」と呼ぶことはできない。
**黙って変えてはいない**: 両方の数値を全85銘柄の `_beta_note` に併記し、
本節に記録し、`batch/rederive_beta.py --source yfinance` で指示どおりの再現もできる。

### 実施結果

**`batch/rederive_beta.py`（新規）** — Adjustments Log から旧 raw を読み戻し、
TOPIX 回帰を実行し、両方を記録して overrides を書き換える。

| 指標 | yfinance フィールド | TOPIX 回帰 |
|---|---:|---:|
| 解決できた銘柄 | 83/85 | **85/85** |
| 最小 / 最大 | −0.201 / +1.561 | −0.002 / +2.280 |
| 平均 | +0.443 | **+0.870** |
| 負のβ | **7件** | 1件 |

- **採用β（Blume 調整後）**: min 0.329 / max 1.858 / 平均 **0.913** / **クランプ発動 0 件**。
  旧ルールでは57件が下限に張り付いていたのに対し、新ルールでは**置換が1件も起きない**。
  クランプが「日常経路」から「稀なガード」に戻ったことが実測で確認できた。
- バッチが未解決として残していた **4568 第一三共・7974 任天堂の β も解決**した
  （旧: テンプレ既定 1.0 を「推定」として採用）。
- 全85件の overrides は `scripts/overrides_validator.py` を**違反ゼロ**で通過。
  差分は `beta` の値と `_beta_note` の追加のみ（1801 で差分キーを確認）。

### 回帰の信頼性が低い銘柄（報告事項）

相関 < 0.20 の9銘柄は回帰βの標準誤差が大きく、点推定を鵜呑みにできない。
Blume 調整はまさにこの種の推定を 1.0 方向へ縮める処理であり、
**弱い推定ほど市場βに寄る**という望ましい性質が働く。ただし記録しておく。

| ticker | TOPIX β | 相関 |
|---|---:|---:|
| 9843 ニトリ | −0.002 | **−0.001** |
| 2871 ニチレイ | 0.033 | 0.025 |
| 9020 JR東日本 | 0.075 | 0.063 |
| 2267 ヤクルト | 0.129 | 0.088 |
| 2269 明治HD | 0.124 | 0.109 |
| 9697 カプコン | 0.326 | 0.133 |
| 4661 OLC | 0.306 | 0.174 |
| 9022 JR東海 | 0.285 | 0.191 |
| 2801 キッコーマン | 0.337 | 0.196 |

相関の分布: min −0.001 / p25 0.272 / 中央値 0.415 / max 0.733（n=104 週、全銘柄）。

---

---

# 追補12 §A-3 — 裁定の生成側内蔵（2026-09-06、フェーズ2 の続き）

## 症状

追補6 §X の脚降格は `batch/demote_dcf_leg.py` が**完成したワークブックを後から編集する
後処理**であり、`generate_dcf.py` はその存在を知らない。したがって
**再生成は必ず裁定を取り消す**。フェーズ2 の全件再生成で17件の裁定がすべて失われ、
6857 アドバンテストの Target は 4,288（PGM単独）→ 8,021（中点平均）に戻った。
これは評価の変化ではなく**裁定の不在**である。

## 実装前に判明した3つの制約（実測）

1. **「validate 実行前」は現行コードのままでは実装できなかった。**
   `arbitrate()` の入口 `check11()` が **validate の出力した `<xlsx>_validation.txt` を
   正規表現で読んで**乖離倍率を得ていたため、validate → 裁定 の順にしか動かない。
   → 乖離を**ワークブックのキャッシュ値から直接算出**する実装に変更した
   （`TV_PGM ÷ Year5 EBITDA` と `C14`、validate のチェック11 と同じ算術）。
   テキスト解析依存が消えるので、これは順序の都合を超えた改善でもある。
2. **recalc が2回必要。** 裁定はキャッシュ値（2脚の株価・WACC・ターミナル価値）を読むので
   Step 8 の後でしか動けず、降格は C10 に数式を書くのでキャッシュが消える。
   そのまま validate すると値レベルのチェックが SKIP → フェーズ2 #8 の規則で FAIL になる。
   整合する順序は **生成 → recalc → 裁定 → recalc → validate** の一通りだけ。
3. **型A〜E を判定する材料がパイプライン側に無かった。**
   `batch_state.json` の `type` はバッチ側の記録で `generate_dcf.py` は読まない。
   overrides にも型のキーは無かった（8410 / 7203 / 8267 いずれも）。

## 変更内容

**`scripts/arbitration.py`（新規）** — 判定ロジックを `batch/` からパイプライン層へ移設。

- §X のラダー本体（`arbitrate`）と入力読み取り（`opm_series` / `peer_band` /
  `midcycle_recorded` / `already_treated` / `legs_and_wacc`）、
  および書き換え（`apply_demotion`）を集約。
- **全関数が銘柄コードではなく明示パスを受け取る**。パイプラインは自分の出力ファイルを
  知っており、基準日の解決を必要としないため。
- `divergence(xlsx)`: 乖離をワークブックから直接算出（上記1）。
  キャッシュ値が無い旧モデル用に検証レポートを読む経路は fallback として残した。
- `resolve_company_type()` / `arbitration_applies()`: 型D をスキップし理由を返す。
- `_log_to_adjustments()`: **降格を Adjustments Log シートにも記録**するようにした。
  従来は Executive Summary の B20 注記のみで、Adjustments Log には残っていなかった
  （「現行どおり」との指示だったが、実装を確認したところ残していなかったので追加した）。
  行はテンプレートが自動記録行と Pipeline Metadata の間に残す隙間に置き、
  隙間が無ければ1行挿入する（metadata はタイトルで探索されるため行ずれに強い）。

**`batch/arbitrate_divergence.py` / `batch/demote_dcf_leg.py`** — 薄い CLI に変更。
基準日の解決（`_d()`）とバッチ全体の走査だけを担う。

**`scripts/overrides_validator.py` / `docs/overrides_schema.md`** —
`company_type`（A〜E）を契約に追加。**型は推測しない**: 8410 の
`net_debt=0` / `base_year_ar=inv=ap=0` は「銀行だから」ではなく ΔNWC を強制ゼロにする
実装手段であり、非銀行も同じ形を取りうるため、構造推定はサイレントな推測になる。

**`scripts/generate_dcf.py`** — **Step 8.5** を新設（recalc と validate の間）。
`--no-arbitration`（デバッグ用の逃げ道）を追加。スキップは必ず理由を出す:

| 条件 | 挙動 |
|---|---|
| `--no-arbitration` | スキップ。「validate は裁定未適用として FAIL する」と告知 |
| `--no-recalc` | スキップ。キャッシュ値が無く乖離を算出できないため |
| 型D | スキップ。「DCF が成立せず Target は DDM+RI が主手法」 |
| 型未宣言 | **適用**し、「型D なら `company_type: "D"` を明示すること」と告知 |
| 乖離を算出できない | 中点平均のまま出力し `final_warnings` に記録 |
| 降格適用済 | 再適用しない |
| 降格が必要 | 適用 → **再 recalc**（失敗したら exit 1。数式だけでキャッシュ値が無い状態で終わらせない） |

**`scripts/validate_output.py`** — **チェック22「追補6 §X arbitration applied」**を追加。
機械化ルールが降格を指示しているのに未適用なら **FAIL**。型D は対象外として PASS。
乖離を算出できない場合は SKIP（＝#8 の規則で FAIL。未検証のワークブックは未検証である）。

## 検証

**(a) 6857 アドバンテストの再生成 — 陰性対照と本番の対**

| 実行 | Target | 判定 |
|---|---:|---|
| `--no-arbitration`（陰性対照） | **8,021**（中点平均） | `[FAIL] 22 機械化ルールは「Exit降格 → Target = PGM 単独」と判定しているが未適用` / **exit 1** |
| 通常（裁定内蔵） | **3,632**（PGM 単独） | `[PASS] 22 適用済 追補6 §X: EXIT 脚を降格 (乖離 6.08x, 成長プレミアム(§Q))` / **exit 0** |

生成ログ:
```
[Step 8.5] 裁定(追補6 §X): PGM/Exit の乖離を判定...
  銘柄型: 未宣言 — DCF 型（A/B/C/E）とみなして適用する
  乖離 6.08x（PGM逆算 5.08x / 仮定Exit 30.90x） | レジーム 成長プレミアム(§Q)
  裁定: Exit降格 → Target = PGM 単独
  降格を反映するため再計算...
```
Adjustments Log 7行目に降格が記録される:
`Executive Summary!C10 | Exit法を[参考]に降格 → Target = PGM 単独（C10 は C16 を参照）| 【追補6 §X】… | PGM/Exit の中点平均 | 裁定適用（追補6 §X）`

**指示の「4,288 側」について**: 4,288 は **v1（旧βルール）の PGM 単独値**である。
フェーズ2 の新βルール（raw 1.776 → Blume 1.520）では PGM 脚は **3,632** になる。
本テストで確認できたのは「**Target が中点平均 8,021 ではなく PGM 単独になる**」という
**基礎の一致**であり、4,288 という数値そのものではない（数値はβの変更で動いている）。

**(b) 移設の無影響**: 86銘柄の §X 走査出力が移設前と**集合として完全一致**（95行、
旧のみ/新のみとも0行）。集計8行もすべて一致。差は表示順のみで、乖離が2桁丸めの値から
全精度になったため近接行が入れ替わった。

**(c) 5726 回帰（46セル）**: **differences: 0**。裁定は「裁定不要（乖離 1.37x）」で
何も書き換えない。

**(d) 既存85件への影響**: チェック22 を入れた状態で85件を再検証 → **FAIL 0件**。
5ゲートは引き続き **85/85 clean**。
