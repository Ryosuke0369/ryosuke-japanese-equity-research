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
