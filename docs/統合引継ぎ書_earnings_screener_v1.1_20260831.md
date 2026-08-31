# 統合引継ぎ書（Claude Code 向け作業指示書）v1.1

作成: 2026-08-31 / 作成側: Kimi（Module A〜D 実装側）
宛先: 本体パイプライン側（Claude Code / ryosuke-japanese-equity-research）

同梱物: `earnings_screener/` ディレクトリ一式（モックデータで全モジュールが動作する状態）

---

## 0. この文書の使い方

この文書は「Kimi側で実装した決算前投資自動化システム（Module A〜D）」を本体パイプラインに統合するための作業指示書です。**丸ごとマージはしないでください**。§1の移植可否表に従い、§3の接続契約だけを満たす形で本体に組み込んでください。

## 0-1. v1.1 変更履歴（Claude Code 検証で指摘された3件の相違への対応）

2026-08-31 の Claude Code による突合検証で発見された相違3件は、すべて v1.1 で修正済みです。

| 指摘 | 内容 | v1.1での対応 | 回帰テスト |
|:---|:---|:---|:---|
| 相違①（最重大） | `forecaster.py` が as_of を受け取りながら5箇所のデータアクセスに渡しておらず、機械予測が未来データを見ていた。`forecast_snapshots` 経由で Module C の凍結基準も破れていた | 全アクセスに as_of を伝播。証拠イベントに `event_date <= as_of` の上限も追加 | `tests/test_pit.py` T2 |
| 相違② | `weekly_screener.py` が `score_ticker` に as_of を渡さず、過去日の再現実行で静かにルックアヘッド | as_of を伝播 | 同 T3 |
| 相違③ | `filings`/`quarterly_standalone` が `UNIQUE(ticker, period_end)` で訂正世代を保持不可能。訂正後の値が as_of 以前の判定に混入 | 全ファクト系テーブルに `generation` 列を導入し訂正版を別世代保持（上書き禁止）。`data_access.visible_generations()` が「as_of時点で発表済みの最新世代」を解決。モックに3905の訂正短信デモを追加 | 同 T1（as_of=3/9 は訂正前1400、as_of=3/10 は訂正後1350を返す） |

あわせて以下も修正:
- **S5のDB直叩き違反**（「スコアラーはDBを直接叩かない」原則の唯一の例外）を解消：`get_latest_forecast()` を data_access に追加し経由化
- **シグナルIDの安定化**：スコアラー辞書のキーを日本語ラベル（"S1_DSO改善"等）から安定ID（"S1"〜"S8"）に変更し、表示名は `SIGNAL_LABELS` に分離。`stop_loss_conditions()` の文字列依存を解消
- **訂正短信の混入防止**：決算カレンダー推定・バックテストのイベント列・モック株価ジャンプはすべて `generation=1`（初回発表）のみを使用

なお相違③について、Claude Code の指摘「本体の `financials_cum`（filing_id, item, context_ref 主キー）を使うべき」は妥当です。**実運用のファクト層は本体構造を正とし、v1.1の generation 機構はモック側でのPIT検証用**と位置づけてください。

---

## 1. 移植可否の一覧

| コンポーネント | パス | 移植判定 |
|:---|:---|:---|
| 日本営業日カレンダー | `common/jp_calendar.py` | **移植可**（外部依存なし。祝日は近似式。SQ日・特別休場・制度変更は未考慮） |
| DBスキーマ | `common/schema.py` | **参照**（本体スキーマへの写像表として使用。§3） |
| 決算カレンダー | `module_a/` | **移植可**（本体に無い層。入力を実filingsに向ける） |
| データアクセス層 | `module_b/data_access.py` | **最重要の接続点**（§3参照。本体ビルダー出力に向き替える） |
| スコアラー S1〜S8 | `module_b/scorers_*.py` | **本体が正**。ただしCF品質・収益認識・希薄化の3指標は本体に無いため参考実装として移植価値あり |
| 機械的予測 A+B+C | `module_b/forecaster.py` | **移植可**（本体に相当機能なし。v1.1でPIT完全対応） |
| 週次選定・レポート | `module_b/weekly_screener.py` | **移植可**（本体P4の参考形式） |
| 出口判定エンジン | `module_c/exit_engine.py` | **移植可**（`decide_exit()` はDB非依存の純粋関数で結合度が低い。本体に出口定義が無い空白を埋める。§5） |
| PITバックテスト | `module_d/backtest.py` | **移植可**（PIT機構は本体に無い生命線。ただし§4の改良が必要） |
| モック一式 | `mock/` + `tests/` | **回帰テスト資産として保持**（実データ接続後も削除しない。`tests/test_pit.py` はPIT機構の不変条件を固定） |
| モックのバックテスト数値 | `data/backtest_*.csv` | **移植しない**（相関を注入したモック世界の結果。エッジの証明ではない。git管理外推奨） |

---

## 2. 実行環境・依存

- Python 3.11+、外部依存は標準ライブラリのみ（検証スクリプト `check_*.py` のみ pandas 使用。本体移植時は不要）
- DBは SQLite 単一ファイル。ORM不使用
- モック環境の全再現コマンド（この順序で依存関係が通る）:

```bash
python mock/generate_mock_data.py        # ユニバース＋開示履歴
python module_a/weekly_batch.py          # 決算カレンダー＋T-15候補
python mock/generate_mock_financials.py  # 財務モック（全ユニバース。訂正短信デモ含む）
python mock/generate_mock_s7_s8.py       # S7/S8補助データ
python mock/generate_mock_rdcf.py        # 逆算DCF要求値モック
python mock/generate_mock_prices.py      # 株価モック
python tests/test_pit.py                 # PIT回帰テスト（T1-T3 全PASS必須）
python module_b/weekly_screener.py       # 週次選定（上位20社CSV）
python mock/generate_mock_actuals.py     # 決算着弾モック
python module_c/exit_engine.py           # 出口判定
python module_d/backtest.py              # バックテスト
```

---

## 3. 接続契約（統合の核心）

Kimi側の全モジュールは `module_b/data_access.py` の関数だけを経由してデータを読みます。本体統合は**この層を本体の単独値ビルダー出力に向き替える作業に集約**されます。

### 3-1. インターフェース仕様

```python
get_pl_series(conn, ticker, as_of=None) -> list[dict]
# 戻り値キー: period_end, quarter_type('1Q'/'2Q'/'3Q'/'FY'), fiscal_year,
#            sales, operating_profit, gross_profit,
#            cogs_quantity, cogs_price, operating_cf（1Q/3QはNoneになり得る）
# 単位: 百万円。四半期単独値（累計ではない）。is_valid=1 かつ可視世代のみ。古い順。

get_bs_series(conn, ticker, item_key, as_of=None) -> list[(period_end, value)]
# item_key の語彙（本体 account_mapping.yaml との写像が必要）:
#   accounts_receivable / contract_liabilities / construction_in_progress /
#   machinery / inventory / deposits_received / advances_paid /
#   rev_over_time / rev_point_in_time / shares_outstanding / diluted_shares
# ※本体語彙例: trade_receivables / machinery_and_equipment / inventories_total

get_adjustments(conn, ticker, as_of=None) -> dict[period_end, 控除合計額]
# 一時収入等の調整。normalized_sales = sales - adjustments[period_end]

get_latest_forecast(conn, ticker, as_of=None) -> Row | None
# 会社通期予想の最新版（as_of指定で当時の版を再現）。v1.1で追加。
# スコアラー・予測器は company_forecasts を直接叩かず必ずこれを経由する
```

### 3-2. as_of（Point-in-Time）の意味論 —— 最重要

- 全関数の `as_of` は「**その日時点で公知だった期だけを返す**」フィルタ
- 判定基準は `period_end` ではなく **`filings.filing_date <= as_of`**（期末の値は発表日に初めて公知になる）
- **訂正世代**：同一 `period_end` に複数世代がある場合、`visible_generations()` が「as_of時点で発表済みの最新世代」を解決する（v1.1で実装）。本体側は `financials_cum` の filing_id 主キーで同等以上のことが自然にできる
- 本体統合時はこの経由を**構造的に強制**すること（スコアラーがDBを直接叩けないようにする）
- 注意1: TDnetは引け後発表が多い。発表日「当日」の判定に当日発表の値を使わない運用（Kimi側は判定を発表日の1営業日前までに行うことで回避）
- 注意2: `forecaster` / `weekly_screener` は v1.1 で as_of 完全対応済みだが、`earnings_calendar` は「実行時点で再構築済み」が前提。**過去日の再現実行では先に `build_calendar(db, as_of)` で当時のカレンダーを再構築すること**

### 3-3. 本体側が用意すべき4テーブル相当

| テーブル | 内容 | 本体での出所 |
|:---|:---|:---|
| quarterly_standalone 相当 | 単独PL+CF+COGS分解+is_validフラグ | 四半期単独値ビルダー出力 |
| balance_sheet_items 相当 | 期×科目のスナップショット | XBRLパーサー＋account_mapping |
| pl_adjustments 相当 | 一時収入の控除（§6-1参照。本体に未存在の層） | **新設が必要** |
| company_forecasts 相当 | 通期予想の履歴（source_dateで版管理） | 短信XBRLの予想科目 |

---

## 4. 既知の弱点・統合判断事項（そのまま移植しない箇所）

### 4-1〜4-4. 弱点（改良指示）

1. **前年同期の取り方が位置ベース（`pl[-5]`）**。is_valid=0 で期が欠けると前年同期でなくなる。→ 本体では `(quarter_type, fiscal_year - 1)` の**キー突合**に変更すること（対象: scorers_s1_s4.py の S1/S2、scorers_s5_s8.py の S6 等）。**v1.1でも未修正のまま残してある**
2. **バックテストの評価が絶対リターンのみ**。出口を T+20/T+60 に拡張する場合は市場指数・セクター指数に対する**超過リターン評価**を追加すること
3. **決算後エントリー（T+1/T+5）をグリッドに載せる場合**、入口条件関数を差し替えること。決算前=証拠スコア、決算後=Module C の4分岐発火（「予想超過＋上方修正」）を転用するのが自然
4. **祝日カレンダーは近似式**（2000-2099年）。SQ日・特別休場は未考慮

### 4-5〜4-8. 統合判断事項（Claude Code 検証で明らかになった結合度）

5. **スコア意味論の相違**：Kimi側は「利用可能な指標の連続値（-1..+1）の単純平均」、本体は「発火 bool＋重み付き和」。`SCORE_THRESHOLD=0.10` と `total = evidence_score × gap_factor` は連続値前提。→ **本体スコアラーに連続値の互換レイヤー（強度を -1..+1 に正規化）を持たせることを推奨**。bool だけでは週次20社の順位付けができない
6. **item_key 語彙の不一致**：写像表1枚を config に追加（例: `trade_receivables → accounts_receivable`、`machinery_and_equipment → machinery`、`inventories_total → inventory`）
7. **四半期表現の相違（要検証）**：Kimi側は '1Q'〜'FY' 文字列、本体は `q_no 1..4 + span_q`。**`span_q=2`（半期粒度の単独値）はKimi側に対応概念がなく、そのまま渡すと6ヶ月の値を四半期として扱う**。→ アダプタで `span_q` を露出し、当面は span_q=1 のみを評価対象とし、span_q=2 は「半期値」として別フラグで扱う設計が必要。**本体P2完了（3441での単独値検証）前に繋がないこと**（Claude Code の §6 指摘に同意）
8. **日次株価2年分の欠如**：バックテストは `daily_prices` の2年分を必要とするが、本体は1日分のみ（J-Quants無料枠の制約）。**バックテストの実データ化は株価蓄積の完了が前提**。それまではモック株価ジェネレーターで計測機構のみ検証する運用

---

## 5. 出口判定（Module C）の最小要件定義

本体仕様書に出口定義が無い状態から起こすための最小要件:

1. **比較基準の凍結**: 発表前に機械予測を `forecast_snapshots` に保存（週次スクリーナーが自動記録）。事後の握り直しを構造的に禁止。v1.1でスナップショット自体のPITも担保済み
2. **実績の正規化**: 短信XBRL→累計→単独変換→**調整レイヤー経由**（一時収入を除く）
3. **4分岐ルール**（順序固定、ミス判定最優先）:
   - 売上 < 予想95% → 即損切り（全量）
   - 売上・利益とも予想超過＋ガイダンス上方修正 → 50%利確・残りT+3追跡
   - 95〜105%レンジ → 全利確
   - それ以外 → 撤退
4. **ガイダンス比較の定義**: 通期営業利益予想の新旧比較（新>旧=上方修正）。FY発表時は翌期初値の扱いを別途定義すること（7月決算銘柄の通期発表では当該年度の期初予想が存在し得ない）
5. **執行ルール**: PTS出来高/直近20日平均 ≥ 30% → PTS売却、未満 → 翌日寄り指値。**PTS出来高のデータソースは未解決**。取れなければ翌日寄り指値固定のfallback運用
6. **判定不能時のエスカレーション**: is_valid=0（期変更・訂正・連結範囲変更）が絡む期は機械判定を放棄し手動判断へ

---

## 6. 設計意図（文書化されていない部分の補足）

### 6-1. pl_adjustments（一時収入分離層）—— 本体に無い最重要の追加層

3905 データセクションの手数料収入55.8億円のような一時項目は、単独値ビルダーが正しくても傾きを汚染する。Kimi側は**全スコアラーと機械予測が「調整後売上（sales − one_time_revenue）」を使う**構造にしてある（`normalized_sales()` 参照）。調整後営業利益は「一時収入はOPに全額フロー」仮定（注記で利益インパクトが分離できる場合は `one_time_gain` 等の別 item_key で上書きする運用）。本体でも単独値ビルダーの後段にこの層を1枚挟むこと。

### 6-2. 逆算DCFとの接続契約

`reverse_dcf_requirements(ticker, as_of_date, required_steady_op_profit, target_year)` の1テーブルが接点。本体DCFパイプラインはこの4値を書き込むだけでよい。織り込み度ギャップ係数 = `clamp(1 + (機械予測の定常OP − 要求値)/|要求値|, 0.2, 2.0)` は仮置き（バックテストで校正）。

### 6-3. モックの相関注入

`mock/generate_mock_prices.py` は財務パターン品質とイベントジャンプを相関させてある（証拠スコアに予見力がある世界の模擬）。これは**バックテストの計測機構の検証用**であり、実データ化後も回帰テスト資産として保持すること（本体の46セル差分ゼロ回帰テストと同思想）。

### 6-4. 「約束 vs 証拠」の機械的実装

`evidence_events.evidence_flag`（1=証拠/0=約束）。MOU・計画・目標は flag=0 で保存のみ、機械予測C経路は flag=1 の金額のみ積む。認識率パラメータ（order 0.6 / contract 0.6 / facility_operation 0.8）は**根拠のない仮置き**であり、重み調整フェーズで最初に感度分析すべき対象。

---

## 7. シグナル体系の統合マッピング（ID不一致の解消）

両者は同じ「S1〜S8」の名前で中身が異なる。統合後は以下の一本化を提案（v1.1でKimi側の辞書キーは安定ID化済み。本体IDへの付け替えは辞書の定義1箇所で完了する）:

| 統合ID | 内容 | 実装の出元 |
|:---|:---|:---|
| S1 | Q単独粗利率の傾き | 本体 |
| S2 | 在庫の価格/数量分解（両義判定） | 本体（Kimi側 scorers_s5_s8.S6 も参照実装） |
| S3/S3b | 建設仮勘定（急増/振替検出） | 本体（Kimi側 scorers_s1_s4.S3 はS3b相当） |
| S4 | 契約負債・前受金の増加 | 本体（Kimi側 scorers_s1_s4.S2 も参照） |
| S5 | 転換＋進捗率異常 | 本体（Kimi側 scorers_s5_s8.S5 は進捗率部分の参照。季節性補正・調整後ベースの考え方を併合） |
| S6 | 文言差分 | 本体（P4） |
| S7 | 修正方向の転換 | 本体（P4） |
| S8 | 売上債権の質（逆） | 本体（Kimi側 scorers_s1_s4.S1 はDSO改善として表裏の関係） |
| S9 | 営業CF vs 営業利益（キャッシュ品質） | **Kimi由来の追加**（scorers_s1_s4.S4） |
| S10 | 収益認識パターン変化（一時点→一定期間） | **Kimi由来の追加**（scorers_s5_s8.S7） |
| S11 | 希薄化リスク（潜在株式オーバーハング） | **Kimi由来の追加**（scorers_s5_s8.S8） |

---

## 8. 要キャリブレーションのパラメータ一覧（バックテストで決める値）

| パラメータ | 現値 | 所在 |
|:---|:---|:---|
| エントリースコア閾値 | 0.10 | module_d/backtest.py `SCORE_THRESHOLD` |
| DSO改善のフルスケール | ±20% | scorers_s1_s4.S1 |
| 契約負債の少額無効ライン | 30百万円 | scorers_s1_s4.S2 |
| CIP振替の稼働率仮定 | 30% | module_b/forecaster.py |
| 証拠イベント認識率 | 0.6/0.6/0.8 | module_b/forecaster.py `C_RECOGNITION` |
| 証拠分の限界利益率 | 40% | module_b/forecaster.py |
| ギャップ係数クリップ | 0.2〜2.0 | module_b/weekly_screener.py |
| 売上ミスライン / インラインレンジ | 95% / 105% | module_c/exit_engine.py |
| PTS出来高閾値 | 30% | module_c/exit_engine.py |
| ポジションサイズ（MDD計算用） | NAVの10% | module_d/backtest.py `POS_SIZE` |

---

## 9. 統合の推奨手順（Claude Code への作業順序）

0. **整備**：重複フォルダの削除、`data/*.csv` と `*.db` は git 管理外（`.gitignore` に追加）、コードには `PROVENANCE.md` を添えてから追跡（同梱済み）
1. `common/jp_calendar.py` をそのまま取り込み、単体テスト（祝日・振替の既知日で検証）
2. 本体の単独値ビルダー出力を調査し、§3-1 のインターフェースを満たすアダプタの**設計案を先に提示**（item_key 写像表・span_q の扱い・連続値互換レイヤーを含む。実装は設計レビュー後）
3. §3-2 の PIT 規則（filing_date 基準・訂正世代）を本体のファクト層に実装。`tests/test_pit.py` と同等の不変条件テストを本体側にも作ること
4. §6-1 の調整レイヤー（pl_adjustments 相当）を新設
5. `module_a`（決算カレンダー）を本体 filings に接続し、精度を前向き検証する記録を開始
6. **本体P2完了後に** `module_d`（バックテスト）を §4-1/4-2 の改良付きで移植し、実データで中核仮説（S2/S3/S4発火→次四半期PL改善の先行性）を検証（日次株価2年分の蓄積が前提、§4-8）
7. `module_c`（出口判定）を §5 の要件定義とともに本体仕様書へ追記の上で実装
8. `module_b/weekly_screener.py` を本体スコアラーに接続して週次運用開始（P4）
