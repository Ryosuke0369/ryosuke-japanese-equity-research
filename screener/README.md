# screener — 傾き検出スクリーナー

仕様書: `docs/傾き検出スクリーナー仕様書_v1_20260829.md`
既存パイプライン(逆算DCF)との接続は仕様書 §8。

## いま動くもの

| フェーズ | 内容 | 状態 |
|---|---|---|
| P0 | TDnet日次アーカイバ + DB + タスクスケジューラ | 稼働中 |
| P1 | EDINET一括取得 / J-Quantsユニバース | 試走完了。全ユニバース夜間実行は未実施。J-Quantsは資格情報待ち |
| P2 | 単独値ビルダー + S1〜S5 | 未着手 |
| P3 | 検証8銘柄バックテスト | 未着手 |
| P4 | 週次レポート + S6/S7/S8 | 未着手 |

## いちばん大事な運用上の事実

**TDnetの公開一覧は約1ヶ月しか残らない。** 取り逃した日は二度と取れない。
だからこのシステムで唯一「落としてはいけない」のは日次アーカイブであり、
`fetch_runs` テーブルは「その日を取得したか」を必ず1行残すために存在する。

- `ok` … 取得成功
- `empty` … 正常に取得できて対象0件(休日・開示ゼロ)。**欠損ではない**
- `partial` … 一覧は取れたが一部ダウンロード失敗
- `failed` … 一覧自体が取れなかった。**欠損として再取得対象に残る**

`--backfill N` は「直近N日の平日のうち ok/empty/partial の記録が無い日」を
すべて取りに行く。1週間PCを止めていても自力で埋まる。

## 使い方

```bash
# 取得(日次)
python -m screener.fetch.tdnet_archiver --today
python -m screener.fetch.tdnet_archiver --date 20260828
python -m screener.fetch.tdnet_archiver --days 3          # 直近3営業日
python -m screener.fetch.tdnet_archiver --backfill 14     # 欠損日だけ埋める
python -m screener.fetch.tdnet_archiver --report 14       # カバレッジ確認のみ

# 抽出(短信XBRL → financials_cum / guidance)
python -m screener.extract.xbrl_parser --all
python -m screener.extract.xbrl_parser --coverage         # §3-1 科目カバレッジ
python -m screener.extract.xbrl_parser --unknown-top 40   # 未マッピングタグ頻度

# ユニバース / EDINET(P1)
python -m screener.fetch.jquants_universe --build
python -m screener.fetch.edinet_bulk --trial              # 検証8銘柄+50社の試走
python -m screener.fetch.edinet_bulk --full --years 3     # 夜間の全ユニバース

# テスト
python -m unittest discover -s screener/tests -t .
```

## 自動実行(Windows タスクスケジューラ)

```powershell
powershell -ExecutionPolicy Bypass -File screener\install_task.ps1          # 登録
powershell -ExecutionPolicy Bypass -File screener\install_task.ps1 -RunNow  # 即実行
powershell -ExecutionPolicy Bypass -File screener\install_task.ps1 -Remove  # 解除
```

タスク名 `ScreenerTdnetArchiver`。平日 **19:00** と **23:15** の2回。
2回にしている理由: TDnetの開示は22:30過ぎまで出る(2026-08-28に実測)。
19:00の1回だけだと遅い開示を構造的に取り逃し、1ヶ月後には回復不能になる。
アーカイバは冪等なので2回目は差分だけを取る。

`-StartWhenAvailable` を付けてあるので、19:00にスリープしていたPCは
起動時にその日の分を実行する。

## ディレクトリ

```
screener/
├── common.py           # パス/DB/スロットル付きHTTP/欠損判定
├── fetch/              # tdnet_archiver.py, edinet_bulk.py, jquants_universe.py
├── extract/            # xbrl_parser.py (quarterly_builder.py は P2)
├── signals/            # P2
├── report/             # P4
├── db/schema.sql       # 仕様書 §5。§5に無い列・表は【§5拡張】と明記
├── config/             # universe_rules.yaml, account_mapping.yaml
├── db/migrations/      # 001_paths_relative_to_data_root.py
└── tests/
```

## データ保存先 (DATA_ROOT) — 2026-08-31 に D: へ移設、2026-09-09 の PC移行で C: へ

取得データはリポジトリの外に置く。場所は**リポジトリルートの `.env`** で決める:

```
DATA_ROOT=C:\screener_data
```

```
C:\screener_data\
├── raw\tdnet\YYYYMMDD\     TDnet 短信PDF + XBRL zip
├── raw\edinet\YYYY-MM-DD\  EDINET 生XBRL zip
├── cache\                   J-Quants トークン (認証情報。持ち出し厳禁)
├── logs\                    screener_YYYYMM.log / run_daily_YYYYMM.log
└── screener.db              SQLite
```

- 解決順は `SCREENER_DATA_ROOT` > `DATA_ROOT` > `screener/data`(旧既定)。
  実装は `common.py:_resolve_data_root()` の**1箇所だけ**。
  他モジュールは `C.RAW_DIR` 等を使い、自前でパスを組み立てないこと。
- `.env` から読むので、**タスクスケジューラの定義に保存先を書く必要はない**。
  `run_daily.ps1` / `install_task.ps1` もログ先を Python に問い合わせる。
- DB の `filings.path / pdf_path / xbrl_path` は **DATA_ROOT 相対**で保存する
  (`C.store_path()` / `C.full_path()`)。ドライブを移してもDB書き換えは不要。
  リポジトリ相対で書かれた旧行は `db/migrations/001_...` が変換する。

移設理由: C: の空きが約5GB しかなく、EDINET 一括取得(下記見積り)が入らない。

`DATA_ROOT` 配下はgit管理外。**アーカイブはリポジトリの成果物ではなくローカル資産**で、
消すと(1ヶ月より前の分は)復元できない。バックアップ対象にすること。

## P1 の現状(2026-08-29)

- **EDINET**: 日付スイープ方式で試走完了。直近12ヶ月・検証8銘柄+50社で
  索引261営業日(84,382件を走査、128件が対象)→ 125件ダウンロード、失敗0。
  検証8銘柄は全社が有報1+半期1で揃った。
  全ユニバース3年の見積り: 索引16分 + ダウンロード約9.3時間 / 約12.6GB。
- **J-Quants**: 403 の原因は API の **V1 終了 (2026-06-01)** だった。トークンでは
  なく世代の問題で、`/v1/token/auth_user` `/v1/token/auth_refresh` は廃止済み。
  V2 はダッシュボード発行の API キーを `x-api-key` ヘッダーで送る方式なので、
  `.env` に `JQUANTS_API_KEY` を置くだけでよい (`JQUANTS_MAIL` /
  `JQUANTS_PASSWORD` / `JQUANTS_REFRESH_TOKEN` は不要)。
  変わったのは認証だけではない —— パス (`/listed/info` → `/equities/master`、
  `/prices/daily_quotes` → `/equities/bars/daily`)、本体キー (`data` に統一)、
  項目名 (`CompanyName`→`CoName`、`Close`→`C`、`TurnoverValue`→`Va` 等) も総取替。
  **レートリミットが分あたりになった** (Free 5 / Light 60 / Standard 120 /
  Premium 500 req/min)。`--rpm` を契約プランに合わせること (既定 5)。
  疎通確認は `python -m screener.fetch.jquants_universe --check-auth`。
- **ユニバース骨格**: J-Quantsが無い間は JPX「東証上場銘柄一覧」で
  コード・銘柄名・市場区分・33業種まで構築済み(4,444社)。
  時価総額・売買代金が無い3,583社は `universe_flag=0` かつ
  `exclude_reason='規模・流動性データ未取得'` ——「条件を満たさない」ではなく
  「まだ判定していない」として区別してある。J-Quantsが通れば
  `apply_universe_rules()` を流し直すだけで確定する。

## 設定を変えるとき

閾値・科目・除外条件はコードではなく `config/*.yaml` にある。

- `universe_rules.yaml` … 時価総額・売買代金・除外業種
- `account_mapping.yaml` … XBRLタグ → 内部項目。**会社間の科目名ゆらぎはここで吸収**

マッピングできなかったタグは捨てずに `unknown_tags` に頻度が積まれる。
`--unknown-top` で上位を見て、意味が確定したものだけを `account_mapping.yaml` に
足していく——これが仕様書 §3-1 の「定期レビュー」の実体。
