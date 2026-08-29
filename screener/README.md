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
├── data/               # raw/(取得原本) cache/ logs/ screener.db  ← 全て .gitignore
└── tests/
```

`data/` 配下はgit管理外。**アーカイブはリポジトリの成果物ではなくローカル資産**で、
消すと(1ヶ月より前の分は)復元できない。バックアップ対象にすること。

## P1 の現状(2026-08-29)

- **EDINET**: 日付スイープ方式で試走完了。直近12ヶ月・検証8銘柄+50社で
  索引261営業日(84,382件を走査、128件が対象)→ 125件ダウンロード、失敗0。
  検証8銘柄は全社が有報1+半期1で揃った。
  全ユニバース3年の見積り: 索引16分 + ダウンロード約9.3時間 / 約12.6GB。
- **J-Quants**: `.env` の `JQUANTS_REFRESH_TOKEN` が HTTP 403 (Forbidden)。
  43文字で、正規のリフレッシュトークン(数百文字のJWT、有効期間1週間)ではない。
  取得しなおして `.env` に入れれば `--source jquants --build` がそのまま通る。
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
