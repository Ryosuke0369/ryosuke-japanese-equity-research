# P1完了 → P2: screener データ保存先を D: へ外出し (2026-08-31)

背景: 新PC到着が3週間遅延。C: 空きが約5GBしかない。
EDINET一括 + TDnet年間 + DB/キャッシュ で最低10GB、余裕をみて20GB必要。

## 計画 / 結果
- [x] 1. `screener/common.py` の DATA_DIR を `.env` の DATA_ROOT で外出し
      解決順 `SCREENER_DATA_ROOT` > `DATA_ROOT` > `screener/data`(旧既定)。
      `load_env()` を import 時に実行(パス定数より前)。dotenv 未導入時は
      「.env を読めていない」を初回 log() で警告する(黙って C: に書かない)。
- [x] 2. `run_daily.ps1` の `screener\data\logs` ハードコードを排除。
      Python に `C.LOG_DIR` を問い合わせ、失敗時は throw(黙って別の場所に
      ログを書かない)。PS 5.1 の NativeCommandError を避けるため `2>$null` は使わない。
- [x] 3. `install_task.ps1` の案内文字列も実パス表示に
- [x] 4. `.env.example` を追加(.gitignore は `!.env.example` で既に許可済み)
- [x] 5. `.env` に `DATA_ROOT=D:\screener_data` を追記
- [x] 6. 既存データ移動 440ファイル / 159,154,881 バイト。
      `diff -r` で完全一致・DB の sha256 一致を確認してから C: 側を削除。
- [x] 7. **想定外の作業**: DB の `filings.path/pdf_path/xbrl_path` が
      **リポジトリ相対**(`screener\data\raw\...`)で保存されていた。
      → `C.store_path()` / `C.full_path()` を追加し **DATA_ROOT 相対**に統一。
      `db/migrations/001_paths_relative_to_data_root.py` で既存353行を変換
      (dry-run → 実在チェック0件欠損 → apply → 冪等性確認)。
      これで新PCへ移す時もDB書き換え不要。
- [x] 8. ユニットテスト31件 OK。タスクスケジューラ `Start-ScheduledTask` で
      実行 → LastTaskResult=0、ログは `D:\screener_data\logs\` に出力、
      当日(08-31)分3件を新パスへ取得、`filings` の新規行もDATA_ROOT相対。
      C: 側に `screener/data` は再生成されていない。
- [x] 9. D: 空き **64.91 GiB** (総容量 79.77 GiB / 使用 14.86 GiB)。目標20GBに対し十分。
- [ ] 10. `--source jquants --build` → **J-Quants API が全面 403 でブロック中(未達)**。
      正しいパスワードでも誤パスワードでも同一の 403 ForbiddenException が返るため
      認証情報の問題ではない(認証評価の手前で弾かれている)。UA を4種類変えても同じ。
      `https://jpx-jquants.com/` 自体は 200。→ アカウント/プラン状態の確認が必要。
      代替として `--source jpx --build` を実行し、業種・市場区分の除外内訳のみ確定。
      規模(時価総額)・流動性(ADV20)の判定は J-Quants 復旧まで保留。

## ユニバース現状 (2026-08-31, JPX data_j.xls as of 20260731)

| 区分 | 社数 |
|---|---|
| companies 総数 | 4,451 (jpx 4,444 + tdnet 由来 7) |
| **universe_flag=1** | **0** ← 規模・流動性データが1社も無いため確定不能 |
| 業種・市場区分で除外 | 858 |
| 規模・流動性 判定保留 | 3,593 |

除外内訳(858):

| 理由 | 社数 |
|---|---|
| 市場区分除外 ETF・ETN | 476 |
| 市場区分除外 PRO Market | 185 |
| 業種除外 銀行業 | 79 |
| 市場区分除外 REIT・ベンチャー/カントリー/インフラファンド | 63 |
| 業種除外 証券、商品先物取引業 | 37 |
| 業種除外 保険業 | 14 |
| 市場区分除外 出資証券 | 2 |
| コード帯除外(市場区分不明のETF/REIT帯) | 2 |

判定保留 3,593 の市場区分内訳: スタンダード1,534 / プライム1,459 / グロース590 /
外国株式5 / 市場区分不明5。ここに時価総額50〜600億円 + ADV20 3,000万円以上を
かけた残りが最終ユニバースになる。

## Review

- コード側のパスは元々 `common.py` に集約されていたため、外出し自体は10行程度で済んだ。
  想定外だったのは **DBに保存されたパスがリポジトリ相対だった**こと。
  ここを DATA_ROOT 相対に変えたので、3週間後の新PC移行は
  「データをコピー + .env の1行を書き換える」だけになる。
- TDnet 日次アーカイブは新パスで完全に動作。カバレッジに欠損営業日なし。
- 唯一の未達は J-Quants。これは当方のコードでは解決できない(API側で遮断)。
