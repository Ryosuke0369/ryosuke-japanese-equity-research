-- screener/db/schema.sql
-- 傾き検出スクリーナー DBスキーマ。仕様書 §5 に準拠。
--
-- 規約:
--   * §5 が定めた列は名前・意味を変えない。実装上必要になった列は §5 の列の
--     「後ろ」に追加し、各テーブルのコメントで【§5拡張】と明示する。黙って足さない。
--   * 取得の失敗・欠測は必ず行として残す(fetch_runs)。仕様書 §2-1
--     「取得失敗日はDBに欠損記録(黙って飛ばさない)」。
--   * マッピング不能タグは捨てず unknown_tags に頻度を積む。仕様書 §3-1。

PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

-- ---------------------------------------------------------------- companies
-- §5: companies(code, name, market, sector, mktcap, adv20, universe_flag)
CREATE TABLE IF NOT EXISTS companies (
    code            TEXT PRIMARY KEY,          -- 4桁 or 3桁+英字 (7203 / 285A)
    name            TEXT,
    market          TEXT,                      -- プライム/スタンダード/グロース 等
    sector          TEXT,                      -- 33業種名
    mktcap          REAL,                      -- JPY mn
    adv20           REAL,                      -- 20日平均売買代金 JPY mn
    universe_flag   INTEGER DEFAULT 0,         -- 1 = universe_rules.yaml を満たす
    -- 【§5拡張】
    sector17        TEXT,
    scale_category  TEXT,
    exclude_reason  TEXT,                      -- universe_flag=0 の理由(銀行/流動性不足等)
    source          TEXT,                      -- jquants / jpx / tdnet
    updated_at      TEXT
);

-- ------------------------------------------------------------------ filings
-- §5: filings(id, code, date, type[短信/修正/有報/半期], source[tdnet/edinet], path, xbrl_ok)
CREATE TABLE IF NOT EXISTS filings (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    code            TEXT,
    date            TEXT,                      -- YYYY-MM-DD (開示日)
    type            TEXT,                      -- 短信 / 修正 / 有報 / 半期 / その他
    source          TEXT,                      -- tdnet / edinet
    path            TEXT,                      -- 主たる保存物(XBRLがあればそれ、無ければPDF)
    xbrl_ok         INTEGER DEFAULT 0,         -- 1 = XBRLを保存できた
    -- 【§5拡張】
    doc_id          TEXT,                      -- TDnet: 140120260828527542 / EDINET: S100XY8O
    title           TEXT,
    disclosed_at    TEXT,                      -- YYYY-MM-DD HH:MM
    subtype         TEXT,                      -- 決算短信 / 業績予想修正 / 配当予想修正 ...
    pdf_path        TEXT,
    xbrl_path       TEXT,
    market_place    TEXT,                      -- TDnet の「東」等
    company_name    TEXT,
    fetched_at      TEXT,
    UNIQUE (source, doc_id)
);
CREATE INDEX IF NOT EXISTS ix_filings_code_date ON filings (code, date);
CREATE INDEX IF NOT EXISTS ix_filings_date      ON filings (date);
CREATE INDEX IF NOT EXISTS ix_filings_type      ON filings (type, subtype);

-- ----------------------------------------------------------- financials_cum
-- §5: financials_cum(filing_id, code, period, q_no, item, value) — 累計原本
CREATE TABLE IF NOT EXISTS financials_cum (
    filing_id       INTEGER,
    code            TEXT,
    period          TEXT,                      -- 会計期間ラベル (FY2026/3 等)
    q_no            INTEGER,                   -- 1..4 (4 = 通期), NULL = 不明
    item            TEXT,                      -- account_mapping.yaml の内部項目名
    value           REAL,
    -- 【§5拡張】
    unit            TEXT,                      -- JPY / JPY mn / shares / pure
    context_ref     TEXT,
    source_tag      TEXT,                      -- 元のXBRL要素名(追跡用)
    PRIMARY KEY (filing_id, item, context_ref),
    FOREIGN KEY (filing_id) REFERENCES filings (id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS ix_cum_code_item ON financials_cum (code, item);

-- 【§5拡張・docs/segment_dimension_design.md】
-- 次元(セグメント等)付きファクト。financials_cum は「見出し数値だけ」という
-- 不変条件を保ちたいので別テーブルにする。cum に segment 列を足すと、全ての
-- 既存クエリが WHERE segment IS NULL を要求するようになり、書き忘れた瞬間に
-- 集計が黙って2〜3倍になる。分けておけばその間違いは構造的に起こらない。
CREATE TABLE IF NOT EXISTS financials_dim (
    filing_id       INTEGER,
    code            TEXT,
    period          TEXT,
    q_no            INTEGER,
    item            TEXT,                      -- account_mapping.yaml の内部項目名
    axis            TEXT,                      -- segment / segment_total / adjustment / other
    member          TEXT,                      -- 正規化後 (Japan / Philippines)
    member_raw      TEXT,                      -- 元文字列。正規化を後から検証できる
    value           REAL,
    unit            TEXT,
    context_ref     TEXT,
    source_tag      TEXT,
    valid_flag      INTEGER DEFAULT 1,         -- 0 = セグメント区分変更で時系列が切れている
    invalid_reason  TEXT,
    PRIMARY KEY (filing_id, item, context_ref),
    FOREIGN KEY (filing_id) REFERENCES filings (id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS ix_dim_code_item ON financials_dim (code, item, member);
CREATE INDEX IF NOT EXISTS ix_dim_axis ON financials_dim (axis);

-- ------------------------------------------------------------- financials_q
-- §5: financials_q(code, period, q_no, item, value, valid_flag) — 単独値(生成)
CREATE TABLE IF NOT EXISTS financials_q (
    code            TEXT,
    period          TEXT,
    q_no            INTEGER,
    item            TEXT,
    value           REAL,
    valid_flag      INTEGER DEFAULT 1,         -- 0 = 決算期変更/遡及修正/連結範囲変更で無効
    -- 【§5拡張】
    invalid_reason  TEXT,
    -- この単独値が何四半期ぶんか。1 = 真の四半期単独値。
    -- 短信が揃っていない会社は EDINET の有報(q4累計)と半期(q2累計)しか無く、
    -- 差分は「下期6ヶ月」になる。それを四半期と名乗らせないための列 ——
    -- 3ヶ月と6ヶ月を同じ土俵に載せると傾きの大きさが二重になる。
    span_q          INTEGER DEFAULT 1,
    built_at        TEXT,
    PRIMARY KEY (code, period, q_no, item)
);

-- ----------------------------------------------------------------- guidance
-- §5: guidance(code, date, fy, item, value, revision_direction)
CREATE TABLE IF NOT EXISTS guidance (
    code                TEXT,
    date                TEXT,
    fy                  TEXT,
    item                TEXT,
    value               REAL,
    revision_direction  TEXT,                  -- up / down / flat / initial
    -- 【§5拡張】
    filing_id           INTEGER,
    prev_value          REAL,
    PRIMARY KEY (code, date, fy, item),
    FOREIGN KEY (filing_id) REFERENCES filings (id) ON DELETE CASCADE
);

-- 【統合引継ぎ書 v1.1 §6-1 / docs/adapter_design.md A】
-- 一時収入の分離層。単独値ビルダーが正しくても、一過性の売上は傾きを汚染する。
-- 3905 の FY2027Q1 は粗利率 90.90%(前年 27.58%)で S1 が +63.3pt の買いシグナルと
-- して誤発火するが、一時収入を除くと粗利は前年割れ。この層が無いと S1/S5 が
-- 実銘柄で誤作動する。
--
-- source_note を NOT NULL + CHECK(20文字以上) にしてあるのは、「金額だけ書いて
-- 済ませる」ことを構造的に不可能にするため。アプリ層の検証ではなくスキーマ制約に
-- 置くのは、検証を通さない経路(手作業のSQL・別スクリプト)からの登録を防ぐため。
-- 一過性収益の特定は機械抽出しない。注記テキスト由来で判断を要するので半手動。
CREATE TABLE IF NOT EXISTS pl_adjustments (
    code             TEXT NOT NULL,
    period           TEXT NOT NULL,      -- FY2027 (本体の period 語彙)
    q_no             INTEGER NOT NULL,
    item_key         TEXT NOT NULL,      -- one_time_revenue / one_time_cost / one_time_gain
    amount           REAL NOT NULL,      -- 控除する額(正値)。単位は本体と同じ「円」
    source_note      TEXT NOT NULL,      -- 短信/有報の注記からの引用(原文ママ)
    source_filing_id INTEGER NOT NULL,   -- 引用元の書類
    source_locator   TEXT NOT NULL,      -- 注記の所在(例:「(セグメント情報等) 3.」)
    confirmed_by     TEXT NOT NULL,      -- 記録した人
    confirmed_at     TEXT NOT NULL,
    note             TEXT,               -- 判断メモ(任意)
    PRIMARY KEY (code, period, q_no, item_key),
    FOREIGN KEY (source_filing_id) REFERENCES filings (id) ON DELETE CASCADE,
    CHECK (length(trim(source_note)) >= 20),
    CHECK (length(trim(source_locator)) > 0),
    CHECK (length(trim(confirmed_by)) > 0),
    CHECK (amount > 0)
);
CREATE INDEX IF NOT EXISTS ix_adj_code ON pl_adjustments (code, period, q_no);

-- ------------------------------------------------------------------ signals
-- §5: signals(code, eval_date, signal_id, value, fired, evidence_text)
CREATE TABLE IF NOT EXISTS signals (
    code            TEXT,
    eval_date       TEXT,
    signal_id       TEXT,                      -- S1..S8
    value           REAL,
    fired           INTEGER,
    evidence_text   TEXT,
    PRIMARY KEY (code, eval_date, signal_id)
);

-- ------------------------------------------------------------------- scores
-- §5: scores(code, week, total, rank)
CREATE TABLE IF NOT EXISTS scores (
    code            TEXT,
    week            TEXT,                      -- ISO week (2026-W35)
    total           REAL,
    rank            INTEGER,
    PRIMARY KEY (code, week)
);

-- ------------------------------------------------------------------- prices
-- §5: prices(code, date, close, volume, adv20) — J-Quants
CREATE TABLE IF NOT EXISTS prices (
    code            TEXT,
    date            TEXT,
    close           REAL,
    volume          REAL,
    adv20           REAL,                      -- 20日平均売買代金 JPY mn
    -- 【§5拡張】
    turnover_value  REAL,                      -- 当日売買代金 JPY
    -- 【§5拡張・2026-08-31】始値と調整後株価。J-Quants は5年ローリングで、
    -- 窓から落ちた日付は二度と取得できない。取れるうちに全部取る。
    --   open   出口ルール「翌日寄り指値」(v1.1 §5-5)の再現に必須
    --   adj_*  株式分割の調整。未調整だと分割が偽の -50% リターンになる
    --          (86970 の 2022-01-04 は C=2518.5 / AdjC=1259.3)
    --   mktcap 日次時価総額。ユニバース条件の時点再現に使える
    open            REAL,
    high            REAL,
    low             REAL,
    adj_factor      REAL,
    adj_close       REAL,
    adj_volume      REAL,
    mktcap          REAL,                      -- JPY mn
    PRIMARY KEY (code, date)
);

-- 【§5拡張・統合引継ぎ書 v1.1 §4-2「超過リターン評価」】
-- 市場指数(TOPIX)の日次終値。バックテストの層別分析(相場環境別)と、
-- 絶対リターンではなく超過リターンで評価するために要る。
-- 「持っていただけ」の効果とイベント効果を分離できないと、エッジの有無が判定できない。
CREATE TABLE IF NOT EXISTS market_index (
    date   TEXT PRIMARY KEY,               -- YYYY-MM-DD
    close  REAL NOT NULL,
    name   TEXT DEFAULT 'TOPIX'
);

-- =====================================================================
-- 以下は §5 に無いが仕様書本文が要求する運用テーブル。追加である旨を明示する。
-- =====================================================================

-- 【§5拡張・仕様書 §2-1 の「欠損記録」要件】
-- 取得を試みた単位(ソース×対象日)を必ず1行残す。status:
--   ok       取得成功(件数 > 0)
--   empty    正常に取得できたが対象0件(休日・非営業日など)
--   partial  一部のダウンロードに失敗(n_failed > 0)
--   failed   一覧の取得自体に失敗 → 再実行対象
CREATE TABLE IF NOT EXISTS fetch_runs (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    source          TEXT NOT NULL,             -- tdnet / edinet / jquants
    target_date     TEXT NOT NULL,             -- YYYY-MM-DD
    status          TEXT NOT NULL,
    n_listed        INTEGER DEFAULT 0,         -- 一覧に載っていた件数
    n_target        INTEGER DEFAULT 0,         -- 保存対象と判定した件数
    n_saved         INTEGER DEFAULT 0,
    n_failed        INTEGER DEFAULT 0,
    started_at      TEXT,
    finished_at     TEXT,
    attempt         INTEGER DEFAULT 1,
    error           TEXT,
    note            TEXT
);
CREATE INDEX IF NOT EXISTS ix_runs_src_date ON fetch_runs (source, target_date);

-- 【§5拡張・仕様書 §3-1 の「マッピング不能タグは unknown として記録し頻度上位を定期レビュー」】
CREATE TABLE IF NOT EXISTS unknown_tags (
    source          TEXT NOT NULL,             -- tdnet_summary / tdnet_attachment / edinet
    tag             TEXT NOT NULL,             -- XBRL 要素名(名前空間接頭辞を除く)
    n               INTEGER DEFAULT 0,         -- 出現回数
    n_filings       INTEGER DEFAULT 0,         -- 出現した filing 数
    sample_value    TEXT,
    first_seen      TEXT,
    last_seen       TEXT,
    PRIMARY KEY (source, tag)
);
CREATE INDEX IF NOT EXISTS ix_unknown_n ON unknown_tags (n DESC);


-- 【投影層・統合引継ぎ書 v1.1 §3-1】別枠 earnings_screener への価格の写像。
--
-- 別枠 module_d/backtest.py は `SELECT ticker, date, close FROM daily_prices`
-- を素の SQL で叩く。**別枠のコードは1行も変えない**という投影層の方針どおり、
-- 本体の prices をその名前・その列名へ射影するビューを本体側に置く。
--
-- ★ このビューの close は「調整後終値(adj_close)」である。未調整の close では
--   分割日に偽の暴落が出る —— 実測: 3110 日東紡績の 5:1 分割 (2026-06-29) は
--   未調整だと 19,630 -> 4,580 の -76.7%、調整後なら 3,926 -> 4,580 の +16.7%。
--   リターン計算は必ずこのビュー経由で行うこと。
--
-- ★ adj_close が NULL の日は行ごと返さない。未調整の close へ黙って
--   フォールバックする経路は**作らない**。「調整後が無い」と「調整後がこの値」を
--   取り違えるのが、この層で最も高くつく故障だから。
--   (実データでは 2021-08-31 の 1,152 行だけが該当。J-Quants Light の
--    5年ローリング窓から落ちた日で、二度と取得できない。)
--
-- ★ 単位が混在する。J-Quants がそう返すので変換せずそのまま置く代わりに、
--   ここに明記する: mktcap は【百万円】、turnover_value は【円】。
--   実測で確認済み (2026-06-30): turnover_value ≒ close × volume (比 1.00)、
--   mktcap の最大 49,076,721 = 49兆円規模。したがって仕様書のユニバース条件は
--     時価総額 50〜1,000億円      -> mktcap BETWEEN 5000 AND 100000
--     20日平均売買代金 3,000万円  -> AVG(turnover_value) >= 30000000
--   となる。片方の桁を取り違えると、フィルタが 100万倍ずれても
--   エラーにならず「該当0件」や「全件通過」として静かに出る。
CREATE VIEW IF NOT EXISTS daily_prices AS
    SELECT code           AS ticker,
           date           AS date,
           adj_close      AS close,            -- 調整後終値
           adj_volume     AS volume,           -- 調整後出来高
           mktcap         AS mktcap,           -- 時価総額【百万円】
           turnover_value AS turnover_value    -- 売買代金【円】
      FROM prices
     WHERE adj_close IS NOT NULL;


-- 【ペーパートレード・docs/paper_trading_design.md §2】判断の凍結。
--
-- 「その時点で本当にそう判断していたのか」を後から証明するための証拠。
-- 人間の記憶も再計算した値も証拠にならない —— DBは更新され、訂正短信は
-- 過去を書き換えるから。判定した瞬間に書き出し、**二度と更新しない**。
--
-- ★ 追記専用。UPDATE も DELETE もしない。訂正が出ても過去の行は残す
--   （「当時はそう見えていた」が事実だから）。
-- ★ inputs_json には**値そのもの**を入れる。参照だけだと、後で値が
--   訂正されたときに当時の判断を再現できない。
CREATE TABLE IF NOT EXISTS forecast_snapshots (
    snapshot_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    as_of           TEXT NOT NULL,          -- 判定日。この日までの情報だけで判断した
    code            TEXT NOT NULL,
    event_date      TEXT,                   -- 対象の決算発表予定日（代理日）
    entry_date      TEXT,                   -- T-15 に相当する営業日
    evidence_score  REAL,
    scores_json     TEXT,                   -- S1〜S8 のスコア・available・evidence
    inputs_json     TEXT,                   -- スコアの入力になった実数値
    filing_ids      TEXT,                   -- 根拠にした本体 filings.id の列
    topix_close     REAL,
    topix_ma200     REAL,
    market_allowed  INTEGER,                -- 市場フィルターの判定
    decision        TEXT NOT NULL,          -- entry / skip_full / skip_filter / skip_score
    decision_note   TEXT,
    frozen_at       TEXT NOT NULL,
    -- 訂正の記録（追記専用を守る仕組み。migration 006/007）。
    -- 旧版は削除せず invalidated=1 を立て、訂正版を revision+1 で追記する。
    invalidated        INTEGER DEFAULT 0,
    invalidated_reason TEXT,
    invalidated_at     TEXT,
    revision           INTEGER NOT NULL DEFAULT 1,
    UNIQUE(as_of, code, event_date, revision)
);
CREATE INDEX IF NOT EXISTS ix_snap_asof ON forecast_snapshots (as_of, decision);

-- 【ペーパートレード §3】仮想トレード台帳。エントリーからエグジットまでを1行で。
--
-- 想定約定価格は**翌営業日の始値**。終値を使うと「引けを見てから建てた」
-- ことになる。バックテスト(終値ベース)との差分はスリッページの実測値に
-- なるので、終値ベースの損益も併記して両方残す。
CREATE TABLE IF NOT EXISTS paper_trades (
    trade_id            INTEGER PRIMARY KEY AUTOINCREMENT,
    snapshot_id         INTEGER REFERENCES forecast_snapshots(snapshot_id),
    code                TEXT NOT NULL,
    event_date          TEXT,
    entry_date          TEXT NOT NULL,      -- 判定日
    entry_fill_date     TEXT,               -- 実際に建てたとみなす日（翌営業日）
    entry_price_assumed REAL,               -- その日の始値
    entry_price_close   REAL,               -- 比較用（判定日の終値）
    position_size       REAL,               -- 建玉時の想定NAVの10%
    exit_rule           TEXT,               -- 4分岐のどれで出たか
    exit_date           TEXT,
    exit_fill_date      TEXT,
    exit_price_assumed  REAL,
    exit_price_close    REAL,
    ret_gross           REAL,
    ret_net             REAL,               -- 往復コスト0.4%控除後
    ret_net_close_basis REAL,               -- 終値ベース（バックテストとの比較用）
    topix_entry         REAL,
    topix_exit          REAL,
    status              TEXT NOT NULL,      -- open / closed
    opened_at           TEXT,
    closed_at           TEXT,
    -- 発表日は「推定」でエントリーし「実績」でエグジットする。
    -- エントリー起点 = 推定発表日の T-15（推定でしか決められない）
    -- エグジット起点 = 実際に観測された発表日の T+2（実績が分かってから）
    event_date_estimated TEXT,
    event_date_actual    TEXT,              -- EDINET/TDnet で捕捉した実際の発表日
    date_error_bdays     INTEGER,           -- 実績 - 推定（営業日）
    -- 訂正の記録（追記専用を守る仕組み。migration 006/007）。
    -- 旧版は削除せず invalidated=1 を立て、訂正版を revision+1 で追記する。
    invalidated        INTEGER DEFAULT 0,
    invalidated_reason TEXT,
    invalidated_at     TEXT,
    revision           INTEGER NOT NULL DEFAULT 1,
    UNIQUE(code, entry_date, event_date, revision)
);
CREATE INDEX IF NOT EXISTS ix_paper_status ON paper_trades (status, entry_date);

-- 【ペーパートレード §4】市場フィルターの作動記録。
--
-- 止めた候補について「建てていたらどうなったか」も並行して記録する
-- (n_blocked / blocked_codes)。フィルターが役に立ったのか単に good trade を
-- 削っただけなのかは、これでしか分からない。
-- ★ この記録を理由に運用途中でフィルターを変更・停止しない（§4-1）。
CREATE TABLE IF NOT EXISTS market_filter_log (
    date            TEXT PRIMARY KEY,
    topix_close     REAL,
    topix_ma200     REAL,
    allowed         INTEGER NOT NULL,
    n_candidates    INTEGER DEFAULT 0,
    n_blocked       INTEGER DEFAULT 0,
    blocked_codes   TEXT,
    logged_at       TEXT
);


-- 【ペーパートレード】発表日の推定と実績の突合（追記専用）。
--
-- なぜ forecast_snapshots を更新しないのか:
--   凍結レコードは「その時点でそう判断した」ことの証拠であり、
--   **後から書き換えたらその瞬間に証拠でなくなる**。実績が判明したことは
--   新しい事実なので、新しい行として追記する。報告時は join して
--   「フラグが付いた凍結レコード」として見せる。
--
-- 3営業日以上のズレに flag を立てる。あとで誤差分布の実測に使う
-- （バックログ 1-2b の代理日誤差を、前向きデータで測り直すことになる）。
CREATE TABLE IF NOT EXISTS calendar_reconciliation (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    snapshot_id     INTEGER REFERENCES forecast_snapshots(snapshot_id),
    code            TEXT NOT NULL,
    event_date_estimated TEXT NOT NULL,
    event_date_actual    TEXT NOT NULL,
    error_bdays     INTEGER NOT NULL,
    flagged         INTEGER NOT NULL,       -- |error| >= 3 営業日
    confidence_at_entry TEXT,               -- 推定時の HIGH/MEDIUM/LOW
    source_filing_id INTEGER,               -- 実績を捕捉した本体 filings.id
    reconciled_at   TEXT NOT NULL,
    UNIQUE(snapshot_id, event_date_actual)
);
CREATE INDEX IF NOT EXISTS ix_recon_flag ON calendar_reconciliation (flagged, code);


-- 【v3 シャドウポートフォリオ・backtest_acceptance_criteria.md v3 事前登録】
--
-- v2(本番ペーパー)は前向き検証の途中なので**一切変更しない**。走っている
-- 検証の途中で判定関数を変えたら、その瞬間に検証の意味が消える。
-- そこで同一データ・同一週次・同一スコアに対して、採否だけを別ルールで
-- 計算した結果をここに記録し、成績を前向きに比較する。
--
-- v3 の追加ルール（v3-1 / v3-2）:
--   A1 available なシグナルが3本未満の銘柄は採用しない
--   B1 同一イベント日の新規エントリー最大5件 / 同一セクター同時保有最大3件
CREATE TABLE IF NOT EXISTS shadow_snapshots (
    snapshot_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    variant         TEXT NOT NULL DEFAULT 'v3',
    as_of           TEXT NOT NULL,
    code            TEXT NOT NULL,
    event_date      TEXT,
    entry_date      TEXT,
    evidence_score  REAL,
    n_available     INTEGER,                -- A1 の判定に使った available 本数
    sector          TEXT,
    decision        TEXT NOT NULL,          -- entry / skip_score / skip_filter /
                                            -- skip_full / skip_evidence / skip_daycap /
                                            -- skip_sectorcap
    decision_note   TEXT,
    frozen_at       TEXT NOT NULL,
    -- 訂正の記録（追記専用を守る仕組み。migration 006/007）。
    -- 旧版は削除せず invalidated=1 を立て、訂正版を revision+1 で追記する。
    invalidated        INTEGER DEFAULT 0,
    invalidated_reason TEXT,
    invalidated_at     TEXT,
    revision           INTEGER NOT NULL DEFAULT 1,
    UNIQUE(variant, as_of, code, event_date, revision)
);
CREATE INDEX IF NOT EXISTS ix_shadow_snap ON shadow_snapshots (variant, as_of, decision);

CREATE TABLE IF NOT EXISTS shadow_trades (
    trade_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    variant         TEXT NOT NULL DEFAULT 'v3',
    snapshot_id     INTEGER REFERENCES shadow_snapshots(snapshot_id),
    code            TEXT NOT NULL,
    sector          TEXT,
    event_date      TEXT,
    entry_date      TEXT NOT NULL,
    entry_fill_date TEXT,
    entry_price_assumed REAL,
    position_size   REAL,
    exit_date       TEXT,
    exit_price_assumed  REAL,
    ret_gross       REAL,
    ret_net         REAL,
    status          TEXT NOT NULL,
    opened_at       TEXT,
    closed_at       TEXT,
    event_date_estimated TEXT,
    event_date_actual    TEXT,
    date_error_bdays     INTEGER,
    -- 訂正の記録（追記専用を守る仕組み。migration 006/007）。
    -- 旧版は削除せず invalidated=1 を立て、訂正版を revision+1 で追記する。
    invalidated        INTEGER DEFAULT 0,
    invalidated_reason TEXT,
    invalidated_at     TEXT,
    revision           INTEGER NOT NULL DEFAULT 1,
    UNIQUE(variant, code, entry_date, event_date, revision)
);
CREATE INDEX IF NOT EXISTS ix_shadow_trades ON shadow_trades (variant, status, entry_date);


-- 【決算説明会資料・2026-09-02】収集した資料の対象期と世代。
--
-- なぜ filings と別に持つか:
--   filings は「開示の索引」で、対象期(period)を持たない。説明会資料は
--   2027年秋に前年ペアを組むとき **タイトルからしか対象期が分からない**。
--   その時に全件を再パースするのは無駄なので、収集時点で確定させておく。
--
-- 世代管理(PIT):
--   （訂正）資料は元の資料と同じ (code, period_label) に属する別世代として
--   記録する。差分を取るときは **as_of 時点で見えている最新世代**を使う。
--   訂正前の資料も残す —— 「当時はそう書かれていた」が事実だから。
--
-- period_label が NULL になる場合:
--   「（2024年3月期第1四半期～2026年3月期）の一部訂正」のように複数期を
--   まとめた訂正は、どの期のものか一意に決まらない。**推測せず NULL**。
--   ペアの構築対象から外れる（0点ではなく欠損）。
CREATE TABLE IF NOT EXISTS presentation_materials (
    filing_id       INTEGER PRIMARY KEY REFERENCES filings(id),
    code            TEXT NOT NULL,
    doc_id          TEXT,
    disclosed_date  TEXT NOT NULL,
    title           TEXT,
    period_label    TEXT,               -- 'FY2027-Q1' 形式。決められなければ NULL
    fiscal_year     INTEGER,            -- 決算期の年（2027年3月期 -> 2027）
    fy_end_month    INTEGER,            -- 決算期末月（3月期 -> 3）
    quarter_type    TEXT,               -- '1Q'/'2Q'/'3Q'/'FY'
    is_correction   INTEGER DEFAULT 0,
    ambiguous       INTEGER DEFAULT 0,  -- 複数期にまたがる訂正など
    -- 資料の種類。同じ期に本編・補足・書き起こし・サマリ・質疑応答が
    -- 並行して出るので、世代を種類ごとに分ける。分けないと「最新世代」が
    -- 訂正版ではなく単に後から出た別文書を指してしまう。
    -- 告知(動画公開のお知らせ等)は内容を持たないので差分母集団から外す。
    -- 分類できないものは 'unknown'。推測で埋めない。
    doc_kind        TEXT,               -- 本編/補足/書き起こし/サマリ/質疑応答/告知/unknown
    generation      INTEGER DEFAULT 1,  -- (code, period_label, doc_kind) 内の開示順
    text_path       TEXT,
    text_chars      INTEGER,
    extract_status  TEXT,               -- ok / short / failed / missing_pdf
    parsed_at       TEXT
);
CREATE INDEX IF NOT EXISTS ix_pres_period
    ON presentation_materials (code, period_label, doc_kind, generation);


-- 【S12 定性文言diff・2026-09-02】各ヒットの根拠。
--
-- **根拠文の無い点数は存在してはならない。** evidence_flag の規律と同じで、
-- 「なぜその点が付いたか」を原文の1文まで遡れない点は信用できない。
-- スコアだけを保存して根拠を捨てると、後から誤検出を潰せなくなる。
CREATE TABLE IF NOT EXISTS s12_evidence (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    filing_id       INTEGER NOT NULL REFERENCES filings(id),   -- 当期の書類
    prior_filing_id INTEGER REFERENCES filings(id),            -- 比較した前期
    code            TEXT NOT NULL,
    period_label    TEXT,
    section         TEXT,               -- 切り出した節の名前
    tier            TEXT NOT NULL,      -- A / B / C / D
    rule_key        TEXT NOT NULL,      -- 辞書のキー
    matched_text    TEXT NOT NULL,      -- マッチした文（根拠）
    score           REAL NOT NULL,      -- この1件が寄与した点数（D は 0）
    dedup_applied   TEXT,               -- 0点にした理由（あれば）
    created_at      TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_s12_filing ON s12_evidence (filing_id, tier);

-- S12 のスコア本体。available=0（ペアが無い）と 0点は別物なので分けて持つ。
CREATE TABLE IF NOT EXISTS s12_scores (
    filing_id       INTEGER PRIMARY KEY REFERENCES filings(id),
    prior_filing_id INTEGER,
    code            TEXT NOT NULL,
    period_label    TEXT,
    available       INTEGER NOT NULL,   -- 0 = ペアが無い/本文が取れない（欠損）
    unavailable_reason TEXT,
    score           REAL,               -- 合成後（clamp 済み）
    tier_a          REAL,
    tier_b          REAL,
    tier_c          REAL,
    segment_changed INTEGER DEFAULT 0,
    new_product_mention INTEGER DEFAULT 0,
    forecast_revision_mentioned INTEGER DEFAULT 0,
    forecast_revision_direction TEXT,   -- guidance から補完。本文では判定しない
    computed_at     TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_s12_code ON s12_scores (code, period_label);
