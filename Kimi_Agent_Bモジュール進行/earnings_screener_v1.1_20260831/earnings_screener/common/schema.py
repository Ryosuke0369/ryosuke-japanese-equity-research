"""
schema.py — 共通DBスキーマ定義（ステップ0：土台）
全モジュールが参照するSQLiteスキーマ。無効化フラグ機構をここに組み込む。
"""
import sqlite3
from pathlib import Path

DDL = """
-- ユニバース（J-Quants選定済み1,327社）
CREATE TABLE IF NOT EXISTS universe (
    ticker            TEXT PRIMARY KEY,   -- 例: '3905', '278A'
    company_name      TEXT,
    market            TEXT,               -- growth / standard / prime
    sector            TEXT,               -- セクター（層別分析用。実運用はJ-Quants由来）
    fiscal_year_end   INTEGER NOT NULL,   -- 決算期末月 (1-12)
    market_cap_ok     INTEGER DEFAULT 1,  -- 時価総額フィルタ通過
    liquidity_ok      INTEGER DEFAULT 1   -- 売買代金フィルタ通過
);

-- 開示履歴（TDnetアーカイブ由来。モック段階では模擬レコード）
CREATE TABLE IF NOT EXISTS filings (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    ticker        TEXT NOT NULL,
    filing_date   TEXT NOT NULL,          -- YYYY-MM-DD（発表日）
    period_end    TEXT NOT NULL,          -- 対象四半期の期末日
    quarter_type  TEXT NOT NULL,          -- '1Q'/'2Q'/'3Q'/'FY'
    fiscal_year   INTEGER NOT NULL,       -- 決算年度（例: 2027年3月期 → 2027）
    source        TEXT DEFAULT 'mock',    -- 'tdnet' / 'mock'
    pdf_path      TEXT,                   -- TDnetアーカイブ内パス（実運用時）
    xbrl_path     TEXT,
    generation    INTEGER DEFAULT 1,      -- 訂正世代（初回=1、訂正短信ごとに+1）
    UNIQUE(ticker, period_end, generation)
);
CREATE INDEX IF NOT EXISTS idx_filings_ticker ON filings(ticker, filing_date);

-- 四半期単独値（Module B以降で使用。累計→単独変換済みを格納）
CREATE TABLE IF NOT EXISTS quarterly_standalone (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    ticker             TEXT NOT NULL,
    fiscal_year        INTEGER NOT NULL,
    quarter_type       TEXT NOT NULL,
    period_end         TEXT NOT NULL,
    sales              REAL,              -- 単独売上高（百万円）
    operating_profit   REAL,              -- 単独営業利益
    gross_profit       REAL,              -- 単独粗利
    cogs_quantity      REAL,              -- 売上原価・数量要因（S6用）
    cogs_price         REAL,              -- 売上原価・価格要因（S6用）
    operating_cf       REAL,              -- 営業CF（単独換算・開示期のみ。S4用）
    is_valid           INTEGER DEFAULT 1, -- ★無効化フラグ：0=決算期変更/遡及修正/連結範囲変更
    invalid_reason     TEXT,              -- 'period_change'/'restatement'/'scope_change'
    generation         INTEGER DEFAULT 1, -- filings.generation と対応（訂正版は別世代で保持・上書きしない）
    UNIQUE(ticker, period_end, generation)
);

-- PL調整項目（§8-3：一時収入の除外など。機械予測は調整後をベースラインにする）
CREATE TABLE IF NOT EXISTS pl_adjustments (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    ticker     TEXT NOT NULL,
    period_end TEXT NOT NULL,
    item_key   TEXT NOT NULL,             -- 'one_time_revenue'/'one_time_gain'...
    amount     REAL NOT NULL,             -- 百万円（売上から控除する場合は正の値）
    note       TEXT,
    generation INTEGER DEFAULT 1,
    UNIQUE(ticker, period_end, item_key, generation)
);

-- 会社通期予想（S5進捗率異常検出・織り込み度比較用）
CREATE TABLE IF NOT EXISTS company_forecasts (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    ticker         TEXT NOT NULL,
    fiscal_year    INTEGER NOT NULL,
    forecast_sales REAL NOT NULL,         -- 通期予想売上高（百万円）
    forecast_op    REAL NOT NULL,         -- 通期予想営業利益（百万円）
    source_date    TEXT NOT NULL,         -- 予想発表日（修正履歴を残す）
    UNIQUE(ticker, fiscal_year, source_date)
);

-- BS科目スナップショット（S1-S3, S6-S8の元データ）
CREATE TABLE IF NOT EXISTS balance_sheet_items (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    ticker        TEXT NOT NULL,
    period_end    TEXT NOT NULL,
    item_key      TEXT NOT NULL,          -- 'accounts_receivable'/'contract_liabilities'/'construction_in_progress'/'machinery'/'inventory'/...
    value         REAL NOT NULL,          -- 百万円
    generation    INTEGER DEFAULT 1,
    UNIQUE(ticker, period_end, item_key, generation)
);

-- 証拠イベント（C経路：適時開示の証拠積み上げ。約束はevidence_flag=0で保存のみ）
CREATE TABLE IF NOT EXISTS evidence_events (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    ticker        TEXT NOT NULL,
    event_date    TEXT NOT NULL,
    event_type    TEXT NOT NULL,          -- 'order'/'contract'/'facility_operation'/'mou'/'plan'
    amount        REAL,                   -- 金額（百万円）。不明はNULL
    evidence_flag INTEGER NOT NULL,       -- ★1=証拠（履行義務あり実契約/計上済み）、0=約束（MOU/計画/目標→スコア根拠にしない）
    source_doc    TEXT,
    note          TEXT
);

-- 逆算DCF要求値（既存バリュエーション層からの入力IF）
CREATE TABLE IF NOT EXISTS reverse_dcf_requirements (
    ticker                    TEXT PRIMARY KEY,
    as_of_date                TEXT NOT NULL,
    required_steady_op_profit REAL NOT NULL,  -- 定常営業利益要求値（百万円）
    target_year               INTEGER,        -- 要求到達年度
    source_run_id             TEXT
);

-- 決算カレンダー（Module A の主出力）
CREATE TABLE IF NOT EXISTS earnings_calendar (
    ticker              TEXT PRIMARY KEY,
    next_earnings_date  TEXT NOT NULL,
    quarter_type        TEXT NOT NULL,        -- 次回が '1Q'/'2Q'/'3Q'/'FY'
    fiscal_year         INTEGER NOT NULL,
    fiscal_year_end     INTEGER NOT NULL,
    confidence_level    TEXT NOT NULL,        -- 'HIGH'/'MEDIUM'/'LOW'
    estimated_from      TEXT NOT NULL,        -- 推定根拠となった前年同日の発表日
    updated_at          TEXT NOT NULL
);

-- 機械予測スナップショット（Module B → Module C の受け渡し。週次で記録し決算後に実績比較）
CREATE TABLE IF NOT EXISTS forecast_snapshots (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    ticker        TEXT NOT NULL,
    as_of_date    TEXT NOT NULL,          -- 予測を出した週次バッチ日
    fiscal_year   INTEGER NOT NULL,
    quarter_type  TEXT NOT NULL,          -- 予測対象の四半期（次回決算）
    target_period_end TEXT,               -- 予測対象の期末日
    pred_sales    REAL NOT NULL,
    pred_op       REAL NOT NULL,
    steady_op_est REAL,
    evidence_score REAL,
    total_score   REAL,
    UNIQUE(ticker, fiscal_year, quarter_type, as_of_date)
);

-- 決算実績（Module C：決算短信XBRLパース結果の格納先。モック段階では模擬着弾）
CREATE TABLE IF NOT EXISTS earnings_actuals (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    ticker        TEXT NOT NULL,
    fiscal_year   INTEGER NOT NULL,
    quarter_type  TEXT NOT NULL,
    period_end    TEXT NOT NULL,
    announce_date TEXT NOT NULL,
    actual_sales  REAL NOT NULL,          -- 当該四半期の単独実績（XBRL累計→単独変換後）
    actual_op     REAL NOT NULL,
    guidance_sales REAL,                  -- 同時開示の（修正後）通期予想。無修正なら従来予想
    guidance_op   REAL,
    source        TEXT DEFAULT 'mock',    -- 'xbrl' / 'mock'
    UNIQUE(ticker, period_end)
);

-- 決算発表時の市場データ（PTS出来高判定用）
CREATE TABLE IF NOT EXISTS pts_observations (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    ticker          TEXT NOT NULL,
    event_date      TEXT NOT NULL,        -- 決算発表日
    pts_volume      REAL,                 -- 発表後PTSの出来高
    daily_avg_volume REAL,                -- 直近20日平均出来高
    UNIQUE(ticker, event_date)
);

-- 日次株価（Module Dバックテスト用。実運用はJ-Quants由来）
CREATE TABLE IF NOT EXISTS daily_prices (
    ticker TEXT NOT NULL,
    date   TEXT NOT NULL,                 -- YYYY-MM-DD
    close  REAL NOT NULL,
    volume REAL,
    PRIMARY KEY (ticker, date)
);
CREATE INDEX IF NOT EXISTS idx_prices_date ON daily_prices(date);

-- 市場指数（日経平均の代替。市場環境別の層別分析用）
CREATE TABLE IF NOT EXISTS market_index (
    date  TEXT PRIMARY KEY,
    close REAL NOT NULL
);
"""


def init_db(db_path: str | Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path))
    conn.executescript(DDL)
    conn.commit()
    return conn


if __name__ == "__main__":
    p = Path(__file__).resolve().parents[1] / "data" / "screener.db"
    init_db(p)
    print(f"schema initialized: {p}")
