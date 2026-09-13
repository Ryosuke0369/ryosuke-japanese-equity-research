"""screener/report/paper_weekly.py — ペーパートレードの週次バッチ。

設計の正本: docs/paper_trading_design.md（2026-09-02 承認）。

やること（設計書 §1 の7段）
  1. 候補抽出       その週に T-15 を迎える銘柄を洗い出す
  2. スコアリング   score_ticker(as_of=判定日) で PIT 評価
  3. 市場フィルター TOPIX 200日移動平均（前日まで）
  4. 建玉判定       10枠制約・score降順
  5. 判断の凍結     forecast_snapshots へ追記（二度と更新しない）
  6. エグジット判定 保有中ポジションの決済
  7. レポート出力   候補リスト + 週次サマリ

実弾の発注機能は**一切作らない**（設計書 §6）。このモジュールは
DBに記録を書くだけで、外部に注文を出す経路を持たない。

    python -m screener.report.paper_weekly --dry-run     # 書き込まない
    python -m screener.report.paper_weekly
    python -m screener.report.paper_weekly --as-of 2026-09-07
"""
from __future__ import annotations

import argparse
import json
import os
import sqlite3
import sys
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import backtest_eval as V1
from screener.report import backtest_v2 as V2

ENTRY_LEAD = V1.MAIN_ENTRY          # T-15
EXIT_LAG = V1.MAIN_EXIT             # T+2
SCORE_THRESHOLD = 0.10
MAX_POSITIONS = V2.MAX_POSITIONS
POS_FRACTION = V2.POS_FRACTION

# ---- v3 シャドウ（backtest_acceptance_criteria.md v3 事前登録）-------------
# 実行後に動かさない。動かしたくなったら v4 として新規に事前登録する。
V3_MIN_AVAILABLE = 3        # A1: available なシグナルが3本未満は採用しない
V3_MAX_PER_DAY = 5          # B1: 同一イベント日の新規エントリー上限
V3_MAX_PER_SECTOR = 3       # B1: 同一セクターの同時保有上限

# ---- シャドウB（エグジット改善+スコア連動サイジング）--------------------
# 事前登録: backtest_acceptance_criteria.md「シャドウB/C 事前登録」
# 初値であり、データを見て変えない。
VB_SIZE_TIERS = ((0.30, 0.15), (0.15, 0.10), (0.10, 0.05))  # (スコア下限, 比率)
VB_MAX_EXPOSURE = 1.00      # 合計エクスポージャの上限
VB_TRAIL_DAYS = 3           # TP50_TRAIL の残り50%は T+3

# ---- シャドウC（高閾値版）-----------------------------------------------
VC_SCORE_THRESHOLD = 0.20   # v2 の 0.10 に対して倍。他は v2 と完全に同一


def vb_position_size(score):
    """スコア連動のポジション比率。閾値未満は 0（建てない）。"""
    for lo, frac in VB_SIZE_TIERS:
        if score >= lo:
            return frac
    return 0.0


def business_days(con):
    return [r[0] for r in con.execute(
        "SELECT DISTINCT date FROM daily_prices ORDER BY date")]


def week_span(as_of):
    """as_of を含む週（月曜〜金曜）を返す。"""
    monday = as_of - timedelta(days=as_of.weekday())
    return monday, monday + timedelta(days=4)


def upcoming_events(pcon, lo, hi):
    """entry_date（= event の ENTRY_LEAD 営業日前）が [lo, hi] に入るイベント。

    投影DBの filings がイベント台帳。event_date は代理発表日。
    """
    bd = business_days(pcon)
    pos = {d: i for i, d in enumerate(bd)}
    out = []
    for r in pcon.execute(
            "SELECT ticker, filing_date, period_end, quarter_type "
            "FROM filings WHERE generation=1 ORDER BY filing_date"):
        d = r[1]
        i = pos.get(d)
        if i is None or i - ENTRY_LEAD < 0:
            continue
        entry = bd[i - ENTRY_LEAD]
        if lo <= entry <= hi:
            out.append({"code": r[0], "event_date": d, "entry_date": entry,
                        "period_end": r[2], "quarter_type": r[3]})

    # 1銘柄1候補に畳む。有報のように1つの書類が複数の期を語ると、
    # 同じ (銘柄, entry_date) が複数行になり、**同じ銘柄が10枠を
    # 2つ以上消費する**（2026-09-02 の初回実行で 3496 / 7022 等で発生）。
    # 残すのは最新の期（period_end 最大）。
    best = {}
    for e in out:
        k = (e["code"], e["entry_date"])
        if k not in best or e["period_end"] > best[k]["period_end"]:
            best[k] = e
    return sorted(best.values(), key=lambda e: (e["entry_date"], e["code"]))


def upcoming_events_from_calendar(pcon, lo, hi):
    """earnings_calendar（前年同期推定）から、その週に T-15 を迎える銘柄を出す。

    価格データは過去しか無いので、営業日の前後計算は別枠の jp_calendar
    （祝日を算出で持つ）に任せる。投影DBの daily_prices を使うと
    未来の営業日が引けない。

    J-Quants の /equities/earnings-calendar は 2026-09-02 実測で**1件しか
    返さず**（しかも過去日）、ユニバースのカバレッジは 0.1%。確定予定日の
    ソースとしては使えないので、推定を主エンジンとして固定した。
    """
    root = V1._external_root()
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from common.jp_calendar import add_business_days

    out = []
    for r in pcon.execute(
            "SELECT ticker, next_earnings_date, quarter_type, fiscal_year, "
            " confidence_level, estimated_from FROM earnings_calendar"):
        ev = date.fromisoformat(r[1])
        entry = add_business_days(ev, -ENTRY_LEAD).isoformat()
        if lo <= entry <= hi:
            out.append({"code": r[0], "event_date": r[1], "entry_date": entry,
                        "period_end": "%s-%s" % (r[3], r[2]),
                        "quarter_type": r[2], "confidence": r[4],
                        "estimated_from": r[5], "date_source": "estimated"})
    return sorted(out, key=lambda e: (e["entry_date"], e["code"]))


def next_business_day(bd, d):
    for x in bd:
        if x > d:
            return x
    return None


def price_on(pcon, code, d, col="close"):
    r = pcon.execute("SELECT %s FROM daily_prices WHERE ticker=? AND date=?"
                     % col, (code, d)).fetchone()
    return r[0] if r else None


def open_price(mcon, code, d):
    r = mcon.execute("SELECT open FROM prices WHERE code=? AND date=?",
                     (code, d)).fetchone()
    return r[0] if r else None


def score_events(pcon, events):
    """別枠のスコアラーで PIT 評価する。別枠のコードは変更しない。"""
    root = V1._external_root()
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from module_b.run_scorers import score_ticker, SCORERS_ALL

    out = []
    for ev in events:
        as_of = date.fromisoformat(ev["entry_date"])
        try:
            res = score_ticker(pcon, ev["code"], SCORERS_ALL, as_of=as_of)
        except Exception as e:                           # pragma: no cover
            res = {"evidence_score": None, "error": "%s: %s" % (type(e).__name__, e)}
        ev = dict(ev)
        ev["scores"] = {k: v for k, v in res.items() if k != "evidence_score"}
        ev["evidence_score"] = res.get("evidence_score")
        out.append(ev)

    # 設定ミスの検出。スコアラーの例外は available=False に潰れるので、
    # 黙っていると「データが無い」と区別できない。全件が ERROR: なら
    # データではなく接続や環境の問題である。
    n_err = sum(1 for e in out
                for v in e["scores"].values()
                if isinstance(v, dict) and str(v.get("evidence", "")).startswith("ERROR:"))
    n_cells = sum(len(e["scores"]) for e in out) or 1
    if n_err and n_err / n_cells > 0.5:
        sample = next(str(v.get("evidence")) for e in out
                      for v in e["scores"].values()
                      if isinstance(v, dict)
                      and str(v.get("evidence", "")).startswith("ERROR:"))
        raise RuntimeError(
            "スコアラーの %d/%d が例外。データ不足ではなく設定の問題を疑う: %s"
            % (n_err, n_cells, sample))
    return out


def reconcile_dates(mcon, dry_run=False):
    """建玉中のトレードについて、推定発表日と実績を突き合わせる。

    **エントリーは推定発表日の T-15、エグジットは実績発表日の T+2** で
    アンカーが違う。推定でしか建てられないが、決済は実際に発表された日を
    起点にするのが正しい（企業がいつ発表したかは事後に確定する）。

    forecast_snapshots は更新しない。凍結レコードを後から書き換えたら
    その瞬間に証拠でなくなる。実績が判明したことは新しい事実なので
    calendar_reconciliation に追記する。3営業日以上のズレに flag を立てる。
    """
    root = V1._external_root()
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from common.jp_calendar import business_days_between

    rows = mcon.execute(
        "SELECT trade_id, snapshot_id, code, entry_date, event_date "
        "FROM paper_trades WHERE status='open' AND event_date_actual IS NULL").fetchall()
    n_rec = n_flag = 0
    for t in rows:
        est = t["event_date"]
        act = mcon.execute(
            "SELECT id, date FROM filings WHERE code=? AND date>=? "
            "AND (source='tdnet' OR subtype IN ('120','140','160')) "
            "ORDER BY date LIMIT 1", (t["code"], t["entry_date"])).fetchone()
        if not act:
            continue                                  # まだ発表されていない
        err = business_days_between(date.fromisoformat(est),
                                    date.fromisoformat(act["date"]))
        flagged = 1 if abs(err) >= 3 else 0
        n_rec += 1
        n_flag += flagged
        if dry_run:
            continue
        conf = mcon.execute(
            "SELECT inputs_json FROM forecast_snapshots WHERE snapshot_id=?",
            (t["snapshot_id"],)).fetchone()
        conf_lv = None
        if conf and conf["inputs_json"]:
            try:
                conf_lv = json.loads(conf["inputs_json"]).get("date_confidence")
            except ValueError:
                pass
        mcon.execute(
            "INSERT OR IGNORE INTO calendar_reconciliation (snapshot_id, code, "
            " event_date_estimated, event_date_actual, error_bdays, flagged, "
            " confidence_at_entry, source_filing_id, reconciled_at) "
            "VALUES (?,?,?,?,?,?,?,?,?)",
            (t["snapshot_id"], t["code"], est, act["date"], err, flagged,
             conf_lv, act["id"], C.utcnow()))
        mcon.execute(
            "UPDATE paper_trades SET event_date_estimated=?, event_date_actual=?, "
            " date_error_bdays=? WHERE trade_id=?",
            (est, act["date"], err, t["trade_id"]))
    if not dry_run:
        mcon.commit()
    return n_rec, n_flag


def n_open_positions(mcon, on_date):
    # 無効化された訂正前の行は数えない（枠を二重に消費してしまう）
    r = mcon.execute(
        "SELECT COUNT(*) FROM paper_trades WHERE status='open' "
        "AND COALESCE(invalidated,0)=0 AND entry_date<=?", (on_date,)).fetchone()
    return r[0]


def decide(pcon, mcon, scored, dry_run=False):
    """市場フィルター → 10枠制約 → 建玉可否。判断は必ず凍結する。"""
    idx = pcon.execute("SELECT date, close FROM market_index ORDER BY date").fetchall()
    idx_d = [r[0] for r in idx]
    idx_c = [r[1] for r in idx]
    bd = business_days(pcon)

    by_day = {}
    for ev in scored:
        by_day.setdefault(ev["entry_date"], []).append(ev)

    frozen, entries = [], []
    for day in sorted(by_day):
        cands = sorted(by_day[day],
                       key=lambda e: (-(e["evidence_score"] or -1), e["code"]))
        allowed, ma = V2.topix_ma_ok(idx_d, idx_c, day)
        topix_close = None
        for i in range(len(idx_d) - 1, -1, -1):
            if idx_d[i] < day:
                topix_close = idx_c[i]
                break
        # 枠は「この実行の中で建てた分」も数える。DBへの書き込みはループ後
        # なので、DBだけを見ると 9/15 に4件建てても 9/17 は空きが10あると
        # 誤認する。実際 2026-09-02 の初回凍結で上限10に対し13件建った。
        free = MAX_POSITIONS - n_open_positions(mcon, day) - len(entries)
        used = 0
        for ev in cands:
            s = ev["evidence_score"]
            if s is None or s < SCORE_THRESHOLD:
                decision = "skip_score"
            elif not allowed:
                decision = "skip_filter"
            elif used >= free:
                decision = "skip_full"
            else:
                decision = "entry"
                used += 1
            rec = {
                "as_of": day, "code": ev["code"], "event_date": ev["event_date"],
                "entry_date": day, "evidence_score": s,
                "scores_json": json.dumps(ev["scores"], ensure_ascii=False,
                                          default=str),
                "inputs_json": json.dumps(
                    {"period_end": ev["period_end"],
                     "quarter_type": ev["quarter_type"],
                     # 発表日が推定か確定かは、後から検証するとき決定的に効く
                     "date_source": ev.get("date_source", "observed"),
                     "date_confidence": ev.get("confidence"),
                     "estimated_from": ev.get("estimated_from")},
                    ensure_ascii=False),
                "topix_close": topix_close, "topix_ma200": ma,
                "market_allowed": int(bool(allowed)), "decision": decision,
                "decision_note": None,
            }
            frozen.append(rec)
            if decision == "entry":
                fill = next_business_day(bd, day)
                entries.append({**rec, "entry_fill_date": fill,
                                "entry_price_assumed": open_price(mcon, ev["code"], fill),
                                "entry_price_close": price_on(pcon, ev["code"], day)})
        # フィルター作動の記録（機会損失の母数も残す）
        blocked = [e["code"] for e in cands if not allowed]
        if not dry_run:
            mcon.execute(
                "INSERT OR REPLACE INTO market_filter_log (date, topix_close, "
                " topix_ma200, allowed, n_candidates, n_blocked, blocked_codes, "
                " logged_at) VALUES (?,?,?,?,?,?,?,?)",
                (day, topix_close, ma, int(bool(allowed)), len(cands),
                 len(blocked), ",".join(blocked[:200]) or None, C.utcnow()))
    return frozen, entries


def decide_v3(pcon, mcon, scored, dry_run=False):
    """v3 シャドウの採否。**v2 の判定には一切触れない。**

    スコアリングは v2 と共通で、違うのは採否だけ:
      A1 available が V3_MIN_AVAILABLE 本未満なら採用しない
      B1 同一イベント日は V3_MAX_PER_DAY 件まで
         同一セクターの同時保有は V3_MAX_PER_SECTOR 件まで
    """
    idx = pcon.execute("SELECT date, close FROM market_index ORDER BY date").fetchall()
    idx_d = [r[0] for r in idx]
    idx_c = [r[1] for r in idx]
    bd = business_days(pcon)
    sectors = {r[0]: r[1] for r in pcon.execute("SELECT ticker, sector FROM universe")}

    by_day = {}
    for ev in scored:
        by_day.setdefault(ev["entry_date"], []).append(ev)

    frozen, entries = [], []
    for day in sorted(by_day):
        cands = sorted(by_day[day],
                       key=lambda e: (-(e["evidence_score"] or -1), e["code"]))
        allowed, _ma = V2.topix_ma_ok(idx_d, idx_c, day)
        # v3 の枠は v3 の建玉だけで数える（v2 とは独立のポートフォリオ）
        open_rows = mcon.execute(
            "SELECT code, sector FROM shadow_trades "
            "WHERE variant='v3' AND status='open' AND COALESCE(invalidated,0)=0 "
            "AND entry_date<=?", (day,)).fetchall()
        # DB の建玉 + この実行で既に建てた分。日跨ぎで累積させる。
        free = MAX_POSITIONS - len(open_rows) - len(entries)
        sec_open = {}
        for r in open_rows:
            sec_open[r["sector"]] = sec_open.get(r["sector"], 0) + 1
        for e in entries:
            sec_open[e["sector"]] = sec_open.get(e["sector"], 0) + 1
        used_day = 0
        for ev in cands:
            s_ = ev["evidence_score"]
            n_av = sum(1 for v in ev["scores"].values()
                       if isinstance(v, dict) and v.get("available"))
            sec = sectors.get(ev["code"])
            if s_ is None or s_ < SCORE_THRESHOLD:
                dec = "skip_score"
            elif n_av < V3_MIN_AVAILABLE:
                dec = "skip_evidence"
            elif not allowed:
                dec = "skip_filter"
            elif free <= 0 or used_day >= free:
                dec = "skip_full"
            elif used_day >= V3_MAX_PER_DAY:
                dec = "skip_daycap"
            elif sec_open.get(sec, 0) >= V3_MAX_PER_SECTOR:
                dec = "skip_sectorcap"
            else:
                dec = "entry"
                used_day += 1
                sec_open[sec] = sec_open.get(sec, 0) + 1
            rec = {"as_of": day, "code": ev["code"], "event_date": ev["event_date"],
                   "entry_date": day, "evidence_score": s_, "n_available": n_av,
                   "sector": sec, "decision": dec}
            frozen.append(rec)
            if dec == "entry":
                fill = next_business_day(bd, day)
                entries.append({**rec, "entry_fill_date": fill,
                                "entry_price_assumed": open_price(mcon, ev["code"], fill)})
    if not dry_run:
        for r in frozen:
            mcon.execute(
                "INSERT OR IGNORE INTO shadow_snapshots (variant, as_of, code, "
                " event_date, entry_date, evidence_score, n_available, sector, "
                " decision, frozen_at) VALUES ('v3',?,?,?,?,?,?,?,?,?)",
                (r["as_of"], r["code"], r["event_date"], r["entry_date"],
                 r["evidence_score"], r["n_available"], r["sector"],
                 r["decision"], C.utcnow()))
        for e in entries:
            sid = mcon.execute(
                "SELECT snapshot_id FROM shadow_snapshots WHERE variant='v3' "
                "AND as_of=? AND code=? AND event_date IS ?",
                (e["as_of"], e["code"], e["event_date"])).fetchone()
            mcon.execute(
                "INSERT OR IGNORE INTO shadow_trades (variant, snapshot_id, code, "
                " sector, event_date, entry_date, entry_fill_date, "
                " entry_price_assumed, position_size, status, opened_at, "
                " event_date_estimated) VALUES ('v3',?,?,?,?,?,?,?,?, 'open',?,?)",
                (sid[0] if sid else None, e["code"], e["sector"], e["event_date"],
                 e["entry_date"], e["entry_fill_date"], e["entry_price_assumed"],
                 POS_FRACTION, C.utcnow(), e["event_date"]))
        mcon.commit()
    return frozen, entries


def decide_variant(pcon, mcon, scored, variant, dry_run=False):
    """シャドウ B / C の採否。**v2 と v3 の判定には一切触れない。**

    B: スコア連動サイジング（合計エクスポージャ上限つき）。エグジットは
       4分岐だが、それは決済時の話なのでここでは建玉だけを決める。
    C: 閾値を 0.20 に上げるだけ。他は v2 と完全に同一。
    """
    idx = pcon.execute("SELECT date, close FROM market_index ORDER BY date").fetchall()
    idx_d = [r[0] for r in idx]
    idx_c = [r[1] for r in idx]
    bd = business_days(pcon)
    sectors = {r[0]: r[1] for r in pcon.execute("SELECT ticker, sector FROM universe")}
    thr = VC_SCORE_THRESHOLD if variant == "C" else SCORE_THRESHOLD

    by_day = {}
    for ev in scored:
        by_day.setdefault(ev["entry_date"], []).append(ev)

    frozen, entries = [], []
    for day in sorted(by_day):
        cands = sorted(by_day[day],
                       key=lambda e: (-(e["evidence_score"] or -1), e["code"]))
        allowed, _ma = V2.topix_ma_ok(idx_d, idx_c, day)
        open_rows = mcon.execute(
            "SELECT code, position_size FROM shadow_trades "
            "WHERE variant=? AND status='open' AND COALESCE(invalidated,0)=0 "
            "AND entry_date<=?", (variant, day)).fetchall()
        # DB の建玉 + この実行で既に建てた分。エクスポージャも同様に累積。
        free = MAX_POSITIONS - len(open_rows) - len(entries)
        exposure = (sum(r["position_size"] or 0 for r in open_rows)
                    + sum(e["size"] for e in entries))
        used = 0
        for ev in cands:
            s_ = ev["evidence_score"]
            size = (vb_position_size(s_) if variant == "B" else POS_FRACTION)                 if s_ is not None else 0.0
            if s_ is None or s_ < thr:
                dec = "skip_score"
            elif not allowed:
                dec = "skip_filter"
            elif free <= 0 or used >= free:
                dec = "skip_full"
            elif variant == "B" and exposure + size > VB_MAX_EXPOSURE + 1e-9:
                dec = "skip_exposure"
            else:
                dec = "entry"
                used += 1
                exposure += size
            rec = {"as_of": day, "code": ev["code"], "event_date": ev["event_date"],
                   "entry_date": day, "evidence_score": s_,
                   "n_available": sum(1 for v in ev["scores"].values()
                                      if isinstance(v, dict) and v.get("available")),
                   "sector": sectors.get(ev["code"]), "decision": dec, "size": size}
            frozen.append(rec)
            if dec == "entry":
                fill = next_business_day(bd, day)
                entries.append({**rec, "entry_fill_date": fill,
                                "entry_price_assumed": open_price(mcon, ev["code"], fill)})
    if not dry_run:
        for r in frozen:
            mcon.execute(
                "INSERT OR IGNORE INTO shadow_snapshots (variant, as_of, code, "
                " event_date, entry_date, evidence_score, n_available, sector, "
                " decision, frozen_at) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (variant, r["as_of"], r["code"], r["event_date"], r["entry_date"],
                 r["evidence_score"], r["n_available"], r["sector"], r["decision"],
                 C.utcnow()))
        for e in entries:
            sid = mcon.execute(
                "SELECT snapshot_id FROM shadow_snapshots WHERE variant=? "
                "AND as_of=? AND code=? AND event_date IS ?",
                (variant, e["as_of"], e["code"], e["event_date"])).fetchone()
            mcon.execute(
                "INSERT OR IGNORE INTO shadow_trades (variant, snapshot_id, code, "
                " sector, event_date, entry_date, entry_fill_date, "
                " entry_price_assumed, position_size, status, opened_at, "
                " event_date_estimated) VALUES (?,?,?,?,?,?,?,?,?, 'open',?,?)",
                (variant, sid[0] if sid else None, e["code"], e["sector"],
                 e["event_date"], e["entry_date"], e["entry_fill_date"],
                 e["entry_price_assumed"], e["size"], C.utcnow(), e["event_date"]))
        mcon.commit()
    return frozen, entries


def next_revision(mcon, table, where_sql, args):
    """訂正版を追記するための次の revision。旧版は残したまま増やす。"""
    r = mcon.execute("SELECT COALESCE(MAX(revision),0)+1 FROM %s WHERE %s"
                     % (table, where_sql), args).fetchone()
    return r[0]


def freeze(mcon, frozen, entries, dry_run=False):
    if dry_run:
        return 0, 0
    n_s = n_t = 0
    for r in frozen:
        cur = mcon.execute(
            "INSERT OR IGNORE INTO forecast_snapshots (as_of, code, event_date, "
            " entry_date, evidence_score, scores_json, inputs_json, topix_close, "
            " topix_ma200, market_allowed, decision, decision_note, frozen_at, "
            " revision) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (r["as_of"], r["code"], r["event_date"], r["entry_date"],
             r["evidence_score"], r["scores_json"], r["inputs_json"],
             r["topix_close"], r["topix_ma200"], r["market_allowed"],
             r["decision"], r["decision_note"], C.utcnow(),
             next_revision(mcon, "forecast_snapshots",
                           "as_of=? AND code=? AND event_date IS ?",
                           (r["as_of"], r["code"], r["event_date"]))))
        n_s += cur.rowcount
    for e in entries:
        sid = mcon.execute(
            "SELECT snapshot_id FROM forecast_snapshots WHERE as_of=? AND code=? "
            "AND event_date IS ? AND COALESCE(invalidated,0)=0 "
            "ORDER BY revision DESC LIMIT 1",
            (e["as_of"], e["code"], e["event_date"])).fetchone()
        cur = mcon.execute(
            "INSERT OR IGNORE INTO paper_trades (snapshot_id, code, event_date, "
            " entry_date, entry_fill_date, entry_price_assumed, entry_price_close, "
            " position_size, status, opened_at, event_date_estimated, revision) "
            "VALUES (?,?,?,?,?,?,?,?, 'open', ?, ?, ?)",
            (sid[0] if sid else None, e["code"], e["event_date"], e["entry_date"],
             e["entry_fill_date"], e["entry_price_assumed"], e["entry_price_close"],
             POS_FRACTION, C.utcnow(), e["event_date"],
             next_revision(mcon, "paper_trades",
                           "code=? AND entry_date=? AND event_date IS ?",
                           (e["code"], e["entry_date"], e["event_date"]))))
        n_t += cur.rowcount
    mcon.commit()
    return n_s, n_t


def report(frozen, entries, as_of):
    C.log("=== ペーパートレード週次レポート  判定週 %s ===" % as_of)
    n = len(frozen)
    by_dec = {}
    for r in frozen:
        by_dec[r["decision"]] = by_dec.get(r["decision"], 0) + 1
    C.log("  候補 %d 件" % n)
    for k in ("entry", "skip_score", "skip_filter", "skip_full"):
        if k in by_dec:
            C.log("    %-12s %d" % (k, by_dec[k]))
    if not entries:
        C.log("  建玉なし")
        return
    C.log("")
    C.log("  --- 建玉候補（採用順）---")
    C.log("  %-6s %-11s %-11s %6s  %-9s %-12s %s"
          % ("銘柄", "判定日", "約定日", "スコア", "想定始値", "対象四半期", "発表日ソース"))
    for e in entries:
        inp = json.loads(e["inputs_json"])
        C.log("  %-6s %-11s %-11s %6.3f  %9s  %-12s %s"
              % (e["code"], e["entry_date"], e["entry_fill_date"] or "-",
                 e["evidence_score"] or 0,
                 ("%.1f" % e["entry_price_assumed"]) if e["entry_price_assumed"] else "未確定",
                 inp.get("period_end", ""),
                 "%s/%s" % (inp.get("date_source", "?"), inp.get("date_confidence") or "-")))


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--as-of", help="判定週の基準日 YYYY-MM-DD（既定は今日）")
    p.add_argument("--dry-run", action="store_true", help="DBに書かない")
    p.add_argument("--lock-wait", type=float, default=600.0)
    a = p.parse_args(argv)

    as_of = date.fromisoformat(a.as_of) if a.as_of else date.today()
    lo, hi = week_span(as_of)
    pdb = os.path.join(C.DATA_DIR, "projection.db")
    pcon = sqlite3.connect("file:%s?mode=ro" % pdb.replace("\\", "/"), uri=True)
    # 別枠の data_access は r["col"] で読む。row_factory を付け忘れると
    # 全スコアラーが TypeError を出し、score_ticker がそれを available=False に
    # 握り潰すので、**設定ミスが「データなし」と見分けられなくなる**
    # （2026-09-02 に実際に踏み、1,773件が全部 skip_score になった）。
    pcon.row_factory = sqlite3.Row

    try:
        with C.writer_lock("paper_weekly", wait_seconds=a.lock_wait):
            mcon = C.init_db()
            # 成功の記録。火曜の健全性チェックがこれを見て、
            # 「その週の週次バッチが走ったか」を判定する。
            run_id = None if a.dry_run else C.start_run(
                mcon, "paper_weekly", lo.isoformat())
            last_obs = pcon.execute(
                "SELECT MAX(filing_date) FROM filings").fetchone()[0]
            if hi.isoformat() > (last_obs or ""):
                ev = upcoming_events_from_calendar(pcon, lo.isoformat(), hi.isoformat())
                src = "推定カレンダー"
            else:
                ev = upcoming_events(pcon, lo.isoformat(), hi.isoformat())
                src = "実績（過去週の再現）"
            C.log("判定週 %s..%s / ソース=%s / T-%d を迎えるイベント %d 件"
                  % (lo, hi, src, ENTRY_LEAD, len(ev)))
            scored = score_events(pcon, ev)
            frozen, entries = decide(pcon, mcon, scored, a.dry_run)
            n_s, n_t = freeze(mcon, frozen, entries, a.dry_run)
            # v3 シャドウ（v2 の判定・記録には触れない）
            f3, e3 = decide_v3(pcon, mcon, scored, a.dry_run)
            d3 = {}
            for r in f3:
                d3[r["decision"]] = d3.get(r["decision"], 0) + 1
            C.log("  [v3シャドウ] 建玉 %d / %s"
                  % (len(e3), " ".join("%s=%d" % kv for kv in sorted(d3.items()))))
            for var in ("B", "C"):
                fx, ex = decide_variant(pcon, mcon, scored, var, a.dry_run)
                dx = {}
                for r in fx:
                    dx[r["decision"]] = dx.get(r["decision"], 0) + 1
                expo = sum(r["size"] for r in ex)
                C.log("  [シャドウ%s] 建玉 %d (エクスポージャ %.0f%%) / %s"
                      % (var, len(ex), expo * 100,
                         " ".join("%s=%d" % kv for kv in sorted(dx.items()))))
            n_rec, n_flag = reconcile_dates(mcon, a.dry_run)
            if n_rec:
                C.log("  発表日の突合: %d 件 / 3営業日以上のズレ %d 件"
                      % (n_rec, n_flag))
            report(frozen, entries, "%s..%s" % (lo, hi))
            if a.dry_run:
                C.log("  (dry-run: DBには書いていない)")
            else:
                C.log("  凍結 %d 件 / 新規建玉 %d 件" % (n_s, n_t))
                C.finish_run(mcon, run_id, "ok", n_target=len(frozen),
                             n_saved=n_t,
                             note="frozen=%d entries=%d recon=%d flagged=%d"
                                  % (n_s, n_t, n_rec, n_flag))
    except C.WriterBusy as e:
        C.log("SKIPPED: %s" % e)
        return 0
    except Exception as e:
        # 失敗も記録に残す。黙って終わると火曜のチェックが
        # 「走らなかった」のか「落ちた」のか区別できない。
        try:
            con = C.init_db()
            rid = C.start_run(con, "paper_weekly", lo.isoformat())
            C.finish_run(con, rid, "failed", error=str(e)[:400])
        except Exception:                                # pragma: no cover
            pass
        raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
