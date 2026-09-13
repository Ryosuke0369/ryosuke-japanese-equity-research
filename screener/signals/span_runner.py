"""screener/signals/span_runner.py — 別枠スコアラーと span-matched 版の合流点。

**別枠のコードは1行も変更しない。** 呼ぶ側でどちらの結果を採るかを決める。

方針(policy)
------------
``prefer_span`` (規則B。設計書 §9、2026-09-02 承認)
    span 版が available ならつねに span 版を使う。span 版が使えない期は
    **別枠の位置ベース比較（根拠期なし）がそのまま通る**。

``span_only``   (旧規則A相当。比較実験用に残す)
    比較粒度が span_matched の期だけ span 版を使い、真の四半期の期は
    別枠の結果をそのまま通す。規則A/Bの差分を測るときに使う。

``evidence_strict`` (2026-09-13 追加。calibration_backlog §31)
    prefer_span に次の4つを足す。**いずれも「根拠が特定・照合できない証拠は
    点にしない」**という1つの原則の具体化。
      1. 根拠期が無い結果（別枠の位置ベース比較・S3/S6/S7/S8 等）は不採用
      2. 根拠期が評価時点の直前四半期から lag 期遅れ: lag=0 は満額、lag=1 は ×0.5、
         lag>=2 は不採用。決算またぎでは直前四半期の証拠でなければ先行性が無い
      3. S1/S2 で当期/前年の売上比が 2.0 以上・0.5 以下（連結範囲変更・M&A 由来の
         前年比破壊の疑い）は不採用
      4. S1 の DSO がどちらかの期で 1日未満（売掛金が売上に対して極小で比が壊れる）は不採用
    不採用にしたシグナルは available=False・score=0 にし、**消さずに** `strict_flags` と
    `raw_score` を残す（監査のため）。

どちらを使ったかは結果の `source` に残す。後から層別できないと、
成績が上がっても下がっても原因が分からない。
"""
from __future__ import annotations

import calendar as _cal
import os
import re
import sys
from datetime import date, timedelta

try:
    from screener import common as C            # noqa: F401
except ImportError:                             # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C            # noqa: F401

from screener.projection import span_matched as SM
from screener.signals import span_scorers as SP

POLICIES = ("prefer_span", "span_only", "evidence_strict")
# 既定の切替は calibration_backlog §31 のシャドウ比較を報告してから行う。
# それまでは明示指定のときだけ evidence_strict を使う（paper/shadow/sanity を黙って変えない）。
DEFAULT_POLICY = "prefer_span"

# evidence_strict のパラメータ（§31 に記録。変更するなら記録先行）
RECENCY_WEIGHT = {0: 1.0, 1: 0.5}      # lag>=2 は不採用
REPORT_LAG_DAYS = 45                   # 四半期末から開示までの法定目安
SCALE_HI, SCALE_LO = 2.0, 0.5
DSO_MIN_DAYS = 1.0
SCOPE_SIGNALS = ("S1", "S2")


def _as_date(as_of):
    if as_of is None:
        return date.today()
    return as_of if isinstance(as_of, date) else date.fromisoformat(str(as_of)[:10])


def fy_end_month(con, ticker):
    try:
        r = con.execute("SELECT fiscal_year_end FROM universe WHERE ticker=?", (ticker,)).fetchone()
    except Exception:
        return None
    return int(r[0]) if r and r[0] else None


def period_ym(period, fym):
    """FY{Y}-Q{q} → 期末 (年, 月)。FY{Y} は期末が Y 年 fym 月の年度。"""
    m = re.match(r"FY(\d{4})-Q(\d)", period or "")
    if not m or not fym:
        return None
    t = int(m.group(1)) * 12 + (fym - 1) - 3 * (4 - int(m.group(2)))
    return (t // 12, t % 12 + 1)


def expected_latest_ym(as_of, fym):
    """as_of 時点で開示が出揃っているはずの最新の四半期末 (年, 月)。

    四半期末は期末月と3ヶ月おきの月末。その月末 + 45日 <= as_of を満たす最新のもの。
    """
    d = _as_date(as_of)
    y, m = d.year, d.month
    for _ in range(15):
        if (m - fym) % 3 == 0:
            end = date(y, m, _cal.monthrange(y, m)[1])
            if end + timedelta(days=REPORT_LAG_DAYS) <= d:
                return (y, m)
        m -= 1
        if m == 0:
            y, m = y - 1, 12
    return None


def lag_quarters(period, as_of, fym):
    ev, ex = period_ym(period, fym), expected_latest_ym(as_of, fym) if fym else None
    if not ev or not ex:
        return None
    return max(0, ((ex[0] * 12 + ex[1]) - (ev[0] * 12 + ev[1])) // 3)


def _sales(con, ticker, pe, span):
    if not pe:
        return None
    r = con.execute("SELECT sales FROM quarterly_standalone_all WHERE ticker=? AND period_end=? "
                    "AND span_q=? AND sales IS NOT NULL", (ticker, pe, span or 1)).fetchone()
    return r[0] if r else None


def _scope_change_same_fy(con, ticker, period):
    m = re.match(r"FY(\d{4})", period or "")
    if not m:
        return None
    r = con.execute("SELECT GROUP_CONCAT(DISTINCT period_end) FROM quarterly_standalone_all "
                    "WHERE ticker=? AND fiscal_year=? AND is_valid=0 "
                    "AND invalid_reason LIKE '%連結範囲%'", (ticker, int(m.group(1)))).fetchone()
    return r[0] if r and r[0] else None


def apply_strict(con, ticker, sid, r, src, as_of, fym):
    """1シグナルに evidence_strict を当てる。新しい dict を返す（元を壊さない）。"""
    r = dict(r)
    if not r.get("available"):
        return r
    flags, notes = [], []
    raw = r.get("score") or 0.0
    period = r.get("period")
    if src != "span_matched" or not period:
        flags.append("no_period")
    else:
        lag = lag_quarters(period, as_of, fym)
        r["evidence_lag_q"] = lag
        if lag is None:
            flags.append("recency_unknown")
        elif lag not in RECENCY_WEIGHT:
            flags.append("stale_lag%d" % lag)
        else:
            r["recency_weight"] = RECENCY_WEIGHT[lag]
        if sid in SCOPE_SIGNALS and r.get("peer_period"):
            # スコアラーが実際に比べた組（前年が Q1+Q2 の合成でもそのまま）から売上を取る。
            # span=2 の前年行を直接引くと、合成された前年（2776 の FY2025-Q2）で空振りする。
            s_now = s_prev = None
            for row in SM.yoy_series(con, ticker, "sales"):
                if row["period_end"] == period and row.get("peer_period") == r.get("peer_period"):
                    s_now, s_prev = row["value"], row["peer_value"]
                    break
            if s_now is None:
                s_now = _sales(con, ticker, period, r.get("span_q"))
                s_prev = _sales(con, ticker, r.get("peer_period"), r.get("span_q"))
            if s_now and s_prev and s_prev > 0:
                ratio = s_now / s_prev
                r["sales_ratio_yoy"] = round(ratio, 3)
                if ratio >= SCALE_HI or ratio <= SCALE_LO:
                    flags.append("scope_change_suspect(x%.2f)" % ratio)
            sc = _scope_change_same_fy(con, ticker, period)
            if sc:
                notes.append("same_fy_scope_invalid(%s)" % sc)
        if sid == "S1":
            det = r.get("details") or {}
            if min(det.get("dso_now", 99), det.get("dso_prev_year", 99)) < DSO_MIN_DAYS:
                flags.append("degenerate_dso")
    r["raw_score"] = raw
    r["strict_flags"] = flags
    r["strict_notes"] = notes
    if flags:
        r["available"] = False
        r["score"] = 0.0
        r["evidence"] = "[不採用:%s] %s" % (",".join(flags), r.get("evidence") or "")
    else:
        w = r.get("recency_weight", 1.0)
        r["score"] = round(raw * w, 3)
        if w != 1.0:
            r["evidence"] = "[直前四半期より1期古い: ×%.1f] %s" % (w, r.get("evidence") or "")
    return r


def score_ticker(con, ticker, scorers, as_of=None, policy: str = DEFAULT_POLICY):
    """別枠 run_scorers.score_ticker と同じ形の dict を返す。

    scorers は別枠の {signal_id: callable}。S1/S2/S4/S5 だけ span 版と
    突き合わせ、残りはそのまま別枠に委ねる。
    """
    if policy not in POLICIES:
        raise ValueError("policy は %s のいずれか: %r" % (POLICIES, policy))
    out, detail, strict = {}, {}, {}
    fym = fy_end_month(con, ticker) if policy == "evidence_strict" else None
    for sid, fn in scorers.items():
        try:
            base = fn(con, ticker, as_of)
        except Exception as e:                  # 別枠が落ちても他を止めない
            base = {"score": 0.0, "available": False,
                    "evidence": "別枠スコアラーが例外: %s" % e, "details": {}}
        r, src = base, "external"
        sp_fn = SP.SPAN_SCORERS.get(sid)
        if sp_fn is not None:
            try:
                cand = sp_fn(con, ticker, as_of)
            except Exception as e:
                cand = {"score": 0.0, "available": False,
                        "evidence": "span版が例外: %s" % e, "mode": None}
            if cand.get("available") and (
                    policy in ("prefer_span", "evidence_strict")
                    or cand.get("mode") == "span_matched"):
                r, src = cand, "span_matched"
        if policy == "evidence_strict":
            r = apply_strict(con, ticker, sid, r, src, as_of, fym)
            strict[sid] = r.get("strict_flags") or []
        # **別枠 score_ticker と同じ形で返す**（各シグナルは dict のまま）。
        # weekly_screen の _fired が dict を前提にしているので、ここで
        # スコアだけに潰すと発火判定が丸ごと死ぬ。
        out[sid] = dict(r, source=src)
        detail[sid] = src
    avail = [v["score"] for v in out.values() if v["available"]]
    out["evidence_score"] = round(sum(avail) / len(avail), 3) if avail else None
    out["_source"] = detail
    if policy == "evidence_strict":
        out["_strict"] = strict
    return out
