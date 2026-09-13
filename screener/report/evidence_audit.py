"""screener/report/evidence_audit.py — スコア根拠の健全性監査（ユニバース全体・読み取り専用）。

なぜ要るか
----------
2026-09-13 の決算窓スキャンで上位2銘柄が偽陽性だった。
  3475 グッドコムアセット: 根拠期なし（別枠の位置ベース比較）。2024年の四半期データ由来、
                           売掛金/売上が極小で「DSO 0日→0日」なのに S1=1.0。
  2776 新都HD           : 売上 1,991→9,971百万円（5倍）と DSO 113.9→30.8日が同時に起き、
                           直後の下期は連結範囲変更で無効化。accounting_change=0 で検知漏れ。
**スコアが出たことではなく、スコアの根拠が特定・照合できるか**を全銘柄で数える。

1銘柄×1シグナルごとに記録すること
  source          span_matched（根拠期つき）/ external（別枠の位置ベース・根拠期なし）
  period / doc_date / doc_id
  doc_age_bucket  根拠期書類の日付を as_of からの絶対日数で 0-3m / 3-6m / 6-12m / 12m+
  doc_check       本体 financials_cum に (期, 四半期) の行があるか: 一致 / 不一致 / 照合不能
  fp_rules        偽陽性の機械的検出（下記）

偽陽性ルール（既知例から作った。閾値は既知例を確実に拾う側に置き、全件の件数を報告する）
  R1_no_period      発火シグナルに根拠期が無い（3475）
  R2_old_evidence   根拠期書類が as_of の 180日超前（2776: 2025-09-11）
  R3_degenerate_dso S1 の DSO がどちらかの期で 1日未満（3475: 0.0 / 0.1日）
  R4_scale_break    S1/S2 で当期/前年の売上比が 2.0 以上または 0.5 以下（2776: 5.0倍）
  R5_scope_change   根拠期と同じ会計年度の financials_q に「連結範囲変更」の無効行がある（2776 FY2026-Q4）

    python -m screener.report.evidence_audit --as-of 2026-09-13 --policy prefer_span \
        --out C:/screener_data/evidence_audit_20260913_before.csv
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import sqlite3
import sys
from collections import Counter, defaultdict
from datetime import date

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import backtest_eval as V1

OLD_DAYS = 180
SCALE_HI, SCALE_LO = 2.0, 0.5
DSO_MIN_DAYS = 1.0


def age_bucket(doc_date, as_of):
    if not doc_date:
        return "根拠期なし"
    days = (as_of - date.fromisoformat(doc_date[:10])).days
    if days <= 91:
        return "0-3m"
    if days <= 182:
        return "3-6m"
    if days <= 365:
        return "6-12m"
    return "12m+"


def doc_check(mcon, code, period, doc_id):
    """本体 financials_cum で (期, 四半期) の実在を確かめる。投影層は使わない。"""
    if not period or not doc_id:
        return "照合不能"
    f = mcon.execute("SELECT id FROM filings WHERE code=? AND doc_id=?", (code, doc_id)).fetchone()
    if not f:
        return "照合不能"
    try:
        fy, q = period.split("-Q")
        n = mcon.execute("SELECT COUNT(*) FROM financials_cum WHERE filing_id=? AND period=? "
                         "AND q_no=?", (f[0], fy, int(q))).fetchone()[0]
    except ValueError:
        return "照合不能"
    return "一致" if n else "不一致"


def scope_change_in_fy(mcon, code, period):
    if not period:
        return None
    fy = period.split("-Q")[0]
    r = mcon.execute("SELECT GROUP_CONCAT(DISTINCT period || '-Q' || q_no) FROM financials_q "
                     "WHERE code=? AND period=? AND valid_flag=0 AND invalid_reason LIKE '%連結範囲%'",
                     (code, fy)).fetchone()
    return r[0] if r and r[0] else None


def fp_rules(sig, v, as_of, mcon, code):
    if not (v.get("available") and (v.get("score") or 0) > 0):
        return []
    out = []
    if not v.get("period"):
        out.append("R1_no_period")
    dd = v.get("doc_date")
    if dd and (as_of - date.fromisoformat(dd[:10])).days > OLD_DAYS:
        out.append("R2_old_evidence")
    det = v.get("details") or {}
    if sig == "S1":
        if min(det.get("dso_now", 99), det.get("dso_prev_year", 99)) < DSO_MIN_DAYS:
            out.append("R3_degenerate_dso")
    if sig in ("S1", "S2"):
        s_now, s_prev = v.get("_sales_now"), v.get("_sales_prev")
        if s_now and s_prev and s_prev > 0:
            ratio = s_now / s_prev
            if ratio >= SCALE_HI or ratio <= SCALE_LO:
                out.append("R4_scale_break(x%.2f)" % ratio)
    sc = scope_change_in_fy(mcon, code, v.get("period"))
    if sc:
        out.append("R5_scope_change(%s)" % sc)
    return out


def _sales_pair(v):
    """evidence 文から「調整後売上 A→B百万円」を読む（S1）。S2 は投影層から引く。"""
    import re
    m = re.search(r"調整後売上\s*([\d.]+)→([\d.]+)百万円", v.get("evidence") or "")
    if m:
        return float(m.group(2)), float(m.group(1))
    return None, None


def run(as_of, policy, codes=None):
    root = V1._external_root()
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from module_b.run_scorers import SCORERS_ALL
    from screener.signals.span_runner import score_ticker

    pdb = os.path.join(C.DATA_DIR, "projection.db").replace("\\", "/")
    pcon = sqlite3.connect("file:%s?mode=ro" % pdb, uri=True)
    pcon.row_factory = sqlite3.Row
    mcon = sqlite3.connect("file:%s?mode=ro" % C.DB_PATH.replace("\\", "/"), uri=True)
    mcon.row_factory = sqlite3.Row
    names = {r["code"]: r["name"] for r in mcon.execute(
        "SELECT code, name FROM companies WHERE universe_flag=1")}
    codes = codes or sorted(names)

    sig_rows, tick_rows = [], []
    for code in codes:
        try:
            res = score_ticker(pcon, code, SCORERS_ALL, as_of=as_of, policy=policy)
        except Exception as e:
            tick_rows.append({"code": code, "name": names.get(code, ""), "evidence_score": "",
                              "error": "%s: %s" % (type(e).__name__, e)})
            continue
        src = res.get("_source") or {}
        fired, rules_all = [], []
        for sig, v in sorted(res.items()):
            if not isinstance(v, dict):
                continue
            if sig == "S2" and v.get("period"):
                r = pcon.execute("SELECT sales FROM quarterly_standalone_all WHERE ticker=? AND "
                                 "period_end=? AND span_q=? AND sales IS NOT NULL",
                                 (code, v["period"], v.get("span_q") or 1)).fetchone()
                r0 = pcon.execute("SELECT sales FROM quarterly_standalone_all WHERE ticker=? AND "
                                  "period_end=? AND span_q=? AND sales IS NOT NULL",
                                  (code, v.get("peer_period"), v.get("span_q") or 1)).fetchone()
                v["_sales_now"], v["_sales_prev"] = (r[0] if r else None), (r0[0] if r0 else None)
            elif sig == "S1":
                v["_sales_now"], v["_sales_prev"] = _sales_pair(v)
            is_fired = bool(v.get("available") and (v.get("score") or 0) > 0)
            rules = fp_rules(sig, v, as_of, mcon, code)
            if is_fired:
                fired.append(sig)
                rules_all += ["%s:%s" % (sig, x) for x in rules]
            sig_rows.append({
                "code": code, "name": names.get(code, ""), "signal": sig,
                "source": src.get(sig, ""), "available": int(bool(v.get("available"))),
                "score": v.get("score"), "fired": int(is_fired),
                "period": v.get("period") or "", "peer_period": v.get("peer_period") or "",
                "doc_id": v.get("doc_id") or "", "doc_date": v.get("doc_date") or "",
                "doc_age": age_bucket(v.get("doc_date"), as_of) if v.get("available") else "",
                "doc_check": doc_check(mcon, code, v.get("period"), v.get("doc_id"))
                if v.get("available") else "",
                "fp_rules": ";".join(rules), "evidence": (v.get("evidence") or "")[:200],
                # evidence_strict の監査列（prefer_span では空）
                "strict_flags": ";".join(v.get("strict_flags") or []),
                "strict_notes": ";".join(v.get("strict_notes") or []),
                "raw_score": v.get("raw_score", ""),
                "evidence_lag_q": "" if v.get("evidence_lag_q") is None else v.get("evidence_lag_q"),
            })
        tick_rows.append({"code": code, "name": names.get(code, ""),
                          "evidence_score": res.get("evidence_score"),
                          "fired": ",".join(fired), "fp_rules": " | ".join(rules_all), "error": ""})
    return sig_rows, tick_rows


def summarize(sig_rows, tick_rows):
    lines = []
    scored = [t for t in tick_rows if t["evidence_score"] not in ("", None)]
    lines.append("銘柄: %d / スコアあり %d / 無評価 %d / 例外 %d" % (
        len(tick_rows), len(scored), sum(1 for t in tick_rows if t["evidence_score"] is None),
        sum(1 for t in tick_rows if t.get("error"))))
    av = [r for r in sig_rows if r["available"]]
    lines.append("available なシグナル: %d / うち根拠期あり %d / 根拠期なし %d" % (
        len(av), sum(1 for r in av if r["period"]), sum(1 for r in av if not r["period"])))
    lines.append("  source 別(available): %s" % dict(Counter(r["source"] for r in av)))
    lines.append("  シグナル×根拠期の有無(available): %s" % dict(
        Counter((r["signal"], "期あり" if r["period"] else "期なし") for r in av)))
    fired = [r for r in sig_rows if r["fired"]]
    lines.append("発火シグナル: %d / 根拠期なしで発火 %d" % (len(fired), sum(1 for r in fired if not r["period"])))
    lines.append("  doc_date 絶対日付分布(available・根拠期あり): %s" % dict(
        Counter(r["doc_age"] for r in av if r["period"])))
    lines.append("  doc_date 絶対日付分布(発火・根拠期あり): %s" % dict(
        Counter(r["doc_age"] for r in fired if r["period"])))
    lines.append("  financials_cum 照合(available): %s" % dict(Counter(r["doc_check"] for r in av)))
    rule_ct = Counter()
    for r in fired:
        for x in (r["fp_rules"] or "").split(";"):
            if x:
                rule_ct[x.split("(")[0]] += 1
    lines.append("偽陽性ルール該当(発火シグナル単位): %s" % dict(rule_ct))
    lines.append("偽陽性ルール該当銘柄(スコア>0 で1本以上該当): %d" % sum(
        1 for t in tick_rows if t.get("fp_rules") and (t["evidence_score"] or 0) > 0))
    return lines


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--as-of", required=True)
    p.add_argument("--policy", default="prefer_span")
    p.add_argument("--codes", nargs="*")
    p.add_argument("--out", required=True, help="シグナル単位CSV。銘柄単位は _tickers.csv")
    a = p.parse_args(argv)
    as_of = date.fromisoformat(a.as_of)
    sig_rows, tick_rows = run(as_of, a.policy, a.codes)
    for path, rows in ((a.out, sig_rows), (a.out.replace(".csv", "_tickers.csv"), tick_rows)):
        keys = []
        for r in rows:
            keys += [k for k in r if k not in keys]
        with open(path, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=keys)
            w.writeheader()
            w.writerows(rows)
    for line in summarize(sig_rows, tick_rows):
        C.log(line)
    C.log("CSV: %s" % a.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
