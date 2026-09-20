"""screener/report/event_strata.py — シャドウE フェーズA 追加測定（E-A2）。

定義の正本は docs/backtest_acceptance_criteria.md「E-A2」（2026-09-20・計算前に確定）。
**このスクリプトは定義を実装するだけで、条件や閾値を決めない。**

測るもの
--------
E-A2-1 スコア発火（合成スコア ≥ 0.10）に限定したイベント→価格反応。
       スコアは `span_runner.score_ticker(policy="evidence_strict")` を
       **as_of = そのイベントの T−15 の営業日**で計算する（エントリー時点の情報だけ）。
E-A2-2 業績予想の修正の方向別（`company_forecasts` の forecast_op の増減）。

なぜテクニカル層に置かないのか
------------------------------
`screener.technical` はスコアラーを呼べない（E-0・test_technical_isolation）。
スコアで層別するのは**測定側の仕事**なので、report 側に置いてイベント反応の
CSV（フェーズAの出力）を読む。スコアへは何も書き戻さない。

    python -m screener.report.event_strata
"""
from __future__ import annotations

import argparse
import csv
import os
import sqlite3
import sys
from collections import defaultdict

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import backtest_eval as V1
from screener.signals.span_runner import score_ticker
from screener.technical import event_response as E


def scorers_all():
    """週次スキャン（weekly_screen.collect）と同じシグナル集合。

    span 版4本（SPAN_SCORERS）だけを渡すと合成スコアの分母が変わり、
    **週次で見ているスコアとは別物**になる。そろえないと層別の意味が消える。
    """
    root = str(V1._external_root())
    if root not in sys.path:
        sys.path.insert(0, root)
    from module_b.run_scorers import SCORERS_ALL
    return SCORERS_ALL

SCORE_THRESHOLD = 0.10          # v2 と同じ（E-A2-1）
POLICY = "evidence_strict"
ENTRY_OFFSET = -15              # as_of は T−15（エントリー時点）


def load_events(events_csv, long_csv):
    """イベント CSV に、long CSV から T−15 の日付を付けて返す。"""
    entry_date = {}
    with open(long_csv, encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            if int(r["offset"]) == ENTRY_OFFSET:
                entry_date[r["event_id"]] = r["date"]
    out = []
    with open(events_csv, encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            if r["window_ok"] != "1":
                continue
            r["entry_date"] = entry_date.get(r["event_id"])
            for k in ("f_pre_rel_topix", "y_init_rel_topix", "y_drift_rel_topix"):
                r[k] = float(r[k]) if r[k] else None
            out.append(r)
    return out


def attach_scores(events, pcon):
    """T−15 時点の合成スコア。評価不能は None のまま残す（0 にしない）。

    `pcon.row_factory = sqlite3.Row` が要る（採点器は行を名前で引く）。
    """
    if pcon.row_factory is not sqlite3.Row:
        raise RuntimeError("pcon.row_factory に sqlite3.Row が要る（全件スコアなしになる）")
    n_err = 0
    scorers = scorers_all()
    for e in events:
        e["score"] = None
        if not e["entry_date"]:
            continue
        try:
            res = score_ticker(pcon, e["code"], scorers,
                               as_of=e["entry_date"], policy=POLICY)
            e["score"] = res.get("evidence_score")
        except Exception:
            n_err += 1
    return n_err


def score_group(e):
    if e["score"] is None:
        return "スコアなし"
    return "発火(≥0.10)" if e["score"] >= SCORE_THRESHOLD else "非発火(<0.10)"


_DIR_JA = {"up": "上方", "down": "下方", "flat": "横ばい", "initial": "不明"}


def revision_direction(mcon, code, disclosed_date, cache={}):
    """業績予想の修正の方向（E-A2-2）。

    出典は本体 `guidance.revision_direction`（2026-09-20 に出典変更。事前登録 E-A2-2 の追記）。
    修正開示の XBRL から「今回予想」と「前回予想」を読んで決めた値で、
    営業利益の予想を見る。旧出典（projection の company_forecasts 前後比較）は
    同一年度に2点を持たず、275件中0件しか決まらなかった。
    """
    key = (code, disclosed_date)
    if key in cache:
        return cache[key]
    # **その開示日ちょうど**の行を引く。近い日の短信（initial）を拾うと、
    # 修正の方向ではなく「直近に出ていた予想」を報告することになる。
    row = mcon.execute(
        "SELECT revision_direction FROM guidance "
        "WHERE code=? AND item='operating_income' AND date=?", (code, disclosed_date)).fetchone()
    out = _DIR_JA.get(row[0] if row else None, "不明")
    cache[key] = out
    return out


def table(title, groups, key):
    L = ["### %s\n" % title, E.HDR]
    for label, es in groups:
        L.append(E._stat_row(label, E.describe([(e.get(key), e["t0"]) for e in es])))
    L.append("")
    return L


def render(events, n_err, meta):
    L = ["# シャドウE フェーズA 追加測定（E-A2・記述統計）\n"]
    for k, v in meta.items():
        L.append("- %s: %s" % (k, v))
    L.append("\n## E-A2-1. スコア発火に限定した反応（as_of = T−15・policy=%s）\n" % POLICY)
    L.append("スコアの内訳: " + " / ".join(
        "%s %d" % (g, sum(1 for e in events if score_group(e) == g))
        for g in ("発火(≥0.10)", "非発火(<0.10)", "スコアなし"))
        + "（スコア計算で例外 %d）" % n_err)
    for typ in (E.TYPE_Q, E.TYPE_FY, E.TYPE_REV, E.TYPE_DIV):
        es = [e for e in events if e["type"] == typ]
        if not es:
            continue
        groups = [("全体", es)]
        for g in ("発火(≥0.10)", "非発火(<0.10)", "スコアなし"):
            groups.append((g, [e for e in es if score_group(e) == g]))
        L.append("\n## %s\n" % typ)
        L += table("① 事前 T−15〜T−1（対TOPIX）", groups, "f_pre_rel_topix")
        L += table("② 初動 T+0〜T+1（対TOPIX）", groups, "y_init_rel_topix")
        L += table("③ ドリフト T+2〜T+20（対TOPIX）", groups, "y_drift_rel_topix")
    L.append("\n## E-A2-2. 業績予想の修正の方向別\n")
    rev = [e for e in events if e["type"] == E.TYPE_REV]
    groups = [("全体", rev)]
    for d in ("上方", "下方", "横ばい", "不明"):
        groups.append((d, [e for e in rev if e["rev_dir"] == d]))
    L += table("② 初動 T+0〜T+1（対TOPIX）", groups, "y_init_rel_topix")
    L += table("③ ドリフト T+2〜T+20（対TOPIX）", groups, "y_drift_rel_topix")
    L += table("① 事前 T−15〜T−1（対TOPIX）", groups, "f_pre_rel_topix")
    return "\n".join(L) + "\n"


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--events", default=os.path.join(C.DATA_DIR, "event_price_response_events.csv"))
    p.add_argument("--long", default=os.path.join(C.DATA_DIR, "event_price_response.csv"))
    p.add_argument("--projection", default=os.path.join(C.DATA_DIR, "projection.db"))
    p.add_argument("--db", default=C.DB_PATH, help="本体DB（予想の改訂方向を読む）")
    p.add_argument("--out", default=os.path.join(C.DATA_DIR, "event_price_response_strata.md"))
    p.add_argument("--csv-out", default=os.path.join(C.DATA_DIR, "event_price_response_scored.csv"))
    a = p.parse_args(argv)

    events = load_events(a.events, a.long)
    pcon = sqlite3.connect("file:%s?mode=ro" % a.projection.replace("\\", "/"), uri=True)
    # 採点器は行を名前で引く。row_factory を付け忘れると例外にならず、
    # 全シグナルが「評価不能」になって静かに全件スコアなしになる（2026-09-20 に踏んだ）。
    pcon.row_factory = sqlite3.Row
    C.log("イベント %d 件にスコアを付ける（as_of は T−15）" % len(events))
    n_err = attach_scores(events, pcon)
    mcon = sqlite3.connect("file:%s?mode=ro" % a.db.replace("\\", "/"), uri=True)
    for e in events:
        e["rev_dir"] = (revision_direction(mcon, e["code"], e["date"])
                        if e["type"] == E.TYPE_REV else "")
    with open(a.csv_out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["event_id", "code", "type", "t0", "entry_date", "score",
                    "score_group", "rev_dir", "f_pre_rel_topix",
                    "y_init_rel_topix", "y_drift_rel_topix"])
        for e in events:
            w.writerow([e["event_id"], e["code"], e["type"], e["t0"], e["entry_date"],
                        "" if e["score"] is None else e["score"], score_group(e),
                        e["rev_dir"]]
                       + ["" if e[k] is None else round(e[k], 6) for k in
                          ("f_pre_rel_topix", "y_init_rel_topix", "y_drift_rel_topix")])
    meta = {"イベント（窓あり）": len(events), "投影DB": a.projection,
            "スコア": "%s / 閾値 %.2f / as_of は T−15" % (POLICY, SCORE_THRESHOLD)}
    text = render(events, n_err, meta)
    with open(a.out, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(text)
    C.log("出力: %s / %s" % (a.out, a.csv_out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
