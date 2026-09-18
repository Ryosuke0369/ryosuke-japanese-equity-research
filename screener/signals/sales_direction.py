"""screener/signals/sales_direction.py — S1 の売上方向ガード（2026-09-18）。

なぜ要るのか
------------
S1（DSO 改善）は「売掛金 ÷ 売上 × 日数」の前年比で点を付ける。分子が
減れば DSO は下がるが、**分子が減る理由は2つある**。

  (a) 同じ売上をより速く回収した                     → 改善（S1 の意図）
  (b) 売上そのものが縮んだので債権も一緒に縮んだ     → 縮小（改善ではない）

別枠の実装は「前年同期と比べて売上が減っていたら score を ×0.5」という
半減ガードを持つ（`scorers_s1_s4.score_s1_dso`）。これは **累計 span の
前年比**でしか効かない。6838 多摩川HD（2026年10月期）は根拠期が H1
（FY2026-Q2 / span=2）で、その H1 の売上は前年 H1 比 +45.2% だったため
半減ガードは作動せず、S1 は満点 1.0 を返した。一方で同じ時点の
**Q単独**売上は 2,051 → 1,691.9 → 1,559.2 百万円と3期連続で減っており、
Q単独 OPM も 24.2% → 15.1% → 11.3% と低下していた。つまり
**減速の証拠が改善シグナルとして加点されていた**。

ここで何を測るか
----------------
「いま売上は伸びているのか」を答えるのは **最新の売上データ**であって、
S1 が根拠にした期ではない。根拠期だけを見ると、より新しい四半期が
DB にあってもそれを無視することになる（累計値で判定して単独値の減速を
見落とす、という同じ穴）。そこで2か所で測る。

  latest   : 最新の売上タイルの方向。**発火の可否はこれで決める。**
  evidence : S1 が根拠にした期の方向。報告にだけ出す。

タイルと run-rate
-----------------
`quarterly_standalone_all` は同じ期に span=1（3ヶ月）と span=2（半期）が
並存する。期ごとに**最も細かい粒度を1本**採って時系列を敷き詰め
（`span_matched.yoy_series` と同じ規則B）、値は **span で割った
1四半期あたりの run-rate** に正規化する。6ヶ月の値と3ヶ月の値をそのまま
並べたら「前期比」の意味が壊れる —— これは累計値で判定するのと同じ誤り。

判定
----
  up          : qoq / yoy のうち取れたものの**いずれかが正**
  down        : 取れたものが**すべて 0 以下**
  unverified  : qoq も yoy も取れない（比較相手が無い）

`quarter_level` は判定に span=1 のタイルを使えたかどうか。False のときは
「Q単独では確認できていない」という警告であって、判定そのものではない。
"""
from __future__ import annotations

import os
import sys

try:
    from screener import common as C            # noqa: F401
except ImportError:                             # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C            # noqa: F401

from screener.projection import span_matched as SM

# 直近推移を何期ぶん報告に載せるか（週次レポートの「Q単独売上の推移」）。
TREND_N = 3


def tiles(con, ticker, vis=None, item="sales"):
    """期ごとに最も細かい粒度を1本だけ採った時系列。昇順。

    `span_matched.yoy_series` の finest 選択と同じ規則。span=4（通期）は
    四半期系列に入れない —— 年次の比較を混ぜると前期比の意味が壊れる。
    """
    rows = con.execute(
        "SELECT period_end, period_start, span_q, %s AS v FROM quarterly_standalone_all "
        "WHERE ticker=? AND is_valid=1 AND %s IS NOT NULL AND span_q<4"
        % (item, item), (ticker,)).fetchall()
    finest = {}
    for r in rows:
        pe, sp = r[0], (r[2] or 1)
        if vis is not None and pe not in vis:
            continue
        if pe not in finest or sp < (finest[pe][2] or 1):
            finest[pe] = r
    out = []
    for pe in sorted(finest):
        pe, ps, sp, v = finest[pe]
        sp = sp or 1
        end = SM.parse_pe(pe)
        if not end:
            continue
        out.append({"period_end": pe, "period_start": ps, "span_q": sp,
                    "value": v, "run_rate": v / sp,
                    "seq_end": SM._seq(*end), "seq_start": SM._seq(*end) - sp + 1})
    return out


def _prev_tile(ts, i):
    """i 番目のタイルの**直前に接する**タイル。重なっていたら採らない。

    同じ期に span=1 と span=2 が並ぶので、単に ts[i-1] を採ると
    「H1 の直前は Q2」のような重複区間を前期として扱ってしまう。
    終端がぴったり接することを要求する（S4 の積み上げと同じ規則）。
    """
    want = ts[i]["seq_start"] - 1
    for j in range(i - 1, -1, -1):
        if ts[j]["seq_end"] == want:
            return ts[j]
    return None


def _yoy_tile(ts, i):
    """1年前（4四半期前）に終わる、同じ span のタイル。"""
    cur = ts[i]
    for j in range(i - 1, -1, -1):
        if (ts[j]["seq_end"] == cur["seq_end"] - 4
                and ts[j]["span_q"] == cur["span_q"]):
            return ts[j]
    return None


def _pct(now, before):
    if now is None or before is None or before <= 0:
        return None
    return (now / before - 1) * 100.0


def _direction_at(ts, i):
    """タイル i の qoq / yoy（run-rate ベース、%）と判定。"""
    cur = ts[i]
    p, y = _prev_tile(ts, i), _yoy_tile(ts, i)
    qoq = _pct(cur["run_rate"], p["run_rate"] if p else None)
    yoy = _pct(cur["run_rate"], y["run_rate"] if y else None)
    have = [x for x in (qoq, yoy) if x is not None]
    if not have:
        status = "unverified"
    elif any(x > 0 for x in have):
        status = "up"
    else:
        status = "down"
    return {"period_end": cur["period_end"], "span_q": cur["span_q"],
            "run_rate": round(cur["run_rate"], 1),
            "qoq_pct": None if qoq is None else round(qoq, 1),
            "yoy_pct": None if yoy is None else round(yoy, 1),
            "qoq_peer": p["period_end"] if p else None,
            "yoy_peer": y["period_end"] if y else None,
            "status": status,
            "quarter_level": cur["span_q"] == 1}


def trend(ts, n=TREND_N):
    """直近 n タイルの run-rate 推移。報告にそのまま出せる形にする。"""
    out = []
    for t in ts[-n:]:
        out.append({"period_end": t["period_end"], "span_q": t["span_q"],
                    "sales": round(t["value"], 1),
                    "per_quarter": round(t["run_rate"], 1)})
    return out


def trend_text(ts, n=TREND_N):
    """「FY2026-Q1(span=1) 2051 → …」の1行。span を必ず書く。"""
    parts = []
    for t in trend(ts, n):
        label = t["period_end"]
        if t["span_q"] != 1:
            label += "(span=%d/1Qあたり%.0f)" % (t["span_q"], t["per_quarter"])
        parts.append("%s %.0f百万円" % (label, t["sales"]))
    return " → ".join(parts)


def evaluate(con, ticker, evidence_period=None, vis=None):
    """S1 の売上方向ガード。

    戻り値:
      status          : up / down / unverified（**latest の判定**）
      quarter_level   : latest の判定に span=1 を使えたか
      latest          : 最新タイルの方向（判定に使う）
      evidence        : 根拠期の方向（報告用。取れなければ None）
      trend           : 直近3タイルの run-rate 推移
      trend_text      : その1行表現
      note            : 人が読む1行
    """
    ts = tiles(con, ticker, vis)
    if not ts:
        return {"status": "unverified", "quarter_level": False,
                "latest": None, "evidence": None, "trend": [], "trend_text": "",
                "note": "売上の四半期系列が無い"}
    latest = _direction_at(ts, len(ts) - 1)
    ev = None
    if evidence_period:
        for i, t in enumerate(ts):
            if t["period_end"] == evidence_period:
                ev = _direction_at(ts, i)
                break
    def _fmt(d):
        if not d:
            return "-"
        return ("%s(span=%d) qoq %s / yoy %s"
                % (d["period_end"], d["span_q"],
                   "-" if d["qoq_pct"] is None else "%+.1f%%" % d["qoq_pct"],
                   "-" if d["yoy_pct"] is None else "%+.1f%%" % d["yoy_pct"]))
    note = "最新 %s → %s" % (_fmt(latest), latest["status"])
    if ev and ev["period_end"] != latest["period_end"]:
        note += " ／ 根拠期 %s → %s" % (_fmt(ev), ev["status"])
    if not latest["quarter_level"]:
        note += " ／ **Q単独では未確認（累計 span での判定）**"
    return {"status": latest["status"], "quarter_level": latest["quarter_level"],
            "latest": latest, "evidence": ev, "trend": trend(ts),
            "trend_text": trend_text(ts), "note": note}
