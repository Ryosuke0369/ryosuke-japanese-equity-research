"""screener/projection/adapter.py — 別枠 earnings_screener への投影層。

統合引継ぎ書 v1.1 §3-1 の4関数と同じシグネチャを本体側で実装する。
**別枠のコードは1行も変更しない。** import 先を data_access からここへ
差し替えるだけで繋がる形にしてある。理由:

  - 別枠は tests/test_pit.py で不変条件が固定されており、書き換えると
    その保証を失う
  - v1.1 の generation 機構は「モック側でのPIT検証用」と v1.1 §0-1 自身が
    位置づけている。本体は financials_cum(filing_id, item, context_ref) で
    同等以上を自然に表現できるので、本体側で実装し直す方が正しい

## 返さないものは黙って返さない

写像が unverified / missing の item_key を要求されたら、空を返したうえで
**なぜ返せないかをログに出す**。0 や None を黙って返すと、下流は
「データが無い」と「まだ繋いでいない」を区別できなくなる —— 本体が
一貫して避けてきた失敗そのもの。
"""
from __future__ import annotations

import os
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.projection import pit

_CFG = None
_WARNED: set = set()


def cfg() -> dict:
    global _CFG
    if _CFG is None:
        _CFG = C.load_yaml("earnings_screener_mapping.yaml")
    return _CFG


def resolve_item(item_key: str) -> str | None:
    """別枠 item_key → 本体 item。繋げないものは None を返しログに残す。"""
    e = (cfg().get("items") or {}).get(item_key)
    if e is None:
        _warn(item_key, f"写像表に item_key '{item_key}' が無い")
        return None
    if e.get("status") != "verified":
        _warn(item_key,
              f"item_key '{item_key}' は status={e.get('status')} のため未接続 "
              f"({str(e.get('note') or '').strip()[:80]})")
        return None
    return e.get("item")


def _warn(key: str, msg: str) -> None:
    if key not in _WARNED:            # 同じ警告で画面を埋めない
        _WARNED.add(key)
        C.log(f"[adapter] {msg}")


def _div() -> float:
    return float(cfg().get("unit_divisor") or 1.0)


def _qtype(q_no: int) -> str | None:
    return (cfg().get("quarter_type") or {}).get(q_no)


def _fy(period: str) -> int | None:
    if period and period.startswith("FY") and period[2:].isdigit():
        return int(period[2:])
    return None


def _period_end(period: str, q_no: int) -> str:
    """別枠は period_end を系列のキーに使う。本体は (period, q_no) を持つので
    ソート可能な合成キーを作る。実日付ではないので日付として解釈しないこと。"""
    return f"{period}-Q{q_no}"


# ------------------------------------------------------- §3-1 の4関数
def get_pl_series(con, ticker, as_of=None, *, mode: str = "strict",
                  allow_span=None) -> list[dict]:
    """単独値PL時系列。有効行のみ・古い順。

    既定は span_q=1（真の四半期単独値）のみ。span_q=2（半期粒度）は別枠に
    対応概念が無く、渡すと6ヶ月の値を四半期として扱う。戻り値には span_q と
    is_half を必ず含め、受け手が無視できないようにする。
    """
    allow = tuple(allow_span or cfg().get("allowed_span_q") or (1,))
    rows = pit.visible_q(con, ticker, as_of, mode, allow_span=allow)
    by_pq: dict = {}
    for r in rows:
        by_pq.setdefault((r["period"], r["q_no"]), {})[r["item"]] = r
    div = _div()

    out = []
    for (period, q_no), items in sorted(by_pq.items()):
        qt = _qtype(q_no)
        if qt is None:
            continue
        span = next((it["span_q"] or 1 for it in items.values()), 1)
        def v(name):
            r = items.get(name)
            return None if r is None or r["value"] is None else r["value"] / div
        out.append({
            "period_end": _period_end(period, q_no),
            "quarter_type": qt,
            "fiscal_year": _fy(period),
            "sales": v("revenue"),
            "operating_profit": v("operating_income"),
            "gross_profit": v("gross_profit"),
            # 原価の数量/価格分解は本体にも XBRL にも該当科目が無い。
            # 別枠 S6 の入力だが作れないので None を明示する（0 にしない）。
            "cogs_quantity": None,
            "cogs_price": None,
            "operating_cf": v("operating_cf"),
            "span_q": span,
            "is_half": span == 2,
            "period": period,
            "q_no": q_no,
        })
    return out


def get_bs_series(con, ticker, item_key, as_of=None, *,
                  mode: str = "strict") -> list[tuple]:
    """BS科目のスナップショット。[(period_end, value)] を古い順で返す。"""
    item = resolve_item(item_key)
    if item is None:
        return []
    rows = pit.visible_q(con, ticker, as_of, mode, allow_span=(1, 2, 3, 4))
    div = _div()
    out = [( _period_end(r["period"], r["q_no"]), r["value"] / div)
           for r in rows if r["item"] == item and r["value"] is not None]
    return sorted(out)


def get_adjustments(con, ticker, as_of=None, *, mode: str = "strict") -> dict:
    """一時収入等の調整。period_end → 売上から控除する合計額。

    one_time_revenue のみを売上控除に使う。one_time_cost / one_time_gain は
    利益側の調整なので別関数（get_adjustment_detail）で返す。
    """
    raw = pit.visible_adjustments(con, ticker, as_of, mode)
    div = _div()
    return {_period_end(p, q): sum(v for k, v in d.items()
                                   if k == "one_time_revenue") / div
            for (p, q), d in raw.items()
            if any(k == "one_time_revenue" for k in d)}


def get_adjustment_detail(con, ticker, as_of=None, *,
                          mode: str = "strict") -> dict:
    """period_end → {item_key: 金額}。根拠つきで何を引いたかを見せるため。"""
    raw = pit.visible_adjustments(con, ticker, as_of, mode)
    div = _div()
    return {_period_end(p, q): {k: v / div for k, v in d.items()}
            for (p, q), d in raw.items()}


def get_latest_forecast(con, ticker, as_of=None, *, mode: str = "strict"):
    """会社通期予想の最新版（as_of 指定でその時点の版）。v1.1 §3-1。"""
    row = pit.visible_guidance(con, ticker, "operating_income", as_of, mode)
    if row is None:
        return None
    sales = pit.visible_guidance(con, ticker, "revenue", as_of, mode)
    div = _div()
    return {"fiscal_year": _fy(row["fy"]),
            "forecast_op": (row["value"] or 0) / div,
            "forecast_sales": ((sales["value"] or 0) / div) if sales else None,
            "source_date": row["date"]}


# ------------------------------------------------------------- 補助
def normalized_sales(pl_row, adjustments) -> float:
    """調整後売上。別枠 data_access.normalized_sales と同じ意味論。"""
    return (pl_row["sales"] or 0) - adjustments.get(pl_row["period_end"], 0)


def unmapped_keys() -> dict:
    """繋がっていない item_key の一覧。統合状況の可視化用。"""
    out = {"unverified": [], "missing": []}
    for k, e in (cfg().get("items") or {}).items():
        st = e.get("status")
        if st in out:
            out[st].append(k)
    return out
