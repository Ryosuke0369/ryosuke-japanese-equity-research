"""screener/signals/signal_defs.py — 指標層 S1〜S5 (仕様書 §4)。

各シグナルは Signal(値, 発火bool, 根拠文字列) を返す。閾値は
config/signal_thresholds.yaml。**根拠文字列には必ず生の数字を入れる** ——
「発火した」だけでは後から検証できず、週次レポートで人が判断できない。

読むのは financials_q(単独値・有効行のみ)。累計を直接読んではいけない。
累計は平均で嘘をつく、というのがこのシステムの前提そのものである。

無効行(valid_flag=0)は計算に使わない。使うと決算期変更や遡及修正を
「傾き」として検出する。
"""
from __future__ import annotations

import os
import sys
from dataclasses import dataclass

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.projection import pit

SIGNAL_IDS = ("S1", "S2", "S3", "S3b", "S4", "S5")


@dataclass
class Signal:
    signal_id: str
    value: float | None
    fired: bool
    evidence: str
    direction: int = 1          # +1 = 加点 / -1 = 減点(逆シグナル)
    # 別枠 earnings_screener との互換用。-1..+1 に正規化した無次元の強度。
    # value を流用しないのは、value の単位が指標ごとに違う(S1はpt、S2は%、
    # S5はpt)ため。そのまま平均すると単位の違う数を足すことになる。
    # 正規化の分母は signal_thresholds.yaml の発火閾値から導くので、
    # 魔法の数字を新しく増やさない。docs/adapter_design.md C-4。
    strength: float | None = None


def _strength(value: float | None, fire_pt: float) -> float | None:
    """発火閾値の2倍で飽和する線形正規化。閾値ちょうどで ±0.5。"""
    if value is None or not fire_pt:
        return None
    return max(-1.0, min(value / (2.0 * abs(fire_pt)), 1.0))


def load_thresholds() -> dict:
    return C.load_yaml("signal_thresholds.yaml")


# ------------------------------------------------------------------ 系列取得
def quarter_key(period: str, q_no: int) -> tuple[int, int]:
    """(FY2026, 3) → (2026, 3)。期をまたいで時系列に並べるためのキー。"""
    y = int(period[2:]) if period.startswith("FY") and period[2:].isdigit() else 0
    return (y, q_no or 0)


class Series:
    """1銘柄ぶんの単独値。`(period, q_no)` の時系列として引ける。

    有効行だけを載せる。無効行を混ぜると、決算期変更や遡及修正を傾きとして
    読んでしまう —— それを防ぐために valid_flag がある。
    """

    def __init__(self, con, code: str, as_of=None, mode: str = "strict",
                 allow_span=(1, 2)):
        """DBを直接叩かず投影層(projection.pit)を経由する。

        経由を強制する理由は2つ。(1) as_of を渡し忘れた瞬間に未来を見る。
        (2) 一時収入の調整が効かなくなる —— 3905 の FY2027Q1 は調整前だと
        粗利率90.9%で S1 が +63.3pt の買いシグナルを誤発火するが、
        一時収入を除くと粗利は前年割れ。tests/test_projection.py の P4 が
        この経由を機械的に強制している。
        """
        self.code = code
        self.data: dict[str, dict[tuple[int, int], float]] = {}
        # span は (item, 期) 単位で持つ。期だけをキーにすると、同じ四半期の
        # ストック項目(span=1)がフロー項目(span=2)を上書きし、6ヶ月の値を
        # 「単独」と表示してしまう。
        self.span: dict[tuple[str, tuple[int, int]], int] = {}
        self.adjusted: set = set()
        self._con = con
        self._as_of = as_of
        self._mode = mode
        rows = pit.visible_q(con, code, as_of, mode, allow_span=allow_span)
        for r in rows:
            k = quarter_key(r["period"], r["q_no"])
            self.data.setdefault(r["item"], {})[k] = r["value"]
            self.span[(r["item"], k)] = r["span_q"] or 1

        # 一時収入の控除。売上と、既定仮定(OPに全額フロー)で営業利益にも効かせる。
        for (period, q_no), d in pit.visible_adjustments(
                con, code, as_of, mode).items():
            k = quarter_key(period, q_no)
            one = d.get("one_time_revenue", 0.0)
            if not one:
                continue
            for item in ("revenue", "gross_profit", "operating_income"):
                if k in self.data.get(item, {}):
                    self.data[item][k] -= one
                    self.adjusted.add(k)

    def span_label(self, item: str, key) -> str:
        """その値が何ヶ月ぶんかを言う。span=2 を『Q単独』と呼ぶと、3ヶ月と
        6ヶ月が同じ言葉になって根拠文が嘘になる。"""
        n = self.span.get((item, key), 1)
        return "単独" if n == 1 else f"{n * 3}ヶ月単独"

    def keys(self, item: str) -> list[tuple[int, int]]:
        return sorted(self.data.get(item, {}))

    def get(self, item: str, key) -> float | None:
        return self.data.get(item, {}).get(key)

    def latest(self, item: str):
        ks = self.keys(item)
        return ks[-1] if ks else None

    @staticmethod
    def prev_q(key):
        y, q = key
        return (y - 1, 4) if q == 1 else (y, q - 1)

    @staticmethod
    def prev_y(key):
        y, q = key
        return (y - 1, q)

    def common_latest(self, items) -> tuple | None:
        """指定した全項目が揃っている最新の四半期。片方だけある期で比率を
        作ると分母と分子の期がずれるので、必ず揃っている期を使う。"""
        sets = [set(self.keys(i)) for i in items]
        if not sets or any(not s for s in sets):
            return None
        common = set.intersection(*sets)
        return max(common) if common else None


def _pct_change(now: float | None, before: float | None) -> float | None:
    if now is None or before is None or before == 0:
        return None
    return (now - before) / abs(before) * 100.0


# ----------------------------------------------------------------------- S1
def s1_gross_margin_slope(s: Series, th: dict) -> Signal:
    """Q単独粗利率の傾き。3441型の主砲。

    粗利率は gross_profit / revenue。gross_profit が無い会社は
    revenue - cogs で補う（短信の様式によっては粗利を直接持たない）。
    """
    cfg = th["s1_gross_margin"]

    def gm(key):
        rev = s.get("revenue", key)
        if not rev:
            return None
        gp = s.get("gross_profit", key)
        if gp is None:
            cogs = s.get("cogs", key)
            gp = None if cogs is None else rev - cogs
        return None if gp is None else gp / rev * 100.0

    keys = [k for k in s.keys("revenue") if gm(k) is not None]
    if not keys:
        return Signal("S1", None, False, "粗利率を作れる四半期が無い")
    k = keys[-1]
    cur = gm(k)
    qoq = None if gm(s.prev_q(k)) is None else cur - gm(s.prev_q(k))
    yoy = None if gm(s.prev_y(k)) is None else cur - gm(s.prev_y(k))
    moves = [x for x in (qoq, yoy) if x is not None]
    if not moves:
        return Signal("S1", cur, False,
                      f"FY{k[0]}Q{k[1]} {s.span_label('revenue', k)}粗利率 {cur:.1f}% "
                      f"(比較対象の四半期が無い)")
    best = max(moves, key=abs)
    parts = [f"FY{k[0]}Q{k[1]} {s.span_label('revenue', k)}粗利率 {cur:.1f}%"
             + ("(一時収入控除後)" if k in s.adjusted else "")]
    if qoq is not None:
        parts.append(f"前Q比 {qoq:+.1f}pt")
    if yoy is not None:
        parts.append(f"前年同Q比 {yoy:+.1f}pt")
    if best >= cfg["fire_pt"]:
        return Signal("S1", best, True, " / ".join(parts), 1,
                      _strength(best, cfg["fire_pt"]))
    if best <= cfg["penalty_pt"]:
        return Signal("S1", best, True, " / ".join(parts) + " ← 低下(逆シグナル)",
                      -1, _strength(best, cfg["fire_pt"]))
    return Signal("S1", best, False, " / ".join(parts), 1,
                  _strength(best, cfg["fire_pt"]))


# ----------------------------------------------------------------------- S2
def s2_inventory_split(s: Series, th: dict) -> Signal:
    """在庫の価格/数量分解。**在庫増は両義的**。

    数量要因の増加 × 売上加速 = 仕込み(加点)
    数量要因の増加 × 売上減速 = 滞留(減点)
    「増えた」だけでは方向が決まらないので、必ず売上の加速/減速と掛ける。
    """
    cfg = th["s2_inventory"]
    k = s.common_latest(["inventories_total", "revenue", "cogs"])
    if k is None:
        return Signal("S2", None, False, "棚卸資産・売上・売上原価が揃う四半期が無い")
    kp = s.prev_q(k)
    d_inv = _pct_change(s.get("inventories_total", k), s.get("inventories_total", kp))
    d_cogs = _pct_change(s.get("cogs", k), s.get("cogs", kp))
    g_now = _pct_change(s.get("revenue", k), s.get("revenue", s.prev_y(k)))
    g_prev = _pct_change(s.get("revenue", kp), s.get("revenue", s.prev_y(kp)))
    if d_inv is None:
        return Signal("S2", None, False, "前四半期の棚卸資産が無い")
    if d_inv < cfg["inventory_change_pct"]:
        return Signal("S2", d_inv, False,
                      f"FY{k[0]}Q{k[1]} 棚卸 前Q比 {d_inv:+.1f}% (閾値未満)")

    # 価格プロキシ = ΔCOGS%。棚卸の増加が ΔCOGS% で説明できる範囲なら価格要因、
    # 大きく上回るなら数量要因（積み増し）。
    gap = None if d_cogs is None else d_inv - d_cogs
    quantity_driven = gap is not None and gap >= cfg["price_proxy_gap_pt"]
    accel = (g_now is not None and g_prev is not None
             and (g_now - g_prev) > cfg["revenue_accel_pt"])
    ev = [f"FY{k[0]}Q{k[1]} 棚卸 前Q比 {d_inv:+.1f}%"]
    if d_cogs is not None:
        ev.append(f"COGS 前Q比 {d_cogs:+.1f}% (価格プロキシ)")
    if gap is not None:
        ev.append(f"差 {gap:+.1f}pt → {'数量要因' if quantity_driven else '価格要因寄り'}")
    if g_now is not None and g_prev is not None:
        ev.append(f"売上YoY {g_prev:+.1f}% → {g_now:+.1f}% ({'加速' if accel else '減速'})")
    if not quantity_driven:
        return Signal("S2", d_inv, False, " / ".join(ev) + " ← 価格要因は発火させない")
    if accel:
        return Signal("S2", d_inv, True, " / ".join(ev) + " ← 仕込み", 1,
                      _strength(d_inv, cfg["inventory_change_pct"]))
    return Signal("S2", d_inv, True, " / ".join(ev) + " ← 滞留(逆シグナル)", -1,
                  _strength(-d_inv, cfg["inventory_change_pct"]))


# ----------------------------------------------------------------------- S3
def s3_construction_in_progress(s: Series, th: dict) -> tuple[Signal, Signal]:
    """建設仮勘定。急増=稼働前投資(S3)、建仮減×機械装置増=振替検出(S3b)。

    振替は「稼働開始の証拠」。3905のQ2チェックと同型。
    """
    cfg = th["s3_cip"]
    k = s.latest("construction_in_progress")
    if k is None:
        return (Signal("S3", None, False, "建設仮勘定のデータが無い"),
                Signal("S3b", None, False, "建設仮勘定のデータが無い"))
    kp = s.prev_q(k)
    d_cip = _pct_change(s.get("construction_in_progress", k),
                        s.get("construction_in_progress", kp))
    d_mac = _pct_change(s.get("machinery_and_equipment", k),
                        s.get("machinery_and_equipment", kp))
    if d_cip is None:
        return (Signal("S3", None, False, "前四半期の建設仮勘定が無い"),
                Signal("S3b", None, False, "前四半期の建設仮勘定が無い"))

    base = f"FY{k[0]}Q{k[1]} 建仮 前Q比 {d_cip:+.1f}%"
    s3 = (Signal("S3", d_cip, True, base + " ← 稼働前投資", 1,
                      _strength(d_cip, cfg["surge_pct"]))
          if d_cip >= cfg["surge_pct"]
          else Signal("S3", d_cip, False, base + " (閾値未満)"))

    if d_mac is None:
        s3b = Signal("S3b", None, False, base + " / 機械装置のデータが無い")
    elif (d_cip <= cfg["transfer_cip_drop_pct"]
          and d_mac >= cfg["transfer_machinery_rise_pct"]):
        s3b = Signal("S3b", d_mac, True,
                     base + f" / 機械装置 {d_mac:+.1f}% ← 振替(稼働開始の証拠)", 1)
    else:
        s3b = Signal("S3b", d_mac, False,
                     base + f" / 機械装置 {d_mac:+.1f}% (振替の形ではない)")
    return s3, s3b


# ----------------------------------------------------------------------- S4
def s4_contract_liabilities(s: Series, th: dict) -> Signal:
    """契約負債・前受金。売上成長率を上回る増加で発火(顧客コミットの現金証拠)。"""
    cfg = th["s4_contract_liabilities"]
    item = ("contract_liabilities" if s.keys("contract_liabilities")
            else "advances_received")
    k = s.common_latest([item, "revenue"])
    if k is None:
        return Signal("S4", None, False, "契約負債・前受金のデータが無い")
    kp = s.prev_q(k)
    d_cl = _pct_change(s.get(item, k), s.get(item, kp))
    d_rev = _pct_change(s.get("revenue", k), s.get("revenue", kp))
    if d_cl is None:
        return Signal("S4", None, False, "前四半期の契約負債・前受金が無い")
    ev = f"FY{k[0]}Q{k[1]} {item} 前Q比 {d_cl:+.1f}%"
    if d_rev is not None:
        ev += f" / 売上 前Q比 {d_rev:+.1f}%"
    if d_cl < cfg["min_change_pct"]:
        return Signal("S4", d_cl, False, ev + " (閾値未満)")
    if d_rev is None:
        return Signal("S4", d_cl, False, ev + " (売上と比較できない)")
    excess = d_cl - d_rev
    if excess >= cfg["excess_over_revenue_pt"]:
        return Signal("S4", excess, True,
                      ev + f" / 超過 {excess:+.1f}pt ← 顧客コミット", 1,
                      _strength(excess, cfg["excess_over_revenue_pt"]))
    return Signal("S4", excess, False, ev + f" / 超過 {excess:+.1f}pt (閾値未満)")


# ----------------------------------------------------------------------- S5
def s5_turnaround_and_progress(con, s: Series, th: dict) -> Signal:
    """①営業損益の赤字→黒字転換(Q単独) ②累計進捗率が通期予想を大幅超過。

    ②は3441の進捗113%型。Q3時点で四半期数比75%に対し113%なら「死んだ
    ガイダンス」——上方修正を先回りする。
    """
    cfg = th["s5_turnaround"]
    k = s.latest("operating_income")
    parts, fired, value = [], False, None

    if k is not None:
        cur = s.get("operating_income", k)
        prev = s.get("operating_income", s.prev_q(k))
        if cur is not None and prev is not None:
            parts.append(f"FY{k[0]}Q{k[1]} 単独営業損益 {prev:,.0f} → {cur:,.0f}")
            if prev < 0 <= cur:
                fired, value = True, cur - prev
                parts[-1] += " ← 赤字→黒字転換"

    # 累計進捗率。分子は単独値の足し上げ（累計を直接読まない方針の一貫性）。
    if k is not None:
        ytd = sum(v for kk, v in s.data.get("operating_income", {}).items()
                  if kk[0] == k[0] and kk[1] <= k[1])
        # guidance も投影層経由。ここを直に読むと as_of が効かず、
        # 「決算後に出た修正予想」で決算前の進捗率を評価してしまう。
        g = pit.visible_guidance(con, s.code, "operating_income",
                                 s._as_of, s._mode)
        if g and g["value"]:
            progress = ytd / g["value"] * 100.0
            pace = k[1] / 4 * 100.0
            excess = progress - pace
            parts.append(f"累計進捗 {progress:.0f}% (四半期数比 {pace:.0f}%, "
                         f"超過 {excess:+.0f}pt)")
            if excess >= cfg["progress_excess_pt"]:
                fired = True
                value = excess if value is None else value
                parts[-1] += " ← 死んだガイダンス"
    if not parts:
        return Signal("S5", None, False, "営業損益の単独値が無い")
    return Signal("S5", value, fired, " / ".join(parts), 1,
                  _strength(value, cfg["progress_excess_pt"]))


# --------------------------------------------------------------------- 実行
def evaluate(con, code: str, th: dict | None = None, as_of=None,
             mode: str = "strict") -> list[Signal]:
    th = th or load_thresholds()
    s = Series(con, code, as_of, mode)
    s3, s3b = s3_construction_in_progress(s, th)
    return [s1_gross_margin_slope(s, th), s2_inventory_split(s, th), s3, s3b,
            s4_contract_liabilities(s, th), s5_turnaround_and_progress(con, s, th)]


def store(con, code: str, eval_date: str, sigs: list[Signal]) -> None:
    for sg in sigs:
        con.execute(
            "INSERT OR REPLACE INTO signals (code, eval_date, signal_id, value, "
            " fired, evidence_text) VALUES (?,?,?,?,?,?)",
            (code, eval_date, sg.signal_id, sg.value, 1 if sg.fired else 0,
             sg.evidence))
    con.commit()
