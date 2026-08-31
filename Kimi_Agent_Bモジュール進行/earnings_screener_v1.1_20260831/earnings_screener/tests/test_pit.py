"""
test_pit.py — PIT（Point-in-Time）機構の回帰テスト
検証項目:
  T1. 訂正世代の切替（as_of以前は訂正前の値、以後は訂正後の値が返る）
  T2. forecaster のPIT（証拠イベントがas_of以後のものを見ない）
  T3. 週次スコアリングのPIT（過去日指定で未来データが混入しない）
実行: python tests/test_pit.py（全モック生成後）
"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from module_b.data_access import connect, get_pl_series
from module_b.forecaster import forecast
from module_b.run_scorers import score_ticker, SCORERS_ALL


def main():
    conn = connect()
    ok = True

    # --- T1: 訂正世代（3905の3Q FY2026: 初回1400 / 訂正後1350） ---
    before = [r for r in get_pl_series(conn, "3905", date(2026, 3, 9))
              if r["period_end"] == "2025-12-31"]
    after = [r for r in get_pl_series(conn, "3905", date(2026, 3, 10))
             if r["period_end"] == "2025-12-31"]
    t1 = before and after and before[0]["sales"] == 1400 and after[0]["sales"] == 1350
    print(f"T1 訂正世代切替: 訂正前(as_of=3/9)={before[0]['sales'] if before else '?'} / "
          f"訂正後(as_of=3/10)={after[0]['sales'] if after else '?'} -> {'PASS' if t1 else 'FAIL'}")
    ok &= bool(t1)

    # --- T2: forecaster のPIT（278Aの装備庁受注 2026-05-20 の前後） ---
    fc_before = forecast(conn, "278A", date(2026, 5, 19))
    fc_after = forecast(conn, "278A", date(2026, 8, 31))
    t2 = (fc_before["available"] and fc_after["available"]
          and fc_before["path_C"]["n_events"] == 0
          and fc_after["path_C"]["evidence_extra_sales"] == 270.0)
    print(f"T2 証拠イベントのPIT: 受注前C経路={fc_before['path_C']['evidence_extra_sales']} "
          f"/ 受注後={fc_after['path_C']['evidence_extra_sales']} -> {'PASS' if t2 else 'FAIL'}")
    ok &= bool(t2)

    # --- T3: スコアリングのPIT（278Aの契約負債急増は1Q決算=2026-05-15発表で初めて公知） ---
    # 前日(5/14)時点では前受金180百万円は見えず、発表日(5/15)以降にのみ反映されるべき
    s_before = score_ticker(conn, "278A", SCORERS_ALL, as_of=date(2026, 5, 14))
    s_after = score_ticker(conn, "278A", SCORERS_ALL, as_of=date(2026, 5, 15))
    s2_b, s2_a = s_before["S2"]["score"], s_after["S2"]["score"]
    t3 = s2_b < s2_a  # 発表日を跨いでスコアが変化すること（=未来データを見ていない）
    print(f"T3 スコアのPIT: S2 発表前日={s2_b:+.3f} / 発表日={s2_a:+.3f} -> {'PASS' if t3 else 'FAIL'}")
    ok &= bool(t3)

    conn.close()
    print("\n=== " + ("ALL PASS" if ok else "FAIL あり") + " ===")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
