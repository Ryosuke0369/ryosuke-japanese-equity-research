"""check_forecast.py — 機械的予測（A+B+C）のテスト銘柄検証"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from module_b.data_access import connect
from module_b.forecaster import forecast

conn = connect()
for t in ["3905", "3441", "278A", "5726", "6501"]:
    fc = forecast(conn, t, date(2026, 8, 31))
    print(f"=== {t} ===")
    if not fc["available"]:
        print("  予測不能:", fc["reason"])
        continue
    print(f"  次Q予測: 売上 {fc['pred_sales']:,.0f} / 営業利益 {fc['pred_op']:,.0f} / 定常OP推定 {fc['steady_op_est']:,.0f}")
    a, b, c = fc["path_A"], fc["path_B"], fc["path_C"]
    print(f"  A(トレンド外挿): 売上 {a['sales']:,.0f}（YoY中央値 {a['yoy_median']:+.1%}）")
    print(f"  B(BS→PL復元)  : 売上 {b['sales']:,.0f}（契約負債増分 {b['d_contract_liab']:,.0f} / 振替寄与 {b['transfer_contrib']:,.0f}）")
    print(f"  C(証拠積上げ) : +{c['evidence_extra_sales']:,.0f}（{c['n_events']}件。MOU等の約束は除外）")
conn.close()
