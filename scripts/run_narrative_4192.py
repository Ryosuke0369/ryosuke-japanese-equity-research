"""SpiderPlus (4192) — Narrative Stage assessment (Block 5).

Runs templates/narrative_stage_template.assess() for the 4192 six-axis scores.
Because 4192's fundamental verdict splits by valuation lens (Comps/Exit = cheap
=> BUY; reverse-DCF/PGM = rich => CAUTION), we run BOTH and contrast the final
verdict. Also reports a TAM x share sensitivity of the earnings-impact metric.

Output: console report + data/narrative/4192_narrative_20260603.json
"""
import sys
import os
import json
from dataclasses import asdict, replace

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates'))

from narrative_stage_template import NarrativeInput, assess

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

# ---------------------------------------------------------------------------
# 4192 six-axis input (scores grounded in this session's Comps / reverse-DCF /
# TAM work). fundamental_verdict is set per-run below.
# ---------------------------------------------------------------------------
BASE_INPUT = NarrativeInput(
    ticker="4192",
    company_name="SpiderPlus & Co.",
    date="2026-06-03",

    axis_1_label_status=1,          # spark: 建設DX/SaaSラベルは付与済、高収益SaaS化は途上
    axis_2a_catalyst_potential=2,   # fuel: 2024年問題・人手不足・ゼネコンDX + ARR構造(強力)
    axis_2b_catalyst_realization=1, # ignition: 1Q営業黒字転換は出たが決定的点火ではない
    axis_3_diffusion_stage=1,       # reach: 一部機関は保有も広く認識されていない(超小型)
    axis_4_earnings_materiality=2,  # impact: TAM中核、成熟時25-99億の営業益余地
    axis_5_narrative_durability=2,  # durability: 解約率0.9%・ARRストック・構造的トレンド

    axis_1_note="建設DX/SaaSラベルは既に付与。ただし黒字化初期で高収益SaaSへの書き換えは途上",
    axis_2a_note="建設2024年問題(時間外規制)・技能者不足・ゼネコンDX。ARR積み上げ構造が燃料",
    axis_2b_note="FY2026 1Qで営業黒字転換(+5百万)。ただし株価は底ばい(234-293円)で決定的点火に至らず",
    axis_3_note="BNYメロン7.6%等の海外機関は保有も、超小型でアナリストカバー薄く認識は限定的",
    axis_4_note="建設現場DX TAM1,250億(2030,矢野)〜建築ConTech 3,042億。シェア10-13%・OPM20-25%で"
                "成熟時25-99億の営業利益創出余地。現営業益ほぼゼロ・時価総額104億に対し十分大",
    axis_5_note="月次解約率0.9%(高スイッチングコスト)、ARRストック、建設DXは一過性でなく構造的トレンド",

    tam_oku_jpy=1250.0,             # 建設現場DX市場 2030年度(矢野経済研究所)
    expected_share_pct=13.0,        # 楽観シェア(アンドパッドに次ぐ2番手級)
    segment_opm_pct=22.0,           # 成熟SaaS営業利益率(eWeLL45%/スマレジ21%参考の中庸)
    current_operating_profit_oku=0.5,  # FY2026計画営業益≒0.5億(ほぼゼロ)

    fundamental_verdict=None,       # set per run
)


def format_result(res, fundamental_label):
    L = []
    L.append("=" * 72)
    L.append(f"  Narrative Stage — {res.company_name} ({res.ticker})  [{res.date}]")
    L.append(f"  Fundamental verdict (input) = {fundamental_label}")
    L.append("=" * 72)
    L.append("  6軸スコア:")
    axis_labels = {
        '1_label_status': '軸1 ラベル現状(火種)',
        '2a_catalyst_potential': '軸2A 触媒ポテンシャル(燃料)',
        '2b_catalyst_realization': '軸2B 触媒実現度(点火)',
        '3_diffusion_stage': '軸3 拡散進行度(到達)',
        '4_earnings_materiality': '軸4 業績寄与の現実性',
        '5_narrative_durability': '軸5 物語持続性',
    }
    for k, lbl in axis_labels.items():
        L.append(f"    {lbl:28s}= {res.axis_scores[k]}")
    L.append(f"    {'合計 total_score':28s}= {res.total_score} / 12")
    L.append("")
    L.append(f"  Stage            : {res.stage}  →  {res.stage_label}")
    L.append(f"  律速軸           : {res.rate_limiting_axis}")
    L.append(f"  理由             : {res.rate_limiting_reason}")
    L.append(f"  業績キャップ適用 : {res.earnings_cap_applied}")
    L.append("")
    L.append(f"  利益インパクト   : {res.earnings_impact_note}")
    if res.earnings_impact_oku is not None:
        L.append(f"    earnings_impact_oku   = {res.earnings_impact_oku:.1f} 億円")
    if res.earnings_uplift_ratio is not None:
        L.append(f"    earnings_uplift_ratio = {res.earnings_uplift_ratio:.1f} 倍")
    L.append("")
    L.append(f"  Headroom signal  : {res.headroom_signal}")
    L.append(f"    {res.headroom_note}")
    L.append("")
    L.append(f"  ★ Final Verdict  : {res.final_verdict}")
    L.append(f"    {res.verdict_rationale}")
    L.append("=" * 72)
    return "\n".join(L)


def tam_sensitivity():
    """TAM x share sensitivity of earnings_impact_oku = TAM * share% * OPM%."""
    opm = BASE_INPUT.segment_opm_pct / 100.0
    cur_op = BASE_INPUT.current_operating_profit_oku
    rows = []
    print("\n" + "=" * 72)
    print("  TAM 感度分析: earnings_impact_oku = TAM × share × OPM(22%)")
    print("=" * 72)
    print(f"  {'TAM(億)':>10s} | {'share':>6s} | {'impact(億)':>11s} | {'uplift vs 現営利0.5億':>18s}")
    print("  " + "-" * 60)
    for tam in (1250.0, 3042.0):
        for share in (10.0, 13.0):
            impact = tam * (share / 100.0) * opm
            uplift = impact / cur_op if cur_op else None
            tam_label = "現場DX" if tam == 1250.0 else "建築ConTech"
            print(f"  {tam:>10,.0f} | {share:>5.0f}% | {impact:>11.1f} | "
                  f"{uplift:>14.0f}倍   ({tam_label})")
            rows.append({'tam_oku': tam, 'share_pct': share,
                         'opm_pct': BASE_INPUT.segment_opm_pct,
                         'impact_oku': round(impact, 2),
                         'uplift_ratio': round(uplift, 1) if uplift else None})
    print("=" * 72)
    return rows


def main():
    runs = {}
    for fv in ("BUY", "CAUTION"):
        res = assess(replace(BASE_INPUT, fundamental_verdict=fv))
        print(format_result(res, fv))
        print()
        runs[fv] = res

    sens = tam_sensitivity()

    # ── contrast summary ──
    print("\n" + "#" * 72)
    print("  総括: ナラティブは本命ゾーン、ファンダはモノサシ次第で両極")
    print("#" * 72)
    print(f"  Stage = {runs['BUY'].stage} ({runs['BUY'].stage_label}) / "
          f"律速軸 = {runs['BUY'].rate_limiting_axis} / Headroom = {runs['BUY'].headroom_signal}")
    print(f"  BUY (Comps/Exit=割安)     × Stage{runs['BUY'].stage} → "
          f"{runs['BUY'].final_verdict}")
    print(f"  CAUTION (逆算DCF/PGM=割高) × Stage{runs['CAUTION'].stage} → "
          f"{runs['CAUTION'].final_verdict}")
    print("#" * 72)

    # ── persist JSON ──
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'narrative')
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, '4192_narrative_20260603.json')
    payload = {
        'input': {k: v for k, v in asdict(BASE_INPUT).items()
                  if k != 'fundamental_verdict'},
        'results': {fv: asdict(runs[fv]) for fv in runs},
        'tam_sensitivity': sens,
    }
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"\nSaved: {os.path.normpath(out_path)}")


if __name__ == "__main__":
    main()
