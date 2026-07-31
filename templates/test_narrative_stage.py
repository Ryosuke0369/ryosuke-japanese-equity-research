"""
test_narrative_stage.py — Validation of Block 5 logic against 4 reference cases

Reference cases (from docs/narrative_stage.md):
  - Core (2359, 2026-05): Stage 1
  - Denyo (6365, 2026-05): Stage 1 (axis 1=2 shows headroom but does NOT gate)
  - IHI (7013, 2026-05): Stage 3
  - Fujikura (5803, 2023-01, retroactive): Stage 1

Stage is gated by reach (axis 3), refined by realization (axis 2B).
Axis 1 (label rewrite) never moves the stage; it is reported as a headroom /
reversal-risk read instead.
"""

import sys

# This file prints em dashes and ✅/❌, neither of which exists in cp932 (the
# default console encoding on this machine) — without this the suite dies on
# its own banner before running a single case.
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:
        pass

from narrative_stage_template import NarrativeInput, assess


def _check(label: str, result, expected_stage, expected_cap=False,
           expected_verdict=None, expected_headroom=None):
    """Print test result and return True if passed."""
    passed = (result.stage == expected_stage
              and result.earnings_cap_applied == expected_cap)
    if expected_verdict is not None:
        passed = passed and result.final_verdict == expected_verdict
    if expected_headroom is not None:
        passed = passed and result.headroom_signal == expected_headroom

    status = "✅ PASS" if passed else "❌ FAIL"
    print(f"\n{status}: {label}")
    print(f"  Stage:          {result.stage} ({result.stage_label})")
    print(f"  Expected:       {expected_stage}")
    print(f"  Rate-limit:     {result.rate_limiting_axis}")
    print(f"  Reason:         {result.rate_limiting_reason}")
    print(f"  Total score:    {result.total_score}")
    print(f"  Earnings cap:   {result.earnings_cap_applied} "
          f"(expected {expected_cap})")
    print(f"  Headroom:       {result.headroom_signal}")
    if result.headroom_note:
        print(f"                  {result.headroom_note}")
    if expected_headroom is not None:
        print(f"  Headroom exp:   {expected_headroom}")
    if expected_verdict is not None:
        print(f"  Final verdict:  {result.final_verdict} "
              f"(expected {expected_verdict})")
    if result.earnings_impact_note:
        print(f"  Impact:         {result.earnings_impact_note}")
    return passed


def test_core():
    """Core (2359) — May 2026 — Stage 1 (reach=1, fired=1)."""
    inp = NarrativeInput(
        ticker='2359.T',
        company_name='コア',
        date='2026-05-19',
        axis_1_label_status=1,
        axis_2a_catalyst_potential=2,
        axis_2b_catalyst_realization=1,
        axis_3_diffusion_stage=1,
        axis_4_earnings_materiality=1,
        axis_5_narrative_durability=2,
        axis_1_note='防衛タグへの言及が個人投資家層で出始め',
        axis_2a_note='防衛省ドローン・スプーフィング実証実験を単独落札(2025/5)',
        axis_2b_note='続報待ち、6月カタリスト期待',
        axis_3_note='個人投資家の一部段階、機関・セルサイドはまだ',
        axis_4_note='現状寄与は中程度、継続案件化で拡大可能',
        axis_5_note='構造的な防衛予算増、競合限定',
        tam_oku_jpy=5000,
        expected_share_pct=10,
        segment_opm_pct=12,
        current_operating_profit_oku=80,
        fundamental_verdict='BUY',
    )
    result = assess(inp)
    # axis_3=1, axis_2b=1 -> Stage 1. axis_1=1 == axis_3=1 -> BALANCED.
    return _check('Core (2359) → Stage 1', result, expected_stage=1,
                  expected_verdict='STRONG_BUY',
                  expected_headroom='BALANCED')


def test_denyo():
    """Denyo (6365) — May 2026 — Stage 1.

    Reach (axis 3) = 1 and realization (axis 2B) = 1, so the grid puts Denyo
    at Stage 1 — the same diffusion段階 as Core. axis 1 = 2 (label well in
    place among retail) does NOT lift the stage: label rewrite and reach are
    independent. The axis1>axis3 gap surfaces instead as a HEADROOM signal.

    axis 4 = 0, but since stage_raw is already 1 the earnings cap does not
    fire (the cap only LOWERS a stage above 2; it never raises one).
    """
    inp = NarrativeInput(
        ticker='6365.T',
        company_name='電業社',
        date='2026-05-19',
        axis_1_label_status=2,
        axis_2a_catalyst_potential=2,
        axis_2b_catalyst_realization=1,
        axis_3_diffusion_stage=1,
        axis_4_earnings_materiality=0,
        axis_5_narrative_durability=1,
        axis_1_note='核融合関連として個人投資家層で定着(置換は進行)',
        axis_2a_note='ITER進捗、国内核融合スタートアップ連携可能性',
        axis_2b_note='大型実需契約はまだ限定的',
        axis_3_note='個人投資家中心、セルサイド/機関未到達(到達は狭い)',
        axis_4_note='核融合関連売上は全社売上の極小部分',
        axis_5_note='時間軸が長すぎ、テーマ賞味期限リスクあり',
        fundamental_verdict='CAUTION',
    )
    result = assess(inp)
    return _check('Denyo (6365) → Stage 1 (axis1=2 = headroom, not a gate)',
                  result, expected_stage=1, expected_cap=False,
                  expected_verdict='AVOID',
                  expected_headroom='HEADROOM')


def test_ihi():
    """IHI (7013) — May 2026 — Stage 3 (reach=2, fired=2)."""
    inp = NarrativeInput(
        ticker='7013.T',
        company_name='IHI',
        date='2026-05-19',
        axis_1_label_status=1,
        axis_2a_catalyst_potential=2,
        axis_2b_catalyst_realization=2,
        axis_3_diffusion_stage=2,
        axis_4_earnings_materiality=2,
        axis_5_narrative_durability=2,
        axis_1_note='新ラベル萌芽あるが「総合重工」が中心ラベル',
        axis_2a_note='中計2026-2034、フェーズ3 OPM 15%+ 明示',
        axis_2b_note='中計発表という強カタリストが発火済み',
        axis_3_note='個人〜メディア〜セルサイドまで拡散',
        axis_4_note='航空エンジン・防衛・原子力が既に主力',
        axis_5_note='地政学・脱炭素・電力需要の構造的需要',
        tam_oku_jpy=50000,
        expected_share_pct=15,
        segment_opm_pct=15,
        current_operating_profit_oku=1700,
        fundamental_verdict='BUY',
    )
    result = assess(inp)
    # axis_1=1 < axis_3=2 -> SHALLOW (reach ahead of label rewrite).
    return _check('IHI (7013) → Stage 3', result, expected_stage=3,
                  expected_verdict='HOLD',
                  expected_headroom='SHALLOW')


def test_fujikura_2023():
    """Fujikura (5803) — Jan 2023 retroactive — Stage 1 (reach=1, fired=1)."""
    inp = NarrativeInput(
        ticker='5803.T',
        company_name='フジクラ',
        date='2023-01-15',
        axis_1_label_status=1,
        axis_2a_catalyst_potential=2,
        axis_2b_catalyst_realization=1,
        axis_3_diffusion_stage=1,
        axis_4_earnings_materiality=2,
        axis_5_narrative_durability=2,
        axis_1_note='電線御三家、データセンター/AI関連は萌芽段階',
        axis_2a_note='SWR/WTC技術優位、北米ハイパースケーラー実供給、生成AI予兆',
        axis_2b_note='構造改革効果は発火、AI需要は予兆段階',
        axis_3_note='個人投資家の一部認識、セルサイド/AIタグ未到達',
        axis_4_note='情報通信セグメント既に主力(30%超)',
        axis_5_note='データセンター投資の構造的長期トレンド、技術優位',
        tam_oku_jpy=500000,
        expected_share_pct=10,
        segment_opm_pct=15,
        current_operating_profit_oku=380,
        fundamental_verdict='BUY',
    )
    result = assess(inp)
    # axis_1=1 == axis_3=1 -> BALANCED.
    return _check('Fujikura 2023 (retroactive) → Stage 1',
                  result, expected_stage=1,
                  expected_verdict='STRONG_BUY',
                  expected_headroom='BALANCED')


def test_edge_cases():
    """Edge cases: axis 4 cap, Stage 0, Stage 4."""
    print("\n--- Edge case tests ---")

    # Edge 1: pure narrative play (high stage scores, axis 4 = 0)
    # Should cap at Stage 2 (this is the genuine axis-4-cap demonstration).
    inp1 = NarrativeInput(
        ticker='TEST1', company_name='Pure Narrative', date='2026-01-01',
        axis_1_label_status=2,
        axis_2a_catalyst_potential=2,
        axis_2b_catalyst_realization=2,
        axis_3_diffusion_stage=2,  # would push to Stage 3
        axis_4_earnings_materiality=0,  # cap trigger
        axis_5_narrative_durability=2,
    )
    r1 = assess(inp1)
    p1 = _check('Edge: Pure narrative play (cap to Stage 2)',
                r1, expected_stage=2, expected_cap=True)

    # The cap overwrites rate_limiting_axis/reason with the axis-4 message, so
    # the pre-cap stage and its binding constraint must survive separately.
    p1_pre = (r1.stage_pre_cap == 3 and bool(r1.pre_cap_reason))
    print(f"\n{'✅ PASS' if p1_pre else '❌ FAIL'}: "
          f"Edge: pre-cap stage preserved")
    print(f"  stage_pre_cap:  {r1.stage_pre_cap} (expected 3)")
    print(f"  pre_cap_reason: {r1.pre_cap_reason or '(empty)'}")


    # Edge 2: Stage 0 — no spark, no fuel
    inp2 = NarrativeInput(
        ticker='TEST2', company_name='Sleeper', date='2026-01-01',
        axis_1_label_status=0,
        axis_2a_catalyst_potential=0,
        axis_2b_catalyst_realization=0,
        axis_3_diffusion_stage=0,
        axis_4_earnings_materiality=1,
        axis_5_narrative_durability=1,
    )
    r2 = assess(inp2)
    p2 = _check('Edge: No narrative movement (Stage 0)',
                r2, expected_stage=0)

    # Edge 3: Stage 4 — full rewrite
    inp3 = NarrativeInput(
        ticker='TEST3', company_name='Fully Rewritten', date='2026-01-01',
        axis_1_label_status=2,
        axis_2a_catalyst_potential=2,
        axis_2b_catalyst_realization=2,
        axis_3_diffusion_stage=2,
        axis_4_earnings_materiality=2,
        axis_5_narrative_durability=2,
    )
    r3 = assess(inp3)
    p3 = _check('Edge: Fully rewritten (Stage 4)',
                r3, expected_stage=4)

    # No cap fired here, so the pre-cap fields must stay empty rather than
    # reporting a cap that never happened.
    p3_pre = (r3.stage_pre_cap is None and r3.pre_cap_reason == '')
    print(f"\n{'✅ PASS' if p3_pre else '❌ FAIL'}: "
          f"Edge: no cap -> pre-cap fields stay empty")
    print(f"  stage_pre_cap:  {r3.stage_pre_cap} (expected None)")

    return p1 and p1_pre and p2 and p3 and p3_pre


def main():
    print("=" * 70)
    print("Block 5: Narrative Stage Assessment — Validation Tests")
    print("=" * 70)

    results = [
        test_core(),
        test_denyo(),
        test_ihi(),
        test_fujikura_2023(),
        test_edge_cases(),
    ]

    print("\n" + "=" * 70)
    passed = sum(results)
    total = len(results)
    print(f"Result: {passed}/{total} test groups passed")
    print("=" * 70)

    if passed != total:
        sys.exit(1)


if __name__ == '__main__':
    main()
