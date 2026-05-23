"""
narrative_stage_template.py — Block 5: Narrative Stage Assessment

Equity Research Framework — Gap 2 (Catalyst/Narrative Gap) Analyzer

This module implements the Narrative Stage Assessment as a companion to the
existing Market Scorecard (Blocks 1-4). It scores a stock across 6 axes,
determines its narrative stage (0-4), and supports the dual-axis Verdict
matrix that combines fundamental scoring with narrative stage.

Design reference: docs/narrative_stage.md

Axis roles (finalized 2026-05-22):
    Axis 1  = is there a spark?   (label rewrite progress; NOT a stage gate)
    Axis 2A = is there fuel?      (existence of catalyst sources)
    Axis 2B = has it ignited?     (catalyst fired into the market)
    Axis 3  = how wide is fire?   (diffusion / reach of recognition) <- stage gate
    Axis 4  = earnings materiality (caps stage at 2 if == 0)
    Axis 5  = narrative durability

Stage is gated by REACH (axis 3), refined by realization (axis 2B).
Axis 1 (label rewrite progress) is independent from axis 3 (reach) and does
NOT move the stage. It is reported separately as a read on headroom / reversal
risk: axis1 > axis3 => room to spread; axis1 < axis3 => shallow rewrite risk.
"""

from dataclasses import dataclass, field
from typing import Optional, Literal

# ============================================================================
# Type definitions
# ============================================================================

AxisScore = Literal[0, 1, 2]
Stage = Literal[0, 1, 2, 3, 4]
BlockScore = Literal['BUY', 'HOLD', 'CAUTION']
Verdict = Literal['STRONG_BUY', 'BUY', 'HOLD', 'OBSERVE', 'TRIM', 'AVOID',
                  'CAUTION', 'EXIT']

# ============================================================================
# Input data structure
# ============================================================================

@dataclass
class NarrativeInput:
    """Block 5 input: six-axis scores + supporting context."""

    # Meta
    ticker: str
    company_name: str
    date: str  # ISO format YYYY-MM-DD

    # Six axis scores (0/+1/+2)
    axis_1_label_status: AxisScore
    axis_2a_catalyst_potential: AxisScore
    axis_2b_catalyst_realization: AxisScore
    axis_3_diffusion_stage: AxisScore
    axis_4_earnings_materiality: AxisScore
    axis_5_narrative_durability: AxisScore

    # Free-form rationale for each axis (audit trail)
    axis_1_note: str = ""
    axis_2a_note: str = ""
    axis_2b_note: str = ""
    axis_3_note: str = ""
    axis_4_note: str = ""
    axis_5_note: str = ""

    # Companion metric: earnings impact at scale (units: 億円 / 100M JPY)
    tam_oku_jpy: Optional[float] = None  # Total Addressable Market
    expected_share_pct: Optional[float] = None  # Optimistic share, e.g. 10.0
    segment_opm_pct: Optional[float] = None  # New segment operating margin
    current_operating_profit_oku: Optional[float] = None  # Current OP

    # Block 2-4 fundamental verdict (for matrix lookup)
    fundamental_verdict: Optional[BlockScore] = None

    def __post_init__(self):
        """Validate inputs."""
        for name, val in [
            ('axis_1_label_status', self.axis_1_label_status),
            ('axis_2a_catalyst_potential', self.axis_2a_catalyst_potential),
            ('axis_2b_catalyst_realization', self.axis_2b_catalyst_realization),
            ('axis_3_diffusion_stage', self.axis_3_diffusion_stage),
            ('axis_4_earnings_materiality', self.axis_4_earnings_materiality),
            ('axis_5_narrative_durability', self.axis_5_narrative_durability),
        ]:
            if val not in (0, 1, 2):
                raise ValueError(
                    f"{name} must be 0, 1, or 2 (got {val})"
                )


# ============================================================================
# Output data structure
# ============================================================================

@dataclass
class NarrativeResult:
    """Block 5 output: stage judgment + companion analysis."""

    ticker: str
    company_name: str
    date: str

    # Stage outputs
    stage: Stage
    stage_label: str  # "Stage 1 (萌芽期)" etc
    rate_limiting_axis: str  # which axis gates the stage
    rate_limiting_reason: str

    # Sum and applied rules (for audit)
    total_score: int
    earnings_cap_applied: bool  # True if axis 4 == 0 triggered cap

    # Companion metric
    earnings_impact_oku: Optional[float] = None  # TAM * share * OPM
    earnings_uplift_ratio: Optional[float] = None  # impact / current OP
    earnings_impact_note: str = ""

    # Companion read: axis 1 (label rewrite) vs axis 3 (reach)
    # 'HEADROOM'  : axis1 > axis3 (rewrite ahead of reach -> room to spread)
    # 'BALANCED'  : axis1 == axis3
    # 'SHALLOW'   : axis1 < axis3 (reach ahead of rewrite -> reversal risk)
    headroom_signal: str = ""
    headroom_note: str = ""

    # Verdict matrix (if fundamental_verdict provided)
    final_verdict: Optional[Verdict] = None
    verdict_rationale: str = ""

    # Echo back for output
    axis_scores: dict = field(default_factory=dict)
    axis_notes: dict = field(default_factory=dict)


# ============================================================================
# Stage determination logic
# ============================================================================

STAGE_LABELS = {
    0: "Stage 0 (ラベル変化なし)",
    1: "Stage 1 (萌芽期)",
    2: "Stage 2 (普及期)",
    3: "Stage 3 (承認期)",
    4: "Stage 4 (定着期)",
}


def _determine_stage(inp: NarrativeInput) -> tuple[Stage, str, str, bool]:
    """
    Determine narrative stage. REACH (axis 3) is the primary gate; catalyst
    realization (axis 2B) refines it. Axis 1 does NOT move the stage.

    Returns:
        (stage, rate_limiting_axis_name, reason, earnings_cap_applied)

    Stage determination grid:
      axis_3=0, axis_2b<=1  -> Stage 1
      axis_3=0, axis_2b=2   -> Stage 1 (catalyst fired but no diffusion yet)
      axis_3=1, axis_2b<=1  -> Stage 1 (late)
      axis_3=1, axis_2b=2   -> Stage 2 (catalyst fired + retail noticed)
      axis_3=2, axis_2b<=1  -> Stage 2 (sell-side reached but unfired)
      axis_3=2, axis_2b=2   -> Stage 3 (full recognition + firing)
      all axes = 2          -> Stage 4 (peer group rewritten)
      axis_1=0 AND axis_2a=0 -> Stage 0 (no spark, no fuel)

    Exception: if axis4 (Earnings Materiality) == 0, cap stage at 2.
    (This LOWERS an otherwise-higher stage to 2. It never RAISES a stage.)
    """
    a1 = inp.axis_1_label_status
    a2a = inp.axis_2a_catalyst_potential
    a2b = inp.axis_2b_catalyst_realization
    a3 = inp.axis_3_diffusion_stage
    a4 = inp.axis_4_earnings_materiality
    a5 = inp.axis_5_narrative_durability

    # Stage 0: no spark and no fuel
    if a1 == 0 and a2a == 0:
        return (0, "axis_1 + axis_2a",
                "ラベル変化なし(火種なし)、触媒源泉も存在しない(燃料なし)", False)

    # Stage 4: peer group rewritten (all axes maxed)
    if (a1 >= 2 and a2a >= 2 and a2b >= 2
            and a3 >= 2 and a4 >= 2 and a5 >= 2):
        stage_raw = 4
        limiting = "全軸最大"
        reason = "全軸=2、ピアグループ書き換え達成(炎が全域に定着)"
    # Stage 3: full diffusion + catalyst fired
    elif a3 >= 2 and a2b >= 2:
        stage_raw = 3
        limiting = "axis_3 + axis_2b"
        reason = f"拡散進行度(reach) = {a3}、触媒実現度 = {a2b}、承認期"
    # Stage 2: (reach + fired) OR (sell-side reach even if unfired)
    elif (a3 >= 1 and a2b >= 2) or a3 >= 2:
        stage_raw = 2
        if a3 >= 2:
            limiting = "axis_2b (Catalyst Realization)"
            reason = (f"拡散進行度(reach) = {a3} でセルサイド到達も、"
                      f"触媒実現度 = {a2b} で続報待ち、Stage 3未到達")
        else:
            limiting = "axis_3 (Diffusion Stage)"
            reason = (f"触媒実現度 = {a2b} で発火、拡散進行度(reach) = {a3}、"
                      f"普及期")
    # Stage 1: spark/fuel present, reach still narrow
    else:
        stage_raw = 1
        if a2b <= 1:
            limiting = "axis_2b (Catalyst Realization)"
            reason = (f"触媒実現度 = {a2b} ≤ 1 で本格発火に至らず、"
                      f"拡散進行度(reach) = {a3}、萌芽期")
        else:
            limiting = "axis_3 (Diffusion Stage)"
            reason = (f"触媒発火済みだが拡散進行度(reach) = {a3} で"
                      f"市場に行き渡っていない、萌芽期")

    # Exception: cap at Stage 2 if Earnings Materiality == 0.
    # This only ever LOWERS the stage; it cannot push a Stage 1 up to 2.
    earnings_cap_applied = False
    if a4 == 0 and stage_raw > 2:
        stage_raw = 2
        limiting = "axis_4 (Earnings Materiality)"
        reason = "業績寄与=0のため Stage 上限 2 に制限(物語先行リスク)"
        earnings_cap_applied = True

    return (stage_raw, limiting, reason, earnings_cap_applied)


# ============================================================================
# Companion read: label rewrite (axis 1) vs reach (axis 3)
# ============================================================================

def _assess_headroom(inp: NarrativeInput) -> tuple[str, str]:
    """
    Axis 1 (label rewrite progress) does not gate the stage, but the gap
    between axis 1 and axis 3 (reach) carries information:

      axis1 > axis3 : the rewrite has run ahead of its reach -> the story is
                      more rewritten than the market has absorbed -> headroom
                      to spread (e.g. Denyo: axis1=2, axis3=1).
      axis1 < axis3 : reach has run ahead of the rewrite -> the label has
                      travelled wider than its substance -> reversal risk.
      axis1 == axis3: rewrite and reach are in step.

    Returns:
        (signal, note) where signal in {'HEADROOM', 'BALANCED', 'SHALLOW'}
    """
    a1 = inp.axis_1_label_status
    a3 = inp.axis_3_diffusion_stage
    if a1 > a3:
        return ('HEADROOM',
                f"軸1(置換)={a1} > 軸3(到達)={a3}: ラベル書き換えが到達範囲に"
                f"先行。伝播の伸びしろあり")
    elif a1 < a3:
        return ('SHALLOW',
                f"軸1(置換)={a1} < 軸3(到達)={a3}: 到達が置換に先行。"
                f"中身に対しラベルが広がりすぎ、逆行リスク")
    else:
        return ('BALANCED',
                f"軸1(置換)={a1} = 軸3(到達)={a3}: 置換と到達が歩調を揃えている")


# ============================================================================
# Companion metric: earnings impact at scale
# ============================================================================

def _calc_earnings_impact(inp: NarrativeInput) -> tuple[
    Optional[float], Optional[float], str
]:
    """
    Calculate earnings impact at scale (companion metric to stage).

    impact = TAM * share * OPM
    uplift_ratio = impact / current_OP

    Returns:
        (impact_oku, uplift_ratio, note)
    """
    if any(v is None for v in [
        inp.tam_oku_jpy,
        inp.expected_share_pct,
        inp.segment_opm_pct,
    ]):
        return (None, None, "TAM/share/OPMいずれか未入力のため算出不可")

    impact = (inp.tam_oku_jpy
              * (inp.expected_share_pct / 100.0)
              * (inp.segment_opm_pct / 100.0))

    if inp.current_operating_profit_oku is None or \
       inp.current_operating_profit_oku == 0:
        return (impact, None,
                f"営業利益インパクト ≈ {impact:.1f}億円(現状営利未入力)")

    uplift = impact / inp.current_operating_profit_oku
    note = (f"営業利益インパクト ≈ {impact:.1f}億円 / "
            f"現状営利 {inp.current_operating_profit_oku:.0f}億円 = "
            f"{uplift:.2f}倍化余地")
    return (impact, uplift, note)


# ============================================================================
# Verdict matrix
# ============================================================================

VERDICT_MATRIX: dict[tuple[BlockScore, str], tuple[Verdict, str]] = {
    # (fundamental_verdict, stage_bucket): (final_verdict, rationale)
    ('BUY', '0-1'): ('STRONG_BUY',
                     'ファンダ的に割安 × ナラティブStage 0-1 = ダブルバガー本命ゾーン'),
    ('BUY', '2'): ('BUY',
                   'ファンダ的に割安 × Stage 2、中期保有想定'),
    ('BUY', '3-4'): ('HOLD',
                     'ファンダ的に割安だが Stage 3-4 まで進行、出口検討'),
    ('HOLD', '0-1'): ('OBSERVE',
                      'ファンダ中立だが Stage 0-1、ナラティブ進行を観察'),
    ('HOLD', '2'): ('HOLD',
                    'ファンダ中立 × Stage 2、現状維持'),
    ('HOLD', '3-4'): ('TRIM',
                      'ファンダ中立 × Stage 3-4、段階利確検討'),
    ('CAUTION', '0-1'): ('AVOID',
                         'ファンダ的に割高 × Stage 0-1、機会費用を回避'),
    ('CAUTION', '2'): ('CAUTION',
                       'ファンダ的に割高 × Stage 2、見送り'),
    ('CAUTION', '3-4'): ('EXIT',
                         'ファンダ的に割高 × Stage 3-4、撤退ゾーン'),
}


def _stage_bucket(stage: Stage) -> str:
    """Map stage to verdict matrix bucket."""
    if stage in (0, 1):
        return '0-1'
    elif stage == 2:
        return '2'
    else:  # 3 or 4
        return '3-4'


def _determine_verdict(
    stage: Stage,
    fundamental: Optional[BlockScore],
) -> tuple[Optional[Verdict], str]:
    """Look up final verdict from matrix."""
    if fundamental is None:
        return (None, "Block 2-4 verdict が未指定のため最終Verdict算出不可")

    bucket = _stage_bucket(stage)
    key = (fundamental, bucket)
    if key not in VERDICT_MATRIX:
        return (None, f"未定義の組み合わせ: {key}")

    return VERDICT_MATRIX[key]


# ============================================================================
# Main API
# ============================================================================

def assess(inp: NarrativeInput) -> NarrativeResult:
    """
    Run Block 5 assessment on a stock.

    Args:
        inp: NarrativeInput with 6-axis scores and optional companion data.

    Returns:
        NarrativeResult with stage, verdict, and audit fields.
    """
    stage, limiting, reason, cap_applied = _determine_stage(inp)
    impact, uplift, impact_note = _calc_earnings_impact(inp)
    headroom_signal, headroom_note = _assess_headroom(inp)
    verdict, verdict_rationale = _determine_verdict(
        stage, inp.fundamental_verdict
    )

    total = (inp.axis_1_label_status
             + inp.axis_2a_catalyst_potential
             + inp.axis_2b_catalyst_realization
             + inp.axis_3_diffusion_stage
             + inp.axis_4_earnings_materiality
             + inp.axis_5_narrative_durability)

    return NarrativeResult(
        ticker=inp.ticker,
        company_name=inp.company_name,
        date=inp.date,
        stage=stage,
        stage_label=STAGE_LABELS[stage],
        rate_limiting_axis=limiting,
        rate_limiting_reason=reason,
        total_score=total,
        earnings_cap_applied=cap_applied,
        earnings_impact_oku=impact,
        earnings_uplift_ratio=uplift,
        earnings_impact_note=impact_note,
        headroom_signal=headroom_signal,
        headroom_note=headroom_note,
        final_verdict=verdict,
        verdict_rationale=verdict_rationale,
        axis_scores={
            '1_label_status': inp.axis_1_label_status,
            '2a_catalyst_potential': inp.axis_2a_catalyst_potential,
            '2b_catalyst_realization': inp.axis_2b_catalyst_realization,
            '3_diffusion_stage': inp.axis_3_diffusion_stage,
            '4_earnings_materiality': inp.axis_4_earnings_materiality,
            '5_narrative_durability': inp.axis_5_narrative_durability,
        },
        axis_notes={
            '1_label_status': inp.axis_1_note,
            '2a_catalyst_potential': inp.axis_2a_note,
            '2b_catalyst_realization': inp.axis_2b_note,
            '3_diffusion_stage': inp.axis_3_note,
            '4_earnings_materiality': inp.axis_4_note,
            '5_narrative_durability': inp.axis_5_note,
        },
    )
