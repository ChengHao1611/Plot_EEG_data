"""Shared-rule fatigue state calculation for the standalone Observe UI."""

from __future__ import annotations

import math
from dataclasses import dataclass

import pandas as pd

from eeg_analysis.fatigue_driving_prediction_system.behavioral_fatigue import (
    GLOBAL_RT_WINDOW_SECONDS,
    PHASE_ONE_DURATION_SECONDS,
    PHASE_ONE_RT_THRESHOLD,
    evaluate_backward_behavioral_fatigue_events,
    evaluate_behavioral_fatigue_events,
    first_behavioral_fatigue,
)
from eeg_analysis.fatigue_driving_prediction_system.physiological_fatigue import (
    TRAINING_POOLED_ALPHA_SCALE,
    TRAINING_POOLED_EYE_SCALE,
    compute_fatigue_score,
    compute_robust_baseline,
    confirmation_runs,
)
from eeg_analysis.statistics_30s_alpha_eyeblink_of_fatigue.record_status_and_eyeblink_to_xlsx import (
    ReactionTimeEvent,
    round_reaction_time,
)


@dataclass(frozen=True)
class FatigueAlgorithmConfig:
    """Observe Phase-1/2 behavioral and physiological settings."""

    baseline_end_second: int = PHASE_ONE_DURATION_SECONDS
    phase_one_rt_threshold: float = PHASE_ONE_RT_THRESHOLD
    global_rt_window_seconds: int = GLOBAL_RT_WINDOW_SECONDS
    complete_window_start_second: int = 30
    score_threshold: float = 1
    confirmation_seconds: int = 1
    pooled_alpha_scale: float = TRAINING_POOLED_ALPHA_SCALE
    pooled_eye_scale: float = TRAINING_POOLED_EYE_SCALE


DEFAULT_CONFIG = FatigueAlgorithmConfig()


def classify_actual_states(
    events_df: pd.DataFrame,
    personalized_rt_threshold: float,
    config: FatigueAlgorithmConfig = DEFAULT_CONFIG,
    *,
    recording_end_second: int | None = None,
) -> pd.DataFrame:
    """Apply Function One backward screening, then Function Two forward RT."""

    events = events_df[["second", "react_time"]].copy()
    events["second"] = pd.to_numeric(events["second"], errors="coerce")
    events["react_time"] = pd.to_numeric(events["react_time"], errors="coerce")
    events = events.dropna(subset=["second", "react_time"])
    events["_input_order"] = range(len(events))
    events = events.sort_values(["second", "_input_order"], kind="stable")

    reaction_events: list[ReactionTimeEvent] = []
    for event_index, event in enumerate(events.itertuples(index=False), start=1):
        raw_second = float(event.second)
        reaction_time = round_reaction_time(float(event.react_time))
        reaction_events.append(
            ReactionTimeEvent(
                event_index=event_index,
                deviation_status=0,
                deviation_time=raw_second,
                correction_start_time=raw_second + reaction_time,
                event_second=int(math.ceil(raw_second)),
                reaction_time=reaction_time,
            )
        )

    personalized_rt_threshold = float(personalized_rt_threshold)
    if not math.isfinite(personalized_rt_threshold) or personalized_rt_threshold <= 0:
        raise ValueError("Phase 2個人化RT門檻必須是大於0的有限數值。")

    phase_one_evaluations = tuple(
        evaluation
        for evaluation in evaluate_backward_behavioral_fatigue_events(
            reaction_events,
            config.phase_one_rt_threshold,
            global_window_seconds=config.global_rt_window_seconds,
            recording_start_second=0,
        )
        if evaluation.event.event_second <= config.baseline_end_second
    )
    phase_one_trigger = first_behavioral_fatigue(phase_one_evaluations)
    phase_one_blocked = phase_one_trigger is not None

    phase_two_evaluations = (
        tuple(
            evaluation
            for evaluation in evaluate_behavioral_fatigue_events(
                reaction_events,
                personalized_rt_threshold,
                global_window_seconds=config.global_rt_window_seconds,
                recording_end_second=recording_end_second,
            )
            if evaluation.event.event_second > config.baseline_end_second
        )
        if not phase_one_blocked
        else ()
    )
    phase_two_trigger = first_behavioral_fatigue(phase_two_evaluations)
    evaluations_by_index = {
        evaluation.event.event_index: evaluation
        for evaluation in (*phase_one_evaluations, *phase_two_evaluations)
    }

    rows: list[dict[str, object]] = []
    fatigue_active = False
    for event in reaction_events:
        evaluation = evaluations_by_index.get(event.event_index)
        if evaluation is not None:
            fatigue_active = fatigue_active or evaluation.behavioral_fatigue
            phase = (
                "PHASE_1"
                if event.event_second <= config.baseline_end_second
                else "PHASE_2"
            )
            global_direction = evaluation.window_direction
            global_rt = evaluation.global_rt
            window_start = evaluation.window_start_second
            window_end = evaluation.window_end_second
            complete_window = evaluation.has_full_global_window
            past_event_count = evaluation.past_event_count
            future_event_count = evaluation.future_event_count
            active_threshold = evaluation.active_threshold
            local_exceed = evaluation.local_exceed
            global_exceed = evaluation.global_exceed
            sustained_fatigue = evaluation.sustained_fatigue
            behavioral_fatigue = evaluation.behavioral_fatigue
            confirmation_second = evaluation.confirmation_second
            trigger_reason = evaluation.trigger_reason
        else:
            # A Phase-1 rejection stops all Phase-2 judgments, but later RT
            # events remain in the table so the top chart can show them.
            phase = "PHASE_2_SKIPPED"
            global_direction = None
            global_rt = None
            window_start = None
            window_end = None
            complete_window = None
            past_event_count = None
            future_event_count = None
            active_threshold = None
            local_exceed = None
            global_exceed = None
            sustained_fatigue = None
            behavioral_fatigue = None
            confirmation_second = None
            trigger_reason = None

        rows.append(
            {
                "second": event.event_second,
                "react_time": event.reaction_time,
                "phase": phase,
                "global_direction": global_direction,
                "global_rt": global_rt,
                "window_start_second": window_start,
                "window_end_second": window_end,
                "has_full_global_window": complete_window,
                "has_full_forward_window": (
                    complete_window if global_direction == "FORWARD" else None
                ),
                "past_event_count": past_event_count,
                "future_event_count": future_event_count,
                "supporting_event_count": (
                    past_event_count
                    if global_direction == "BACKWARD"
                    else future_event_count
                ),
                "active_threshold": active_threshold,
                "local_exceed": local_exceed,
                "global_exceed": global_exceed,
                "sustained_fatigue": sustained_fatigue,
                "behavioral_fatigue": behavioral_fatigue,
                "confirmation_second": confirmation_second,
                "trigger_reason": trigger_reason,
                "state_val": int(fatigue_active),
            }
        )

    result = pd.DataFrame(
        rows,
        columns=[
            "second",
            "react_time",
            "phase",
            "global_direction",
            "global_rt",
            "window_start_second",
            "window_end_second",
            "has_full_global_window",
            "has_full_forward_window",
            "past_event_count",
            "future_event_count",
            "supporting_event_count",
            "active_threshold",
            "local_exceed",
            "global_exceed",
            "sustained_fatigue",
            "behavioral_fatigue",
            "confirmation_second",
            "trigger_reason",
            "state_val",
        ],
    )
    result.attrs["personalized_rt_threshold"] = float(personalized_rt_threshold)
    result.attrs["phase_one_rt_threshold"] = config.phase_one_rt_threshold
    result.attrs["phase_one_behavioral_rule"] = (
        f"(Local>={config.phase_one_rt_threshold:g} AND "
        f"BackwardGlobal{config.global_rt_window_seconds}>="
        f"{config.phase_one_rt_threshold:g} AND complete backward window "
        "AND prior RT exists)"
    )
    result.attrs["phase_two_behavioral_rule"] = (
        "(Local>=personalized threshold AND "
        f"ForwardGlobal{config.global_rt_window_seconds}>=personalized threshold "
        "AND complete forward window AND future RT exists)"
    )
    result.attrs["behavioral_rule"] = (
        "Phase 1: "
        + result.attrs["phase_one_behavioral_rule"]
        + "; Phase 2: "
        + result.attrs["phase_two_behavioral_rule"]
    )
    result.attrs["global_rt_window_seconds"] = config.global_rt_window_seconds
    first_trigger = phase_one_trigger or phase_two_trigger
    result.attrs["first_confirmation_second"] = (
        first_trigger.confirmation_second if first_trigger else None
    )
    result.attrs["phase_one_blocked"] = phase_one_blocked
    return result


def predict_fatigue_states(
    master_df: pd.DataFrame,
    config: FatigueAlgorithmConfig = DEFAULT_CONFIG,
) -> pd.DataFrame:
    """Compute Median/MAD Z scores and a confirmed ``min`` warning state."""
    required = {"second", "alpha_sum", "eye_sum"}
    missing = required.difference(master_df.columns)
    if missing:
        raise ValueError(f"生理資料缺少欄位：{', '.join(sorted(missing))}")

    result = master_df.copy()
    baseline_mask = result["second"].between(
        config.complete_window_start_second,
        config.baseline_end_second,
        inclusive="both",
    )
    alpha_baseline = compute_robust_baseline(
        result.loc[baseline_mask, "alpha_sum"].to_numpy(dtype=float),
        pooled_scale=config.pooled_alpha_scale,
    )
    eye_baseline = compute_robust_baseline(
        result.loc[baseline_mask, "eye_sum"].to_numpy(dtype=float),
        pooled_scale=config.pooled_eye_scale,
    )
    invalid_features = [
        name
        for name, baseline in (
            ("Alpha", alpha_baseline),
            ("Eye blink", eye_baseline),
        )
        if not baseline.valid
    ]
    if invalid_features:
        raise ValueError(
            "、".join(invalid_features)
            + " 的前300秒MAD與IQR皆為0，且未提供訓練資料Pooled Scale。"
        )

    z_alpha_values: list[float] = []
    z_eye_values: list[float] = []
    scores: list[float | None] = []
    for row in result.itertuples(index=False):
        if row.second <= config.baseline_end_second:
            z_alpha_values.append(float("nan"))
            z_eye_values.append(float("nan"))
            scores.append(None)
            continue
        score = compute_fatigue_score(
            row.alpha_sum,
            row.eye_sum,
            alpha_baseline=alpha_baseline,
            eye_baseline=eye_baseline,
        )
        z_alpha_values.append(
            float(score.z_alpha) if score.z_alpha is not None else float("nan")
        )
        z_eye_values.append(
            float(score.z_eye) if score.z_eye is not None else float("nan")
        )
        scores.append(score.score)

    above, runs, warnings = confirmation_runs(
        scores,
        threshold=config.score_threshold,
        confirmation_seconds=config.confirmation_seconds,
    )
    fatigue_active = False
    pred_states: list[int] = []
    for warning in warnings:
        fatigue_active = fatigue_active or warning
        pred_states.append(int(fatigue_active))

    result["z_alpha"] = z_alpha_values
    result["z_eye"] = z_eye_values
    result["fatigue_score"] = [
        float(value) if value is not None else float("nan") for value in scores
    ]
    result["score_above_threshold"] = [int(value) for value in above]
    result["consecutive_seconds"] = runs
    result["alpha_alarm"] = (
        pd.Series(z_alpha_values, index=result.index) >= config.score_threshold
    ).astype(int)
    result["eye_alarm"] = (
        pd.Series(z_eye_values, index=result.index) >= config.score_threshold
    ).astype(int)
    result["pred_state"] = pred_states
    result.attrs["alpha_baseline"] = alpha_baseline
    result.attrs["eye_baseline"] = eye_baseline
    return result


__all__ = [
    "DEFAULT_CONFIG",
    "FatigueAlgorithmConfig",
    "classify_actual_states",
    "predict_fatigue_states",
]
