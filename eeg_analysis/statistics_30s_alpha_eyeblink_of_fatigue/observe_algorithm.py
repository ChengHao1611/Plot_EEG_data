"""Shared-rule fatigue state calculation for the standalone Observe UI."""

from __future__ import annotations

import math
from dataclasses import dataclass

import pandas as pd

from eeg_analysis.fatigue_driving_prediction_system.behavioral_fatigue import (
    GLOBAL_RT_WINDOW_SECONDS,
    PHASE_ONE_DURATION_SECONDS,
    evaluate_behavioral_fatigue_events,
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
    """Phase-2 physiological warning settings."""

    baseline_end_second: int = PHASE_ONE_DURATION_SECONDS
    global_rt_window_seconds: int = GLOBAL_RT_WINDOW_SECONDS
    complete_window_start_second: int = 30
    score_threshold: float = 0.8
    confirmation_seconds: int = 4
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
    """Build the Function Two behavioral state from event-level RT data.

    Event seconds are upward-rounded to integers and Local RT values use the
    same conventional one-decimal rounding as the EDF parser. Forward Global RT
    is the inclusive mean from each abnormal event through the following 90
    seconds. The displayed behavioral state begins at the first forward-
    confirmed trigger; an onset through second 300 represents Phase 1 rejection.
    """

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
        raise ValueError("個人化RT門檻必須是大於0的有限數值。")

    evaluations = evaluate_behavioral_fatigue_events(
        reaction_events,
        personalized_rt_threshold,
        global_window_seconds=config.global_rt_window_seconds,
        recording_end_second=recording_end_second,
    )

    rows: list[dict[str, float | int | bool | str]] = []
    fatigue_active = False
    for evaluation in evaluations:
        fatigue_active = fatigue_active or evaluation.behavioral_fatigue
        rows.append(
            {
                "second": evaluation.event.event_second,
                "react_time": evaluation.event.reaction_time,
                "global_rt": evaluation.global_rt,
                "window_start_second": evaluation.window_start_second,
                "window_end_second": evaluation.window_end_second,
                "has_full_forward_window": evaluation.has_full_global_window,
                "future_event_count": evaluation.future_event_count,
                "active_threshold": evaluation.active_threshold,
                "local_exceed": evaluation.local_exceed,
                "global_exceed": evaluation.global_exceed,
                "sustained_fatigue": evaluation.sustained_fatigue,
                "behavioral_fatigue": evaluation.behavioral_fatigue,
                "confirmation_second": evaluation.confirmation_second,
                "trigger_reason": evaluation.trigger_reason,
                "state_val": int(fatigue_active),
            }
        )

    result = pd.DataFrame(
        rows,
        columns=[
            "second",
            "react_time",
            "global_rt",
            "window_start_second",
            "window_end_second",
            "has_full_forward_window",
            "future_event_count",
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
    result.attrs["behavioral_rule"] = (
        "(Local>=threshold AND "
        f"ForwardGlobal{config.global_rt_window_seconds}>=threshold AND "
        "complete forward window AND future RT exists)"
    )
    result.attrs["global_rt_window_seconds"] = config.global_rt_window_seconds
    first_trigger = next(
        (
            evaluation
            for evaluation in evaluations
            if evaluation.behavioral_fatigue
        ),
        None,
    )
    result.attrs["first_confirmation_second"] = (
        first_trigger.confirmation_second if first_trigger else None
    )
    result.attrs["phase_one_blocked"] = any(
        evaluation.behavioral_fatigue
        and evaluation.event.event_second <= config.baseline_end_second
        for evaluation in evaluations
    )
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
