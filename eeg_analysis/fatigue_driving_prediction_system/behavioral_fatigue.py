"""Shared event-level behavioral-fatigue evaluation.

Reaction times are rounded by the EDF event parser before they reach this
module.  Event windows deliberately use the integer ``event_second`` values
used by the rest of the prediction workflow.
"""

from __future__ import annotations

from bisect import bisect_right
from dataclasses import dataclass
from decimal import Decimal
from typing import Iterable, Sequence

from eeg_analysis.fatigue_driving_prediction_system.physiological_fatigue import (
    ceil_to_one_decimal,
)
from eeg_analysis.statistics_30s_alpha_eyeblink_of_fatigue.record_status_and_eyeblink_to_xlsx import (
    ReactionTimeEvent,
)


GLOBAL_RT_WINDOW_SECONDS = 90
PHASE_ONE_DURATION_SECONDS = 300
PHASE_ONE_RT_THRESHOLD = 1.6
PERSONALIZED_RT_MULTIPLIER = 1.5


@dataclass(frozen=True)
class BehavioralFatigueEvaluation:
    """Behavioral-fatigue flags calculated for one lane-deviation event."""

    event: ReactionTimeEvent
    global_rt: float
    active_threshold: float
    window_start_second: int
    window_end_second: int
    has_full_global_window: bool
    future_event_count: int
    local_exceed: bool
    global_exceed: bool
    sustained_fatigue: bool
    behavioral_fatigue: bool
    confirmation_second: int | None
    trigger_reason: str


def calculate_personalized_rt_threshold(
    reaction_times: Iterable[float],
    *,
    multiplier: float = PERSONALIZED_RT_MULTIPLIER,
    maximum_threshold: float = PHASE_ONE_RT_THRESHOLD,
) -> float:
    """Calculate the shared capped, upward-rounded personal RT threshold."""
    values = [float(value) for value in reaction_times]
    if not values:
        raise ValueError("personalized RT threshold requires baseline RT events")
    if multiplier <= 0 or maximum_threshold <= 0:
        raise ValueError("personalized RT threshold parameters must be positive")
    decimal_mean = sum(Decimal(str(value)) for value in values) / Decimal(
        len(values)
    )
    scaled_threshold = decimal_mean * Decimal(str(multiplier))
    return min(
        maximum_threshold,
        ceil_to_one_decimal(float(scaled_threshold)),
    )


def evaluate_behavioral_fatigue_events(
    events: Sequence[ReactionTimeEvent],
    active_threshold: float,
    *,
    global_window_seconds: int = GLOBAL_RT_WINDOW_SECONDS,
    recording_end_second: int | None = None,
) -> tuple[BehavioralFatigueEvaluation, ...]:
    """Confirm abnormal Local RT events with an inclusive forward window.

    For event ``i`` at integer second ``s``, Forward Global RT is the mean of
    rounded Local RT values for the current and later events satisfying
    ``s <= event_second <= s + global_window_seconds``.  Events earlier in the
    same integer second are excluded so each candidate only uses itself and
    events that follow it in event order.

    Behavioral fatigue requires a complete 90-second future interval, at least
    one event in a strictly later second, and both Local RT and Forward Global
    RT at or above the active threshold.  The onset remains ``s`` while the
    confirmation time is ``s + global_window_seconds``.
    """
    if active_threshold <= 0:
        raise ValueError("active_threshold must be greater than zero")
    if global_window_seconds <= 0:
        raise ValueError("global_window_seconds must be greater than zero")

    ordered_events = sorted(
        events,
        key=lambda event: (
            event.event_second,
            event.deviation_time,
            event.event_index,
        ),
    )
    if not ordered_events:
        return ()

    resolved_recording_end = (
        int(recording_end_second)
        if recording_end_second is not None
        else ordered_events[-1].event_second
    )
    if resolved_recording_end < 0:
        raise ValueError("recording_end_second must not be negative")
    if resolved_recording_end < ordered_events[-1].event_second:
        raise ValueError(
            "recording_end_second must include every supplied RT event"
        )

    event_seconds = [event.event_second for event in ordered_events]
    prefix_rt = [Decimal("0")]
    for event in ordered_events:
        prefix_rt.append(prefix_rt[-1] + Decimal(str(event.reaction_time)))
    threshold_decimal = Decimal(str(active_threshold))

    evaluations: list[BehavioralFatigueEvaluation] = []

    for index, event in enumerate(ordered_events):
        window_start = event.event_second
        window_end = window_start + global_window_seconds
        right_index = bisect_right(event_seconds, window_end, lo=index)
        window_count = right_index - index
        global_rt_decimal = (
            prefix_rt[right_index] - prefix_rt[index]
        ) / Decimal(window_count)
        global_rt = float(global_rt_decimal)
        first_strictly_later = bisect_right(
            event_seconds,
            window_start,
            lo=index + 1,
            hi=right_index,
        )
        future_event_count = right_index - first_strictly_later
        has_full_window = window_end <= resolved_recording_end
        local_exceed = event.reaction_time >= active_threshold
        global_exceed = global_rt_decimal >= threshold_decimal
        sustained_fatigue = (
            has_full_window
            and future_event_count > 0
            and local_exceed
            and global_exceed
        )
        behavioral_fatigue = sustained_fatigue
        evaluations.append(
            BehavioralFatigueEvaluation(
                event=event,
                global_rt=global_rt,
                active_threshold=active_threshold,
                window_start_second=window_start,
                window_end_second=window_end,
                has_full_global_window=has_full_window,
                future_event_count=future_event_count,
                local_exceed=local_exceed,
                global_exceed=global_exceed,
                sustained_fatigue=sustained_fatigue,
                behavioral_fatigue=behavioral_fatigue,
                confirmation_second=window_end if behavioral_fatigue else None,
                trigger_reason=(
                    "LOCAL_AND_FORWARD_GLOBAL"
                    if behavioral_fatigue
                    else "NONE"
                ),
            )
        )

    return tuple(evaluations)


def first_behavioral_fatigue(
    evaluations: Sequence[BehavioralFatigueEvaluation],
    *,
    after_second: int | None = None,
) -> BehavioralFatigueEvaluation | None:
    """Return the first triggered evaluation, optionally after a boundary."""
    return next(
        (
            evaluation
            for evaluation in evaluations
            if evaluation.behavioral_fatigue
            and (
                after_second is None
                or evaluation.event.event_second > after_second
            )
        ),
        None,
    )


__all__ = [
    "BehavioralFatigueEvaluation",
    "GLOBAL_RT_WINDOW_SECONDS",
    "PERSONALIZED_RT_MULTIPLIER",
    "PHASE_ONE_DURATION_SECONDS",
    "PHASE_ONE_RT_THRESHOLD",
    "calculate_personalized_rt_threshold",
    "evaluate_behavioral_fatigue_events",
    "first_behavioral_fatigue",
]
