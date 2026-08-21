"""Shared event-level behavioral-fatigue evaluation.

Reaction times are rounded by the EDF event parser before they reach this
module.  Event windows deliberately use the integer ``event_second`` values
used by the rest of the prediction workflow.
"""

from __future__ import annotations

from collections import deque
from dataclasses import dataclass
from decimal import Decimal
from statistics import fmean
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
CRITICAL_LOCAL_RT_THRESHOLD = 3.2
PERSONALIZED_RT_MULTIPLIER = 1.5


@dataclass(frozen=True)
class BehavioralFatigueEvaluation:
    """Behavioral-fatigue flags calculated for one lane-deviation event."""

    event: ReactionTimeEvent
    global_rt: float
    active_threshold: float
    has_full_global_window: bool
    local_exceed: bool
    global_exceed: bool
    sustained_fatigue: bool
    critical_lapse: bool
    behavioral_fatigue: bool
    trigger_reason: str


def _trigger_reason(*, sustained_fatigue: bool, critical_lapse: bool) -> str:
    if sustained_fatigue and critical_lapse:
        return "SUSTAINED_AND_CRITICAL"
    if critical_lapse:
        return "CRITICAL_LOCAL_RT"
    if sustained_fatigue:
        return "LOCAL_AND_GLOBAL"
    return "NONE"


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
    critical_local_rt_threshold: float = CRITICAL_LOCAL_RT_THRESHOLD,
    recording_start_second: int = 0,
) -> tuple[BehavioralFatigueEvaluation, ...]:
    """Evaluate Local RT, inclusive trailing Global RT, and critical lapses.

    For event ``i`` at integer second ``s``, Global RT is the mean of rounded
    Local RT values from already-observed events satisfying
    ``s - global_window_seconds <= event_second <= s``.  The current event is
    appended before the mean is calculated, while later events sharing the
    same integer second are excluded to prevent look-ahead.

    Sustained fatigue requires a complete recording-time window plus both
    Local RT and Global RT at or above the active threshold.  A critical Local
    RT triggers immediately, including during the initial Global RT warm-up.
    """
    if active_threshold <= 0:
        raise ValueError("active_threshold must be greater than zero")
    if global_window_seconds <= 0:
        raise ValueError("global_window_seconds must be greater than zero")
    if critical_local_rt_threshold <= 0:
        raise ValueError("critical_local_rt_threshold must be greater than zero")

    ordered_events = sorted(
        events,
        key=lambda event: (
            event.event_second,
            event.deviation_time,
            event.event_index,
        ),
    )
    trailing_events: deque[ReactionTimeEvent] = deque()
    evaluations: list[BehavioralFatigueEvaluation] = []

    for event in ordered_events:
        window_start = event.event_second - global_window_seconds
        while trailing_events and trailing_events[0].event_second < window_start:
            trailing_events.popleft()
        trailing_events.append(event)

        global_rt = fmean(item.reaction_time for item in trailing_events)
        has_full_window = (
            event.event_second - recording_start_second >= global_window_seconds
        )
        local_exceed = event.reaction_time >= active_threshold
        global_exceed = global_rt >= active_threshold
        sustained_fatigue = has_full_window and local_exceed and global_exceed
        critical_lapse = event.reaction_time >= critical_local_rt_threshold
        behavioral_fatigue = sustained_fatigue or critical_lapse
        evaluations.append(
            BehavioralFatigueEvaluation(
                event=event,
                global_rt=global_rt,
                active_threshold=active_threshold,
                has_full_global_window=has_full_window,
                local_exceed=local_exceed,
                global_exceed=global_exceed,
                sustained_fatigue=sustained_fatigue,
                critical_lapse=critical_lapse,
                behavioral_fatigue=behavioral_fatigue,
                trigger_reason=_trigger_reason(
                    sustained_fatigue=sustained_fatigue,
                    critical_lapse=critical_lapse,
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
    "CRITICAL_LOCAL_RT_THRESHOLD",
    "GLOBAL_RT_WINDOW_SECONDS",
    "PERSONALIZED_RT_MULTIPLIER",
    "PHASE_ONE_DURATION_SECONDS",
    "PHASE_ONE_RT_THRESHOLD",
    "calculate_personalized_rt_threshold",
    "evaluate_behavioral_fatigue_events",
    "first_behavioral_fatigue",
]
