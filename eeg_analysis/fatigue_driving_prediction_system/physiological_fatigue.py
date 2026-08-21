"""Shared robust physiological-fatigue scoring utilities.

Both Function Two and the standalone Observe UI use this module so that the
online warning rule cannot silently diverge between the two entry points.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_CEILING
from typing import Iterable, Sequence

import numpy as np


MAD_NORMAL_SCALE = 1.4826
IQR_NORMAL_SCALE = 1.349

# Frozen from training_pooled_baseline.xlsx so production analysis does not
# depend on rescanning train_data at runtime.  These values are fallbacks only:
# a recording's own MAD or IQR scale still takes precedence when valid.
TRAINING_POOLED_ALPHA_SCALE = 1.4826
TRAINING_POOLED_EYE_SCALE = 4.4478


@dataclass(frozen=True)
class RobustBaseline:
    """Location and robust scale calculated from complete baseline windows."""

    median: float
    mad: float
    iqr: float
    scale: float | None
    scale_method: str
    sample_count: int

    @property
    def valid(self) -> bool:
        return self.scale is not None and self.scale > 0


@dataclass(frozen=True)
class FatigueScore:
    """Standardized Alpha/Eye evidence for one window ending at one second."""

    z_alpha: float | None
    z_eye: float | None
    score: float | None


def ceil_to_one_decimal(value: float) -> float:
    """Round a finite value upward to one decimal place.

    Decimal-from-string avoids turning an already exact value such as 1.6
    into 1.7 because of binary floating-point representation.
    """
    decimal_value = Decimal(str(value))
    if not decimal_value.is_finite():
        raise ValueError("value must be finite")
    return float(decimal_value.quantize(Decimal("0.1"), rounding=ROUND_CEILING))


def rolling_event_counts(
    event_seconds: Iterable[int],
    *,
    start_second: int,
    end_second: int,
    window_seconds: int = 30,
) -> dict[int, int]:
    """Count event-bearing seconds in trailing inclusive windows.

    At second 30 with a 30-second window, the result covers seconds 1--30.
    Events before ``start_second`` may still contribute to a later requested
    window, which is required when Phase 2 begins at second 301.
    """
    if window_seconds <= 0:
        raise ValueError("window_seconds must be positive")
    if end_second < start_second:
        return {}

    ordered_events = sorted({int(second) for second in event_seconds})
    counts: dict[int, int] = {}
    left = 0
    right = 0
    for second in range(start_second, end_second + 1):
        window_start = second - window_seconds + 1
        while left < len(ordered_events) and ordered_events[left] < window_start:
            left += 1
        if right < left:
            right = left
        while right < len(ordered_events) and ordered_events[right] <= second:
            right += 1
        counts[second] = right - left
    return counts


def compute_robust_baseline(
    values: Sequence[float] | np.ndarray,
    *,
    pooled_scale: float | None = None,
) -> RobustBaseline:
    """Return Median/MAD, with IQR then training-pooled scale fallbacks."""
    array = np.asarray(values, dtype=float)
    array = array[np.isfinite(array)]
    if array.size == 0:
        raise ValueError("baseline requires at least one finite complete window")

    median = float(np.median(array))
    mad = float(np.median(np.abs(array - median)))
    q1, q3 = np.percentile(array, [25, 75])
    iqr = float(q3 - q1)
    mad_scale = MAD_NORMAL_SCALE * mad
    iqr_scale = iqr / IQR_NORMAL_SCALE

    if mad_scale > 0:
        scale: float | None = mad_scale
        method = "MAD"
    elif iqr_scale > 0:
        scale = iqr_scale
        method = "IQR"
    elif pooled_scale is not None and np.isfinite(pooled_scale) and pooled_scale > 0:
        scale = float(pooled_scale)
        method = "POOLED"
    else:
        scale = None
        method = "INVALID"

    return RobustBaseline(
        median=median,
        mad=mad,
        iqr=iqr,
        scale=scale,
        scale_method=method,
        sample_count=int(array.size),
    )


def estimate_pooled_scale(baselines: Iterable[RobustBaseline]) -> float | None:
    """Return the median non-pooled valid scale from training baselines."""
    scales = [
        float(item.scale)
        for item in baselines
        if item.valid and item.scale_method in {"MAD", "IQR"}
    ]
    return float(np.median(scales)) if scales else None


def compute_fatigue_score(
    alpha_count: float,
    eye_count: float,
    *,
    alpha_baseline: RobustBaseline,
    eye_baseline: RobustBaseline,
) -> FatigueScore:
    """Return directional robust Z scores and their conjunctive minimum."""
    if not alpha_baseline.valid or not eye_baseline.valid:
        return FatigueScore(None, None, None)
    assert alpha_baseline.scale is not None
    assert eye_baseline.scale is not None
    z_alpha = (float(alpha_count) - alpha_baseline.median) / alpha_baseline.scale
    # Lower Eye30 is the fatigue direction, so reverse its sign.
    z_eye = (eye_baseline.median - float(eye_count)) / eye_baseline.scale
    return FatigueScore(z_alpha, z_eye, min(z_alpha, z_eye))


def confirmation_runs(
    scores: Iterable[float | None],
    *,
    threshold: float = 1.0,
    confirmation_seconds: int = 5,
) -> tuple[list[bool], list[int], list[bool]]:
    """Return threshold flags, run lengths, and real-time confirmed warnings."""
    if confirmation_seconds <= 0:
        raise ValueError("confirmation_seconds must be positive")
    above_values: list[bool] = []
    run_lengths: list[int] = []
    warnings: list[bool] = []
    run_length = 0
    for score in scores:
        above = score is not None and np.isfinite(score) and score >= threshold
        run_length = run_length + 1 if above else 0
        above_values.append(bool(above))
        run_lengths.append(run_length)
        warnings.append(run_length >= confirmation_seconds)
    return above_values, run_lengths, warnings


def classify_lead_time(
    lead_seconds: int | float | None,
    *,
    min_lead_seconds: int = 30,
    max_lead_seconds: int = 60,
) -> str:
    """Classify a warning relative to the required pre-fatigue target interval."""
    if lead_seconds is None:
        return "TARGET_WITHOUT_WARNING"
    if lead_seconds > max_lead_seconds:
        return "WARNING_TOO_EARLY"
    if min_lead_seconds <= lead_seconds <= max_lead_seconds:
        return "TARGET_PREDICTED_30_TO_60"
    if 0 < lead_seconds < min_lead_seconds:
        return "WARNING_TOO_LATE"
    return "WARNING_AT_OR_AFTER_TARGET"


__all__ = [
    "FatigueScore",
    "RobustBaseline",
    "TRAINING_POOLED_ALPHA_SCALE",
    "TRAINING_POOLED_EYE_SCALE",
    "ceil_to_one_decimal",
    "classify_lead_time",
    "compute_fatigue_score",
    "compute_robust_baseline",
    "confirmation_runs",
    "estimate_pooled_scale",
    "rolling_event_counts",
]
