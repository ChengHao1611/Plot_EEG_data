from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path
from typing import Iterable, Sequence

import matplotlib
import numpy as np

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FormatStrFormatter

EYEBLINK_AMPLITUDE_MIN = 5000.0
EYEBLINK_AMPLITUDE_MAX = 15000.0
EYEBLINK_BIN_WIDTH = 500.0
DEFAULT_BIN_COUNT = 20
PEAK_DISTRIBUTION_X_MIN = 0.0
PEAK_DISTRIBUTION_X_MAX = 0.0002
PEAK_DISTRIBUTION_X_TICK = 0.00001


def normalize_user_path(raw_path: str) -> Path:
    cleaned = raw_path.strip().strip('"').strip("'")
    return Path(cleaned).expanduser().resolve()


def collect_amplitudes_by_seconds(
    amplitudes: np.ndarray,
    seconds: Sequence[int],
) -> np.ndarray:
    values: list[float] = []
    for second in seconds:
        index = int(second) - 1
        if index < 0 or index >= amplitudes.size:
            continue
        value = float(amplitudes[index])
        if not np.isfinite(value):
            continue
        values.append(value)
    return np.asarray(values, dtype=float)


def clip_amplitudes(
    values: np.ndarray,
    *,
    min_value: float = EYEBLINK_AMPLITUDE_MIN,
    max_value: float = EYEBLINK_AMPLITUDE_MAX,
) -> np.ndarray:
    if values.size == 0:
        return values.astype(float, copy=True)
    return np.clip(np.asarray(values, dtype=float), float(min_value), float(max_value))


def _to_finite_array(values: Iterable[float]) -> np.ndarray:
    result: list[float] = []
    for value in values:
        value_float = float(value)
        if np.isfinite(value_float):
            result.append(value_float)
    return np.asarray(result, dtype=float)


def clip_peak_distribution_values(
    values: np.ndarray,
    *,
    min_value: float = PEAK_DISTRIBUTION_X_MIN,
    max_value: float = PEAK_DISTRIBUTION_X_MAX,
) -> np.ndarray:
    if values.size == 0:
        return values.astype(float, copy=True)
    return np.clip(np.asarray(values, dtype=float), float(min_value), float(max_value))


def compute_peak_to_shoulder_amplitude(
    peak_value: float,
    left_min: float,
    right_min: float,
) -> float | None:
    peak_float = float(peak_value)
    left_float = float(left_min)
    right_float = float(right_min)
    if not (np.isfinite(peak_float) and np.isfinite(left_float) and np.isfinite(right_float)):
        return None
    return peak_float - max(left_float, right_float)


def compute_shared_bin_edges(
    groups: Sequence[np.ndarray],
    *,
    bin_count: int = DEFAULT_BIN_COUNT,
    min_value: float | None = None,
    max_value: float | None = None,
) -> np.ndarray:
    finite_groups = [np.asarray(group, dtype=float) for group in groups if np.asarray(group).size > 0]
    if finite_groups:
        all_values = np.concatenate(finite_groups)
        auto_min = float(np.min(all_values))
        auto_max = float(np.max(all_values))
    else:
        auto_min = 0.0
        auto_max = 1.0

    x_min = auto_min if min_value is None else float(min_value)
    x_max = auto_max if max_value is None else float(max_value)

    if np.isclose(x_min, x_max):
        padding = max(abs(x_min) * 0.05, 1e-9)
        x_min -= padding
        x_max += padding

    final_bin_count = max(int(bin_count), 1)
    return np.linspace(x_min, x_max, final_bin_count + 1, dtype=float)


def plot_distribution(
    ax: plt.Axes,
    values: np.ndarray,
    *,
    label: str,
    color: str,
    bin_edges: np.ndarray,
) -> None:
    if values.size == 0:
        return

    ax.hist(
        values,
        bins=bin_edges,
        color=color,
        alpha=0.72,
        edgecolor="white",
        linewidth=0.8,
        label=f"{label} (n={values.size})",
    )


def save_eyeblink_amplitude_distribution(
    output_path: Path,
    one_hz_amplitudes: np.ndarray,
    true_seconds: Sequence[int],
    false_seconds: Sequence[int],
) -> tuple[int, int]:
    true_values = clip_amplitudes(collect_amplitudes_by_seconds(one_hz_amplitudes, true_seconds))
    false_values = clip_amplitudes(collect_amplitudes_by_seconds(one_hz_amplitudes, false_seconds))

    bin_edges = np.arange(
        float(EYEBLINK_AMPLITUDE_MIN),
        float(EYEBLINK_AMPLITUDE_MAX) + float(EYEBLINK_BIN_WIDTH),
        float(EYEBLINK_BIN_WIDTH),
        dtype=float,
    )

    fig, ax = plt.subplots(figsize=(13, 7))
    plot_distribution(
        ax,
        true_values,
        label="true eyeblink 1Hz amplitude",
        color="#2a9d8f",
        bin_edges=bin_edges,
    )
    plot_distribution(
        ax,
        false_values,
        label="false eyeblink 1Hz amplitude",
        color="#e76f51",
        bin_edges=bin_edges,
    )

    if true_values.size == 0 and false_values.size == 0:
        ax.text(
            0.5,
            0.5,
            "No eyeblink amplitude data available",
            transform=ax.transAxes,
            ha="center",
            va="center",
            fontsize=13,
        )

    ax.set_title("Eyeblink 1Hz Amplitude Distribution")
    ax.set_xlabel("1Hz Amplitude (clipped to 5000-15000)")
    ax.set_ylabel("Count")
    ax.set_xlim(float(EYEBLINK_AMPLITUDE_MIN), float(EYEBLINK_AMPLITUDE_MAX))
    ax.set_xticks(bin_edges)
    ax.grid(True, alpha=0.2)
    handles, labels = ax.get_legend_handles_labels()
    if handles:
        ax.legend(handles, labels, loc="upper right")

    fig.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=180, bbox_inches="tight")
    plt.close(fig)

    return int(true_values.size), int(false_values.size)


def save_eye_movement_peak_distribution(
    output_path: Path,
    positive_amplitudes: Sequence[float],
    false_amplitudes: Sequence[float],
    *,
    bin_count: int = DEFAULT_BIN_COUNT,
    min_value: float | None = None,
    max_value: float | None = None,
) -> tuple[int, int]:
    positive_values = clip_peak_distribution_values(_to_finite_array(positive_amplitudes))
    false_values = clip_peak_distribution_values(_to_finite_array(false_amplitudes))
    x_min = PEAK_DISTRIBUTION_X_MIN if min_value is None else float(min_value)
    x_max = PEAK_DISTRIBUTION_X_MAX if max_value is None else float(max_value)
    bin_edges = compute_shared_bin_edges(
        [positive_values, false_values],
        bin_count=bin_count,
        min_value=x_min,
        max_value=x_max,
    )
    tick_positions = np.arange(
        x_min,
        x_max + (PEAK_DISTRIBUTION_X_TICK / 2.0),
        PEAK_DISTRIBUTION_X_TICK,
        dtype=float,
    )

    fig, ax = plt.subplots(figsize=(13, 7))

    plot_distribution(
        ax,
        positive_values,
        label="true + miss amplitude",
        color="#2a9d8f",
        bin_edges=bin_edges,
    )
    plot_distribution(
        ax,
        false_values,
        label="false amplitude",
        color="#e76f51",
        bin_edges=bin_edges,
    )

    if positive_values.size == 0 and false_values.size == 0:
        ax.text(
            0.5,
            0.5,
            "No eye movement amplitude data available",
            transform=ax.transAxes,
            ha="center",
            va="center",
            fontsize=12,
        )

    ax.set_title("Eye Movement Amplitude Distribution Comparison")
    ax.set_xlim(x_min, x_max)
    ax.set_xticks(tick_positions)
    ax.xaxis.set_major_formatter(FormatStrFormatter("%.6f"))
    ax.set_xlabel("Peak - max(left_min, right_min)")
    ax.set_ylabel("Count")
    ax.grid(True, alpha=0.2)
    ax.tick_params(axis="x", rotation=45)
    handles, labels = ax.get_legend_handles_labels()
    if handles:
        ax.legend(handles, labels, loc="upper right")
    fig.tight_layout()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=180, bbox_inches="tight")
    plt.close(fig)

    return int(positive_values.size), int(false_values.size)


def load_event_amplitudes(event_summary_path: Path) -> tuple[np.ndarray, np.ndarray]:
    positive_values: list[float] = []
    false_values: list[float] = []

    with event_summary_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
        reader = csv.DictReader(csv_file)
        for row in reader:
            category = (row.get("category") or "").strip().lower()
            current_val_raw = (row.get("current_val") or "").strip()
            left_min_raw = (row.get("left_min") or "").strip()
            right_min_raw = (row.get("right_min") or "").strip()
            if not current_val_raw or not left_min_raw or not right_min_raw:
                continue

            try:
                current_val = float(current_val_raw)
                left_min = float(left_min_raw)
                right_min = float(right_min_raw)
            except ValueError:
                continue

            amplitude = compute_peak_to_shoulder_amplitude(current_val, left_min, right_min)
            if amplitude is None:
                continue

            if category in {"true", "miss"}:
                positive_values.append(amplitude)
            elif category == "false":
                false_values.append(amplitude)

    return np.asarray(positive_values, dtype=float), np.asarray(false_values, dtype=float)


def resolve_event_summary_path(input_path: Path) -> Path:
    if input_path.is_dir():
        candidate = input_path / "event_summary.csv"
        if candidate.exists():
            return candidate
        raise FileNotFoundError(f"event_summary.csv not found in {input_path}")

    if input_path.name.lower() == "event_summary.csv":
        return input_path

    raise FileNotFoundError(
        "Input must be a record_arousal_test output folder or a direct event_summary.csv path."
    )


def default_output_path(event_summary_path: Path) -> Path:
    return event_summary_path.with_name("eye_movement_peak_distribution.png")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Plot eye movement peak-amplitude distributions from record_arousal_test output. "
            "The distribution overlays (true + miss) and false on one shared axis, using "
            "amplitude = peak - max(left_min, right_min)."
        )
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        help="record_arousal_test output folder or direct event_summary.csv path.",
    )
    parser.add_argument(
        "--output",
        help="Optional PNG output path. Defaults to eye_movement_peak_distribution.png next to event_summary.csv.",
    )
    parser.add_argument(
        "--bins",
        type=int,
        default=DEFAULT_BIN_COUNT,
        help="Number of histogram bins shared by both subplots.",
    )
    parser.add_argument(
        "--x-min",
        type=float,
        help="Optional shared x-axis minimum. By default it is derived from the data.",
    )
    parser.add_argument(
        "--x-max",
        type=float,
        help="Optional shared x-axis maximum. By default it is derived from the data.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    raw_input_path = args.input_path
    if not raw_input_path:
        raw_input_path = input("請輸入 record_arousal_test 輸出資料夾或 event_summary.csv 路徑: ")

    input_path = normalize_user_path(raw_input_path)
    try:
        event_summary_path = resolve_event_summary_path(input_path)
    except FileNotFoundError as exc:
        print(f"Error: {exc}")
        return 1

    positive_values, false_values = load_event_amplitudes(event_summary_path)
    output_path = normalize_user_path(args.output) if args.output else default_output_path(event_summary_path)
    positive_count, false_count = save_eye_movement_peak_distribution(
        output_path,
        positive_values,
        false_values,
        bin_count=int(args.bins),
        min_value=args.x_min,
        max_value=args.x_max,
    )

    print(f"Event summary: {event_summary_path}")
    print(f"Output: {output_path}")
    print(f"(true + miss) count: {positive_count}")
    print(f"false count: {false_count}")
    print(f"Shared bin count: {int(args.bins)}")
    if args.x_min is not None or args.x_max is not None:
        print(f"Shared x-range override: min={args.x_min}, max={args.x_max}")

    return 0


if __name__ == "__main__":
    sys.exit(main())


__all__ = [
    "save_eyeblink_amplitude_distribution",
    "save_eye_movement_peak_distribution",
    "load_event_amplitudes",
]
