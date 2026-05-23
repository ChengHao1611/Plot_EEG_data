from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import defaultdict
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import mne
import numpy as np

from eye_movement_distribution import (
    compute_peak_to_shoulder_amplitude,
    save_eye_movement_peak_distribution,
)


@dataclass
class CandidateDiagnostic:
    peak_index: int
    second: int
    peak_time: float
    current_val: float
    left_min_idx: int
    left_min: float
    right_min_idx: int
    right_min: float
    left_rise: float
    right_fall: float
    start_idx: int
    end_idx: int
    start_value: float
    end_value: float
    floor_level: float
    height_thresh: float
    prominence_thresh: float
    window_ticks: int
    broad_window_ticks: int
    is_local_max: bool
    height_ok: bool
    prominence_ok: bool
    floor_ok: bool
    passes: bool
    source: str
    rejection_reasons: list[str]


@dataclass
class ComparisonEvent:
    category: str
    second: int
    diagnostic: CandidateDiagnostic


def get_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Detect eye movements from an EDF file, compare them with manual arousal DAT "
            "labels, and export diagnostic time-domain plots for true/false/miss events."
        )
    )
    parser.add_argument("--file", type=str, help="The path to the EDF file")
    parser.add_argument(
        "--manual-dat",
        type=str,
        help="Optional manual DAT path. Defaults to <edf_stem>_arousal info.dat in the same folder.",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        help="Optional output directory. Defaults to <edf_stem>_arousal_compare under the EDF folder.",
    )
    parser.add_argument(
        "--plot-window-sec",
        type=float,
        default=1.0,
        help="Seconds shown before and after each event in diagnostic plots.",
    )
    parser.add_argument(
        "--height-mad-scale",
        type=float,
        default=0.5,
        help="Height threshold is median + height_mad_scale * MAD.",
    )
    parser.add_argument(
        "--prominence",
        type=float,
        default=0.00006,
        help="Prominence threshold used by the custom detector.",
    )
    return parser.parse_args()


def normalize_user_path(raw_path: str) -> Path:
    cleaned = raw_path.strip().strip('"').strip("'")
    return Path(cleaned).expanduser().resolve()


def parse_numeric_tokens(text: str) -> list[int]:
    matches = re.findall(r"[-+]?\d+(?:\.\d+)?", text)
    return [int(round(float(token))) for token in matches]


def load_dat_seconds(dat_path: Path) -> list[int]:
    content = dat_path.read_text(encoding="utf-8").strip()
    if not content:
        return []

    numbers = parse_numeric_tokens(content)
    if not numbers:
        return []

    if len(numbers) >= 2 and numbers[0] == len(numbers) - 1:
        seconds = numbers[1:]
    elif len(numbers) >= 2 and numbers[0] == 0:
        seconds = numbers[1:]
    else:
        seconds = numbers

    return sorted(int(second) for second in seconds)


def save_dat_seconds(output_path: Path, seconds: Iterable[int]) -> None:
    ordered_seconds = sorted({int(second) for second in seconds})
    values = [str(len(ordered_seconds)), *(str(second) for second in ordered_seconds)]
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(",".join(values) + "\n", encoding="utf-8")


def safe_divide(numerator: int | float, denominator: int | float) -> float | None:
    if denominator == 0:
        return None
    return float(numerator) / float(denominator)


def format_metric(value: float | None) -> str:
    if value is None:
        return "N/A"
    return f"{value:.6f}"


def derive_prefix_from_edf(edf_path: Path) -> str:
    stem = edf_path.stem
    if stem.lower().endswith("_raw"):
        return stem[:-4]
    return stem


def resolve_manual_dat_path(edf_path: Path, manual_override: Path | None = None) -> Path:
    if manual_override is not None:
        return manual_override

    folder = edf_path.parent
    stem = edf_path.stem
    prefix = derive_prefix_from_edf(edf_path)
    candidate_names = [
        f"{stem}_arousal info.dat",
        f"{stem}_raw_arousal info.dat",
        f"{prefix}_arousal info.dat",
        f"{prefix}_raw_arousal info.dat",
    ]

    seen: set[str] = set()
    for candidate_name in candidate_names:
        if candidate_name in seen:
            continue
        seen.add(candidate_name)
        candidate_path = folder / candidate_name
        if candidate_path.exists():
            return candidate_path

    expected_path = folder / f"{stem}_arousal info.dat"
    raise FileNotFoundError(
        f"Manual DAT not found. Expected something like: {expected_path}"
    )


def select_target_channels(raw: mne.io.BaseRaw) -> list[str]:
    return [ch for ch in raw.ch_names if "fp1" in ch.lower() or "fp2" in ch.lower()]


def second_from_time(time_value: float) -> int:
    return max(1, int(np.ceil(float(time_value))))


def build_candidate_diagnostic(
    signal: np.ndarray,
    times: np.ndarray,
    peak_index: int,
    sfreq: float,
    height_thresh: float,
    prominence_thresh: float,
    *,
    source: str,
) -> CandidateDiagnostic:
    window_ticks = max(int(round(sfreq * 0.15)), 1)
    broad_window_ticks = max(int(round(sfreq * 0.35)), 1)

    left_start = max(0, peak_index - window_ticks)
    left_end = peak_index
    right_start = peak_index
    right_end = min(len(signal), peak_index + window_ticks)

    left_window = signal[left_start:left_end]
    right_window = signal[right_start:right_end]

    if left_window.size == 0 or right_window.size == 0:
        raise ValueError("Peak diagnostic cannot be computed at the very edge of the signal.")

    left_offset = int(np.argmin(left_window))
    right_offset = int(np.argmin(right_window))
    left_min_idx = left_start + left_offset
    right_min_idx = right_start + right_offset

    current_val = float(signal[peak_index])
    left_min = float(signal[left_min_idx])
    right_min = float(signal[right_min_idx])
    left_rise = current_val - left_min
    right_fall = current_val - right_min

    start_idx = max(0, peak_index - broad_window_ticks)
    end_idx = min(len(signal), peak_index + broad_window_ticks)
    if end_idx <= start_idx:
        end_idx = min(len(signal), start_idx + 1)

    floor_level = current_val - (float(prominence_thresh) * 0.7)
    start_value = float(signal[start_idx])
    end_value = float(signal[end_idx - 1])

    prev_value = float(signal[peak_index - 1]) if peak_index > 0 else current_val
    next_value = float(signal[peak_index + 1]) if peak_index + 1 < len(signal) else current_val
    current_window = signal[left_start:right_end]
    max_index = left_start + int(np.argmax(current_window))
    is_local_max = (peak_index == max_index)
    height_ok = current_val >= float(height_thresh)
    prominence_ok = left_rise >= float(prominence_thresh) and right_fall >= float(prominence_thresh)
    floor_ok = start_value <= floor_level and end_value <= floor_level
    floor_ok = True
    passes = is_local_max and height_ok and prominence_ok and floor_ok

    rejection_reasons: list[str] = []
    if not is_local_max:
        rejection_reasons.append("not_local_max")
    if not height_ok:
        rejection_reasons.append("below_height_threshold")
    if not prominence_ok:
        rejection_reasons.append("below_prominence_threshold")
    if not floor_ok:
        rejection_reasons.append("floor_not_recovered")

    return CandidateDiagnostic(
        peak_index=int(peak_index),
        second=second_from_time(float(times[peak_index])),
        peak_time=float(times[peak_index]),
        current_val=current_val,
        left_min_idx=int(left_min_idx),
        left_min=left_min,
        right_min_idx=int(right_min_idx),
        right_min=right_min,
        left_rise=float(left_rise),
        right_fall=float(right_fall),
        start_idx=int(start_idx),
        end_idx=int(end_idx),
        start_value=start_value,
        end_value=end_value,
        floor_level=float(floor_level),
        height_thresh=float(height_thresh),
        prominence_thresh=float(prominence_thresh),
        window_ticks=int(window_ticks),
        broad_window_ticks=int(broad_window_ticks),
        is_local_max=bool(is_local_max),
        height_ok=bool(height_ok),
        prominence_ok=bool(prominence_ok),
        floor_ok=bool(floor_ok),
        passes=bool(passes),
        source=source,
        rejection_reasons=rejection_reasons,
    )


def collect_candidate_diagnostics(
    signal: np.ndarray,
    times: np.ndarray,
    sfreq: float,
    height_thresh: float,
    prominence_thresh: float,
) -> tuple[list[CandidateDiagnostic], dict[int, list[CandidateDiagnostic]], dict[int, list[CandidateDiagnostic]]]:
    window_ticks = max(int(round(sfreq * 0.15)), 1)
    passing_candidates: list[CandidateDiagnostic] = []
    passing_by_second: dict[int, list[CandidateDiagnostic]] = defaultdict(list)
    all_candidates_by_second: dict[int, list[CandidateDiagnostic]] = defaultdict(list)

    for peak_index in range(window_ticks, len(signal) - window_ticks):
        current_val = signal[peak_index]
        if current_val <= signal[peak_index - 1] or current_val <= signal[peak_index + 1]:
            continue

        diagnostic = build_candidate_diagnostic(
            signal,
            times,
            peak_index,
            sfreq,
            height_thresh,
            prominence_thresh,
            source="local_max",
        )
        all_candidates_by_second[diagnostic.second].append(diagnostic)
        if diagnostic.passes:
            passing_candidates.append(diagnostic)
            passing_by_second[diagnostic.second].append(diagnostic)

    return passing_candidates, passing_by_second, all_candidates_by_second


def choose_best_diagnostic(candidates: list[CandidateDiagnostic]) -> CandidateDiagnostic:
    return max(
        candidates,
        key=lambda candidate: (
            float(candidate.current_val),
            float(candidate.left_rise + candidate.right_fall),
            -abs(float(candidate.peak_time) - float(candidate.second)),
        ),
    )


def build_second_fallback_diagnostic(
    signal: np.ndarray,
    times: np.ndarray,
    sample_seconds: np.ndarray,
    second: int,
    sfreq: float,
    height_thresh: float,
    prominence_thresh: float,
) -> CandidateDiagnostic | None:
    indices = np.flatnonzero(sample_seconds == int(second))
    if indices.size == 0:
        return None

    relative_index = int(np.argmax(signal[indices]))
    peak_index = int(indices[relative_index])

    if peak_index <= 0:
        peak_index = 1
    if peak_index >= len(signal) - 1:
        peak_index = len(signal) - 2
    if peak_index <= 0 or peak_index >= len(signal) - 1:
        return None

    return build_candidate_diagnostic(
        signal,
        times,
        peak_index,
        sfreq,
        height_thresh,
        prominence_thresh,
        source="second_fallback",
    )


def compare_seconds(
    predicted_seconds: Iterable[int],
    manual_seconds: Iterable[int],
    total_seconds: int,
) -> dict[str, int | float | None | list[int]]:
    predicted_set = {int(second) for second in predicted_seconds}
    manual_set = {int(second) for second in manual_seconds}

    true_seconds = sorted(predicted_set & manual_set)
    false_seconds = sorted(predicted_set - manual_set)
    miss_seconds = sorted(manual_set - predicted_set)

    true_count = len(true_seconds)
    false_count = len(false_seconds)
    miss_count = len(miss_seconds)
    true_negative = max(int(total_seconds) - true_count - false_count - miss_count, 0)

    accuracy = safe_divide(true_count, true_count + false_count + miss_count)
    precision = safe_divide(true_count, true_count + false_count)
    recall = safe_divide(true_count, true_count + miss_count)

    return {
        "true_seconds": true_seconds,
        "false_seconds": false_seconds,
        "miss_seconds": miss_seconds,
        "true_count": true_count,
        "false_count": false_count,
        "miss_count": miss_count,
        "true_negative": true_negative,
        "accuracy": accuracy,
        "precision": precision,
        "recall": recall,
    }


def clamp_seconds(seconds: Iterable[int], total_seconds: int) -> tuple[list[int], list[int]]:
    kept: list[int] = []
    dropped: list[int] = []
    for second in seconds:
        second_int = int(second)
        if 1 <= second_int <= int(total_seconds):
            kept.append(second_int)
        else:
            dropped.append(second_int)
    return sorted(set(kept)), sorted(set(dropped))


def plot_event(
    signal: np.ndarray,
    times: np.ndarray,
    diagnostic: CandidateDiagnostic,
    category: str,
    output_path: Path,
    plot_window_sec: float,
) -> None:
    plot_half_width = max(float(plot_window_sec), 0.1)
    center_time = float(diagnostic.peak_time)
    start_time = center_time - plot_half_width
    end_time = center_time + plot_half_width
    mask = (times >= start_time) & (times <= end_time)
    if not np.any(mask):
        mask[max(diagnostic.peak_index - 1, 0) : min(diagnostic.peak_index + 2, len(signal))] = True

    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(times[mask], signal[mask], color="navy", linewidth=1.5, label="combined_signal")

    peak_time = float(times[diagnostic.peak_index])
    left_time = float(times[diagnostic.left_min_idx])
    right_time = float(times[diagnostic.right_min_idx])
    start_time_value = float(times[diagnostic.start_idx])
    end_time_value = float(times[diagnostic.end_idx - 1])

    ax.scatter([peak_time], [diagnostic.current_val], color="red", s=40, zorder=5, label="peak")
    ax.scatter([left_time], [diagnostic.left_min], color="green", s=35, zorder=5, label="left_min")
    ax.scatter([right_time], [diagnostic.right_min], color="orange", s=35, zorder=5, label="right_min")
    ax.scatter(
        [start_time_value],
        [diagnostic.start_value],
        color="purple",
        s=35,
        zorder=5,
        label="signal[start_idx]",
    )
    ax.scatter(
        [end_time_value],
        [diagnostic.end_value],
        color="brown",
        s=35,
        zorder=5,
        label="signal[end_idx-1]",
    )
    ax.axhline(diagnostic.floor_level, color="black", linestyle="--", linewidth=1.2, label="floor_level")
    ax.plot(
        [left_time, peak_time],
        [diagnostic.left_min, diagnostic.current_val],
        color="green",
        linewidth=2.0,
        label="left_rise",
    )
    ax.plot(
        [peak_time, right_time],
        [diagnostic.current_val, diagnostic.right_min],
        color="orange",
        linewidth=2.0,
        label="right_fall",
    )
    ax.axvline(peak_time, color="red", linestyle=":", linewidth=1.0)

    info_lines = [
        f"category={category}",
        f"second={diagnostic.second}",
        f"peak_index={diagnostic.peak_index}",
        f"peak_time={diagnostic.peak_time:.4f}s",
        f"left_rise={diagnostic.left_rise:.8f}",
        f"right_fall={diagnostic.right_fall:.8f}",
        f"floor_level={diagnostic.floor_level:.8f}",
        f"signal[start_idx]={diagnostic.start_value:.8f} @ {diagnostic.start_idx}",
        f"signal[end_idx-1]={diagnostic.end_value:.8f} @ {diagnostic.end_idx - 1}",
        f"height_thresh={diagnostic.height_thresh:.8f}",
        f"prominence_thresh={diagnostic.prominence_thresh:.8f}",
        f"passes={diagnostic.passes}",
        f"source={diagnostic.source}",
    ]
    if diagnostic.rejection_reasons:
        info_lines.append(f"reasons={','.join(diagnostic.rejection_reasons)}")
    else:
        info_lines.append("reasons=pass")

    ax.text(
        0.02,
        0.98,
        "\n".join(info_lines),
        transform=ax.transAxes,
        va="top",
        ha="left",
        fontsize=9,
        bbox={"facecolor": "white", "alpha": 0.85, "edgecolor": "gray"},
    )

    ax.set_title(f"{category.upper()} second {diagnostic.second}")
    ax.set_xlabel("Time (s)")
    ax.set_ylabel("Combined |FP1|/|FP2| amplitude")
    ax.grid(True, alpha=0.25)
    ax.legend(loc="upper right")
    fig.tight_layout()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=150)
    plt.close(fig)


def save_event_summary_csv(output_path: Path, events: list[ComparisonEvent]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "category",
        "second",
        "peak_index",
        "peak_time",
        "current_val",
        "left_min_idx",
        "left_min",
        "right_min_idx",
        "right_min",
        "left_rise",
        "right_fall",
        "start_idx",
        "end_idx",
        "start_value",
        "end_value",
        "floor_level",
        "height_thresh",
        "prominence_thresh",
        "is_local_max",
        "height_ok",
        "prominence_ok",
        "floor_ok",
        "passes",
        "source",
        "rejection_reasons",
    ]

    with output_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        for event in events:
            row = asdict(event.diagnostic)
            row["category"] = event.category
            row["second"] = event.second
            row["rejection_reasons"] = ",".join(event.diagnostic.rejection_reasons)
            writer.writerow({field: row.get(field) for field in fieldnames})


def save_metrics_report(
    output_path: Path,
    *,
    edf_path: Path,
    manual_dat_path: Path,
    auto_dat_path: Path,
    output_dir: Path,
    total_seconds: int,
    predicted_seconds: list[int],
    manual_seconds: list[int],
    dropped_manual_seconds: list[int],
    metrics: dict[str, int | float | None | list[int]],
) -> None:
    lines = [
        f"EDF: {edf_path}",
        f"Manual DAT: {manual_dat_path}",
        f"Auto DAT: {auto_dat_path}",
        f"Output dir: {output_dir}",
        f"Total seconds: {total_seconds}",
        f"Predicted seconds ({len(predicted_seconds)}): {predicted_seconds}",
        f"Manual seconds ({len(manual_seconds)}): {manual_seconds}",
        f"true ({metrics['true_count']}): {metrics['true_seconds']}",
        f"false ({metrics['false_count']}): {metrics['false_seconds']}",
        f"miss ({metrics['miss_count']}): {metrics['miss_seconds']}",
        f"true_negative: {metrics['true_negative']}",
        f"accuracy: {format_metric(metrics['accuracy'])}",
        f"precision: {format_metric(metrics['precision'])}",
        f"recall: {format_metric(metrics['recall'])}",
    ]
    if dropped_manual_seconds:
        lines.append(f"dropped_manual_seconds: {dropped_manual_seconds}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def detect_eye_movements(
    raw: mne.io.BaseRaw,
    target_channels: list[str],
    *,
    height_mad_scale: float,
    prominence_thresh: float,
) -> dict[str, object]:
    picked_raw = raw.copy().pick(target_channels[:2])
    data, times = picked_raw.get_data(), picked_raw.times
    sfreq = float(picked_raw.info["sfreq"])
    #combined_signal = (np.abs(data[0]) + np.abs(data[1])) / 2.0
    combined_signal = data[1]

    median = float(np.median(combined_signal))
    mad = float(np.median(np.abs(combined_signal - median)))
    height_thresh = median + float(height_mad_scale) * mad

    passing_candidates, passing_by_second, all_candidates_by_second = collect_candidate_diagnostics(
        combined_signal,
        times,
        sfreq,
        height_thresh,
        prominence_thresh,
    )

    predicted_seconds = sorted(passing_by_second.keys())
    sample_seconds = np.maximum(1, np.ceil(times).astype(int))
    total_seconds = int(sample_seconds[-1]) if sample_seconds.size else 0

    return {
        "signal": combined_signal,
        "times": times,
        "sfreq": sfreq,
        "sample_seconds": sample_seconds,
        "height_thresh": height_thresh,
        "prominence_thresh": float(prominence_thresh),
        "passing_candidates": passing_candidates,
        "passing_by_second": passing_by_second,
        "all_candidates_by_second": all_candidates_by_second,
        "predicted_seconds": predicted_seconds,
        "total_seconds": total_seconds,
    }


def build_comparison_events(
    comparison: dict[str, int | float | None | list[int]],
    detection: dict[str, object],
) -> list[ComparisonEvent]:
    signal = detection["signal"]
    times = detection["times"]
    sample_seconds = detection["sample_seconds"]
    sfreq = float(detection["sfreq"])
    height_thresh = float(detection["height_thresh"])
    prominence_thresh = float(detection["prominence_thresh"])
    passing_by_second = detection["passing_by_second"]
    all_candidates_by_second = detection["all_candidates_by_second"]

    events: list[ComparisonEvent] = []

    for category_key, category_name in (
        ("false_seconds", "false"),
        ("miss_seconds", "miss"),
        ("true_seconds", "true"),
    ):
        for second in comparison[category_key]:
            second_int = int(second)
            if category_name in {"true", "false"}:
                candidates = list(passing_by_second.get(second_int, []))
            else:
                candidates = list(all_candidates_by_second.get(second_int, []))

            if candidates:
                diagnostic = choose_best_diagnostic(candidates)
            else:
                diagnostic = build_second_fallback_diagnostic(
                    signal,
                    times,
                    sample_seconds,
                    second_int,
                    sfreq,
                    height_thresh,
                    prominence_thresh,
                )
                if diagnostic is None:
                    continue

            events.append(
                ComparisonEvent(
                    category=category_name,
                    second=second_int,
                    diagnostic=diagnostic,
                )
            )

    return events


def main() -> int:
    args = get_args()
    raw_file = args.file
    if raw_file is None:
        print("Hint: The --file parameter was not detected.")
        raw_file = input("Please enter the path to the EDF file: ")

    edf_path = normalize_user_path(raw_file)
    if not edf_path.exists():
        print(f"Error: The file '{edf_path}' does not exist.")
        return 1

    manual_dat_override = normalize_user_path(args.manual_dat) if args.manual_dat else None
    try:
        manual_dat_path = resolve_manual_dat_path(edf_path, manual_dat_override)
    except FileNotFoundError as exc:
        print(f"Error: {exc}")
        return 1

    output_dir = (
        normalize_user_path(args.output_dir)
        if args.output_dir
        else (edf_path.parent / f"{edf_path.stem}_arousal_compare").resolve()
    )
    auto_dat_path = output_dir / f"{edf_path.stem}_auto_eye_movement.dat"
    event_summary_csv = output_dir / "event_summary.csv"
    distribution_output = output_dir / "eye_movement_peak_distribution.png"
    metrics_report_path = output_dir / "metrics_report.txt"

    raw = mne.io.read_raw_edf(str(edf_path), preload=True, verbose=False)
    target_channels = select_target_channels(raw)
    if len(target_channels) < 2:
        print("找不到 FP1 或 FP2 通道")
        return 1

    detection = detect_eye_movements(
        raw,
        target_channels,
        height_mad_scale=float(args.height_mad_scale),
        prominence_thresh=float(args.prominence),
    )

    predicted_seconds = list(detection["predicted_seconds"])
    total_seconds = int(detection["total_seconds"])
    predicted_seconds, dropped_predicted_seconds = clamp_seconds(predicted_seconds, total_seconds)
    save_dat_seconds(auto_dat_path, predicted_seconds)

    manual_seconds_raw = load_dat_seconds(manual_dat_path)
    manual_seconds, dropped_manual_seconds = clamp_seconds(manual_seconds_raw, total_seconds)
    comparison = compare_seconds(predicted_seconds, manual_seconds, total_seconds)
    events = build_comparison_events(comparison, detection)
    positive_amplitudes = [
        amplitude
        for event in events
        for amplitude in [
            compute_peak_to_shoulder_amplitude(
                event.diagnostic.current_val,
                event.diagnostic.left_min,
                event.diagnostic.right_min,
            )
        ]
        if amplitude is not None
        if event.category in {"true", "miss"}
    ]
    false_amplitudes = [
        amplitude
        for event in events
        for amplitude in [
            compute_peak_to_shoulder_amplitude(
                event.diagnostic.current_val,
                event.diagnostic.left_min,
                event.diagnostic.right_min,
            )
        ]
        if amplitude is not None
        if event.category == "false"
    ]
    positive_count, false_count = save_eye_movement_peak_distribution(
        distribution_output,
        positive_amplitudes,
        false_amplitudes,
    )

    for event in events:
        plot_path = output_dir / event.category / f"{event.category}_second_{event.second:04d}.png"
        plot_event(
            detection["signal"],
            detection["times"],
            event.diagnostic,
            event.category,
            plot_path,
            float(args.plot_window_sec),
        )

    save_event_summary_csv(event_summary_csv, events)
    save_metrics_report(
        metrics_report_path,
        edf_path=edf_path,
        manual_dat_path=manual_dat_path,
        auto_dat_path=auto_dat_path,
        output_dir=output_dir,
        total_seconds=total_seconds,
        predicted_seconds=predicted_seconds,
        manual_seconds=manual_seconds,
        dropped_manual_seconds=dropped_manual_seconds,
        metrics=comparison,
    )

    print(f"EDF: {edf_path}")
    print(f"Manual DAT: {manual_dat_path}")
    print(f"Output dir: {output_dir}")
    print(f"Auto DAT: {auto_dat_path}")
    print(f"Event summary CSV: {event_summary_csv}")
    print(f"Distribution graph: {distribution_output}")
    print(f"Metrics report: {metrics_report_path}")
    print(f"Channels used: {target_channels[:2]}")
    print(f"Total seconds: {total_seconds}")
    print(f"Predicted seconds: {len(predicted_seconds)}")
    print(f"Manual seconds: {len(manual_seconds)}")
    print(f"(true + miss) amplitude count: {positive_count}")
    print(f"false amplitude count: {false_count}")
    print(f"true: {comparison['true_count']}")
    print(f"false: {comparison['false_count']}")
    print(f"miss: {comparison['miss_count']}")
    print(f"true_negative: {comparison['true_negative']}")
    print(f"accuracy: {format_metric(comparison['accuracy'])}")
    print(f"precision: {format_metric(comparison['precision'])}")
    print(f"recall: {format_metric(comparison['recall'])}")

    if dropped_predicted_seconds:
        print(
            "Warning: predicted seconds outside the EDF duration were ignored: "
            f"{dropped_predicted_seconds}"
        )
    if dropped_manual_seconds:
        print(
            "Warning: manual DAT contains seconds outside the EDF duration and they were ignored: "
            f"{dropped_manual_seconds}"
        )

    return 0


if __name__ == "__main__":
    sys.exit(main())
