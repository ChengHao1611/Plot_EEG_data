from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Sequence

import numpy as np
import pyedflib
from scipy.signal import butter, sosfiltfilt

from alpha_detection_distribution import save_alpha_energy_normal_distribution
from alpha_detection_paths import default_plot_dir, normalize_user_path, resolve_input_paths
from alpha_detection_plotting import export_classified_plots


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Standalone alpha detection workflow: find raw EDF/alpha DAT/eye DAT in a folder, "
            "compute per-second FFT outputs, compare labeled seconds, and export true/false/miss plots."
        )
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        help="Folder containing one *_raw.EDF file, or a direct raw EDF path.",
    )
    parser.add_argument("--channel", default="FP2", help="Channel name to match (default: FP2)")
    parser.add_argument(
        "--plot-dir",
        help="Optional output folder for true/false/miss plots.",
    )
    parser.add_argument(
        "--normal-dist-output",
        help="Optional PNG path for the (true+miss) vs false alpha energy normal distribution graph.",
    )
    parser.add_argument(
        "--alpha-threshold",
        type=float,
        default=1500.0,
        help="A second is predicted positive when alpha_peak == 1 and alpha_amplitude > this threshold.",
    )
    parser.add_argument(
        "--nfft-pow2",
        action="store_true",
        help="Use FFT length as the next power of 2 >= fs. Default uses nfft=fs.",
    )
    parser.add_argument(
        "--highpass-1",
        action="store_true",
        help="Apply a zero-phase 1 Hz high-pass filter before FFT and plotting.",
    )
    parser.add_argument(
        "--lowpass-30",
        action="store_true",
        help="Apply a zero-phase 30 Hz low-pass filter before FFT and plotting. If both filters are enabled, high-pass runs first.",
    )
    parser.add_argument("--theta-low", type=float, default=4.0, help="Theta band low cutoff (Hz)")
    parser.add_argument("--theta-high", type=float, default=7.8, help="Theta band high cutoff (Hz)")
    parser.add_argument("--alpha-low", type=float, default=7.8, help="Alpha band low cutoff (Hz)")
    parser.add_argument("--alpha-high", type=float, default=12.5, help="Alpha band high cutoff (Hz)")
    parser.add_argument("--beta-low", type=float, default=12.5, help="Beta band low cutoff (Hz)")
    parser.add_argument("--beta-high", type=float, default=30.0, help="Beta band high cutoff (Hz)")
    parser.add_argument(
        "--peak-low",
        type=float,
        default=1.9,
        help="Peak search low cutoff (Hz).",
    )
    parser.add_argument(
        "--peak-high",
        type=float,
        default=30.0,
        help="Peak search high cutoff (Hz).",
    )
    parser.add_argument(
        "--max-plot-freq",
        type=float,
        default=100.0,
        help="Maximum frequency shown in the spectrum plot (Hz).",
    )
    return parser.parse_args()


def parse_numeric_tokens(text: str) -> list[int]:
    matches = re.findall(r"[-+]?\d+(?:\.\d+)?", text)
    return [int(round(float(token))) for token in matches]


def load_dat_seconds(dat_path: Path | None) -> list[int]:
    if dat_path is None or not dat_path.exists():
        return []

    content = dat_path.read_text(encoding="utf-8").strip()
    if not content:
        return []

    numbers = parse_numeric_tokens(content)
    if not numbers:
        return []

    seconds = numbers[1:]
    seconds.sort()
    return seconds


def find_channel_index(labels: Sequence[str], target: str) -> int | None:
    target_lower = target.lower()
    for index, label in enumerate(labels):
        if target_lower in label.lower():
            return index
    return None


def load_channel_signal(edf_path: Path, channel_name: str) -> tuple[np.ndarray, float]:
    if not edf_path.exists():
        raise FileNotFoundError(f"EDF not found: {edf_path}")

    reader = pyedflib.EdfReader(str(edf_path))
    try:
        labels = reader.getSignalLabels()
        channel_index = find_channel_index(labels, channel_name)
        if channel_index is None:
            raise ValueError(f"Channel '{channel_name}' not found in {edf_path}")

        signal = reader.readSignal(channel_index).astype(float)
        fs = float(reader.getSampleFrequency(channel_index))
    finally:
        reader.close()

    return signal, fs


def next_power_of_2(n: int) -> int:
    if n <= 1:
        return 1
    return 1 << (n - 1).bit_length()


def apply_highpass_filter_1(signal: np.ndarray, fs: float, *, order: int = 4) -> np.ndarray:
    if fs <= 0:
        raise ValueError(f"Invalid sample rate for filtering: {fs}")

    nyquist = float(fs) / 2.0
    cutoff = 1.0
    if cutoff >= nyquist:
        raise ValueError(
            f"Cannot apply 1 Hz high-pass filter when Nyquist is {nyquist:.6f} Hz. "
            f"Sample rate must be greater than 2 Hz."
        )

    sos = butter(
        int(order),
        float(cutoff),
        btype="highpass",
        fs=float(fs),
        output="sos",
    )
    return sosfiltfilt(sos, np.asarray(signal, dtype=float))


def apply_lowpass_filter_30(signal: np.ndarray, fs: float, *, order: int = 4) -> np.ndarray:
    if fs <= 0:
        raise ValueError(f"Invalid sample rate for filtering: {fs}")

    nyquist = float(fs) / 2.0
    cutoff = 30.0
    if cutoff >= nyquist:
        raise ValueError(
            f"Cannot apply 30 Hz low-pass filter when Nyquist is {nyquist:.6f} Hz. "
            f"Sample rate must be greater than 60 Hz."
        )

    sos = butter(
        int(order),
        float(cutoff),
        btype="lowpass",
        fs=float(fs),
        output="sos",
    )
    return sosfiltfilt(sos, np.asarray(signal, dtype=float))


def safe_divide(numerator: int, denominator: int) -> float | None:
    if denominator == 0:
        return None
    return numerator / denominator


def summarize_metrics(true_count: int, false_count: int, miss_count: int) -> dict[str, float | None]:
    precision = safe_divide(true_count, true_count + false_count)
    accuracy = safe_divide(true_count, true_count + false_count + miss_count)
    recall = safe_divide(true_count, true_count + miss_count)

    f1_score: float | None = None
    if precision is not None and recall is not None:
        denominator = precision + recall
        f1_score = 0.0 if denominator == 0 else (2.0 * precision * recall / denominator)

    return {
        "precision": precision,
        "accuracy": accuracy,
        "recall": recall,
        "f1_score": f1_score,
    }


def format_metric(value: float | None) -> str:
    if value is None:
        return "N/A"
    return f"{value:.6f}"


def default_normal_dist_output(plot_dir: Path, prefix: str) -> Path:
    return plot_dir / f"{prefix}_alpha_energy_normal_distribution.png"


def compute_band_powers_and_ratios_fft(
    signal: np.ndarray,
    fs: float,
    *,
    theta_low: float,
    theta_high: float,
    alpha_low: float,
    alpha_high: float,
    beta_low: float,
    beta_high: float,
    peak_low: float,
    peak_high: float,
    nfft: int | None = None,
) -> tuple[np.ndarray, np.ndarray, list[np.ndarray]]:
    del theta_low, theta_high, beta_low, beta_high

    if signal.size == 0 or fs <= 0:
        return np.array([], dtype=float), np.array([], dtype=float), []

    win = int(round(fs))
    if win <= 0:
        return np.array([], dtype=float), np.array([], dtype=float), []

    nfft_final = win if nfft is None else int(nfft)
    if nfft_final < win:
        nfft_final = win

    n_secs = int(signal.size // win)
    if n_secs == 0:
        return np.array([], dtype=float), np.array([], dtype=float), []

    freqs = np.fft.rfftfreq(nfft_final, d=1.0 / fs)
    peak_mask = (freqs >= peak_low) & (freqs <= peak_high)
    if not np.any(peak_mask):
        return np.array([], dtype=float), np.array([], dtype=float), []

    alpha_amplitude = np.full(n_secs, np.nan, dtype=float)
    alpha_peak = np.full(n_secs, np.nan, dtype=float)
    power_segments: list[np.ndarray] = []

    for second_index in range(n_secs):
        start = second_index * win
        end = start + win
        segment = signal[start:end]

        if segment.size < win or np.isnan(segment).any():
            power_segments.append(np.array([], dtype=float))
            continue

        spectrum = np.fft.rfft(segment, n=nfft_final)
        power = np.abs(spectrum).astype(float, copy=False)
        power_segments.append(power.copy())

        peak_rel_index = int(np.argmax(power[peak_mask]))
        peak_abs_index = int(np.flatnonzero(peak_mask)[peak_rel_index])
        peak_freq = float(freqs[peak_abs_index])

        alpha_amplitude[second_index] = float(power[peak_abs_index])
        alpha_peak[second_index] = 1.0 if alpha_low <= peak_freq < alpha_high else 0.0

    return alpha_amplitude, alpha_peak, power_segments


def build_classified_row(
    second: int,
    label: str,
    alpha_amplitude: np.ndarray,
    alpha_peak: np.ndarray,
    power_segments: Sequence[np.ndarray],
) -> dict[str, object]:
    index = second - 1
    amplitude_value: float | None = None
    peak_value: int | None = None
    power_value = np.array([], dtype=float)

    if 0 <= index < alpha_amplitude.size:
        raw_value = float(alpha_amplitude[index])
        if np.isfinite(raw_value):
            amplitude_value = raw_value

    if 0 <= index < alpha_peak.size:
        raw_peak = float(alpha_peak[index])
        if np.isfinite(raw_peak):
            peak_value = int(round(raw_peak))

    if 0 <= index < len(power_segments):
        power_value = np.asarray(power_segments[index], dtype=float)

    return {
        "second": second,
        "label": label,
        "alpha_amplitude": amplitude_value,
        "alpha_peak": peak_value,
        "power": power_value,
    }


def compare_seconds(
    alpha_dat_path: Path,
    eye_dat_path: Path | None,
    alpha_amplitude: np.ndarray,
    alpha_peak: np.ndarray,
    power_segments: Sequence[np.ndarray],
    *,
    alpha_threshold: float,
) -> tuple[int, int, int, list[dict[str, object]]]:
    dat_seconds = load_dat_seconds(alpha_dat_path)
    eye_dat_seconds = set(load_dat_seconds(eye_dat_path))

    predicted_seconds: list[int] = []
    for index in range(min(alpha_amplitude.size, alpha_peak.size)):
        second = index + 1
        peak_value = float(alpha_peak[index])
        amplitude_value = float(alpha_amplitude[index])
        if not np.isfinite(peak_value) or not np.isfinite(amplitude_value):
            continue
        if second in eye_dat_seconds:
        #if second in eye_dat_seconds or (second - 1 in eye_dat_seconds and second + 1 in eye_dat_seconds):
            continue
        if int(round(peak_value)) == 1 and amplitude_value > float(alpha_threshold):
            predicted_seconds.append(second)

    predicted_seconds.sort()
    dat_seconds.sort()

    pointer = 0
    true_count = 0
    false_count = 0
    miss_count = 0
    result_rows: list[dict[str, object]] = []

    for second in predicted_seconds:
        while pointer < len(dat_seconds) and dat_seconds[pointer] < second:
            miss_second = dat_seconds[pointer]
            result_rows.append(
                build_classified_row(
                    miss_second,
                    "miss",
                    alpha_amplitude,
                    alpha_peak,
                    power_segments,
                )
            )
            miss_count += 1
            pointer += 1

        if pointer < len(dat_seconds) and dat_seconds[pointer] == second:
            result_rows.append(
                build_classified_row(
                    second,
                    "true",
                    alpha_amplitude,
                    alpha_peak,
                    power_segments,
                )
            )
            true_count += 1
            pointer += 1
        else:
            result_rows.append(
                build_classified_row(
                    second,
                    "false",
                    alpha_amplitude,
                    alpha_peak,
                    power_segments,
                )
            )
            false_count += 1

    while pointer < len(dat_seconds):
        miss_second = dat_seconds[pointer]
        result_rows.append(
            build_classified_row(
                miss_second,
                "miss",
                alpha_amplitude,
                alpha_peak,
                power_segments,
            )
        )
        miss_count += 1
        pointer += 1

    return true_count, false_count, miss_count, result_rows


def main() -> int:
    args = parse_args()
    raw_input_path = args.input_path
    if not raw_input_path:
        raw_input_path = input("請輸入資料夾或 raw EDF 路徑: ")

    input_path = normalize_user_path(raw_input_path)
    folder, edf_path, prefix, alpha_dat_path, eye_dat_path = resolve_input_paths(input_path)
    plot_dir = normalize_user_path(args.plot_dir) if args.plot_dir else default_plot_dir(folder, prefix)
    normal_dist_output = (
        normalize_user_path(args.normal_dist_output)
        if args.normal_dist_output
        else default_normal_dist_output(plot_dir, prefix)
    )

    signal, fs = load_channel_signal(edf_path, args.channel)
    if bool(args.highpass_1):
        signal = apply_highpass_filter_1(signal, fs)
    if bool(args.lowpass_30):
        signal = apply_lowpass_filter_30(signal, fs)
    window = int(round(fs))
    if window <= 0:
        raise ValueError(f"Invalid sample rate: {fs}")

    nfft_value = next_power_of_2(window) if bool(args.nfft_pow2) else window
    alpha_amplitude, alpha_peak, power_segments = compute_band_powers_and_ratios_fft(
        signal,
        fs,
        theta_low=args.theta_low,
        theta_high=args.theta_high,
        alpha_low=args.alpha_low,
        alpha_high=args.alpha_high,
        beta_low=args.beta_low,
        beta_high=args.beta_high,
        peak_low=args.peak_low,
        peak_high=args.peak_high,
        nfft=nfft_value,
    )

    true_count, false_count, miss_count, result_rows = compare_seconds(
        alpha_dat_path,
        eye_dat_path,
        alpha_amplitude,
        alpha_peak,
        power_segments,
        alpha_threshold=float(args.alpha_threshold),
    )
    metrics = summarize_metrics(true_count, false_count, miss_count)

    plotted_true, plotted_false, plotted_miss, warnings = export_classified_plots(
        result_rows,
        output_dir=plot_dir,
        signal=signal,
        fs=fs,
        nfft=nfft_value,
        alpha_low=float(args.alpha_low),
        alpha_high=float(args.alpha_high),
        max_plot_freq=float(args.max_plot_freq),
    )
    normal_true_miss_count, normal_false_count = save_alpha_energy_normal_distribution(
        normal_dist_output,
        result_rows,
    )

    print(f"Folder: {folder}")
    print(f"EDF: {edf_path}")
    print(f"Alpha DAT: {alpha_dat_path}")
    print(f"Eye DAT: {eye_dat_path if eye_dat_path else 'None'}")
    print(f"Channel: {args.channel}")
    print(f"fs: {fs}")
    print(f"highpass_1: {bool(args.highpass_1)}")
    print(f"lowpass_30: {bool(args.lowpass_30)}")
    print(f"nfft: {nfft_value}")
    print(f"Total seconds: {len(power_segments)}")
    print(f"true: {true_count}")
    print(f"false: {false_count}")
    print(f"miss: {miss_count}")
    print(f"precision: {format_metric(metrics['precision'])}")
    print(f"accuracy: {format_metric(metrics['accuracy'])}")
    print(f"recall: {format_metric(metrics['recall'])}")
    print(f"f1-score: {format_metric(metrics['f1_score'])}")
    print(f"Plot dir: {plot_dir}")
    print(f"Normal dist graph: {normal_dist_output}")
    print(f"Normal dist true+miss count: {normal_true_miss_count}")
    print(f"Normal dist false count: {normal_false_count}")
    print(f"Saved true plots: {plotted_true}")
    print(f"Saved false plots: {plotted_false}")
    print(f"Saved miss plots: {plotted_miss}")

    if warnings:
        print("--- Plot warnings ---")
        for message in warnings:
            print(message)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
