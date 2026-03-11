import argparse
import csv
import os
from typing import List, Tuple

import numpy as np
import pyedflib

def find_channel_index(labels: List[str], target: str) -> int | None:
    target_lower = target.lower()
    for i, label in enumerate(labels):
        if target_lower in label.lower():
            return i
    return None

def list_edf_files(input_dir: str) -> List[str]:
    if not os.path.isdir(input_dir):
        raise NotADirectoryError(f"Not a directory: {input_dir}")

    files: List[str] = []
    for name in os.listdir(input_dir):
        if name.lower().endswith(".edf"):
            files.append(os.path.join(input_dir, name))

    files.sort()
    return files


def load_channel_signal(edf_path: str, channel_name: str) -> Tuple[np.ndarray, float]:
    if not os.path.exists(edf_path):
        raise FileNotFoundError(f"EDF not found: {edf_path}")

    f = pyedflib.EdfReader(edf_path)
    try:
        labels = f.getSignalLabels()
        ch_idx = find_channel_index(labels, channel_name)
        if ch_idx is None:
            raise ValueError(f"Channel '{channel_name}' not found in {edf_path}")

        signal = f.readSignal(ch_idx).astype(float)
        fs = float(f.getSampleFrequency(ch_idx))
    finally:
        f.close()

    return signal, fs


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
) -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    if signal.size == 0 or fs <= 0:
        empty = np.array([], dtype=float)
        return empty, empty, empty, empty, empty

    win = int(round(fs))
    if win <= 0:
        empty = np.array([], dtype=float)
        return empty, empty, empty, empty, empty

    n_secs = int(signal.size // win)
    if n_secs == 0:
        empty = np.array([], dtype=float)
        return empty, empty, empty, empty, empty

    freqs = np.fft.rfftfreq(win, d=1.0 / fs)
    theta_mask = (freqs >= theta_low) & (freqs < theta_high)
    alpha_mask = (freqs >= alpha_low) & (freqs < alpha_high)
    beta_mask = (freqs >= beta_low) & (freqs <= beta_high)

    if not np.any(theta_mask) or not np.any(alpha_mask) or not np.any(beta_mask):
        empty = np.array([], dtype=float)
        return empty, empty, empty, empty, empty

    window = np.hanning(win)
    theta_power = np.full(n_secs, np.nan, dtype=float)
    alpha_power = np.full(n_secs, np.nan, dtype=float)
    beta_power = np.full(n_secs, np.nan, dtype=float)
    alpha_beta = np.full(n_secs, np.nan, dtype=float)
    alpha_theta = np.full(n_secs, np.nan, dtype=float)
    alpha_total = np.full(n_secs, np.nan, dtype=float)


    for s in range(n_secs):
        start = s * win
        end = start + win
        seg = signal[start:end]
        if seg.size < win:
            continue
        if np.isnan(seg).any():
            continue
        seg = seg - float(np.mean(seg))
        seg = seg * window
        spec = np.fft.rfft(seg)
        power = np.abs(spec) ** 2

        p_theta = float(np.sum(power[theta_mask]))
        p_alpha = float(np.sum(power[alpha_mask]))
        p_beta = float(np.sum(power[beta_mask]))

        theta_power[s] = p_theta
        alpha_power[s] = p_alpha
        beta_power[s] = p_beta

        if p_beta > 0.0:
            alpha_beta[s] = p_alpha / p_beta
        if p_theta > 0.0:
            alpha_theta[s] = p_alpha / p_theta
        
        alpha_total[s] = p_alpha / (p_alpha + p_beta + p_theta)

    return theta_power, alpha_power, beta_power, alpha_beta, alpha_theta, alpha_total


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Compute per-second theta/alpha/beta FFT band power and "
            "alpha/beta, alpha/theta ratios for FP2 channel."
        )
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--edf", nargs="+", help="EDF file paths")

    parser.add_argument("--channel", default="FP2", help="Channel name to match (default: FP2)")

    parser.add_argument("--theta-low", type=float, default=4.0, help="Theta band low cutoff (Hz)")
    parser.add_argument("--theta-high", type=float, default=8.0, help="Theta band high cutoff (Hz)")
    parser.add_argument("--alpha-low", type=float, default=8.0, help="Alpha band low cutoff (Hz)")
    parser.add_argument("--alpha-high", type=float, default=12.0, help="Alpha band high cutoff (Hz)")
    parser.add_argument("--beta-low", type=float, default=12.0, help="Beta band low cutoff (Hz)")
    parser.add_argument("--beta-high", type=float, default=30.0, help="Beta band high cutoff (Hz)")

    parser.add_argument("--save-csv", default=None, help="Optional CSV output path")

    return parser.parse_args()


def save_csv(output_path: str, rows: List[Tuple[str, int, float, float, float, float, float]]) -> None:
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "file",
                "second",
                "theta_power",
                "alpha_power",
                "beta_power",
                "alpha_beta",
                "alpha_theta",
                "alpha_total"
            ]
        )
        writer.writerows(rows)


def main() -> int:
    args = parse_args()
    edf_files = []

    if args.edf:
        edf_files = [os.path.abspath(p) for p in args.edf]

    if len(edf_files) == 0:
        print("No EDF files to process.")
        return 1

    total_secs = 0
    total_valid = 0
    total_skipped = 0
    csv_rows: List[Tuple[str, int, float, float, float, float, float, float]] = []

    for path in edf_files:
        signal, fs = load_channel_signal(path, args.channel)
        theta_power, alpha_power, beta_power, alpha_beta, alpha_theta, alpha_total = compute_band_powers_and_ratios_fft(
            signal,
            fs,
            theta_low=args.theta_low,
            theta_high=args.theta_high,
            alpha_low=args.alpha_low,
            alpha_high=args.alpha_high,
            beta_low=args.beta_low,
            beta_high=args.beta_high,
        )

        n_secs = int(theta_power.size)
        valid_mask = ~(np.isnan(theta_power) | np.isnan(alpha_power) | np.isnan(beta_power))
        valid_secs = int(np.sum(valid_mask))
        skipped = int(n_secs - valid_secs)

        total_secs += n_secs
        total_valid += valid_secs
        total_skipped += skipped

        print(
            f"{os.path.basename(path)}: seconds={n_secs}, "
            f"valid={valid_secs}, skipped={skipped}"
        )

        if args.save_csv:
            for sec_idx in range(n_secs):
                t_val = float(theta_power[sec_idx]) if not np.isnan(theta_power[sec_idx]) else float("nan")
                a_val = float(alpha_power[sec_idx]) if not np.isnan(alpha_power[sec_idx]) else float("nan")
                b_val = float(beta_power[sec_idx]) if not np.isnan(beta_power[sec_idx]) else float("nan")
                ab_val = float(alpha_beta[sec_idx]) if not np.isnan(alpha_beta[sec_idx]) else float("nan")
                at_val = float(alpha_theta[sec_idx]) if not np.isnan(alpha_theta[sec_idx]) else float("nan")
                atotal_val = float(alpha_total[sec_idx]) if not np.isnan(alpha_total[sec_idx]) else float("nan")
                csv_rows.append((os.path.basename(path), sec_idx + 1, t_val, a_val, b_val, ab_val, at_val, atotal_val))

    print("--- Summary ---")
    print(f"Files processed: {len(edf_files)}")
    print(f"Total seconds: {total_secs}")
    print(f"Valid seconds: {total_valid}")
    print(f"Skipped seconds (invalid): {total_skipped}")

    if args.save_csv:
        save_csv(args.save_csv, csv_rows)
        print(f"Saved per-second values to: {args.save_csv}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
