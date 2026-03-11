import argparse
import csv
import os
from typing import Dict, List, Tuple

import numpy as np
import pyedflib

def load_eyeblinkning(data_path: str) -> List[int]:
    if not os.path.exists(data_path):
        print(f"Warning: arousal data file not found: {data_path}")
        return []
    try:
        with open(data_path, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                return []
            parts = content.split(",")
            blinks_seconds = [int(float(p)) for p in parts[1:]]  # Skip the first part (count)
            return blinks_seconds
    except Exception as e:
        print(f"Error reading arousal data from {data_path}: {e}")
        return []

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
) -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
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


def extract_react_times(edf_path: str) -> Dict[int, float]:
    """
    使用 stage 1->2 的規則，react_time = sec_253 - sec_251。
    t = sec，sec 為 1-based 秒數。
    回傳 {second_key: react_time}，second_key = round(sec_251 * 10) 的整數 (1 位小數)。
    """
    f = pyedflib.EdfReader(edf_path)
    try:
        labels = f.getSignalLabels()
        status_idx = find_channel_index(labels, "status")
        if status_idx is None:
            raise ValueError("Status channel not found in EDF.")

        status_signal = f.readSignal(status_idx)
        fs = float(f.getSampleFrequency(status_idx))
        if fs <= 0:
            raise ValueError("Invalid sample rate.")

        total_samples = len(status_signal)
        total_seconds = int(total_samples // fs)

        stage = 1
        sec_251 = None
        events: Dict[int, float] = {}

        for sec in range(1, total_seconds + 1):
            start = int((sec - 1) * fs)
            end = int(sec * fs)
            segment = status_signal[start:end]

            for i in range(len(segment)):
                if segment[i] > 1:
                    t = float(sec)
                    if stage == 1:
                        sec_251 = t + 0.002 * i
                        stage = 2
                    elif stage == 2:
                        sec_253 = t + 0.002 * i
                        stage = 3
                    elif stage == 3:
                        if sec_251 is not None and sec_253 is not None:
                            sec_251_round = int(round(sec_251, 0))
                            react_time = round(sec_253 - sec_251, 1)
                            key = int(round(sec_251_round * 10))
                            print(sec_251, sec_251_round,key, react_time)
                            if key not in events:
                                events[key] = react_time
                        stage = 1
    finally:
        f.close()

    return events


def merge_react_time_into_csv(csv_path: str, events: Dict[int, float]) -> None:
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"CSV not found: {csv_path}")

    with open(csv_path, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames or []
        rows = list(reader)

    if "second" not in fieldnames:
        raise ValueError("CSV must contain 'second' column.")

    if "react_time" not in fieldnames:
        fieldnames.append("react_time")

    def second_to_key(val: str) -> int | None:
        try:
            return int(round(float(val) * 10))
        except Exception:
            return None

    existing_keys = set()
    for row in rows:
        key = second_to_key(row.get("second", ""))
        if key is None:
            continue
        existing_keys.add(key)
        if key in events:
            row["react_time"] = f"{events[key]:.1f}"

    for key, react_time in events.items():
        if key in existing_keys:
            continue
        sec_value = key / 10.0
        new_row = {name: "" for name in fieldnames}
        new_row["second"] = f"{sec_value}"
        new_row["react_time"] = f"{react_time:.1f}"
        rows.append(new_row)

    def sort_key(row: Dict[str, str]) -> float:
        try:
            return float(row.get("second", ""))
        except Exception:
            return float("inf")

    rows.sort(key=sort_key)

    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


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


def save_csv(output_path: str, rows: List[Tuple[str, int, float, float, float, float, float, int]]) -> None:
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
                "alpha_total",
                "eyeblinking_count"
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
    csv_rows: List[Tuple[str, int, float, float, float, float, float, float, int]] = []

    for path in edf_files:
        dat_path = path.replace(".edf", "_arousal info.dat")
        blink_seconds = load_eyeblinkning(dat_path)
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
                t = sec_idx + 1
                start_range = max(0, t - 30)
                count = sum(1 for blink in blink_seconds if start_range <= blink <= t)
                t_val = float(theta_power[sec_idx]) if not np.isnan(theta_power[sec_idx]) else float("nan")
                a_val = float(alpha_power[sec_idx]) if not np.isnan(alpha_power[sec_idx]) else float("nan")
                b_val = float(beta_power[sec_idx]) if not np.isnan(beta_power[sec_idx]) else float("nan")
                ab_val = float(alpha_beta[sec_idx]) if not np.isnan(alpha_beta[sec_idx]) else float("nan")
                at_val = float(alpha_theta[sec_idx]) if not np.isnan(alpha_theta[sec_idx]) else float("nan")
                atotal_val = float(alpha_total[sec_idx]) if not np.isnan(alpha_total[sec_idx]) else float("nan")

                csv_rows.append((
                    os.path.basename(path),
                    t,
                    t_val, a_val, b_val,
                    ab_val, at_val, atotal_val,
                    count
                ))

    print("--- Summary ---")
    print(f"Files processed: {len(edf_files)}")
    print(f"Total seconds: {total_secs}")
    print(f"Valid seconds: {total_valid}")
    print(f"Skipped seconds (invalid): {total_skipped}")

    if args.save_csv:
        save_csv(args.save_csv, csv_rows)

        if len(edf_files) == 1:
            events = extract_react_times(edf_files[0])
            merge_react_time_into_csv(args.save_csv, events)
        else:
            print("Multiple EDF files provided. Skipping react_time merge.")

        print(f"Saved per-second values to: {args.save_csv}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
