"""Alpha-dominant band-power detection for the fatigue-driving system."""

from __future__ import annotations

import argparse
import statistics
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

import numpy as np
import pyedflib

from eeg_analysis.common.filters import apply_highpass_filter, apply_lowpass_filter


THETA_LOW_HZ = 4.0
THETA_HIGH_HZ = 7.0
ALPHA_LOW_HZ = 8.0
ALPHA_HIGH_HZ = 12.0
BETA_LOW_HZ = 13.0
BETA_HIGH_HZ = 20.0
HIGHPASS_HZ = 1.0
LOWPASS_HZ = 30.0
USE_NFFT_POW2 = True


@dataclass(frozen=True)
class BandPowerRecord:
    """Band powers and Alpha qualification for one one-based second."""

    second: int
    excluded_by_eye: bool
    theta_power: float | None
    alpha_power: float | None
    beta_power: float | None
    alpha_qualified: bool


@dataclass(frozen=True)
class AlphaDetectionResult:
    """Complete per-second result returned to Function One."""

    channel_name: str
    sample_rate: float
    records: tuple[BandPowerRecord, ...]

    @property
    def alpha_seconds(self) -> list[int]:
        return [record.second for record in self.records if record.alpha_qualified]

    @property
    def qualified_alpha_powers(self) -> list[float]:
        return [
            float(record.alpha_power)
            for record in self.records
            if record.alpha_qualified and record.alpha_power is not None
        ]

    @property
    def alpha_mean(self) -> float | None:
        values = self.qualified_alpha_powers
        return statistics.mean(values) if values else None

    @property
    def alpha_median(self) -> float | None:
        values = self.qualified_alpha_powers
        return statistics.median(values) if values else None

    @property
    def excluded_eye_seconds(self) -> int:
        return sum(record.excluded_by_eye for record in self.records)

    @property
    def valid_fft_seconds(self) -> int:
        return sum(
            not record.excluded_by_eye and record.alpha_power is not None
            for record in self.records
        )


def save_dat_seconds(output_path: str | Path, seconds: Sequence[int]) -> None:
    """Save unique one-based Alpha seconds as ``count,second,...``."""
    output = Path(output_path)
    ordered_seconds = sorted({int(second) for second in seconds})
    values = [str(len(ordered_seconds)), *(str(second) for second in ordered_seconds)]
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(",".join(values) + "\n", encoding="utf-8")


def load_dat_seconds(input_path: str | Path) -> list[int]:
    """Read and validate the project's common DAT second-list format."""
    path = Path(input_path)
    if not path.is_file():
        raise FileNotFoundError(f"DAT not found: {path}")
    content = path.read_text(encoding="utf-8").strip()
    if not content:
        raise ValueError(f"DAT is empty: {path}")
    try:
        values = [int(part.strip()) for part in content.split(",") if part.strip()]
    except ValueError as exc:
        raise ValueError(f"DAT contains a non-integer value: {path}") from exc
    if not values:
        raise ValueError(f"DAT does not contain an event count: {path}")
    declared_count, seconds = values[0], values[1:]
    if declared_count != len(seconds):
        raise ValueError(
            f"DAT count mismatch in {path}: declared {declared_count}, found {len(seconds)}"
        )
    return seconds


def next_power_of_2(value: int) -> int:
    if value <= 1:
        return 1
    return 1 << (value - 1).bit_length()


def find_channel_index(labels: Sequence[str], target: str) -> int | None:
    normalized_target = str(target).strip().casefold()
    for index, label in enumerate(labels):
        if str(label).strip().casefold() == normalized_target:
            return index
    for index, label in enumerate(labels):
        if normalized_target in str(label).strip().casefold():
            return index
    return None


def resolve_alpha_channel_name(target_channels: Sequence[str] | None = None) -> str:
    if target_channels:
        fp2_channel = next(
            (str(channel) for channel in target_channels if "fp2" in str(channel).casefold()),
            None,
        )
        if fp2_channel is not None:
            return fp2_channel
        return str(list(target_channels)[-1])
    return "FP2"


def load_channel_signal(
    edf_path: str | Path, channel_name: str
) -> tuple[np.ndarray, float, str]:
    path = Path(edf_path)
    if not path.is_file():
        raise FileNotFoundError(f"EDF not found: {path}")
    reader = pyedflib.EdfReader(str(path))
    try:
        labels = reader.getSignalLabels()
        channel_index = find_channel_index(labels, channel_name)
        if channel_index is None:
            raise ValueError(f"Channel '{channel_name}' not found in {path}")
        signal = reader.readSignal(channel_index).astype(float)
        sample_rate = float(reader.getSampleFrequency(channel_index))
        resolved_name = str(labels[channel_index])
    finally:
        reader.close()
    return signal, sample_rate, resolved_name


def _band_power(
    amplitude_squared: np.ndarray,
    frequencies: np.ndarray,
    low: float,
    high: float,
) -> float:
    low_index, high_index = resolve_band_bin_indices(frequencies, low, high)
    return float(np.sum(amplitude_squared[low_index : high_index + 1]))


def resolve_band_bin_indices(
    frequencies: np.ndarray,
    low_hz: float,
    high_hz: float,
) -> tuple[int, int]:
    """Return FFT bins nearest to the requested inclusive band boundaries.

    For example, fs=500 and nfft=512 gives a 0.9765625 Hz bin spacing.  The
    nominal 8--12 Hz Alpha band therefore maps to bins 8--12, whose actual
    center frequencies are 7.8125--11.71875 Hz.
    """
    frequency_values = np.asarray(frequencies, dtype=float)
    if frequency_values.ndim != 1 or frequency_values.size == 0:
        raise ValueError("frequencies must be a non-empty one-dimensional array")
    if not np.all(np.isfinite(frequency_values)):
        raise ValueError("frequencies must contain only finite values")
    if float(low_hz) > float(high_hz):
        raise ValueError("low_hz must be less than or equal to high_hz")

    low_index = int(np.argmin(np.abs(frequency_values - float(low_hz))))
    high_index = int(np.argmin(np.abs(frequency_values - float(high_hz))))
    if low_index > high_index:
        low_index, high_index = high_index, low_index
    return low_index, high_index


def compute_band_power_records(
    signal: np.ndarray,
    sample_rate: float,
    *,
    eye_seconds: Iterable[int] = (),
    start_second: int = 1,
    end_second: int | None = None,
    use_nfft_pow2: bool = USE_NFFT_POW2,
) -> list[BandPowerRecord]:
    """Compute Theta/Alpha/Beta power for non-eye one-second windows."""
    if sample_rate <= 0:
        raise ValueError(f"Invalid sample rate: {sample_rate}")
    if start_second < 1:
        raise ValueError("start_second must be at least 1")
    if end_second is not None and end_second < start_second:
        raise ValueError("end_second must be greater than or equal to start_second")

    window_size = int(round(sample_rate))
    available_seconds = int(len(signal) // window_size)
    final_second = available_seconds if end_second is None else min(end_second, available_seconds)
    if final_second < start_second:
        return []

    nfft = next_power_of_2(window_size) if use_nfft_pow2 else window_size
    frequencies = np.fft.rfftfreq(nfft, d=1.0 / float(sample_rate))
    excluded_seconds = {int(second) for second in eye_seconds}
    records: list[BandPowerRecord] = []

    for second in range(start_second, final_second + 1):
        if second in excluded_seconds:
            records.append(BandPowerRecord(second, True, None, None, None, False))
            continue

        start_index = (second - 1) * window_size
        segment = np.asarray(signal[start_index : start_index + window_size], dtype=float)
        if segment.size != window_size or not np.all(np.isfinite(segment)):
            records.append(BandPowerRecord(second, False, None, None, None, False))
            continue

        spectrum = np.fft.rfft(segment, n=nfft)
        amplitude_squared = np.abs(spectrum).astype(float, copy=False) ** 2
        theta_power = _band_power(
            amplitude_squared, frequencies, THETA_LOW_HZ, THETA_HIGH_HZ
        )
        alpha_power = _band_power(
            amplitude_squared, frequencies, ALPHA_LOW_HZ, ALPHA_HIGH_HZ
        )
        beta_power = _band_power(
            amplitude_squared, frequencies, BETA_LOW_HZ, BETA_HIGH_HZ
        )
        records.append(
            BandPowerRecord(
                second=second,
                excluded_by_eye=False,
                theta_power=theta_power,
                alpha_power=alpha_power,
                beta_power=beta_power,
                alpha_qualified=alpha_power > theta_power and alpha_power > beta_power,
            )
        )
    return records


def detect_alpha(
    edf_path: str | Path,
    output_path: str | Path,
    target_channels: Sequence[str] | None = None,
    *,
    eye_seconds: Iterable[int] = (),
    start_second: int = 1,
    end_second: int | None = None,
    alpha_power_threshold: float | None = None,
) -> AlphaDetectionResult:
    """Detect Alpha-dominant non-eye seconds and save them to DAT.

    When ``alpha_power_threshold`` is supplied, a second qualifies only when
    Alpha is dominant over Theta/Beta and its power is strictly above that
    threshold. Function Two uses the personal Alpha median for this value.
    """
    channel_name = resolve_alpha_channel_name(target_channels)
    signal, sample_rate, resolved_name = load_channel_signal(edf_path, channel_name)
    window_size = int(round(sample_rate))
    available_seconds = int(len(signal) // window_size)
    final_second = available_seconds if end_second is None else min(end_second, available_seconds)
    if final_second <= 0:
        raise ValueError("EDF does not contain a complete one-second FP2 window")

    # Limit filtering to the requested prefix so Function One never uses data
    # after second 300 when deriving its baseline.
    signal = signal[: final_second * window_size]
    signal = apply_highpass_filter(signal, sample_rate, cutoff_hz=HIGHPASS_HZ)
    signal = apply_lowpass_filter(signal, sample_rate, cutoff_hz=LOWPASS_HZ)
    records = compute_band_power_records(
        signal,
        sample_rate,
        eye_seconds=eye_seconds,
        start_second=start_second,
        end_second=final_second,
    )
    if alpha_power_threshold is not None:
        threshold = float(alpha_power_threshold)
        records = [
            BandPowerRecord(
                second=record.second,
                excluded_by_eye=record.excluded_by_eye,
                theta_power=record.theta_power,
                alpha_power=record.alpha_power,
                beta_power=record.beta_power,
                alpha_qualified=(
                    record.alpha_qualified
                    and record.alpha_power is not None
                    and record.alpha_power > threshold
                ),
            )
            for record in records
        ]
    result = AlphaDetectionResult(
        channel_name=resolved_name,
        sample_rate=sample_rate,
        records=tuple(records),
    )
    save_dat_seconds(output_path, result.alpha_seconds)

    print("處理完成！")
    if alpha_power_threshold is None:
        condition_text = "Alpha > Theta 且 Alpha > Beta"
    else:
        condition_text = (
            "Alpha > Theta、Alpha > Beta 且 "
            f"Alpha Power > {float(alpha_power_threshold):.6g}"
        )
    print(f"符合 {condition_text} 的秒數：{len(result.alpha_seconds)}")
    print(f"眼動排除秒數：{result.excluded_eye_seconds}")
    print(f"結果已存入：{output_path}")
    return result


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Detect Alpha-dominant FP2 seconds using Theta/Alpha/Beta band power."
    )
    parser.add_argument("--file", required=True, help="Input EDF path.")
    parser.add_argument("--output", help="Output DAT path (default: Alpha.dat beside EDF).")
    parser.add_argument("--eye-dat", help="Optional eye DAT whose seconds will be excluded.")
    parser.add_argument("--channel", default="FP2")
    parser.add_argument("--start-second", type=int, default=1)
    parser.add_argument("--end-second", type=int)
    parser.add_argument(
        "--alpha-power-threshold",
        type=float,
        help="Optional strict lower bound for qualified Alpha Power.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output_path = Path(args.output) if args.output else Path(args.file).parent / "Alpha.dat"
    eye_seconds = load_dat_seconds(args.eye_dat) if args.eye_dat else []
    result = detect_alpha(
        args.file,
        output_path,
        [args.channel],
        eye_seconds=eye_seconds,
        start_second=args.start_second,
        end_second=args.end_second,
        alpha_power_threshold=args.alpha_power_threshold,
    )
    print(f"Alpha Power平均：{result.alpha_mean}")
    print(f"Alpha Power中位數：{result.alpha_median}")


if __name__ == "__main__":
    main()


__all__ = [
    "ALPHA_HIGH_HZ",
    "ALPHA_LOW_HZ",
    "AlphaDetectionResult",
    "BETA_HIGH_HZ",
    "BETA_LOW_HZ",
    "BandPowerRecord",
    "THETA_HIGH_HZ",
    "THETA_LOW_HZ",
    "compute_band_power_records",
    "detect_alpha",
    "find_channel_index",
    "load_channel_signal",
    "load_dat_seconds",
    "next_power_of_2",
    "resolve_band_bin_indices",
    "resolve_alpha_channel_name",
    "save_dat_seconds",
]
