from __future__ import annotations

from pathlib import Path
from typing import Sequence

import numpy as np
import pyedflib

from filters import apply_highpass_filter, apply_lowpass_filter, apply_notch_filter

# Detection settings mirrored from the standalone alpha workflow.
ALPHA_THRESHOLD = 1500.0
ALPHA_LOW_HZ = 7.8
ALPHA_HIGH_HZ = 12.5
PEAK_LOW_HZ = 1.9
PEAK_HIGH_HZ = 30.0
USE_NFFT_POW2 = True
APPLY_HIGHPASS = True
APPLY_LOWPASS = True
NOTCH_FREQ: float | None = None


def save_dat_seconds(output_path: str | Path, seconds: Sequence[int]) -> None:
    output_path_obj = Path(output_path)
    ordered_seconds = sorted({int(second) for second in seconds})
    values = [str(len(ordered_seconds)), *(str(second) for second in ordered_seconds)]
    output_path_obj.parent.mkdir(parents=True, exist_ok=True)
    output_path_obj.write_text(",".join(values) + "\n", encoding="utf-8")


def next_power_of_2(n: int) -> int:
    if n <= 1:
        return 1
    return 1 << (n - 1).bit_length()


def find_channel_index(labels: Sequence[str], target: str) -> int | None:
    target_lower = target.lower()
    for index, label in enumerate(labels):
        if target_lower in label.lower():
            return index
    return None


def resolve_alpha_channel_name(target_channels: Sequence[str] | None = None) -> str:
    if target_channels:
        channels = list(target_channels)
        if len(channels) >= 2:
            return str(channels[1])
        return str(channels[0])
    return "FP2"


def load_channel_signal(edf_path: str | Path, channel_name: str) -> tuple[np.ndarray, float]:
    edf_path_obj = Path(edf_path)
    if not edf_path_obj.exists():
        raise FileNotFoundError(f"EDF not found: {edf_path_obj}")

    reader = pyedflib.EdfReader(str(edf_path_obj))
    try:
        labels = reader.getSignalLabels()
        channel_index = find_channel_index(labels, channel_name)
        if channel_index is None:
            raise ValueError(f"Channel '{channel_name}' not found in {edf_path_obj}")

        signal = reader.readSignal(channel_index).astype(float)
        sfreq = float(reader.getSampleFrequency(channel_index))
    finally:
        reader.close()

    return signal, sfreq


def compute_alpha_fft_features(
    signal: np.ndarray,
    fs: float,
    *,
    alpha_low: float = ALPHA_LOW_HZ,
    alpha_high: float = ALPHA_HIGH_HZ,
    peak_low: float = PEAK_LOW_HZ,
    peak_high: float = PEAK_HIGH_HZ,
    nfft: int | None = None,
) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    if signal.size == 0 or fs <= 0:
        return np.array([], dtype=float), np.array([], dtype=float), np.array([], dtype=float)

    window_size = int(round(fs))
    if window_size <= 0:
        return np.array([], dtype=float), np.array([], dtype=float), np.array([], dtype=float)

    nfft_final = window_size if nfft is None else int(nfft)
    if nfft_final < window_size:
        nfft_final = window_size

    total_seconds = int(signal.size // window_size)
    if total_seconds <= 0:
        return np.array([], dtype=float), np.array([], dtype=float), np.array([], dtype=float)

    freqs = np.fft.rfftfreq(nfft_final, d=1.0 / fs)
    peak_mask = (freqs >= float(peak_low)) & (freqs <= float(peak_high))
    if not np.any(peak_mask):
        return np.array([], dtype=float), np.array([], dtype=float), np.array([], dtype=float)

    peak_indices = np.flatnonzero(peak_mask)
    alpha_amplitude = np.full(total_seconds, np.nan, dtype=float)
    alpha_peak = np.full(total_seconds, np.nan, dtype=float)
    peak_freqs = np.full(total_seconds, np.nan, dtype=float)

    for second_index in range(total_seconds):
        start_idx = second_index * window_size
        end_idx = start_idx + window_size
        segment = signal[start_idx:end_idx]
        if segment.size < window_size or np.isnan(segment).any():
            continue

        spectrum = np.fft.rfft(segment, n=nfft_final)
        amplitude = np.abs(spectrum).astype(float, copy=False)
        peak_rel_index = int(np.argmax(amplitude[peak_mask]))
        peak_abs_index = int(peak_indices[peak_rel_index])
        peak_freq = float(freqs[peak_abs_index])

        alpha_amplitude[second_index] = float(amplitude[peak_abs_index])
        alpha_peak[second_index] = 1.0 if float(alpha_low) <= peak_freq < float(alpha_high) else 0.0
        peak_freqs[second_index] = peak_freq

    return alpha_amplitude, alpha_peak, peak_freqs


def classify_alpha_seconds(
    alpha_amplitude: np.ndarray,
    alpha_peak: np.ndarray,
    *,
    alpha_threshold: float = ALPHA_THRESHOLD,
) -> list[int]:
    predicted_seconds: list[int] = []
    total_seconds = min(alpha_amplitude.size, alpha_peak.size)

    for second_index in range(total_seconds):
        peak_flag = float(alpha_peak[second_index])
        amplitude_value = float(alpha_amplitude[second_index])
        if not np.isfinite(peak_flag) or not np.isfinite(amplitude_value):
            continue
        if int(round(peak_flag)) == 1 and amplitude_value > float(alpha_threshold):
            predicted_seconds.append(second_index + 1)

    return predicted_seconds


def detect_alpha(
    edf_path: str | Path,
    output_path: str | Path,
    target_channels: Sequence[str] | None = None,
) -> list[int]:
    channel_name = resolve_alpha_channel_name(target_channels)
    signal, sfreq = load_channel_signal(edf_path, channel_name)

    if APPLY_HIGHPASS:
        signal = apply_highpass_filter(signal, sfreq, cutoff_hz=1.0)
    if APPLY_LOWPASS:
        signal = apply_lowpass_filter(signal, sfreq, cutoff_hz=30.0)
    if NOTCH_FREQ is not None:
        signal = apply_notch_filter(signal, sfreq, notch_freq=float(NOTCH_FREQ))

    window_size = int(round(sfreq))
    if window_size <= 0:
        raise ValueError(f"Invalid sample rate: {sfreq}")

    nfft_value = next_power_of_2(window_size) if USE_NFFT_POW2 else window_size
    alpha_amplitude, alpha_peak, _ = compute_alpha_fft_features(
        signal,
        sfreq,
        alpha_low=ALPHA_LOW_HZ,
        alpha_high=ALPHA_HIGH_HZ,
        peak_low=PEAK_LOW_HZ,
        peak_high=PEAK_HIGH_HZ,
        nfft=nfft_value,
    )
    predicted_seconds = classify_alpha_seconds(
        alpha_amplitude,
        alpha_peak,
        alpha_threshold=ALPHA_THRESHOLD,
    )

    save_dat_seconds(output_path, predicted_seconds)

    print(f"處理完成！")
    print(f"總α秒數：{len(predicted_seconds)}")
    print(f"結果已存入：{output_path}")
    return predicted_seconds


__all__ = [
    "ALPHA_HIGH_HZ",
    "ALPHA_LOW_HZ",
    "ALPHA_THRESHOLD",
    "APPLY_HIGHPASS",
    "APPLY_LOWPASS",
    "NOTCH_FREQ",
    "PEAK_HIGH_HZ",
    "PEAK_LOW_HZ",
    "USE_NFFT_POW2",
    "classify_alpha_seconds",
    "compute_alpha_fft_features",
    "detect_alpha",
    "save_dat_seconds",
]
