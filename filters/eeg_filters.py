from __future__ import annotations

import numpy as np
from scipy.signal import butter, filtfilt, iirnotch, sosfiltfilt


def _validate_sample_rate(fs: float) -> float:
    fs_value = float(fs)
    if fs_value <= 0:
        raise ValueError(f"Invalid sample rate for filtering: {fs}")
    return fs_value


def _validate_cutoff(cutoff_hz: float, *, fs: float, filter_name: str) -> float:
    cutoff_value = float(cutoff_hz)
    if cutoff_value <= 0:
        raise ValueError(f"{filter_name} cutoff must be greater than 0 Hz: {cutoff_hz}")

    nyquist = fs / 2.0
    if cutoff_value >= nyquist:
        raise ValueError(
            f"Cannot apply {filter_name} cutoff at {cutoff_value:.6f} Hz when Nyquist is "
            f"{nyquist:.6f} Hz."
        )

    return cutoff_value


def apply_highpass_filter(
    signal: np.ndarray,
    fs: float,
    *,
    cutoff_hz: float,
    order: int = 4,
) -> np.ndarray:
    fs_value = _validate_sample_rate(fs)
    cutoff_value = _validate_cutoff(cutoff_hz, fs=fs_value, filter_name="high-pass filter")
    sos = butter(
        int(order),
        cutoff_value,
        btype="highpass",
        fs=fs_value,
        output="sos",
    )
    return sosfiltfilt(sos, np.asarray(signal, dtype=float))


def apply_lowpass_filter(
    signal: np.ndarray,
    fs: float,
    *,
    cutoff_hz: float,
    order: int = 4,
) -> np.ndarray:
    fs_value = _validate_sample_rate(fs)
    cutoff_value = _validate_cutoff(cutoff_hz, fs=fs_value, filter_name="low-pass filter")
    sos = butter(
        int(order),
        cutoff_value,
        btype="lowpass",
        fs=fs_value,
        output="sos",
    )
    return sosfiltfilt(sos, np.asarray(signal, dtype=float))


def apply_notch_filter(
    signal: np.ndarray,
    fs: float,
    *,
    notch_freq: float,
    quality_factor: float = 30.0,
) -> np.ndarray:
    fs_value = _validate_sample_rate(fs)
    notch_value = _validate_cutoff(notch_freq, fs=fs_value, filter_name="notch filter")
    q_value = float(quality_factor)
    if q_value <= 0:
        raise ValueError(f"Notch filter quality factor must be greater than 0: {quality_factor}")

    b, a = iirnotch(notch_value, q_value, fs=fs_value)
    return filtfilt(b, a, np.asarray(signal, dtype=float))
