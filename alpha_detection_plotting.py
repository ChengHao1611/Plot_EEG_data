from __future__ import annotations

from pathlib import Path
from typing import Sequence

import matplotlib
import numpy as np

matplotlib.use("Agg")
import matplotlib.pyplot as plt

MIN_PLOT_FREQ_HZ = 2.0


def to_float_or_none(value: object) -> float | None:
    if not isinstance(value, (int, float, np.floating, np.integer)):
        return None
    number = float(value)
    if not np.isfinite(number):
        return None
    return number


def to_int_or_none(value: object) -> int | None:
    number = to_float_or_none(value)
    if number is None:
        return None
    return int(round(number))


def get_second_segment(signal: np.ndarray, fs: float, second: int) -> tuple[np.ndarray, np.ndarray]:
    window = int(round(fs))
    if window <= 0:
        raise ValueError(f"Invalid fs: {fs}")
    if second <= 0:
        raise ValueError(f"Second must be positive: {second}")

    start = (second - 1) * window
    end = start + window
    if end > signal.size:
        raise IndexError(f"Second {second} exceeds available signal length.")

    segment = signal[start:end].astype(float)
    times = np.arange(segment.size, dtype=float) / float(fs)
    return times, segment


def save_class_plot(
    output_path: Path,
    *,
    second: int,
    label: str,
    signal: np.ndarray,
    fs: float,
    freqs: np.ndarray,
    power_segment: np.ndarray,
    alpha_low: float,
    alpha_high: float,
    max_plot_freq: float,
    alpha_amplitude: float | None,
    alpha_peak: int | None,
) -> None:
    times, segment = get_second_segment(signal, fs, second)

    spectrum_mask = (
        np.isfinite(freqs)
        & (freqs >= float(MIN_PLOT_FREQ_HZ))
        & (freqs <= float(max_plot_freq))
    )
    power_masked = power_segment[spectrum_mask] if power_segment.size == freqs.size else np.array([], dtype=float)
    freqs_masked = freqs[spectrum_mask] if power_segment.size == freqs.size else np.array([], dtype=float)

    fig, axes = plt.subplots(2, 1, figsize=(12, 8))

    axes[0].plot(freqs_masked, power_masked, color="#1f77b4", linewidth=1.0)
    axes[0].axvspan(alpha_low, alpha_high, color="#f4d35e", alpha=0.25)
    axes[0].set_xlim(float(MIN_PLOT_FREQ_HZ), float(max_plot_freq))
    axes[0].set_xlabel("Frequency (Hz)")
    axes[0].set_ylabel("Amplitude")
    axes[0].set_title(f"Frequency Domain ({int(MIN_PLOT_FREQ_HZ)}-{int(round(float(max_plot_freq)))} Hz)")
    axes[0].grid(True, alpha=0.2)

    axes[1].plot(times, segment, color="#222222", linewidth=0.9)
    axes[1].set_xlim(times[0], times[-1] if times.size > 1 else 1.0)
    axes[1].set_xlabel("Time (s)")
    axes[1].set_ylabel("Signal")
    axes[1].set_title("Time Domain")
    axes[1].grid(True, alpha=0.2)

    amplitude_text = "NA" if alpha_amplitude is None else f"{alpha_amplitude:.6f}"
    peak_text = "NA" if alpha_peak is None else str(alpha_peak)
    fig.suptitle(
        f"{label} | second {second}\nalpha_amplitude={amplitude_text} alpha_peak={peak_text}",
        fontsize=13,
    )
    fig.tight_layout(rect=(0.0, 0.0, 1.0, 0.95))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def export_classified_plots(
    result_rows: Sequence[dict[str, object]],
    *,
    output_dir: Path,
    signal: np.ndarray,
    fs: float,
    nfft: int,
    alpha_low: float,
    alpha_high: float,
    max_plot_freq: float,
) -> tuple[int, int, int, list[str]]:
    freqs = np.fft.rfftfreq(int(nfft), d=1.0 / float(fs))
    plotted_true = 0
    plotted_false = 0
    plotted_miss = 0
    warnings: list[str] = []

    for label in ("true", "false", "miss"):
        (output_dir / label).mkdir(parents=True, exist_ok=True)

    for row in result_rows:
        second_raw = row.get("second")
        label_raw = row.get("label")
        power_raw = row.get("power")
        if not isinstance(second_raw, int) or not isinstance(label_raw, str):
            warnings.append(f"Skipped row without valid second/label: {row}")
            continue

        second = second_raw
        label = label_raw
        power_segment = np.asarray(power_raw, dtype=float) if power_raw is not None else np.array([], dtype=float)
        output_path = output_dir / label / f"s{second:05d}_{label}.png"

        try:
            save_class_plot(
                output_path,
                second=second,
                label=label,
                signal=signal,
                fs=fs,
                freqs=freqs,
                power_segment=power_segment,
                alpha_low=float(alpha_low),
                alpha_high=float(alpha_high),
                max_plot_freq=float(max_plot_freq),
                alpha_amplitude=to_float_or_none(row.get("alpha_amplitude")),
                alpha_peak=to_int_or_none(row.get("alpha_peak")),
            )
        except Exception as exc:
            warnings.append(f"Skipped second {second} ({label}): {exc}")
            continue

        if label == "true":
            plotted_true += 1
        elif label == "false":
            plotted_false += 1
        elif label == "miss":
            plotted_miss += 1

    return plotted_true, plotted_false, plotted_miss, warnings


__all__ = [
    "export_classified_plots",
]
