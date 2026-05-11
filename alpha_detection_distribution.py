from __future__ import annotations

from pathlib import Path
from typing import Sequence

import matplotlib
import numpy as np

matplotlib.use("Agg")
import matplotlib.pyplot as plt

ALPHA_ENERGY_MAX = 4000.0


def to_float_or_none(value: object) -> float | None:
    if not isinstance(value, (int, float, np.floating, np.integer)):
        return None
    number = float(value)
    if not np.isfinite(number):
        return None
    return number


def collect_alpha_energy_by_labels(
    result_rows: Sequence[dict[str, object]],
    labels: set[str],
) -> np.ndarray:
    values: list[float] = []
    for row in result_rows:
        label = row.get("label")
        if label not in labels:
            continue
        alpha_value = to_float_or_none(row.get("alpha_amplitude"))
        if alpha_value is None:
            continue
        values.append(alpha_value)
    return np.asarray(values, dtype=float)


def clip_alpha_energy(values: np.ndarray, *, max_value: float = ALPHA_ENERGY_MAX) -> np.ndarray:
    if values.size == 0:
        return values.astype(float, copy=True)
    clipped = np.asarray(values, dtype=float).copy()
    clipped[clipped > float(max_value)] = float(max_value)
    return clipped


def normal_pdf(x: np.ndarray, mean: float, std: float) -> np.ndarray:
    coefficient = 1.0 / (std * np.sqrt(2.0 * np.pi))
    exponent = -0.5 * ((x - mean) / std) ** 2
    return coefficient * np.exp(exponent)


def plot_hist_and_normal(
    ax: plt.Axes,
    values: np.ndarray,
    *,
    color: str,
    label: str,
    bin_edges: np.ndarray,
) -> None:
    if values.size == 0:
        return

    bin_width = float(bin_edges[1] - bin_edges[0]) if bin_edges.size > 1 else 1.0

    ax.hist(
        values,
        bins=bin_edges,
        density=False,
        alpha=0.28,
        color=color,
        edgecolor=color,
        linewidth=1.0,
        label=f"{label} histogram",
    )

    mean = float(np.mean(values))
    std = float(np.std(values))
    if std <= 1e-12:
        ax.axvline(mean, color=color, linewidth=2.0, label=f"{label} mean={mean:.2f}")
        return

    x_min = float(bin_edges[0])
    x_max = float(bin_edges[-1])
    if x_min == x_max:
        ax.axvline(mean, color=color, linewidth=2.0, label=f"{label} mean={mean:.2f}")
        return

    x_values = np.linspace(x_min, x_max, 400, dtype=float)
    pdf = normal_pdf(x_values, mean, std) * float(values.size) * bin_width
    ax.plot(
        x_values,
        pdf,
        color=color,
        linewidth=2.2,
        label=f"{label} normal fit (mu={mean:.2f}, sigma={std:.2f})",
    )


def save_alpha_energy_normal_distribution(
    output_path: Path,
    result_rows: Sequence[dict[str, object]],
) -> tuple[int, int]:
    true_miss_values = clip_alpha_energy(
        collect_alpha_energy_by_labels(result_rows, {"true", "miss"})
    )
    false_values = clip_alpha_energy(
        collect_alpha_energy_by_labels(result_rows, {"false"})
    )

    fig, ax = plt.subplots(figsize=(12, 7))

    max_count = max(true_miss_values.size, false_values.size, 1)
    bins = max(10, min(40, int(np.sqrt(max_count) * 2)))
    bin_edges = np.linspace(0.0, float(ALPHA_ENERGY_MAX), bins + 1, dtype=float)

    plot_hist_and_normal(
        ax,
        true_miss_values,
        color="#2a9d8f",
        label="true+miss alpha energy",
        bin_edges=bin_edges,
    )
    plot_hist_and_normal(
        ax,
        false_values,
        color="#e76f51",
        label="false alpha energy",
        bin_edges=bin_edges,
    )

    if true_miss_values.size == 0 and false_values.size == 0:
        ax.text(
            0.5,
            0.5,
            "No alpha energy data available",
            transform=ax.transAxes,
            ha="center",
            va="center",
            fontsize=13,
        )

    ax.set_title("Alpha Energy Normal Distribution")
    ax.set_xlabel("Alpha Energy (alpha_amplitude)")
    ax.set_ylabel("Count")
    ax.set_xlim(0.0, float(ALPHA_ENERGY_MAX))
    ax.grid(True, alpha=0.2)
    ax.legend()

    fig.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=180, bbox_inches="tight")
    plt.close(fig)

    return int(true_miss_values.size), int(false_values.size)


__all__ = [
    "save_alpha_energy_normal_distribution",
]
