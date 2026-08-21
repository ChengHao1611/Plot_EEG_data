"""Calculate training-derived fallback scales for physiological baselines."""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .physiological_fatigue import (
    RobustBaseline,
    compute_robust_baseline,
    estimate_pooled_scale,
    rolling_event_counts,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_MANIFEST_PATH = PROJECT_ROOT / "train_data"
DEFAULT_RESULTS_ROOT = (
    PROJECT_ROOT / "data" / "derived" / "fatigue_driving_prediction_system"
)
DEFAULT_REPORT_PATH = DEFAULT_RESULTS_ROOT / "training_pooled_baseline.xlsx"


@dataclass(frozen=True)
class TrainingRecordBaseline:
    """Alpha and eye baseline statistics for one training recording."""

    record_id: str
    alpha: RobustBaseline
    eye: RobustBaseline


@dataclass(frozen=True)
class TrainingBaselineSummary:
    """Per-record statistics and the two supported pooled summaries."""

    records: tuple[TrainingRecordBaseline, ...]
    combined_alpha: RobustBaseline
    combined_eye: RobustBaseline
    pooled_alpha_scale: float
    pooled_eye_scale: float


def read_training_record_ids(manifest_path: str | Path) -> list[str]:
    """Read record IDs separated by whitespace or Chinese/ASCII commas."""
    path = Path(manifest_path)
    if not path.is_file():
        raise FileNotFoundError(f"找不到訓練資料清單：{path}")

    record_ids = [
        token.strip()
        for token in re.split(
            r"[\s,，、]+", path.read_text(encoding="utf-8-sig")
        )
        if token.strip()
    ]
    if not record_ids:
        raise ValueError(f"訓練資料清單是空的：{path}")

    duplicates = sorted(
        record_id
        for record_id in set(record_ids)
        if record_ids.count(record_id) > 1
    )
    if duplicates:
        raise ValueError("train_data 含重複資料：" + "、".join(duplicates))
    return record_ids


def load_event_seconds(path: str | Path) -> list[int]:
    """Load detector event seconds from a header-prefixed DAT file."""
    dat_path = Path(path)
    if not dat_path.is_file():
        raise FileNotFoundError(f"找不到訓練特徵檔：{dat_path}")

    frame = pd.read_csv(dat_path, header=None, sep=None, engine="python")
    flattened = frame.to_numpy().ravel()
    if flattened.size <= 1:
        return []
    seconds = pd.to_numeric(
        pd.Series(flattened[1:]), errors="coerce"
    ).dropna()
    return seconds.astype(int).tolist()


def calculate_training_baselines(
    manifest_path: str | Path = DEFAULT_MANIFEST_PATH,
    results_root: str | Path = DEFAULT_RESULTS_ROOT,
    *,
    complete_window_start_second: int = 30,
    baseline_end_second: int = 300,
    window_seconds: int = 30,
) -> TrainingBaselineSummary:
    """Recalculate training MAD/IQR values and fallback pooled scales.

    The combined baselines describe all complete training windows concatenated
    together.  The fallback scales follow the existing production definition:
    the median of valid per-record MAD/IQR scales.
    """
    record_ids = read_training_record_ids(manifest_path)
    root = Path(results_root)
    records: list[TrainingRecordBaseline] = []
    all_alpha_windows: list[int] = []
    all_eye_windows: list[int] = []

    for record_id in record_ids:
        record_root = root / record_id
        alpha_seconds = load_event_seconds(
            record_root / "Alpha_function_one.dat"
        )
        eye_seconds = load_event_seconds(record_root / "eyeblink.dat")

        alpha_windows = list(
            rolling_event_counts(
                alpha_seconds,
                start_second=complete_window_start_second,
                end_second=baseline_end_second,
                window_seconds=window_seconds,
            ).values()
        )
        eye_windows = list(
            rolling_event_counts(
                eye_seconds,
                start_second=complete_window_start_second,
                end_second=baseline_end_second,
                window_seconds=window_seconds,
            ).values()
        )
        alpha_baseline = compute_robust_baseline(alpha_windows)
        eye_baseline = compute_robust_baseline(eye_windows)
        records.append(
            TrainingRecordBaseline(
                record_id=record_id,
                alpha=alpha_baseline,
                eye=eye_baseline,
            )
        )
        all_alpha_windows.extend(alpha_windows)
        all_eye_windows.extend(eye_windows)

    pooled_alpha_scale = estimate_pooled_scale(
        record.alpha for record in records
    )
    pooled_eye_scale = estimate_pooled_scale(record.eye for record in records)
    if pooled_alpha_scale is None or pooled_eye_scale is None:
        invalid_features = []
        if pooled_alpha_scale is None:
            invalid_features.append("Alpha")
        if pooled_eye_scale is None:
            invalid_features.append("Eye blink")
        raise ValueError(
            "訓練資料無法建立有效的 Pooled Scale："
            + "、".join(invalid_features)
        )

    return TrainingBaselineSummary(
        records=tuple(records),
        combined_alpha=compute_robust_baseline(all_alpha_windows),
        combined_eye=compute_robust_baseline(all_eye_windows),
        pooled_alpha_scale=float(pooled_alpha_scale),
        pooled_eye_scale=float(pooled_eye_scale),
    )


def write_training_baseline_report(
    summary: TrainingBaselineSummary,
    output_path: str | Path = DEFAULT_REPORT_PATH,
) -> Path:
    """Write combined and per-record MAD/IQR statistics to an Excel report."""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    combined_rows = []
    for feature, baseline, fallback_scale in (
        ("Alpha", summary.combined_alpha, summary.pooled_alpha_scale),
        ("Eye", summary.combined_eye, summary.pooled_eye_scale),
    ):
        combined_rows.append(
            {
                "Feature": feature,
                "Training Records": len(summary.records),
                "Combined Windows": baseline.sample_count,
                "Combined Median": baseline.median,
                "Combined MAD": baseline.mad,
                "Combined IQR": baseline.iqr,
                "Combined Scale": baseline.scale,
                "Combined Scale Method": baseline.scale_method,
                "Fallback Pooled Scale": fallback_scale,
                "Fallback Definition": "Median of valid per-record scales",
            }
        )

    record_rows = []
    for record in summary.records:
        for feature, baseline in (
            ("Alpha", record.alpha),
            ("Eye", record.eye),
        ):
            record_rows.append(
                {
                    "Record ID": record.record_id,
                    "Feature": feature,
                    "Complete Windows": baseline.sample_count,
                    "Median": baseline.median,
                    "MAD": baseline.mad,
                    "IQR": baseline.iqr,
                    "Scale": baseline.scale,
                    "Scale Method": baseline.scale_method,
                }
            )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(combined_rows).to_excel(
            writer, sheet_name="Pooled Summary", index=False
        )
        pd.DataFrame(record_rows).to_excel(
            writer, sheet_name="Per Record", index=False
        )
    return output


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Calculate training MAD/IQR and pooled fallback scales."
    )
    parser.add_argument("--manifest", default=DEFAULT_MANIFEST_PATH)
    parser.add_argument("--results-root", default=DEFAULT_RESULTS_ROOT)
    parser.add_argument("--output", default=DEFAULT_REPORT_PATH)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    summary = calculate_training_baselines(args.manifest, args.results_root)
    output = write_training_baseline_report(summary, args.output)
    print(f"Alpha fallback pooled scale: {summary.pooled_alpha_scale:g}")
    print(f"Eye fallback pooled scale: {summary.pooled_eye_scale:g}")
    print(f"Report: {output}")


if __name__ == "__main__":
    main()


__all__ = [
    "DEFAULT_MANIFEST_PATH",
    "DEFAULT_REPORT_PATH",
    "DEFAULT_RESULTS_ROOT",
    "TrainingBaselineSummary",
    "TrainingRecordBaseline",
    "calculate_training_baselines",
    "load_event_seconds",
    "read_training_record_ids",
    "write_training_baseline_report",
]
