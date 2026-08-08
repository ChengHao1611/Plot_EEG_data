"""Function Two: predict the first fatigue event after the 300-second baseline."""

from __future__ import annotations

import argparse
import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

if __package__ in {None, ""}:
    project_root = Path(__file__).resolve().parents[2]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import mne
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from eeg_analysis.detection.record_alpha import (
    AlphaDetectionResult,
    BandPowerRecord,
    detect_alpha,
)
from eeg_analysis.detection.record_arousal import detect_eye_movements
from eeg_analysis.fatigue_driving_prediction_system.function_one import (
    FunctionOneResult,
    analyze_function_one,
)
from eeg_analysis.statistics_30s_alpha_eyeblink_of_fatigue.record_status_and_eyeblink_to_xlsx import (
    ReactionTimeEvent,
    extract_reaction_time_events,
)


@dataclass(frozen=True)
class FunctionTwoConfig:
    baseline_end_second: int = 300
    eye_window_seconds: int = 30
    eye_alert_threshold: int = 10
    alpha_window_seconds: int = 10
    alpha_alert_threshold: int = 3
    plot_seconds_before_fatigue: int = 30


DEFAULT_CONFIG = FunctionTwoConfig()


@dataclass(frozen=True)
class FunctionTwoFeatureRecord:
    second: int
    seconds_before_fatigue: int | None
    eye_detected: bool
    eye_window_count: int
    eye_alert: bool
    alpha_valid: bool
    theta_power: float | None
    alpha_power: float | None
    beta_power: float | None
    alpha_qualified: bool
    alpha_window_count: int
    alpha_alert: bool
    warning: bool
    warning_reason: str
    target_fatigue: bool


@dataclass(frozen=True)
class FunctionTwoResult:
    record_id: str
    status: str
    analysis_start_second: int
    analysis_end_second: int
    personalized_rt_threshold: float | None
    alpha_median_baseline: float | None
    target_event: ReactionTimeEvent | None
    first_warning_second: int | None
    first_warning_reason: str | None
    prediction_success: bool | None
    lead_seconds: int | None
    features: tuple[FunctionTwoFeatureRecord, ...]
    post_baseline_events: tuple[ReactionTimeEvent, ...]
    eye_seconds: tuple[int, ...]
    alpha_result: AlphaDetectionResult | None
    output_dir: Path


def find_first_post_baseline_fatigue_event(
    events: Sequence[ReactionTimeEvent],
    personalized_rt_threshold: float,
    baseline_end_second: int = 300,
) -> ReactionTimeEvent | None:
    """Return the first post-baseline RT event meeting the personal threshold."""
    eligible_events = sorted(
        (
            event
            for event in events
            if event.event_second > baseline_end_second
            and event.reaction_time >= personalized_rt_threshold
        ),
        key=lambda event: (event.event_second, event.deviation_time),
    )
    return eligible_events[0] if eligible_events else None


def classify_warning(
    eye_window_count: int,
    alpha_window_count: int,
    config: FunctionTwoConfig = DEFAULT_CONFIG,
) -> tuple[bool, str]:
    """Apply the requested OR rule and return warning state plus its reason."""
    eye_alert = eye_window_count < config.eye_alert_threshold
    alpha_alert = alpha_window_count >= config.alpha_alert_threshold
    if eye_alert and alpha_alert:
        return True, "EYE_AND_ALPHA"
    if eye_alert:
        return True, "EYE"
    if alpha_alert:
        return True, "ALPHA"
    return False, "NONE"


def build_function_two_features(
    *,
    start_second: int,
    end_second: int,
    eye_seconds: Iterable[int],
    alpha_records: Iterable[BandPowerRecord],
    alpha_median_baseline: float,
    target_second: int | None,
    config: FunctionTwoConfig = DEFAULT_CONFIG,
) -> list[FunctionTwoFeatureRecord]:
    """Build per-second windows, carrying pre-300 records into Function Two."""
    if start_second < 1:
        raise ValueError("start_second must be at least 1")
    if end_second < start_second:
        return []
    if not math.isfinite(alpha_median_baseline):
        raise ValueError("alpha_median_baseline must be finite")

    eye_second_set = {int(second) for second in eye_seconds}
    alpha_by_second = {record.second: record for record in alpha_records}
    qualified_alpha_seconds = {
        record.second
        for record in alpha_by_second.values()
        if record.alpha_qualified
        and record.alpha_power is not None
        and record.alpha_power > alpha_median_baseline
    }

    features: list[FunctionTwoFeatureRecord] = []
    for second in range(start_second, end_second + 1):
        eye_window_start = second - config.eye_window_seconds + 1
        alpha_window_start = second - config.alpha_window_seconds + 1
        eye_window_count = sum(
            window_second in eye_second_set
            for window_second in range(eye_window_start, second + 1)
        )
        alpha_window_count = sum(
            window_second in qualified_alpha_seconds
            for window_second in range(alpha_window_start, second + 1)
        )
        warning, warning_reason = classify_warning(
            eye_window_count, alpha_window_count, config
        )
        alpha_record = alpha_by_second.get(second)
        alpha_valid = bool(
            alpha_record is not None
            and not alpha_record.excluded_by_eye
            and alpha_record.alpha_power is not None
        )
        alpha_qualified = second in qualified_alpha_seconds
        eye_alert = eye_window_count < config.eye_alert_threshold
        alpha_alert = alpha_window_count >= config.alpha_alert_threshold
        features.append(
            FunctionTwoFeatureRecord(
                second=second,
                seconds_before_fatigue=(
                    target_second - second if target_second is not None else None
                ),
                eye_detected=second in eye_second_set,
                eye_window_count=eye_window_count,
                eye_alert=eye_alert,
                alpha_valid=alpha_valid,
                theta_power=(alpha_record.theta_power if alpha_record else None),
                alpha_power=(alpha_record.alpha_power if alpha_record else None),
                beta_power=(alpha_record.beta_power if alpha_record else None),
                alpha_qualified=alpha_qualified,
                alpha_window_count=alpha_window_count,
                alpha_alert=alpha_alert,
                warning=warning,
                warning_reason=warning_reason,
                target_fatigue=target_second == second,
            )
        )
    return features


def _find_fp2_channel(channel_names: Sequence[str]) -> str:
    for channel_name in channel_names:
        if str(channel_name).strip().casefold() == "fp2":
            return str(channel_name)
    for channel_name in channel_names:
        if "fp2" in str(channel_name).casefold():
            return str(channel_name)
    raise ValueError("EDF中找不到FP2通道。")


def _autosize_worksheet(worksheet) -> None:
    for column_index in range(1, worksheet.max_column + 1):
        values = [
            str(worksheet.cell(row=row, column=column_index).value or "")
            for row in range(1, worksheet.max_row + 1)
        ]
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(
            max(max(map(len, values), default=0) + 2, 12), 36
        )


def write_function_two_workbook(
    result: FunctionTwoResult,
    output_path: str | Path,
    config: FunctionTwoConfig = DEFAULT_CONFIG,
) -> Path:
    """Write the Function Two summary, per-second features, and RT ground truth."""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    summary = workbook.active
    summary.title = "功能二摘要"
    summary.append(["項目", "數值"])
    summary_rows = [
        ("record_id", result.record_id),
        ("功能二狀態", result.status),
        ("分析開始秒", result.analysis_start_second),
        ("分析結束秒", result.analysis_end_second),
        ("個人化RT疲勞門檻", result.personalized_rt_threshold),
        ("Alpha Power中位數Baseline", result.alpha_median_baseline),
        ("眼動窗口_秒", config.eye_window_seconds),
        ("眼動警報條件", f"EyeWindow < {config.eye_alert_threshold}"),
        ("Alpha窗口_秒", config.alpha_window_seconds),
        ("Alpha警報條件", f"AlphaWindow >= {config.alpha_alert_threshold}"),
        ("警報組合", "OR"),
        (
            "第一個疲勞事件_秒",
            result.target_event.event_second if result.target_event else None,
        ),
        (
            "第一個疲勞事件RT_秒",
            result.target_event.reaction_time if result.target_event else None,
        ),
        ("第一次警報_秒", result.first_warning_second),
        ("第一次警報原因", result.first_warning_reason),
        (
            "成功提前預測",
            "是"
            if result.prediction_success is True
            else "否"
            if result.prediction_success is False
            else None,
        ),
        ("提前秒數", result.lead_seconds),
        ("逐秒資料筆數", len(result.features)),
    ]
    for row in summary_rows:
        summary.append(list(row))

    feature_sheet = workbook.create_sheet("功能二逐秒特徵")
    feature_sheet.append(
        [
            "秒數",
            "距離疲勞事件_秒",
            "該秒有眼動",
            f"EyeWindow{config.eye_window_seconds}",
            f"EyeWindow<{config.eye_alert_threshold}",
            "該秒Alpha有效",
            "Theta Power",
            "Alpha Power",
            "Beta Power",
            "符合Alpha條件",
            f"AlphaWindow{config.alpha_window_seconds}",
            f"AlphaWindow>={config.alpha_alert_threshold}",
            "警報",
            "警報原因",
            "第一個疲勞事件",
        ]
    )
    for feature in result.features:
        feature_sheet.append(
            [
                feature.second,
                feature.seconds_before_fatigue,
                feature.eye_detected,
                feature.eye_window_count,
                feature.eye_alert,
                feature.alpha_valid,
                feature.theta_power,
                feature.alpha_power,
                feature.beta_power,
                feature.alpha_qualified,
                feature.alpha_window_count,
                feature.alpha_alert,
                feature.warning,
                feature.warning_reason,
                feature.target_fatigue,
            ]
        )

    rt_sheet = workbook.create_sheet("300秒後RT事件")
    rt_sheet.append(
        [
            "事件編號",
            "事件秒數",
            "Reaction Time_秒",
            "達個人化疲勞門檻",
            "第一個疲勞事件",
        ]
    )
    target_index = result.target_event.event_index if result.target_event else None
    for event in result.post_baseline_events:
        rt_sheet.append(
            [
                event.event_index,
                event.event_second,
                event.reaction_time,
                (
                    result.personalized_rt_threshold is not None
                    and event.reaction_time >= result.personalized_rt_threshold
                ),
                event.event_index == target_index,
            ]
        )
    for cell in rt_sheet["C"][1:]:
        cell.number_format = "0.0"

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        _autosize_worksheet(worksheet)
    workbook.save(output)
    return output


def save_pre_fatigue_plot(
    result: FunctionTwoResult,
    output_path: str | Path,
    config: FunctionTwoConfig = DEFAULT_CONFIG,
) -> Path | None:
    """Plot EyeWindow30 and AlphaWindow10 for the 30 seconds before target."""
    if result.target_event is None or not result.features:
        return None

    target_second = result.target_event.event_second
    plot_start = target_second - config.plot_seconds_before_fatigue
    plot_features = [
        feature for feature in result.features if feature.second >= plot_start
    ]
    if not plot_features:
        return None

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    relative_seconds = [feature.second - target_second for feature in plot_features]
    eye_values = [feature.eye_window_count for feature in plot_features]
    alpha_values = [feature.alpha_window_count for feature in plot_features]

    figure, (eye_axis, alpha_axis) = plt.subplots(
        2, 1, figsize=(13, 8), sharex=True
    )
    eye_axis.plot(
        relative_seconds,
        eye_values,
        color="#4472C4",
        marker="o",
        markersize=3.5,
        linewidth=1.8,
        label=f"EyeWindow{config.eye_window_seconds}",
    )
    eye_axis.axhline(
        config.eye_alert_threshold,
        color="#C00000",
        linestyle="--",
        label=f"警報門檻 < {config.eye_alert_threshold}",
    )
    eye_alert_x = [
        feature.second - target_second
        for feature in plot_features
        if feature.eye_alert
    ]
    eye_alert_y = [
        feature.eye_window_count for feature in plot_features if feature.eye_alert
    ]
    eye_axis.scatter(eye_alert_x, eye_alert_y, color="#C00000", s=28, zorder=4)
    eye_axis.set_ylabel(f"{config.eye_window_seconds}秒眼動次數")
    eye_axis.set_title("眼動窗口")
    eye_axis.grid(True, linestyle="--", alpha=0.3)

    alpha_axis.plot(
        relative_seconds,
        alpha_values,
        color="#70AD47",
        marker="o",
        markersize=3.5,
        linewidth=1.8,
        label=f"AlphaWindow{config.alpha_window_seconds}",
    )
    alpha_axis.axhline(
        config.alpha_alert_threshold,
        color="#C00000",
        linestyle="--",
        label=f"警報門檻 >= {config.alpha_alert_threshold}",
    )
    alpha_alert_x = [
        feature.second - target_second
        for feature in plot_features
        if feature.alpha_alert
    ]
    alpha_alert_y = [
        feature.alpha_window_count
        for feature in plot_features
        if feature.alpha_alert
    ]
    alpha_axis.scatter(alpha_alert_x, alpha_alert_y, color="#C00000", s=28, zorder=4)
    alpha_axis.set_ylabel(f"{config.alpha_window_seconds}秒Alpha特徵數")
    alpha_axis.set_xlabel("距離第一個疲勞事件的秒數（0 = 疲勞事件）")
    alpha_axis.set_title("Alpha窗口")
    alpha_axis.grid(True, linestyle="--", alpha=0.3)

    for axis in (eye_axis, alpha_axis):
        axis.axvline(0, color="#7030A0", linewidth=1.8, label="疲勞事件")
        axis.set_xlim(min(relative_seconds), 0.5)
        axis.legend(loc="best")

    warning_text = (
        f"第一次警報：第{result.first_warning_second}秒，"
        f"提前{result.lead_seconds}秒，原因={result.first_warning_reason}"
        if result.first_warning_second is not None
        else "第一個疲勞事件前未發出警報"
    )
    figure.suptitle(
        f"{result.record_id}：第一個疲勞事件前30秒窗口\n"
        f"疲勞事件=第{target_second}秒；{warning_text}",
        fontsize=14,
    )
    figure.tight_layout(rect=(0, 0, 1, 0.92))
    figure.savefig(output, dpi=300, bbox_inches="tight")
    plt.close(figure)
    return output


def _empty_result(
    function_one_result: FunctionOneResult,
    *,
    status: str,
    output_dir: Path,
    personalized_rt_threshold: float | None,
    alpha_median_baseline: float | None,
    config: FunctionTwoConfig,
) -> FunctionTwoResult:
    result = FunctionTwoResult(
        record_id=function_one_result.record_id,
        status=status,
        analysis_start_second=config.baseline_end_second + 1,
        analysis_end_second=config.baseline_end_second,
        personalized_rt_threshold=personalized_rt_threshold,
        alpha_median_baseline=alpha_median_baseline,
        target_event=None,
        first_warning_second=None,
        first_warning_reason=None,
        prediction_success=None,
        lead_seconds=None,
        features=(),
        post_baseline_events=(),
        eye_seconds=(),
        alpha_result=None,
        output_dir=output_dir,
    )
    write_function_two_workbook(
        result, output_dir / "function_two_results.xlsx", config
    )
    return result


def analyze_function_two(
    edf_path: str | Path,
    function_one_result: FunctionOneResult,
    output_dir: str | Path | None = None,
    config: FunctionTwoConfig = DEFAULT_CONFIG,
) -> FunctionTwoResult:
    """Run Function Two using the baseline returned by Function One."""
    path = Path(edf_path)
    if not path.is_file():
        raise FileNotFoundError(f"EDF not found: {path}")
    resolved_output_dir = (
        Path(output_dir) if output_dir is not None else function_one_result.output_dir
    )
    resolved_output_dir.mkdir(parents=True, exist_ok=True)

    personalized_threshold = function_one_result.personalized_rt_threshold
    alpha_median = (
        function_one_result.alpha_result.alpha_median
        if function_one_result.alpha_result is not None
        else None
    )
    if not function_one_result.allow_driving:
        return _empty_result(
            function_one_result,
            status="SKIPPED_FUNCTION_ONE_NOT_PASSED",
            output_dir=resolved_output_dir,
            personalized_rt_threshold=personalized_threshold,
            alpha_median_baseline=alpha_median,
            config=config,
        )
    if personalized_threshold is None:
        return _empty_result(
            function_one_result,
            status="INSUFFICIENT_RT_BASELINE",
            output_dir=resolved_output_dir,
            personalized_rt_threshold=None,
            alpha_median_baseline=alpha_median,
            config=config,
        )
    if alpha_median is None or not math.isfinite(alpha_median):
        return _empty_result(
            function_one_result,
            status="INSUFFICIENT_ALPHA_BASELINE",
            output_dir=resolved_output_dir,
            personalized_rt_threshold=personalized_threshold,
            alpha_median_baseline=alpha_median,
            config=config,
        )

    all_events = extract_reaction_time_events(path)
    post_baseline_events = tuple(
        event
        for event in all_events
        if event.event_second > config.baseline_end_second
    )
    target_event = find_first_post_baseline_fatigue_event(
        all_events,
        personalized_threshold,
        config.baseline_end_second,
    )

    raw = mne.io.read_raw_edf(path, preload=False, verbose=False)
    try:
        fp2_channel = _find_fp2_channel(raw.ch_names)
        sample_rate = float(raw.info["sfreq"])
        samples_per_second = max(int(round(sample_rate)), 1)
        recording_end_second = int(raw.n_times // samples_per_second)
        requested_end_second = (
            target_event.event_second if target_event else recording_end_second
        )
        analysis_end_second = min(requested_end_second, recording_end_second)
        analysis_start_second = config.baseline_end_second + 1
        if analysis_end_second < analysis_start_second:
            return _empty_result(
                function_one_result,
                status="INSUFFICIENT_POST_BASELINE_DATA",
                output_dir=resolved_output_dir,
                personalized_rt_threshold=personalized_threshold,
                alpha_median_baseline=alpha_median,
                config=config,
            )

        post_eye_seconds = tuple(
            detect_eye_movements(
                raw,
                [fp2_channel],
                resolved_output_dir / "eyeblink_function_two.dat",
                start_second=analysis_start_second,
                end_second=analysis_end_second,
            )
        )
    finally:
        raw.close()

    combined_eye_seconds = tuple(
        sorted(set(function_one_result.eye_seconds) | set(post_eye_seconds))
    )
    alpha_result = detect_alpha(
        path,
        resolved_output_dir / "Alpha_function_two.dat",
        [fp2_channel],
        eye_seconds=combined_eye_seconds,
        start_second=analysis_start_second,
        end_second=analysis_end_second,
        alpha_power_threshold=alpha_median,
    )
    baseline_alpha_records = (
        function_one_result.alpha_result.records
        if function_one_result.alpha_result is not None
        else ()
    )
    combined_alpha_records = (*baseline_alpha_records, *alpha_result.records)
    features = build_function_two_features(
        start_second=analysis_start_second,
        end_second=analysis_end_second,
        eye_seconds=combined_eye_seconds,
        alpha_records=combined_alpha_records,
        alpha_median_baseline=alpha_median,
        target_second=target_event.event_second if target_event else None,
        config=config,
    )
    first_warning = next((feature for feature in features if feature.warning), None)
    first_warning_second = first_warning.second if first_warning else None
    first_warning_reason = first_warning.warning_reason if first_warning else None

    lead_seconds: int | None = None
    prediction_success: bool | None = None
    if target_event is not None:
        if first_warning_second is None:
            status = "TARGET_WITHOUT_WARNING"
            prediction_success = False
        else:
            lead_seconds = target_event.event_second - first_warning_second
            prediction_success = lead_seconds > 0
            status = "TARGET_PREDICTED" if prediction_success else "WARNING_AT_TARGET"
    else:
        status = (
            "NO_TARGET_FATIGUE_WITH_WARNING"
            if first_warning_second is not None
            else "NO_TARGET_FATIGUE"
        )

    result = FunctionTwoResult(
        record_id=function_one_result.record_id,
        status=status,
        analysis_start_second=analysis_start_second,
        analysis_end_second=analysis_end_second,
        personalized_rt_threshold=personalized_threshold,
        alpha_median_baseline=alpha_median,
        target_event=target_event,
        first_warning_second=first_warning_second,
        first_warning_reason=first_warning_reason,
        prediction_success=prediction_success,
        lead_seconds=lead_seconds,
        features=tuple(features),
        post_baseline_events=tuple(
            event
            for event in post_baseline_events
            if event.event_second <= analysis_end_second
        ),
        eye_seconds=post_eye_seconds,
        alpha_result=alpha_result,
        output_dir=resolved_output_dir,
    )
    write_function_two_workbook(
        result, resolved_output_dir / "function_two_results.xlsx", config
    )
    save_pre_fatigue_plot(
        result,
        resolved_output_dir / "function_two_pre_fatigue_30s.png",
        config,
    )
    return result


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run Function One and then Function Two for one EDF."
    )
    parser.add_argument("--file", required=True, help="Input EDF path.")
    parser.add_argument("--output-dir", help="Output folder for this recording.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    function_one_result = analyze_function_one(args.file, args.output_dir)
    print(f"功能一結果：{function_one_result.decision}")
    if not function_one_result.allow_driving:
        print("功能一未通過，不進入功能二。")
        print(f"輸出資料夾：{function_one_result.output_dir}")
        return

    result = analyze_function_two(
        args.file,
        function_one_result,
        function_one_result.output_dir,
    )
    print(f"功能二狀態：{result.status}")
    print(
        "第一個疲勞事件："
        + (
            f"第{result.target_event.event_second}秒"
            if result.target_event is not None
            else "無"
        )
    )
    print(
        "第一次警報："
        + (
            f"第{result.first_warning_second}秒（{result.first_warning_reason}）"
            if result.first_warning_second is not None
            else "無"
        )
    )
    print(f"輸出資料夾：{result.output_dir}")


if __name__ == "__main__":
    main()


__all__ = [
    "DEFAULT_CONFIG",
    "FunctionTwoConfig",
    "FunctionTwoFeatureRecord",
    "FunctionTwoResult",
    "analyze_function_two",
    "build_function_two_features",
    "classify_warning",
    "find_first_post_baseline_fatigue_event",
    "save_pre_fatigue_plot",
    "write_function_two_workbook",
]
