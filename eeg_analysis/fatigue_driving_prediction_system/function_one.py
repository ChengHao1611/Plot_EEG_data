"""Function One: initial 300-second fatigue screening and baseline export."""

from __future__ import annotations

import argparse
import statistics
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

# When this file is launched directly, Python adds only this subdirectory to
# sys.path.  Add the repository root so absolute ``eeg_analysis.*`` imports
# work for both:
#   python eeg_analysis/fatigue_driving_prediction_system/function_one.py
#   python -m eeg_analysis.fatigue_driving_prediction_system.function_one
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

from eeg_analysis.detection.record_alpha import AlphaDetectionResult, detect_alpha
from eeg_analysis.detection.record_arousal import detect_eye_movements
from eeg_analysis.statistics_30s_alpha_eyeblink_of_fatigue.record_status_and_eyeblink_to_xlsx import (
    ReactionTimeEvent,
    extract_reaction_time_events,
    write_reaction_time_events_xlsx,
)


@dataclass(frozen=True)
class FunctionOneConfig:
    baseline_start_second: int = 1
    baseline_end_second: int = 300
    fatigue_reaction_threshold: float = 1.6
    fatigue_window_seconds: int = 60
    personalized_multiplier: float = 1.5


DEFAULT_CONFIG = FunctionOneConfig()


@dataclass(frozen=True)
class FunctionOneResult:
    record_id: str
    decision: str
    allow_driving: bool
    events: tuple[ReactionTimeEvent, ...]
    trigger_pair: tuple[ReactionTimeEvent, ReactionTimeEvent] | None
    trigger_window_start_event: ReactionTimeEvent | None
    rt_mean: float | None
    rt_median: float | None
    personalized_rt_threshold: float | None
    eye_seconds: tuple[int, ...]
    alpha_result: AlphaDetectionResult | None
    output_dir: Path


def record_id_from_edf(edf_path: str | Path) -> str:
    stem = Path(edf_path).stem
    return stem[:-4] if stem.casefold().endswith("_raw") else stem


def select_baseline_events(
    events: Sequence[ReactionTimeEvent],
    config: FunctionOneConfig = DEFAULT_CONFIG,
) -> list[ReactionTimeEvent]:
    return [
        event
        for event in events
        if config.baseline_start_second
        <= event.event_second
        <= config.baseline_end_second
    ]


def find_fatigue_trigger_pair(
    events: Sequence[ReactionTimeEvent],
    config: FunctionOneConfig = DEFAULT_CONFIG,
) -> tuple[ReactionTimeEvent, ReactionTimeEvent] | None:
    """Return fatigue events that satisfy the follow-up reaction-window rule.

    After an initial fatigue event, the next reaction event starts an inclusive
    60-second window. A fatigue event at the window start or before its end
    triggers Function One.
    """
    details = find_fatigue_trigger_details(events, config)
    if details is None:
        return None
    initial_fatigue, _, followup_fatigue = details
    return initial_fatigue, followup_fatigue


def find_fatigue_trigger_details(
    events: Sequence[ReactionTimeEvent],
    config: FunctionOneConfig = DEFAULT_CONFIG,
) -> tuple[ReactionTimeEvent, ReactionTimeEvent, ReactionTimeEvent] | None:
    """Return initial fatigue, next reaction/window start, and follow-up fatigue."""
    ordered_events = sorted(
        events,
        key=lambda event: (event.event_second, event.deviation_time),
    )
    for initial_index, initial_event in enumerate(ordered_events[:-1]):
        if initial_event.reaction_time < config.fatigue_reaction_threshold:
            continue

        window_start_event = ordered_events[initial_index + 1]
        window_end_second = (
            window_start_event.event_second + config.fatigue_window_seconds
        )
        for followup_event in ordered_events[initial_index + 1 :]:
            if followup_event.event_second > window_end_second:
                break
            if followup_event.reaction_time >= config.fatigue_reaction_threshold:
                return initial_event, window_start_event, followup_event
    return None


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


def write_function_one_workbook(
    result: FunctionOneResult,
    output_path: str | Path,
    config: FunctionOneConfig = DEFAULT_CONFIG,
) -> Path:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    summary = workbook.active
    summary.title = "功能一摘要"

    alpha_result = result.alpha_result
    summary_rows = [
        ("record_id", result.record_id),
        ("功能一結果", result.decision),
        ("允許繼續駕駛", "是" if result.allow_driving else "否"),
        ("Baseline開始秒", config.baseline_start_second),
        ("Baseline結束秒", config.baseline_end_second),
        ("固定疲勞RT門檻_秒", config.fatigue_reaction_threshold),
        ("疲勞事件窗口_秒", config.fatigue_window_seconds),
        ("前300秒RT事件數", len(result.events)),
        (
            "前300秒疲勞RT事件數",
            sum(
                event.reaction_time >= config.fatigue_reaction_threshold
                for event in result.events
            ),
        ),
        ("RT平均Baseline", result.rt_mean),
        ("RT中位數Baseline", result.rt_median),
        ("個人化RT疲勞門檻", result.personalized_rt_threshold),
        ("眼動秒數", len(result.eye_seconds) if result.allow_driving else None),
        (
            "眼動排除秒數",
            alpha_result.excluded_eye_seconds if alpha_result is not None else None,
        ),
        (
            "有效FFT秒數",
            alpha_result.valid_fft_seconds if alpha_result is not None else None,
        ),
        (
            "符合Alpha條件秒數",
            len(alpha_result.alpha_seconds) if alpha_result is not None else None,
        ),
        ("Alpha Power平均Baseline", alpha_result.alpha_mean if alpha_result else None),
        ("Alpha Power中位數Baseline", alpha_result.alpha_median if alpha_result else None),
        (
            "觸發疲勞事件1_秒",
            result.trigger_pair[0].event_second if result.trigger_pair else None,
        ),
        (
            "後續60秒窗口開始_秒",
            result.trigger_window_start_event.event_second
            if result.trigger_window_start_event
            else None,
        ),
        (
            "觸發疲勞事件2_秒",
            result.trigger_pair[1].event_second if result.trigger_pair else None,
        ),
    ]
    summary.append(["項目", "數值"])
    for row in summary_rows:
        summary.append(list(row))

    event_sheet = workbook.create_sheet("Reaction Time事件")
    event_sheet.append(
        [
            "事件編號",
            "偏移Status",
            "偏移開始時間_秒",
            "導正開始時間_秒",
            "向上取整事件秒數",
            "Reaction Time_秒",
            "RT>=1.6",
            "後續60秒窗口起點",
            "觸發功能一",
        ]
    )
    trigger_ids = {
        event.event_index for event in result.trigger_pair
    } if result.trigger_pair else set()
    for event in result.events:
        event_sheet.append(
            [
                event.event_index,
                event.deviation_status,
                event.deviation_time,
                event.correction_start_time,
                event.event_second,
                event.reaction_time,
                event.reaction_time >= config.fatigue_reaction_threshold,
                result.trigger_window_start_event is not None
                and event.event_index == result.trigger_window_start_event.event_index,
                event.event_index in trigger_ids,
            ]
        )

    if alpha_result is not None:
        alpha_sheet = workbook.create_sheet("Alpha Power逐秒")
        alpha_sheet.append(
            [
                "秒數",
                "眼動排除",
                "Theta Power",
                "Alpha Power",
                "Beta Power",
                "Alpha>Theta且Alpha>Beta",
            ]
        )
        for record in alpha_result.records:
            alpha_sheet.append(
                [
                    record.second,
                    record.excluded_by_eye,
                    record.theta_power,
                    record.alpha_power,
                    record.beta_power,
                    record.alpha_qualified,
                ]
            )

    for cell in event_sheet["F"][1:]:
        cell.number_format = "0.0"
    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        _autosize_worksheet(worksheet)
    workbook.save(output)
    return output


def save_rt_validation_plot(
    result: FunctionOneResult,
    output_path: str | Path,
    config: FunctionOneConfig = DEFAULT_CONFIG,
) -> Path:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    figure, axis = plt.subplots(figsize=(13, 6.5))
    normal_events = [
        event
        for event in result.events
        if event.reaction_time < config.fatigue_reaction_threshold
    ]
    fatigue_events = [
        event
        for event in result.events
        if event.reaction_time >= config.fatigue_reaction_threshold
    ]
    if normal_events:
        axis.scatter(
            [event.event_second for event in normal_events],
            [event.reaction_time for event in normal_events],
            color="#4472C4",
            label="RT < 1.6秒",
            zorder=3,
        )
    if fatigue_events:
        axis.scatter(
            [event.event_second for event in fatigue_events],
            [event.reaction_time for event in fatigue_events],
            color="#C00000",
            label="RT >= 1.6秒",
            zorder=4,
        )
    axis.axhline(
        config.fatigue_reaction_threshold,
        color="#C00000",
        linestyle="--",
        linewidth=1.5,
        label="固定疲勞門檻 1.6秒",
    )
    if result.personalized_rt_threshold is not None:
        axis.axhline(
            result.personalized_rt_threshold,
            color="#70AD47",
            linestyle=":",
            linewidth=1.8,
            label=f"個人化門檻 {result.personalized_rt_threshold:.3f}秒",
        )
    if result.trigger_pair is not None and result.trigger_window_start_event is not None:
        first, second = result.trigger_pair
        window_start_second = result.trigger_window_start_event.event_second
        axis.axvspan(
            window_start_second,
            window_start_second + config.fatigue_window_seconds,
            color="#F4CCCC",
            alpha=0.45,
            label="下一反應事件起算的60秒窗口",
        )
        for event in (first, second):
            axis.annotate(
                f"{event.event_second}s, RT={event.reaction_time:.1f}s",
                (event.event_second, event.reaction_time),
                xytext=(6, 8),
                textcoords="offset points",
                fontsize=9,
            )

    detail_lines = [
        f"結果：{result.decision}",
        f"RT事件數：{len(result.events)}",
    ]
    if result.rt_mean is not None:
        detail_lines.append(f"RT平均：{result.rt_mean:.3f}秒")
    if result.rt_median is not None:
        detail_lines.append(f"RT中位數：{result.rt_median:.3f}秒")
    axis.text(
        0.99,
        0.97,
        "\n".join(detail_lines),
        transform=axis.transAxes,
        ha="right",
        va="top",
        bbox={"boxstyle": "round", "facecolor": "white", "alpha": 0.85},
    )
    axis.set_xlim(config.baseline_start_second, config.baseline_end_second)
    axis.set_xlabel("記錄秒數（向上取整）")
    axis.set_ylabel("Reaction Time（秒）")
    axis.set_title(f"{result.record_id}：前300秒Reaction Time驗證圖")
    axis.grid(True, linestyle="--", alpha=0.3)
    axis.legend(loc="upper left")
    figure.tight_layout()
    figure.savefig(output, dpi=300, bbox_inches="tight")
    plt.close(figure)
    return output


def analyze_function_one(
    edf_path: str | Path,
    output_dir: str | Path | None = None,
    config: FunctionOneConfig = DEFAULT_CONFIG,
) -> FunctionOneResult:
    path = Path(edf_path)
    if not path.is_file():
        raise FileNotFoundError(f"EDF not found: {path}")
    record_id = record_id_from_edf(path)
    resolved_output_dir = (
        Path(output_dir)
        if output_dir is not None
        else Path("data")
        / "derived"
        / "fatigue_driving_prediction_system"
        / record_id
    )
    resolved_output_dir.mkdir(parents=True, exist_ok=True)

    all_events = extract_reaction_time_events(path)
    events = select_baseline_events(all_events, config)
    write_reaction_time_events_xlsx(
        events, resolved_output_dir / "reaction_time_events.xlsx"
    )
    trigger_details = find_fatigue_trigger_details(events, config)
    trigger_pair = (
        (trigger_details[0], trigger_details[2])
        if trigger_details is not None
        else None
    )
    trigger_window_start_event = trigger_details[1] if trigger_details else None
    if not events:
        decision = "INSUFFICIENT_DATA"
        allow_driving = False
    elif trigger_pair is not None:
        decision = "FATIGUE"
        allow_driving = False
    else:
        decision = "NON_FATIGUE"
        allow_driving = True

    rt_mean: float | None = None
    rt_median: float | None = None
    personalized_threshold: float | None = None
    eye_seconds: tuple[int, ...] = ()
    alpha_result: AlphaDetectionResult | None = None

    if allow_driving:
        rt_values = [event.reaction_time for event in events]
        rt_mean = statistics.mean(rt_values)
        rt_median = statistics.median(rt_values)
        personalized_threshold = min(
            config.fatigue_reaction_threshold,
            rt_mean * config.personalized_multiplier,
        )

        raw = mne.io.read_raw_edf(path, preload=False, verbose=False)
        fp2_channel = _find_fp2_channel(raw.ch_names)
        eye_output = resolved_output_dir / "eyeblink.dat"
        eye_seconds = tuple(
            detect_eye_movements(
                raw,
                [fp2_channel],
                eye_output,
                start_second=config.baseline_start_second,
                end_second=config.baseline_end_second,
            )
        )
        alpha_result = detect_alpha(
            path,
            resolved_output_dir / "Alpha.dat",
            [fp2_channel],
            eye_seconds=eye_seconds,
            start_second=config.baseline_start_second,
            end_second=config.baseline_end_second,
        )

    result = FunctionOneResult(
        record_id=record_id,
        decision=decision,
        allow_driving=allow_driving,
        events=tuple(events),
        trigger_pair=trigger_pair,
        trigger_window_start_event=trigger_window_start_event,
        rt_mean=rt_mean,
        rt_median=rt_median,
        personalized_rt_threshold=personalized_threshold,
        eye_seconds=eye_seconds,
        alpha_result=alpha_result,
        output_dir=resolved_output_dir,
    )
    write_function_one_workbook(
        result, resolved_output_dir / "function_one_results.xlsx", config
    )
    save_rt_validation_plot(
        result, resolved_output_dir / "rt_validation.png", config
    )
    return result


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run Function One of the fatigue-driving prediction system."
    )
    parser.add_argument("--file", required=True, help="Input EDF path.")
    parser.add_argument("--output-dir", help="Output folder for this recording.")
    parser.add_argument(
        "--function-one-only",
        action="store_true",
        help="Stop after Function One instead of continuing to Function Two.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    result = analyze_function_one(args.file, args.output_dir)
    print(f"功能一結果：{result.decision}")
    print(f"允許繼續駕駛：{'是' if result.allow_driving else '否'}")
    if result.allow_driving and not args.function_one_only:
        from eeg_analysis.fatigue_driving_prediction_system.function_two import (
            analyze_function_two,
        )

        function_two_result = analyze_function_two(
            args.file,
            result,
            result.output_dir,
        )
        print(f"功能二狀態：{function_two_result.status}")
        if function_two_result.target_event is not None:
            print(
                "第一個疲勞事件："
                f"第{function_two_result.target_event.event_second}秒"
            )
        else:
            print("第一個疲勞事件：無")
        if function_two_result.first_warning_second is not None:
            print(
                "第一次警報："
                f"第{function_two_result.first_warning_second}秒"
                f"（{function_two_result.first_warning_reason}）"
            )
        else:
            print("第一次警報：無")
    print(f"輸出資料夾：{result.output_dir}")


if __name__ == "__main__":
    main()
