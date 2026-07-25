"""Summarise 60 km/h estimates for events with Status reactions over one second.

The script reads every XLSX workbook in the sibling ``data`` directory.  It
keeps event rows with ``Status反應時間_秒 > 1``, then saves the corresponding
``60km_h推估時間_秒`` values and a 0.1-second frequency distribution.  Values
greater than or equal to 2.5 seconds are capped in the 2.5-second bucket.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


EVENT_SHEET = "事件分析"
EVENT_ID_COLUMN = "事件編號"
STATUS_REACTION_COLUMN = "Status反應時間_秒"
ESTIMATED_TIME_COLUMN = "60km_h推估時間_秒"
BIN_START_SECONDS = 1.0
BIN_WIDTH_SECONDS = 0.1
CAP_SECONDS = 2.5


def bucket_label(value: float) -> float:
    """Return the lower boundary of the 0.1-second bucket for a value.

    The last bucket is labelled 2.5 and contains every value at or above 2.5.
    """
    if value >= CAP_SECONDS:
        return CAP_SECONDS
    return round(
        BIN_START_SECONDS
        + np.floor((value - BIN_START_SECONDS) / BIN_WIDTH_SECONDS) * BIN_WIDTH_SECONDS,
        1,
    )


def read_selected_events(data_dir: Path) -> tuple[pd.DataFrame, list[dict[str, object]]]:
    """Read all event-analysis workbooks and retain qualifying event rows."""
    selected_frames: list[pd.DataFrame] = []
    file_summary: list[dict[str, object]] = []

    for workbook_path in sorted(data_dir.glob("*.xlsx")):
        events = pd.read_excel(workbook_path, sheet_name=EVENT_SHEET)
        required_columns = {
            EVENT_ID_COLUMN,
            STATUS_REACTION_COLUMN,
            ESTIMATED_TIME_COLUMN,
        }
        missing_columns = required_columns - set(events.columns)
        if missing_columns:
            missing_text = ", ".join(sorted(missing_columns))
            raise ValueError(f"{workbook_path.name} 缺少欄位：{missing_text}")

        events[STATUS_REACTION_COLUMN] = pd.to_numeric(
            events[STATUS_REACTION_COLUMN], errors="coerce"
        )
        events[ESTIMATED_TIME_COLUMN] = pd.to_numeric(
            events[ESTIMATED_TIME_COLUMN], errors="coerce"
        )
        selected = events.loc[
            (events[STATUS_REACTION_COLUMN] > 1)
            & events[ESTIMATED_TIME_COLUMN].notna()
        ].copy()
        selected.insert(0, "來源檔案", workbook_path.name)
        selected.insert(1, "來源工作表", EVENT_SHEET)
        selected["分佈時間桶_秒"] = selected[ESTIMATED_TIME_COLUMN].map(bucket_label)

        selected_frames.append(selected)
        file_summary.append(
            {
                "來源檔案": workbook_path.name,
                "事件總數": len(events),
                "Status反應時間_秒大於1的event數": len(selected),
                "60km_h推估時間_秒大於等於2.5的event數": int(
                    (selected[ESTIMATED_TIME_COLUMN] >= CAP_SECONDS).sum()
                ),
            }
        )

    if not selected_frames:
        raise FileNotFoundError(f"找不到 XLSX 檔案：{data_dir}")

    return pd.concat(selected_frames, ignore_index=True), file_summary


def build_distribution(selected_events: pd.DataFrame) -> pd.DataFrame:
    """Count events in every requested 0.1-second bucket."""
    labels = np.round(
        np.arange(BIN_START_SECONDS, CAP_SECONDS + BIN_WIDTH_SECONDS / 2, BIN_WIDTH_SECONDS),
        1,
    )
    counts = selected_events["分佈時間桶_秒"].value_counts().reindex(labels, fill_value=0)
    display_labels = [f"{label:.1f}" for label in labels]
    display_labels[-1] = "2.5以上"
    return pd.DataFrame(
        {
            "時間桶_秒": labels,
            "圖表標籤": display_labels,
            "event個數": counts.to_numpy(dtype=int),
        }
    )


def write_workbook(
    output_path: Path,
    selected_events: pd.DataFrame,
    distribution: pd.DataFrame,
    file_summary: list[dict[str, object]],
) -> None:
    """Save the selected event rows and frequency table in one XLSX file."""
    summary = pd.DataFrame(file_summary)
    totals = pd.DataFrame(
        [
            {
                "來源檔案": "合計",
                "事件總數": summary["事件總數"].sum(),
                "Status反應時間_秒大於1的event數": len(selected_events),
                "60km_h推估時間_秒大於等於2.5的event數": int(
                    (selected_events[ESTIMATED_TIME_COLUMN] >= CAP_SECONDS).sum()
                ),
            }
        ]
    )
    summary = pd.concat([summary, totals], ignore_index=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        selected_events.to_excel(writer, sheet_name="符合條件event", index=False)
        distribution.to_excel(writer, sheet_name="0.1秒分佈", index=False)
        summary.to_excel(writer, sheet_name="檔案統計", index=False)

        notes = pd.DataFrame(
            {
                "項目": ["篩選條件", "分箱方式", "上限處理"],
                "內容": [
                    "Status反應時間_秒 > 1",
                    "1.0 秒起，以 0.1 秒為區間；各桶採下界標示。",
                    "60km_h推估時間_秒 >= 2.5 秒全部記錄於 2.5以上 桶。",
                ],
            }
        )
        notes.to_excel(writer, sheet_name="說明", index=False)

        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            for column_cells in worksheet.columns:
                width = max(len(str(cell.value or "")) for cell in column_cells) + 2
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(width, 48)


def plot_distribution(distribution: pd.DataFrame, output_path: Path) -> None:
    """Draw the requested seconds-versus-count bar chart."""
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei", "Microsoft YaHei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    labels = distribution["圖表標籤"].tolist()
    counts = distribution["event個數"].tolist()
    positions = np.arange(len(labels))

    figure, axis = plt.subplots(figsize=(14, 7))
    axis.bar(positions, counts, width=0.82, color="#2B6CB0", edgecolor="white")
    axis.set_title("Status反應時間 > 1 秒之 60 km/h 推估時間分佈")
    axis.set_xlabel("60km_h推估時間_秒（0.1 秒區間；2.5 秒以上合併）")
    axis.set_ylabel("event 個數")
    axis.set_xticks(positions)
    axis.set_xticklabels(labels, rotation=45, ha="right")
    axis.grid(axis="y", alpha=0.25)
    axis.set_axisbelow(True)
    figure.tight_layout()
    figure.savefig(output_path, dpi=200, bbox_inches="tight")
    plt.close(figure)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=Path(__file__).with_name("data"),
        help="包含事件分析 XLSX 的資料夾。",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path(__file__).with_name("output"),
        help="輸出 XLSX 與 PNG 的資料夾。",
    )
    args = parser.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)
    selected_events, file_summary = read_selected_events(args.data_dir)
    distribution = build_distribution(selected_events)

    workbook_path = args.output_dir / "status_reaction_over_1_60kmh_distribution.xlsx"
    figure_path = args.output_dir / "status_reaction_over_1_60kmh_distribution.png"
    write_workbook(workbook_path, selected_events, distribution, file_summary)
    plot_distribution(distribution, figure_path)

    capped_count = int((selected_events[ESTIMATED_TIME_COLUMN] >= CAP_SECONDS).sum())
    print(f"讀取檔案數：{len(file_summary)}")
    print(f"符合條件的 event 數：{len(selected_events)}")
    print(f"合併至 2.5 秒以上桶的 event 數：{capped_count}")
    print(f"Excel：{workbook_path.resolve()}")
    print(f"圖表：{figure_path.resolve()}")


if __name__ == "__main__":
    main()
