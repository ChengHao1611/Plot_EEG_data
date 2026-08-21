from pathlib import Path
import math
import sys
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import matplotlib.pyplot as plt
import threading

if __package__ in {None, ""}:
    project_root = Path(__file__).resolve().parents[2]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

from eeg_analysis.fatigue_driving_prediction_system.physiological_fatigue import (
    TRAINING_POOLED_ALPHA_SCALE,
    TRAINING_POOLED_EYE_SCALE,
)

try:
    from .observe_algorithm import (
        FatigueAlgorithmConfig,
        classify_actual_states,
        predict_fatigue_states,
    )
except ImportError:
    # 支援直接以 python observe.py 執行。
    from observe_algorithm import (
        FatigueAlgorithmConfig,
        classify_actual_states,
        predict_fatigue_states,
    )

# 設定字體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False


# 事件 Excel 可以使用舊版英文欄位，或由逐秒統計表輸出的中文欄位。
EVENT_COLUMN_ALIASES = {
    "second": ("second", "秒數", "事件秒數", "向上取整事件秒數"),
    "react_time": (
        "react_time",
        "事件反應時間",
        "Local RT_秒",
        "Reaction Time_秒",
    ),
}

PHASE_2_START_SECOND = 300
PLOT_SECONDS_AFTER_FIRST_FATIGUE = 500
REACTION_TIME_CAP_SECONDS = 10.0
REACTION_TIME_OVERFLOW_Y = 10.8
REACTION_TIME_AXIS_MAX = 11.5


def configure_reaction_time_coordinate_display(
    ax_top,
    ax_twin,
    reaction_events_df: pd.DataFrame,
) -> None:
    """Label toolbar coordinates and snap them to nearby RT event markers."""
    event_seconds = reaction_events_df["second"].to_numpy(dtype=float)
    actual_reaction_times = reaction_events_df["react_time"].to_numpy(
        dtype=float
    )
    displayed_reaction_times = np.where(
        actual_reaction_times > REACTION_TIME_CAP_SECONDS,
        REACTION_TIME_OVERFLOW_Y,
        actual_reaction_times,
    )

    def format_reaction_time_coordinate(x, y):
        if event_seconds.size and math.isfinite(x) and math.isfinite(y):
            cursor_pixel = ax_twin.transData.transform((x, y))
            event_pixels = ax_twin.transData.transform(
                np.column_stack((event_seconds, displayed_reaction_times))
            )
            distances = np.hypot(
                event_pixels[:, 0] - cursor_pixel[0],
                event_pixels[:, 1] - cursor_pixel[1],
            )
            nearest_index = int(np.argmin(distances))
            if distances[nearest_index] <= 10:
                return (
                    "Right RT (x, y) = ("
                    f"{event_seconds[nearest_index]:.2f}, "
                    f"{actual_reaction_times[nearest_index]:.2f})"
                )
        return f"Right RT (x, y) = ({x:.2f}, {y:.2f})"

    def format_accumulation_coordinate(x, y):
        display_coordinate = ax_top.transData.transform((x, y))
        rt_x, rt_y = ax_twin.transData.inverted().transform(
            display_coordinate
        )
        left_coordinate = (
            f"Left Alpha/Eye (x, y) = ({x:.2f}, {y:.2f})"
        )
        right_coordinate = format_reaction_time_coordinate(rt_x, rt_y)
        return f"{left_coordinate} | {right_coordinate}"

    ax_top.format_coord = format_accumulation_coordinate
    ax_twin.format_coord = format_reaction_time_coordinate


def calculate_plot_end_second(
    first_fatigue_time: float | None,
    all_time: float,
    seconds_after_fatigue: int = PLOT_SECONDS_AFTER_FIRST_FATIGUE,
) -> float:
    """Return ``min(first fatigue + tail, all time)`` for the plot x-limit."""
    if first_fatigue_time is None:
        return float(all_time)
    return min(
        float(first_fatigue_time) + seconds_after_fatigue,
        float(all_time),
    )


def normalize_event_columns(events_df: pd.DataFrame) -> pd.DataFrame:
    """將支援的中英文事件欄位名稱統一為程式內部欄位名稱。"""
    columns_by_normalized_name = {}
    for column in events_df.columns:
        normalized_name = str(column).strip().casefold()
        columns_by_normalized_name.setdefault(normalized_name, column)

    rename_columns = {}
    missing_columns = []
    for internal_name, aliases in EVENT_COLUMN_ALIASES.items():
        source_column = next(
            (
                columns_by_normalized_name[alias.casefold()]
                for alias in aliases
                if alias.casefold() in columns_by_normalized_name
            ),
            None,
        )
        if source_column is None:
            missing_columns.append(" / ".join(aliases))
        elif source_column != internal_name:
            rename_columns[source_column] = internal_name

    if missing_columns:
        raise ValueError(
            "事件 Excel 缺少必要欄位："
            f"{', '.join(missing_columns)}。"
            "請使用支援的事件秒數與Local RT欄位。"
        )

    return events_df.rename(columns=rename_columns)


def find_first_fatigue_onset(
    state_df: pd.DataFrame,
    state_column: str,
    start_second: float = PHASE_2_START_SECOND,
) -> float | None:
    """找出指定時間後第一次由清醒（0）轉為疲勞（1）的秒數。"""
    if state_df.empty or state_column not in state_df.columns:
        return None

    states = state_df[["second", state_column]].copy()
    states["second"] = pd.to_numeric(states["second"], errors="coerce")
    states[state_column] = pd.to_numeric(states[state_column], errors="coerce")
    states = states.dropna(subset=["second", state_column]).sort_values(
        "second", kind="stable"
    )

    is_fatigued = states[state_column].eq(1)
    fatigue_onsets = states.loc[
        is_fatigued & ~is_fatigued.shift(fill_value=False)
    ]
    fatigue_onsets = fatigue_onsets.loc[
        fatigue_onsets["second"] >= start_second, "second"
    ]
    if fatigue_onsets.empty:
        return None
    return float(fatigue_onsets.iloc[0])


def format_plot_second(second: float) -> str:
    """秒數若為整數則不顯示小數點。"""
    return f"{second:g}"


def annotate_fatigue_timing(
    ax,
    state_df: pd.DataFrame,
    pred_df: pd.DataFrame,
    *,
    show_physiological: bool = True,
) -> tuple[float | None, float | None]:
    """在主圖標出 phase 2、兩種首次疲勞與兩者之間的 lead time。"""
    if show_physiological:
        ax.axvline(
            PHASE_2_START_SECOND,
            color="#4472C4",
            linestyle="--",
            linewidth=1.5,
            label=f"Phase 2 Start ({PHASE_2_START_SECOND} s)",
            zorder=4,
        )

    physiological_second = (
        find_first_fatigue_onset(pred_df, "pred_state")
        if show_physiological
        else None
    )
    behavioral_second = find_first_fatigue_onset(
        state_df,
        "state_val",
        start_second=0,
    )

    behavioral_label = "First Behavioral Fatigue"
    if (
        behavioral_second is not None
        and behavioral_second <= PHASE_2_START_SECOND
    ):
        behavioral_label += " (Phase 1)"

    fatigue_lines = (
        (physiological_second, "#008C95", "First Physiological Fatigue"),
        (behavioral_second, "#C00000", behavioral_label),
    )
    available_lines = [item for item in fatigue_lines if item[0] is not None]
    left_line_index = None
    if len(available_lines) == 2:
        left_line_index = min(
            range(len(available_lines)), key=lambda index: available_lines[index][0]
        )

    for line_index, (second, color, label) in enumerate(available_lines):
        second_text = format_plot_second(second)
        ax.axvline(
            second,
            color=color,
            linestyle="-.",
            linewidth=1.8,
            label=f"{label} ({second_text} s)",
            zorder=4,
        )

        # 文字放在兩條疲勞線的外側，避免與中間的 lead time 重疊。
        is_left_line = line_index == left_line_index
        x_offset = -5 if is_left_line else 5
        horizontal_alignment = "right" if is_left_line else "left"
        ax.annotate(
            f"{label}\n{second_text} s",
            xy=(second, 0.62),
            xycoords=ax.get_xaxis_transform(),
            xytext=(x_offset, 0),
            textcoords="offset points",
            ha=horizontal_alignment,
            va="top",
            fontsize=9,
            fontweight="bold",
            color=color,
            bbox={
                "boxstyle": "round,pad=0.25",
                "facecolor": "white",
                "edgecolor": color,
                "linewidth": 1,
                "alpha": 1,
            },
            zorder=8,
        )

    if (
        physiological_second is not None
        and behavioral_second is not None
        and behavioral_second > PHASE_2_START_SECOND
    ):
        lead_time = behavioral_second - physiological_second
        lead_time_text = format_plot_second(lead_time)
        midpoint = (physiological_second + behavioral_second) / 2

        if physiological_second != behavioral_second:
            ax.annotate(
                "",
                xy=(behavioral_second, 0.47),
                xytext=(physiological_second, 0.47),
                xycoords=ax.get_xaxis_transform(),
                textcoords=ax.get_xaxis_transform(),
                arrowprops={
                    "arrowstyle": "<->",
                    "color": "#404040",
                    "linewidth": 2,
                    "shrinkA": 0,
                    "shrinkB": 0,
                },
                zorder=8,
            )
        ax.text(
            midpoint,
            0.49,
            f"Lead Time: {lead_time_text} s",
            transform=ax.get_xaxis_transform(),
            ha="center",
            va="bottom",
            fontsize=10,
            fontweight="bold",
            color="#404040",
            bbox={
                "boxstyle": "round,pad=0.25",
                "facecolor": "white",
                "edgecolor": "#808080",
                "alpha": 1,
            },
            zorder=9,
        )

    return physiological_second, behavioral_second


class EEGBrowserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("EEG 駕駛狀態分析系統")
        self.root.geometry("760x360")

        self.path_xlsx = tk.StringVar()
        self.path_alpha = tk.StringVar()
        self.path_eye = tk.StringVar()
        self.personalized_rt_threshold = tk.StringVar(value="1.6")
        self.pooled_alpha_scale = tk.StringVar(
            value=f"{TRAINING_POOLED_ALPHA_SCALE:g}"
        )
        self.pooled_eye_scale = tk.StringVar(
            value=f"{TRAINING_POOLED_EYE_SCALE:g}"
        )
        self.create_layout()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_layout(self):
        # 1. 數據來源設定
        top_frame = tk.LabelFrame(self.root, text="數據來源", padx=15, pady=10)
        top_frame.pack(side='top', fill='x', padx=20, pady=10)

        self.add_file_row(
            top_frame,
            "事件 Excel (.xlsx):",
            self.path_xlsx,
            True,
        )
        self.add_file_row(top_frame, "Alpha 數據 (.dat):", self.path_alpha, False)
        self.add_file_row(top_frame, "眼動 數據 (.dat):", self.path_eye, False)
        threshold_row = tk.Frame(top_frame)
        threshold_row.pack(fill='x', pady=2)
        tk.Label(threshold_row, text="個人化RT門檻:", width=15).pack(side='left')
        tk.Entry(
            threshold_row,
            textvariable=self.personalized_rt_threshold,
        ).pack(side='left', fill='x', expand=True, padx=5)
        tk.Label(threshold_row, text="秒（手動輸入）").pack(side='right')
        self.add_number_row(
            top_frame,
            "Pooled Alpha Scale:",
            self.pooled_alpha_scale,
            "（手動輸入）",
        )
        self.add_number_row(
            top_frame,
            "Pooled Eye Scale:",
            self.pooled_eye_scale,
            "（手動輸入）",
        )

        # 2. 進度條
        self.progress_frame = tk.Frame(self.root, padx=20)
        self.progress_frame.pack(fill='x')
        self.progress_label = tk.Label(self.progress_frame, text="準備就緒", font=('Arial', 9))
        self.progress_label.pack(side='top', anchor='w')
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill='x', pady=5)

        self.run_btn = tk.Button(self.root, text="執行駕駛狀態分析並匯出結果", command=self.start_thread, 
                                 bg='#2196F3', fg='white', font=('Arial', 10, 'bold'), height=2)
        self.run_btn.pack(fill='x', padx=20, pady=5)

    def start_thread(self):
        # 關鍵：在主執行緒先把資料取出來，避免背景執行緒存取 UI 變數
        plt.close('all')
        xlsx_val = self.path_xlsx.get()
        alpha_val = self.path_alpha.get()
        eye_val = self.path_eye.get()
        if not all([xlsx_val, alpha_val, eye_val]):
            messagebox.showwarning("警告", "請先選擇所有必要的數據文件。")
            return
        try:
            rt_threshold = float(self.personalized_rt_threshold.get().strip())
            pooled_alpha_scale = float(self.pooled_alpha_scale.get().strip())
            pooled_eye_scale = float(self.pooled_eye_scale.get().strip())
            input_values = (
                rt_threshold,
                pooled_alpha_scale,
                pooled_eye_scale,
            )
            if any(not math.isfinite(value) or value <= 0 for value in input_values):
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "警告",
                "個人化RT門檻與兩個Pooled Scale都必須是大於0的有限數值。",
            )
            return

        self.run_btn.config(state='disabled')
        # 把參數傳進去，不要讓 Thread 直接讀 self.path_xlsx
        threading.Thread(
            target=self.process_data,
            args=(
                xlsx_val,
                alpha_val,
                eye_val,
                rt_threshold,
                pooled_alpha_scale,
                pooled_eye_scale,
            ),
            daemon=True,
        ).start()

    def process_data(
        self,
        xlsx_p,
        alpha_p,
        eye_p,
        rt_threshold,
        pooled_alpha_scale,
        pooled_eye_scale,
    ):
        try:
            self.safe_update_status("數據讀取中...")
            
            # 讀取數據 (使用傳進來的字串路徑)
            df_main = normalize_event_columns(pd.read_excel(xlsx_p))
            df_main['second'] = pd.to_numeric(df_main['second'], errors='coerce')
            df_main['react_time'] = pd.to_numeric(df_main['react_time'], errors='coerce')
            df_main = df_main.dropna(subset=['second'])

            alpha_counts = self.load_dat_counts(alpha_p)
            eye_counts = self.load_dat_counts(eye_p)
            
            # 建立主時間軸
            max_s = int(max(df_main['second'].max(), alpha_counts.index.max() or 0, eye_counts.index.max() or 0))
            master_df = pd.DataFrame({'second': range(max_s + 1)})
            master_df = master_df.merge(alpha_counts.rename('a_raw'), left_on='second', right_index=True, how='left').fillna(0)
            master_df = master_df.merge(eye_counts.rename('e_raw'), left_on='second', right_index=True, how='left').fillna(0)
            
            # rolling windows (改)
            master_df['alpha_sum'] = master_df['a_raw'].rolling(window=30, min_periods=30).sum()
            master_df['eye_sum'] = master_df['e_raw'].rolling(window=30, min_periods=30).sum()

            self.process_driving_state(
                df_main,
                master_df,
                xlsx_p,
                rt_threshold,
                pooled_alpha_scale,
                pooled_eye_scale,
            )

        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: messagebox.showerror("錯誤", f"處理中斷: {err_msg}"))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state='normal'))

    def safe_update_status(self, text):
        # update UI from main thread
        self.root.after(0, lambda: self.progress_label.config(text=text))

    def process_driving_state(
        self,
        df_main,
        master_df,
        event_xlsx_path,
        rt_threshold,
        pooled_alpha_scale,
        pooled_eye_scale,
    ):
        """生成駕駛狀態隨時間演變圖（生理預測與實際反應時間）。"""
        self.safe_update_status("正在生成狀態演變圖...")

        # 輸出名稱以事件 Excel 為準；例如 s01_060926_1n_raw.xlsx
        # 會產生本程式資料夾下的 data/observe/s01_060926_1n_observe.png。
        record_name = Path(event_xlsx_path).stem
        if record_name.lower().endswith("_raw"):
            record_name = record_name[:-4]
        figure_output_path = (
            Path(__file__).resolve().parent
            / "data"
            / "observe"
            / f"{record_name}_observe.png"
        )
        
        # debug excel file
        def export_to_excel(pred_df, state_df):
            import openpyxl

            behavior_mapping = (
                state_df.drop_duplicates(subset=['second'], keep='last')
                .set_index('second')
                .to_dict(orient='index')
            )

            wb = openpyxl.Workbook()
            summary = wb.active
            summary.title = "Baseline Summary"
            summary.append(["Item", "Value"])
            alpha_baseline = pred_df.attrs.get("alpha_baseline")
            eye_baseline = pred_df.attrs.get("eye_baseline")
            summary_rows = [
                ("Personalized RT Threshold", rt_threshold),
                (
                    "Behavioral Global RT Window Seconds",
                    state_df.attrs["global_rt_window_seconds"],
                ),
                (
                    "Critical Local RT Threshold",
                    state_df.attrs["critical_local_rt_threshold"],
                ),
                (
                    "Behavioral Rule",
                    state_df.attrs["behavioral_rule"],
                ),
                ("Alpha Window Seconds", 30),
                ("Eye Window Seconds", 30),
                (
                    "Physiological Rule",
                    "Z-Alpha>=h AND Z-Eye>=h",
                ),
                ("h", algorithm_config.score_threshold),
                (
                    "Confirmation Seconds",
                    algorithm_config.confirmation_seconds,
                ),
                ("Input Pooled Alpha Scale", pooled_alpha_scale),
                ("Input Pooled Eye Scale", pooled_eye_scale),
                ("Pooled Scale Source", "Manual Observe input"),
                ("First Behavioral Fatigue Second", first_behavioral_second),
                ("Phase One Blocked", state_df.attrs["phase_one_blocked"]),
                ("Plot End Second", plot_end_second),
                ("Alpha Median", alpha_baseline.median if alpha_baseline else None),
                ("Alpha MAD", alpha_baseline.mad if alpha_baseline else None),
                ("Alpha IQR", alpha_baseline.iqr if alpha_baseline else None),
                ("Alpha Scale", alpha_baseline.scale if alpha_baseline else None),
                ("Alpha Scale Method", alpha_baseline.scale_method if alpha_baseline else None),
                ("Eye Median", eye_baseline.median if eye_baseline else None),
                ("Eye MAD", eye_baseline.mad if eye_baseline else None),
                ("Eye IQR", eye_baseline.iqr if eye_baseline else None),
                ("Eye Scale", eye_baseline.scale if eye_baseline else None),
                ("Eye Scale Method", eye_baseline.scale_method if eye_baseline else None),
            ]
            for item in summary_rows:
                summary.append(item)

            ws = wb.create_sheet("Driving State Analysis Results")
            ws.title = "Driving State Analysis Results"
            ws.views.sheetView[0].showGridLines = True
            headers = [
                "Second", "Eye30", "Alpha30", "Z_Eye", "Z_Alpha",
                "Joint_Z_Min",
                f"Both_Z>={algorithm_config.score_threshold:g}",
                "Consecutive_Seconds", "React_time",
                "Personalized_RT_Threshold", "Eye_Alarm", "Alpha_Alarm",
                "Pred_State", "Behavioral_State", "Global_RT",
                "Local>=Threshold", "Global>=Threshold", "Sustained_Fatigue",
                "Critical_Lapse", "Behavioral_Trigger", "Trigger_Reason",
            ]
            ws.append(headers)

            for _, row in pred_df.iterrows():
                sec = row['second']
                behavior = behavior_mapping.get(int(sec))
                react_t = behavior.get('react_time') if behavior else None

                ws.append([
                    int(sec), 
                    row['eye_sum'], 
                    row['alpha_sum'], 
                    row['z_eye'],
                    row['z_alpha'],
                    row['fatigue_score'],
                    int(row['score_above_threshold']),
                    int(row['consecutive_seconds']),
                    react_t if react_t is not None else "", 
                    rt_threshold,
                    int(row['eye_alarm']), 
                    int(row['alpha_alarm']),
                    int(row['pred_state']),
                    int(behavior['state_val']) if behavior else "",
                    behavior['global_rt'] if behavior else "",
                    int(behavior['local_exceed']) if behavior else "",
                    int(behavior['global_exceed']) if behavior else "",
                    int(behavior['sustained_fatigue']) if behavior else "",
                    int(behavior['critical_lapse']) if behavior else "",
                    int(behavior['behavioral_fatigue']) if behavior else "",
                    behavior['trigger_reason'] if behavior else "",
                ])
            excel_filename = figure_output_path.with_suffix(".xlsx")
            excel_filename.parent.mkdir(parents=True, exist_ok=True)
            wb.save(excel_filename)
            print(f"📊 Excel 檔案已成功儲存至: {excel_filename}")

        # 疲勞判斷完全由 observe_algorithm.py 負責；此檔案只處理 UI、繪圖與匯出。
        algorithm_config = FatigueAlgorithmConfig(
            pooled_alpha_scale=pooled_alpha_scale,
            pooled_eye_scale=pooled_eye_scale,
        )
        state_df = classify_actual_states(
            df_main,
            personalized_rt_threshold=rt_threshold,
            config=algorithm_config,
        )
        pred_df = predict_fatigue_states(master_df, config=algorithm_config)
        first_behavioral_second = find_first_fatigue_onset(
            state_df,
            "state_val",
            start_second=0,
        )
        all_time = float(master_df["second"].max())
        plot_end_second = calculate_plot_end_second(
            first_behavioral_second,
            all_time,
        )

        try:
            export_to_excel(pred_df, state_df)
        except Exception as e:
            print(f"❌ 導出 Excel 失敗: {str(e)}")
            
        def draw_driving_state():
            try:
                plt.close('all')
                phase_one_blocked = bool(
                    state_df.attrs.get("phase_one_blocked", False)
                )
                if phase_one_blocked:
                    fig, ax_top = plt.subplots(figsize=(12, 6.5))
                    ax_score = None
                else:
                    fig, (ax_top, ax_score) = plt.subplots(
                        2, 1, figsize=(12, 8.5), sharex=True,
                        gridspec_kw={"height_ratios": [1.35, 1]},
                    )

                # Keep reaction time as a subdued background layer.  Drawing
                # the main axis above the twin axis prevents these stems from
                # covering the fatigue annotations and information legend.
                ax_twin = ax_top.twinx()
                ax_top.set_zorder(ax_twin.get_zorder() + 1)
                ax_top.patch.set_visible(False)
                ax_twin.axhline(
                    rt_threshold,
                    color='#2F5597',
                    linestyle='--',
                    linewidth=1.6,
                    label=(
                        'Personalized RT Threshold '
                        f'({rt_threshold:g} s)'
                    ),
                    zorder=2,
                )

                reaction_events_df = state_df.dropna(subset=['react_time'])
                regular_reactions = reaction_events_df[
                    reaction_events_df['react_time'].le(
                        REACTION_TIME_CAP_SECONDS
                    )
                ]
                overflow_reactions = reaction_events_df[
                    reaction_events_df['react_time'].gt(
                        REACTION_TIME_CAP_SECONDS
                    )
                ]

                if not regular_reactions.empty:
                    markerline, stemlines, _ = ax_twin.stem(
                        regular_reactions['second'],
                        regular_reactions['react_time'],
                        markerfmt='ko',
                        linefmt='k-',
                        basefmt=" ",
                        label='Reaction Time',
                    )
                    plt.setp(
                        markerline,
                        markersize=3.5,
                        alpha=0.45,
                        zorder=1,
                    )
                    plt.setp(
                        stemlines,
                        linewidth=0.5,
                        alpha=0.2,
                        zorder=1,
                    )

                if not overflow_reactions.empty:
                    overflow_y = [
                        REACTION_TIME_OVERFLOW_Y
                    ] * len(overflow_reactions)
                    overflow_marker, overflow_stems, _ = ax_twin.stem(
                        overflow_reactions['second'],
                        overflow_y,
                        markerfmt='k^',
                        linefmt='k-',
                        basefmt=" ",
                        label='Reaction Time > 10 s',
                    )
                    plt.setp(
                        overflow_marker,
                        markersize=5,
                        alpha=0.7,
                        zorder=1,
                    )
                    plt.setp(
                        overflow_stems,
                        linewidth=0.6,
                        alpha=0.25,
                        zorder=1,
                    )

                # Keep the physiological curves in the background so the
                # fatigue timing annotations remain visually dominant.
                ax_top.plot(
                    master_df['second'],
                    master_df['alpha_sum'],
                    color='#7030A0',
                    label='Alpha Accumulation',
                    alpha=0.35,
                    linewidth=1,
                    zorder=1,
                )
                ax_top.plot(
                    master_df['second'],
                    master_df['eye_sum'],
                    color='#ED7D31',
                    label='Eye Accumulation',
                    alpha=0.35,
                    linewidth=1,
                    zorder=1,
                )

                physiological_second, behavioral_second = annotate_fatigue_timing(
                    ax_top,
                    state_df,
                    pred_df,
                    show_physiological=not phase_one_blocked,
                )

                if ax_score is not None:
                    ax_score.plot(
                        pred_df['second'], pred_df['z_alpha'],
                        color='#7030A0', linewidth=1, label='Z-Alpha',
                    )
                    ax_score.plot(
                        pred_df['second'], pred_df['z_eye'],
                        color='#ED7D31', linewidth=1, label='Z-Eye',
                    )
                    ax_score.axhline(
                        algorithm_config.score_threshold,
                        color='#C00000', linestyle='--', linewidth=1.4,
                        label=f'h = {algorithm_config.score_threshold:g}',
                    )
                    ax_score.axvline(
                        PHASE_2_START_SECOND,
                        color='#4472C4', linestyle='--', linewidth=1.2,
                        label='Phase 2 Start',
                    )
                    if physiological_second is not None:
                        ax_score.axvline(
                            physiological_second,
                            color='#008C95', linestyle='-.', linewidth=1.8,
                            label='Confirmed Physiological Warning',
                        )
                    if behavioral_second is not None:
                        ax_score.axvline(
                            behavioral_second,
                            color='#C00000', linestyle='-.', linewidth=1.8,
                            label='Behavioral Fatigue',
                        )
                    ax_score.set_ylabel("Robust Z")
                    ax_score.set_xlabel("Test Time (Seconds)")
                    ax_score.grid(True, linestyle='--', alpha=0.3)
                    ax_score.set_title(
                        "Z-Alpha and Z-Eye both >= "
                        f"{algorithm_config.score_threshold:g} for "
                        f"{algorithm_config.confirmation_seconds} seconds"
                    )
                    ax_score.legend(
                        loc='upper right', fontsize=8, framealpha=1,
                        facecolor='white', edgecolor='#808080',
                    )
                else:
                    ax_top.set_xlabel("Test Time (Seconds)")

                ax_top.set_ylabel("Total Accumulation")
                ax_twin.set_ylabel("Reaction Time (Seconds)")
                ax_twin.set_ylim(0, REACTION_TIME_AXIS_MAX)
                ax_twin.set_yticks(
                    [0, 2, 4, 6, 8, 10, REACTION_TIME_OVERFLOW_Y]
                )
                ax_twin.set_yticklabels(
                    ["0", "2", "4", "6", "8", "10", ">10"]
                )
                configure_reaction_time_coordinate_display(
                    ax_top,
                    ax_twin,
                    reaction_events_df,
                )
                ax_top.set_xlim(0, plot_end_second)
                ax_top.set_title(
                    "The Evolution of Driving States\n"
                    f"Personalized RT threshold = {rt_threshold:g} s",
                    fontsize=14,
                    pad=15,
                )
                handles_top, labels_top = ax_top.get_legend_handles_labels()
                handles_twin, labels_twin = ax_twin.get_legend_handles_labels()
                ax_top.legend(
                    handles_top + handles_twin,
                    labels_top + labels_twin,
                    loc='upper right',
                    fontsize=9,
                    framealpha=1,
                    facecolor='white',
                    edgecolor='#808080',
                )
                plt.tight_layout()
                figure_output_path.parent.mkdir(parents=True, exist_ok=True)
                fig.savefig(figure_output_path, dpi=300, bbox_inches='tight')
                plt.show()
                self.safe_update_status(f"駕駛狀態圖已儲存：{figure_output_path.name}")

            except Exception as e:
                messagebox.showerror("繪圖錯誤", f"駕駛狀態圖生成失敗：\n{str(e)}")
                self.safe_update_status("駕駛狀態圖生成失敗")

        self.root.after(0, draw_driving_state)

    def load_dat_counts(self, path):
        try:
            df = pd.read_csv(path, header=None, sep=None, engine='python')
            v = df.values.flatten()[1:]
            s = pd.to_numeric(pd.Series(v), errors='coerce').dropna().astype(int)
            return s.value_counts().sort_index()
        except: return pd.Series()

    def add_number_row(self, master, text, var, suffix):
        row = tk.Frame(master)
        row.pack(fill='x', pady=2)
        tk.Label(row, text=text, width=20).pack(side='left')
        tk.Entry(row, textvariable=var).pack(
            side='left', fill='x', expand=True, padx=5
        )
        tk.Label(row, text=suffix).pack(side='right')

    def add_file_row(self, master, text, var, is_xlsx):
        row = tk.Frame(master); row.pack(fill='x', pady=2)
        tk.Label(row, text=text, width=15).pack(side='left')
        tk.Entry(row, textvariable=var, state='readonly').pack(side='left', fill='x', expand=True, padx=5)
        tk.Button(row, text="瀏覽", command=lambda: self.select_file(var, is_xlsx)).pack(side='right')

    def select_file(self, var, is_xlsx):
        ext = [("Excel", "*.xlsx *.xls")] if is_xlsx else [("Data", "*.dat")]
        f = filedialog.askopenfilename(filetypes=ext)
        if f: var.set(f)

    def on_closing(self):
        plt.close('all')
        self.root.quit() # 使用 quit() 確保主迴圈安全結束
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = EEGBrowserGUI(root)
    root.mainloop()
