from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import matplotlib.pyplot as plt
import threading

try:
    from .observe_algorithm import (
        DEFAULT_CONFIG as OBSERVE_CONFIG,
        classify_actual_states,
        predict_fatigue_states,
    )
except ImportError:
    # 支援直接以 python observe.py 執行。
    from observe_algorithm import (
        DEFAULT_CONFIG as OBSERVE_CONFIG,
        classify_actual_states,
        predict_fatigue_states,
    )

# 設定字體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False


# 事件 Excel 可以使用舊版英文欄位，或由逐秒統計表輸出的中文欄位。
EVENT_COLUMN_ALIASES = {
    "second": ("second", "秒數"),
    "react_time": ("react_time", "事件反應時間"),
}


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
            "請使用 second、react_time，或 秒數、事件反應時間。"
        )

    return events_df.rename(columns=rename_columns)


class EEGBrowserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("EEG 駕駛狀態分析系統")
        self.root.geometry("760x300")

        self.path_xlsx = tk.StringVar()
        self.path_alpha = tk.StringVar()
        self.path_eye = tk.StringVar()
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

        self.run_btn.config(state='disabled')
        # 把參數傳進去，不要讓 Thread 直接讀 self.path_xlsx
        threading.Thread(target=self.process_data, args=(xlsx_val, alpha_val, eye_val), daemon=True).start()

    def process_data(self, xlsx_p, alpha_p, eye_p):
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
            master_df['alpha_sum'] = master_df['a_raw'].rolling(window=10, min_periods=1).sum()
            master_df['eye_sum'] = master_df['e_raw'].rolling(window=30, min_periods=1).sum()

            self.process_driving_state(df_main, master_df, xlsx_p)

        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: messagebox.showerror("錯誤", f"處理中斷: {err_msg}"))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state='normal'))

    def safe_update_status(self, text):
        # update UI from main thread
        self.root.after(0, lambda: self.progress_label.config(text=text))

    def process_driving_state(self, df_main, master_df, event_xlsx_path):
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
        def export_to_excel(pred_df, df_main):
            import openpyxl

            react_mapping = df_main.dropna(subset=['react_time']).set_index('second')['react_time'].to_dict()

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Driving State Analysis Results"
            ws.views.sheetView[0].showGridLines = True
            headers = ["Second", "Eye_Sum", "Alpha_Sum", "React_time", "Eye_Alarm", 
                       "Alpha_Alarm", "Pred_State"]
            ws.append(headers)

            for _, row in pred_df.iterrows():
                sec = row['second']
                react_t = react_mapping.get(sec, None)

                ws.append([
                    int(sec), 
                    row['eye_sum'], 
                    row['alpha_sum'], 
                    react_t if react_t is not None else "", 
                    int(row['eye_alarm']), 
                    int(row['alpha_alarm']),
                    int(row['pred_state'])
                ])
            excel_filename = "駕駛狀態分析結果.xlsx"
            wb.save(excel_filename)
            print(f"📊 Excel 檔案已成功儲存至: {excel_filename}")

        # 疲勞判斷完全由 observe_algorithm.py 負責；此檔案只處理 UI、繪圖與匯出。
        state_df = classify_actual_states(df_main)
        pred_df = predict_fatigue_states(master_df)

        try:
            export_to_excel(pred_df, df_main)
        except Exception as e:
            print(f"❌ 導出 Excel 失敗: {str(e)}")
            
        def draw_driving_state():
            try:
                plt.close('all') 
                fig, (ax_top, ax_bot) = plt.subplots(2, 1, figsize=(12, 8), sharex=True, 
                                                     gridspec_kw={'height_ratios': [2, 1]})

                # The top cumulative indicators and react time plot
                ax_twin = ax_top.twinx()
                lns1 = ax_top.plot(master_df['second'], master_df['alpha_sum'], color='#7030A0', 
                                   label='Alpha 累積', alpha=0.6, linewidth=1)
                lns2 = ax_top.plot(master_df['second'], master_df['eye_sum'], color='#ED7D31', 
                                   label='Eye 累積', alpha=0.6, linewidth=1)
                
                reaction_events_df = state_df.dropna(subset=['react_time'])
                react_times = reaction_events_df['react_time'].clip(upper=10)
                markerline, stemlines, baseline = ax_twin.stem(
                    reaction_events_df['second'], react_times,
                    markerfmt='ko', linefmt='k-', basefmt=" ", label='反應時間'
                )
                plt.setp(markerline, markersize=4, zorder=5)
                plt.setp(stemlines, linewidth=0.5, alpha=0.5)
                ax_twin.axhline(
                    OBSERVE_CONFIG.fatigue_start_reaction_threshold,
                    color='#C00000',
                    linestyle='--',
                    linewidth=1.4,
                    label=(
                        '疲勞RT門檻 '
                        f'({OBSERVE_CONFIG.fatigue_start_reaction_threshold:g}秒)'
                    ),
                    zorder=1,
                )
                ax_top.axvline(
                    300,
                    color='#4472C4',
                    linestyle='--',
                    linewidth=1.5,
                    label='Baseline結束（第300秒）',
                    zorder=4,
                )

                target_df = master_df if 'sleep' in master_df.columns else df_main

                sleep_start = None
                has_sleep_legend = False
                if 'sleep' in target_df.columns:
                    for idx, row in target_df.iterrows():
                        if row['sleep'] == 1:
                            sleep_start = row['second']
                        elif row['sleep'] == 2 and sleep_start is not None:
                            sleep_end = row['second']

                            lbl = 'Sleep Period' if not has_sleep_legend else ""
                            ax_top.axvspan(sleep_start, sleep_end, color='#FADBD8', alpha=0.5, label=lbl, zorder=0)
                            has_sleep_legend = True
                            sleep_start = None
                    if sleep_start is not None:
                        lbl = 'Sleep Period' if not has_sleep_legend else ""
                        ax_top.axvspan(sleep_start, target_df['second'].max(), color='#FADBD8', alpha=0.5, label=lbl, zorder=0)


                ax_top.set_ylabel("Total Accumulation")
                ax_twin.set_ylabel("React Time (Seconds)")
                ax_top.set_title("The Evolution of Driving States", fontsize=14, pad=15)

                handles_top, labels_top = ax_top.get_legend_handles_labels()
                handles_twin, labels_twin = ax_twin.get_legend_handles_labels()
                ax_top.legend(handles_top + handles_twin, labels_top + labels_twin, 
                               loc='upper left', fontsize=9)

                # the bottom state evolution step plot
                ax_bot.step(state_df['second'], state_df['state_val'], where='post', 
                            color='#444444', linewidth=2, markersize=4, label='實際狀態 (React)', zorder=2)

                # Predicted state step plot
                ax_bot.step(pred_df['second'], pred_df['pred_state'], where='post', 
                            color='red', linewidth=1.5, alpha=0.7, 
                            label='Predicted State (e/a Logic)', zorder=3)
                ax_bot.axvline(
                    300,
                    color='#4472C4',
                    linestyle='--',
                    linewidth=1.5,
                    label='_nolegend_',
                    zorder=4,
                )
                
                ax_bot.set_yticks([0, 1])
                ax_bot.set_yticklabels(['Alert (Awake)', 'Fatigue (Tired)'], fontsize=9)
                ax_bot.set_ylim(-0.5, 1.5)
                ax_bot.set_ylabel("State Classification")
                ax_bot.set_xlabel("Test Time (Seconds)")
                ax_bot.grid(axis='y', linestyle='--', alpha=0.5)

                # 背景顏色
                ax_bot.axhspan(-0.5, 0.5, facecolor='#C6EFCE', alpha=0.2)
                ax_bot.axhspan(0.5, 1.5, facecolor='#FFEB9C', alpha=0.2)
                ax_bot.legend(loc='upper left', fontsize=9)

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
