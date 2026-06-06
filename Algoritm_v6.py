import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import threading

# 設定字體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False

class EEGBrowserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("EEG 數據動態分析系統 v2.0")
        self.root.geometry("1100x950")

        self.path_xlsx = tk.StringVar()
        self.path_alpha = tk.StringVar()
        self.path_eye = tk.StringVar()
        self.analysis_mode = tk.IntVar(value=1)  # 1: 原本模式, 2: 趨勢對照模式
        
        self.all_groups_data = {}
        self.current_group_id = None
        self.total_groups = 0

        self.create_layout()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_layout(self):
        # 1. 數據來源設定
        top_frame = tk.LabelFrame(self.root, text="數據來源與模式設定", padx=15, pady=10)
        top_frame.pack(side='top', fill='x', padx=20, pady=10)

        self.add_file_row(top_frame, "事件 Excel (.xlsx):", self.path_xlsx, True)
        self.add_file_row(top_frame, "Alpha 數據 (.dat):", self.path_alpha, False)
        self.add_file_row(top_frame, "眼動 數據 (.dat):", self.path_eye, False)

        # 模式選擇
        mode_frame = tk.Frame(top_frame)
        mode_frame.pack(fill='x', pady=5)
        tk.Label(mode_frame, text="分析模式：", font=('Arial', 10, 'bold')).pack(side='left')
        tk.Radiobutton(mode_frame, text="模式一：事件切片(逐筆)", variable=self.analysis_mode, value=1).pack(side='left', padx=10)
        tk.Radiobutton(mode_frame, text="模式二：整體趨勢與反應時間對照", variable=self.analysis_mode, value=2).pack(side='left', padx=10)
        tk.Radiobutton(mode_frame, text="模式三：駕駛狀態判定與統計", variable=self.analysis_mode, value=3).pack(side='left', padx=10)

        # 2. 進度條
        self.progress_frame = tk.Frame(self.root, padx=20)
        self.progress_frame.pack(fill='x')
        self.progress_label = tk.Label(self.progress_frame, text="準備就緒", font=('Arial', 9))
        self.progress_label.pack(side='top', anchor='w')
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill='x', pady=5)

        self.run_btn = tk.Button(self.root, text="執行分析並匯出報表", command=self.start_thread, 
                                 bg='#2196F3', fg='white', font=('Arial', 10, 'bold'), height=2)
        self.run_btn.pack(fill='x', padx=20, pady=5)

        # 3. 圖表區
        self.chart_frame = tk.Frame(self.root, bg='white', bd=2, relief='sunken')
        self.chart_frame.pack(side='top', fill='both', expand=True, padx=20, pady=5)
        self.fig, self.ax1 = plt.subplots(figsize=(8, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_frame)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)

        # 4. 控制區
        self.control_frame = tk.Frame(self.root, pady=10)
        self.control_frame.pack(fill='x')
        self.prev_btn = tk.Button(self.control_frame, text="<< 上一組", command=self.prev_group)
        self.prev_btn.pack(side='left', padx=5)
        self.group_index_entry = tk.Entry(self.control_frame, width=8, justify='center')
        self.group_index_entry.pack(side='left', padx=2)
        self.total_label = tk.Label(self.control_frame, text="/ 00")
        self.total_label.pack(side='left', padx=5)
        self.next_btn = tk.Button(self.control_frame, text="下一組 >>", command=self.next_group)
        self.next_btn.pack(side='left', padx=5)
        self.time_info_label = tk.Label(self.control_frame, text=" | 狀態: --", font=('Arial', 10, 'bold'))
        self.time_info_label.pack(side='left', padx=15)

    def start_thread(self):
        # 關鍵：在主執行緒先把資料取出來，避免背景執行緒存取 UI 變數
        plt.close('all')
        xlsx_val = self.path_xlsx.get()
        alpha_val = self.path_alpha.get()
        eye_val = self.path_eye.get()
        mode_val = self.analysis_mode.get()

        if not all([xlsx_val, alpha_val, eye_val]):
            messagebox.showwarning("警告", "請先選擇所有必要的數據文件。")
            return

        self.run_btn.config(state='disabled')
        # 把參數傳進去，不要讓 Thread 直接讀 self.path_xlsx
        threading.Thread(target=self.process_data, args=(xlsx_val, alpha_val, eye_val, mode_val), daemon=True).start()

    def process_data(self, xlsx_p, alpha_p, eye_p, mode):
        try:
            self.safe_update_status("數據讀取中...")
            
            # 讀取數據 (使用傳進來的字串路徑)
            df_main = pd.read_excel(xlsx_p)
            df_main['second'] = pd.to_numeric(df_main['second'], errors='coerce')
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

            if mode == 1:
                self.process_mode_one(df_main, master_df)
            elif mode == 2:
                self.process_mode_two(df_main, master_df)
            elif mode == 3:
                self.process_mode_three(df_main, master_df)

        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: messagebox.showerror("錯誤", f"處理中斷: {err_msg}"))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state='normal'))

    def safe_update_status(self, text):
        # update UI from main thread
        self.root.after(0, lambda: self.progress_label.config(text=text))

    def process_mode_one(self, df_main, master_df):
        events = df_main[df_main['react_time'].notna()].copy()
        temp_groups = {}
        
        for idx, row in events.iterrows():
            t_event = row['second']
            react_s = row['react_time']
            t_start, t_end = t_event - (30 + react_s), t_event + 5
            group_df = master_df[(master_df['second'] >= t_start) & (master_df['second'] <= t_end)].copy()
            if not group_df.empty:
                group_df.attrs = {'react_val': react_s, 'event_time': t_event}
                temp_groups[idx + 1] = group_df

        self.root.after(0, lambda: self.finalize_mode_one(temp_groups))

    def finalize_mode_one(self, data):
        self.all_groups_data = data
        self.total_groups = len(data)
        self.total_label.config(text=f"/ {self.total_groups:02d}")
        if self.total_groups > 0:
            self.display_group(list(data.keys())[0])
        self.progress_label.config(text="模式一處理完成")
        messagebox.showinfo("完成", "模式一數據處理完成！")

    def process_mode_two(self, df_main, master_df):
        self.safe_update_status("正在準備趨勢圖...")
        plot_df = master_df.merge(df_main[['second', 'react_time']], on='second', how='left')
        # 丟回主執行緒開啟新視窗
        self.root.after(0, lambda: self.popup_trend_chart(plot_df))

    def popup_trend_chart(self, plot_df):
        try:
            # 關鍵：先關閉所有舊視窗，避免資源衝突卡死
            plt.close('all') 
            
            fig, ax1 = plt.subplots(figsize=(12, 7))
            ax2 = ax1.twinx()

            lns1 = ax1.plot(plot_df['second'], plot_df['alpha_sum'], color='#7030A0', label='Alpha (左軸)')
            lns2 = ax1.plot(plot_df['second'], plot_df['eye_sum'], color='#FFC000', label='Eye (左軸)')
            ax1.set_xlabel('時間 (秒)')
            ax1.set_ylabel('累積量')

            markerline, stemlines, baseline = ax2.stem(
                plot_df['second'], plot_df['react_time'], 
                linefmt='C0--', markerfmt='C0o', label='反應時間 (右軸)', basefmt=" "
            )
            ax2.set_ylabel('反應時間 (秒)', color='C0')

            plt.title("EEG 整體趨勢對照圖 (模式二)")
            plt.tight_layout()
            self.progress_label.config(text="趨勢圖已開啟")
            plt.show() 
        except Exception as e:
            messagebox.showerror("繪圖錯誤", str(e))

    # mode 3
    def process_mode_three(self, df_main, master_df):
        """模式三：駕駛狀態隨時間演變圖 (生理峰值追蹤預測 vs 實際反應時間)"""
        self.safe_update_status("正在生成狀態演變圖...")
        
        # 設定反應時間判定門檻
        ALERT_THRESHOLD = 3.0
        FATIGUE_THRESHOLD = 7.0
        
        # 黑線
        def get_state_value(t):
            if pd.isna(t): return None
            if t < ALERT_THRESHOLD: return 0  # Alert
            if t < FATIGUE_THRESHOLD: return 1  # Fatigue
            return 1  # Drowsy

        # 紅線 (改)
        def calculate_all_pred_states(df):
            pred_results = []
            eye_alarm_list = []
            alpha_alarm_list = []

            max_e = 0
            pre_e = 0
            pre_a = 0
            pre_state = 0
            
            triggered_eye_val = None 
            
            monitor_enabled = True # switch for fatigue monitoring, starts enabled
            INIT_BUFFER_SECONDS = 30
            EYE_ALERT_MIN_THRESHOLD = 7
            BASE_RECOVERY_THRESHOLD = 10 # 改名以區分，這是保底的絕對恢復門檻
            
            for i, row in df.iterrows():
                current_second = row['second']
                a = row['alpha_sum']
                e = row['eye_sum']

                # warm-up period
                if current_second < INIT_BUFFER_SECONDS:
                    pred_results.append(0)  # Alert during initial buffer period
                    eye_alarm_list.append(0)
                    alpha_alarm_list.append(0)
                    pre_e = e
                    pre_a = a
                    pre_state = pre_state
                    max_e = e  # 初始化 max_e 為第一秒的 eye_sum
                    continue
                
                # check for recovery condition when monitor is disabled
                if not monitor_enabled:
                    dynamic_recovery_limit = BASE_RECOVERY_THRESHOLD
                    if triggered_eye_val is not None:
                        if triggered_eye_val > 7 and e <= 7:
                            triggered_eye_val = 7
                        dynamic_recovery_limit = min(28, max(BASE_RECOVERY_THRESHOLD, triggered_eye_val + 3))
                    
                    # 只要達到動態門檻且正在上升，就允許恢復清醒
                    if e >= dynamic_recovery_limit and e > pre_e: 
                        monitor_enabled = True  
                        triggered_eye_val = None # 成功恢復後，重設紀錄值
                
                # update max_e only when monitor is enabled, to track the new peak after recovery
                if monitor_enabled and (e > pre_e):
                    if e >= BASE_RECOVERY_THRESHOLD:
                        max_e = e

                current_eye_alarm = 0
                current_alpha_alarm = 0

                # alert only when monitor is enabled
                if monitor_enabled:
                    eye_alarm = (e <= pre_e) and (e > 0) if max_e > 0 else False
                    alpha_alarm = (a >= 3)
                    current_eye_alarm = 1 if eye_alarm else 0                    
                    current_alpha_alarm = 1 if alpha_alarm else 0

                    if e <= 0.1 and a <= 0.1:
                        state = 1  # 2
                    elif (eye_alarm and alpha_alarm) or (e <= EYE_ALERT_MIN_THRESHOLD and e != max_e):
                        state = 1  # fatigue
                        monitor_enabled = False  # 【關鍵】觸發後立刻關閉開關，進入鎖定
                        
                        # 【新增】在這裡紀錄觸發當下的眼動值
                        triggered_eye_val = e 
                    else:
                        state = 0  # 清醒
                else:
                    # monitor disabled, maintain fatigue state unless clear recovery condition is met
                    state = 1
                    if e <= 0.1 and a <= 0.1: #until clear sleep condition is met
                        state = 1 # 2

                pred_results.append(state)
                eye_alarm_list.append(current_eye_alarm)
                alpha_alarm_list.append(current_alpha_alarm)

                pre_e = e
                pre_a = a
                pre_state = state
                
            return pred_results, eye_alarm_list, alpha_alarm_list
        
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

        # tidy up state dataframe for plotting
        state_df = df_main[['second', 'react_time']].copy()
        state_df['state_val'] = state_df['react_time'].apply(get_state_value)
        state_df = state_df.dropna(subset=['state_val'])

        # calculate predicted states based on the master_df
        pred_df = master_df.copy()
        pred_df['pred_state'], pred_df['eye_alarm'], pred_df['alpha_alarm'] = calculate_all_pred_states(pred_df)

        try:
            export_to_excel(pred_df, df_main)
        except Exception as e:
            print(f"❌ 導出 Excel 失敗: {str(e)}")
            
        def draw_mode_three():
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
                
                react_times = state_df['react_time'].clip(upper=10)
                markerline, stemlines, baseline = ax_twin.stem(
                    state_df['second'], react_times,
                    markerfmt='ko', linefmt='k-', basefmt=" ", label='反應時間'
                )
                plt.setp(markerline, markersize=4, zorder=5)
                plt.setp(stemlines, linewidth=0.5, alpha=0.5)

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
                plt.show()
                self.safe_update_status("模式三離散訊號圖已成功生成")

            except Exception as e:
                messagebox.showerror("繪圖錯誤", f"模式三生成失敗：\n{str(e)}")
                self.safe_update_status("模式三失敗")

        self.root.after(0, draw_mode_three)


    def display_group(self, orig_id):
        # 這裡也要防錯，確保存取屬性時資料存在
        if orig_id not in self.all_groups_data: return
        self.current_group_id = orig_id
        data = self.all_groups_data[orig_id]
        
        self.ax1.clear()
        if hasattr(self, 'ax2'): 
            try: self.ax2.remove()
            except: pass
            del self.ax2
        
        self.ax1.plot(data['second'], data['alpha_sum'], color='purple', label='Alpha')
        self.ax1.plot(data['second'], data['eye_sum'], color='orange', label='Eye')
        self.ax1.axvline(x=data.attrs['event_time'], color='red', linestyle='--', label='事件點')
        
        self.ax1.set_title(f"模式一：原始第 {orig_id} 組 (反應: {data.attrs['react_val']}s)")
        self.ax1.legend(loc='upper left')
        
        self.group_index_entry.delete(0, tk.END)
        self.group_index_entry.insert(0, str(orig_id))
        self.time_info_label.config(text=f" | 反應時間: {data.attrs['react_val']} s")
        self.canvas.draw()

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

    def next_group(self):
        if not self.all_groups_data: return
        keys = list(self.all_groups_data.keys())
        try:
            idx = keys.index(self.current_group_id)
            if idx < len(keys) - 1: self.display_group(keys[idx + 1])
        except: self.display_group(keys[0])

    def prev_group(self):
        if not self.all_groups_data: return
        keys = list(self.all_groups_data.keys())
        try:
            idx = keys.index(self.current_group_id)
            if idx > 0: self.display_group(keys[idx - 1])
        except: self.display_group(keys[0])

    def on_closing(self):
        plt.close('all')
        self.root.quit() # 使用 quit() 確保主迴圈安全結束
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = EEGBrowserGUI(root)
    root.mainloop()