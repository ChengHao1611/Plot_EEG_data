import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# 設定 Matplotlib 字型
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False

class EEGBrowserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("EEG 綜合數據分析瀏覽器")
        self.root.geometry("1100x850")

        # --- 狀態變數 ---
        # 1. 預設改為 multi (三檔關聯)
        self.mode = tk.StringVar(value="multi") 
        self.path_xlsx = tk.StringVar()
        self.path_alpha = tk.StringVar()
        self.path_eye = tk.StringVar()
        
        self.all_groups_data = {}
        self.current_group_id = None
        self.total_groups = 0

        self.create_layout()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_layout(self):
        top_frame = tk.LabelFrame(self.root, text="數據來源設定", padx=15, pady=10)
        top_frame.pack(side='top', fill='x', padx=20, pady=10)

        mode_frame = tk.Frame(top_frame)
        mode_frame.pack(fill='x', pady=5)
        tk.Label(mode_frame, text="分析模式:", font=('Arial', 10, 'bold')).pack(side='left')
        tk.Radiobutton(mode_frame, text="三檔關聯 (Alpha/眼動累加)", variable=self.mode, 
                       value="multi", command=self.toggle_mode).pack(side='left', padx=10)
        tk.Radiobutton(mode_frame, text="單一 Excel (Ratio/眨眼)", variable=self.mode, 
                       value="single", command=self.toggle_mode).pack(side='left', padx=10)

        self.row_xlsx = self.add_file_row(top_frame, "事件 Excel (.xlsx):", self.path_xlsx, is_excel=True)
        self.row_alpha = self.add_file_row(top_frame, "Alpha 數據 (.dat):", self.path_alpha, is_excel=False)
        self.row_eye = self.add_file_row(top_frame, "眼動 數據 (.dat):", self.path_eye, is_excel=False)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(side='top', fill='x', padx=20)
        
        self.run_btn = tk.Button(btn_frame, text="開始解析並整合數據", command=self.process_data, 
                                 bg='#4CAF50', fg='white', font=('Arial', 10, 'bold'), height=2)
        self.run_btn.pack(fill='x', pady=5)

        # 圖表顯示區
        self.chart_frame = tk.Frame(self.root, bg='white', bd=2, relief='sunken')
        self.chart_frame.pack(side='top', fill='both', expand=True, padx=20, pady=5)
        
        self.fig, self.ax1 = plt.subplots(figsize=(8, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_frame)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)

        # 互動控制區
        self.control_frame = tk.Frame(self.root, pady=10, bd=1, relief='groove')
        
        tk.Label(self.control_frame, text="瀏覽組別:", font=('Arial', 11, 'bold')).pack(side='left', padx=15)
        self.btn_prev = tk.Button(self.control_frame, text="<< 上一組", command=self.prev_group)
        self.btn_prev.pack(side='left', padx=5)

        tk.Label(self.control_frame, text="Group_").pack(side='left')
        self.group_index_entry = tk.Entry(self.control_frame, width=5, justify='center', font=('Arial', 11))
        self.group_index_entry.pack(side='left', padx=2)
        self.group_index_entry.bind('<Return>', lambda e: self.show_requested_group())
        
        self.total_label = tk.Label(self.control_frame, text="/ 00", font=('Arial', 11))
        self.total_label.pack(side='left', padx=5)

        tk.Button(self.control_frame, text="跳轉", command=self.show_requested_group).pack(side='left', padx=5)
        self.btn_next = tk.Button(self.control_frame, text="下一組 >>", command=self.next_group)
        self.btn_next.pack(side='left', padx=5)

        self.time_info_label = tk.Label(self.control_frame, text=" | 反應時間: -- s", 
                                        font=('Arial', 11, 'bold'), fg='#2c3e50')
        self.time_info_label.pack(side='left', padx=15)
        
        # initialize mode
        self.toggle_mode()

    def on_closing(self):
        """完全關閉程式，釋放所有資源"""
        plt.close('all')
        self.root.destroy()
        import sys
        sys.exit()

    # 透過 is_excel 參數來正確過濾檔案類型
    def add_file_row(self, master, label_text, var, is_excel=True):
        row = tk.Frame(master)
        row.pack(fill='x', pady=2)
        tk.Label(row, text=label_text, width=18, anchor='w').pack(side='left')
        tk.Entry(row, textvariable=var, state='readonly').pack(side='left', fill='x', expand=True, padx=5)
        tk.Button(row, text="瀏覽", command=lambda: self.select_file(var, is_excel)).pack(side='right')
        return row

    def select_file(self, var, is_excel):
        if is_excel:
            ext = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        else:
            ext = [("Data files", "*.dat"), ("All files", "*.*")]
        
        file = filedialog.askopenfilename(filetypes=ext)
        if file: var.set(file)

    def toggle_mode(self):
        if self.mode.get() == "single":
            self.row_alpha.pack_forget()
            self.row_eye.pack_forget()
        else:
            self.row_alpha.pack(fill='x', pady=2)
            self.row_eye.pack(fill='x', pady=2)
        
        if hasattr(self, 'control_frame'):
            self.control_frame.pack_forget()
        self.clear_plot()

    def load_dat_counts(self, path):
        """讀取 .dat，排除檔案中的第一個數字，統計每秒次數"""
        try:
            df = pd.read_csv(path, header=None, sep=None, engine='python')
            all_data = df.values.flatten()
            valid_data = all_data[1:]
            
            # 4. 轉換為整數秒數並統計頻率
            s_data = pd.Series(valid_data)
            s_data = pd.to_numeric(s_data, errors='coerce').dropna().astype(int)
            
            return s_data.value_counts().sort_index()
        except Exception as e:
            print(f"讀取 .dat 出錯: {e}")
            return pd.Series()

    def process_data(self):
        if not self.path_xlsx.get():
            messagebox.showwarning("警告", "請先選擇主要 Excel 檔案")
            return

        try:
            self.run_btn.config(state='disabled', text="正在整合數據...")
            self.root.update()

            # 1. 讀取核心 Excel
            df_main = pd.read_excel(self.path_xlsx.get())
            df_main['second'] = pd.to_numeric(df_main['second']).round(2)
            
            # 2. 根據模式處理數據
            if self.mode.get() == "multi":
                if not self.path_alpha.get() or not self.path_eye.get():
                    messagebox.showwarning("警告", "三檔模式下請選擇所有檔案")
                    self.run_btn.config(state='normal', text="開始解析並整合數據")
                    return
                
                alpha_counts = self.load_dat_counts(self.path_alpha.get())
                eye_counts = self.load_dat_counts(self.path_eye.get())
                
                max_s = int(max(df_main['second'].max(), 
                                alpha_counts.index.max() if not alpha_counts.empty else 0,
                                eye_counts.index.max() if not eye_counts.empty else 0))
                
                master_df = pd.DataFrame({'second': range(max_s + 1)})
                master_df = master_df.merge(alpha_counts.rename('a_raw'), left_on='second', right_index=True, how='left').fillna(0)
                master_df = master_df.merge(eye_counts.rename('e_raw'), left_on='second', right_index=True, how='left').fillna(0)
                
                # 計算 30 秒滑動累加
                master_df['alpha_sum'] = master_df['a_raw'].rolling(window=30, min_periods=1).sum()
                master_df['eye_sum'] = master_df['e_raw'].rolling(window=30, min_periods=1).sum()
                try:
                    output_name = "alpha_eye count.xlsx"
                    output_path = os.path.join(os.path.dirname(self.path_xlsx.get()), output_name)
                    master_df.to_excel(output_path, index=False)
                    print(f"數據已匯出至: {output_path}")
                except Exception as e:
                    print(f"匯出 Excel 失敗: {e}")
                plot_source = master_df
            else:
                plot_source = df_main

            events = df_main[df_main['react_time'].notna()].sort_values(by='react_time')
            self.all_groups_data = {}
            
            for i, (_, row) in enumerate(events.iterrows(), 1):
                t = row['second']
                mask = (plot_source['second'] > (t - 30)) & (plot_source['second'] <= t)
                group_df = plot_source[mask].copy()
                if not group_df.empty:
                    group_df.attrs['react_val'] = row['react_time']
                    self.all_groups_data[i] = group_df

            self.total_groups = len(self.all_groups_data)
            self.total_label.config(text=f"/ {self.total_groups:02d}")
            self.control_frame.pack(side='bottom', fill='x', padx=20, pady=10)
            self.display_group(1)

        except Exception as e:
            messagebox.showerror("整合錯誤", f"錯誤訊息: {str(e)}")
        finally:
            self.run_btn.config(state='normal', text="開始解析並整合數據")

    def clear_plot(self):
        self.ax1.clear()
        if hasattr(self, 'ax2'):
            self.ax2.remove()
            del self.ax2
        self.canvas.draw()

    def display_group(self, idx):
        if not self.all_groups_data or idx not in self.all_groups_data: return
        
        self.current_group_id = idx
        data = self.all_groups_data[idx]
        target_t = data['second'].max()
        react_val = data.attrs.get('react_val', 0)

        self.group_index_entry.delete(0, tk.END)
        self.group_index_entry.insert(0, f"{idx:02d}")
        self.time_info_label.config(text=f" | 反應時間: {react_val:.2f} s")

        self.clear_plot()
        self.ax1.grid(True, linestyle=':', alpha=0.5)

        if self.mode.get() == "single":
            if 'alpha_beta' in data.columns:
                self.ax1.plot(data['second'], data['alpha_beta'], label='α/β Ratio', color='blue')
                self.ax1.set_ylabel('Ratio (α/β)', color='blue', fontweight='bold')
            if 'eyeblinking_count' in data.columns:
                self.ax2 = self.ax1.twinx()
                self.ax2.plot(data['second'], data['eyeblinking_count'], label='眨眼次數', color='red')
                self.ax2.set_ylabel('眨眼次數', color='red', fontweight='bold')
        else:
            self.ax1.plot(data['second'], data['alpha_sum'], label='Alpha 30s 累加', color='purple', linewidth=2)
            self.ax1.set_ylabel('Alpha 總秒數', color='purple', fontweight='bold')
            
            self.ax2 = self.ax1.twinx()
            self.ax2.plot(data['second'], data['eye_sum'], label='眼動 30s 累加', color='darkorange', linewidth=2)
            self.ax2.set_ylabel('眼動 總秒數', color='darkorange', fontweight='bold')

        self.ax1.axvline(x=target_t, color='black', linestyle='--', alpha=0.7, label='反應點')
        self.ax1.set_title(f"EEG 分析 - Group {idx:02d} (於 {target_t:.2f}s)")
        self.ax1.set_xlabel("時間 (s)")
        
        h1, l1 = self.ax1.get_legend_handles_labels()
        if hasattr(self, 'ax2'):
            h2, l2 = self.ax2.get_legend_handles_labels()
            self.ax1.legend(h1+h2, l1+l2, loc='upper left')
        else:
            self.ax1.legend(loc='upper left')

        self.canvas.draw()

    def show_requested_group(self):
        try:
            idx = int(self.group_index_entry.get())
            if 1 <= idx <= self.total_groups: self.display_group(idx)
        except: pass

    def next_group(self):
        if self.current_group_id < self.total_groups: self.display_group(self.current_group_id + 1)

    def prev_group(self):
        if self.current_group_id > 1: self.display_group(self.current_group_id - 1)

if __name__ == "__main__":
    root = tk.Tk()
    app = EEGBrowserGUI(root)
    root.mainloop()