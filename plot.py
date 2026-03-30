import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

class EEGPlotterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("EEG 數據可視化工具")
        self.root.geometry("400x250")

        # 設定變數
        self.file_path = tk.StringVar()
        self.mode = tk.IntVar(value=2)  # 預設模式 2

        # 建立介面佈局
        self.create_widgets()

    def create_widgets(self):
        # 檔案選擇區
        tk.Label(self.root, text="步驟 1: 選擇 Excel 檔案", font=('Arial', 10, 'bold')).pack(pady=10)
        
        file_frame = tk.Frame(self.root)
        file_frame.pack(fill='x', padx=20)
        
        tk.Entry(file_frame, textvariable=self.file_path, state='readonly').pack(side='left', fill='x', expand=True)
        tk.Button(file_frame, text="瀏覽...", command=self.browse_file).pack(side='right', padx=5)

        # 模式選擇區
        tk.Label(self.root, text="步驟 2: 選擇顯示模式", font=('Arial', 10, 'bold')).pack(pady=10)
        
        mode_frame = tk.Frame(self.root)
        mode_frame.pack()
        
        tk.Radiobutton(mode_frame, text="使用α/Total 比例作圖", variable=self.mode, value=1).pack(side='left', padx=10)
        tk.Radiobutton(mode_frame, text="使用α/β & α/θ 比例作圖", variable=self.mode, value=2).pack(side='left', padx=10)

        # 執行按鈕
        tk.Button(self.root, text="生成圖表", command=self.run_plot, bg='#4CAF50', fg='white', font=('Arial', 12, 'bold')).pack(pady=20)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="選擇 EEG Excel 檔案",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filename:
            self.file_path.set(filename)

    def plot_eeg_data(self, xlsx_file, mode):
        try:
            df = pd.read_excel(xlsx_file)
            eeg_file = os.path.basename(xlsx_file).split('_raw_')[0]
            
            # 建立圖表
            fig, ax1 = plt.subplots(figsize=(12, 6))
            events = df[df['react_time'] > 0]

            event_patch = None
            shaded_label_added = False
            for _, row in events.iterrows():
                start_t = row['second']
                duration = row['react_time']
                end_t = start_t + duration
                label = 'Event' if not shaded_label_added else ""
                p = ax1.axvspan(start_t, end_t, color='yellow', alpha=0.3, label=label, zorder=1)
                if not shaded_label_added:
                    event_patch = p
                    shaded_label_added = True

            # 繪製主軸數據
            if mode == 2:
                line1, = ax1.plot(df['second'], df['alpha_beta'], label='α/β Ratio', color='blue', alpha=0.7, zorder=3)
                line2, = ax1.plot(df['second'], df['alpha_theta'], label='α/θ Ratio', color='green', alpha=0.7, zorder=3)
                lines_main = [line1, line2]
            else:
                line1, = ax1.plot(df['second'], df['alpha_total'], label='α/total Ratio', color='blue', alpha=0.7, zorder=3)
                lines_main = [line1]

            ax1.set_xlabel('Time(s)')
            ax1.set_ylabel('Ratio', color='black')
            ax1.grid(True, which='both', linestyle='--', linewidth=0.5, zorder=0)

            # 繪製次軸 (眨眼)
            ax2 = ax1.twinx()
            line3, = ax2.plot(df['second'], df['eyeblinking_count'], label='Eyeblink Count', color='red', linewidth=2, linestyle='-', zorder=4)
            ax2.set_ylabel('Eyeblink Count (per 30s)', color='red')
            ax2.tick_params(axis='y', labelcolor='red')

            # 合併圖例
            elements = []
            if event_patch: elements.append(event_patch)
            elements.extend(lines_main + [line3])
            
            labels = [obj.get_label() for obj in elements]
            ax1.legend(elements, labels, loc='upper right')
            
            plt.title(f"EEG Analysis: {eeg_file}")
            plt.tight_layout()
            plt.show()

        except Exception as e:
            messagebox.showerror("錯誤", f"繪圖過程中發生錯誤：\n{str(e)}")

    def run_plot(self):
        path = self.file_path.get()
        if not path:
            messagebox.showwarning("警告", "請先選擇一個 Excel 檔案！")
            return
        if not os.path.exists(path):
            messagebox.showerror("錯誤", "找不到該檔案。")
            return
        
        self.plot_eeg_data(path, self.mode.get())

if __name__ == "__main__":
    root = tk.Tk()
    app = EEGPlotterGUI(root)
    root.mainloop()
