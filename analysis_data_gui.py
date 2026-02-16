from tkinter import Tk, filedialog, messagebox, ttk, StringVar, IntVar, W, E, N, S, LEFT
from matplotlib.axes import Axes
from matplotlib.figure import Figure
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from pandas import DataFrame
from plot_data import *

matplotlib.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]  # 設定中文字體
matplotlib.rcParams["axes.unicode_minus"] = False  # 解決負號顯示問題

def process_data(df: DataFrame) -> DataFrame:
    """
    處理數據
    """
    # 從第二個 row 開始讀取數據（跳過第一個 row）
    data = df.iloc[1:].copy()
    data["秒數"] = pd.to_numeric(data["秒數"], errors="coerce")
    data["事件反應時間"] = pd.to_numeric(data["事件反應時間"], errors="coerce")
    data["α波時間"] = pd.to_numeric(data["α波時間"], errors="coerce")
    data["導回車道用時"] = pd.to_numeric(data["導回車道用時"], errors="coerce")
    data["睡著"] = pd.to_numeric(data["睡著"], errors="coerce")
    data["眼動次數"] = pd.to_numeric(data["眼動次數"], errors="coerce")
    
    # 將 NaN 值替換為 0
    data = data.fillna(0)
    return data

class AnalysisDataGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("數據分析與圖表生成工具")
        self.root.geometry("800x600")
        
        self.df = None
        self.file_path = None
        
        # 圖表選項（排除「睡著」，因為第三個圖表固定為睡著）
        self.chart_options = ["事件反應時間", "α波時間", "導回車道用時", "眼動次數"]
        self.chart_indices = [1, 2, 3, 5]  # 對應 data_column_name 的索引（排除睡著索引4）
        
        # 每個圖表選項的專屬顏色
        self.chart_colors = {
            "事件反應時間": "#4169E1",      # 皇家藍
            "α波時間": "#FF6347",           # 番茄紅
            "導回車道用時": "#9370DB",       # 中紫色
            "眼動次數": "#FF8C00",           # 深橙色
            "睡著": "#32CD32"                # 酸橙綠
        }
        
        self.setup_ui()
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(W, E, N, S))
        
        # 檔案選擇區域
        file_frame = ttk.LabelFrame(main_frame, text="檔案選擇", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(W, E), pady=5)
        
        self.file_label = ttk.Label(file_frame, text="尚未選擇檔案")
        self.file_label.grid(row=0, column=0, sticky=W, padx=5)
        
        ttk.Button(file_frame, text="選擇 Excel 檔案", command=self.select_file).grid(row=0, column=1, padx=5)
        
        # 圖表選擇區域（固定為三張圖疊加，第三個固定為睡著）
        self.chart_frame = ttk.LabelFrame(main_frame, text="圖表選擇（三張圖疊加，第三個固定為睡著）", padding="10")
        self.chart_frame.grid(row=1, column=0, columnspan=2, sticky=(W, E, N, S), pady=5)
        
        # 生成按鈕
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="生成圖表", command=self.generate_chart).pack(side=LEFT, padx=5)
        ttk.Button(button_frame, text="清除選擇", command=self.clear_selection).pack(side=LEFT, padx=5)
        ttk.Button(button_frame, text="退出", command=self.root.quit).pack(side=LEFT, padx=5)
        
        # 配置權重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # 初始化圖表選擇區域（固定為兩個圖表選擇）
        self.setup_chart_selection_ui()
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")]
        )
        
        if file_path:
            try:
                self.file_path = file_path
                self.df = pd.read_excel(file_path)
                self.df = process_data(self.df)
                self.file_label.config(text=f"已選擇: {file_path.split('/')[-1]}")
                messagebox.showinfo("成功", "檔案載入成功！")
            except Exception as e:
                messagebox.showerror("錯誤", f"載入檔案時發生錯誤：{str(e)}")
                self.df = None
                self.file_path = None
                self.file_label.config(text="尚未選擇檔案")
    
    def setup_chart_selection_ui(self):
        """設置圖表選擇界面（只選擇兩個圖表，第三個固定為睡著）"""
        # 第一個圖表
        ttk.Label(self.chart_frame, text="第一個圖表：").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        
        self.chart1_var = StringVar(value=self.chart_options[0])
        chart1_combo = ttk.Combobox(self.chart_frame, textvariable=self.chart1_var, values=self.chart_options, state="readonly", width=20)
        chart1_combo.grid(row=0, column=1, padx=5, pady=5)
        
        self.chart1_style_var = IntVar(value=1)
        self.chart1_style_frame = ttk.Frame(self.chart_frame)
        self.chart1_style_frame.grid(row=0, column=2, padx=5, pady=5)
        ttk.Radiobutton(self.chart1_style_frame, text="折線圖", variable=self.chart1_style_var, value=1).pack(side=LEFT)
        ttk.Radiobutton(self.chart1_style_frame, text="長條圖", variable=self.chart1_style_var, value=2).pack(side=LEFT)
        
        # 第二個圖表
        ttk.Label(self.chart_frame, text="第二個圖表：").grid(row=1, column=0, sticky=W, padx=5, pady=5)
        
        # 第二個圖表預設選擇不同的選項（如果有的話）
        default_chart2 = self.chart_options[1] if len(self.chart_options) > 1 else self.chart_options[0]
        self.chart2_var = StringVar(value=default_chart2)
        chart2_combo = ttk.Combobox(self.chart_frame, textvariable=self.chart2_var, values=self.chart_options, state="readonly", width=20)
        chart2_combo.grid(row=1, column=1, padx=5, pady=5)
        
        self.chart2_style_var = IntVar(value=1)
        self.chart2_style_frame = ttk.Frame(self.chart_frame)
        self.chart2_style_frame.grid(row=1, column=2, padx=5, pady=5)
        ttk.Radiobutton(self.chart2_style_frame, text="折線圖", variable=self.chart2_style_var, value=1).pack(side=LEFT)
        ttk.Radiobutton(self.chart2_style_frame, text="長條圖", variable=self.chart2_style_var, value=2).pack(side=LEFT)
        
        # 第三個圖表（固定為睡著，只顯示標籤）
        ttk.Label(self.chart_frame, text="第三個圖表：").grid(row=2, column=0, sticky=W, padx=5, pady=5)
        ttk.Label(self.chart_frame, text="睡著（固定）", foreground="gray").grid(row=2, column=1, sticky=W, padx=5, pady=5)
    
    def get_chart_index(self, chart_name):
        """將圖表名稱轉換為索引"""
        return self.chart_indices[self.chart_options.index(chart_name)]
    
    def get_chart_color(self, chart_name):
        """獲取圖表的專屬顏色"""
        return self.chart_colors.get(chart_name, "blueviolet")
    
    def generate_chart(self):
        if self.df is None:
            messagebox.showwarning("警告", "請先選擇 Excel 檔案！")
            return
        
        try:
            # 固定為三張圖疊加模式
            chart1_name = self.chart1_var.get()
            chart2_name = self.chart2_var.get()
            
            chart1 = self.get_chart_index(chart1_name)
            chart2 = self.get_chart_index(chart2_name)
            chart3 = 4  # 第三個圖表固定為睡著（索引4）
            
            mode1 = self.chart1_style_var.get()
            mode2 = self.chart2_style_var.get()
            mode3 = None  # 睡著不需要模式
            
            # 獲取每個圖表的專屬顏色
            color1 = self.get_chart_color(chart1_name)
            color2 = self.get_chart_color(chart2_name)
            color3 = self.get_chart_color("睡著")
            
            plot_data_triple(self.df, chart1, mode1, chart2, mode2, chart3, mode3, color1, color2, color3)
        except Exception as e:
            messagebox.showerror("錯誤", f"生成圖表時發生錯誤：{str(e)}")
    
    def clear_selection(self):
        self.df = None
        self.file_path = None
        self.file_label.config(text="尚未選擇檔案")
        messagebox.showinfo("清除", "選擇已清除")


if __name__ == "__main__":
    root = Tk()
    app = AnalysisDataGUI(root)
    root.mainloop()

