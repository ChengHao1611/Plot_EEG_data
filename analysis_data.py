from matplotlib.axes import Axes
from matplotlib.figure import Figure
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from pandas import DataFrame
matplotlib.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]  # 設定中文字體
matplotlib.rcParams["axes.unicode_minus"] = False  # 解決負號顯示問題
from plot_data import *

#data_column_name = ["秒數", "事件反應時間", "α波時間", "導回車道用時", "睡著"]

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


def input_picture_format(df: DataFrame):
    """
    輸入要產生的模式和圖表，並產生圖表
    """
    print("請輸入要產生的模式: 1. 合併圖表 2. 單獨圖表 3. 三張圖疊加")
    mode = int(input("請輸入數字："))
    print("請輸入要產生哪些圖表：1. 事件反應時間 2. α波時間 3. 導回車道用時 4. 睡著, 5. 眼動次數")
    if mode == 1:
        chart1 = int(input("請輸入第一個圖表的數字："))
        if chart1 != 4:
            print("請輸入要產生折線圖還是長條圖：1. 折線圖 2. 長條圖")
            mode1 = int(input("請輸入數字："))
            fig1, ax1 = check_ouput_picture(df, data_column_name[chart1], mode1)
        else:
            fig1, ax1 = plot_sleep_area(df, color="blueviolet")
        
        chart2 = int(input("請輸入第二個圖表的數字："))
        if chart2 != 4:
            print("請輸入要產生折線圖還是長條圖：1. 折線圖 2. 長條圖")
            mode2 = int(input("請輸入數字："))
            fig2, ax2 = check_ouput_picture(df, data_column_name[chart2], mode2)
        else:
            mode2 = None
            fig2, ax2 = plot_sleep_area(df, color="salmon")
        plot_data_combined(fig1, ax1, chart1, mode1, fig2, ax2, chart2, mode2, df)
    elif mode == 2:
        chart3 = int(input("請輸入數字："))
        if chart3 != 4:
            print("請輸入要產生折線圖還是長條圖：1. 折線圖 2. 長條圖")
            mode3 = int(input("請輸入數字："))
            fig3, ax3 = check_ouput_picture(df, data_column_name[chart3], mode3)
        else:
            fig3, ax3 = plot_sleep_area(df, color="blueviolet")

        # 顯示圖表
        fig3.tight_layout()
        fig3.show()  
        #plt.close(fig1)  # 關閉圖表，釋放記憶體
    elif mode == 3:
        print("請輸入三個圖表的數字：")
        chart1 = int(input("第一個圖表："))
        if chart1 != 4:
            print("請輸入要產生折線圖還是長條圖：1. 折線圖 2. 長條圖")
            mode1 = int(input("請輸入數字："))
        else:
            mode1 = None

        chart2 = int(input("第二個圖表："))
        if chart2 != 4:
            print("請輸入要產生折線圖還是長條圖：1. 折線圖 2. 長條圖")
            mode2 = int(input("請輸入數字："))
        else:
            mode2 = None

        chart3 = int(input("第三個圖表："))
        if chart3 != 4:
            print("請輸入要產生折線圖還是長條圖：1. 折線圖 2. 長條圖")
            mode3 = int(input("請輸入數字："))
        else:
            mode3 = None

        plot_data_triple(df, chart1, mode1, chart2, mode2, chart3, mode3)
    else:
        print("輸入錯誤")


if __name__ == "__main__":
    excel_file = input("請輸入 Excel 檔案路徑: ")
    #excel_file = "E:\專題\data\s09_060317n.set\s09_060317n.xlsx" 
    excel_file = excel_file.replace('"', '')
    
    try:
        df = pd.read_excel(excel_file)
        df = process_data(df)
        while(1): input_picture_format(df)
        exit()
        
    except Exception as e:
        print(f"發生錯誤: {e}")