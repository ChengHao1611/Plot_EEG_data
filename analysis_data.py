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
    
    # 將 NaN 值替換為 0
    data = data.fillna(0)
    return data


def check_ouput_picture(df: DataFrame, column_name: str, mode: int) -> tuple[Figure, Axes]:
    if mode == 1:
        return plot_line_data(df, column_name)
    elif mode == 2:
        return plot_bar_data(df, column_name)
    
def plot_data_triple(df: DataFrame,
                     chart1: int, mode1: int,
                     chart2: int, mode2: int,
                     chart3: int, mode3: int):
    """
    三張圖疊加：以最長的 X 軸為基準，第三張不顯示 Y 軸
    """
    # 主圖 (第一張)
    fig, ax1 = check_ouput_picture(df, data_column_name[chart1], mode1)

    # 第二張：建立右側 y 軸
    ax2 = ax1.twinx()
    column_name2 = data_column_name[chart2]
    filtered_df2 = df[(df[column_name2].notna()) & (df[column_name2] != 0)]
    if mode2 == 2:
        ax2.bar(filtered_df2["秒數"], filtered_df2[column_name2],
                width=0.8, alpha=0.6, color="salmon", label=column_name2)
    else:
        ax2.plot(filtered_df2["秒數"], filtered_df2[column_name2],
                 color="salmon", marker="o", markersize=4,
                 linewidth=2, label=column_name2)
    ax2.set_ylabel(column_name2, fontsize=12, color="salmon")
    ax2.tick_params(axis="y", labelcolor="salmon")

    # 第三張：再建立一個共享 X 軸的 overlay
    ax3 = ax1.twinx()
    column_name3 = data_column_name[chart3]
    filtered_df3 = df[(df[column_name3].notna()) & (df[column_name3] != 0)]
    if mode3 == 2:
        ax3.bar(filtered_df3["秒數"], filtered_df3[column_name3],
                width=0.8, alpha=0.4, color="green", label=column_name3)
    elif mode3 == None:
        plot_sleep_area_on_ax(ax3, df, color="green", alpha=0.3)
        column_name3 = "睡著"
    else:
        ax3.plot(filtered_df3["秒數"], filtered_df3[column_name3],
                 color="green", marker="x", markersize=4,
                 linewidth=2, label=column_name3)

    # 隱藏第三張的 Y 軸
    ax3.get_yaxis().set_visible(False)

    # 以最長的 X 軸為基準
    xmin = min(df["秒數"].min(), filtered_df2["秒數"].min(), filtered_df3["秒數"].min())
    xmax = max(df["秒數"].max(), filtered_df2["秒數"].max(), filtered_df3["秒數"].max())
    buffer = (xmax - xmin) * 0.05
    ax1.set_xlim(xmin - buffer, xmax + buffer)
    ax2.set_xlim(xmin - buffer, xmax + buffer)
    ax3.set_xlim(xmin - buffer, xmax + buffer)

    # 合併圖例
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    lines3, labels3 = ax3.get_legend_handles_labels()
    ax1.legend(lines1 + lines2 + lines3, labels1 + labels2 + labels3, loc="upper left")

    fig.tight_layout()
    fig.show()
    return fig, ax1

def input_picture_format(df: DataFrame):
    """
    輸入要產生的模式和圖表，並產生圖表
    """
    print("請輸入要產生的模式: 1. 合併圖表 2. 單獨圖表 3. 三張圖疊加")
    mode = int(input("請輸入數字："))
    print("請輸入要產生哪些圖表：1. 事件反應時間 2. α波時間 3. 導回車道用時 4. 睡著")
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
    #excel_file = "E:\專題\data\s02_050921m.set\s02_050921m.xlsx" 
    
    try:
        df = pd.read_excel(excel_file)
        df = process_data(df)
        while(1): input_picture_format(df)
        exit()
        
    except Exception as e:
        print(f"發生錯誤: {e}")