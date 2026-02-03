from matplotlib.axes import Axes
from matplotlib.figure import Figure
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from pandas import DataFrame
matplotlib.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]  # 設定中文字體
matplotlib.rcParams["axes.unicode_minus"] = False  # 解決負號顯示問題

data_column_name = ["秒數", "事件反應時間", "α波時間", "導回車道用時", "睡著", "眼動次數"]

def align_yaxis(ax: Axes, df: DataFrame, column_name: str):
    """
    確保 Y 軸的 0 位準與範圍邏輯統一
    """
    c_min = df[column_name].min()
    c_max = df[column_name].max()
    
    # 處理全為 NaN 或數據為空的情況
    if pd.isna(c_min) or pd.isna(c_max):
        return

    if c_min >= 0:
        ax_bottom = 0
        ax_top = c_max * 1.1 if c_max > 0 else 1
    elif c_max <= 0:
        ax_bottom = c_min * 1.1
        ax_top = 0
    else:
        # 正負值都有時，強制 0 置中
        max_abs = max(abs(c_min), abs(c_max))
        ax_bottom = -max_abs * 1.1
        ax_top = max_abs * 1.1
    
    ax.set_ylim(bottom=ax_bottom, top=ax_top)

def plot_line_data(df: DataFrame, column_name: str) -> tuple[Figure, Axes]:
    """
    繪製折線圖
    """
    #filtered_df = df[df[column_name].notna() & (df[column_name] != 0)]
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(df["秒數"], df[column_name], color="blueviolet", marker="o", markersize=4, linewidth=2, label=column_name)
    ax.set_xlabel("秒數", fontsize=12)
    ax.set_ylabel(column_name, fontsize=12, color="blueviolet")
    ax.tick_params(axis="y", labelcolor="blueviolet")
    ax.grid(True, alpha=0.3)
    ax.set_title(column_name + "分析", fontsize=14, fontweight="bold")
    ax.legend(loc="upper left")
    
    # 計算 ax 的 Y 軸範圍
    column_name_min = df[column_name].min()
    column_name_max = df[column_name].max()
    if column_name_min >= 0:
        ax_bottom = 0
        ax_top = None
    elif column_name_max <= 0:
        ax_bottom = None
        ax_top = 0
    else:
        max_abs = max(abs(column_name_min), abs(column_name_max))
        ax_bottom = -max_abs
        ax_top = max_abs
    
    # 設定 Y 軸範圍，讓 0 對齊 X 軸
    if ax_bottom is not None and ax_top is not None:
        ax.set_ylim(bottom=ax_bottom, top=ax_top)
    elif ax_bottom is not None:
        ax.set_ylim(bottom=ax_bottom)
    elif ax_top is not None:
        ax.set_ylim(top=ax_top)
    return fig, ax

def plot_bar_data(df: DataFrame, column_name: str) -> tuple[Figure, Axes]:
    """
        繪製長條圖
    """
    fig, ax = plt.subplots(figsize=(12, 6))
    
    ax.bar(df["秒數"], df[column_name], 
                  width=1, alpha=0.6, color="blueviolet", label=column_name)
    ax.set_xlabel("秒數", fontsize=12)
    ax.set_ylabel(column_name, fontsize=12, color="blueviolet")
    ax.tick_params(axis="y", labelcolor="blueviolet")
    ax.grid(True, alpha=0.3)
    ax.set_title(column_name + "分析", fontsize=14, fontweight="bold")
    ax.legend(loc="upper left")
    return fig, ax

def plot_sleep_area(df: DataFrame, color="blueviolet") -> tuple[Figure, Axes]:
    """
    特殊繪圖：只針對「睡著」欄位，從=1開始到=2停止，區域換底色
    """
    fig, ax = plt.subplots(figsize=(12, 6))

    # 繪製基礎折線圖 (可選，或只顯示底色)
    # ax.plot(df["秒數"], df["睡著"], color="gray", alpha=0.5, label="睡著狀態")

    # 找出所有區段
    start = None
    for i, val in enumerate(df["睡著"]):
        if val == 1 and start is None:
            start = df["秒數"].iloc[i]   # 記錄起始點
        elif val == 2 and start is not None:
            end = df["秒數"].iloc[i]     # 記錄結束點
            # 畫底色區域
            ax.axvspan(start, end, color=color, alpha=0.3 if start == df["秒數"].iloc[0] else 0.0)
            start = None  # 重置，準備下一段

    ax.set_xlabel("秒數", fontsize=12)
    ax.set_ylabel("睡著狀態", fontsize=12)
    ax.set_title("睡著區域標記", fontsize=14, fontweight="bold")
    # ax.legend(loc="upper left")
    ax.grid(True, alpha=0.3)

    return fig, ax

def plot_sleep_area_on_ax(ax: Axes, df: DataFrame, color="blueviolet", alpha: float = 0.3):
    # 保證排序
    df_sorted = df.sort_values("秒數")
    xlim = ax.get_xlim()  # 記住原本 X 範圍

    start = None
    label_added = False
    for i, val in enumerate(df_sorted["睡著"]):
        if val == 1 and start is None:
            start = df_sorted["秒數"].iloc[i]
        elif val == 2 and start is not None:
            end = df_sorted["秒數"].iloc[i]
            ax.axvspan(start, end, color=color, alpha=alpha, label="睡著區域" if not label_added else None)
            label_added = True
            start = None

    # 恢復原本 X 範圍以避免平移
    ax.set_xlim(xlim)

def plot_data_combined(fig1: Figure, ax1: Axes, chart1: int, mode1: int, 
                       fig2: Figure, ax2: Axes, chart2: int, mode2: int, df: DataFrame):
    """
    合併兩個圖表到一個圖中（使用雙軸）
    """
    # 關閉第二個圖表，因為我們要將它合併到第一個圖表中
    plt.close(fig2)
    
    # 獲取第二個圖表的數據
    column_name2 = data_column_name[chart2]
    color2 = "salmon"  # 預設顏色

    # 過濾0和NAN
    #filtered_df2 = df[df[column_name2].notna() & (df[column_name2] != 0)]
    
    # 在 ax1 上創建第二個 y 軸
    ax2_new = ax1.twinx()
    
    if chart2 == 4:
        plot_sleep_area_on_ax(ax2_new, df, color=color2)
        column_name2 = "睡著"
    else:
        # 根據模式判斷第二個圖表是折線圖還是長條圖，並繪製
        if mode2 == 2:
            # 如果是長條圖
            ax2_new.bar(df["秒數"], df[column_name2], width=1, alpha=0.6, color=color2, label=column_name2)
        else:
            # 如果是折線圖
            ax2_new.plot(df["秒數"], df[column_name2], color=color2, marker="o", markersize=4, 
                        linewidth=2, label=column_name2)


    # 設定第二個 y 軸的標籤和顏色
    ax2_new.set_ylabel(column_name2, fontsize=12, color=color2)
    ax2_new.tick_params(axis="y", labelcolor=color2)
    
    # 計算第二個 y 軸的範圍
    column_name2_min = df[column_name2].min()
    column_name2_max = df[column_name2].max()
    if column_name2_min >= 0:
        ax2_bottom = 0
        ax2_top = None
    elif column_name2_max <= 0:
        ax2_bottom = None
        ax2_top = 0
    else:
        max_abs = max(abs(column_name2_min), abs(column_name2_max))
        ax2_bottom = -max_abs
        ax2_top = max_abs
    
    # 設定第二個 y 軸的範圍
    if ax2_bottom is not None and ax2_top is not None:
        ax2_new.set_ylim(bottom=ax2_bottom, top=ax2_top)
    elif ax2_bottom is not None:
        ax2_new.set_ylim(bottom=ax2_bottom)
    elif ax2_top is not None:
        ax2_new.set_ylim(top=ax2_top)
    
    # 設定標題
    ax1.set_title(f"{data_column_name[chart1]}與{data_column_name[chart2]}分析", fontsize=14, fontweight="bold")
    
    # 合併圖例
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2_new, labels2_new = ax2_new.get_legend_handles_labels()
    ax1.legend(lines1 + lines2_new, labels1 + labels2_new, loc="upper left")
    
    # 調整版面 
    fig1.tight_layout()
    fig1.show()  # 顯示第一個圖表（使用 fig 對象的方法）
    #plt.close(fig1)  # 關閉第一個圖表，釋放記憶體


def plot_data_triple(df: DataFrame,
                     chart1: int, mode1: int,
                     chart2: int, mode2: int,
                     chart3: int, mode3: int):
    """
    三張圖疊加：以最長的 X 軸為基準，第三張不顯示 Y 軸
    """
    # 主圖 (第一張)
    fig, ax1 = check_ouput_picture(df, data_column_name[chart1], mode1)
    align_yaxis(ax1, df, data_column_name[chart1])

    # 第二張：建立右側 y 軸
    ax2 = ax1.twinx()
    column_name2 = data_column_name[chart2]
    #filtered_df2 = df[(df[column_name2].notna()) & (df[column_name2] != 0)]
    if mode2 == 2:
        ax2.bar(df["秒數"], df[column_name2],
                width=0.8, alpha=0.6, color="salmon", label=column_name2)
    else:
        ax2.plot(df["秒數"], df[column_name2],
                 color="salmon", marker="o", markersize=4,
                 linewidth=2, label=column_name2)
    align_yaxis(ax2, df, column_name2)
    ax2.set_ylabel(column_name2, fontsize=12, color="salmon", alpha=0.1)
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
    xmin = min(df["秒數"].min(), df["秒數"].min(), filtered_df3["秒數"].min())
    xmax = max(df["秒數"].max(), df["秒數"].max(), filtered_df3["秒數"].max())
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

def check_ouput_picture(df: DataFrame, column_name: str, mode: int) -> tuple[Figure, Axes]:
    if mode == 1:
        return plot_line_data(df, column_name)
    elif mode == 2:
        return plot_bar_data(df, column_name)