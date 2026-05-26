import argparse
import os
import sys
import mne
import numpy as np
from scipy.signal import find_peaks

DYNAMIC_PROMINENCE = 0.00007


def custom_eye_movement_peaks(signal, sfreq, height_thresh, prominence_thresh):
    """
    Parameters:
    - signal: 綜合後的訊號 (1D numpy array)
    - sfreq: 取樣率 (例如 250)
    - height_thresh: 絕對高度門檻 (median + 0.5 * mad)
    - prominence_thresh: 突起度門檻
    """
    peaks = []
    
    # 將秒數轉換為點數 (Ticks)
    window_ticks = int(sfreq * 0.15) 
    
    # 為了安全檢查，前後留出一個窗口的邊界
    for i in range(window_ticks, len(signal) - window_ticks):

        current_val = signal[i]
        
        # ensure that the current point is a local maximum
        if current_val < signal[i - 1] or current_val < signal[i + 1]:
            continue
            
        # -------------------------------------------------------------
        # 關卡 2：絕對高度檢查
        # -------------------------------------------------------------
        if current_val < height_thresh:
            continue
            
        # -------------------------------------------------------------
        # 關卡 3：🛡️ 雙邊驟升與驟降檢查（直接消滅「單邊階梯跳躍」）
        # -------------------------------------------------------------
        # 撈出左邊與右邊窗口內的最低點
        left_start = max(0, i - window_ticks)
        left_end = i
        right_start = i
        right_end = min(len(signal), i + window_ticks)

        left_window = signal[left_start:left_end]
        right_window = signal[right_start:right_end]
        
        left_min = np.min(left_window)
        right_min = np.min(right_window)
        left_min_idx = left_start + np.argmin(left_window)
        right_min_idx = i + np.argmin(right_window)
        window_max_idx = left_start + np.argmax(signal[left_start : right_end + 1])

        if window_max_idx != i:
            continue
        
        # 計算左爬升與右跌落
        left_rise = current_val - left_min
        right_fall = current_val - right_min
        
        # 雙邊都必須大於你設定的突起度門檻（確保不是單邊跳躍，也不是微小毛刺）
        if min(left_rise, right_fall) < prominence_thresh:
            continue
            
        # 成功通過所有幾何形狀安檢！這絕對是一顆標準的眼動脈衝
        peaks.append(i)
        
    return np.array(peaks)

def detect_eye_movements(raw, target_channels, output_path):
    raw.pick(target_channels)
    data, times = raw.get_data(), raw.times
    sfreq = raw.info['sfreq']
    signal = data[1]

    # 使用 MAD 方法計算動態閾值
    median = np.median(signal)
    mad = np.median(np.abs(signal - median))
    threshold = median + 0.5 * mad

    peaks = custom_eye_movement_peaks(signal, sfreq, threshold, DYNAMIC_PROMINENCE)
    eye_move_seconds = np.unique(np.ceil(times[peaks])).astype(int)
    total_duration = int(times[-1])
    eye_move_seconds = eye_move_seconds[eye_move_seconds <= total_duration]
    
    n_count = len(eye_move_seconds)
    output_data = [n_count] + eye_move_seconds.tolist()
    output_string = ",".join(map(str, output_data))
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(output_string)

    print(f"處理完成！")
    print(f"總眼動秒數：{n_count}")
    print(f"結果已存入：{output_path}")