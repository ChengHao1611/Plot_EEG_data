import argparse
import os
import sys
import mne
import numpy as np
from scipy.signal import find_peaks

def get_args():
    parser = argparse.ArgumentParser(description='Detect eye movements from an EDF file')
    parser.add_argument("--file", type=str, help='The path to the EDF file')
    return parser.parse_args()

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
    max_width_ticks = int(sfreq * 0.3)
    
    # 為了安全檢查，前後留出一個窗口的邊界
    for i in range(window_ticks, len(signal) - window_ticks):
        current_val = signal[i]
        
        # ensure that the current point is a local maximum
        if current_val <= signal[i - 1] or current_val <= signal[i + 1]:
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
        left_window = signal[i - window_ticks : i]
        right_window = signal[i : i + window_ticks]
        
        left_min = np.min(left_window)
        right_min = np.min(right_window)
        
        # 計算左爬升與右跌落
        left_rise = current_val - left_min
        right_fall = current_val - right_min
        
        # 雙邊都必須大於你設定的突起度門檻（確保不是單邊跳躍，也不是微小毛刺）
        if left_rise < prominence_thresh or right_fall < prominence_thresh:
            continue
            
        # -------------------------------------------------------------
        # 關卡 4：🛡️ 寬度與非平頂檢查（直接消滅「2秒平頂大山丘」）
        # -------------------------------------------------------------
        # 檢查在更廣的範圍（例如 0.35 秒）內，訊號有沒有確實「吐回低點」
        # 如果訊號一直維持在高位（平頂山），那它的最低點就不會夠低
        broad_window_ticks = int(sfreq * 0.35)
        start_idx = max(0, i - broad_window_ticks)
        end_idx = min(len(signal), i + broad_window_ticks)
        
        # 在這個稍寬的窗口內，訊號必須跌回原本高度的 30% 以下
        # 如果跌不回去，代表它是一個肥胖的山丘或平頂，直接淘汰！
        floor_level = current_val - (prominence_thresh * 0.7)
        if signal[start_idx] > floor_level or signal[end_idx - 1] > floor_level:
            continue
            
        # 成功通過所有幾何形狀安檢！這絕對是一顆標準的眼動脈衝
        peaks.append(i)
        
    return np.array(peaks)

def detect_eye_movements(raw, target_channels, output_path):
    raw.pick(target_channels)
    raw.filter(l_freq=0.5, h_freq=10, fir_design='firwin', verbose=False)
    data, times = raw.get_data(), raw.times
    sfreq = raw.info['sfreq']
    combined_signal = (np.abs(data[0]) + np.abs(data[1])) / 2

    # 使用 MAD 方法計算動態閾值
    median = np.median(combined_signal)
    mad = np.median(np.abs(combined_signal - median))
    threshold = median + 0.5 * mad
    # 動態調整 prominence threshold
    dynamic_prominence = 0.00006
    # peaks, _ = find_peaks(combined_signal, height=threshold, 
    #                       prominence=dynamic_prominence,
    #                       distance=int(sfreq * 0.4),
    #                       width=(int(sfreq * 0), int(sfreq * 0.5)))
    peaks = custom_eye_movement_peaks(combined_signal, sfreq, threshold, dynamic_prominence)
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

if __name__ == "__main__":
    args = get_args()
    file_path = args.file
    if file_path is None:
        print('Hint: The --file parameter was not detected.')
        file_path = input('Please enter the path to the EDF file: ').strip().replace('"', '').replace("'", "")
    if not os.path.exists(file_path):
        print(f"Error: The file '{file_path}' does not exist.")
        sys.exit(1)
    file_dir = os.path.dirname(file_path)
    output_path = os.path.join(file_dir, "eyeblink.dat")
    raw = mne.io.read_raw_edf(file_path, preload=True)

    raw.filter(1.5, 10, fir_design='firwin')
    all_channels = raw.ch_names
    target_channels = [ch for ch in all_channels if 'fp1' in ch.lower() or 'fp2' in ch.lower()]

    if len(target_channels) < 2:
        print("找不到 FP1 或 FP2 通道")
    else:
        detect_eye_movements(raw, target_channels, output_path)