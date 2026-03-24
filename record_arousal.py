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

def detect_eye_movements(raw, target_channels, output_path):
    raw.pick(target_channels)
    data, times = raw.get_data(), raw.times
    sfreq = raw.info['sfreq']
    combined_signal = (np.abs(data[0]) + np.abs(data[1])) / 2

    # 使用 MAD 方法計算動態閾值
    median = np.median(combined_signal)
    mad = np.median(np.abs(combined_signal - median))
    threshold = median + 3 * mad
    # 動態調整 prominence threshold
    dynamic_prominence = np.percentile(combined_signal, 95)
    peaks, _ = find_peaks(combined_signal, height=threshold, 
                          prominence=dynamic_prominence,
                          distance=int(sfreq * 0.4),
                          width=(int(sfreq * 0), int(sfreq * 0.5)))
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