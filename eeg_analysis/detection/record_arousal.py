import argparse
import os
import sys
import mne
import numpy as np

# === 偵測參數（與 eye_movement_validation.py 對齊）===
HEIGHT_MAD_SCALE = 0.5          # 高度門檻 = median + HEIGHT_MAD_SCALE * MAD
MIN_PEAK_TO_SHOULDER_DELTA = 0.00007  # peak 與左右兩側較高谷值的最小落差門檻
PEAK_WINDOW_SEC = 0.15          # 左右觀察窗口（秒）

# 濾波設定（zero-phase band-pass，與 eye_movement_validation.py 預設一致）
L_FREQ = 0.1   # 高通截止頻率 (Hz)，去除慢速基線飄移；設 None/0 可關閉
H_FREQ = 10.0  # 低通截止頻率 (Hz)，去除高頻肌電/雜訊；設 None/0 可關閉
NOTCH_FREQ = 0.0  # 電源雜訊陷波頻率 (Hz)；預設關閉


def compute_peak_to_shoulder_amplitude(current_val, left_min, right_min):
    """peak 與左右兩側「較高」谷值之間的落差（越大代表訊號起伏越明顯）。"""
    if current_val is None or left_min is None or right_min is None:
        return None
    return float(current_val) - float(max(left_min, right_min))


def apply_noise_filters(picked_raw, target_channels, *, l_freq, h_freq, notch_freq):
    """套用電源雜訊陷波 + zero-phase band-pass 濾波。

    使用 zero-phase 濾波（MNE 預設的 FIR/firwin 設計）是為了避免 peak 的時間點被
    偏移，因為後續會把偵測到的 peak 對應回特定的秒數。
    """
    sfreq = float(picked_raw.info["sfreq"])
    nyquist = sfreq / 2.0

    if notch_freq and notch_freq > 0:
        harmonics = np.arange(notch_freq, nyquist, notch_freq)
        if harmonics.size:
            picked_raw.notch_filter(
                harmonics,
                picks=target_channels[:2],
                fir_design="firwin",
                phase="zero",
                verbose=False,
            )

    l_freq_arg = float(l_freq) if l_freq and l_freq > 0 else None
    h_freq_arg = float(h_freq) if h_freq and h_freq > 0 else None
    if l_freq_arg is not None or h_freq_arg is not None:
        picked_raw.filter(
            l_freq=l_freq_arg,
            h_freq=h_freq_arg,
            picks=target_channels[:2],
            fir_design="firwin",
            phase="zero",
            verbose=False,
        )

    return picked_raw


def custom_eye_movement_peaks(signal, sfreq, height_thresh, peak_to_shoulder_thresh):
    """
    Parameters:
    - signal: 訊號 (1D numpy array)
    - sfreq: 取樣率 (例如 250)
    - height_thresh: 絕對高度門檻 (median + HEIGHT_MAD_SCALE * mad)
    - peak_to_shoulder_thresh: peak 與左右較高谷值的最小落差門檻

    判斷邏輯與 eye_movement_validation.py 的 build_candidate_diagnostic /
    collect_candidate_diagnostics 保持一致：
      1. 嚴格區域最大值（嚴格大於左右相鄰點）
      2. 絕對高度門檻
      3. 在左右窗口內，該點必須是整個窗口內的最大值（is_local_max）
      4. peak 與左右兩側「較高」谷值的落差必須大於門檻（peak_to_shoulder_ok）
    """
    window_ticks = max(int(round(sfreq * PEAK_WINDOW_SEC)), 1)
    peaks = []

    for i in range(window_ticks, len(signal) - window_ticks):
        current_val = signal[i]

        # -------------------------------------------------------------
        # 關卡 1：嚴格區域最大值（與驗證腳本一致，使用 <= 而非 <）
        # -------------------------------------------------------------
        if current_val <= signal[i - 1] or current_val <= signal[i + 1]:
            continue

        # -------------------------------------------------------------
        # 關卡 2：絕對高度檢查
        # -------------------------------------------------------------
        if current_val < height_thresh:
            continue

        # -------------------------------------------------------------
        # 關卡 3：窗口內最大值檢查（確認不是被更高的鄰近點蓋過）
        # -------------------------------------------------------------
        left_start = max(0, i - window_ticks)
        left_end = i
        right_start = i
        right_end = min(len(signal), i + window_ticks)

        left_window = signal[left_start:left_end]
        right_window = signal[right_start:right_end]

        if left_window.size == 0 or right_window.size == 0:
            continue

        left_min = np.min(left_window)
        right_min = np.min(right_window)

        current_window = signal[left_start:right_end]
        max_index = left_start + int(np.argmax(current_window))
        if max_index != i:
            continue

        # -------------------------------------------------------------
        # 關卡 4：peak-to-shoulder 落差檢查
        # -------------------------------------------------------------
        peak_to_shoulder_delta = compute_peak_to_shoulder_amplitude(current_val, left_min, right_min)
        if peak_to_shoulder_delta is None or peak_to_shoulder_delta <= peak_to_shoulder_thresh:
            continue

        # 成功通過所有檢查！這是一顆標準的眼動脈衝
        peaks.append(i)

    return np.array(peaks)


def detect_eye_movements(
    raw,
    target_channels,
    output_path,
    *,
    start_second=1,
    end_second=None,
):
    """Detect eye-movement seconds, save DAT, and return the saved seconds.

    Seconds are one-based and event timestamps are rounded upward.  Optional
    bounds allow Function One to derive its eye mask using only seconds 1--300.
    """
    if start_second < 1:
        raise ValueError("start_second 必須大於等於 1")
    if end_second is not None and end_second < start_second:
        raise ValueError("end_second 必須大於等於 start_second")

    picked_raw = raw.copy().pick(target_channels[:2])
    if end_second is not None:
        available_end = float(picked_raw.times[-1])
        picked_raw.crop(tmin=0.0, tmax=min(float(end_second), available_end))
    picked_raw.load_data()
    picked_raw = apply_noise_filters(
        picked_raw,
        target_channels,
        l_freq=L_FREQ,
        h_freq=H_FREQ,
        notch_freq=NOTCH_FREQ,
    )

    data, times = picked_raw.get_data(), picked_raw.times
    sfreq = float(picked_raw.info["sfreq"])
    fp2_index = next(
        (
            index
            for index, channel_name in enumerate(picked_raw.ch_names)
            if "fp2" in channel_name.casefold()
        ),
        len(picked_raw.ch_names) - 1,
    )
    signal = data[fp2_index]

    # 使用 MAD 方法計算動態高度閾值
    median = np.median(signal)
    mad = np.median(np.abs(signal - median))
    height_thresh = median + HEIGHT_MAD_SCALE * mad

    peaks = custom_eye_movement_peaks(signal, sfreq, height_thresh, MIN_PEAK_TO_SHOULDER_DELTA)
    eye_move_seconds = np.unique(np.ceil(times[peaks])).astype(int)
    last_included_second = int(np.ceil(times[-1]))
    if end_second is not None:
        last_included_second = min(last_included_second, int(end_second))
    eye_move_seconds = eye_move_seconds[
        (eye_move_seconds >= int(start_second))
        & (eye_move_seconds <= last_included_second)
    ]

    n_count = len(eye_move_seconds)
    output_data = [n_count] + eye_move_seconds.tolist()
    output_string = ",".join(map(str, output_data))
    output_dir = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(output_dir, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(output_string + "\n")

    print(f"處理完成！")
    print(f"總眼動秒數：{n_count}")
    print(f"結果已存入：{output_path}")
    return eye_move_seconds.tolist()
