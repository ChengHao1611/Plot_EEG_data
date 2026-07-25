import numpy as np

def load_dat_file(filepath):
    """讀取 .dat 檔案並回傳秒數集合"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            parts = content.replace('\n', '').split(',')
            # 第一個數字是總數 n，後面是秒數
            seconds = set(map(int, parts[1:]))
            return seconds
    except Exception as e:
        print(f"讀取檔案 {filepath} 失敗: {e}")
        return set()

if __name__ == "__main__":
    manual_file = "E:\專題\data\s53_090918n.set\s53_090918n_raw_arousal info.dat" # 手動標記的 .dat 檔案路徑
    auto_file = "E:\專題\data\s53_090918n.set\eyeblink.dat"
    # edf_file = "./EDF/s09_060317n_raw.EDF"  # 用於比對的 EDF 檔案路徑
    report_file = "comparison_details.txt" # 詳細報告輸出的路徑

    # 載入數據
    manual_set = load_dat_file(manual_file)
    auto_set = load_dat_file(auto_file)

    # 計算差異
    missed = sorted(list(manual_set - auto_set))  # 漏抓 (FN)
    extra = sorted(list(auto_set - manual_set))   # 誤報 (FP)
    matches = sorted(list(manual_set.intersection(auto_set))) # 正確 (TP)

    # 計算指標
    tp = len(matches)
    fp = len(extra)
    fn = len(missed)

    precision = tp / len(auto_set) if auto_set else 0
    recall = tp / len(manual_set) if manual_set else 0
    csi = tp / (tp + fp + fn) if (tp + fp + fn) > 0 else 0

    # --- 1. Terminal 輸出 (只保留關鍵績效指標) ---
    print(f"{' 偵測效能總結 ':=^40}")
    print(f"正確對應 (TP): {tp}")
    print(f"漏抓數量 (FN): {fn}")
    print(f"誤報數量 (FP): {fp}")
    print("-" * 40)
    print(f"準確率 (Precision): {precision:.2%}")
    print(f"回溯率 (Recall): {recall:.2%}")
    print(f"完整率 (CSI): {csi:.2%}")
    print(f"詳細清單已寫入: {report_file}")
    print("=" * 40)

    # --- 2. 寫入檔案 (存放詳細秒數清單) ---
    with open(report_file, "w", encoding="utf-8") as f:
        f.write(f"EEG 眼動偵測比對詳細報告\n")
        f.write(f"==========================\n\n")
        
        f.write(f"【正確對應 (Matches/TP)】共 {tp} 秒:\n")
        f.write(f"{matches}\n\n")
        
        f.write(f"【程式漏抓 (Missed/FN)】共 {fn} 秒:\n")
        f.write(f"說明：手動標記有，但程式沒偵測到的秒數。\n")
        f.write(f"{missed}\n\n")
        
        f.write(f"【程式誤報 (Extra/FP)】共 {fp} 秒:\n")
        f.write(f"說明：程式有偵測到，但手動標記沒有的秒數。\n")
        f.write(f"{extra}\n")