import pyedflib
import numpy as np
from openpyxl import Workbook  # 新增
import os  # 新增
from openpyxl.utils import get_column_letter

def process_alpha_data(ws, base_path, total_seconds):
    """
    處理 α 波資料並以 10 秒滑動視窗填入 C 欄。
    例如輸入 abc_raw.EDF 時，讀取同資料夾的 abc_alpha.dat。
    """
    dat_dir = os.path.dirname(base_path)
    edf_stem = os.path.basename(base_path)
    record_id = edf_stem[:-4] if edf_stem.lower().endswith("_raw") else edf_stem
    alpha_path = os.path.join(dat_dir, f"{record_id}_alpha.dat")

    numbers = []
    if not os.path.exists(alpha_path):
        print(f"警告：找不到 α 波資料檔案 {alpha_path}")
    else:
        with open(alpha_path, "r", encoding="utf-8") as f:
            content = f.read().strip()

        if not content:
            print(f"警告：{alpha_path} 檔案為空")
        else:
            try:
                numbers = [int(x.strip()) for x in content.split(",") if x.strip()]
            except ValueError:
                print(f"警告：{alpha_path} 資料格式不正確")
                numbers = []

    last_second = int(total_seconds)

    # 第一個數字為 α 波秒數的總數量，後面才是實際秒數。
    alpha_counts = {}
    for second in numbers[1:]:
        if 0 <= second <= last_second:
            alpha_counts[second] = alpha_counts.get(second, 0) + 1
        else:
            print(f"警告：忽略超出 EDF 時間範圍的 α 波秒數 {second}")

    # 每秒統計當前秒與前 9 秒，等同 rolling(window=10) 的結果。
    rolling_alpha_counts = {}
    rolling_count = 0
    for second in range(last_second + 1):
        rolling_count += alpha_counts.get(second, 0)
        if second >= 10:
            rolling_count -= alpha_counts.get(second - 10, 0)
        rolling_alpha_counts[second] = rolling_count

    # 保留原本的 Status 資料列，再補上尚不存在的每一整數秒資料列。
    rows = []
    existing_integer_seconds = set()
    for row in range(2, ws.max_row + 1):
        values = [ws.cell(row=row, column=column).value for column in range(1, 7)]
        a_value = values[0]
        if a_value is None:
            continue

        try:
            second_value = float(a_value)
        except (TypeError, ValueError):
            continue

        if second_value.is_integer() and 0 <= second_value <= last_second:
            integer_second = int(second_value)
            values[2] = rolling_alpha_counts[integer_second]
            existing_integer_seconds.add(integer_second)

        rows.append((second_value, row, values))

    for second, alpha_count in rolling_alpha_counts.items():
        if second not in existing_integer_seconds:
            rows.append((float(second), -1, [float(second), None, alpha_count, None, None, None]))

    rows.sort(key=lambda item: (item[0], item[1]))

    # 重新寫回排序後的列，讓每一整數秒都有 α 波滑動視窗結果。
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    for row, (_, _, values) in enumerate(rows, start=2):
        for column, value in enumerate(values, start=1):
            ws.cell(row=row, column=column, value=value)

def process_eye_blink_data(ws, base_path, total_seconds):
    """
    處理眼動資料並以 30 秒滑動視窗填入 F 欄。
    base_path: EDF 檔案的路徑（不含副檔名）
    """
    # 尋找對應的 .dat 檔案
    dat_dir = os.path.dirname(base_path)
    edf_stem = os.path.basename(base_path)
    record_id = edf_stem[:-4] if edf_stem.lower().endswith("_raw") else edf_stem
    dat_filename = record_id + "_raw_arousal info.dat"
    dat_path = os.path.join(dat_dir, dat_filename)
    
    numbers = []
    if not os.path.exists(dat_path):
        print(f"警告：找不到眼動資料檔案 {dat_path}")
    else:
        with open(dat_path, "r", encoding="utf-8") as f:
            content = f.read().strip()

        if not content:
            print(f"警告：{dat_path} 檔案為空")
        else:
            try:
                numbers = [int(x.strip()) for x in content.split(",") if x.strip()]
            except ValueError:
                print(f"警告：{dat_path} 資料格式不正確")
                numbers = []

    last_second = int(total_seconds)
    eye_counts = {}
    for second in numbers[1:]:
        if 0 <= second <= last_second:
            eye_counts[second] = eye_counts.get(second, 0) + 1
        else:
            print(f"警告：忽略超出 EDF 時間範圍的眼動秒數 {second}")

    # 每秒統計當前秒與前 29 秒，等同 rolling(window=30) 的結果。
    rolling_eye_counts = {}
    rolling_count = 0
    for second in range(last_second + 1):
        rolling_count += eye_counts.get(second, 0)
        if second >= 30:
            rolling_count -= eye_counts.get(second - 30, 0)
        rolling_eye_counts[second] = rolling_count

    # process_alpha_data 已建立完整逐秒時間軸；只需將結果填入每個整數秒。
    for row in range(2, ws.max_row + 1):
        a_value = ws.cell(row=row, column=1).value
        if a_value is None:
            continue

        try:
            second_value = float(a_value)
        except (TypeError, ValueError):
            continue

        if second_value.is_integer() and 0 <= second_value <= last_second:
            ws.cell(row=row, column=6, value=rolling_eye_counts[int(second_value)])

def check_status_253(edf_path, tolerance=0.05):
    f = pyedflib.EdfReader(edf_path)

    # --- 準備 xlsx 檔案與表頭 ---
    base_path, _ = os.path.splitext(edf_path)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "data")
    os.makedirs(output_dir, exist_ok=True)
    xlsx_filename = os.path.basename(base_path) + ".xlsx"
    xlsx_path = os.path.join(output_dir, xlsx_filename)

    wb = Workbook()
    ws = wb.active
    ws.title = "result"

    # A1~F1 標題
    ws["A1"] = "秒數"
    ws["B1"] = "事件反應時間"
    ws["C1"] = "α波次數（10秒滑動）"
    ws["D1"] = "導回車道用時"
    ws["E1"] = "睡著"
    ws["F1"] = "眼動次數（30秒滑動）"

    next_row = 2  # 下一筆資料要寫入的列數（從第2列開始）

    # ... 下面保持原本程式 ...
    channel_labels = f.getSignalLabels()
    status_index = None
    for i, label in enumerate(channel_labels):
        if 'status' in label.lower():
            status_index = i
            break

    if status_index is None:
        print("未找到 Status 通道，請確認標籤名稱")
        f.close()
        return

    status_signal = f.readSignal(status_index)
    sample_rate = int(f.getSampleFrequency(status_index))
    total_samples = len(status_signal)
    total_seconds = total_samples // sample_rate

    print(f"取樣率: {sample_rate} Hz, 總長度: {total_seconds} 秒")

    stage = 1
    for sec in range(total_seconds):
        start = sec * sample_rate
        end = start + sample_rate
        segment = status_signal[start:end]
        for i in range(len(segment)):
            #print(segment[i])
            if segment[i] > 1:
                if stage == 1:
                    sec_251 = sec + 0.002 * i
                    stage = 2
                elif stage == 2:
                    sec_253 = sec + 0.002 * i
                    stage = 3
                elif stage == 3:
                    sec_254 = sec + 0.002 * i

                    # --- 新增：在 stage==3 時把資料寫入 xlsx ---
                    ws.cell(row=next_row, column=1, value=int(sec_251))                              # 秒數（無條件捨去）
                    ws.cell(row=next_row, column=2, value=float(f"{sec_253 - sec_251:.1f}"))        # 事件反應時間
                    # C 欄 α波資料會在 process_alpha_data 補入
                    ws.cell(row=next_row, column=4, value=float(f"{sec_254 - sec_253:.1f}"))        # 導回車道用時
                    # E 欄 睡著：暫時留空

                    # print(
                    #     f"第{sec_253:.1f}秒的事件反應時間：{sec_253 - sec_251:.1f}秒, "
                    #     f"導回車道用時：{sec_254 - sec_253:.1f}秒"
                    # )

                    next_row += 1
                    stage = 1
                # 若有需要，可把 debug 的 print 打開
                # print(f"秒數 {sec + 0.002 * i} -> Status 平均值: {segment[i]:.2f}, stage={stage}")

    f.close()

    # --- 處理 α 波資料並填入 C 欄 ---
    process_alpha_data(ws, base_path, total_seconds)

    # --- 處理眼動資料並填入 F 欄 ---
    process_eye_blink_data(ws, base_path, total_seconds)
    
    # --- 新增：儲存 xlsx 檔案 ---
    wb.save(xlsx_path)
    #print(f"結果已儲存到: {xlsx_path}")

if __name__ == "__main__":
    edf_file = input("請輸入 EDF 檔案路徑: ")
    edf_file = edf_file.replace('"', '')
    check_status_253(edf_file)
