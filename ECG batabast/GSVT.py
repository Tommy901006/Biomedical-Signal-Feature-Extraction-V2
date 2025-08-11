import os
import shutil
import pandas as pd

# 設定路徑
excel_path = "GSVT.xlsx"  # 你的 Excel 檔案
source_folder = r"D:\ECGDataDenoised\labeled_output"  # A 資料夾路徑
target_folder = r"D:\ECGDataDenoised\labeled_output\GSVT"  # B 資料夾路徑

# 讀取 Excel
df = pd.read_excel(excel_path)

# 確保 FileName 欄位存在
if "FileName" not in df.columns:
    raise ValueError("Excel 沒有 'FileName' 欄位")

# 建立目標資料夾（如果不存在）
os.makedirs(target_folder, exist_ok=True)

# 遍歷檔案清單
for file_name in df["FileName"]:
    # 自動補上 .csv 副檔名（如果沒有）
    if not file_name.lower().endswith(".csv"):
        file_name = f"{file_name}.csv"
    
    src_path = os.path.join(source_folder, file_name)
    dst_path = os.path.join(target_folder, file_name)

    if os.path.exists(src_path):
        shutil.move(src_path, dst_path)  # 剪下貼上（移動檔案）
        print(f"已移動：{file_name}")
    else:
        print(f"找不到檔案：{file_name}")

print("移動完成！")
