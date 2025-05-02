import os
import glob
import pyedflib
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

def convert_edf_to_csv(edf_folder, csv_folder, log_widget):
    edf_paths = glob.glob(os.path.join(edf_folder, '*.edf'))
    if not edf_paths:
        messagebox.showwarning("警告", "找不到任何 EDF 檔案。")
        return

    os.makedirs(csv_folder, exist_ok=True)

    for edf_path in edf_paths:
        try:
            with pyedflib.EdfReader(edf_path) as edf:
                n_signals = edf.signals_in_file
                labels = edf.getSignalLabels()
                n_samples = edf.getNSamples()[0]
                data = np.zeros((n_signals, n_samples))
                for i in range(n_signals):
                    data[i, :] = edf.readSignal(i)

            df = pd.DataFrame(data.T, columns=labels)
            base = os.path.splitext(os.path.basename(edf_path))[0]
            csv_path = os.path.join(csv_folder, f'{base}.csv')
            df.to_csv(csv_path, index=False)
            log_widget.insert(tk.END, f'✔ 已儲存：{csv_path}\n')
            log_widget.see(tk.END)
        except Exception as e:
            log_widget.insert(tk.END, f'❌ 錯誤：{edf_path} - {str(e)}\n')
            log_widget.see(tk.END)

    messagebox.showinfo("完成", "全部轉換完成！")

def browse_edf_folder():
    folder = filedialog.askdirectory()
    if folder:
        edf_folder_var.set(folder)

def browse_csv_folder():
    folder = filedialog.askdirectory()
    if folder:
        csv_folder_var.set(folder)

def start_conversion():
    edf_folder = edf_folder_var.get()
    csv_folder = csv_folder_var.get()
    if not edf_folder or not csv_folder:
        messagebox.showwarning("警告", "請選擇 EDF 資料夾與 CSV 輸出資料夾。")
        return
    convert_edf_to_csv(edf_folder, csv_folder, log_area)

# GUI 建立
root = tk.Tk()
root.title("EDF 轉 CSV 工具")

tk.Label(root, text="EDF 資料夾：").grid(row=0, column=0, sticky="e")
edf_folder_var = tk.StringVar()
tk.Entry(root, textvariable=edf_folder_var, width=50).grid(row=0, column=1)
tk.Button(root, text="瀏覽", command=browse_edf_folder).grid(row=0, column=2)

tk.Label(root, text="CSV 輸出資料夾：").grid(row=1, column=0, sticky="e")
csv_folder_var = tk.StringVar()
tk.Entry(root, textvariable=csv_folder_var, width=50).grid(row=1, column=1)
tk.Button(root, text="瀏覽", command=browse_csv_folder).grid(row=1, column=2)

tk.Button(root, text="開始轉換", command=start_conversion, bg="lightblue").grid(row=2, column=1, pady=10)

log_area = scrolledtext.ScrolledText(root, width=70, height=20)
log_area.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

root.mainloop()
