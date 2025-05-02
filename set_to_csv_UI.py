import os
import mne
import pandas as pd
import tkinter as tk
import ttkbootstrap as ttk
from tkinter import filedialog, messagebox
from ttkbootstrap.constants import *
import subprocess

class EEGConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("EEG .SET 轉 .CSV 工具")
        master.geometry("760x480")
        master.resizable(False, False)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # 標題
        ttk.Label(self.master, text="EEG .set 檔案轉換工具", font=("Arial", 16, "bold")).pack(pady=10)

        # 資料夾選擇區塊
        frame1 = ttk.Labelframe(self.master, text="輸入資料夾", padding=10)
        frame1.pack(padx=20, pady=5, fill='x')
        ttk.Entry(frame1, textvariable=self.input_path, width=80).pack(side='left', padx=5)
        ttk.Button(frame1, text="瀏覽", command=self.browse_input).pack(side='right')

        frame2 = ttk.Labelframe(self.master, text="輸出資料夾", padding=10)
        frame2.pack(padx=20, pady=5, fill='x')
        ttk.Entry(frame2, textvariable=self.output_path, width=80).pack(side='left', padx=5)
        ttk.Button(frame2, text="瀏覽", command=self.browse_output).pack(side='right')

        # 開始轉換按鈕與進度條
        ttk.Button(self.master, text="開始轉換", bootstyle=SUCCESS, command=self.convert_files).pack(pady=10)
        self.progress = ttk.Progressbar(self.master, mode='determinate', length=700)
        self.progress.pack(pady=5)

        # 日誌顯示區
        self.log_box = tk.Text(self.master, height=14, font=("Courier", 10))
        self.log_box.pack(padx=20, pady=10, fill="both", expand=True)

    def browse_input(self):
        path = filedialog.askdirectory(title="選擇 .set 檔案資料夾")
        if path:
            self.input_path.set(path)

    def browse_output(self):
        path = filedialog.askdirectory(title="選擇輸出資料夾")
        if path:
            self.output_path.set(path)

    def log(self, message):
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.master.update()

    def convert_files(self):
        input_folder = self.input_path.get()
        output_folder = self.output_path.get()

        if not input_folder or not output_folder:
            messagebox.showerror("錯誤", "請選擇輸入與輸出資料夾")
            return

        files = [f for f in os.listdir(input_folder) if f.endswith(".set")]
        if not files:
            self.log("⚠️ 找不到 .set 檔案")
            return

        total = len(files)
        self.progress["maximum"] = total
        self.progress["value"] = 0

        for idx, file_name in enumerate(files, start=1):
            try:
                self.log(f"🔄 處理中：{file_name}")
                raw = mne.io.read_raw_eeglab(os.path.join(input_folder, file_name), preload=True)
                data = raw.get_data()
                df = pd.DataFrame(data.T, columns=raw.info['ch_names'])

                output_path = os.path.join(output_folder, file_name.replace(".set", ".csv"))
                df.to_csv(output_path, index=False)
                self.log(f"✅ 已輸出至：{output_path}")
            except Exception as e:
                self.log(f"❌ 轉換失敗：{file_name}\n   錯誤：{str(e)}")

            self.progress["value"] = idx
            self.master.update_idletasks()

        self.log("🎉 所有檔案已完成轉換！")
        self.open_folder(output_folder)

    def open_folder(self, path):
        try:
            if os.name == 'nt':  # Windows
                subprocess.Popen(f'explorer "{path}"')
            elif os.name == 'posix':  # macOS or Linux
                subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', path])
        except Exception as e:
            self.log(f"⚠️ 無法開啟資料夾：{e}")

# 啟動程式
if __name__ == "__main__":
    root = ttk.Window(themename="flatly")  # 你也可以改用 'journal', 'darkly', 'superhero' 等主題
    app = EEGConverterApp(root)
    root.mainloop()
