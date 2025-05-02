import os
import glob
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

def process_eea_file(input_file, output_base_name, output_folder, export_csv=True, export_excel=False, num_channels=16, samples_per_channel=7680):
    """
    直接讀取 .eea 檔案並處理資料，輸出重組後的 CSV 或 Excel。
    """
    data_list = []
    with open(input_file, 'r', encoding='utf-8') as infile:
        for line in infile:
            row_values = line.strip().split(',')
            row_values = [x for x in row_values if x]
            data_list.extend(row_values)
    
    try:
        data_array = np.array(data_list, dtype=float)
    except ValueError:
        raise ValueError("資料無法轉換為數值，請確認 .eea 檔案內容格式")
    
    total_samples = num_channels * samples_per_channel
    if len(data_array) != total_samples:
        raise ValueError(f"資料筆數不正確，預期 {total_samples} 筆，但實際讀取 {len(data_array)} 筆")
    
    reshaped = np.reshape(data_array, (num_channels, samples_per_channel)).T
    columns = [
        "F7", "F3", "F4", "F8",
        "T3", "C3", "Cz", "C4",
        "T4", "T5", "P3", "Pz",
        "P4", "T6", "O1", "O2"
    ]
    new_df = pd.DataFrame(reshaped, columns=columns)

    if export_csv:
        new_df.to_csv(os.path.join(output_folder, output_base_name + ".csv"), index=False)
    if export_excel:
        new_df.to_excel(os.path.join(output_folder, output_base_name + ".xlsx"), index=False)

def convert_all_files(input_folder, output_folder, log_callback=None, export_csv=True, export_excel=False):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    eea_files = glob.glob(os.path.join(input_folder, "*.eea"))
    if not eea_files:
        if log_callback:
            log_callback(f"在 {input_folder} 中找不到任何 .eea 檔案")
        return
    
    for eea_file in eea_files:
        base_name = os.path.splitext(os.path.basename(eea_file))[0]
        try:
            process_eea_file(eea_file, base_name, output_folder, export_csv=export_csv, export_excel=export_excel)
            if log_callback:
                formats = []
                if export_csv: formats.append("CSV")
                if export_excel: formats.append("Excel")
                log_callback(f"轉換完成（{'、'.join(formats)}）：\n{eea_file}")
        except Exception as e:
            if log_callback:
                log_callback(f"轉換 {eea_file} 時發生錯誤：{e}\n")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EEA to CSV/Excel Converter")
        self.geometry("600x500")
        self.configure(bg="#f0f0f0")
        
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.export_csv = tk.BooleanVar(value=True)
        self.export_excel = tk.BooleanVar(value=False)
        
        tk.Label(self, text="輸入資料夾 (包含 .eea 檔案):", bg="#f0f0f0", font=("Arial", 12)).pack(pady=5)
        frame_input = tk.Frame(self, bg="#f0f0f0")
        frame_input.pack()
        tk.Entry(frame_input, textvariable=self.input_folder, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_input, text="瀏覽", command=self.browse_input, bg="#4CAF50", fg="white").pack(side=tk.LEFT)

        tk.Label(self, text="輸出資料夾 (儲存位置):", bg="#f0f0f0", font=("Arial", 12)).pack(pady=5)
        frame_output = tk.Frame(self, bg="#f0f0f0")
        frame_output.pack()
        tk.Entry(frame_output, textvariable=self.output_folder, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_output, text="瀏覽", command=self.browse_output, bg="#4CAF50", fg="white").pack(side=tk.LEFT)

        # 輸出格式選擇
        tk.Label(self, text="選擇輸出格式：", bg="#f0f0f0", font=("Arial", 12)).pack(pady=5)
        format_frame = tk.Frame(self, bg="#f0f0f0")
        format_frame.pack()
        tk.Checkbutton(format_frame, text="CSV", variable=self.export_csv, bg="#f0f0f0", font=("Arial", 10)).pack(side=tk.LEFT, padx=10)
        tk.Checkbutton(format_frame, text="Excel", variable=self.export_excel, bg="#f0f0f0", font=("Arial", 10)).pack(side=tk.LEFT, padx=10)

        tk.Button(self, text="開始轉換", command=self.start_conversion, bg="#2196F3", fg="white", font=("Arial", 12)).pack(pady=15)

        self.log_text = scrolledtext.ScrolledText(self, width=70, height=12, font=("Arial", 10))
        self.log_text.pack(pady=10)
    
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
    
    def browse_input(self):
        folder = filedialog.askdirectory(title="選擇包含 .eea 檔案的資料夾")
        if folder:
            self.input_folder.set(folder)
    
    def browse_output(self):
        folder = filedialog.askdirectory(title="選擇輸出資料夾")
        if folder:
            self.output_folder.set(folder)
    
    def start_conversion(self):
        input_folder = self.input_folder.get()
        output_folder = self.output_folder.get()
        if not input_folder or not output_folder:
            messagebox.showerror("錯誤", "請選擇輸入與輸出資料夾")
            return
        if not self.export_csv.get() and not self.export_excel.get():
            messagebox.showwarning("提醒", "請至少選擇一種輸出格式（CSV 或 Excel）")
            return
        self.log("開始轉換...")
        convert_all_files(
            input_folder,
            output_folder,
            log_callback=self.log,
            export_csv=self.export_csv.get(),
            export_excel=self.export_excel.get()
        )
        self.log("全部轉換完成！")

def main():
    app = App()
    app.mainloop()

if __name__ == '__main__':
    main()
