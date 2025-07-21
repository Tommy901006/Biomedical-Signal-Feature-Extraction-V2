import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd

# Correct ECG 12-lead names
ecg_leads = ['I', 'II', 'III', 'aVR', 'aVL', 'aVF', 'V1', 'V2', 'V3', 'V4', 'V5', 'V6']

def process_folder(folder_path):
    output_folder = os.path.join(folder_path, "labeled_output")
    os.makedirs(output_folder, exist_ok=True)
    processed_files = 0
    skipped_files = 0

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".csv"):
            file_path = os.path.join(folder_path, filename)
            try:
                df = pd.read_csv(file_path, header=None)
                if df.shape[1] == 12:
                    df.columns = ecg_leads
                    output_path = os.path.join(output_folder, filename.replace(".csv", "_labeled.csv"))
                    df.to_csv(output_path, index=False)
                    processed_files += 1
                else:
                    skipped_files += 1
            except Exception as e:
                skipped_files += 1

    return processed_files, skipped_files, output_folder

def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        processed, skipped, output_path = process_folder(folder_path)
        messagebox.showinfo("Batch ECG Labeling Complete",
                            f"✅ Processed: {processed} file(s)\n"
                            f"⚠️ Skipped: {skipped} file(s)\n\n"
                            f"Labeled files saved to:\n{output_path}")

# GUI setup
root = tk.Tk()
root.title("ECG 12-Lead CSV Labeler")
root.geometry("400x200")

label = tk.Label(root, text="Select a folder containing raw ECG CSV files:", font=("Arial", 12))
label.pack(pady=20)

button = tk.Button(root, text="Select Folder", command=select_folder, font=("Arial", 12), width=20)
button.pack(pady=10)

root.mainloop()
