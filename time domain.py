import os
import threading
import pandas as pd
import numpy as np
from scipy.stats import skew
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk

class StatisticsApp:
    MAX_COLS = 5

    def __init__(self, master):
        self.master = master
        master.title("Statistics Calculator")
        master.geometry("850x750")

        container = ttk.Frame(master, padding=10)
        container.pack(fill='both', expand=True)

        input_frame = ttk.Labelframe(container, text="Input/Output Settings", padding=10)
        input_frame.pack(fill='x', pady=5)

        ttk.Label(input_frame, text="Folder:").grid(row=0, column=0, sticky='w')
        self.entry_folder = ttk.Entry(input_frame, width=60)
        self.entry_folder.grid(row=0, column=1, sticky='ew')
        ttk.Button(input_frame, text="Browse", command=self.browse_folder).grid(row=0, column=2)

        ttk.Button(input_frame, text="Load Columns", command=self.load_columns).grid(row=1, column=1, pady=5)

        ttk.Label(input_frame, text="Output File:").grid(row=2, column=0, sticky='w')
        self.entry_output = ttk.Entry(input_frame, width=60)
        self.entry_output.grid(row=2, column=1, sticky='ew')
        ttk.Button(input_frame, text="Browse", command=self.browse_output).grid(row=2, column=2)

        input_frame.columnconfigure(1, weight=1)

        column_frame = ttk.Labelframe(container, text="Select Columns (up to 5)", padding=10)
        column_frame.pack(fill='x', pady=5)
        self.combo_cols = []
        for i in range(self.MAX_COLS):
            ttk.Label(column_frame, text=f"Column {i+1}:").grid(row=i, column=0, sticky='e')
            combo = ttk.Combobox(column_frame, state="readonly", width=30)
            combo.grid(row=i, column=1, sticky='w', padx=5, pady=2)
            self.combo_cols.append(combo)
        column_frame.columnconfigure(1, weight=1)

        progress_frame = ttk.Frame(container, padding=0)
        progress_frame.pack(fill='both', expand=True, pady=5)

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill='x', pady=5)

        self.log = scrolledtext.ScrolledText(progress_frame, height=15, wrap='word')
        self.log.pack(fill='both', expand=True)

        ttk.Button(container, text="Start Calculation", command=self.start).pack(pady=10)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.entry_folder.delete(0, tk.END)
            self.entry_folder.insert(0, folder)
            for combo in self.combo_cols:
                combo['values'] = []
                combo.set('')

    def load_columns(self):
        folder = self.entry_folder.get()
        if not os.path.isdir(folder):
            messagebox.showerror("Invalid folder", "Please select a valid folder first.")
            return
        files = [f for f in os.listdir(folder) if f.lower().endswith((".xls", ".xlsx", ".csv"))]
        if not files:
            messagebox.showwarning("No files found", "No Excel/CSV files in the folder.")
            return
        try:
            path = os.path.join(folder, files[0])
            df = pd.read_excel(path) if path.endswith(('.xls', '.xlsx')) else pd.read_csv(path)
            cols = [''] + list(df.columns)
            for combo in self.combo_cols:
                combo['values'] = cols
                combo.set('')
            messagebox.showinfo("Columns Loaded", f"Loaded columns from {files[0]}")
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    def browse_output(self):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, file)

    def log_message(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.yview(tk.END)

    def start(self):
        folder = self.entry_folder.get()
        output = self.entry_output.get()
        cols = [c.get() for c in self.combo_cols if c.get()]
        if not os.path.isdir(folder) or not cols or not output:
            messagebox.showerror("Missing info", "Ensure folder, columns, and output are set.")
            return
        threading.Thread(target=self.process_files, args=(folder, output, cols), daemon=True).start()

    def compute_statistics(self, df, column):
        mean = df[column].mean()
        std = df[column].std()
        skewness = skew(df[column].dropna())
        return mean, std, skewness

    def process_files(self, folder, output, cols):
        files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith(('.xlsx', '.csv'))]
        self.progress['maximum'] = len(files)
        self.progress['value'] = 0
        all_results = []

        for file in files:
            basename = os.path.basename(file)
            try:
                df = pd.read_excel(file) if file.endswith(('.xls', '.xlsx')) else pd.read_csv(file)
                row = {'File Name': basename}
                logs = []
                for col in cols:
                    if col in df.columns:
                        mean, std, skewness = self.compute_statistics(df, col)
                        row[f"{col} Mean"] = mean
                        row[f"{col} Std"] = std
                        row[f"{col} Skewness"] = skewness
                        logs.append(f"{col} computed")
                    else:
                        logs.append(f"{col} skipped")
                all_results.append(row)
                self.log_message(f"{basename}: " + ", ".join(logs))
            except Exception as e:
                self.log_message(f"Error {basename}: {e}")
            self.progress['value'] += 1

        if all_results:
            pd.DataFrame(all_results).to_excel(output, index=False)
            self.log_message(f"Results saved to {output}")
            messagebox.showinfo("Completed", "Calculation finished.")
        else:
            messagebox.showwarning("No Data", "No valid files processed.")

if __name__ == "__main__":
    root = tk.Tk()
    app = StatisticsApp(root)
    root.mainloop()
