import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pandas as pd
import numpy as np
from NLIDOOP3 import RecurrenceAnalysis

class NLIDApp:
    def __init__(self, master):
        self.master = master
        master.title("NLID 批次分析工具（支援參數輸入）")
        master.geometry("800x600")

        container = ttk.Frame(master, padding=10)
        container.pack(fill='both', expand=True)

        input_frame = ttk.Labelframe(container, text="Input Settings", padding=10)
        input_frame.pack(fill='x', pady=5)

        ttk.Label(input_frame, text="Folder:").grid(row=0, column=0, sticky='w')
        self.entry_folder = ttk.Entry(input_frame, width=60)
        self.entry_folder.grid(row=0, column=1, sticky='ew')
        ttk.Button(input_frame, text="Browse", command=self.browse_folder).grid(row=0, column=2)

        ttk.Button(input_frame, text="Load Columns", command=self.load_columns).grid(row=1, column=1, pady=5)

        input_frame.columnconfigure(1, weight=1)

        column_frame = ttk.Labelframe(container, text="Select 2 Columns", padding=10)
        column_frame.pack(fill='x', pady=5)
        self.combo_col_x = ttk.Combobox(column_frame, state="readonly", width=30)
        self.combo_col_y = ttk.Combobox(column_frame, state="readonly", width=30)
        ttk.Label(column_frame, text="Column X:").grid(row=0, column=0, sticky='e')
        self.combo_col_x.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        ttk.Label(column_frame, text="Column Y:").grid(row=1, column=0, sticky='e')
        self.combo_col_y.grid(row=1, column=1, sticky='w', padx=5, pady=2)

        param_frame = ttk.Labelframe(container, text="Parameters", padding=10)
        param_frame.pack(fill='x', pady=5)
        ttk.Label(param_frame, text="Embedding dimension (m):").grid(row=0, column=0, sticky='w')
        self.entry_m = ttk.Entry(param_frame, width=10)
        self.entry_m.insert(0, "3")
        self.entry_m.grid(row=0, column=1, sticky='w', padx=5)
        ttk.Label(param_frame, text="Delay (tau):").grid(row=1, column=0, sticky='w')
        self.entry_tau = ttk.Entry(param_frame, width=10)
        self.entry_tau.insert(0, "1")
        self.entry_tau.grid(row=1, column=1, sticky='w', padx=5)

        progress_frame = ttk.Frame(container, padding=0)
        progress_frame.pack(fill='both', expand=True, pady=5)

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill='x', pady=5)

        self.log = scrolledtext.ScrolledText(progress_frame, height=15, wrap='word')
        self.log.pack(fill='both', expand=True)

        ttk.Button(container, text="Start Analysis", command=self.start).pack(pady=10)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.entry_folder.delete(0, tk.END)
            self.entry_folder.insert(0, folder)
            self.combo_col_x['values'] = []
            self.combo_col_y['values'] = []
            self.combo_col_x.set('')
            self.combo_col_y.set('')

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
            cols = [''] + list(df.columns.str.strip())
            self.combo_col_x['values'] = cols
            self.combo_col_y['values'] = cols
            self.combo_col_x.set('')
            self.combo_col_y.set('')
            messagebox.showinfo("Columns Loaded", f"Loaded columns from {files[0]}")
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    def log_message(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.yview(tk.END)

    def start(self):
        folder = self.entry_folder.get()
        col_x = self.combo_col_x.get()
        col_y = self.combo_col_y.get()
        try:
            m = int(self.entry_m.get())
            tau = int(self.entry_tau.get())
        except ValueError:
            messagebox.showerror("Invalid input", "m and tau must be integers.")
            return
        if not os.path.isdir(folder) or not col_x or not col_y:
            messagebox.showerror("Missing info", "Ensure folder and two columns are selected.")
            return
        threading.Thread(target=self.process_files, args=(folder, col_x, col_y, m, tau), daemon=True).start()

    def process_files(self, folder, col_x, col_y, m, tau):
        files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith(('.xlsx', '.csv'))]
        self.progress['maximum'] = len(files)
        self.progress['value'] = 0
        results = []

        for file in files:
            basename = os.path.basename(file)
            try:
                df = pd.read_excel(file) if file.endswith(('.xls', '.xlsx')) else pd.read_csv(file)
                df.columns = df.columns.str.strip().str.upper()
                col_x = col_x.strip().upper()
                col_y = col_y.strip().upper()

                if col_x not in df.columns or col_y not in df.columns:
                    self.log_message(f"{basename}: missing selected columns.")
                    continue

                x = df[col_x].dropna().values
                y = df[col_y].dropna().values
                min_len = min(len(x), len(y))
                if min_len < 1:
                    self.log_message(f"{basename}: not enough data.")
                    continue

                x = x[:min_len]
                y = y[:min_len]

                ra_x = RecurrenceAnalysis(x, m=m, tau=tau)
                ps_x = ra_x.reconstruct_phase_space()

                ra_y = RecurrenceAnalysis(y, m=m, tau=tau)
                ps_y = ra_y.reconstruct_phase_space()

                AR_X = RecurrenceAnalysis.compute_reconstruction_matrix(ps_x, threshold=0.1, threshold_type="dynamic")
                AR_Y = RecurrenceAnalysis.compute_reconstruction_matrix(ps_y, threshold=0.1, threshold_type="dynamic")

                NLID_XY, NLID_YX = RecurrenceAnalysis.calculate_nlid(AR_X, AR_Y)

                results.append({
                    "檔名": basename,
                    f"NLID({col_x}|{col_y})": NLID_XY,
                    f"NLID({col_y}|{col_x})": NLID_YX
                })
                self.log_message(f"Processed: {basename}")
            except Exception as e:
                self.log_message(f"Error {basename}: {e}")
            self.progress['value'] += 1

        if results:
            result_df = pd.DataFrame(results)
            output_path = os.path.join(folder, "NLID_Results.xlsx")
            result_df.to_excel(output_path, index=False)
            self.log_message(f"Results saved to {output_path}")
            messagebox.showinfo("Done", f"Analysis completed. Saved to: {output_path}")
        else:
            messagebox.showwarning("No Data", "No valid files processed.")

if __name__ == "__main__":
    root = tk.Tk()
    app = NLIDApp(root)
    root.mainloop()