import os
import threading
import pandas as pd
import numpy as np
from scipy.stats import skew
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
from tkinter import ttk
import matplotlib.pyplot as plt
import smtplib
from email.message import EmailMessage

class StatisticsApp:
    MAX_COLS = 5

    def __init__(self, master):
        self.master = master
        master.title("Statistics Calculator with Window Control")
        master.geometry("900x800")

        self.recipient_email = None

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

        option_frame = ttk.Labelframe(container, text="Options", padding=10)
        option_frame.pack(fill='x', pady=5)

        self.var_window = tk.BooleanVar(value=False)
        self.var_plot = tk.BooleanVar(value=False)
        self.var_email = tk.BooleanVar(value=False)

        ttk.Checkbutton(option_frame, text="Enable Sliding Window", variable=self.var_window).grid(row=0, column=0, sticky='w')
        ttk.Checkbutton(option_frame, text="Plot Segment Trends", variable=self.var_plot).grid(row=0, column=1, sticky='w')
        ttk.Checkbutton(option_frame, text="Email Result", variable=self.var_email).grid(row=0, column=2, sticky='w')

        ttk.Label(option_frame, text="Window Size:").grid(row=1, column=0, sticky='e')
        self.entry_window = ttk.Entry(option_frame, width=10)
        self.entry_window.insert(0, "100")
        self.entry_window.grid(row=1, column=1, sticky='w')

        ttk.Label(option_frame, text="Overlap (%):").grid(row=1, column=2, sticky='e')
        self.entry_overlap = ttk.Entry(option_frame, width=10)
        self.entry_overlap.insert(0, "50")
        self.entry_overlap.grid(row=1, column=3, sticky='w')

        self.output_mode = tk.StringVar(value="average")
        ttk.Label(option_frame, text="Output Mode:").grid(row=2, column=0, sticky='e')
        ttk.Radiobutton(option_frame, text="Average Only", variable=self.output_mode, value="average").grid(row=2, column=1, sticky='w')
        ttk.Radiobutton(option_frame, text="Per Segment", variable=self.output_mode, value="segment").grid(row=2, column=2, sticky='w')

        progress_frame = ttk.Frame(container)
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
        files = [f for f in os.listdir(folder) if f.endswith((".xls", ".xlsx", ".csv"))]
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

    def compute_statistics(self, series):
        return series.mean(), series.std(), skew(series.dropna())

    def start(self):
        folder = self.entry_folder.get()
        output = self.entry_output.get()
        cols = [c.get() for c in self.combo_cols if c.get()]
        if not os.path.isdir(folder) or not cols or not output:
            messagebox.showerror("Missing info", "Ensure folder, columns, and output are set.")
            return
        if self.var_email.get():
            self.recipient_email = simpledialog.askstring("Email", "Enter recipient email:")
        threading.Thread(target=self.process_files, args=(folder, output, cols), daemon=True).start()

    def process_files(self, folder, output, cols):
        files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith(('.xlsx', '.csv'))]
        self.progress['maximum'] = len(files)
        self.progress['value'] = 0

        use_window = self.var_window.get()
        only_mean = self.output_mode.get() == "average"
        full_stats = self.output_mode.get() == "segment"
        plot_segment = self.var_plot.get()
        send_email = self.var_email.get()

        window_size = int(self.entry_window.get())
        overlap = float(self.entry_overlap.get()) / 100

        summary_results = []
        segment_results = {}

        for file in files:
            basename = os.path.basename(file)
            try:
                df = pd.read_excel(file) if file.endswith(('.xls', '.xlsx')) else pd.read_csv(file)
                row = {"File": basename}
                seg_rows = []

                for col in cols:
                    if col not in df.columns:
                        continue
                    data = df[col].dropna().reset_index(drop=True)

                    if use_window:
                        length = len(data)
                        step = int(window_size * (1 - overlap))
                        all_means, all_stds, all_skews = [], [], []

                        for i in range((length - window_size) // step + 1):
                            start, end = i * step, i * step + window_size
                            segment = data[start:end]
                            mean, std, skw = self.compute_statistics(segment)
                            all_means.append(mean)
                            all_stds.append(std)
                            all_skews.append(skw)

                            if full_stats:
                                seg_rows.append({
                                    "File": basename,
                                    "Column": col,
                                    "Segment": f"Segment{i+1}",
                                    "Mean": mean,
                                    "Std": std,
                                    "Skewness": skw
                                })

                        if only_mean:
                            row[f"{col} Mean"] = np.mean(all_means)
                            row[f"{col} Std"] = np.mean(all_stds)
                            row[f"{col} Skewness"] = np.mean(all_skews)

                        if plot_segment and all_means:
                            plt.figure()
                            plt.plot(all_means, marker='o')
                            plt.title(f"{basename} - {col} Mean (Sliding)")
                            plt.xlabel("Segment")
                            plt.ylabel("Mean")
                            plt.grid()
                            plt.tight_layout()
                            plt.show()

                    else:
                        mean, std, skw = self.compute_statistics(data)
                        row[f"{col} Mean"] = mean
                        row[f"{col} Std"] = std
                        row[f"{col} Skewness"] = skw

                if use_window and full_stats:
                    combined_seg_dict = {}
                    for seg in seg_rows:
                        seg_id = seg["Segment"]
                        if seg_id not in combined_seg_dict:
                            combined_seg_dict[seg_id] = {"File": basename, "Segment": seg_id}
                        col = seg["Column"]
                        combined_seg_dict[seg_id][f"Column_{col}_Mean"] = seg["Mean"]
                        combined_seg_dict[seg_id][f"Column_{col}_Std"] = seg["Std"]
                        combined_seg_dict[seg_id][f"Column_{col}_Skewness"] = seg["Skewness"]
                    segment_results[basename] = list(combined_seg_dict.values())
                else:
                    summary_results.append(row)

            except Exception as e:
                self.log_message(f"Error {basename}: {e}")
            self.progress['value'] += 1

        if summary_results:
            pd.DataFrame(summary_results).to_excel(output, index=False)

        if segment_results:
            seg_path = os.path.join(os.path.dirname(output), "PerSegment_Output.xlsx")
            with pd.ExcelWriter(seg_path, engine='openpyxl') as writer:
                for fname, rows in segment_results.items():
                    pd.DataFrame(rows).to_excel(writer, sheet_name=fname[:31], index=False)

        if send_email and self.recipient_email:
            attachments = [output]
            if segment_results:
                attachments.append(seg_path)
            self.send_email(self.recipient_email, attachments)

        messagebox.showinfo("Done", "All files processed.")

    def send_email(self, to_email, attachments):
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = "4A930099@stust.edu.tw"
        sender_password = "trwvoyligttxqcjy"
        msg = EmailMessage()
        msg['Subject'] = "Statistics Report"
        msg['From'] = sender_email
        msg['To'] = to_email
        msg.set_content("Attached are your statistical analysis results.")
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            for f in attachments:
                with open(f, 'rb') as fp:
                    msg.add_attachment(fp.read(), maintype='application', subtype='octet-stream', filename=os.path.basename(f))
            server.send_message(msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = StatisticsApp(root)
    root.mainloop()
