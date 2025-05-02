import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from scipy import signal
import matplotlib.pyplot as plt
import smtplib
from email.message import EmailMessage

class CoherenceAnalysisGUI:
    def __init__(self, master):
        self.master = master
        master.title("Batch Coherence Calculator")
        master.geometry("800x620")

        self.selected_cols = []
        self.file_columns = []
        self.combo_cols = []

        self.build_interface()

    def build_interface(self):
        frame_path = ttk.LabelFrame(self.master, text="Folder and Column Selection", padding=10)
        frame_path.pack(fill='x', padx=10, pady=5)

        ttk.Button(frame_path, text="Select Folder", command=self.select_folder).pack(anchor='w')
        self.lbl_folder = ttk.Label(frame_path, text="")
        self.lbl_folder.pack(anchor='w', pady=5)

        frame_cols = ttk.Frame(frame_path)
        frame_cols.pack()
        for i in range(2):
            ttk.Label(frame_cols, text=f"Column {i+1}:").grid(row=i, column=0, sticky='e')
            cb = ttk.Combobox(frame_cols, width=30, state="readonly")
            cb.grid(row=i, column=1, padx=5, pady=2)
            self.combo_cols.append(cb)

        frame_sampling = ttk.LabelFrame(self.master, text="Settings", padding=10)
        frame_sampling.pack(fill='x', padx=10, pady=5)

        ttk.Label(frame_sampling, text="Sampling Rate (Hz):").pack(side='left')
        self.entry_fs = ttk.Entry(frame_sampling, width=10)
        self.entry_fs.insert(0, "1000")
        self.entry_fs.pack(side='left', padx=5)

        self.var_save_curve = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_sampling, text="Save Coherence Curves", variable=self.var_save_curve).pack(side='left', padx=10)

        self.var_plot_single = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_sampling, text="Plot Single File Charts", variable=self.var_plot_single).pack(side='left', padx=10)

        self.var_plot_all = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_sampling, text="Plot All Files Comparison", variable=self.var_plot_all).pack(side='left', padx=10)

        ttk.Label(frame_sampling, text="Send results to email:").pack(side='left', padx=(10, 0))
        self.entry_email = ttk.Entry(frame_sampling, width=25)
        self.entry_email.pack(side='left', padx=5)

        frame_log = ttk.LabelFrame(self.master, text="Execution Log", padding=10)
        frame_log.pack(fill='both', expand=True, padx=10, pady=5)

        self.log = scrolledtext.ScrolledText(frame_log, height=12)
        self.log.pack(fill='both', expand=True)

        ttk.Button(self.master, text="Start Batch Processing", command=self.start_processing).pack(pady=10)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.lbl_folder.config(text=folder)
            files = [f for f in os.listdir(folder) if f.endswith(('.xlsx', '.csv'))]
            if files:
                path = os.path.join(folder, files[0])
                df = pd.read_excel(path) if path.endswith(('xlsx', 'xls')) else pd.read_csv(path)
                self.file_columns = df.columns.tolist()
                for cb in self.combo_cols:
                    cb['values'] = [''] + self.file_columns
                    cb.set('')
            else:
                messagebox.showwarning("No Files", "No Excel or CSV files found in the folder.")

    def log_message(self, msg):
        self.log.insert(tk.END, msg + '\n')
        self.log.yview(tk.END)

    def calculate_coherence(self, X, Y, fs=1000, nperseg=None):
        f, Cxy = signal.coherence(X, Y, fs=fs, nperseg=nperseg)
        return f, Cxy, np.mean(Cxy)

    def send_email(self, to_email, subject, body, attachment_path):
        try:
            smtp_server = "smtp.gmail.com"
            smtp_port = 587
            sender_email = "4A930099@stust.edu.tw"        # <--- 替換成你的 Gmail
            sender_password = "trwvoyligttxqcjy"         # <--- 替換成你的 App 密碼（不是 Gmail 登入密碼）

            msg = EmailMessage()
            msg['Subject'] = subject
            msg['From'] = sender_email
            msg['To'] = to_email
            msg.set_content(body)

            with open(attachment_path, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(attachment_path)
                msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)

            self.log_message(f"Email sent to {to_email}")
        except Exception as e:
            self.log_message(f"Failed to send email: {e}")

    def start_processing(self):
        folder = self.lbl_folder.cget("text")
        selected_cols = [cb.get() for cb in self.combo_cols if cb.get()]

        try:
            fs = int(self.entry_fs.get())
        except ValueError:
            messagebox.showerror("Error", "Sampling rate must be a number.")
            return

        if not folder or len(selected_cols) != 2:
            messagebox.showerror("Missing Information", "Please select a folder and exactly two columns.")
            return

        save_curves = self.var_save_curve.get()
        plot_single = self.var_plot_single.get()
        plot_all = self.var_plot_all.get()

        results = []
        curves = {}
        files = [f for f in os.listdir(folder) if f.endswith(('.xlsx', '.csv'))]
        for file in files:
            try:
                path = os.path.join(folder, file)
                df = pd.read_excel(path) if path.endswith(('xlsx', 'xls')) else pd.read_csv(path)

                col_x = df[selected_cols[0]].dropna()
                col_y = df[selected_cols[1]].dropna()

                min_len = min(len(col_x), len(col_y))
                col_x = col_x.iloc[:min_len]
                col_y = col_y.iloc[:min_len]

                f, Cxy, coh_value = self.calculate_coherence(col_x, col_y, fs=fs)

                results.append({
                    "File Name": file,
                    "Coherence": coh_value
                })

                curves[file] = (f, Cxy)

                self.log_message(f"Processed {file}: Coherence = {coh_value:.4f}")

                if plot_single:
                    plt.figure(figsize=(10, 6))
                    plt.plot(f, Cxy)
                    plt.title(f"Coherence - {file}")
                    plt.xlabel("Frequency (Hz)")
                    plt.ylabel("Coherence")
                    plt.grid()
                    plt.tight_layout()
                    single_plot_path = os.path.join(folder, f"{os.path.splitext(file)[0]}_Coherence.png")
                    plt.savefig(single_plot_path)
                    plt.close()

            except Exception as e:
                self.log_message(f"Error processing {file}: {e}")

        if results:
            out_summary = os.path.join(folder, "Coherence_Summary.xlsx")
            df_results = pd.DataFrame(results)
            df_results.to_excel(out_summary, index=False)

        if save_curves and curves:
            out_curves = os.path.join(folder, "Coherence_Curves.xlsx")
            with pd.ExcelWriter(out_curves) as writer:
                for file, (freqs, coh) in curves.items():
                    df_curve = pd.DataFrame({"Frequency (Hz)": freqs, "Coherence": coh})
                    safe_sheet = file[:30].replace('/', '_').replace('\\', '_').replace('?', '_')
                    df_curve.to_excel(writer, sheet_name=safe_sheet, index=False)

        if plot_all and curves:
            plt.figure(figsize=(12, 8))
            for file, (freqs, coh) in curves.items():
                plt.plot(freqs, coh, label=file)
            plt.title("Coherence Curves Comparison")
            plt.xlabel("Frequency (Hz)")
            plt.ylabel("Coherence")
            plt.grid()
            plt.legend()
            plt.tight_layout()
            plot_path = os.path.join(folder, "Coherence_Comparison.png")
            plt.savefig(plot_path)
            plt.close()

        to_email = self.entry_email.get().strip()
        if to_email and results:
            self.send_email(
                to_email,
                subject="Your Coherence Analysis Results",
                body="Attached is the summary of your batch coherence analysis.",
                attachment_path=out_summary
            )

        messagebox.showinfo("Completed", "All tasks completed!")

        try:
            os.startfile(folder)
        except Exception as e:
            self.log_message(f"Failed to open folder: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CoherenceAnalysisGUI(root)
    root.mainloop()
