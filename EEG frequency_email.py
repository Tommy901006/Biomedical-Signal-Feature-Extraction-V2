import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
from scipy.signal import ellip, filtfilt
from scipy.integrate import simpson
import matplotlib.pyplot as plt
import smtplib
from email.message import EmailMessage

class EEGAnalysisGUI:
    def __init__(self, master):
        self.master = master
        master.title("EEG Band Power Analysis")
        master.geometry("1000x800")

        # 啟動時詢問 Email（可跳過）
        self.recipient_email = simpledialog.askstring("收件人 Email", "請輸入收件者 Email（可跳過）:")

        self.selected_cols = []
        self.file_columns = []
        self.combo_cols = []
        self.frequency_bands = {
            'delta': (0.5, 4),
            'theta': (4, 8),
            'alpha': (8, 13),
            'beta': (13, 25),
            'gamma': (25, 45)
        }

        self.build_interface()

    def build_interface(self):
        frame_folder = ttk.LabelFrame(self.master, text="Folder and Column Selection", padding=10)
        frame_folder.pack(fill='x', padx=10, pady=5)

        ttk.Button(frame_folder, text="Select Folder", command=self.select_folder).pack(anchor='w')
        self.lbl_folder = ttk.Label(frame_folder, text="")
        self.lbl_folder.pack(anchor='w', pady=5)

        frame_cols = ttk.Frame(frame_folder)
        frame_cols.pack()
        for i in range(5):
            ttk.Label(frame_cols, text=f"Column {i+1}:").grid(row=i, column=0, sticky='e')
            cb = ttk.Combobox(frame_cols, width=30, state="readonly")
            cb.grid(row=i, column=1, padx=5, pady=2)
            self.combo_cols.append(cb)

        frame_freq = ttk.LabelFrame(self.master, text="Frequency Band Settings", padding=10)
        frame_freq.pack(fill='x', padx=10, pady=5)

        self.band_entries = {}
        for idx, band in enumerate(self.frequency_bands):
            ttk.Label(frame_freq, text=f"{band.capitalize()} Low:").grid(row=idx, column=0, sticky='e')
            low_entry = ttk.Entry(frame_freq, width=7)
            low_entry.insert(0, str(self.frequency_bands[band][0]))
            low_entry.grid(row=idx, column=1)

            ttk.Label(frame_freq, text=f"{band.capitalize()} High:").grid(row=idx, column=2, sticky='e')
            high_entry = ttk.Entry(frame_freq, width=7)
            high_entry.insert(0, str(self.frequency_bands[band][1]))
            high_entry.grid(row=idx, column=3)

            self.band_entries[band] = (low_entry, high_entry)

        frame_sampling = ttk.LabelFrame(self.master, text="Sampling Rate", padding=10)
        frame_sampling.pack(fill='x', padx=10, pady=5)
        ttk.Label(frame_sampling, text="Sampling Rate (Hz):").pack(side='left')
        self.entry_sampling_rate = ttk.Entry(frame_sampling, width=10)
        self.entry_sampling_rate.insert(0, "500")
        self.entry_sampling_rate.pack(side='left', padx=5)

        self.var_plot = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_sampling, text="Draw FFT Spectrum", variable=self.var_plot).pack(side='left', padx=10)

        frame_log = ttk.LabelFrame(self.master, text="Execution Log", padding=10)
        frame_log.pack(fill='both', expand=True, padx=10, pady=5)
        self.log = scrolledtext.ScrolledText(frame_log, height=12)
        self.log.pack(fill='both', expand=True)

        ttk.Button(self.master, text="Start Analysis", command=self.start_processing).pack(pady=10)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.lbl_folder.config(text=folder)
            files = [f for f in os.listdir(folder) if f.endswith('.csv')]
            if files:
                df = pd.read_csv(os.path.join(folder, files[0]))
                self.file_columns = df.columns.tolist()
                for cb in self.combo_cols:
                    cb['values'] = [''] + self.file_columns
                    cb.set('')
            else:
                messagebox.showwarning("No Files", "No CSV files found in the folder.")

    def log_message(self, msg):
        self.log.insert(tk.END, msg + '\n')
        self.log.yview(tk.END)

    def get_frequency_bands(self):
        bands = {}
        for band, (low_entry, high_entry) in self.band_entries.items():
            try:
                low = float(low_entry.get())
                high = float(high_entry.get())
                bands[band] = (low, high)
            except ValueError:
                messagebox.showerror("Frequency Error", f"Invalid frequency input for {band}.")
                return None
        return bands

    def safe_sheet_name(self, name):
        safe_name = ''.join(c if c not in '[]:*?/\\' else '_' for c in str(name))
        return safe_name[:31]

    def start_processing(self):
        folder = self.lbl_folder.cget("text")
        selected_cols = [cb.get() for cb in self.combo_cols if cb.get()]
        bands = self.get_frequency_bands()
        plot_required = self.var_plot.get()

        try:
            sampling_rate = int(self.entry_sampling_rate.get())
        except ValueError:
            messagebox.showerror("Error", "Sampling rate must be a number.")
            return

        if not folder or not selected_cols or bands is None:
            messagebox.showerror("Error", "Please check folder, columns and frequency bands.")
            return

        files = [f for f in os.listdir(folder) if f.endswith('.csv')]
        output_excel_path = os.path.join(folder, "EEG_Band_Analysis_Results.xlsx")

        with pd.ExcelWriter(output_excel_path) as writer:
            for col in selected_cols:
                results = []

                for file_name in files:
                    try:
                        file_path = os.path.join(folder, file_name)
                        data = pd.read_csv(file_path)
                        if col not in data.columns:
                            self.log_message(f"Skipped {file_name} (missing column {col})")
                            continue

                        channel_data = data[col]
                        n = len(channel_data)

                        fft_values_raw = np.fft.fft(channel_data)
                        frequencies_raw = np.fft.fftfreq(n, d=1/sampling_rate)
                        pos_freqs = frequencies_raw[:n // 2]
                        pos_fft_raw = np.abs(fft_values_raw[:n // 2]) ** 2

                        total_mask = (pos_freqs >= 0.5) & (pos_freqs <= 100)
                        total_power = simpson(pos_fft_raw[total_mask], x=pos_freqs[total_mask])

                        power_bands = {}
                        rel_power_bands = {}

                        for band, (lowcut, highcut) in bands.items():
                            nyquist = 0.5 * sampling_rate
                            low = lowcut / nyquist
                            high = highcut / nyquist
                            b, a = ellip(1, 0.5, 45, [low, high], btype='band')
                            filtered = filtfilt(b, a, channel_data)

                            fft_values = np.fft.fft(filtered)
                            frequencies = np.fft.fftfreq(n, d=1/sampling_rate)
                            positive = np.abs(fft_values[:n // 2]) ** 2
                            freqs = frequencies[:n // 2]

                            band_mask = (freqs >= lowcut) & (freqs <= highcut)
                            band_power = simpson(positive[band_mask], x=freqs[band_mask])
                            power_bands[f"{band.capitalize()} Band Power"] = band_power
                            rel_power_bands[f"{band.capitalize()} Band Relative Power"] = band_power / total_power

                        result = {"File Name": file_name}
                        result.update(power_bands)
                        result.update(rel_power_bands)
                        results.append(result)

                        if plot_required:
                            plt.figure(figsize=(12, 6))
                            plt.plot(pos_freqs[total_mask], pos_fft_raw[total_mask], label='Raw FFT Power', color='blue')
                            for i, (band, (low, high)) in enumerate(bands.items()):
                                mask = (pos_freqs >= low) & (pos_freqs <= high)
                                plt.fill_between(pos_freqs[mask], pos_fft_raw[mask], alpha=0.5, label=band)
                            plt.title(f'FFT - {file_name} - {col}')
                            plt.xlabel('Hz')
                            plt.ylabel('Power')
                            plt.legend()
                            plt.tight_layout()
                            plt.show()

                        self.log_message(f"Processed {file_name} Column {col}")

                    except Exception as e:
                        self.log_message(f"Error processing {file_name}: {e}")

                if results:
                    df_results = pd.DataFrame(results)
                    sheet_name = self.safe_sheet_name(col)
                    df_results.to_excel(writer, sheet_name=sheet_name, index=False)

        messagebox.showinfo("Completed", f"Results saved to {output_excel_path}")

        # ✅ 自動寄出 Email
        if self.recipient_email:
            self.send_email(self.recipient_email, output_excel_path)

        try:
            os.startfile(folder)
        except Exception as e:
            self.log_message(f"Failed to open folder: {e}")

    def send_email(self, to_email, file_path):
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = "4A930099@stust.edu.tw"  # 改為你的 Gmail
        sender_password = "trwvoyligttxqcjy"     # Gmail App 密碼

        try:
            msg = EmailMessage()
            msg['Subject'] = 'EEG 頻段能量分析結果'
            msg['From'] = sender_email
            msg['To'] = to_email
            msg.set_content("您好，附件為 EEG 頻段能量分析 Excel 結果。")

            with open(file_path, 'rb') as f:
                msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=os.path.basename(file_path))

            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)

            self.log_message("📧 Email 已寄出至 " + to_email)
        except Exception as e:
            self.log_message("❌ Email 發送失敗：" + str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = EEGAnalysisGUI(root)
    root.mainloop()
