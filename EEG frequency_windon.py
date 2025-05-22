import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from scipy.integrate import simpson
import matplotlib.pyplot as plt

class EEGAnalysisGUI:
    def __init__(self, master):
        self.master = master
        master.title("EEG Band Power Analysis")
        master.geometry("1000x850")

        self.selected_cols = []
        self.file_columns = []
        self.combo_cols = []
        self.frequency_bands = {
            'delta': (0.5, 4),
            'theta': (4, 8),
            'alpha': (8, 13),
            'beta': (14, 30),
            'gamma': (30, 100)
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
        row_idx = 0
        for band in self.frequency_bands:
            ttk.Label(frame_freq, text=f"{band.capitalize()} Low:").grid(row=row_idx, column=0, sticky='e')
            low_entry = ttk.Entry(frame_freq, width=7)
            low_entry.insert(0, str(self.frequency_bands[band][0]))
            low_entry.grid(row=row_idx, column=1)

            ttk.Label(frame_freq, text=f"{band.capitalize()} High:").grid(row=row_idx, column=2, sticky='e')
            high_entry = ttk.Entry(frame_freq, width=7)
            high_entry.insert(0, str(self.frequency_bands[band][1]))
            high_entry.grid(row=row_idx, column=3)

            self.band_entries[band] = (low_entry, high_entry)
            row_idx += 1

        frame_sampling = ttk.LabelFrame(self.master, text="Sampling & Sliding Window", padding=10)
        frame_sampling.pack(fill='x', padx=10, pady=5)

        ttk.Label(frame_sampling, text="Sampling Rate (Hz):").grid(row=0, column=0, sticky='e')
        self.entry_sampling_rate = ttk.Entry(frame_sampling, width=10)
        self.entry_sampling_rate.insert(0, "500")
        self.entry_sampling_rate.grid(row=0, column=1, padx=5)

        self.var_plot = tk.BooleanVar(value=False)
        self.var_sliding = tk.BooleanVar(value=False)

        ttk.Checkbutton(frame_sampling, text="Use Sliding Window (with Overlap)", variable=self.var_sliding).grid(row=0, column=2, padx=10)
        ttk.Checkbutton(frame_sampling, text="Draw FFT Spectrum", variable=self.var_plot).grid(row=0, column=3, padx=10)

        ttk.Label(frame_sampling, text="Window Size (sec):").grid(row=1, column=0, sticky='e', pady=5)
        self.entry_window_size = ttk.Entry(frame_sampling, width=10)
        self.entry_window_size.insert(0, "2")
        self.entry_window_size.grid(row=1, column=1)

        ttk.Label(frame_sampling, text="Overlap (%):").grid(row=1, column=2, sticky='e')
        self.entry_overlap = ttk.Entry(frame_sampling, width=10)
        self.entry_overlap.insert(0, "50")
        self.entry_overlap.grid(row=1, column=3)

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
        use_sliding = self.var_sliding.get()
        plot_required = self.var_plot.get()

        try:
            sampling_rate = int(self.entry_sampling_rate.get())
            window_size = float(self.entry_window_size.get())
            overlap_pct = float(self.entry_overlap.get())
        except ValueError:
            messagebox.showerror("Error", "Sampling rate, window size, and overlap must be numbers.")
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

                        channel_data = data[col].values
                        n = len(channel_data)

                        if use_sliding:
                            window_len = int(window_size * sampling_rate)
                            step = int(window_len * (1 - overlap_pct / 100))
                            if step <= 0:
                                messagebox.showerror("Error", "Overlap too high, step size is zero.")
                                return

                            power_list = {b: [] for b in bands}
                            rel_power_list = {b: [] for b in bands}
                            total_powers = []

                            for start in range(0, n - window_len + 1, step):
                                segment = channel_data[start:start + window_len]
                                fft_vals = np.fft.fft(segment)
                                freqs = np.fft.fftfreq(len(segment), d=1 / sampling_rate)
                                pos_mask = (freqs >= 0.5) & (freqs <= 100)
                                freqs = freqs[pos_mask]
                                powers = np.abs(fft_vals[pos_mask]) ** 2

                                total_power = simpson(powers, x=freqs)
                                total_powers.append(total_power)

                                for band, (low, high) in bands.items():
                                    band_mask = (freqs >= low) & (freqs <= high)
                                    band_power = simpson(powers[band_mask], x=freqs[band_mask])
                                    power_list[band].append(band_power)
                                    rel_power_list[band].append(band_power / total_power if total_power > 0 else 0)

                            result = {"File Name": file_name}
                            result["Total Power (0.5–100Hz)"] = np.mean(total_powers)
                            for band in bands:
                                result[f"{band.capitalize()} Band Power"] = np.mean(power_list[band])
                                result[f"{band.capitalize()} Band Relative Power"] = np.mean(rel_power_list[band])
                            results.append(result)

                        else:
                            # 整段處理邏輯
                            fft_values = np.fft.fft(channel_data)
                            freqs = np.fft.fftfreq(n, d=1 / sampling_rate)
                            pos_mask = (freqs >= 0.5) & (freqs <= 100)
                            freqs = freqs[pos_mask]
                            powers = np.abs(fft_values[pos_mask]) ** 2

                            total_power = simpson(powers, x=freqs)
                            result = {"File Name": file_name}
                            result["Total Power (0.5–100Hz)"] = total_power
                            for band, (low, high) in bands.items():
                                band_mask = (freqs >= low) & (freqs <= high)
                                band_power = simpson(powers[band_mask], x=freqs[band_mask])
                                rel_power = band_power / total_power if total_power > 0 else 0
                                result[f"{band.capitalize()} Band Power"] = band_power
                                result[f"{band.capitalize()} Band Relative Power"] = rel_power
                            results.append(result)

                        self.log_message(f"Processed {file_name} Column {col}")

                    except Exception as e:
                        self.log_message(f"Error processing {file_name}: {e}")

                if results:
                    df_results = pd.DataFrame(results)
                    df_results.to_excel(writer, sheet_name=self.safe_sheet_name(col), index=False)

        messagebox.showinfo("Completed", f"Results saved to {output_excel_path}")
        try:
            os.startfile(folder)
        except Exception as e:
            self.log_message(f"Failed to open folder: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = EEGAnalysisGUI(root)
    root.mainloop()
