import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class FileConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSV ↔ Excel Batch Converter")
        self.geometry("600x400")
        self.configure(bg="#f4f4f4")

        self.input_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_format = tk.StringVar(value="both")

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="Input Folder:", bg="#f4f4f4", font=("Arial", 11)).pack(pady=(10, 0))
        frame_input = tk.Frame(self, bg="#f4f4f4")
        frame_input.pack(pady=5)
        tk.Entry(frame_input, textvariable=self.input_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_input, text="Browse", command=self.browse_input, bg="#4CAF50", fg="white").pack(side=tk.LEFT)

        tk.Label(self, text="Output Folder:", bg="#f4f4f4", font=("Arial", 11)).pack(pady=(10, 0))
        frame_output = tk.Frame(self, bg="#f4f4f4")
        frame_output.pack(pady=5)
        tk.Entry(frame_output, textvariable=self.output_folder, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_output, text="Browse", command=self.browse_output, bg="#4CAF50", fg="white").pack(side=tk.LEFT)

        tk.Label(self, text="Output Format:", bg="#f4f4f4", font=("Arial", 11)).pack(pady=(10, 0))
        formats_frame = tk.Frame(self, bg="#f4f4f4")
        formats_frame.pack(pady=5)
        tk.Radiobutton(formats_frame, text="CSV", variable=self.output_format, value="csv", bg="#f4f4f4").pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(formats_frame, text="Excel", variable=self.output_format, value="excel", bg="#f4f4f4").pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(formats_frame, text="Both", variable=self.output_format, value="both", bg="#f4f4f4").pack(side=tk.LEFT, padx=10)

        tk.Button(self, text="Start Conversion", command=self.convert_files, bg="#2196F3", fg="white", font=("Arial", 12)).pack(pady=20)

        self.progress = ttk.Progressbar(self, orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=(10, 5))

        self.status = tk.Label(self, text="", bg="#f4f4f4", font=("Arial", 10))
        self.status.pack()

    def browse_input(self):
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            self.input_path.set(folder)

    def browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder.set(folder)

    def convert_files(self):
        input_path = self.input_path.get()
        output_folder = self.output_folder.get()
        fmt = self.output_format.get()

        if not input_path or not output_folder:
            messagebox.showerror("Error", "Please select both input and output folders.")
            return

        file_list = []
        for file in os.listdir(input_path):
            if file.endswith(".csv") or file.endswith(".xlsx"):
                file_list.append(os.path.join(input_path, file))

        total_files = len(file_list)
        self.progress["maximum"] = total_files
        self.progress["value"] = 0

        for i, file_path in enumerate(file_list, start=1):
            try:
                df = pd.read_csv(file_path) if file_path.endswith(".csv") else pd.read_excel(file_path)
                base_name = os.path.splitext(os.path.basename(file_path))[0]

                if fmt in ("excel", "both"):
                    df.to_excel(os.path.join(output_folder, base_name + ".xlsx"), index=False)
                if fmt in ("csv", "both"):
                    df.to_csv(os.path.join(output_folder, base_name + ".csv"), index=False)

                self.status.config(text=f"Converted: {base_name}")
            except Exception as e:
                self.status.config(text=f"Error: {e}")

            self.progress["value"] = i
            self.update_idletasks()

        messagebox.showinfo("Done", "All files have been converted.")

if __name__ == "__main__":
    app = FileConverterApp()
    app.mainloop()
