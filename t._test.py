import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from scipy import stats
from tkinter import font
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os

class TTestGUI:
    def __init__(self, master):
        self.master = master
        master.title("📊 T 檢定互動式工具")
        master.geometry("700x820")
        master.resizable(False, False)
        master.configure(bg='#f9f9f9')

        # Style configuration
        style = ttk.Style()
        try:
            style.theme_use('vista')
        except:
            style.theme_use('clam')
        style.configure('.', font=('Microsoft JhengHei', 11), background='#f9f9f9')
        style.configure('TButton', padding=8)
        style.configure('TLabel', background='#f9f9f9')
        style.configure('TCombobox', padding=5)
        style.configure('TLabelframe', background='#f0f8ff', borderwidth=2, relief='groove')
        style.configure('TLabelframe.Label', font=('Microsoft JhengHei', 12, 'bold'))

        # Title
        title_font = font.Font(family='Microsoft JhengHei', size=18, weight='bold')
        lbl_title = ttk.Label(master, text="T 檢定互動式工具", font=title_font, foreground='#333')
        lbl_title.pack(pady=(15, 10))

        # Notebook for pages
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill='both', expand=True, padx=20, pady=10)

        # ---- 設定分頁 ----
        self.page1 = ttk.Frame(self.notebook)
        self.notebook.add(self.page1, text='🔧 設定')

        # 檔案選擇
        frame_file = ttk.Labelframe(self.page1, text="📂 選擇檔案")
        frame_file.pack(fill='x', pady=5)
        ttk.Button(frame_file, text="開啟 Excel/CSV", command=self.load_file).pack(side='left', padx=10, pady=10)
        self.lbl_status = ttk.Label(frame_file, text="尚未載入檔案", foreground='#555')
        self.lbl_status.pack(side='left', padx=10)

        # 變數選擇
        frame_vars = ttk.Labelframe(self.page1, text="📊 變數選擇")
        frame_vars.pack(fill='x', pady=5)
        ttk.Label(frame_vars, text="變數 A：").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.combo1 = ttk.Combobox(frame_vars, state='readonly', width=30)
        self.combo1.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(frame_vars, text="變數 B：").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.combo2 = ttk.Combobox(frame_vars, state='readonly', width=30)
        self.combo2.grid(row=1, column=1, padx=5, pady=5)

        # 檢定選項
        frame_opts = ttk.Labelframe(self.page1, text="⚙️ 檢定選項")
        frame_opts.pack(fill='x', pady=5)
        ttk.Label(frame_opts, text="檢定類型：").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.test_type = ttk.Combobox(frame_opts, state='readonly', width=32,
                                      values=['獨立樣本 (自動判斷變異)', '配對樣本'])
        self.test_type.current(0)
        self.test_type.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(frame_opts, text="尾端檢定：").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.tail_type = ttk.Combobox(frame_opts, state='readonly', width=32,
                                      values=['雙尾', '單尾'])
        self.tail_type.current(0)
        self.tail_type.grid(row=1, column=1, padx=5, pady=5)

        # 圖表設定
        frame_chart = ttk.Labelframe(self.page1, text="📋 圖表設定")
        frame_chart.pack(fill='x', pady=5)
        ttk.Label(frame_chart, text="Bar Color：").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.color_combo = ttk.Combobox(frame_chart, state='readonly', width=15,
                                        values=['blue','green','red','orange','purple','gray'])
        self.color_combo.current(0)
        self.color_combo.grid(row=0, column=1, padx=5, pady=5)
        self.show_labels_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_chart, text='顯示數值標籤', variable=self.show_labels_var).grid(row=0, column=2, padx=10)
        ttk.Label(frame_chart, text="Error capsize：").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.capsize_spin = ttk.Spinbox(frame_chart, from_=0, to=20, width=5)
        self.capsize_spin.set(10)
        self.capsize_spin.grid(row=1, column=1, padx=5, pady=5)

        # 執行按鈕
        ttk.Button(self.page1, text="▶️ 執行 T 檢定", command=self.run_ttest).pack(pady=10)

        # ---- 結果＆圖表分頁 ----
        self.page2 = ttk.Frame(self.notebook)
        self.notebook.add(self.page2, text='📈 結果＆圖表')
        frame_res = ttk.Labelframe(self.page2, text="🔍 檢定結果")
        frame_res.pack(fill='x', pady=5, padx=10)
        self.txt_result = tk.Text(frame_res, height=8, font=('Microsoft JhengHei', 11), bg='#fcfcfc')
        self.txt_result.pack(fill='x', padx=10, pady=5)
        self.txt_result.configure(state='disabled')
        self.chart_frame = ttk.Labelframe(self.page2, text="Mean and Std Dev Chart")
        self.chart_frame.pack(fill='both', expand=True, pady=5, padx=10)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[('Excel', '*.xls;*.xlsx'), ('CSV', '*.csv')])
        if not file_path:
            return
        try:
            self.df = pd.read_excel(file_path) if file_path.lower().endswith(('xls','xlsx')) else pd.read_csv(file_path)
            cols = list(self.df.columns)
            self.combo1['values'] = cols
            self.combo2['values'] = cols
            self.lbl_status.config(text=f"已載入: {file_path.split('/')[-1]}")
        except Exception as e:
            messagebox.showerror('錯誤', f'檔案讀取失敗:\n{e}')

    def run_ttest(self):
        # Pre-check
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showwarning('警告', '請先載入資料檔案！')
            return
        col1, col2 = self.combo1.get(), self.combo2.get()
        if not col1 or not col2:
            messagebox.showwarning('警告', '請選擇 A 與 B 兩組變數！')
            return

        # Compute stats
        data1, data2 = self.df[col1].dropna(), self.df[col2].dropna()
        if '獨立樣本' in self.test_type.get():
            lev_stat, lev_p = stats.levene(data1, data2)
            equal_var = lev_p > 0.05
            var_note = '等變異' if equal_var else '不等變異'
            t_stat, p_two = stats.ttest_ind(data1, data2, equal_var=equal_var)
        else:
            lev_p, var_note = None, None
            t_stat, p_two = stats.ttest_rel(data1, data2)
        mean1, mean2 = data1.mean(), data2.mean()
        std1, std2 = data1.std(ddof=0), data2.std(ddof=0)
        if self.tail_type.get() == '雙尾':
            p_val, tail_desc = p_two, '雙尾'
        else:
            dir_ = t_stat>0 if mean1>mean2 else t_stat<0
            tail_desc = '單尾 (A > B)' if mean1>mean2 else '單尾 (A < B)'
            p_val = p_two/2 if dir_ else 1-p_two/2

        # Display text
        self.txt_result.configure(state='normal')
        self.txt_result.delete('1.0', tk.END)
        self.txt_result.insert(tk.END, f"T 統計量: {t_stat:.4f}\n")
        self.txt_result.insert(tk.END, f"P 值 ({tail_desc}): {p_val:.4f}\n")
        if lev_p is not None:
            self.txt_result.insert(tk.END, f"Levene p: {lev_p:.4f} ({var_note})\n")
        self.txt_result.insert(tk.END, f"均值 A ({col1}): {mean1:.4f}\n")
        self.txt_result.insert(tk.END, f"母體標準差 A: {std1:.4f}\n")
        self.txt_result.insert(tk.END, f"均值 B ({col2}): {mean2:.4f}\n")
        self.txt_result.insert(tk.END, f"母體標準差 B: {std2:.4f}\n")
        self.txt_result.configure(state='disabled')
        self.notebook.select(self.page2)

        # Plot chart with custom options
        color = self.color_combo.get()
        capsize = int(self.capsize_spin.get())
        show_lbls = self.show_labels_var.get()
        labels = ['A','B']; means=[mean1,mean2]; stds=[std1,std2]
        fig, ax = plt.subplots(figsize=(4,3))
        bars = ax.bar(labels, means, yerr=stds, capsize=capsize, color=color)
        ax.set_ylabel('Value')
        ax.set_title('Mean (error bars = STDEV.P)')

        # value labels
        if show_lbls:
            for bar,m in zip(bars, means):
                ax.text(bar.get_x()+bar.get_width()/2, m, f'{m:.2f}', ha='center', va='bottom')

        # significance
        sig = '***' if p_val<0.001 else '**' if p_val<0.01 else '*' if p_val<0.05 else ''
        if sig:
            y = max(means)+max(stds)*0.5
            ax.text(0.5, y+max(stds)*0.1, sig, ha='center', va='bottom', fontsize=16)

        # embed
        for w in self.chart_frame.winfo_children(): w.destroy()
        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw(); canvas.get_tk_widget().pack(fill='both', expand=True)

        # Save chart?
        if messagebox.askyesno('詢問','是否另存圖表為圖片？'):
            img_path = filedialog.asksaveasfilename(defaultextension='.png', filetypes=[('PNG','*.png')])
            if img_path: fig.savefig(img_path); messagebox.showinfo('儲存成功',f'圖表已儲存至:\n{img_path}')

        # Save Excel
        save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')])
        if save_path:
            result = {'T 統計量':t_stat, 'P 值':p_val,'尾端':tail_desc,
                      'Levene p': lev_p or '', '變異判定':var_note or '',
                      f'均值 A ({col1})':mean1,'母體標準差 A':std1,
                      f'均值 B ({col2})':mean2,'母體標準差 B':std2}
            pd.DataFrame([result]).to_excel(save_path,index=False)
            messagebox.showinfo('儲存成功',f'結果已儲存至:\n{save_path}')
            try: os.startfile(os.path.dirname(save_path))
            except: pass

if __name__=='__main__':
    root=tk.Tk(); app=TTestGUI(root); root.mainloop()
