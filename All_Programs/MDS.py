import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from sklearn.manifold import MDS
from sklearn.preprocessing import StandardScaler
import matplotlib.pyplot as plt
from scipy.spatial.distance import pdist, squareform
import io
from itertools import combinations
import threading

# (คลาส AboutWindow และ MDSConfigWindow ไม่มีการเปลี่ยนแปลง)
class AboutWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("คำแนะนำการตั้งค่าและการคำนวณ")
        self.geometry("900x800")
        self.transient(parent)
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(expand=True, fill=tk.BOTH)
        scrollbar = ttk.Scrollbar(main_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget = tk.Text(
            main_frame, wrap='word', yscrollcommand=scrollbar.set,
            padx=15, pady=15, font=("Segoe UI", 11),
            background=self.cget('bg'), relief=tk.FLAT, spacing1=5, spacing3=5
        )
        text_widget.pack(expand=True, fill=tk.BOTH)
        scrollbar.config(command=text_widget.yview)
        self._populate_text(text_widget)
        text_widget.config(state=tk.DISABLED)
        button_frame = ttk.Frame(self, padding=(0, 10, 0, 15))
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="ปิด", command=self.destroy).pack()
        self.grab_set()
        self.focus_set()

    def _populate_text(self, text_widget):
        text_widget.tag_configure("h1", font=("Segoe UI", 18, "bold"), spacing3=15)
        text_widget.tag_configure("h2", font=("Segoe UI", 14, "bold"), spacing3=10, lmargin1=10)
        text_widget.tag_configure("h3", font=("Segoe UI", 12, "bold"), spacing3=8, lmargin1=20)
        text_widget.tag_configure("bold", font=("Segoe UI", 11, "bold"))
        text_widget.tag_configure("section", lmargin1=20, lmargin2=20)
        text_widget.tag_configure("bullet", lmargin1=40, lmargin2=40)
        text_widget.insert(tk.END, "คำแนะนำการตั้งค่าและการคำนวณ\n", "h1")
        text_widget.insert(tk.END, "การตั้งค่าที่เหมาะสมกับลักษณะของข้อมูล จะทำให้ได้ผลลัพธ์การวิเคราะห์ที่มีคุณภาพสูงสุด\n\n")
        text_widget.insert(tk.END, "ส่วนที่ 1: การตั้งค่า Input (หน้าต่างหลัก)\n", "h2")
        text_widget.insert(tk.END, "1. Input: Distance Metric\n", "h3")
        text_widget.insert(tk.END, "คืออะไร: ", ("section", "bold"))
        text_widget.insert(tk.END, "สูตรทางคณิตศาสตร์ที่ใช้คำนวณ 'ความไม่เหมือนกัน' (Dissimilarity) ระหว่างคอลัมน์ข้อมูลที่คุณเลือก\n", "section")
        text_widget.insert(tk.END, "ควรเลือกเมื่อใด:\n", ("section", "bold"))
        text_widget.insert(tk.END, " • ", "bullet")
        text_widget.insert(tk.END, "Euclidean (ค่าเริ่มต้น): ", ("bullet", "bold"))
        text_widget.insert(tk.END, "เป็นค่ามาตรฐาน เหมาะสำหรับข้อมูลที่เป็นค่าต่อเนื่องทั่วไป (Ratio/Interval Scale) ที่ทุกคอลัมน์มีหน่วยและสเกลเดียวกัน\n", "bullet")
        text_widget.insert(tk.END, " • ", "bullet")
        text_widget.insert(tk.END, "Cityblock (Manhattan): ", ("bullet", "bold"))
        text_widget.insert(tk.END, "เป็นอีกทางเลือกของ Euclidean อาจทำงานได้ดีกว่าในบางกรณีที่ข้อมูลมีมิติสูง\n", "bullet")
        text_widget.insert(tk.END, " • ", "bullet")
        text_widget.insert(tk.END, "Correlation: ", ("bullet", "bold"))
        text_widget.insert(tk.END, "สำคัญมาก! ให้เลือกใช้เมื่อคุณสนใจ 'รูปแบบ' หรือ 'ทิศทาง' ของข้อมูล มากกว่า 'ขนาด' ของค่า เช่น ใช้เปรียบเทียบ Brand ที่มีภาพลักษณ์คล้ายกัน โดยไม่สนใจว่า Brand นั้นมีคะแนนสูงหรือต่ำกว่า Brand อื่นโดยรวม\n\n", "bullet")
        text_widget.insert(tk.END, "2. Input: Standardize Data (Z-score)\n", "h3")
        text_widget.insert(tk.END, "คืออะไร: ", ("section", "bold"))
        text_widget.insert(tk.END, "กระบวนการปรับสเกลข้อมูลในแต่ละคอลัมน์ ให้มีค่าเฉลี่ยเป็น 0 และส่วนเบี่ยงเบนมาตรฐานเป็น 1\n", "section")
        text_widget.insert(tk.END, "ควรเลือกเมื่อใด:\n", ("section", "bold"))
        text_widget.insert(tk.END, " • ", "bullet")
        text_widget.insert(tk.END, "ต้องเลือกเสมอ (สำคัญมาก!): ", ("bullet", "bold"))
        text_widget.insert(tk.END, "เมื่อคอลัมน์ที่เลือกมาวิเคราะห์มี 'หน่วย' หรือ 'สเกล' ที่ต่างกัน เช่น นำคอลัมน์ 'ราคา' (หลักพัน) มาวิเคราะห์ร่วมกับ 'คะแนนความพึงพอใจ' (1-5) หากไม่เลือกตัวนี้ คอลัมน์ 'ราคา' จะมีอิทธิพลต่อผลลัพธ์ทั้งหมด ทำให้ผลวิเคราะห์ผิดเพี้ยน\n", "bullet")
        text_widget.insert(tk.END, " • ", "bullet")
        text_widget.insert(tk.END, "อาจไม่เลือกก็ได้: ", ("bullet", "bold"))
        text_widget.insert(tk.END, "เมื่อทุกคอลัมน์ที่เลือกมีหน่วยและสเกลเดียวกันอยู่แล้ว เช่น เป็นคะแนน 1-7 ทุกคอลัมน์\n\n", "bullet")
        text_widget.insert(tk.END, "ส่วนที่ 2: การตั้งค่า MDS (หน้าต่าง MDS Settings)\n", "h2")
        text_widget.insert(tk.END, "3. Dimensions (comma-separated)\n", "h3")
        text_widget.insert(tk.END, "คืออะไร: ", ("section", "bold"))
        text_widget.insert(tk.END, "จำนวนมิติ (หรือแกน) ของ 'แผนที่' ที่ต้องการสร้างขึ้นเพื่อแสดงความสัมพันธ์ของข้อมูล\n", "section")
        text_widget.insert(tk.END, "คำแนะนำ: ", ("section", "bold"))
        text_widget.insert(tk.END, "ใช้ `2, 3` (ค่าเริ่มต้น) สำหรับการนำเสนอผลทั่วไป เพราะเป็นมิติที่มองเห็นและเข้าใจได้ง่ายที่สุด\n\n", "section")
        text_widget.insert(tk.END, "4. MDS Model Type (นี่คือการตั้งค่าที่สำคัญที่สุด!)\n", "h3")
        text_widget.insert(tk.END, "คืออะไร: ", ("section", "bold"))
        text_widget.insert(tk.END, "กฎที่อัลกอริทึมจะใช้ในการสร้างแผนที่ให้ใกล้เคียงกับข้อมูลตั้งต้นมากที่สุด\n", "section")
        text_widget.insert(tk.END, "ควรเลือกเมื่อใด:\n", ("section", "bold"))
        text_widget.insert(tk.END, " • ", "bullet")
        text_widget.insert(tk.END, "Metric (Absolute): ", ("bullet", "bold"))
        text_widget.insert(tk.END, "เมื่อข้อมูลตั้งต้นเป็นค่าเชิงปริมาณที่มีความหมายจริงๆ เช่น ระยะทาง, ราคา (ข้อมูลดิบเป็น Ratio Scale)\n", "bullet")
        text_widget.insert(tk.END, " • ", "bullet")
        text_widget.insert(tk.END, "Non-Metric: ", ("bullet", "bold"))
        text_widget.insert(tk.END, "เมื่อข้อมูลเป็นข้อมูลเชิงอันดับ (Ordinal) เช่น คะแนนความพึงพอใจ 1-5, ผลสำรวจ หรือเมื่อใช้ Distance Metric แบบ Correlation\n\n", "bullet")
        text_widget.insert(tk.END, "5. Repetitions (n_init)\n", "h3")
        text_widget.insert(tk.END, "คืออะไร: ", ("section", "bold"))
        text_widget.insert(tk.END, "จำนวนครั้งที่โปรแกรมจะ 'เริ่มคำนวณใหม่จากจุดสุ่ม' เพื่อหาผลลัพธ์ที่ดีที่สุด\n", "section")
        text_widget.insert(tk.END, "คำแนะนำ: ", ("section", "bold"))
        text_widget.insert(tk.END, "ค่าเริ่มต้น `10` มักจะเพียงพอ อาจเพิ่มเป็น `20` หรือ `50` หากข้อมูลซับซ้อนมากเพื่อให้ผลลัพธ์เสถียรขึ้น\n\n", "section")
        text_widget.insert(tk.END, "6. Max Iterations (max_iter)\n", "h3")
        text_widget.insert(tk.END, "คืออะไร: ", ("section", "bold"))
        text_widget.insert(tk.END, "จำนวนรอบการปรับตำแหน่งสูงสุดในแต่ละครั้งของการคำนวณ\n", "section")
        text_widget.insert(tk.END, "คำแนะนำ: ", ("section", "bold"))
        text_widget.insert(tk.END, "เพิ่มค่าเป็น `500` หรือ `1000` ก็ต่อเมื่อโปรแกรมแสดงคำเตือนว่า `MDS did not converge` เท่านั้น\n\n", "section")
        text_widget.insert(tk.END, "7. Random Seed (random_state)\n", "h3")
        text_widget.insert(tk.END, "คืออะไร: ", ("section", "bold"))
        text_widget.insert(tk.END, "ตัวเลขสำหรับ 'ล็อก' ผลการสุ่ม เพื่อให้ได้ผลลัพธ์หน้าตาเหมือนเดิมทุกครั้งที่รัน\n", "section")
        text_widget.insert(tk.END, "คำแนะนำ: ", ("section", "bold"))
        text_widget.insert(tk.END, "ควรใส่ตัวเลขไว้เสมอ เพื่อให้ผลการวิเคราะห์สามารถทำซ้ำได้ (Reproducible) หากเว้นว่างไว้ โปรแกรมจะสุ่มเลขให้และบันทึกไว้ใน Report\n", "section")

class MDSConfigWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.transient(app.root)
        self.app = app
        self.title("MDS Settings")
        self.geometry("450x260")
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(1, weight=1)
        def create_row(parent, label_text, row_idx):
            ttk.Label(parent, text=label_text).grid(row=row_idx, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        create_row(main_frame, "Dimensions (comma-separated):", 0)
        self.dims_var = tk.StringVar(value=self.app.mds_settings['dimensions'])
        ttk.Entry(main_frame, textvariable=self.dims_var).grid(row=0, column=1, sticky=tk.EW)
        create_row(main_frame, "MDS Model Type:", 1)
        self.metric_var = tk.StringVar(value='Metric (Absolute)' if self.app.mds_settings['metric'] else 'Non-Metric')
        ttk.Combobox(main_frame, textvariable=self.metric_var, values=['Metric (Absolute)', 'Non-Metric'], state='readonly').grid(row=1, column=1, sticky=tk.EW)
        create_row(main_frame, "Repetitions (n_init):", 2)
        self.n_init_var = tk.StringVar(value=str(self.app.mds_settings['n_init']))
        ttk.Entry(main_frame, textvariable=self.n_init_var).grid(row=2, column=1, sticky=tk.EW)
        create_row(main_frame, "Max Iterations (max_iter):", 3)
        self.max_iter_var = tk.StringVar(value=str(self.app.mds_settings['max_iter']))
        ttk.Entry(main_frame, textvariable=self.max_iter_var).grid(row=3, column=1, sticky=tk.EW)
        create_row(main_frame, "Random Seed (leave blank for random):", 4)
        self.random_state_var = tk.StringVar(value=str(self.app.mds_settings['random_state'] or ''))
        ttk.Entry(main_frame, textvariable=self.random_state_var).grid(row=4, column=1, sticky=tk.EW)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Save", command=self.save).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=10)
        self.grab_set()
        self.focus_set()
    def save(self):
        try:
            dims = [int(d.strip()) for d in self.dims_var.get().split(',') if d.strip()]
            if not all(d > 0 for d in dims): raise ValueError("Dimensions must be positive integers.")
            self.app.mds_settings.update({
                'dimensions': self.dims_var.get(),
                'metric': self.metric_var.get() == 'Metric (Absolute)',
                'n_init': int(self.n_init_var.get()),
                'max_iter': int(self.max_iter_var.get()),
                'random_state': int(self.random_state_var.get()) if self.random_state_var.get() else None
            })
            messagebox.showinfo("Saved", "MDS settings updated.", parent=self)
            self.destroy()
        except (ValueError, TypeError) as e:
            messagebox.showerror("Invalid Input", f"Please check your inputs.\nError: {e}", parent=self)

class ProximityMatrixApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel MDS Full Report Generator (v29 - Explanation Added)")
        self.root.geometry("850x750")
        
        self.mds_settings = {'dimensions': '2, 3', 'metric': True, 'n_init': 10, 'max_iter': 300, 'random_state': 42}
        self.df, self.proximity_df = None, None
        
        file_frame = ttk.LabelFrame(self.root, text="1. โหลดข้อมูล", padding="10")
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(file_frame, text="โหลดไฟล์ Excel", command=self.load_excel).pack(side=tk.LEFT)
        self.file_label = ttk.Label(file_frame, text="ยังไม่ได้เลือกไฟล์")
        self.file_label.pack(side=tk.LEFT, padx=10)

        col_frame = ttk.LabelFrame(self.root, text="2. ตั้งค่าและเลือกคอลัมน์", padding="10")
        col_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        calc_options_frame = ttk.Frame(col_frame)
        calc_options_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(calc_options_frame, text="Input: Distance Metric:").pack(side=tk.LEFT, padx=(0, 5))
        self.metric_method_var = tk.StringVar(value='euclidean')
        ttk.Combobox(calc_options_frame, textvariable=self.metric_method_var, 
                     values=['euclidean', 'cityblock', 'correlation', 'cosine'], 
                     state='readonly', width=12).pack(side=tk.LEFT, padx=5)

        self.standardize_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(calc_options_frame, text="Input: Standardize Data (Z-score)", variable=self.standardize_var).pack(side=tk.LEFT, padx=10)

        btn_box = ttk.Frame(col_frame)
        btn_box.pack(fill=tk.X, pady=5)
        ttk.Button(btn_box, text="เลือกทั้งหมด", command=lambda: self.cols_list.selection_set(0, tk.END)).pack(side=tk.LEFT)
        ttk.Button(btn_box, text="ไม่เลือกทั้งหมด", command=lambda: self.cols_list.selection_clear(0, tk.END)).pack(side=tk.LEFT, padx=5)
        
        self.cols_list = tk.Listbox(col_frame, selectmode=tk.MULTIPLE, height=8)
        self.cols_list.pack(fill=tk.BOTH, expand=True)

        act_frame = ttk.LabelFrame(self.root, text="3. ดำเนินการ", padding="10")
        act_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(act_frame, text="คำนวณ Matrix", command=self.calc_matrix).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(act_frame, text="ตั้งค่า MDS", command=lambda: MDSConfigWindow(self)).pack(side=tk.LEFT, padx=(0, 10))
        self.exp_btn = ttk.Button(act_frame, text="Export Report", command=self.export_report_threaded, state=tk.DISABLED)
        self.exp_btn.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(act_frame, text="คำแนะนำ (About)", command=lambda: AboutWindow(self.root)).pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=5)

        res_frame = ttk.LabelFrame(self.root, text="Proximity Matrix Preview", padding="10")
        res_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.tree = ttk.Treeview(res_frame, show='headings')
        self.tree.pack(fill=tk.BOTH, expand=True, pady=5)
        
    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=(("Excel", "*.xlsx *.xls"),))
        if path:
            try:
                self.df = pd.read_excel(path, index_col=0)
                self.file_label.config(text=path.split('/')[-1])
                self.cols_list.delete(0, tk.END)
                for c in self.df.columns: self.cols_list.insert(tk.END, c)
                self.exp_btn.config(state=tk.DISABLED)
                for item in self.tree.get_children(): self.tree.delete(item)
            except Exception as e: messagebox.showerror("Error loading file", str(e))

    def calc_matrix(self):
        sels = [self.cols_list.get(i) for i in self.cols_list.curselection()]
        if len(sels) < 2: return messagebox.showwarning("Selection required", "Please select at least 2 columns.")
        try:
            data_to_process = self.df[sels].apply(pd.to_numeric, errors='coerce')
            if data_to_process.isnull().any().any():
                raise ValueError("Selected columns contain non-numeric data.")

            if self.standardize_var.get():
                scaler = StandardScaler()
                data_to_process = pd.DataFrame(scaler.fit_transform(data_to_process), 
                                                columns=data_to_process.columns, 
                                                index=data_to_process.index)

            metric_method = self.metric_method_var.get()
            dist_matrix = squareform(pdist(data_to_process.T, metric_method))
            self.proximity_df = pd.DataFrame(dist_matrix, index=sels, columns=sels)

            for item in self.tree.get_children(): self.tree.delete(item)
            self.tree["columns"] = [""] + sels
            self.tree.column("#0", width=0, stretch=tk.NO)
            self.tree.column("", width=120, anchor=tk.W)
            self.tree.heading("", text="Item")
            for c in sels:
                self.tree.column(c, width=80, anchor=tk.CENTER)
                self.tree.heading(c, text=c)
            for index, row in self.proximity_df.iterrows():
                self.tree.insert("", tk.END, values=[index] + [f"{val:.3f}" for val in row])
            self.exp_btn.config(state=tk.NORMAL)
        except Exception as e: messagebox.showerror("Calculation Error", str(e))

    def export_report_threaded(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel Files", "*.xlsx"),))
        if not path: return
        self.exp_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0
        thread = threading.Thread(target=self.export_report, args=(path,))
        thread.start()
        self.monitor_thread(thread)

    def monitor_thread(self, thread):
        if thread.is_alive():
            self.root.after(100, lambda: self.monitor_thread(thread))
        else:
            self.exp_btn.config(state=tk.NORMAL)
            self.progress['value'] = 100 

    def export_report(self, path):
        seed = self.mds_settings['random_state'] or np.random.randint(0, 2**31 - 1)
        try:
            with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
                self.proximity_df.to_excel(writer, sheet_name='Proximity_Matrix')
                dims_to_run = [int(d.strip()) for d in self.mds_settings['dimensions'].split(',') if d.strip()]
                self.progress['maximum'] = len(dims_to_run)
                for i, dim in enumerate(dims_to_run):
                    self.create_mds_sheet(writer, dim, seed)
                    self.root.after(0, lambda v=i+1: self.progress.config(value=v))
                
                # *** NEW: Add the explanation sheet at the end ***
                self._create_explanation_sheet(writer)

            self.root.after(0, lambda: messagebox.showinfo("Success", f"Report saved successfully to:\n{path}"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Export Error", str(e)))

    def create_mds_sheet(self, writer, n_dim, seed):
        mds = MDS(n_components=n_dim, dissimilarity='precomputed', random_state=seed,
                  metric=self.mds_settings['metric'], n_init=self.mds_settings['n_init'],
                  max_iter=self.mds_settings['max_iter'], normalized_stress=False)
        coords = mds.fit_transform(self.proximity_df)
        labels = self.proximity_df.columns
        metrics = self._calculate_mds_metrics(self.proximity_df.values, coords, mds)
        summary_df = self._create_summary_df(n_dim, seed, metrics)
        sheet_name = f'MDS_{n_dim}D_Results'
        summary_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
        
        row_cursor = len(summary_df) + 4
        all_tables = self._create_main_result_tables(n_dim, coords, labels, metrics)
        for title, df in all_tables:
            writer.sheets[sheet_name].write(row_cursor, 0, title)
            write_index = not title.startswith("Pairwise")
            df.to_excel(writer, sheet_name=sheet_name, startrow=row_cursor + 1, index=write_index)
            row_cursor += len(df) + 3

        img_data_shepard = self._plot_to_bytes(self._create_shepard_diagram, metrics)
        writer.sheets[sheet_name].write(row_cursor, 0, 'Shepard Diagram:')
        writer.sheets[sheet_name].insert_image(row_cursor + 1, 0, '', {'image_data': img_data_shepard})
        row_cursor += 25

        if n_dim >= 2:
            img_data_scatter2d = self._plot_to_bytes(self._create_scatter_plot, coords, labels)
            writer.sheets[sheet_name].write(row_cursor, 0, '2D Perceptual Map (Dim1 vs Dim2):')
            writer.sheets[sheet_name].insert_image(row_cursor + 1, 0, '', {'image_data': img_data_scatter2d})
            row_cursor += 25
        
        if n_dim >= 3:
            img_data_scatter3d = self._plot_to_bytes(self._create_scatter_plot, coords, labels, is_3d=True)
            writer.sheets[sheet_name].write(row_cursor, 0, '3D Perceptual Map:')
            writer.sheets[sheet_name].insert_image(row_cursor + 1, 0, '', {'image_data': img_data_scatter3d})
            self._create_3d_extra_sheet(writer, coords, labels)
            
    def _calculate_mds_metrics(self, original_matrix, coords, mds_model):
        final_dist_matrix = squareform(pdist(coords, 'euclidean'))
        disparities = original_matrix if self.mds_settings['metric'] else mds_model.disparities_
        iu = np.triu_indices(len(original_matrix), k=1)
        original_flat = original_matrix[iu]
        final_flat = final_dist_matrix[iu]
        delta_vals = disparities[iu]
        num = np.sum((final_flat - delta_vals) ** 2)
        den = np.sum(final_flat ** 2)
        kruskal_stress = np.sqrt(num / den) if den > 0 else 0
        r_square_correct = np.corrcoef(original_flat, final_flat)[0, 1] ** 2
        r_square_client = np.corrcoef(original_matrix.flatten(), final_dist_matrix.flatten())[0, 1] ** 2
        return {'final_dist_matrix': final_dist_matrix, 'disparities': disparities, 'original_flat': original_flat, 'final_flat': final_flat, 'delta_vals': delta_vals, 'kruskal_stress': kruskal_stress, 'r_square_correct': r_square_correct, 'r_square_client': r_square_client}

    def _create_summary_df(self, n_dim, seed, metrics):
        summary_data = {'Setting': ['Dimensional Space', 'Kruskal\'s Stress (1)', 'R-square (Corrected)', "R-square (Full Matrix)", 'MDS Model Type', 'Input: Distance Metric', 'Input: Standardized', 'Repetitions (n_init)', 'Max Iterations (max_iter)', 'Seed (random numbers)'], 'Value': [f'{n_dim} Dimensions', f"{metrics['kruskal_stress']:.4f}", f"{metrics['r_square_correct']:.4f}", f"{metrics['r_square_client']:.4f}", 'Metric (Absolute)' if self.mds_settings['metric'] else 'Non-Metric', self.metric_method_var.get(), str(self.standardize_var.get()), self.mds_settings['n_init'], self.mds_settings['max_iter'], str(seed)]}
        return pd.DataFrame(summary_data)

    def _create_main_result_tables(self, n_dim, coords, labels, metrics):
        main_tables = [('Configuration:', pd.DataFrame(coords, index=labels, columns=[f'Dim{i+1}' for i in range(n_dim)]))]
        if n_dim >= 2:
            pairwise_r2 = []
            final_distances_in_nd_space = metrics['final_flat']
            for pair_indices in combinations(range(n_dim), 2):
                distances_in_2d_plane = pdist(coords[:, list(pair_indices)], 'euclidean')
                r2_pair = np.corrcoef(final_distances_in_nd_space, distances_in_2d_plane)[0, 1] ** 2
                pair_name = f"dim{pair_indices[0]+1}-dim{pair_indices[1]+1}"
                pairwise_r2.append({'Output': pair_name, 'R-square': f"{r2_pair*100:.2f}"})
            main_tables.append(("Pairwise Dimension R-square:", pd.DataFrame(pairwise_r2)))
        main_tables.extend([('Distances measured in the representation space:', pd.DataFrame(metrics['final_dist_matrix'], index=labels, columns=labels)), ('Disparities computed using the model:', pd.DataFrame(metrics['disparities'], index=labels, columns=labels)), ('Residual distances:', pd.DataFrame(metrics['disparities'] - metrics['final_dist_matrix'], index=labels, columns=labels))])
        return main_tables
    
    def _create_3d_extra_sheet(self, writer, coords, labels):
        extra_sheet_name = 'MDS_3D_Extra_Analysis'
        extra_ws = writer.book.add_worksheet(extra_sheet_name)
        extra_row_cursor = 0

        def get_sq_euclidean_df(coordinates):
            sq_dist_matrix = squareform(pdist(coordinates, metric='sqeuclidean'))
            return pd.DataFrame(sq_dist_matrix, index=labels, columns=labels)
        
        sq_prox_df = self.proximity_df ** 2
        sq_dist_3d_df = get_sq_euclidean_df(coords)
        sq_dist_d1d2_df = get_sq_euclidean_df(coords[:, [0, 1]])
        sq_dist_d1d3_df = get_sq_euclidean_df(coords[:, [0, 2]])
        sq_dist_d2d3_df = get_sq_euclidean_df(coords[:, [1, 2]])
        rsq_proof_results = []
        iu = np.triu_indices(len(labels), k=1)
        base_sq_upper_triangle = sq_dist_3d_df.values[iu]
        base_sq_full_matrix = sq_dist_3d_df.values.flatten()
        
        planes_to_test = [("Dim1-Dim2 (Table 2 vs Table 1)", sq_dist_d1d2_df), ("Dim1-Dim3 (Table 3 vs Table 1)", sq_dist_d1d3_df), ("Dim2-Dim3 (Table 4 vs Table 1)", sq_dist_d2d3_df)]
        for name, df_plane in planes_to_test:
            plane_sq_upper_triangle = df_plane.values[iu]
            r_square_upper = np.corrcoef(base_sq_upper_triangle, plane_sq_upper_triangle)[0, 1] ** 2
            plane_sq_full_matrix = df_plane.values.flatten()
            if np.var(base_sq_full_matrix) == 0 or np.var(plane_sq_full_matrix) == 0:
                 r_square_full = 1.0 if np.array_equal(base_sq_full_matrix, plane_sq_full_matrix) else 0.0
            else:
                r_square_full = np.corrcoef(base_sq_full_matrix, plane_sq_full_matrix)[0, 1] ** 2
            rsq_proof_results.append({'Comparison': name, 'R-square (Upper Triangle)': f"{r_square_upper*100:.5f}", 'R-square (Excel RSQ - Full Matrix)': f"{r_square_full*100:.5f}"})
        rsq_proof_df = pd.DataFrame(rsq_proof_results)
        extra_tables = [("Table 0: Squared Proximity Matrix (Original Distances^2)", sq_prox_df), ("Table 1: Squared Euclidean Distances from 3D Space (Dim1, Dim2, Dim3)", sq_dist_3d_df), ("Table 2: Squared Euclidean Distances from Dim1-Dim2 Plane", sq_dist_d1d2_df), ("Table 3: Squared Euclidean Distances from Dim1-Dim3 Plane", sq_dist_d1d3_df), ("Table 4: Squared Euclidean Distances from Dim2-Dim3 Plane", sq_dist_d2d3_df), ("R-square Comparison (Planes vs. Full 3D Space - Table 1)", rsq_proof_df)]
        for title, df in extra_tables:
            extra_ws.write(extra_row_cursor, 0, title)
            write_index = not title.startswith("R-square")
            df.to_excel(writer, sheet_name=extra_sheet_name, startrow=extra_row_cursor + 1, index=write_index)
            extra_row_cursor += len(df) + 3

    def _plot_to_bytes(self, plot_function, *args, **kwargs):
        fig = plot_function(*args, **kwargs)
        img_data = io.BytesIO()
        fig.savefig(img_data, format='png', bbox_inches='tight')
        plt.close(fig)
        return img_data

    def _create_shepard_diagram(self, metrics):
        fig, ax = plt.subplots(figsize=(6, 5))
        ax.scatter(metrics['original_flat'], metrics['final_flat'], facecolors='none', edgecolors='k', alpha=0.7, label='Distances')
        sorted_indices = np.argsort(metrics['original_flat'])
        ax.plot(metrics['original_flat'][sorted_indices], metrics['delta_vals'][sorted_indices], 'r-', label='Disparities (Monotone Fit)')
        ax.set_title(f"Shepard Diagram (Stress-1 = {metrics['kruskal_stress']:.3f})")
        ax.set_xlabel("Dissimilarity (Input)"); ax.set_ylabel("Distance / Disparity (Output)")
        ax.legend(); ax.grid(True)
        return fig
    
    def _create_scatter_plot(self, coords, labels, is_3d=False):
        fig = plt.figure(figsize=(7, 6))
        if is_3d and coords.shape[1] >= 3:
            ax = fig.add_subplot(111, projection='3d')
            ax.scatter(coords[:, 0], coords[:, 1], coords[:, 2])
            for i, label in enumerate(labels):
                ax.text(coords[i, 0], coords[i, 1], coords[i, 2], label, size=8)
            ax.set_xlabel("Dimension 1"); ax.set_ylabel("Dimension 2"); ax.set_zlabel("Dimension 3")
            ax.set_title("3D Perceptual Map")
        else:
            ax = fig.add_subplot(111)
            ax.scatter(coords[:, 0], coords[:, 1])
            for i, label in enumerate(labels):
                ax.text(coords[i, 0], coords[i, 1], label, size=9)
            ax.set_xlabel("Dimension 1"); ax.set_ylabel("Dimension 2")
            ax.set_title("2D Perceptual Map")
            ax.grid(True); ax.axhline(0, color='grey', lw=0.5); ax.axvline(0, color='grey', lw=0.5)
        fig.tight_layout()
        return fig

    def _create_explanation_sheet(self, writer):
        """Creates a new sheet in the Excel file with detailed explanations."""
        wb = writer.book
        ws = wb.add_worksheet("R-square Explanation")

        # --- Define Formats ---
        title_format = wb.add_format({'bold': True, 'font_size': 16, 'valign': 'vcenter'})
        header_format = wb.add_format({'bold': True, 'font_size': 12, 'underline': True, 'valign': 'vcenter'})
        bold_format = wb.add_format({'bold': True, 'font_size': 11})
        body_format = wb.add_format({'font_size': 11, 'valign': 'top', 'text_wrap': True})
        table_header_format = wb.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        table_body_format = wb.add_format({'border': 1, 'valign': 'top', 'text_wrap': True})

        # --- Set Column Widths ---
        ws.set_column('A:A', 60)
        ws.set_column('B:B', 45)
        ws.set_column('C:C', 45)
        
        row = 0
        # --- Main Title ---
        ws.write(row, 0, "คำอธิบายความแตกต่างของค่า R-square ในตารางต่างๆ", title_format)
        row += 2

        # --- Introduction ---
        ws.write(row, 0, "การที่ค่า R-square ในตาราง 'Pairwise Dimension R-square' และ 'R-square Comparison' ไม่เท่ากันแต่มีค่าใกล้เคียงกันนั้น เป็นสิ่งที่ถูกต้องและเกิดขึ้นจากวัตถุประสงค์และวิธีการคำนวณที่แตกต่างกัน ดังนี้:", body_format)
        row += 2
        
        # --- Section 1: Pairwise Dimension R-square ---
        ws.write(row, 0, "1. ตาราง 'Pairwise Dimension R-square' (ในชีต MDS_2D/3D_Results)", header_format)
        row += 1
        ws.write_rich_string(row, 0, body_format, 'ตอบคำถามว่า: "ถ้าเรามอง \'แผนที่ 3 มิติ\' แล้วฉายภาพลงบนระนาบ 2 มิติ ภาพ 2 มิตินั้นจะยังคงรักษาระยะห่างสัมพัทธ์ของจุดต่างๆ จากภาพ 3 มิติเดิมได้ดีแค่ไหน?"')
        row += 2
        ws.write_rich_string(row, 0, bold_format, 'ข้อมูลที่ใช้คำนวณ:', body_format, ' คำนวณจาก ', bold_format, '\"ระยะทางจริง\" (Euclidean Distance)')
        row += 1
        ws.write(row, 0, " •  ชุดข้อมูล A: ระยะทางจริงระหว่างทุกคู่จุดในพื้นที่ N มิติ (เช่น 3 มิติ)", body_format)
        row += 1
        ws.write(row, 0, " •  ชุดข้อมูล B: ระยะทางจริงระหว่างทุกคู่จุดในระนาบ 2 มิติ (เช่น Dim1-Dim2)", body_format)
        row += 2
        ws.write_rich_string(row, 0, bold_format, 'การตีความ:', body_format, ' เป็นค่าที่เข้าใจง่ายที่สุด บอกความสัมพันธ์ของ \"ระยะห่าง\" ตรงๆ ว่าแผนภาพ 2 มิติ กับแผนภาพ 3 มิติ ให้ผลเรื่องระยะทางสอดคล้องกันมากน้อยเพียงใด')
        row += 3

        # --- Section 2: R-square Comparison ---
        ws.write(row, 0, "2. ตาราง 'R-square Comparison' (ในชีต MDS_3D_Extra_Analysis)", header_format)
        row += 1
        ws.write_rich_string(row, 0, body_format, 'ตอบคำถามว่า: "ความแปรปรวน (Variance) ทั้งหมดของระยะห่างในพื้นที่ 3 มิติ สามารถถูกอธิบายโดยความแปรปรวนของระยะห่างในระนาบ 2 มิตินี้ได้กี่เปอร์เซ็นต์?"')
        row += 2
        ws.write_rich_string(row, 0, bold_format, 'ข้อมูลที่ใช้คำนวณ:', body_format, ' คำนวณจาก ', bold_format, '\"ระยะทางยกกำลังสอง\" (Squared Euclidean Distance)')
        row += 1
        ws.write(row, 0, " •  ชุดข้อมูล X: ระยะทางยกกำลังสองระหว่างทุกคู่จุดในพื้นที่ 3 มิติ (จาก Table 1)", body_format)
        row += 1
        ws.write(row, 0, " •  ชุดข้อมูล Y: ระยะทางยกกำลังสองระหว่างทุกคู่จุดในระนาบ 2 มิติ (จาก Table 2, 3, 4)", body_format)
        row += 2
        ws.write_rich_string(row, 0, bold_format, 'การตีความ:', body_format, ' เป็นการวิเคราะห์เชิงเทคนิคเพื่อพิสูจน์การกระจายตัวของความแปรปรวน (Partitioning of Variance) ซึ่งมีความสำคัญทางคณิตศาสตร์ เพราะสอดคล้องกับทฤษฎีบทพีทาโกรัส (`a² + b² = c²`)')
        row += 3

        # --- Summary Table ---
        ws.write(row, 0, "สรุปเปรียบเทียบ", header_format)
        row += 1
        headers = ["ประเด็น (Aspect)", "Pairwise Dimension R-square (ชีตหลัก)", "R-square Comparison (ชีต Extra)"]
        ws.write_row(row, 0, headers, table_header_format)
        row += 1
        
        table_data = [
            ["เป้าหมาย", "วัดความสอดคล้องของ \"ระยะห่าง\" ที่มองเห็นได้", "พิสูจน์การกระจายตัวของ \"ความแปรปรวน\""],
            ["ข้อมูลที่ใช้", "ระยะทางจริง (Euclidean Distance)", "ระยะทางยกกำลังสอง (Squared Distance)"],
            ["การตีความ", "เข้าใจง่าย, เน้นการนำไปใช้แสดงผล", "เทคนิค, เน้นการพิสูจน์เชิงสถิติ"]
        ]
        
        for data_row in table_data:
            ws.write_row(row, 0, data_row, table_body_format)
            ws.set_row(row, 45) # Set row height
            row += 1



    
# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None): # ชื่อฟังก์ชันนี้จะถูกใช้ใน Launcher
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน QuotaSamplerApp.
    """
    print(f"--- QUOTA_SAMPLER_INFO: Starting 'QuotaSamplerApp' via run_this_app() ---")
    try:
        # --- โค้ดที่ย้ายมาจาก if __name__ == "__main__": เดิมจะมาอยู่ที่นี่ ---
        # --- ส่วนที่ใช้รันโปรแกรม ---
    #if __name__ == "__main__":
        root = tk.Tk()
        style = ttk.Style(root)
        try:
            available_themes = style.theme_names()
            if 'clam' in available_themes:
                style.theme_use('clam')
            elif 'vista' in available_themes:
                style.theme_use('vista')
        except tk.TclError:
            print("ttk themes not available.")
        app = ProximityMatrixApp(root)
        root.mainloop()
        
        print(f"--- QUOTA_SAMPLER_INFO: QuotaSamplerApp mainloop finished. ---")

    except Exception as e:
        # ดักจับ Error ที่อาจเกิดขึ้นระหว่างการสร้างหรือรัน App
        print(f"QUOTA_SAMPLER_ERROR: An error occurred during QuotaSamplerApp execution: {e}")
        # แสดง Popup ถ้ามีปัญหา
        if 'root' not in locals() or not root.winfo_exists(): # สร้าง root ชั่วคราวถ้ายังไม่มี
            root_temp = tk.Tk()
            root_temp.withdraw()
            messagebox.showerror("Application Error (Quota Sampler)",
                               f"An unexpected error occurred:\n{e}", parent=root_temp)
            root_temp.destroy()
        else:
            messagebox.showerror("Application Error (Quota Sampler)",
                               f"An unexpected error occurred:\n{e}", parent=root) # ใช้ root ที่มีอยู่ถ้าเป็นไปได้
        sys.exit(f"Error running QuotaSamplerApp: {e}") # อาจจะ exit หรือไม่ก็ได้ ขึ้นกับการออกแบบ


# --- ส่วน Run Application เมื่อรันไฟล์นี้โดยตรง (สำหรับ Test) ---
if __name__ == "__main__":
    print("--- Running QuotaSamplerApp.py directly for testing ---")
    # (ถ้ามีการตั้งค่า DPI ด้านบน มันจะทำงานอัตโนมัติ)

    # เรียกฟังก์ชัน Entry Point ที่เราสร้างขึ้น
    run_this_app()

    print("--- Finished direct execution of QuotaSamplerApp.py ---")
# <<< END OF CHANGES >>>