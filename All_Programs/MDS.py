# วางบรรทัดเหล่านี้ไว้บนสุดของไฟล์ MDS.py
import sys, io, os
# ถ้ารันจากไฟล์ที่ถูก freeze และไม่มีคอนโซล ให้ผูก stdout/stderr เข้ากับที่ทิ้ง
if getattr(sys, 'frozen', False):
    if sys.stdout is None:
        sys.stdout = open(os.devnull, 'w', buffering=1)
    if sys.stderr is None:
        sys.stderr = open(os.devnull, 'w', buffering=1)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from sklearn.manifold import MDS
import matplotlib.pyplot as plt
from scipy.spatial.distance import pdist, squareform
import io

# ===================================================================
# ฟังก์ชันสำหรับจัดหน้าต่างให้อยู่กลางจอ
# ===================================================================
def center_window(win):
    """
    ฟังก์ชันสำหรับจัดหน้าต่าง (Tk หรือ Toplevel) ให้อยู่กลางหน้าจอ
    """
    win.update_idletasks()  # อัปเดตข้อมูลขนาดของหน้าต่างให้เป็นปัจจุบันก่อน
    width = win.winfo_width()
    height = win.winfo_height()
    x = (win.winfo_screenwidth() // 2) - (width // 2)
    y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry(f'{width}x{height}+{x}+{y}')

# ===================================================================
# NEW CLASS: หน้าต่างสำหรับแสดงคำแนะนำ
# ===================================================================
class AboutWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw() # ซ่อนหน้าต่างทันทีที่สร้าง

        self.title("เกี่ยวกับค่า Config และคำแนะนำ")
        self.geometry("800x650") # กำหนดขนาดให้ใหญ่พอสำหรับข้อความ
        self.transient(parent)

        # --- สร้าง Frame หลักและ Scrollbar ---
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(expand=True, fill=tk.BOTH)

        scrollbar = ttk.Scrollbar(main_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # --- สร้าง Text widget สำหรับแสดงข้อความ ---
        text_widget = tk.Text(
            main_frame,
            wrap='word',
            yscrollcommand=scrollbar.set,
            padx=10,
            pady=10,
            font=("Segoe UI", 10),
            background=self.cget('bg'),
            relief=tk.FLAT
        )
        text_widget.pack(expand=True, fill=tk.BOTH)
        scrollbar.config(command=text_widget.yview)

        # --- ใส่เนื้อหาคำอธิบาย ---
        self._populate_text(text_widget)
        text_widget.config(state=tk.DISABLED)

        # --- ปุ่มปิด ---
        button_frame = ttk.Frame(self, padding=(0, 5, 0, 10))
        button_frame.pack(fill=tk.X)
        close_button = ttk.Button(button_frame, text="ปิด", command=self.destroy)
        close_button.pack()

        self.grab_set()
        self.focus_set()
        
        # จัดตำแหน่งกลางจอ แล้วค่อยแสดงผล เพื่อความนุ่มนวล
        center_window(self)
        self.deiconify()

    def _populate_text(self, text_widget):
        text_widget.tag_configure("h1", font=("Segoe UI", 14, "bold"), spacing3=10)
        text_widget.tag_configure("h2", font=("Segoe UI", 11, "bold"), spacing3=8)
        text_widget.tag_configure("bold", font=("Segoe UI", 10, "bold"))
        text_widget.tag_configure("bullet", lmargin1=20, lmargin2=20)
        text_widget.insert(tk.END, "คำอธิบายค่า Config ในหน้าต่าง MDS Settings\n\n", "h1")
        text_widget.insert(tk.END, "หน้าต่างนี้ใช้สำหรับปรับแต่งพารามิเตอร์ของอัลกอริทึม Multidimensional Scaling (MDS) ซึ่งเป็นเทคนิคทางสถิติที่ใช้วาดแผนภาพแสดงความสัมพันธ์ (ความคล้าย/แตกต่าง) ระหว่างข้อมูลต่างๆ\n\n")
        text_widget.insert(tk.END, "1. Dimensions (comma-separated)\n", "h2")
        text_widget.insert(tk.END, "คือจำนวน 'มิติ' ที่คุณต้องการให้ผลลัพธ์แสดงออกมา เช่น 2 (สำหรับแผนภาพ 2 มิติ) และ 3 (สำหรับแผนภาพ 3 มิติ) คุณสามารถใส่หลายค่าโดยคั่นด้วยจุลภาค (comma) เพื่อสร้างผลลัพธ์หลายรูปแบบ\n\n")
        text_widget.insert(tk.END, "2. MDS Model Type\n", "h2")
        text_widget.insert(tk.END, "คือรูปแบบของ MDS ที่จะใช้ในการคำนวณ:\n")
        text_widget.insert(tk.END, "• Metric (Absolute): ", ("bullet", "bold"))
        text_widget.insert(tk.END, "พยายามรักษาระยะห่าง 'จริงๆ' ของข้อมูลให้ได้มากที่สุด เหมาะสำหรับข้อมูลที่เป็นตัวเลขและมีสเกลที่ชัดเจน\n", "bullet")
        text_widget.insert(tk.END, "• Non-Metric: ", ("bullet", "bold"))
        text_widget.insert(tk.END, "พยายามรักษาแค่ 'ลำดับ' ของระยะห่าง (อะไรใกล้กว่าอะไร) เหมาะสำหรับข้อมูลเชิงอันดับ เช่น ผลสำรวจความพึงพอใจ\n\n", "bullet")
        text_widget.insert(tk.END, "3. Repetitions (n_init)\n", "h2")
        text_widget.insert(tk.END, "คือจำนวนครั้งที่อัลกอริทึมจะ 'สุ่มจุดเริ่มต้น' แล้วเริ่มคำนวณใหม่ทั้งหมด โปรแกรมจะเลือกผลลัพธ์จากรอบที่ดีที่สุด (มีค่าความคลาดเคลื่อนน้อยที่สุด) มาเป็นคำตอบสุดท้าย การเพิ่มค่านี้จะเพิ่มโอกาสได้ผลลัพธ์ที่ดีขึ้น แต่ใช้เวลาคำนวณนานขึ้น\n\n")
        text_widget.insert(tk.END, "4. Max Iterations (max_iter)\n", "h2")
        text_widget.insert(tk.END, "คือจำนวน 'รอบการปรับปรุง' สูงสุดในแต่ละครั้งของการคำนวณ เป็นการป้องกันไม่ให้โปรแกรมทำงานนานเกินไป หากผลลัพธ์ยังไม่นิ่งแต่ครบจำนวนรอบแล้ว โปรแกรมจะหยุดและใช้ผลลัพธ์นั้น\n\n")
        text_widget.insert(tk.END, "5. Random Seed (random_state)\n", "h2")
        text_widget.insert(tk.END, "คือตัวเลขสำหรับ 'ล็อกผลการสุ่ม' เพื่อให้ได้ผลลัพธ์หน้าตาเหมือนเดิมทุกครั้งที่รันด้วยข้อมูลและค่าตั้งต้นชุดเดียวกัน ซึ่งสำคัญมากสำหรับการทำซ้ำผลการทดลอง (Reproducibility) หากเว้นว่างไว้ ผลลัพธ์อาจมีหน้าตาต่างกันเล็กน้อยในแต่ละครั้งที่รัน\n\n")
        text_widget.insert(tk.END, "คำแนะนำ: ค่าเริ่มต้นที่ให้มาเป็นค่ามาตรฐานที่ดี\n", "h1")
        text_widget.insert(tk.END, "ค่าเหล่านี้เป็นจุดเริ่มต้นที่ปลอดภัยและเหมาะสมกับการใช้งานส่วนใหญ่ แต่คุณควรปรับเปลี่ยนตามลักษณะของข้อมูลของคุณ:\n")
        text_widget.insert(tk.END, "• ใช้ Non-Metric: ", ("bullet", "bold"))
        text_widget.insert(tk.END, "ถ้าข้อมูลของคุณมาจากการสำรวจ, ความรู้สึก, หรือการจัดอันดับ\n", "bullet")
        text_widget.insert(tk.END, "• เพิ่ม max_iter: ", ("bullet", "bold"))
        text_widget.insert(tk.END, "ถ้าโปรแกรมแจ้งเตือนว่า 'MDS did not converge'\n", "bullet")
        text_widget.insert(tk.END, "• เพิ่ม n_init: ", ("bullet", "bold"))
        text_widget.insert(tk.END, "ถ้าข้อมูลมีความซับซ้อนสูง และต้องการความมั่นใจในผลลัพธ์มากที่สุด\n", "bullet")


class MDSConfigWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.withdraw() # ซ่อนหน้าต่างทันทีที่สร้าง

        self.transient(app.root)
        self.app = app
        self.title("MDS Settings")
        self.geometry("450x260")
        self.result = None

        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(1, weight=1)

        def create_setting_row(parent, label_text, row_index):
            label = ttk.Label(parent, text=label_text)
            label.grid(row=row_index, column=0, sticky=tk.W, padx=(0, 10), pady=5)

        create_setting_row(main_frame, "Dimensions (comma-separated):", 0)
        self.dims_var = tk.StringVar(value=self.app.mds_settings['dimensions'])
        dims_entry = ttk.Entry(main_frame, textvariable=self.dims_var)
        dims_entry.grid(row=0, column=1, sticky=tk.EW)

        create_setting_row(main_frame, "MDS Model Type:", 1)
        self.metric_var = tk.StringVar(value='Metric (Absolute)' if self.app.mds_settings['metric'] else 'Non-Metric')
        model_combo = ttk.Combobox(main_frame, textvariable=self.metric_var, values=['Metric (Absolute)', 'Non-Metric'], state='readonly')
        model_combo.grid(row=1, column=1, sticky=tk.EW)

        create_setting_row(main_frame, "Repetitions (n_init):", 2)
        self.n_init_var = tk.StringVar(value=str(self.app.mds_settings['n_init']))
        n_init_entry = ttk.Entry(main_frame, textvariable=self.n_init_var)
        n_init_entry.grid(row=2, column=1, sticky=tk.EW)

        create_setting_row(main_frame, "Max Iterations (max_iter):", 3)
        self.max_iter_var = tk.StringVar(value=str(self.app.mds_settings['max_iter']))
        max_iter_entry = ttk.Entry(main_frame, textvariable=self.max_iter_var)
        max_iter_entry.grid(row=3, column=1, sticky=tk.EW)

        create_setting_row(main_frame, "Random Seed (leave blank for random):", 4)
        self.random_state_var = tk.StringVar(value=str(self.app.mds_settings['random_state'] or ''))
        seed_entry = ttk.Entry(main_frame, textvariable=self.random_state_var)
        seed_entry.grid(row=4, column=1, sticky=tk.EW)

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="Save", command=self.save_settings).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=10)

        self.grab_set()
        self.focus_set()
        
        # จัดตำแหน่งกลางจอ แล้วค่อยแสดงผล เพื่อความนุ่มนวล
        center_window(self)
        self.deiconify()

    def save_settings(self):
        try:
            dims_str = self.dims_var.get()
            if not dims_str: raise ValueError("Dimensions field cannot be empty.")
            dims = [int(d.strip()) for d in dims_str.split(',') if d.strip()]
            if not all(d > 0 for d in dims): raise ValueError("Dimensions must be positive integers.")

            new_settings = {
                'dimensions': dims_str,
                'metric': self.metric_var.get() == 'Metric (Absolute)',
                'n_init': int(self.n_init_var.get()),
                'max_iter': int(self.max_iter_var.get()),
                'random_state': int(self.random_state_var.get()) if self.random_state_var.get() else None
            }
            self.app.mds_settings = new_settings
            messagebox.showinfo("Settings Saved", "MDS settings have been updated.", parent=self)
            self.destroy()
        except ValueError as e:
            messagebox.showerror("Invalid Input", f"Please check your inputs.\nError: {e}", parent=self)

class ProximityMatrixApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel MDS Full Report Generator (v12 - Final)")
        self.root.geometry("800x650")
        center_window(self.root) # ตั้งค่าตำแหน่ง แต่ยังไม่แสดงผล

        self.mds_settings = {
            'dimensions': '2, 3',
            'metric': True,
            'n_init': 10,
            'max_iter': 300,
            'random_state': 42
        }

        self.df = None
        self.source_data_df = None
        self.proximity_df = None

        file_frame = ttk.Frame(self.root, padding="10")
        file_frame.pack(fill=tk.X)
        self.load_button = ttk.Button(file_frame, text="1. โหลดไฟล์ Excel", command=self.load_excel)
        self.load_button.pack(side=tk.LEFT)
        self.file_label = ttk.Label(file_frame, text="ยังไม่ได้เลือกไฟล์")
        self.file_label.pack(side=tk.LEFT, padx=10)

        columns_frame = ttk.Frame(self.root, padding="10")
        columns_frame.pack(fill=tk.BOTH, expand=True)
        columns_label = ttk.Label(columns_frame, text="2. เลือกคอลัมน์ที่ต้องการวิเคราะห์:")
        columns_label.pack(anchor=tk.W)
        selection_buttons_frame = ttk.Frame(columns_frame)
        selection_buttons_frame.pack(fill=tk.X, pady=(5, 2))
        self.select_all_button = ttk.Button(selection_buttons_frame, text="เลือกทั้งหมด", command=self.select_all_columns)
        self.select_all_button.pack(side=tk.LEFT)
        self.deselect_all_button = ttk.Button(selection_buttons_frame, text="ไม่เลือกทั้งหมด", command=self.deselect_all_columns)
        self.deselect_all_button.pack(side=tk.LEFT, padx=5)
        self.columns_listbox = tk.Listbox(columns_frame, selectmode=tk.MULTIPLE, height=8)
        self.columns_listbox.pack(fill=tk.BOTH, expand=True, pady=5)

        action_frame = ttk.Frame(self.root, padding="10")
        action_frame.pack(fill=tk.X)
        self.calculate_button = ttk.Button(action_frame, text="3. คำนวณ Proximity Matrix", command=self.calculate_proximity_matrix)
        self.calculate_button.pack(side=tk.LEFT, padx=(0, 10))
        self.config_button = ttk.Button(action_frame, text="4. ตั้งค่า MDS", command=self.open_mds_config)
        self.config_button.pack(side=tk.LEFT, padx=(0, 10))
        self.export_report_button = ttk.Button(action_frame, text="5. Export Full MDS Report", command=self.export_mds_report, state=tk.DISABLED)
        self.export_report_button.pack(side=tk.LEFT, padx=(0, 10))
        self.about_button = ttk.Button(action_frame, text="คำแนะนำ (About)", command=self.show_about_window)
        self.about_button.pack(side=tk.LEFT, padx=(0, 10))

        result_frame = ttk.Frame(self.root, padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True)
        result_label = ttk.Label(result_frame, text="Proximity Matrix (Euclidean Distance):")
        result_label.pack(anchor=tk.W)
        self.tree = ttk.Treeview(result_frame, show='headings')
        self.tree.pack(fill=tk.BOTH, expand=True, pady=5)

    def open_mds_config(self):
        MDSConfigWindow(self)

    def show_about_window(self):
        AboutWindow(self.root)

    def select_all_columns(self):
        if self.columns_listbox.size() > 0: self.columns_listbox.selection_set(0, tk.END)

    def deselect_all_columns(self):
        if self.columns_listbox.size() > 0: self.columns_listbox.selection_clear(0, tk.END)

    def clear_results(self, clear_data=True):
        for item in self.tree.get_children(): self.tree.delete(item)
        self.tree["columns"] = []
        self.export_report_button.config(state=tk.DISABLED)
        if clear_data:
            self.proximity_df = None
            self.source_data_df = None

    def load_excel(self):
        file_path = filedialog.askopenfilename(title="เลือกไฟล์ Excel", filetypes=(("Excel Files", "*.xlsx *.xls"), ("All files", "*.*")))
        if not file_path: return
        try:
            self.df = pd.read_excel(file_path, index_col=0)
            self.file_label.config(text=file_path.split('/')[-1])
            self.columns_listbox.delete(0, tk.END)
            for col in self.df.columns:
                self.columns_listbox.insert(tk.END, col)
            self.clear_results()
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถโหลดไฟล์ได้:\n{e}")
            self.clear_results()

    def calculate_proximity_matrix(self):
        selected_indices = self.columns_listbox.curselection()
        if len(selected_indices) < 2:
            messagebox.showwarning("เลือกไม่ถูกต้อง", "กรุณาเลือกคอลัมน์อย่างน้อย 2 คอลัมน์")
            return
        selected_columns = [self.columns_listbox.get(i) for i in selected_indices]
        try:
            numeric_data = self.df[selected_columns].apply(pd.to_numeric, errors='coerce')
            if numeric_data.isnull().values.any():
                messagebox.showerror("ข้อมูลผิดพลาด", "ข้อมูลในคอลัมน์ที่เลือกมีบางเซลล์ที่ไม่ใช่ตัวเลข")
                self.clear_results()
                return

            self.source_data_df = numeric_data
            data = self.source_data_df.to_numpy()
            num_cols = len(selected_columns)
            dist_matrix = np.zeros((num_cols, num_cols))
            for i in range(num_cols):
                for j in range(i, num_cols):
                    dist = np.sqrt(np.sum((data[:, i] - data[:, j])**2))
                    dist_matrix[i, j] = dist
                    dist_matrix[j, i] = dist
            self.proximity_df = pd.DataFrame(dist_matrix, index=selected_columns, columns=selected_columns)

            self.display_matrix_in_gui(self.proximity_df)
            self.export_report_button.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการคำนวณ", f"รายละเอียด:\n{e}")
            self.clear_results()

    def display_matrix_in_gui(self, df):
        self.clear_results(clear_data=False)
        self.tree["columns"] = [""] + list(df.columns)
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.column("", anchor=tk.W, width=100)
        self.tree.heading("", text=" ", anchor=tk.W)
        for col in list(df.columns):
            self.tree.column(col, anchor=tk.CENTER, width=100)
            self.tree.heading(col, text=col, anchor=tk.CENTER)
        for index, row in df.iterrows():
            formatted_row = [f"{val:.3f}" for val in row.values]
            self.tree.insert("", tk.END, values=[index] + formatted_row)

    def export_mds_report(self):
        if self.proximity_df is None:
            messagebox.showwarning("ไม่มีข้อมูล", "กรุณากดคำนวณก่อนทำการ Export")
            return
        file_path = filedialog.asksaveasfilename(title="บันทึกรายงาน MDS", defaultextension=".xlsx", filetypes=(("Excel Files", "*.xlsx"),))
        if not file_path: return
        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                self.proximity_df.to_excel(writer, sheet_name='Proximity_Matrix', index=True)
                dims_to_run_str = self.mds_settings['dimensions']
                dims_to_run = [int(d.strip()) for d in dims_to_run_str.split(',') if d.strip()]
                for n_dim in dims_to_run:
                    self.create_mds_sheet(writer, n_dim)
            messagebox.showinfo("สำเร็จ", f"รายงาน MDS ฉบับสมบูรณ์ถูกบันทึกเรียบร้อยแล้วที่:\n{file_path}")
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ได้:\n{e}")

    def create_mds_sheet(self, writer, n_dim):
        mds = MDS(n_components=n_dim,
                  metric=self.mds_settings['metric'],
                  n_init=self.mds_settings['n_init'],
                  max_iter=self.mds_settings['max_iter'],
                  dissimilarity='precomputed',
                  random_state=self.mds_settings['random_state'],
                  normalized_stress='auto')
        coords = mds.fit_transform(self.proximity_df)
        labels = self.proximity_df.columns
        config_df = pd.DataFrame(coords, index=labels, columns=[f'Dim{i+1}' for i in range(n_dim)])
        new_distances_condensed = pdist(coords)
        new_distances_df = pd.DataFrame(squareform(new_distances_condensed), index=labels, columns=labels)
        if self.mds_settings['metric']:
            disparities_df = self.proximity_df.copy()
        else:
            disparities_df = pd.DataFrame(mds.disparities_, index=labels, columns=labels)
        residuals_df = disparities_df - new_distances_df
        upper_triangle_indices = np.triu_indices(len(labels), k=1)
        pairs = [f'{labels[i]} - {labels[j]}' for i, j in zip(*upper_triangle_indices)]
        comp_data = {
            'Pair': pairs,
            'Dissimilarity': self.proximity_df.values[upper_triangle_indices],
            'Disparity': disparities_df.values[upper_triangle_indices],
            'Distance': new_distances_df.values[upper_triangle_indices]
        }
        comparative_df = pd.DataFrame(comp_data)
        for col in ['Dissimilarity', 'Disparity', 'Distance']:
            comparative_df[f'Rank ({col})'] = comparative_df[col].rank().astype(int)
        fig, ax = plt.subplots(figsize=(6, 5))
        ax.scatter(comparative_df['Dissimilarity'], comparative_df['Distance'], facecolors='none', edgecolors='k', label='Distances')
        sorted_disp = comparative_df.sort_values('Dissimilarity')
        ax.plot(sorted_disp['Dissimilarity'], sorted_disp['Disparity'], 'k-o', markersize=4, label='Disparities')
        ax.set_xlabel('Dissimilarity')
        ax.set_ylabel('Disparity / Distance')
        ax.set_title(f"Shepard Diagram (Kruskal's Stress = {mds.stress_:.3f})")
        ax.legend()
        ax.grid(True)
        img_data = io.BytesIO()
        fig.savefig(img_data, format='png')
        plt.close(fig)
        sheet_name = f'MDS_{n_dim}D_Results'
        workbook = writer.book
        worksheet = workbook.add_worksheet(sheet_name)
        row_cursor = 0
        stress_info = pd.DataFrame([f"Kruskal's stress (1) = {mds.stress_:.4f}"], columns=[f"Results for a {n_dim}-dimensional space:"])
        stress_info.to_excel(writer, sheet_name=sheet_name, startrow=row_cursor, index=False)
        row_cursor += 4
        def write_df_to_sheet(df, title):
            nonlocal row_cursor
            worksheet.write(row_cursor, 0, title)
            df.to_excel(writer, sheet_name=sheet_name, startrow=row_cursor + 1)
            row_cursor += len(df) + 3
        write_df_to_sheet(config_df, 'Configuration:')
        write_df_to_sheet(new_distances_df, 'Distances measured in the representation space:')
        write_df_to_sheet(disparities_df, 'Disparities computed using the model:')
        write_df_to_sheet(residuals_df, 'Residual distances:')
        write_df_to_sheet(comparative_df, 'Comparative table:')
        worksheet.write(row_cursor, 0, 'Shepard diagram:')
        worksheet.insert_image(row_cursor + 1, 0, '', {'image_data': img_data})


def run_this_app(working_dir=None):
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน ProximityMatrixApp.
    """
    print(f"--- MDS_APP_INFO: Starting 'ProximityMatrixApp' via run_this_app() ---")
    try:
        root = tk.Tk()
        root.withdraw() # ซ่อนหน้าต่างทันทีที่สร้างขึ้น

        # สร้าง instance ของแอป (ซึ่งจะมีการเรียก center_window อยู่ข้างใน __init__)
        app = ProximityMatrixApp(root)

        root.deiconify() # แสดงหน้าต่างขึ้นมาหลังจากที่จัดตำแหน่งและเตรียมทุกอย่างเสร็จแล้ว
        root.mainloop()

        print(f"--- MDS_APP_INFO: ProximityMatrixApp mainloop finished. ---")
    except Exception as e:
        print(f"MDS_APP_ERROR: An error occurred during ProximityMatrixApp execution: {e}")
        if 'root' not in locals() or not root.winfo_exists():
            root_temp = tk.Tk()
            root_temp.withdraw()
            messagebox.showerror("Application Error", f"An unexpected error occurred:\n{e}", parent=root_temp)
            root_temp.destroy()
        else:
            messagebox.showerror("Application Error", f"An unexpected error occurred:\n{e}", parent=root)
        sys.exit(f"Error running ProximityMatrixApp: {e}")

if __name__ == "__main__":
    print("--- Running MDSApp.py directly for testing ---")
    run_this_app()
    print("--- Finished direct execution of MDSApp.py ---")