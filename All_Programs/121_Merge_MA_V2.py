import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import pandas as pd
import pyreadstat
import re
import os
import multiprocessing

class SpssProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("โปรแกรมดูดติดข้อ MA (SPSS/Excel) By Bell V1")

        # ตั้งขนาดและวางหน้าต่างกลางจอ
        window_width = 950
        window_height = 700
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

        self.df = None
        self.file_path = ""
        self.selected_sheet = None

        # หัวเรื่อง
        title_frame = ttk.Frame(self.root, bootstyle="dark")
        title_frame.pack(fill="x", pady=(0, 20))

        title_label = ttk.Label(
            title_frame,
            text="🔧 โปรแกรมดูดติดข้อ MA (SPSS/Excel)",
            font=('Arial', 18, 'bold'),
            bootstyle="inverse-dark"
        )
        title_label.pack(pady=20)

        # กรอบปุ่มหลัก
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=20, pady=10)

        self.btn_load = ttk.Button(
            button_frame,
            text="📂 โหลดไฟล์ (SPSS/Excel)",
            command=self.load_file,
            bootstyle="info",
            width=25
        )
        self.btn_load.pack(side="left", padx=5)

        self.btn_ma = ttk.Button(
            button_frame,
            text="🔀 รวมข้อ _O (MA)",
            command=self.process_ma_columns,
            state="disabled",
            bootstyle="success",
            width=25
        )
        self.btn_ma.pack(side="left", padx=5)

        self.btn_save = ttk.Button(
            button_frame,
            text="💾 บันทึกเป็น Excel (.xlsx)",
            command=self.save_to_excel,
            state="disabled",
            bootstyle="danger",
            width=25
        )
        self.btn_save.pack(side="left", padx=5)

        # กรอบตั้งค่า
        settings_frame = ttk.LabelFrame(
            self.root,
            text="⚙️ ตั้งค่า",
            bootstyle="primary",
            padding=15
        )
        settings_frame.pack(fill="x", padx=20, pady=10)

        # ตัวเลือกลบคอลัมน์
        self.delete_var = ttk.BooleanVar(value=True)
        self.chk_delete = ttk.Checkbutton(
            settings_frame,
            text="🗑️ ลบคอลัมน์เดิมหลังรวม (เช่น Q4_O1, Q4_O2 จะถูกลบออก)",
            variable=self.delete_var,
            bootstyle="primary-round-toggle"
        )
        self.chk_delete.grid(row=0, column=0, columnspan=4, sticky='w', pady=(0, 15))

        # ตัวเลือกตัวคั่น
        separator_label = ttk.Label(
            settings_frame,
            text="📝 ระบุตัวคั่นสำหรับรวมข้อมูล MA:",
            font=('Arial', 10, 'bold')
        )
        separator_label.grid(row=1, column=0, sticky='w', padx=(0, 10))

        self.separator_var = ttk.StringVar(value=',')

        separator_entry = ttk.Entry(
            settings_frame,
            textvariable=self.separator_var,
            font=('Arial', 11),
            width=10,
            justify='center',
            bootstyle="info"
        )
        separator_entry.grid(row=1, column=1, sticky='w', padx=5)

        hint_label = ttk.Label(
            settings_frame,
            text="(เช่น , หรือ | หรือ ; หรือช่องว่าง)",
            font=('Arial', 9),
            bootstyle="secondary"
        )
        hint_label.grid(row=1, column=2, sticky='w', padx=5)

        # สถานะไฟล์
        status_frame = ttk.Frame(self.root, bootstyle="secondary")
        status_frame.pack(fill="x", padx=20, pady=10)

        self.lbl_file_path = ttk.Label(
            status_frame,
            text="📁 ยังไม่ได้เลือกไฟล์",
            font=('Arial', 10),
            bootstyle="secondary",
            anchor="w"
        )
        self.lbl_file_path.pack(fill="x", padx=15, pady=10)

        # กรอบตาราง
        table_frame = ttk.LabelFrame(
            self.root,
            text="📊 ตัวอย่างข้อมูล (10 แถวแรก × 10 คอลัมน์แรก)",
            bootstyle="info",
            padding=10
        )
        table_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Treeview พร้อม Scrollbar
        tree_scroll_frame = ttk.Frame(table_frame)
        tree_scroll_frame.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(tree_scroll_frame, orient="vertical", bootstyle="info-round")
        hsb = ttk.Scrollbar(tree_scroll_frame, orient="horizontal", bootstyle="info-round")

        self.tree = ttk.Treeview(
            tree_scroll_frame,
            show='headings',
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            bootstyle="info"
        )

        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree.pack(side='left', fill='both', expand=True)

    def _select_sheet_dialog(self, sheet_names):
        dialog = ttk.Toplevel(self.root)
        dialog.title("เลือกชีต")

        # ตั้งขนาดและวางหน้าต่างกลางจอ
        dialog_width = 450
        dialog_height = 400
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        center_x = int((screen_width - dialog_width) / 2)
        center_y = int((screen_height - dialog_height) / 2)
        dialog.geometry(f"{dialog_width}x{dialog_height}+{center_x}+{center_y}")
        dialog.resizable(False, False)

        dialog.transient(self.root)
        dialog.grab_set()

        # หัวข้อ
        header_frame = ttk.Frame(dialog, bootstyle="primary")
        header_frame.pack(fill="x")

        ttk.Label(
            header_frame,
            text="📋 เลือกชีตที่ต้องการ",
            font=('Arial', 16, 'bold'),
            bootstyle="inverse-primary"
        ).pack(pady=20)

        ttk.Label(
            dialog,
            text="ไฟล์ Excel นี้มีหลายชีต\nกรุณาเลือกชีตที่ต้องการประมวลผล:",
            font=('Arial', 11)
        ).pack(pady=20)

        listbox_frame = ttk.Frame(dialog)
        listbox_frame.pack(padx=25, pady=10, fill="both", expand=True)

        scrollbar = ttk.Scrollbar(listbox_frame, bootstyle="primary-round")
        scrollbar.pack(side="right", fill="y")

        from tkinter import Listbox
        listbox = Listbox(
            listbox_frame,
            height=10,
            font=('Arial', 11),
            yscrollcommand=scrollbar.set,
            selectmode="single",
            relief="flat",
            bd=2,
            highlightthickness=1
        )
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        for name in sheet_names:
            listbox.insert("end", name)
        listbox.selection_set(0)
        listbox.focus_set()

        self.selected_sheet = None

        def on_ok():
            try:
                selected_index = listbox.curselection()[0]
                self.selected_sheet = listbox.get(selected_index)
                dialog.destroy()
            except IndexError:
                messagebox.showwarning("ยังไม่ได้เลือก", "กรุณาเลือกชีตก่อน", parent=dialog)

        def on_double_click(event):
            on_ok()

        listbox.bind("<Double-Button-1>", on_double_click)
        listbox.bind("<Return>", lambda e: on_ok())

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)

        ok_button = ttk.Button(
            button_frame,
            text="✓ เลือกชีตนี้",
            command=on_ok,
            bootstyle="success",
            width=20
        )
        ok_button.pack()

        self.root.wait_window(dialog)

        return self.selected_sheet

    def _clean_dataframe(self, df):
        """
        ทำความสะอาดข้อความในทุกคอลัมน์ที่มีแนวโน้มเป็นข้อความ
        - แปลง bytes -> utf-8 string
        - แปลง dtype เป็น pandas StringDtype ก่อนใช้ .str
        - คงค่า NaN ไว้ ไม่แปลงเป็น "nan"
        """
        import pandas as pd
        from pandas.api.types import is_object_dtype, is_string_dtype

        for col in df.columns:
            s = df[col]

            # เอาเฉพาะคอลัมน์ที่เป็น object หรือ string dtype เท่านั้น
            if is_object_dtype(s) or is_string_dtype(s):
                # บางทีจาก .sav จะได้ค่าเป็น bytes -> แปลงเป็น str ก่อน
                s = s.apply(lambda x: x.decode('utf-8', 'ignore') if isinstance(x, (bytes, bytearray)) else x)

                # แปลงเป็น pandas StringDtype (คง NaN เป็น <NA>)
                s = s.astype("string")

                # ทำความสะอาดเฉพาะค่าที่ไม่ว่าง
                s = (s.str.replace('_x000D_', '', regex=False)
                    .str.replace('\r', '', regex=False)
                    .str.replace('\n', ' ', regex=False))

                df[col] = s  # ใส่กลับหลังทำความสะอาด
        return df

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ SPSS หรือ Excel",
            filetypes=(
                ("Data Files", "*.sav *.xlsx *.xls"),
                ("SPSS Files", "*.sav"),
                ("Excel Files", "*.xlsx *.xls"),
                ("All files", "*.*")
            )
        )
        if not file_path:
            return

        try:
            file_extension = os.path.splitext(file_path)[1].lower()
            loaded_df = None

            if file_extension == '.sav':
                df, meta = pyreadstat.read_sav(file_path, apply_value_formats=False)
                loaded_df = df

            elif file_extension in ['.xlsx', '.xls']:
                xls = pd.ExcelFile(file_path)
                sheet_names = xls.sheet_names

                chosen_sheet = None
                if len(sheet_names) == 1:
                    chosen_sheet = sheet_names[0]
                else:
                    chosen_sheet = self._select_sheet_dialog(sheet_names)

                if chosen_sheet:
                    loaded_df = pd.read_excel(file_path, sheet_name=chosen_sheet)
                else:
                    self.lbl_file_path.config(text="📁 ยกเลิกการโหลดไฟล์")
                    return

            else:
                messagebox.showerror("ไฟล์ไม่รองรับ", f"ไม่รองรับไฟล์ประเภท: {file_extension}")
                return

            self.df = self._clean_dataframe(loaded_df)

            self.file_path = file_path
            self.lbl_file_path.config(
                text=f"📁 ไฟล์ที่เลือก: {os.path.basename(self.file_path)}",
                bootstyle="success"
            )
            messagebox.showinfo("สำเร็จ", f"โหลดไฟล์ '{os.path.basename(self.file_path)}' สำเร็จ!")
            self.update_treeview()

            self.btn_ma.config(state="normal")
            self.btn_save.config(state="normal")

        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการโหลดไฟล์", f"ไม่สามารถโหลดไฟล์ได้:\n{e}")

    def process_ma_columns(self):
        if self.df is None:
            messagebox.showwarning("คำเตือน", "กรุณาโหลดไฟล์ก่อน")
            return

        try:
            # ดึงค่าตัวคั่นที่ผู้ใช้กรอก
            separator = self.separator_var.get()
            if not separator:
                separator = ','  # ใช้ค่าเริ่มต้นถ้าไม่ได้กรอก

            o_cols = [col for col in self.df.columns if re.search(r'_O\d+$', col)]

            if not o_cols:
                messagebox.showinfo("ไม่พบข้อมูล", "ไม่พบคอลัมน์ที่ต้องรวม (เช่น Q1_O1, Q1_O2)")
                return

            groups = {}
            for col in o_cols:
                prefix = re.sub(r'\d+$', '', col)
                if prefix not in groups:
                    groups[prefix] = []
                groups[prefix].append(col)

            new_cols_count = 0
            for prefix, cols_to_join in groups.items():
                last_col_in_group = cols_to_join[-1]
                insert_location = self.df.columns.get_loc(last_col_in_group) + 1

                def join_without_decimal(row):
                    # .dropna() จะจัดการกับค่าว่าง (NaN) ให้อยู่แล้ว
                    values_to_join = []
                    for v in row.dropna():
                        s_val = str(v)
                        if isinstance(v, float) and v.is_integer():
                            values_to_join.append(str(int(v)))
                        else:
                            # การทำความสะอาดหลักเกิดขึ้นที่ _clean_dataframe แล้ว
                            # ที่นี่อาจไม่ต้องทำซ้ำ แต่มีไว้ก็ไม่เสียหาย
                            cleaned_val = s_val.replace('_x000D_', '').replace('\r', '').replace('\n', ' ')
                            values_to_join.append(cleaned_val)
                    return separator.join(values_to_join)

                combined_series = self.df[cols_to_join].apply(join_without_decimal, axis=1)

                self.df.insert(
                    loc=insert_location,
                    column=prefix,
                    value=combined_series
                )

                new_cols_count += 1

            if self.delete_var.get():
                cols_to_drop = [col for group in groups.values() for col in group]
                self.df.drop(columns=cols_to_drop, inplace=True)

            self.update_treeview()

            # แสดงตัวคั่นที่ใช้ในข้อความแจ้งเตือน
            sep_display = separator if separator != ' ' else '[ช่องว่าง]'
            messagebox.showinfo(
                "สำเร็จ",
                f"รวมข้อมูล {new_cols_count} คอลัมน์ใหม่เรียบร้อยแล้ว!\n"
                f"ตัวคั่นที่ใช้: {sep_display}"
            )

        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"เกิดข้อผิดพลาดระหว่างรวมคอลัมน์:\n{e}")

    def save_to_excel(self):
        if self.df is None:
            messagebox.showwarning("คำเตือน", "ไม่มีข้อมูลสำหรับบันทึก")
            return

        save_path = filedialog.asksaveasfilename(
            title="บันทึกเป็นไฟล์ Excel",
            defaultextension=".xlsx",
            filetypes=(("Excel Files", "*.xlsx"), ("All files", "*.*"))
        )
        if not save_path:
            return

        try:
            # ไม่จำเป็นต้องทำความสะอาดข้อมูลซ้ำอีก
            # เพราะ self.df สะอาดตั้งแต่ตอนโหลดแล้ว
            self.df.to_excel(save_path, index=False)
            messagebox.showinfo("สำเร็จ", f"บันทึกไฟล์ Excel เรียบร้อยแล้วที่:\n{save_path}")
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ได้:\n{e}")

    def update_treeview(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        if self.df is None:
            return

        # แสดงเฉพาะ 10 คอลัมน์แรก
        display_columns = list(self.df.columns[:10])
        self.tree["columns"] = display_columns

        for col in display_columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor='w')

        # แสดงเฉพาะ 10 แถวแรก
        df_head = self.df.head(10)
        for index, row in df_head.iterrows():

            def format_for_display(value):
                # ฟังก์ชันนี้จัดการกับ pd.isna() ได้อย่างถูกต้องอยู่แล้ว
                if pd.isna(value):
                    return ""
                if isinstance(value, float) and value.is_integer():
                    return str(int(value))
                return str(value)

            # แสดงเฉพาะค่าของ 10 คอลัมน์แรก
            formatted_values = [format_for_display(v) for v in row[:10].tolist()]
            self.tree.insert("", "end", values=formatted_values)




# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None): # ชื่อฟังก์ชันนี้จะถูกใช้ใน Launcher
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน QuotaSamplerApp.
    """
    print(f"--- QUOTA_SAMPLER_INFO: Starting 'QuotaSamplerApp' via run_this_app() ---")
    try:
    # --- ส่วนที่ใช้รันโปรแกรม ---
    #if __name__ == "__main__":
        multiprocessing.freeze_support()

        root = ttk.Window(themename="cosmo")  # ใช้ธีม cosmo (สีฟ้า-สวย)
        app = SpssProcessorApp(root)
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