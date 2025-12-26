# -*- coding: utf-8 -*-

# CHANGED: นำเข้าไลบรารีที่จำเป็น
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog
import pandas as pd
from openpyxl.styles import Font
import sys

class ModernExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("โปรแกรม Check Rotation Diary V1 (Bootstrap Fixed)")

        # --- การจัดหน้าต่างกลางจอภาพ ยังคงเดิม ---
        window_width = 900
        window_height = 700
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        position_x = int((screen_width / 2) - (window_width / 2))
        position_y = int((screen_height / 2) - (window_height / 2))
        self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
        
        self.file_path = None
        self.df_final = None
        self.df_raw = None

        self.export_icon_base64 = """
            R0lGODlhEAAQAMQAAORHVO5KV+5MX/FNX/JRY/JTY/NVa/Nak/Nav/VwZ/ZxaPZzaPZ/b/aAc/aDd/aIe/aNe/iQf/iTgf2ahv2ehv2ih/2qj/2uj/21k/2/mv3Co/3Gqf3Osv3auv3dygAAACH5BAEAABoALAAAAAAQABAAAAVv4C6QxGqAGeoVGIoAoSsMCIvA8YkFIsL4iUYsQJG4WARE4Ag2MAwYDALAhZrABJ4APqF4EGgYFACBAeAMHgoEDwLWaAEPAg+gUgGBAQYABgYIAYCAgIABAoEAnJ4GBAcadQoDgJ+ZBoUDjZgIAQA7
        """
        self.export_icon = ttk.PhotoImage(data=self.export_icon_base64)
        
        self.create_widgets()

    def create_widgets(self):
        # --- ส่วนของ Control Frame ยังคงเดิม ---
        control_frame = ttk.LabelFrame(self.root, text="ส่วนควบคุม", padding=15)
        control_frame.pack(fill=BOTH, padx=10, pady=10)
        control_frame.grid_columnconfigure(1, weight=1)
        
        self.btn_load = ttk.Button(control_frame, text="1. เลือกไฟล์ Excel", command=self.load_excel, bootstyle=PRIMARY, width=20)
        self.btn_load.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.lbl_file = ttk.Label(control_frame, text="ยังไม่ได้เลือกไฟล์", anchor="w", bootstyle="secondary")
        self.lbl_file.grid(row=0, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        lbl_sheet = ttk.Label(control_frame, text="เลือก Sheet:")
        lbl_sheet.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sheet_var = ttk.StringVar()
        self.sheet_menu = ttk.Combobox(control_frame, textvariable=self.sheet_var, state="disabled")
        self.sheet_menu.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        lbl_check_count = ttk.Label(control_frame, text="จำนวนชิ้นที่ตรวจสอบ:")
        lbl_check_count.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.check_count_var = ttk.StringVar(value="10")
        self.entry_check_count = ttk.Entry(control_frame, textvariable=self.check_count_var, width=10)
        self.entry_check_count.grid(row=2, column=1, padx=10, pady=5, sticky="w")
        action_frame = ttk.Frame(control_frame)
        action_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky="ew")
        action_frame.grid_columnconfigure(0, weight=1)
        action_frame.grid_columnconfigure(1, weight=1)

        self.btn_process = ttk.Button(action_frame, text="2. ประมวลผลข้อมูล", command=self.process_data, state="disabled", bootstyle=WARNING)
        self.btn_process.grid(row=0, column=0, padx=5, sticky="ew")
        
        self.btn_export = ttk.Button(
            action_frame, text="3. Export ผลลัพธ์", command=self.export_to_excel, 
            state="disabled", image=self.export_icon, compound=LEFT, bootstyle=SUCCESS
        )
        self.btn_export.image = self.export_icon
        self.btn_export.grid(row=0, column=1, padx=5, sticky="ew")
        
        # --- ส่วนของ Result Frame ---
        result_frame = ttk.LabelFrame(self.root, text="ผลลัพธ์", padding=15)
        result_frame.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))
        self.lbl_total_count = ttk.Label(result_frame, text="จำนวนข้อมูลทั้งหมด: 0 ชุด", font=("-size 10 -weight bold"))
        self.lbl_total_count.grid(row=0, column=0, sticky="w", pady=(0, 10))

        self.tree = ttk.Treeview(result_frame, show="headings")
        self.tree.bind("<Double-1>", self.show_raw_data_for_selection)

        style = ttk.Style()
        
        style.configure('Treeview.Heading', 
                        background=style.colors.primary, 
                        foreground='white',
                        font=('-size 10 -weight bold'))
        
        even_row_color = style.colors.get('light')
        odd_row_color = '#F0F0F0' 

        self.tree.tag_configure('evenrow', background=even_row_color)
        self.tree.tag_configure('oddrow', background=odd_row_color)
        self.tree.tag_configure('highlight', foreground=style.colors.get('danger'))

        vsb = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview, bootstyle="round")
        hsb = ttk.Scrollbar(result_frame, orient="horizontal", command=self.tree.xview, bootstyle="round")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        result_frame.grid_rowconfigure(1, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)
        self.tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")

    def process_data(self):
        try:
            n_check = int(self.check_count_var.get())
            if n_check <= 0:
                self.root.update_idletasks()
                Messagebox.show_error("ข้อมูลไม่ถูกต้อง", "จำนวนชิ้นที่ตรวจสอบต้องเป็นเลขจำนวนเต็มบวก", parent=self.root)
                return
        except ValueError:
            self.root.update_idletasks()
            Messagebox.show_error("ข้อมูลไม่ถูกต้อง", "กรุณากรอกจำนวนชิ้นที่ตรวจสอบเป็นตัวเลขเท่านั้น", parent=self.root)
            return
        if not self.file_path or not self.sheet_var.get():
            self.root.update_idletasks()
            Messagebox.show_warning("ข้อมูลไม่ครบ", "กรุณาเลือกไฟล์และ Sheet ก่อน", parent=self.root)
            return
        try:
            df = pd.read_excel(self.file_path, sheet_name=self.sheet_var.get(), header=None)
            header_row_index = -1
            for i, row in df.iterrows():
                if 'rotation' in [str(x).strip().lower() for x in row.values]:
                    header_row_index = i
                    break
            if header_row_index == -1:
                self.root.update_idletasks()
                Messagebox.show_error("หา Header ไม่พบ", "ไม่พบคอลัมน์ 'Rotation' ใน Sheet ที่เลือก", parent=self.root)
                return
            df.columns = df.iloc[header_row_index]
            df = df.iloc[header_row_index + 1:].reset_index(drop=True)
            df.columns = [str(c).strip() for c in df.columns]
            required_cols = ['ID', 'Rotation', 'No.', 'Sample No.']
            df = df[required_cols]
            df['ID'] = pd.to_numeric(df['ID'], errors='coerce')
            df['No.'] = pd.to_numeric(df['No.'], errors='coerce')
            df['Rotation'] = pd.to_numeric(df['Rotation'], errors='coerce')
            df = df.dropna(subset=['ID', 'Rotation'])
            df['ID'] = df['ID'].astype(int)
            df['Rotation'] = df['Rotation'].astype(int)
            self.df_raw = df.copy()

            def _summarize_qp(series):
                if series.empty: return ("", 0)
                s_cleaned = series.astype(str).str.upper().str.strip()
                q_count = s_cleaned.str.startswith('Q').sum()
                p_count = s_cleaned.str.startswith('P').sum()
                parts = []
                if q_count > 0: parts.append(f"Q={q_count}")
                if p_count > 0: parts.append(f"P={p_count}")
                summary_str = " ".join(parts)
                return (summary_str, len(s_cleaned))

            def summarize_rotation(series):
                if series.empty:
                    return ""
                counts = series.value_counts()
                if len(counts) <= 1:
                    return counts.index[0] if not counts.empty else ""
                else:
                    primary_rotation = counts.index[0]
                    other_rotations = counts.index[1:]
                    other_str = ", ".join(map(str, other_rotations))
                    return f"{primary_rotation} > Check Rotation {other_str}"

            self.df_final = self.df_raw.groupby('ID').agg(
                Rotation=('Rotation', summarize_rotation),
                No_Count=('No.', 'count'),
                First_N_Summary=('Sample No.', lambda s: _summarize_qp(s.head(n_check))),
                Last_N_Summary=('Sample No.', lambda s: _summarize_qp(s.tail(n_check)))
            ).reset_index()

            highlights = []
            for index, row in self.df_final.iterrows():
                first_summary, first_count = row['First_N_Summary']
                last_summary, last_count = row['Last_N_Summary']
                highlight = False
                if (first_count > 0 and first_count < n_check) or \
                   (last_count > 0 and last_count < n_check):
                    highlight = True
                if ' ' in first_summary or ' ' in last_summary:
                    highlight = True
                if isinstance(row['Rotation'], str) and 'Check Rotation' in row['Rotation']:
                    highlight = True
                highlights.append(highlight)
            self.df_final['_is_highlighted'] = highlights

            self.display_results()
            self.btn_export.config(state="normal")
        except KeyError as e:
            self.root.update_idletasks()
            Messagebox.show_error("คอลัมน์ไม่ถูกต้อง", f"ไม่พบคอลัมน์ที่จำเป็น: {e}\nกรุณาตรวจสอบชื่อคอลัมน์ในไฟล์ Excel (ID, Rotation, No., Sample No.)", parent=self.root)
        except Exception as e:
            self.root.update_idletasks()
            Messagebox.show_error("เกิดข้อผิดพลาด", f"เกิดข้อผิดพลาดระหว่างประมวลผล:\n{e}", parent=self.root)

    def display_results(self):
        try:
            n_check = int(self.check_count_var.get())
        except ValueError:
            n_check = 10
        total_count = len(self.df_final)
        self.lbl_total_count.config(text=f"จำนวนข้อมูลทั้งหมด: {total_count} ชุด")
        for i in self.tree.get_children():
            self.tree.delete(i)

        col_first_n = f'สรุป {n_check} ชิ้นแรก'
        col_last_n = f'สรุป {n_check} ชิ้นสุดท้าย'

        for index, row in self.df_final.iterrows():
            first_summary, _ = row['First_N_Summary']
            last_summary, _ = row['Last_N_Summary']

            base_tag = 'evenrow' if index % 2 == 0 else 'oddrow'

            if row['_is_highlighted']:
                tags = (base_tag, 'highlight')
            else:
                tags = (base_tag,)

            display_row = [row['ID'], row['Rotation'], row['No_Count'], first_summary, last_summary, tags]

            self.tree.insert("", "end", values=display_row[:-1], tags=display_row[-1])

        self.tree["columns"] = ['ID', 'Rotation', 'No.', col_first_n, col_last_n]
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            width = 80 if col in ['ID', 'Rotation', 'No.'] else 200
            self.tree.column(col, anchor="center", width=width)

    def export_to_excel(self):
        if self.df_final is None:
            self.root.update_idletasks()
            Messagebox.show_warning("ไม่มีข้อมูล", "ยังไม่มีข้อมูลให้ Export", parent=self.root)
            return
        try:
            n_check = int(self.check_count_var.get())
        except ValueError:
            n_check = 10
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="บันทึกผลลัพธ์เป็น")
        if not save_path:
            return
        try:
            df_to_export = self.df_final.drop(columns=['_is_highlighted'])
            df_to_export['First_N_Summary'] = df_to_export['First_N_Summary'].apply(lambda x: x[0])
            df_to_export['Last_N_Summary'] = df_to_export['Last_N_Summary'].apply(lambda x: x[0])
            df_to_export.rename(columns={
                'No_Count': 'No.',
                'First_N_Summary': f'Sample No. ({n_check} ชิ้นแรก)',
                'Last_N_Summary': f'Sample No. ({n_check} ชิ้นสุดท้าย)'
            }, inplace=True)
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                df_to_export.to_excel(writer, index=False, sheet_name='Results')
                workbook = writer.book
                worksheet = writer.sheets['Results']
                red_font = Font(color="FF0000")
                for row_idx, is_highlighted in enumerate(self.df_final['_is_highlighted'], start=2):
                    if is_highlighted:
                        for col_idx in range(1, worksheet.max_column + 1):
                            worksheet.cell(row=row_idx, column=col_idx).font = red_font
            self.root.update_idletasks()
            Messagebox.show_info("สำเร็จ", f"บันทึกไฟล์เรียบร้อยแล้วที่:\n{save_path}", parent=self.root)
        except Exception as e:
            self.root.update_idletasks()
            Messagebox.show_error("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ได้:\n{e}", parent=self.root)

    def show_raw_data_for_selection(self, event):
        if self.df_raw is None or not self.tree.selection():
            return
        selected_item = self.tree.selection()[0]
        row_values = self.tree.item(selected_item, 'values')
        clicked_id = int(row_values[0])

        raw_data_subset = self.df_raw[self.df_raw['ID'] == clicked_id].copy()

        detail_window = ttk.Toplevel(self.root)
        detail_window.title(f"ข้อมูลดิบ (Raw Data) สำหรับ ID: {clicked_id}")

        self.root.update_idletasks()
        main_x, main_y = self.root.winfo_x(), self.root.winfo_y()
        main_width, main_height = self.root.winfo_width(), self.root.winfo_height()
        detail_width, detail_height = 600, 400
        detail_pos_x = main_x + (main_width // 2) - (detail_width // 2)
        detail_pos_y = main_y + (main_height // 2) - (detail_height // 2)
        detail_window.geometry(f"{detail_width}x{detail_height}+{detail_pos_x}+{detail_pos_y}")

        detail_window.transient(self.root)
        detail_window.grab_set()

        detail_frame = ttk.Frame(detail_window, padding="10")
        detail_frame.pack(fill=BOTH, expand=True)

        detail_tree = ttk.Treeview(detail_frame, show="headings", bootstyle="info")
        columns = list(raw_data_subset.columns)
        detail_tree["columns"] = columns
        for col in columns:
            detail_tree.heading(col, text=col)
            detail_tree.column(col, anchor="center", width=100)
        for index, row in raw_data_subset.iterrows():
            detail_tree.insert("", "end", values=list(row))

        vsb = ttk.Scrollbar(detail_frame, orient="vertical", command=detail_tree.yview, bootstyle="round-info")
        hsb = ttk.Scrollbar(detail_frame, orient="horizontal", command=detail_tree.xview, bootstyle="round-info")
        detail_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        detail_frame.grid_rowconfigure(0, weight=1)
        detail_frame.grid_columnconfigure(0, weight=1)
        detail_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if not path:
            return
        self.file_path = path
        self.lbl_file.config(text=self.file_path.split('/')[-1], bootstyle="success")
        try:
            xls = pd.ExcelFile(self.file_path)
            sheet_names = xls.sheet_names
            self.sheet_menu['values'] = sheet_names
            self.sheet_menu.config(state="readonly")
            if sheet_names:
                self.sheet_menu.set(sheet_names[0])
            self.btn_process.config(state="normal")
        except Exception as e:
            self.root.update_idletasks()
            Messagebox.show_error("เกิดข้อผิดพลาด", f"ไม่สามารถอ่านไฟล์ Excel ได้:\n{e}", parent=self.root)

def run_app():
    """ฟังก์ชันสำหรับสร้างและรันแอปพลิเคชัน"""
    # ลองเปลี่ยนธีมได้ เช่น 'cosmo', 'litera', 'darkly', 'superhero', 'flatly'
    root = ttk.Window(themename="cosmo")
    app = ModernExcelProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    run_app()
