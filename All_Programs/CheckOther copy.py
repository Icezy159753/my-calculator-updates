import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pyreadstat
import re
from natsort import natsorted

# --- ฟังก์ชัน find_others_in_spss และ find_all_other_codes (เหมือนเดิม ไม่มีการเปลี่ยนแปลง) ---
def find_others_in_spss(spss_path, encoding=None):
    try:
        df_raw, meta = pyreadstat.read_sav(spss_path, apply_value_formats=False, encoding=encoding)
        search_terms, temp_results, group_order, all_vars = ['อื่นๆ', 'อื่น ๆ'], {}, {}, meta.column_names
        for i, var_name in enumerate(all_vars):
            if var_name not in meta.variable_value_labels or var_name not in df_raw.columns: continue
            value_labels, freq_counts = meta.variable_value_labels[var_name], df_raw[var_name].value_counts()
            for code, label in value_labels.items():
                if any(term in str(label).strip() for term in search_terms):
                    frequency = freq_counts.get(float(code), 0)
                    if frequency > 0:
                        display_group_name = var_name; pos = max(var_name.rfind('_O'), var_name.rfind('$'))
                        if pos != -1 and var_name[pos + 2:].isdigit(): display_group_name = var_name[:pos + 2]
                        group_key = (display_group_name, int(code))
                        if group_key not in temp_results:
                            last_var_in_group = display_group_name
                            if display_group_name.endswith(('_O', '$')):
                                group_vars = natsorted([v for v in all_vars if v.startswith(display_group_name)])
                                if group_vars: last_var_in_group = group_vars[-1]
                            next_variable = "N/A (End of File)"
                            try:
                                last_var_index = all_vars.index(last_var_in_group)
                                if last_var_index + 1 < len(all_vars): next_variable = all_vars[last_var_index + 1]
                            except ValueError: next_variable = "Error finding var"
                            temp_results[group_key] = {'label': label, 'frequency': 0, 'order': i, 'next_variable': next_variable}
                        temp_results[group_key]['frequency'] += int(frequency)
        final_results = []
        for group_key, data in temp_results.items():
            display_group_name, code = group_key
            final_results.append({'variable': display_group_name, 'next_variable': data['next_variable'], 'code': code, 'label': data['label'], 'frequency': data['frequency'], 'source_variable': display_group_name, 'order': data['order']})
        final_results.sort(key=lambda x: x['order'])
        return df_raw, meta, final_results
    except Exception as e:
        messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถอ่านไฟล์ SPSS ได้:\n{e}")
        return None, None, None

def find_all_other_codes(spss_path, encoding=None):
    try:
        _, meta = pyreadstat.read_sav(spss_path, apply_value_formats=True, encoding=encoding)
        search_terms, temp_results, all_vars = ['อื่นๆ', 'อื่น ๆ'], {}, meta.column_names
        for i, var_name in enumerate(all_vars):
            if var_name not in meta.variable_value_labels: continue
            for code, label in meta.variable_value_labels[var_name].items():
                if any(term in str(label).strip() for term in search_terms):
                    display_group_name = var_name; pos = max(var_name.rfind('_O'), var_name.rfind('$'))
                    if pos != -1 and var_name[pos + 2:].isdigit(): display_group_name = var_name[:pos + 2]
                    group_key = (display_group_name, int(code))
                    if group_key not in temp_results:
                        last_var_in_group = display_group_name
                        if display_group_name.endswith(('_O', '$')):
                            group_vars = natsorted([v for v in all_vars if v.startswith(display_group_name)])
                            if group_vars: last_var_in_group = group_vars[-1]
                        next_variable = "N/A (End of File)"
                        try:
                            last_var_index = all_vars.index(last_var_in_group)
                            if last_var_index + 1 < len(all_vars): next_variable = all_vars[last_var_index + 1]
                        except ValueError: next_variable = "Error finding var"
                        temp_results[group_key] = {'label': label, 'order': i, 'next_variable': next_variable}
        final_results = [{'variable': gk[0], 'next_variable': d['next_variable'], 'code': gk[1], 'label': d['label'], 'order': d['order']} for gk, d in temp_results.items()]
        final_results.sort(key=lambda x: x['order'])
        return final_results
    except Exception as e:
        messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถอ่านไฟล์ SPSS ได้:\n{e}")
        return None

# --- ส่วนของหน้าตาโปรแกรม (GUI) ---
class App:
    # --- ค่าคงที่สำหรับสีและฟอนต์ ---
    BG_COLOR = "#F0F2F5"
    FRAME_COLOR = "#FFFFFF"
    TEXT_COLOR = "#212121"
    BUTTON_TEXT_COLOR = "#FFFFFF"
    SUMMARY_COLOR = "#D32F2F"
    HIGHLIGHT_COLOR_FOUND = "#FFCDD2"
    ROW_ALT_COLOR = "#F5F5F5"
    
    PRIMARY_COLOR = "#005A9C"
    ACCENT_COLOR = "#0078D7"
    
    SUCCESS_COLOR = "#28a745"
    SUCCESS_DARK_COLOR = "#218838"
    
    DEFAULT_FONT = ("Segoe UI", 10)
    BOLD_FONT = ("Segoe UI", 10, "bold")
    TITLE_FONT = ("Segoe UI", 12, "bold")
    SUMMARY_FONT = ("Segoe UI", 11, "bold")

    def __init__(self, root):
        self.root = root
        self.root.title("โปรแกรม Check Codes_Other V1")
        
        # --- เพิ่มโค้ดสำหรับจัดหน้าจอให้อยู่ตรงกลาง ---
        window_width = 1100
        window_height = 650
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width / 2)
        center_y = int(screen_height/2 - window_height / 2)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        # --- สิ้นสุดส่วนที่เพิ่ม ---

        self.root.configure(bg=self.BG_COLOR)
        
        self.configure_styles()

        self.raw_df_spss = None; self.spss_path = tk.StringVar(); self.selected_encoding = tk.StringVar(value='Auto-Detect (None)')

        self.notebook = ttk.Notebook(root); self.notebook.pack(pady=15, padx=15, fill="both", expand=True)
        self.tab1 = ttk.Frame(self.notebook, padding="10"); self.tab2 = ttk.Frame(self.notebook, padding="10"); self.tab3 = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.tab1, text="  ค้นหา 'อื่นๆ' (มี Freq)  "); self.notebook.add(self.tab3, text="  'อื่นๆ' ทั้งหมด (All Others)  "); self.notebook.add(self.tab2, text="  ตรวจสอบ Raw Data (Excel)  ")

        self.create_spss_finder_tab()
        self.create_data_checker_tab()
        self.create_all_others_tab()

    def configure_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
            
        style.configure('.', font=self.DEFAULT_FONT, background=self.BG_COLOR, foreground=self.TEXT_COLOR)
        style.configure('TFrame', background=self.BG_COLOR)
        
        style.configure('TLabelFrame', background=self.FRAME_COLOR, borderwidth=1, relief="solid")
        style.configure('TLabelFrame.Label', font=self.TITLE_FONT, background=self.FRAME_COLOR, foreground=self.PRIMARY_COLOR)

        style.configure('Accent.TButton', font=self.BOLD_FONT, background=self.ACCENT_COLOR, foreground=self.BUTTON_TEXT_COLOR, borderwidth=0)
        style.map('Accent.TButton', background=[('active', self.PRIMARY_COLOR), ('pressed', self.PRIMARY_COLOR)], foreground=[('active', self.BUTTON_TEXT_COLOR)])
        
        style.configure('Success.TButton', font=self.BOLD_FONT, background=self.SUCCESS_COLOR, foreground=self.BUTTON_TEXT_COLOR, borderwidth=0)
        style.map('Success.TButton', background=[('active', self.SUCCESS_DARK_COLOR), ('pressed', self.SUCCESS_DARK_COLOR)], foreground=[('active', self.BUTTON_TEXT_COLOR)])

        style.configure('Treeview', rowheight=25, background=self.FRAME_COLOR, fieldbackground=self.FRAME_COLOR)
        style.configure('Treeview.Heading', font=self.BOLD_FONT, background=self.PRIMARY_COLOR, foreground=self.BUTTON_TEXT_COLOR, relief='flat')
        style.map('Treeview.Heading', background=[('active', self.ACCENT_COLOR)])
        
        style.configure('TNotebook', background=self.BG_COLOR, borderwidth=0)
        style.configure('TNotebook.Tab', font=self.BOLD_FONT, padding=[10, 5], background=self.BG_COLOR, borderwidth=1, relief='raised')
        style.map('TNotebook.Tab', background=[('selected', self.FRAME_COLOR)], foreground=[('selected', self.PRIMARY_COLOR)], relief=[('selected', 'flat')])
        
        style.configure('TCombobox', fieldbackground=self.FRAME_COLOR, selectbackground=self.FRAME_COLOR, selectforeground=self.TEXT_COLOR)

    def create_spss_finder_tab(self):
        file_frame = ttk.LabelFrame(self.tab1, text="ขั้นตอนที่ 1: เลือกไฟล์และตั้งค่า", padding=(15, 10))
        file_frame.pack(fill=tk.X, pady=5, ipady=5)
        spss_label = ttk.Label(file_frame, text="ไฟล์ SPSS (.sav):", background=self.FRAME_COLOR); spss_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        spss_entry = ttk.Entry(file_frame, textvariable=self.spss_path, width=70, font=self.DEFAULT_FONT); spss_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        browse_spss_button = ttk.Button(file_frame, text="เลือกไฟล์...", command=lambda: self.browse_file(self.spss_path, "SPSS")); browse_spss_button.grid(row=0, column=2, sticky="e", padx=5, pady=5)
        encoding_label = ttk.Label(file_frame, text="การเข้ารหัส (Encoding):", background=self.FRAME_COLOR); encoding_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        encoding_combo = ttk.Combobox(file_frame, textvariable=self.selected_encoding, state="readonly", width=30, font=self.DEFAULT_FONT)
        encoding_combo['values'] = ('Auto-Detect (None)', 'UTF-8', 'TIS-620', 'CP874 (Windows-874)'); encoding_combo.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        file_frame.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(self.tab1)
        button_frame.pack(pady=10)
        process_button = ttk.Button(button_frame, text="เริ่มการค้นหา (เฉพาะที่มีความถี่)", command=self.process_spss_data, style="Accent.TButton")
        process_button.pack(side=tk.LEFT, ipady=5, padx=(0, 5))
        save_button = ttk.Button(button_frame, text="บันทึกผลลัพธ์และข้อมูลดิบเป็นไฟล์ Excel...", command=self.save_results_to_excel, style="Success.TButton")
        save_button.pack(side=tk.LEFT, ipady=5, padx=(5, 0))

        result_frame = ttk.LabelFrame(self.tab1, text="ผลลัพธ์การค้นหา", padding=(15, 10))
        result_frame.pack(fill=tk.BOTH, expand=True)
        self.summary_label = ttk.Label(result_frame, text="", font=self.SUMMARY_FONT, foreground=self.SUMMARY_COLOR, background=self.FRAME_COLOR)
        self.summary_label.pack(anchor="w", pady=(0, 10))
        columns = ('variable', 'next_var', 'code', 'label', 'frequency')
        self.spss_tree = ttk.Treeview(result_frame, columns=columns, show='headings'); self.spss_tree.heading('variable', text='กลุ่มตัวแปร (Group)'); self.spss_tree.heading('next_var', text='ตัวแปรถัดไป (Next Var)'); self.spss_tree.heading('code', text='Code'); self.spss_tree.heading('label', text='ความหมาย (Label)'); self.spss_tree.heading('frequency', text='ความถี่ (Freq)'); self.spss_tree.column('variable', width=150, anchor=tk.W); self.spss_tree.column('next_var', width=150, anchor=tk.W); self.spss_tree.column('code', width=80, anchor=tk.CENTER); self.spss_tree.column('label', width=450, anchor=tk.W); self.spss_tree.column('frequency', width=120, anchor=tk.CENTER)
        v_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.spss_tree.yview); h_scrollbar = ttk.Scrollbar(result_frame, orient="horizontal", command=self.spss_tree.xview); self.spss_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set); v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y); h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X); self.spss_tree.pack(fill=tk.BOTH, expand=True); self.spss_tree.tag_configure('oddrow', background=self.ROW_ALT_COLOR); self.spss_tree.tag_configure('evenrow', background=self.FRAME_COLOR)

    def create_data_checker_tab(self):
        self.settings_path = tk.StringVar(); self.raw_data_path = tk.StringVar()
        file_frame = ttk.LabelFrame(self.tab2, text="ขั้นตอนที่ 1: เลือกไฟล์สำหรับตรวจสอบ", padding=(15, 10))
        file_frame.pack(fill=tk.X, pady=5, ipady=5)
        settings_label = ttk.Label(file_frame, text="ไฟล์ตั้งค่า (.xlsx):", background=self.FRAME_COLOR); settings_label.grid(row=0, column=0, sticky="w", padx=5, pady=5); settings_entry = ttk.Entry(file_frame, textvariable=self.settings_path, width=70, font=self.DEFAULT_FONT); settings_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5); browse_settings_button = ttk.Button(file_frame, text="เลือกไฟล์...", command=lambda: self.browse_file(self.settings_path, "Excel")); browse_settings_button.grid(row=0, column=2, sticky="e", padx=5, pady=5)
        raw_data_label = ttk.Label(file_frame, text="ไฟล์ Raw Data (.xlsx):", background=self.FRAME_COLOR); raw_data_label.grid(row=1, column=0, sticky="w", padx=5, pady=5); raw_data_entry = ttk.Entry(file_frame, textvariable=self.raw_data_path, width=70, font=self.DEFAULT_FONT); raw_data_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5); browse_raw_data_button = ttk.Button(file_frame, text="เลือกไฟล์...", command=lambda: self.browse_file(self.raw_data_path, "Excel")); browse_raw_data_button.grid(row=1, column=2, sticky="e", padx=5, pady=5)
        file_frame.columnconfigure(1, weight=1)
        check_button = ttk.Button(self.tab2, text="เริ่มการตรวจสอบ", command=self.start_checking, style="Accent.TButton")
        check_button.pack(pady=10, ipady=5)
        check_result_frame = ttk.LabelFrame(self.tab2, text="ผลการตรวจสอบ", padding=(15, 10)); check_result_frame.pack(fill=tk.BOTH, expand=True)
        check_columns = ('variable', 'code', 'label', 'new_frequency'); self.check_tree = ttk.Treeview(check_result_frame, columns=check_columns, show='headings'); self.check_tree.heading('variable', text='กลุ่มตัวแปร (Group)'); self.check_tree.heading('code', text='Code'); self.check_tree.heading('label', text='ความหมาย (Label)'); self.check_tree.heading('new_frequency', text='ความถี่ที่พบ (New Freq)'); self.check_tree.column('variable', width=150, anchor=tk.W); self.check_tree.column('code', width=80, anchor=tk.CENTER); self.check_tree.column('label', width=450, anchor=tk.W); self.check_tree.column('new_frequency', width=120, anchor=tk.CENTER)
        v_scrollbar_check = ttk.Scrollbar(check_result_frame, orient="vertical", command=self.check_tree.yview); h_scrollbar_check = ttk.Scrollbar(check_result_frame, orient="horizontal", command=self.check_tree.xview); self.check_tree.configure(yscrollcommand=v_scrollbar_check.set, xscrollcommand=h_scrollbar_check.set); v_scrollbar_check.pack(side=tk.RIGHT, fill=tk.Y); h_scrollbar_check.pack(side=tk.BOTTOM, fill=tk.X); self.check_tree.pack(fill=tk.BOTH, expand=True); self.check_tree.tag_configure('found', background=self.HIGHLIGHT_COLOR_FOUND)

    def create_all_others_tab(self):
        button_frame = ttk.Frame(self.tab3)
        button_frame.pack(pady=10)
        all_others_button = ttk.Button(button_frame, text="ค้นหารายการ 'อื่นๆ' ทั้งหมด (โดยไม่สนความถี่)", command=self.process_all_others_data, style="Accent.TButton")
        all_others_button.pack(side=tk.LEFT, ipady=5, padx=(0, 5))
        save_all_button = ttk.Button(button_frame, text="บันทึกรายการทั้งหมดเป็น Excel...", command=self.save_all_others_to_excel, style="Success.TButton")
        save_all_button.pack(side=tk.LEFT, ipady=5, padx=(5, 0))

        result_frame = ttk.LabelFrame(self.tab3, text="รายการ 'อื่นๆ' ทั้งหมดที่พบในไฟล์", padding=(15, 10))
        result_frame.pack(fill=tk.BOTH, expand=True)
        columns = ('variable', 'next_var', 'code', 'label'); self.all_others_tree = ttk.Treeview(result_frame, columns=columns, show='headings'); self.all_others_tree.heading('variable', text='กลุ่มตัวแปร (Group)'); self.all_others_tree.heading('next_var', text='ตัวแปรถัดไป (Next Var)'); self.all_others_tree.heading('code', text='Code'); self.all_others_tree.heading('label', text='ความหมาย (Label)'); self.all_others_tree.column('variable', width=150, anchor=tk.W); self.all_others_tree.column('next_var', width=150, anchor=tk.W); self.all_others_tree.column('code', width=80, anchor=tk.CENTER); self.all_others_tree.column('label', width=500, anchor=tk.W)
        v_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.all_others_tree.yview); h_scrollbar = ttk.Scrollbar(result_frame, orient="horizontal", command=self.all_others_tree.xview); self.all_others_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set); v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y); h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X); self.all_others_tree.pack(fill=tk.BOTH, expand=True); self.all_others_tree.tag_configure('oddrow', background=self.ROW_ALT_COLOR); self.all_others_tree.tag_configure('evenrow', background=self.FRAME_COLOR)
        
    # --- ส่วนของฟังก์ชันการทำงานของโปรแกรม (ไม่เปลี่ยนแปลง Logic) ---
    def browse_file(self, path_var, file_type):
        filetypes = {"SPSS": (("SPSS Data File", "*.sav"),), "Excel": (("Excel Files", "*.xlsx;*.xls"),)}.get(file_type, (("All files", "*.*"),)); filepath = filedialog.askopenfilename(title=f"เลือกไฟล์ {file_type}", filetypes=filetypes)
        if filepath: path_var.set(filepath)
    def process_spss_data(self):
        for item in self.spss_tree.get_children(): self.spss_tree.delete(item)
        spss_file = self.spss_path.get();
        if not spss_file: messagebox.showwarning("เลือกไฟล์", "กรุณาเลือกไฟล์ SPSS (.sav) ก่อน"); return
        encoding_choice = self.selected_encoding.get(); selected_enc = None if "Auto-Detect" in encoding_choice else encoding_choice.split(' ')[0]
        self.raw_df_spss, _, results = find_others_in_spss(spss_file, encoding=selected_enc)
        if results is None: return
        self.summary_label.config(text=f"พบรายการ 'อื่นๆ' ที่มีความถี่ > 0 จำนวน {len(results)} รายการ:")
        for i, res in enumerate(results):
            tag = 'oddrow' if i % 2 else 'evenrow'; values = (res['variable'], res['next_variable'], res['code'], res['label'], f"{res['frequency']} คำตอบ")
            self.spss_tree.insert('', tk.END, values=values, tags=(tag,))
    def process_all_others_data(self):
        for item in self.all_others_tree.get_children(): self.all_others_tree.delete(item)
        spss_file = self.spss_path.get()
        if not spss_file: messagebox.showwarning("เลือกไฟล์", "กรุณาเลือกไฟล์ SPSS (.sav) ที่ Tab แรกก่อน"); return
        encoding_choice = self.selected_encoding.get(); selected_enc = None if "Auto-Detect" in encoding_choice else encoding_choice.split(' ')[0]
        results = find_all_other_codes(spss_file, encoding=selected_enc)
        if results is None: return
        messagebox.showinfo("สำเร็จ", f"ค้นพบรายการ 'อื่นๆ' ทั้งหมด {len(results)} รายการ (รวมที่ไม่มีความถี่)")
        for i, res in enumerate(results):
            tag = 'oddrow' if i % 2 else 'evenrow'; values = (res['variable'], res['next_variable'], res['code'], res['label'])
            self.all_others_tree.insert('', tk.END, values=values, tags=(tag,))
    def _create_safe_sheet_name(self, name, max_len=31):
        safe_name = re.sub(r'[\\/*?:"<>|]', "", name); return safe_name[:max_len]
    def _find_next_oth_block(self, all_columns, start_index):
        oth_block = []
        for i in range(start_index + 1, len(all_columns)):
            current_var = all_columns[i]
            if 'oth' in str(current_var).lower(): oth_block.append(current_var)
            else: break
        return oth_block
    def save_results_to_excel(self):
        if not self.spss_tree.get_children(): messagebox.showwarning("ไม่มีข้อมูล", "กรุณาค้นหาข้อมูล (ที่มี Freq) ก่อนทำการบันทึก"); return
        if self.raw_df_spss is None: messagebox.showerror("ไม่มีข้อมูลดิบ", "ไม่พบข้อมูลดิบจากไฟล์ SPSS\nกรุณากดค้นหาข้อมูลอีกครั้ง"); return
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Workbook", "*.xlsx")], title="บันทึกไฟล์ตั้งค่าและข้อมูลดิบ")
        if not filepath: return
        writer = None
        try:
            summary_data = []
            for item_id in self.spss_tree.get_children():
                values = self.spss_tree.item(item_id, 'values')
                summary_data.append({'variable': values[0], 'next_variable': values[1], 'code': int(values[2]), 'label': values[3], 'frequency': int(values[4].split(' ')[0])})
            df_summary = pd.DataFrame(summary_data); df_summary['Go to'] = ''
            unique_groups = df_summary.drop_duplicates(subset=['variable']); grouped_summary = df_summary.groupby('variable')
            writer = pd.ExcelWriter(filepath, engine='xlsxwriter'); df_summary.to_excel(writer, sheet_name='Summary', index=False)
            workbook = writer.book; link_format = workbook.add_format({'color': 'blue', 'underline': 1})
            highlight_colors = ['#FFFF00', '#90EE90', '#ADD8E6', '#FFB6C1', '#FFA07A', '#DDA0DD']
            id_column = self.raw_df_spss.columns[0]; all_spss_columns = list(self.raw_df_spss.columns); summary_sheet = writer.sheets['Summary']
            for index, summary_row in unique_groups.iterrows():
                group_name = summary_row['variable']; group_df = grouped_summary.get_group(group_name); target_codes = group_df['code'].tolist()
                sheet_name = self._create_safe_sheet_name(group_name)
                rows_to_link = df_summary.index[df_summary['variable'] == group_name].tolist()
                for row_idx in rows_to_link:
                    summary_sheet.write_url(row_idx + 1, df_summary.columns.get_loc('Go to'), f"internal:'{sheet_name}'!A1", link_format, f"Go to {sheet_name}")
                vars_in_group = natsorted([c for c in all_spss_columns if str(c).startswith(group_name)]) if group_name.endswith(('_O', '$')) else [group_name]
                extra_oth_cols = []
                if vars_in_group:
                    try:
                        last_var_index = all_spss_columns.index(vars_in_group[-1])
                        extra_oth_cols = self._find_next_oth_block(all_spss_columns, last_var_index)
                    except ValueError: pass
                cols_to_show = list(dict.fromkeys([id_column] + vars_in_group + extra_oth_cols))
                mask = pd.Series(False, index=self.raw_df_spss.index)
                for var in vars_in_group:
                    if var in self.raw_df_spss.columns: mask |= self.raw_df_spss[var].isin(target_codes)
                detail_df = self.raw_df_spss.loc[mask, cols_to_show].copy()
                detail_df.to_excel(writer, sheet_name=sheet_name, index=False)
                detail_sheet = writer.sheets[sheet_name]; detail_sheet.write_url('A1', "internal:'Summary'!A1", link_format, 'Sbjnum')
                (max_row, max_col) = detail_df.shape
                for i, code_to_highlight in enumerate(target_codes):
                    color = highlight_colors[i % len(highlight_colors)]
                    cell_format = workbook.add_format({'bg_color': color})
                    detail_sheet.conditional_format(1, 1, max_row, max_col, {'type': 'cell', 'criteria': '==', 'value': code_to_highlight, 'format': cell_format})
            messagebox.showinfo("สำเร็จ", f"บันทึกผลลัพธ์และข้อมูลดิบลงในไฟล์:\n{filepath}")
        except Exception as e: messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ได้:\n{e}")
        finally:
            if writer: writer.close()
    def save_all_others_to_excel(self):
        if not self.all_others_tree.get_children(): messagebox.showwarning("ไม่มีข้อมูล", "กรุณาค้นหารายการ 'อื่นๆ' ทั้งหมดก่อนทำการบันทึก"); return
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Workbook", "*.xlsx")], title="บันทึกรายการ 'อื่นๆ' ทั้งหมด")
        if not filepath: return
        try:
            data_to_save = [{'variable': v[0], 'next_variable': v[1], 'code': int(v[2]), 'label': v[3]} for v in (self.all_others_tree.item(item_id, 'values') for item_id in self.all_others_tree.get_children())]
            pd.DataFrame(data_to_save).to_excel(filepath, index=False)
            messagebox.showinfo("สำเร็จ", f"บันทึกรายการทั้งหมดลงในไฟล์:\n{filepath}")
        except Exception as e: messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ได้:\n{e}")
    def start_checking(self):
        settings_file, raw_data_file = self.settings_path.get(), self.raw_data_path.get()
        if not settings_file or not raw_data_file: messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกทั้งไฟล์ตั้งค่าและไฟล์ Raw Data"); return
        for item in self.check_tree.get_children(): self.check_tree.delete(item)
        try:
            df_settings, df_raw = pd.read_excel(settings_file), pd.read_excel(raw_data_file)
            required_cols = ['variable', 'code']
            if not all(col in df_settings.columns for col in required_cols): messagebox.showerror("ไฟล์ไม่ถูกต้อง", f"ไฟล์ตั้งค่าต้องมีคอลัมน์: {', '.join(required_cols)}"); return
            for _, row in df_settings.iterrows():
                group_name, target_code, label = row['variable'], float(row['code']), row.get('label', '')
                new_freq = 0
                if group_name.endswith(('_O', '$')):
                    related_cols = [col for col in df_raw.columns if str(col).startswith(group_name)]
                    if related_cols: new_freq = (df_raw[related_cols] == target_code).any(axis=1).sum()
                elif group_name in df_raw.columns:
                    new_freq = (df_raw[group_name] == target_code).sum()
                tag = ('found',) if new_freq > 0 else ()
                self.check_tree.insert('', tk.END, values=(group_name, int(target_code), label, new_freq), tags=tag)
            messagebox.showinfo("สำเร็จ", "การตรวจสอบเสร็จสิ้น")
        except Exception as e: messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถประมวลผลไฟล์ได้:\n{e}")



# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None): # ชื่อฟังก์ชันนี้จะถูกใช้ใน Launcher
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน QuotaSamplerApp.
    """
    print(f"--- QUOTA_SAMPLER_INFO: Starting 'QuotaSamplerApp' via run_this_app() ---")
    try:
        # --- โค้ดที่ย้ายมาจาก if __name__ == "__main__": เดิมจะมาอยู่ที่นี่ ---
    #if __name__ == "__main__":
        root = tk.Tk()
        app = App(root)
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