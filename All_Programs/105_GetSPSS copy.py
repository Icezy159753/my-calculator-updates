# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import pyreadstat
import pandas as pd
import re # <<< เพิ่ม: import re สำหรับการเรียงลำดับ

# ==============================================================================
#  ส่วนที่ 1: โค้ดสำหรับฟังก์ชันเรียงลำดับตัวแปร
# ==============================================================================

def natural_sort_key(s):
    """
    แยกสตริงออกเป็นส่วนของตัวอักษรและตัวเลข แล้วแปลงเป็นคีย์ที่เสถียร
    """
    parts = re.split(r'(\d+)', s)
    key = []
    for part in parts:
        if not part: continue
        if part.isdigit():
            key.append((1, int(part)))
        else:
            key.append((0, part.lower()))
    return key

def get_sort_key_for_vars(var_name):
    """
    ฟังก์ชันสร้าง "กุญแจ" สำหรับเรียงลำดับตัวแปรโดยเฉพาะ
    """
    parts = str(var_name).split('_')
    
    if len(parts) < 3 or not parts[0].isalpha():
        return ((2, var_name),) 

    try:
        respondent_id = int(parts[1])
    except (ValueError, IndexError):
        return ((2, var_name),)

    question_part_str = '_'.join(parts[2:])

    match = re.search(r'(_O\d+)', question_part_str)
    if match:
        o_part_str = match.group(1)
        stem_str = question_part_str[:match.start()]
        return (natural_sort_key(stem_str), respondent_id, natural_sort_key(o_part_str))
    else:
        return (natural_sort_key(question_part_str), respondent_id, [])

# ==============================================================================
#  ส่วนที่ 2: คลาสโปรแกรมหลัก SpssVariableEditor
# ==============================================================================

class SpssVariableEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Get Syntax from SPSS Variables V2")
        self.root.geometry("950x600") # ขยายความกว้างเล็กน้อย
        self.root.configure(background='white')

        self.df = None
        self.meta = None
        self.variables_data = [] 
        self.original_filepath = None

        # --- การตั้งค่า Style ---
        style = ttk.Style()
        style.theme_use("clam")

        COLOR_GREEN = '#28a745'; COLOR_GREEN_ACTIVE = '#218838'
        COLOR_BLUE = '#007bff'; COLOR_BLUE_ACTIVE = '#0069d9'
        COLOR_RED = '#dc3545'; COLOR_RED_ACTIVE = '#c82333'
        COLOR_ORANGE = '#fd7e14'; COLOR_ORANGE_ACTIVE = '#e67311' # <<< เพิ่ม: สีส้มสำหรับปุ่ม Sort
        
        style.configure("White.TFrame", background="white")
        style.configure("TLabel", background="white", font=("Segoe UI", 9))
        
        style.configure('Green.TButton', foreground='white', background=COLOR_GREEN, font=('Segoe UI', 9, 'bold'), borderwidth=0)
        style.map('Green.TButton', background=[('active', COLOR_GREEN_ACTIVE)])
        style.configure('Blue.TButton', foreground='white', background=COLOR_BLUE, font=('Segoe UI', 10, 'bold'), padding=5)
        style.map('Blue.TButton', background=[('active', COLOR_BLUE_ACTIVE)])
        style.configure('Red.TButton', foreground='white', background=COLOR_RED, font=('Segoe UI', 9, 'bold'))
        style.map('Red.TButton', background=[('active', COLOR_RED_ACTIVE)])
        style.configure('Orange.TButton', foreground='white', background=COLOR_ORANGE, font=('Segoe UI', 9, 'bold'))
        style.map('Orange.TButton', background=[('active', COLOR_ORANGE_ACTIVE)])
        style.configure("White.Vertical.TScrollbar", troughcolor='white', background='lightgrey', bordercolor='white')
        style.configure("White.Horizontal.TScrollbar", troughcolor='white', background='lightgrey', bordercolor='white')
        
        GRID_LINE_COLOR = "#DCDCDC"; HEADER_BG = "#EFEFEF"; ODD_ROW_BG = "#F0F8FF"
        EVEN_ROW_BG = "#FFFFFF"; SELECT_BG = "#0078D7"; SELECT_FG = "white"
        style.configure("Treeview.Heading", background=HEADER_BG, foreground="black", relief="flat", font=("Segoe UI", 10, "bold"))
        style.map("Treeview.Heading", relief=[('active','groove'),('pressed','sunken')])
        style.map('Treeview', background=[('selected', SELECT_BG)], foreground=[('selected', SELECT_FG)])
        style.configure("Treeview", rowheight=25, background=GRID_LINE_COLOR, fieldbackground=GRID_LINE_COLOR, font=("Segoe UI", 9))
        
        # --- สร้างส่วนของ UI ---
        control_frame = ttk.Frame(self.root, padding="10", style="White.TFrame")
        control_frame.pack(fill=tk.X)

        self.load_button = ttk.Button(control_frame, text="โหลดไฟล์ SPSS", command=self.load_file, style='Green.TButton')
        self.load_button.pack(side=tk.LEFT, padx=5)
        
        # <<< จุดที่แก้ไข: ลบ Combobox สำหรับเลือก Encoding ออก >>>
        # encoding_label = ttk.Label(control_frame, text="Encoding:")
        # encoding_label.pack(side=tk.LEFT, padx=(10, 2))
        # self.encoding_combo = ttk.Combobox(control_frame, values=['cp874', 'tis-620', 'utf-8'], width=10, state="readonly")
        # self.encoding_combo.set('cp874'); self.encoding_combo.pack(side=tk.LEFT, padx=2)

        # --- กรอบสำหรับปุ่มจัดการตัวแปร ---
        manipulation_frame = ttk.Frame(control_frame, style="White.TFrame")
        manipulation_frame.pack(side=tk.LEFT, padx=(15, 2))

        self.move_up_button = ttk.Button(manipulation_frame, text="▲", command=self.move_up, state=tk.DISABLED, style='Blue.TButton', width=3)
        self.move_up_button.pack(side=tk.LEFT, padx=2)
        self.move_down_button = ttk.Button(manipulation_frame, text="▼", command=self.move_down, state=tk.DISABLED, style='Blue.TButton', width=3)
        self.move_down_button.pack(side=tk.LEFT, padx=2)
        
        self.sort_button = ttk.Button(manipulation_frame, text="Sort loop", command=self.sort_variables, state=tk.DISABLED, style='Orange.TButton')
        self.sort_button.pack(side=tk.LEFT, padx=(10,2))
        
        self.save_button = ttk.Button(control_frame, text="สร้าง Syntax...", command=self.save_syntax_file, state=tk.DISABLED, style='Red.TButton')
        self.save_button.pack(side=tk.RIGHT, padx=5)
        
        instruction_label = ttk.Label(control_frame, text="เลือกแถวแล้วกด 'Delete' เพื่อลบ / ดับเบิลคลิกเพื่อดู Values")
        instruction_label.pack(side=tk.LEFT, padx=20)

        # --- สร้างตารางแสดงผล (Treeview) ---
        tree_frame = ttk.Frame(self.root, padding="10", style="White.TFrame")
        tree_frame.pack(fill=tk.BOTH, expand=True)
        columns = ("#", "Name", "Type", "Label", "Values")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("#", text="ลำดับ"); self.tree.heading("Name", text="Name", anchor='w'); self.tree.heading("Type", text="Type"); self.tree.heading("Label", text="Label", anchor='w'); self.tree.heading("Values", text="Values", anchor='w')
        self.tree.column("#", width=50, anchor=tk.CENTER, stretch=False); self.tree.column("Name", width=150, anchor='w', stretch=True); self.tree.column("Type", width=80, anchor=tk.CENTER, stretch=False); self.tree.column("Label", width=300, anchor='w', stretch=True); self.tree.column("Values", width=300, anchor='w', stretch=True)
        self.tree.tag_configure('oddrow', background=ODD_ROW_BG); self.tree.tag_configure('evenrow', background=EVEN_ROW_BG)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview, style="White.Vertical.TScrollbar")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview, style="White.Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y); hsb.pack(side=tk.BOTTOM, fill=tk.X); self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Delete>", self.delete_selected); self.tree.bind("<<TreeviewSelect>>", self.on_selection_change); self.tree.bind("<Double-1>", self.show_value_labels)

    # <<< จุดที่แก้ไข: ปรับปรุงฟังก์ชัน load_file ทั้งหมด >>>
    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ SPSS",
            filetypes=(("SPSS Data File", "*.sav"), ("All files", "*.*"))
        )
        if not file_path:
            return

        self.original_filepath = file_path
        
        # รายการ Encoding ที่จะใช้ทดลองเปิดไฟล์ (เรียงจากที่พบบ่อยสำหรับภาษาไทย)
        encodings_to_try = ['cp874', 'tis-620', 'utf-8']
        
        df_loaded = None
        meta_loaded = None
        successful_encoding = None

        # วนลูปเพื่อทดลองเปิดไฟล์ด้วย Encoding แต่ละตัว
        for encoding in encodings_to_try:
            try:
                # พยายามอ่านไฟล์ด้วย encoding ปัจจุบัน
                df_loaded, meta_loaded = pyreadstat.read_sav(
                    self.original_filepath,
                    encoding=encoding,
                    apply_value_formats=False,
                    dates_as_pandas_datetime=True
                )
                # ถ้าสำเร็จ ให้เก็บชื่อ encoding ไว้ แล้วออกจาก loop
                successful_encoding = encoding
                print(f"File loaded successfully with encoding: {successful_encoding}")
                break
            except Exception as e:
                # ถ้าไม่สำเร็จ ให้ข้ามไปลองตัวถัดไป (อาจจะ print error ไว้ดูใน console)
                print(f"Failed to load with encoding '{encoding}': {e}")
                continue
        
        # ตรวจสอบว่าโหลดไฟล์สำเร็จหรือไม่
        if successful_encoding:
            # ถ้าสำเร็จ, ดำเนินการต่อเหมือนเดิม
            self.df = df_loaded
            self.meta = meta_loaded
            
            self.variables_data.clear()
            for i, name in enumerate(self.meta.column_names):
                display_type = "Numeric"
                if 'datetime' in str(self.df[name].dtype.name).lower():
                    display_type = "Date"
                elif hasattr(self.meta, 'readstat_variable_types') and self.meta.readstat_variable_types.get(name, '').upper().startswith('A'):
                    display_type = "String"
                
                self.variables_data.append({
                    'name': name,
                    'display_type': display_type,
                    'label': self.meta.column_labels[i] if self.meta.column_labels else "",
                    'values': self.meta.variable_value_labels.get(name, {})
                })
                
            self.populate_treeview()
            self.save_button['state'] = tk.NORMAL
            self.sort_button['state'] = tk.NORMAL
            messagebox.showinfo(
                "สำเร็จ",
                f"โหลดไฟล์ {os.path.basename(self.original_filepath)} เรียบร้อยแล้ว\n(ใช้ Encoding: {successful_encoding})"
            )
        else:
            # ถ้าลองทุกตัวแล้วยังไม่สำเร็จ ให้แสดงข้อความแจ้งเตือน
            messagebox.showerror(
                "เกิดข้อผิดพลาด",
                f"ไม่สามารถอ่านไฟล์ SPSS ได้\n"
                f"ได้ทดลองใช้ Encoding ต่อไปนี้แล้ว แต่ไม่สำเร็จ:\n"
                f"{', '.join(encodings_to_try)}\n\n"
                "กรุณาตรวจสอบไฟล์ หรือลองหา Encoding ที่ถูกต้อง"
            )

    def sort_variables(self):
        if not self.variables_data:
            messagebox.showwarning("คำเตือน", "ยังไม่มีข้อมูลตัวแปรให้เรียงลำดับ")
            return
        
        try:
            original_vars_data = self.variables_data

            def is_sortable(var_info):
                var_name = var_info['name']
                key = get_sort_key_for_vars(var_name)
                return key[0][0] != 2

            final_ordered_list = []
            current_sortable_block = []
            
            for var_info in original_vars_data:
                if is_sortable(var_info):
                    current_sortable_block.append(var_info)
                else:
                    if current_sortable_block:
                        sorted_block = sorted(current_sortable_block, key=lambda v: get_sort_key_for_vars(v['name']))
                        final_ordered_list.extend(sorted_block)
                        current_sortable_block = []
                    
                    final_ordered_list.append(var_info)

            if current_sortable_block:
                sorted_block = sorted(current_sortable_block, key=lambda v: get_sort_key_for_vars(v['name']))
                final_ordered_list.extend(sorted_block)

            self.variables_data = final_ordered_list
            self.populate_treeview()
            messagebox.showinfo("สำเร็จ", "เรียงลำดับตัวแปรเรียบร้อยแล้ว")
        
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดขณะเรียงลำดับตัวแปร:\n{e}")
            import traceback
            traceback.print_exc()

    def populate_treeview(self):
        selection = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        for i, var_info in enumerate(self.variables_data):
            tag = 'oddrow' if (i + 1) % 2 != 0 else 'evenrow'
            values_str = str(var_info['values']) if var_info['values'] else "None"
            if len(values_str) > 100: values_str = values_str[:100] + "..."
            self.tree.insert("", tk.END, iid=i, values=(i + 1, var_info['name'], var_info['display_type'], var_info['label'], values_str), tags=(tag,))
        if selection:
            try: self.tree.selection_set(selection)
            except tk.TclError: pass
        self.on_selection_change()

    def show_value_labels(self, event):
        selected_item = self.tree.focus()
        if not selected_item: return
        index = int(selected_item)
        var_info = self.variables_data[index]
        
        popup = tk.Toplevel(self.root); popup.title(f"Value Labels for: {var_info['name']}")
        popup.geometry("400x300"); popup.transient(self.root); popup.grab_set()
        popup.configure(background='white')
        text_frame = ttk.Frame(popup, padding="10", style="White.TFrame")
        text_frame.pack(expand=True, fill=tk.BOTH)
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Segoe UI", 10), background='white', borderwidth=0, relief='flat')
        text_widget.pack(expand=True, fill=tk.BOTH, side=tk.LEFT)
        text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview, style="White.Vertical.TScrollbar")
        text_scrollbar.pack(fill=tk.Y, side=tk.RIGHT); text_widget.config(yscrollcommand=text_scrollbar.set)
        formatted_text = f"Variable: {var_info['name']}\n" + "=" * 30 + "\n\n"
        if not var_info['values']: formatted_text += "No defined value labels."
        else:
            for key, value in var_info['values'].items(): formatted_text += f"{float(key):g} : {value}\n"
        text_widget.insert(tk.END, formatted_text); text_widget.config(state=tk.DISABLED)
        button_frame = ttk.Frame(popup, padding=(0, 0, 0, 10), style="White.TFrame")
        button_frame.pack(fill=tk.X); close_button = ttk.Button(button_frame, text="Close", command=popup.destroy); close_button.pack()
        self.root.wait_window(popup)

    def delete_selected(self, event=None):
        selected_items = self.tree.selection()
        if not selected_items: messagebox.showwarning("คำเตือน", "กรุณาเลือกแถวที่ต้องการลบก่อน"); return
        indices_to_delete = sorted([int(item) for item in selected_items], reverse=True)
        for index in indices_to_delete: del self.variables_data[index]
        self.populate_treeview(); self.tree.selection_set()
        messagebox.showinfo("สำเร็จ", f"ลบตัวแปรจำนวน {len(selected_items)} รายการแล้ว")

    def move_up(self):
        selected_items = self.tree.selection();
        if not selected_items: return
        indices = sorted([int(item) for item in selected_items])
        if indices[0] == 0: return
        for i in indices: self.variables_data.insert(i - 1, self.variables_data.pop(i))
        self.populate_treeview()
        new_selection_ids = [str(i - 1) for i in indices]
        self.tree.selection_set(new_selection_ids); self.tree.focus(new_selection_ids[0]); self.tree.see(new_selection_ids[0])

    def move_down(self):
        selected_items = self.tree.selection();
        if not selected_items: return
        indices = sorted([int(item) for item in selected_items], reverse=True)
        if indices[0] == len(self.variables_data) - 1: return
        for i in indices: self.variables_data.insert(i + 1, self.variables_data.pop(i))
        self.populate_treeview()
        new_selection_ids = [str(i + 1) for i in sorted([int(item) for item in selected_items])]
        self.tree.selection_set(new_selection_ids); self.tree.focus(new_selection_ids[0]); self.tree.see(new_selection_ids[0])

    def on_selection_change(self, event=None):
        if self.tree.selection(): self.move_up_button['state'] = tk.NORMAL; self.move_down_button['state'] = tk.NORMAL
        else: self.move_up_button['state'] = tk.DISABLED; self.move_down_button['state'] = tk.DISABLED

    def save_syntax_file(self):
        if not self.original_filepath: messagebox.showerror("ข้อผิดพลาด", "ยังไม่มีข้อมูลให้บันทึก"); return
        base_name = os.path.basename(self.original_filepath)
        default_syntax_name = os.path.splitext(base_name)[0] + "_Final.sps"
        syntax_path = filedialog.asksaveasfilename(title="บันทึกไฟล์ Syntax เป็น", initialfile=default_syntax_name, defaultextension=".sps", filetypes=(("SPSS Syntax File", "*.sps"), ("All files", "*.*")))
        if not syntax_path: return
        try:
            original_file_path_spss = self.original_filepath.replace('\\', '/')
            dir_name = os.path.dirname(original_file_path_spss); file_name, file_ext = os.path.splitext(os.path.basename(original_file_path_spss))
            final_sav_path_spss = f"{dir_name}/{file_name}_Final{file_ext}"
            keep_vars = "\n".join([var['name'] for var in self.variables_data])
            syntax_content = f"""* Generated by SpssVariableEditor.
GET FILE='{original_file_path_spss}'.
ADD FILES /FILE=* /KEEP={keep_vars}.
EXECUTE.
SAVE OUTFILE='{final_sav_path_spss}'
/COMPRESSED.
EXECUTE.
""".strip()
            with open(syntax_path, 'w', encoding='utf-8') as f: f.write(syntax_content)
            messagebox.showinfo("สำเร็จ", f"บันทึกไฟล์ Syntax ที่:\n{syntax_path}\nเรียบร้อยแล้ว\n\nคำสั่งใน Syntax จะทำการเรียงลำดับและลบตัวแปรตามที่แสดงในโปรแกรม")
        except Exception as e: messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถสร้างไฟล์ Syntax ได้:\n{e}")

# (ส่วนที่เหลือของโค้ดเหมือนเดิม)
def run_this_app(working_dir=None):
    print(f"--- SPssVariableEditor: Starting app ---")
    try:
        root = tk.Tk()
        app = SpssVariableEditor(root)
        root.mainloop()
        print(f"--- SPssVariableEditor: mainloop finished. ---")
    except Exception as e:
        print(f"SPssVariableEditor_ERROR: An error occurred during execution: {e}")
        import sys
        if 'root' not in locals() or not root.winfo_exists():
            root_temp = tk.Tk()
            root_temp.withdraw()
            messagebox.showerror("Application Error", f"An unexpected error occurred:\n{e}", parent=root_temp)
            root_temp.destroy()
        else:
            messagebox.showerror("Application Error", f"An unexpected error occurred:\n{e}", parent=root)
        sys.exit(f"Error running app: {e}")

if __name__ == "__main__":
    print("--- Running SpssVariableEditor.py directly for testing ---")
    run_this_app()
    print("--- Finished direct execution of SpssVariableEditor.py ---")