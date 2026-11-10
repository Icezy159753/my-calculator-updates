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

LOOP_VAR_RE = re.compile(r'^[Ii]_(\d+)_(.+)$')                  # I_1_q1_1_O2
LOOP_OPTION_RE = re.compile(r'_(O\d+|oth\d*)$', re.IGNORECASE)   # _O1 (ตัวเลือก MA), _oth (ระบุอื่นๆ)
LOOP_ATTR_RE = re.compile(r'_([A-Za-z]{1,3})\d+$')               # _l1, _r1 (Loop Attribute)
OTH_RE = re.compile(r'_oth\d*$', re.IGNORECASE)                  # q1_1_oth
SPEC_VAR_RE = re.compile(r'^S_\d+_\d+$', re.IGNORECASE)          # S_180_5 (ช่องระบุของระบบ)
NUM_LOOP_RE = re.compile(r'^(.*\d.*?)_(\d+)$')                   # q1_1a_207 -> (q1_1a, 207)
HASH_LOOP_RE = re.compile(r'^(.+?)#(\d+)(\$.*)?$')               # C1#3$2 -> (C1, 3)

def loop_question_base(rest):
    """
    หาชื่อข้อของตัวแปร I_ Loop:
      p6_l3 -> p6_l (Attribute l1..ln คือข้อเดียวกัน), p6_l3_O1 -> p6_l
      q1_1_O2 -> q1_1, na1_oth -> na1
      p6_2 / p6_3 / b1_10 / Q_195 -> คงเดิม (เลขล้วนถือเป็นคนละข้อ)
    """
    base = LOOP_OPTION_RE.sub('', rest) or rest
    m = LOOP_ATTR_RE.search(base)
    if m and m.start() > 0:
        base = base[:m.start()] + '_' + m.group(1)
    return base

def normalize_label(text):
    """ตัดข้อความในวงเล็บ (มักเป็นชื่อแบรนด์/ชื่อ Loop) ตัวเลข และเครื่องหมายออก เหลือตัวโจทย์ไว้เทียบกัน"""
    text = re.sub(r'\([^)]*\)|\[[^\]]*\]|\{[^}]*\}', '', str(text or ''))
    return re.sub(r'[\W\d_]+', '', text).lower()

def classify_var(var):
    """
    คืน
      ('member', family, loop, signature) = ตัวแปรของข้อ Loop
      ('attach', base)                    = ตัวแปร _oth / S_<n>_<n> ที่ติดมากับข้อ
      None                                = ตัวแปรทั่วไป
    - I_<loop>_...  : ดูจากชื่ออย่างเดียว (ชัดเจนอยู่แล้ว)
    - รูปแบบอื่น เช่น q1_1a_207_O1, q26f_1_O3, q5_2_r1, C1#2$5 : signature = (โจทย์, Value Labels)
      ใช้ตรวจใน sort_loop_order ว่าทุกตัวในข้อเหมือนกันหมดจริงๆ
    """
    name = str(var['name'])
    m = LOOP_VAR_RE.match(name)
    if m:
        return ('member', ('I', loop_question_base(m.group(2)).lower()), int(m.group(1)), None)
    if SPEC_VAR_RE.match(name):
        return ('attach', None)
    if OTH_RE.search(name):
        return ('attach', OTH_RE.sub('', name).lower())
    m = HASH_LOOP_RE.match(name)
    if not m:
        base = LOOP_OPTION_RE.sub('', name)
        attr = LOOP_ATTR_RE.search(base)
        if attr and attr.start() > 0:
            base = base[:attr.start()]
        m = NUM_LOOP_RE.match(base)
    if not m:
        return None
    stem, loop = m.group(1), m.group(2)
    values = tuple(sorted((str(k), str(v)) for k, v in (var.get('values') or {}).items()))
    return ('member', ('N', stem.lower()), loop, (normalize_label(var.get('label')), values))

def loop_item_name(name):
    """ชื่อข้อย่อยของตัวแปร (ตัด _O<n> ออก) ใช้จับคู่กับ _oth ของมันเอง"""
    return LOOP_OPTION_RE.sub('', str(name)).lower()

def sort_loop_order(variables):
    """
    variables: list ของ dict ที่มี 'name', 'label', 'values'
    คืน (ลำดับ index ใหม่, จำนวนข้อ Loop ที่พบ, จำนวน _oth / S_ ที่ถูกย้ายไปท้ายข้อ)

    ตัวแปรที่ติดกันเป็นบล็อก (ตัวแปร Loop + _oth/S_) จะถูกจัดเป็น:
      ข้อ A: Loop 1 (ข้อย่อยตามลำดับเดิม) -> Loop 2 -> ... -> _oth/S_ ที่คั่นอยู่
      ข้อ B: ...
    ตัวแปรทั่วไปอยู่ที่เดิมและเป็นตัวคั่นบล็อก (ไม่ย้ายอะไรข้ามมัน)
    """
    infos = [classify_var(v) for v in variables]

    loops_per_family, signatures = {}, {}
    for info in infos:
        if info and info[0] == 'member':
            loops_per_family.setdefault(info[1], set()).add(info[2])
            signatures.setdefault(info[1], set()).add(info[3])
    # เป็นข้อ Loop ได้เมื่อ: มีมากกว่า 1 Loop และ (ถ้าไม่ใช่ I_) โจทย์ + Value Labels ต้องเหมือนกันทุกตัว
    loop_families = {f for f, loops in loops_per_family.items()
                     if len(loops) > 1 and (f[0] == 'I' or len(signatures[f]) == 1)}
    infos = [None if info and info[0] == 'member' and info[1] not in loop_families else info
             for info in infos]

    # ข้อที่ไม่ใช่ I_ ต้องอยู่ในบล็อกเดียวครบทุกตัว ถ้าถูกตัวแปรอื่นคั่นแยกเป็นหลายช่วง -> ไม่แตะข้อนั้น
    while True:
        block_of, block_id = {}, 0
        for info in infos:
            if not info:
                block_id += 1
            elif info[0] == 'member' and info[1][0] == 'N':
                block_of.setdefault(info[1], set()).add(block_id)
        split = {f for f, blocks in block_of.items() if len(blocks) > 1}
        if not split:
            break
        infos = [None if info and info[0] == 'member' and info[1] in split else info
                 for info in infos]
    loop_families -= {f for f in loop_families if f[0] == 'N' and f not in block_of}

    result, block = [], []
    n_moved = 0

    def flush():
        nonlocal n_moved
        groups, tails, lead = {}, {}, []
        ranks = {}
        current = None   # (family, loop_key, ชื่อข้อของตัวล่าสุด)
        for idx in block:
            info = infos[idx]
            if info[0] == 'member':
                family, loop = info[1], info[2]
                rank = ranks.setdefault(family, {})
                rank.setdefault(loop, len(rank))
                loop_key = loop if family[0] == 'I' else rank[loop]
                groups.setdefault(family, []).append((loop_key, idx))
                current = (family, loop_key, loop_item_name(variables[idx]['name']))
            elif current is None:
                lead.append(idx)
            elif info[1] is not None and (info[1] == current[2] or info[1].startswith(current[2] + '_')):
                # _oth ของข้อย่อยนี้เอง (q75_1 -> q75_1_oth, a1_1_1 -> a1_1_1_1001_oth) อยู่ติดกับ Loop ของมัน
                groups[current[0]].append((current[1], idx))
            else:
                tails.setdefault(current[0], []).append(idx)
        result.extend(lead)
        for family, items in groups.items():
            result.extend(idx for _, idx in sorted(items))
            tail = tails.get(family, [])
            last_member = max(idx for _, idx in items)
            n_moved += sum(1 for idx in tail if idx < last_member)
            result.extend(tail)
        block.clear()

    for i, info in enumerate(infos):
        if info:
            block.append(i)
        else:
            flush()
            result.append(i)
    flush()
    return result, len(loop_families), n_moved


def compute_auto_trim(names, start_var, end_var, head_anchor="Date", keep_vars=("Status",)):
    """
    คืน index ของตัวแปรที่ต้องตัด:
    - หัว: ตั้งแต่ head_anchor (ถ้าไม่พบ เริ่มจากตัวแรก) จนถึงก่อน start_var โดยเว้น keep_vars ไว้
    - ท้าย: ทุกตัวแปรหลัง end_var
    """
    start_idx = names.index(start_var)
    end_idx = names.index(end_var)
    if start_idx > end_idx:
        raise ValueError(f"ตัวแปรเริ่มต้น '{start_var}' อยู่หลังตัวแปรสุดท้าย '{end_var}'")
    lower = [n.lower() for n in names]
    head_from = lower.index(head_anchor.lower()) if head_anchor.lower() in lower else 0
    keep = {k.lower() for k in keep_vars}
    head = [i for i in range(head_from, start_idx) if lower[i] not in keep]
    tail = list(range(end_idx + 1, len(names)))
    return head, tail

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
        style.configure("TRadiobutton", background="white", font=("Segoe UI", 10))
        
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
            self.ask_trim_option(successful_encoding)
        else:
            # ถ้าลองทุกตัวแล้วยังไม่สำเร็จ ให้แสดงข้อความแจ้งเตือน
            messagebox.showerror(
                "เกิดข้อผิดพลาด",
                f"ไม่สามารถอ่านไฟล์ SPSS ได้\n"
                f"ได้ทดลองใช้ Encoding ต่อไปนี้แล้ว แต่ไม่สำเร็จ:\n"
                f"{', '.join(encodings_to_try)}\n\n"
                "กรุณาตรวจสอบไฟล์ หรือลองหา Encoding ที่ถูกต้อง"
            )

    def ask_trim_option(self, encoding):
        names = [v['name'] for v in self.variables_data]

        dlg = tk.Toplevel(self.root)
        dlg.title("ตัดตัวแปรหัว-ท้าย")
        dlg.configure(background='white')
        dlg.transient(self.root)
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=15, style="White.TFrame")
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frm,
            text=f"โหลดไฟล์ {os.path.basename(self.original_filepath)} แล้ว "
                 f"({len(names)} ตัวแปร, Encoding: {encoding})\n\nต้องการตัดตัวแปรหัว-ท้ายอย่างไร?",
            font=("Segoe UI", 10)
        ).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 8))

        mode = tk.StringVar(value="auto")
        input_frame = ttk.Frame(frm, padding=(20, 5, 0, 5), style="White.TFrame")

        def toggle():
            if mode.get() == "auto":
                input_frame.grid()
                start_combo.focus_set()
            else:
                input_frame.grid_remove()

        ttk.Radiobutton(frm, text="1. ตัดอัตโนมัติ (หัว-ท้าย)", variable=mode, value="auto",
                        command=toggle).grid(row=1, column=0, columnspan=2, sticky='w')
        input_frame.grid(row=2, column=0, columnspan=2, sticky='w')
        ttk.Radiobutton(frm, text="2. เลือกตัดเอง", variable=mode, value="manual",
                        command=toggle).grid(row=3, column=0, columnspan=2, sticky='w', pady=(4, 0))

        ttk.Label(input_frame, text="ตัวแปรเริ่มต้น:").grid(row=0, column=0, sticky='w', pady=3)
        start_combo = ttk.Combobox(input_frame, values=names, width=30)
        start_combo.grid(row=0, column=1, sticky='w', padx=5, pady=3)
        ttk.Label(input_frame, text="ตัวแปรสุดท้าย:").grid(row=1, column=0, sticky='w', pady=3)
        end_combo = ttk.Combobox(input_frame, values=names, width=30)
        end_combo.grid(row=1, column=1, sticky='w', padx=5, pady=3)
        if names:
            end_combo.set(names[-1])
        ttk.Label(
            input_frame,
            text="* ตัดตั้งแต่ Date จนถึงก่อนตัวแปรเริ่มต้น (เว้น Status ไว้)\n"
                 "* ตัดทุกตัวแปรที่อยู่หลังตัวแปรสุดท้าย",
            foreground="#666666"
        ).grid(row=2, column=0, columnspan=2, sticky='w', pady=(5, 0))

        def resolve(text):
            text = text.strip()
            if text in names:
                return text
            matches = [n for n in names if n.lower() == text.lower()]
            return matches[0] if matches else None

        def preview(label, idxs):
            shown = [names[i] for i in idxs[:15]]
            more = f" ... (+{len(idxs) - 15})" if len(idxs) > 15 else ""
            return f"{label} {len(idxs)} ตัว: " + (", ".join(shown) + more if idxs else "-")

        def on_ok():
            if mode.get() == "manual":
                dlg.destroy()
                return
            start_var = resolve(start_combo.get())
            end_var = resolve(end_combo.get())
            if not start_var or not end_var:
                messagebox.showwarning("คำเตือน", "ไม่พบชื่อตัวแปรเริ่มต้นหรือตัวแปรสุดท้ายที่กรอก", parent=dlg)
                return
            try:
                head, tail = compute_auto_trim(names, start_var, end_var)
            except ValueError as e:
                messagebox.showwarning("คำเตือน", str(e), parent=dlg)
                return
            if not head and not tail:
                messagebox.showinfo("แจ้งเตือน", "ไม่มีตัวแปรที่ต้องตัด", parent=dlg)
                dlg.destroy()
                return
            msg = (f"{preview('ตัดหัว', head)}\n\n{preview('ตัดท้าย', tail)}\n\n"
                   f"รวมตัดทั้งหมด {len(head) + len(tail)} ตัวแปร ยืนยันหรือไม่?")
            if not messagebox.askyesno("ยืนยันการตัดตัวแปร", msg, parent=dlg):
                return
            remove = set(head) | set(tail)
            self.variables_data = [v for i, v in enumerate(self.variables_data) if i not in remove]
            self.populate_treeview()
            dlg.destroy()
            messagebox.showinfo("สำเร็จ", f"ตัดตัวแปร {len(remove)} ตัวแล้ว เหลือ {len(self.variables_data)} ตัวแปร")

        btn_frame = ttk.Frame(frm, style="White.TFrame")
        btn_frame.grid(row=4, column=0, columnspan=2, pady=(12, 0))
        ttk.Button(btn_frame, text="ตกลง", command=on_ok, style='Green.TButton', width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="ยกเลิก", command=dlg.destroy, width=10).pack(side=tk.LEFT, padx=5)

        dlg.bind("<Return>", lambda e: on_ok())
        dlg.bind("<Escape>", lambda e: dlg.destroy())

        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        dlg.grab_set()
        start_combo.focus_set()
        self.root.wait_window(dlg)

    def sort_variables(self):
        if not self.variables_data:
            messagebox.showwarning("คำเตือน", "ยังไม่มีข้อมูลตัวแปรให้เรียงลำดับ")
            return

        names = [v['name'] for v in self.variables_data]
        order, n_loops, n_moved = sort_loop_order(self.variables_data)
        if order == list(range(len(names))):
            messagebox.showinfo("แจ้งเตือน", "ลำดับตัวแปรถูกต้องอยู่แล้ว ไม่มีอะไรต้องเรียง")
            return

        self.variables_data = [self.variables_data[i] for i in order]
        self.populate_treeview()
        lines = []
        if n_loops:
            lines.append(f"- พบข้อ Loop {n_loops} ข้อ")
        if n_moved:
            lines.append(f"- ย้ายตัวแปร _oth / S_ ที่คั่นกลางชุดคำถาม ไปท้ายชุด {n_moved} ตัว")
        messagebox.showinfo("สำเร็จ", "จัดเรียงตัวแปรเรียบร้อยแล้ว\n" + "\n".join(lines))

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