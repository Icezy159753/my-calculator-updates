from PyQt6 import QtCore, QtGui, QtWidgets
import pandas as pd
import pyreadstat
import re
import tempfile
import os
import sys
import subprocess
import numpy as np
import traceback
import faulthandler
# import traceback # Optional: for detailed error logging
# --- ฟังก์ชัน parse_list_range (จำเป็นสำหรับ expand_wildcard ที่อัปเดต) ---


# --- Global Variables ---
saved_conditions = []
original_df_columns_list = []
lower_df_columns_set = set()
lower_to_original_map = {}
current_df = None
condition_counts = {} # <--- เพิ่มตัวแปรนี้
spss_meta = None # <--- เพิ่มตัวแปร global สำหรับเก็บ Metadata
STRICT_MODE = True  # ถ้าต้องการปิดโหมดเข้มงวดให้ตั้งเป็น False

# --- Constants ---
HELP_TEXT = (
    "รูปแบบเงื่อนไข:\n"
    "  = หรือ == คือ เท่ากับ\n"
    "  != คือ ไม่เท่ากับ\n"
    "  | หรือ OR คือ หรือ\n"
    "  & หรือ AND คือ และ\n"
    "  > คือ มากกว่า\n"
    "  < คือ น้อยกว่า\n"
    "  >= คือ มากกว่าหรือเท่ากับ\n"
    "  <= คือ น้อยกว่าหรือเท่ากับ\n"
    "  =ROFF คือ เท่ากับค่าว่าง (is null) - ไม่สนตัวพิมพ์เล็ก/ใหญ่\n"
    "  =RON คือ เท่ากับมีข้อมูล (is not null) - ไม่สนตัวพิมพ์เล็ก/ใหญ่\n"
    "  !=ROFF คือ ไม่เท่ากับค่าว่าง (is not null)\n"
    "  !=RON คือ ไม่เท่ากับค่าว่าง (is null)\n"
    "\n"
    "หมายเหตุ: ชื่อคอลัมน์ (ตัวแปร) ไม่สนตัวพิมพ์เล็ก/ใหญ่ (เช่น S3 กับ s3 เหมือนกัน)\n"
    "\n"
    "การใช้ List/Range (คั่นด้วย , หรือ -):\n"
    "  s3 = 1,3-5   (หมายถึง s3 เท่ากับ 1 หรือ 3 หรือ 4 หรือ 5)\n"
    "  s3 != 1,3-5  (หมายถึง s3 ไม่ใช่ 1 และไม่ใช่ 3 และไม่ใช่ 4 และไม่ใช่ 5)\n"
    "\n"
    "การใช้ Wildcard (สำหรับคอลัมน์ที่ลงท้ายด้วย _O ตามด้วยตัวเลข เช่น s8_1_O1, s8_1_O2,...):\n"
    "รูปแบบ: <ชื่อฐาน_O> <Operator> <ค่า>\n"
    "  - Operator ที่รองรับ: =, ==, !=\n"
    "  - ค่าที่รองรับ: RON, ROFF, ตัวเลขเดี่ยว (ยังไม่รองรับ List/Range ใน Wildcard)\n"
    "ตัวอย่าง Wildcard:\n"
    "  s8_1_O = RON   (มี s8_1_O<เลข> อย่างน้อย 1 คอลัมน์ที่มีข้อมูล)\n"
    "  s8_1_O = ROFF  (ทุกคอลัมน์ s8_1_O<เลข> ต้องเป็นค่าว่างทั้งหมด)\n"
    "  s8_1_O = 12    (มี s8_1_O<เลข> อย่างน้อย 1 คอลัมน์ที่เท่ากับ 12)\n"
    "  s8_1_O != 12   (ทุกคอลัมน์ s8_1_O<เลข> ต้องไม่เท่ากับ 12)\n"
    "  s8_1_O != ROFF (มี s8_1_O<เลข> อย่างน้อย 1 คอลัมน์ที่มีข้อมูล)\n"
    "  s8_1_O != RON  (ทุกคอลัมน์ s8_1_O<เลข> ต้องเป็นค่าว่างทั้งหมด)\n"
    "\n"
    "ตัวอย่างเงื่อนไขผสม:\n"
    "  s3=1,3-5 & s4_range!=3\n"
    "  (s8_1_O=12 | s8_2_O=12) & s4_range=3\n"
    "  AGE > 25 AND INCOME <= 50000\n"
)

# --- UI Helpers (PyQt6) ---
ui = None

class UiVar:
    """Simple get/set wrapper to replace Tkinter StringVar behavior."""
    def __init__(self, getter=None, setter=None):
        self._getter = getter
        self._setter = setter
        self._value = ""

    def get(self):
        return self._getter() if self._getter else self._value

    def set(self, value):
        self._value = value if value is not None else ""
        if self._setter:
            self._setter(self._value)

def _report_fatal_error(title, message):
    """พยายามรายงานข้อผิดพลาดโดยไม่ให้แอปล้มซ้ำ"""
    try:
        if ui is not None:
            ui.show_error(title, message)
            return
    except Exception:
        pass
    print(f"{title}: {message}")

def _append_runtime_log(message):
    try:
        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "runtime_debug.log")
        ts = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {message}\n")
    except Exception:
        pass

class SweetAlert(QtWidgets.QDialog):
    def __init__(self, title, message, kind="info", buttons=("OK",), parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowType.FramelessWindowHint)
        self._result = None

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(16)

        header = QtWidgets.QHBoxLayout()
        icon_label = QtWidgets.QLabel()
        icon_label.setFixedSize(36, 36)
        icon_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        icon_label.setObjectName("iconLabel")
        header.addWidget(icon_label)

        title_label = QtWidgets.QLabel(title)
        title_label.setObjectName("titleLabel")
        header.addWidget(title_label)
        header.addStretch(1)
        root.addLayout(header)

        msg = QtWidgets.QLabel(message)
        msg.setWordWrap(True)
        msg.setObjectName("messageLabel")
        root.addWidget(msg)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch(1)
        for b in buttons:
            btn = QtWidgets.QPushButton(b)
            btn.clicked.connect(self._make_handler(b))
            btn_row.addWidget(btn)
        root.addLayout(btn_row)

        color_map = {
            "info": "#3b82f6",
            "success": "#16a34a",
            "warning": "#f59e0b",
            "error": "#ef4444",
            "question": "#6366f1",
        }
        accent = color_map.get(kind, "#3b82f6")
        icon_label.setStyleSheet(
            f"background:{accent};color:white;border-radius:18px;font-weight:bold;"
        )
        icon_label.setText("!")

        self.setStyleSheet("""
            QDialog { background: #ffffff; border-radius: 12px; }
            QLabel#titleLabel { font-size: 14px; font-weight: 700; color: #111827; }
            QLabel#messageLabel { font-size: 12px; color: #374151; }
            QPushButton {
                background: #111827; color: white; border: none; padding: 8px 16px;
                border-radius: 8px; font-weight: 600;
            }
            QPushButton:hover { background: #0b1220; }
        """)

        self.adjustSize()

    def _make_handler(self, label):
        def handler():
            self._result = label
            self.accept()
        return handler

    @property
    def result(self):
        return self._result

def set_button_enabled(btn, enabled):
    if btn is not None:
        btn.setEnabled(enabled)

def process_events():
    app = QtWidgets.QApplication.instance()
    if app:
        app.processEvents()

# --- Helper Functions ---

def resource_path(relative_name):
    """ คืน path ที่ถูกต้อง ไม่ว่าจะ run จากไฟล์ .py หรือ bundle เป็น exe """
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_name)




def calculate_single_count(condition, df, original_cols, lower_cols_set, lower_to_orig_map):
    """คำนวณ Count สำหรับเงื่อนไขเดียว"""
    if df is None:
        return "N/A" # หรือ None หรือค่าที่เหมาะสมถ้ายังไม่โหลดข้อมูล
    try:
        # ทำขั้นตอนเหมือนใน compute_counts แต่สำหรับเงื่อนไขเดียว
        error_msg = validate_condition(condition, original_cols, lower_cols_set, lower_to_orig_map)
        if error_msg:
            # อาจจะคืนค่า Error หรือ raise exception เพื่อให้ save_condition จัดการ
            # print(f"Validation Error for '{condition}': {error_msg}") # Debug
            return "Error" # คืนค่า Error ไปแสดงผล

        expanded = expand_wildcard(condition, original_cols, lower_to_orig_map)
        converted = auto_convert(expanded, lower_to_orig_map)
        count = len(df.query(converted))
        return count

    except Exception as e:
        # print(f"Error calculating single count for '{condition}': {e}") # Debug
        return "Error" # คืนค่า Error ไปแสดงผล

# --- ฟังก์ชัน parse_list_range (จำเป็นสำหรับ expand_wildcard ที่อัปเดต) ---
def parse_list_range(val):
    """ แปลงสตริง list/range เช่น '1,3-5' เป็น list ตัวเลข [1, 3, 4, 5] """
    items = []
    try:
        for part in val.split(','):
            part = part.strip()
            if '-' in part:
                a,b = map(int, part.split('-',1))
                if a > b: a, b = b, a
                items.extend(range(a, b+1))
            elif part:
                items.append(int(part))
    except ValueError:
        raise ValueError(f"รูปแบบค่า '{val}' ไม่ถูกต้อง (ต้องเป็นตัวเลข, คอมม่า, หรือขีดกลาง)")
    return sorted(list(set(items)))

# --- Core Logic Functions ---

def load_dataframe(filepath):
    """ โหลด DataFrame จากไฟล์ SPSS """
    if not filepath or not os.path.exists(filepath): return None, "ไม่พบไฟล์ที่ระบุ"
    try:
        df, meta = pyreadstat.read_sav(filepath)
        # Optional: Clean column names if necessary (e.g., remove leading/trailing spaces)
        # df.columns = [col.strip() for col in df.columns]
        return df, None
    except Exception as e:
        return None, f"เกิดข้อผิดพลาดในการอ่านไฟล์ SPSS: {e}"

def validate_condition(expr, original_cols, lower_cols_set, lower_to_orig_map):
    """
    ตรวจสอบความถูกต้องเบื้องต้นของเงื่อนไข (Case-Insensitive สำหรับชื่อคอลัมน์)
    Returns: str (error message) or None (if valid)
    """
    if not expr.strip(): return "เงื่อนไขว่างเปล่า"
    if expr.count('(') != expr.count(')'): return "จำนวนวงเล็บเปิด '(' และปิด ')' ไม่เท่ากัน"

    # Check invalid operators with ROFF/RON
    roff_ron_pattern = re.compile(r"\b(\w+)\s*([<>]=?)\s*(ROFF|RON)\b", flags=re.IGNORECASE)
    match = roff_ron_pattern.search(expr)
    if match:
        var, op, keyword = match.groups()
        return f"ใช้ Operator '{op}' กับ '{keyword.upper()}' ไม่ได้ (ต้องใช้ = หรือ !=)"

    # Check for invalid column names/wildcards (Case-Insensitive)
    potential_vars = set(re.findall(r'\b[a-zA-Z_][a-zA-Z0-9_]*\b', expr))
    known_keywords = {'and', 'or', 'not', 'in', 'ron', 'roff', 'isnull', 'notnull', 'isin'}
    lower_wildcard_bases = {m.group(1).lower() for m in re.finditer(r'\b(\w+_O)\b', expr, flags=re.IGNORECASE)}
    invalid_cols_display = []

    for token in potential_vars:
        token_lower = token.lower()
        if token_lower in known_keywords or token.isdigit(): continue

        # Check Wildcard Base
        if token_lower in lower_wildcard_bases:
            has_matching_cols = any(
                col.lower().startswith(token_lower) and re.match(r'^\d+$', col[len(token_lower):])
                for col in original_cols
            )
            if not has_matching_cols:
                invalid_cols_display.append(f"{token}* (ไม่พบคอลัมน์ {token}<ตัวเลข>)")
            continue # Skip to next token if it's a wildcard base

        # Check normal variable name
        if token_lower not in lower_cols_set:
            invalid_cols_display.append(token) # Show the name as user typed it

    if invalid_cols_display:
        return f"ไม่พบชื่อคอลัมน์หรือรูปแบบ Wildcard: {', '.join(invalid_cols_display)}"

    # Check List/Range format after =, ==, !=
    list_range_pattern = re.compile(r"\b\w+\s*([=!]?=)\s*([\d,\s-]+)\b")
    for match in list_range_pattern.finditer(expr):
        # Only validate the format if operator is for equality/inequality
        if match.group(1) in ['=', '==', '!=']:
            vals = match.group(2)
            try:
                parse_list_range(vals) # Try parsing
            except ValueError as e:
                # Provide context in the error message
                return f"รูปแบบรายการ/ช่วงตัวเลขไม่ถูกต้องใน '{match.group(0)}': {e}"

    # All preliminary checks passed
    return None

# --- ฟังก์ชัน expand_wildcard ที่อัปเดตแล้ว ---
def expand_wildcard(expr, original_cols, lower_to_orig_map):
    """
    ขยาย wildcard (Case-Insensitive matching, Original case output)
    รองรับ: RON, ROFF (รวม '' ด้วย), ตัวเลขเดี่ยว, และ List/Range กับ Operator =, ==, !=
    """
    # Regex pattern ที่อัปเดตเพื่อจับ value ที่เป็น list/range ได้ด้วย
    pattern = re.compile(
        r"""
        \b(\w+_O)\b                     # Capture base name (Group 1) e.g., s16a_O
        \s*([=!]?=)\s* # Capture operator =, ==, != (Group 2)
        # Capture value: RON, ROFF, or sequence of digits, commas, spaces, hyphens
        # Negative lookahead (?!\w) ป้องกันการจับส่วนหนึ่งของคำอื่น ถ้าไม่ใช่ RON/ROFF
        (RON|ROFF|[\d,\s-]+(?!\w))
        """,
        flags=re.IGNORECASE | re.VERBOSE
    )

    def repl(m):
        # ดึงค่าที่จับได้จาก Regex
        prefix_as_typed = m.group(1)
        prefix_lower = prefix_as_typed.lower()
        op_raw = m.group(2).strip()
        val_str = m.group(3).strip()
        val_str_upper = val_str.upper()

        # ค้นหาคอลัมน์ทั้งหมดที่ตรงกับ pattern (เทียบ lowercase) แต่เก็บชื่อเดิม (original case)
        cols_original_case = sorted([
            orig_col for orig_col in original_cols
            if orig_col.lower().startswith(prefix_lower) and re.match(r'^\d+$', orig_col[len(prefix_lower):])
        ])

        # ถ้าไม่พบคอลัมน์ที่ตรง wildcard เลย ให้แจ้ง error
        if not cols_original_case:
            raise ValueError(f"ไม่พบคอลัมน์ที่ตรงกับรูปแบบ wildcard '{prefix_as_typed}*'")

        parts = []  # List เก็บส่วนของ query string สำหรับแต่ละคอลัมน์
        joiner = '' # ตัวเชื่อมระหว่างส่วนต่างๆ (' | ' หรือ ' & ')

        # --- กรณี: ค่าเป็น RON หรือ ROFF (ปรับปรุงให้ตรวจ '' ด้วย) ---
        if val_str_upper in ('RON', 'ROFF'):
            # กำหนดว่าเป็นเงื่อนไขเช็คเท่ากับ (=, ==) หรือ ไม่เท่ากับ (!=)
            if op_raw == '=' or op_raw == '==': is_positive_match = True
            elif op_raw == '!=': is_positive_match = False
            else: raise ValueError(f"Operator '{op_raw}' ใช้กับ '{val_str_upper}' ใน wildcard ไม่ได้")

            # กำหนด query template และ joiner ตาม logic ที่อัปเดต
            # =ROFF หรือ !=RON -> เช็ค (isnull() หรือ == '') สำหรับ *ทุก* คอลัมน์
            if (val_str_upper == 'ROFF' and is_positive_match) or \
               (val_str_upper == 'RON' and not is_positive_match):
                # Template สำหรับเช็ค null หรือ empty string
                col_expr_template = "( (`{c_orig}`.isnull()) | (`{c_orig}` == '') )"
                joiner = ' & ' # ใช้ AND เพราะทุกคอลัมน์ต้องเข้าเงื่อนไขนี้
            # =RON หรือ !=ROFF -> เช็ค (notnull() และ != '') สำหรับ *อย่างน้อยหนึ่ง* คอลัมน์
            elif (val_str_upper == 'RON' and is_positive_match) or \
                 (val_str_upper == 'ROFF' and not is_positive_match):
                # Template สำหรับเช็ค not null และ not empty string
                col_expr_template = "( (`{c_orig}`.notnull()) & (`{c_orig}` != '') )"
                joiner = ' | ' # ใช้ OR เพราะขอแค่คอลัมน์เดียวเข้าเงื่อนไขนี้
            else: # กรณีที่ไม่ควรเกิดขึ้น
                 raise ValueError("Logic error ในการประมวลผล wildcard RON/ROFF")

            # สร้าง query string สำหรับแต่ละคอลัมน์
            parts = [col_expr_template.format(c_orig=c_orig) for c_orig in cols_original_case]

        # --- กรณี: ค่าเป็นตัวเลขเดี่ยว ---
        elif re.fullmatch(r'\d+', val_str):
            val_num = val_str
            if op_raw == '=' or op_raw == '==':
                op, joiner = '==', ' | '
                parts = [f"(`{c_orig}` {op} {val_num})" for c_orig in cols_original_case]
            elif op_raw == '!=':
                joiner = ' & '
                if STRICT_MODE:
                    # โหมดเข้มงวด: ตัด missing/ค่าว่างออกจากเงื่อนไข !=
                    parts = [f"((`{c_orig}`.notnull()) & (`{c_orig}` != '') & (`{c_orig}` != {val_num}))"
                             for c_orig in cols_original_case]
                else:
                    parts = [f"(`{c_orig}` != {val_num})" for c_orig in cols_original_case]
            else: raise ValueError(f"Operator '{op_raw}' กับตัวเลข '{val_num}' ใน wildcard ไม่รองรับ")

        # --- กรณี: ค่าเป็น List/Range ---
        elif re.fullmatch(r'[\d,\s-]+', val_str):
            try:
                nums_list = parse_list_range(val_str)
                if not nums_list: raise ValueError(f"List/range '{val_str}' ให้ค่าว่างเปล่า")
            except ValueError as e_parse: raise ValueError(f"List/range '{val_str}' ผิดรูปแบบ: {e_parse}")

            if op_raw == '=' or op_raw == '==':
                joiner = ' | '
                col_expr_template = "`{c_orig}`.isin({nums})"
            elif op_raw == '!=':
                joiner = ' & '
                if STRICT_MODE:
                    # โหมดเข้มงวด: ตัด missing/ค่าว่างออกจากเงื่อนไข !=
                    col_expr_template = "((`{c_orig}`.notnull()) & (`{c_orig}` != '') & ~(`{c_orig}`.isin({nums})))"
                else:
                    col_expr_template = "~(`{c_orig}`.isin({nums}))"
            else: raise ValueError(f"Operator '{op_raw}' กับ list/range '{val_str}' ใน wildcard ไม่รองรับ")

            nums_repr = repr(nums_list)
            parts = [col_expr_template.format(c_orig=c_orig, nums=nums_repr) for c_orig in cols_original_case]

        # --- กรณี: ค่าเป็นรูปแบบอื่นที่ไม่รองรับ ---
        else:
            raise ValueError(f"ค่า '{val_str}' ไม่รองรับใน wildcard '{prefix_as_typed}'")

        # --- รวม query string ของแต่ละคอลัมน์ ---
        if not parts: return 'True' if joiner == ' & ' else 'False'
        else: return '(' + joiner.join(parts) + ')'

    # --- ใช้การแทนที่แบบวนซ้ำ (Iterative Replacement) ---
    processed_expr = expr
    iterations = 0
    max_iterations = len(expr) + 5 # ตั้งค่าเผื่อความยาวที่เพิ่มขึ้น

    while iterations < max_iterations:
        match = pattern.search(processed_expr)
        if not match: break
        try:
            replacement = repl(match)
            start, end = match.span()
            processed_expr = processed_expr[:start] + replacement + processed_expr[end:]
        except ValueError as e: raise e # ส่งต่อ error จาก repl
        iterations += 1

    # แจ้งเตือนถ้าวนลูปนานผิดปกติ
    if iterations >= max_iterations:
        print(f"Warning: Wildcard expansion วนลูปนานผิดปกติ: {expr}")

    # คืนค่า expression ที่ผ่านการขยาย wildcard แล้ว
    return processed_expr

def auto_convert(expr, lower_to_orig_map):
    """
    แปลงเงื่อนไขส่วนที่เหลือ (Case-Insensitive matching, Original case output)
    ทำงาน *หลังจาก* expand_wildcard
    ปรับปรุงให้ =ROFF ตรวจจับ '' ด้วย
    """
    # Helper to safely get original case name and quote it
    def get_original_quoted(var_match_lower):
        original_var = lower_to_orig_map.get(var_match_lower)
        return f"`{original_var}`" if original_var else None

    # --- Fallback for ROFF/RON (ปรับปรุงให้ตรวจ '' ด้วย) ---
    def repl_roff_ron_fb(m):
        var_quoted = get_original_quoted(m.group(1).lower())
        if not var_quoted: return m.group(0) # Return original if var not found
        op_raw = m.group(2) # =, ==, !=
        keyword = m.group(3).upper() # ROFF/RON

        # =ROFF หรือ !=RON -> เช็ค isnull() หรือ == ''
        if (keyword == 'ROFF' and (op_raw == '=' or op_raw == '==')) or \
           (keyword == 'RON' and op_raw == '!='):
            # สร้าง query string ที่ตรวจสอบทั้ง null และ empty string
            return f"(({var_quoted}.isnull()) | ({var_quoted} == ''))"
        # =RON หรือ !=ROFF -> เช็ค notnull() และ != ''
        elif (keyword == 'RON' and (op_raw == '=' or op_raw == '==')) or \
             (keyword == 'ROFF' and op_raw == '!='):
            # สร้าง query string ที่ตรวจสอบทั้ง not null และ not empty string
            return f"(({var_quoted}.notnull()) & ({var_quoted} != ''))"
        else: # กรณีที่ไม่ควรเกิดขึ้น
             return m.group(0)
    # Regex with lookarounds to avoid replacing already processed parts
    # จับเฉพาะตัวแปรที่ไม่ได้อยู่ใน backticks หรือเป็นส่วนหนึ่งของ method call แล้ว
    expr = re.sub(r'\b(?<![`\w.])(\w+)(?![._\w])\s*([=!]?=)\s*(ROFF|RON)\b', repl_roff_ron_fb, expr, flags=re.IGNORECASE)

    # --- Handle List/Range for non-wildcard variables ---
    # (ไม่มีการเปลี่ยนแปลงจากเวอร์ชันก่อนหน้า)
    def repl_list_range(m):
        var_quoted = get_original_quoted(m.group(1).lower())
        if not var_quoted: return m.group(0)
        op_raw, vals_str = m.group(2), m.group(3)
        try: nums = parse_list_range(vals_str)
        except ValueError as e: raise ValueError(f"'{vals_str}': {e}") # Propagate parse error
        op = '==' if op_raw == '=' else op_raw # Convert = to ==
        nums_str = [str(n) for n in nums]
        if op == '==':
            return f"({var_quoted}.isin({nums}) | {var_quoted}.isin({nums_str}))"
        elif op == '!=':
            if STRICT_MODE:
                # โหมดเข้มงวด: ตัด missing/ค่าว่างออกจากเงื่อนไข !=
                return f"(({var_quoted}.notnull()) & ({var_quoted} != '') & ~(({var_quoted}.isin({nums}) | {var_quoted}.isin({nums_str}))))"
            return f"~(({var_quoted}.isin({nums}) | {var_quoted}.isin({nums_str})))"
        else: return m.group(0) # Should not happen if regex is correct
    expr = re.sub(r'\b(?<![`\w.])(\w+)(?![._\w])\s*([=!]?=)\s*([\d,\s-]+)\b', repl_list_range, expr)

    # --- Handle Simple Comparisons (e.g., s3=1, Age > 20) ---
    # (ไม่มีการเปลี่ยนแปลงจากเวอร์ชันก่อนหน้า)
    def repl_simple_comp(m):
        var_quoted = get_original_quoted(m.group(1).lower())
        if not var_quoted: return m.group(0)
        op = m.group(2).strip()
        val_str = m.group(3).strip()
        op_final = '==' if op == '=' else op # Convert = to ==
        # Check if value is likely numeric or boolean, otherwise quote it using repr
        if val_str.lower() in ['true', 'false'] or re.fullmatch(r'-?\d+(\.\d+)?([eE][-+]?\d+)?', val_str):
            val_final = val_str
        else:
            # Remove surrounding quotes if present before using repr
            if (val_str.startswith("'") and val_str.endswith("'")) or \
               (val_str.startswith('"') and val_str.endswith('"')):
                 val_str_unquoted = val_str[1:-1]
            else: val_str_unquoted = val_str
            val_final = repr(val_str_unquoted) # Safely quote the string content
        if op_final == '!=':
            if STRICT_MODE:
                # โหมดเข้มงวด: ตัด missing/ค่าว่างออกจากเงื่อนไข !=
                return f"(({var_quoted}.notnull()) & ({var_quoted} != '') & ({var_quoted} {op_final} {val_final}))"
            return f"{var_quoted} {op_final} {val_final}"
        return f"{var_quoted} {op_final} {val_final}"
    # Regex to match variable <op> value (run last) - includes quoted strings
    expr = re.sub(r"""
        \b(?<![`\w.])(\w+)(?![.\w])    # Capture variable name (Group 1)
        \s*([=!]?=|>=?|<=?)          # Capture operator (Group 2)
        \s*('.*?'|".*?"|\S+)         # Capture value: quoted or non-space (Group 3)
        """, repl_simple_comp, expr, flags=re.VERBOSE)

    # --- Final conversion of AND/OR ---
    expr = re.sub(r"\bAND\b", " & ", expr, flags=re.IGNORECASE)
    expr = re.sub(r"\bOR\b", " | ", expr, flags=re.IGNORECASE)

    return expr
# --- สิ้นสุดฟังก์ชัน auto_convert ที่อัปเดตแล้ว ---

# --- ตรวจสอบและแก้ไข extract_cols_from_raw_condition ---
def extract_cols_from_raw_condition(expr, original_cols, lower_cols_set, lower_to_orig_map): # <--- ตรวจสอบว่ารับ map นี้
    """ดึงคอลัมน์ที่ใช้จากเงื่อนไขดิบ (Case-Insensitive matching, Original case output)"""
    potential_vars = set(re.findall(r'\b[a-zA-Z_][a-zA-Z0-9_]*\b', expr))
    used_original_case = []
    added_lower = set() # Keep track of added columns (lowercase) to prevent duplicates

    # --- กรองและหา Original Case ---
    for token in potential_vars:
        token_lower = token.lower()
        if token_lower in lower_cols_set:
            # *** ใช้ map ที่ส่งเข้ามาในการค้นหา ***
            original_var = lower_to_orig_map.get(token_lower)
            # ------------------------------------
            if original_var and token_lower not in added_lower:
                 used_original_case.append(original_var)
                 added_lower.add(token_lower)

    # --- จัดการ Wildcard (เหมือนเดิม - ใช้ original_cols ในการหา matching) ---
    wildcard_patterns = re.findall(r'\b(\w+_O)\b', expr, flags=re.IGNORECASE)
    for pattern_base in wildcard_patterns:
        prefix_lower = pattern_base.lower()
        matching_cols = [
            orig_col for orig_col in original_cols
            if orig_col.lower().startswith(prefix_lower) and re.match(r'^\d+$', orig_col[len(prefix_lower):])
        ]
        # เพิ่มเฉพาะคอลัมน์ที่ยังไม่มี
        for mc in matching_cols:
            mc_lower = mc.lower()
            if mc_lower not in added_lower:
                 used_original_case.append(mc)
                 added_lower.add(mc_lower)

    # --- เพิ่ม Primary Key (คอลัมน์แรก) ถ้ายังไม่มี ---
    if original_cols:
        pk_original = original_cols[0]
        pk_lower = pk_original.lower()
        if pk_lower not in added_lower:
             used_original_case.insert(0, pk_original)
             added_lower.add(pk_lower) # Mark PK as added

    # ถ้า list ยังว่าง ให้ใส่ PK
    if original_cols and not used_original_case:
         used_original_case.append(original_cols[0])

    return used_original_case
# --- สิ้นสุดการแก้ไข extract_cols_from_raw_condition ---

def compute_counts(df, original_cols, lower_cols_set, lower_to_orig_map):
    """คำนวณ counts สำหรับ saved_conditions"""
    counts = {}
    errors = {}
    if df is None:
        # Return counts for conditions, marking all as 'N/A' or similar if no DF
        for cond in saved_conditions: counts[cond] = 'N/A'
        return counts, "DataFrame ไม่ได้โหลด"

    # Process each saved condition
    for cond in saved_conditions:
        try:
            # 1. Validate
            error_msg = validate_condition(cond, original_cols, lower_cols_set, lower_to_orig_map)
            if error_msg:
                errors[cond] = f"รูปแบบผิด: {error_msg}"
                counts[cond] = "Error"
                continue # Skip to next condition

            # 2. Expand Wildcard
            expanded = expand_wildcard(cond, original_cols, lower_to_orig_map)

            # 3. Auto Convert (remaining parts)
            converted = auto_convert(expanded, lower_to_orig_map)
            # print(f"DEBUG Compute: Raw='{cond}'\nExp='{expanded}'\nConv='{converted}'") # Debug

            # 4. Query DataFrame
            counts[cond] = len(df.query(converted))

        except Exception as e:
            # Log detailed error for debugging if needed
            # import traceback
            # print(f"Error processing condition '{cond}': {e}\n{traceback.format_exc()}")
            errors[cond] = f"ประมวลผลผิดพลาด: {e}"
            counts[cond] = "Error"

    error_summary = "\n".join([f"- '{c}': {m}" for c, m in errors.items()]) if errors else None
    return counts, error_summary

# --- Function for Excel Export with Index, Counts, and Backlinks ---
# --- Function for Excel Export with Index, Counts, and Backlinks ---
def open_multi_excel(df_dict, counts_dict, filename_base="Result"):
    """
    ส่งออก DataFrame หลายอันไปยังชีทต่างๆ ในไฟล์ Excel เดียว
    พร้อมสร้าง Index Sheet ที่มี ID, Condition, Sheet Name, Count, และ Hyperlink
    และเพิ่ม Hyperlink ในแต่ละชีทข้อมูลเพื่อกลับมายัง Index Sheet
    (เวอร์ชันแก้ไข: ปรับปรุงการสร้างชื่อชีทให้ไม่เกิน 31 ตัวอักษร)
    Args:
        df_dict (dict): Dictionary ที่มี key เป็น condition string และ value เป็น DataFrame ผลลัพธ์.
        counts_dict (dict): Dictionary ที่มี key เป็น condition string และ value เป็น count (หรือ "Error"/"N/A").
        filename_base (str): ชื่อพื้นฐานสำหรับไฟล์ Excel ที่จะส่งออก.
    """
    # ตรวจสอบว่ามีข้อมูลที่จะ export หรือไม่ (result_dict)
    if not df_dict:
        # แม้ไม่มี df_dict แต่ก็อาจจะมี counts ที่ไม่ใช่ 0 (กรณี Error หรือ N/A)
        # แต่โดยทั่วไป ถ้า df_dict ว่างเปล่า หมายถึงไม่มีเงื่อนไขที่ Count > 0
        ui.show_info("ไม่มีข้อมูล", "ไม่มีเงื่อนไขใดที่มีข้อมูล (Count > 0) ที่จะส่งออก")
        return

    # ถามผู้ใช้ว่าจะบันทึกไฟล์ที่ไหนและชื่ออะไร
    save_path = ui.save_file_dialog(
        defaultextension='.xlsx',
        filetypes=[("Excel files", "*.xlsx")],
        initialfile=f'{filename_base}_{pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
        title="บันทึกผลลัพธ์เป็น Excel"
    )
    # ถ้าผู้ใช้กดยกเลิก
    if not save_path:
        return

    sheet_info_list = [] # List เพื่อเก็บข้อมูลสำหรับสร้าง Index Sheet

    try:
        # ใช้ pd.ExcelWriter พร้อม engine xlsxwriter
        with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
            workbook = writer.book # เข้าถึง workbook object

            # --- 1. สร้าง Index Sheet เป็นชีทแรก ---
            index_sheet = workbook.add_worksheet('Index')

            # --- 2. วนลูปเขียน Data Sheets (เฉพาะเงื่อนไขที่มีข้อมูลใน df_dict) ---
            sheet_idx = 1
            generated_sheet_names = {}

            # วนลูปเฉพาะเงื่อนไขที่มี DataFrame ผลลัพธ์ (คือมี Count > 0)
            for cond, df_ in df_dict.items():
                # --- [แก้ไข] สร้างชื่อชีทที่ถูกต้องและไม่ซ้ำ (ไม่เกิน 31 ตัวอักษร) ---
                # 1. ทำความสะอาด condition string และสร้างส่วนหลักของชื่อ
                safe_cond_part = re.sub(r'[\\/*?:\[\]]', '_', cond)
                prefix = f"Cond{sheet_idx}" # เช่น "Cond1", "Cond12"

                # 2. คำนวณความยาวสูงสุดที่เหลือสำหรับ condition และสร้างชื่อพื้นฐาน
                # 31 คือขีดจำกัด, -1 สำหรับ '_' ที่จะคั่นกลาง
                max_len_for_cond = 31 - len(prefix) - 1
                base_sheet_name = f"{prefix}_{safe_cond_part[:max_len_for_cond]}"

                # 3. จัดการกรณีชื่อซ้ำโดยการเติม suffix (_1, _2, ...)
                sheet_name = base_sheet_name
                name_count = 1
                while sheet_name in generated_sheet_names:
                    suffix = f"_{name_count}"
                    # คำนวณความยาวของชื่อพื้นฐานใหม่เพื่อให้พอดีกับ suffix
                    max_len_for_base = 31 - len(suffix)
                    sheet_name = f"{base_sheet_name[:max_len_for_base]}{suffix}"
                    name_count += 1
                # --- [สิ้นสุดการแก้ไข] ---
                
                generated_sheet_names[sheet_name] = True

                # เก็บข้อมูลสำหรับ Index (รวม Count จาก counts_dict)
                count_value = counts_dict.get(cond, 'N/A') # ดึง Count จาก dict ที่ส่งมา
                sheet_info_list.append({
                    'id': sheet_idx,
                    'condition': cond,
                    'sheet_name': sheet_name,
                    'count': count_value # <--- เก็บ Count ไว้ด้วย
                })

                # --- เขียน Data Sheet ---
                pd.DataFrame([{"Condition": cond}]).to_excel(writer, sheet_name=sheet_name, startrow=0, index=False, header=False)
                df_.to_excel(writer, sheet_name=sheet_name, startrow=2, index=False)
                worksheet = writer.sheets[sheet_name]

                # --- เพิ่ม Hyperlink กลับไปยัง Index Sheet ---
                backlink_url = "internal:'Index'!A1"
                worksheet.write_url('B1', backlink_url, string='<- กลับไปหน้า Index')

                # --- จัดรูปแบบคอลัมน์ใน Data Sheet ---
                for idx_col, col_name in enumerate(df_.columns):
                     try:
                         max_len_data = df_[col_name].astype(str).map(len).max()
                         if pd.isna(max_len_data): max_len_data = 0
                     except Exception: max_len_data = 10
                     max_len = max(max_len_data, len(str(col_name))) + 2
                     worksheet.set_column(idx_col, idx_col, max_len)

                sheet_idx += 1 # เพิ่มลำดับชีท

            # --- 3. กลับมา Populate ข้อมูลใน Index Sheet ---
            header_format = workbook.add_format({'bold': True})
            # *** เพิ่ม Header สำหรับ Count ***
            index_sheet.write('A1', 'ID', header_format)
            index_sheet.write('B1', 'Condition', header_format)
            index_sheet.write('C1', 'Sheet Name', header_format)
            index_sheet.write('D1', 'Count', header_format) # <--- Header Count ใหม่
            index_sheet.write('E1', 'Link to Sheet', header_format) # <--- Header Link เลื่อนไป E

            max_cond_len = 10
            max_sheet_name_len = 12

            # วนลูปเขียนข้อมูลและลิงก์สำหรับแต่ละชีทที่สร้าง (จาก sheet_info_list)
            for idx, info in enumerate(sheet_info_list):
                row_num = idx + 1 # แถวข้อมูลเริ่มที่ 2 (index 1)
                index_sheet.write(row_num, 0, info['id'])         # คอลัมน์ A: ID
                index_sheet.write(row_num, 1, info['condition'])  # คอลัมน์ B: Condition
                index_sheet.write(row_num, 2, info['sheet_name']) # คอลัมน์ C: Sheet Name
                # *** เขียนค่า Count ลงคอลัมน์ D ***
                index_sheet.write(row_num, 3, info['count'])      # คอลัมน์ D: Count
                # *** สร้างและเขียน Hyperlink ในคอลัมน์ E ***
                link_url = f"internal:'{info['sheet_name']}'!A1"
                index_sheet.write_url(row_num, 4, link_url, string='Go to Sheet') # คอลัมน์ E: Link

                # อัปเดตความกว้างสูงสุด
                if len(info['condition']) > max_cond_len: max_cond_len = len(info['condition'])
                if len(info['sheet_name']) > max_sheet_name_len: max_sheet_name_len = len(info['sheet_name'])

            # --- 4. จัดรูปแบบคอลัมน์ใน Index Sheet ---
            index_sheet.set_column('A:A', 5)  # ID
            index_sheet.set_column('B:B', min(max_cond_len + 2, 80)) # Condition
            index_sheet.set_column('C:C', max_sheet_name_len + 2) # Sheet Name
            index_sheet.set_column('D:D', 10) # Count <--- ความกว้างคอลัมน์ Count
            index_sheet.set_column('E:E', 15) # Link <--- ความกว้างคอลัมน์ Link

            # --- 5. ตั้งให้ Index Sheet เป็นชีทที่แสดงเมื่อเปิดไฟล์ ---
            index_sheet.activate()

        # --- สิ้นสุดบล็อก with (ไฟล์ Excel ถูกบันทึก) ---
        ui.show_info("สำเร็จ", f"ส่งออกข้อมูลไปยังไฟล์:\n{os.path.basename(save_path)}\nเรียบร้อยแล้ว (พร้อม Index, Count, และ Backlink)")

        # ถามผู้ใช้ว่าต้องการเปิดไฟล์หรือไม่
        if ui.ask_yes_no("เปิดไฟล์", "ต้องการเปิดไฟล์ Excel ที่ส่งออกหรือไม่?"):
            try:
                if sys.platform.startswith('win'): os.startfile(save_path)
                elif sys.platform.startswith('darwin'): subprocess.run(["open", save_path], check=True)
                else: subprocess.run(["xdg-open", save_path], check=True)
            except FileNotFoundError:
                 ui.show_error("ไม่สามารถเปิดไฟล์", f"ไม่พบคำสั่งเปิดไฟล์สำหรับระบบปฏิบัติการของคุณ\nกรุณาเปิดไฟล์ด้วยตนเองที่:\n{save_path}")
            except Exception as e:
                ui.show_error("ไม่สามารถเปิดไฟล์", f"เกิดข้อผิดพลาดในการเปิดไฟล์: {e}\n\nกรุณาเปิดไฟล์ด้วยตนเองที่:\n{save_path}")

    except Exception as e:
         ui.show_error("ส่งออก Excel ล้มเหลว", f"เกิดข้อผิดพลาดระหว่างการเขียนไฟล์ Excel: {e}")
# --- สิ้นสุดฟังก์ชัน open_multi_excel ---


# --- UI Functions ---

def update_table(counts=None):
    """อัปเดตตารางเงื่อนไข (PyQt6)"""
    if ui is None:
        return
    table = ui.conditions_table
    table.setRowCount(0)
    for idx, cond in enumerate(saved_conditions, start=1):
        cnt = counts.get(cond, '') if counts is not None else ''
        row = table.rowCount()
        table.insertRow(row)

        id_item = QtWidgets.QTableWidgetItem(str(idx))
        cond_item = QtWidgets.QTableWidgetItem(cond)
        cnt_item = QtWidgets.QTableWidgetItem(str(cnt))

        id_item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        cnt_item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        font_bold = QtGui.QFont()
        font_bold.setBold(True)

        if isinstance(cnt, (int, float)) and cnt > 0:
            cnt_item.setForeground(QtGui.QBrush(QtGui.QColor("#d11f1f")))
            cnt_item.setFont(font_bold)
        elif str(cnt) == "Error":
            cnt_item.setForeground(QtGui.QBrush(QtGui.QColor("#e67e22")))
        elif str(cnt) == "N/A":
            cnt_item.setForeground(QtGui.QBrush(QtGui.QColor("#6b7280")))

        table.setItem(row, 0, id_item)
        table.setItem(row, 1, cond_item)
        table.setItem(row, 2, cnt_item)

def show_help():
    """แสดงคำอธิบายวิธีใช้งาน"""
    if ui:
        ui.show_info("วิธีใช้งาน", HELP_TEXT)

# --- ฟังก์ชัน load_file (เวอร์ชันแก้ไข Log ไม่ให้ซ้ำ) ---
# --- ฟังก์ชัน load_file (เวอร์ชันแก้ไข Log ไม่ให้ซ้ำ และใช้ global UI vars) ---
def load_file():
    """เลือกและโหลดไฟล์ SPSS, อัปเดต Global Vars (รวม Metadata) และ UI"""
    global current_df, original_df_columns_list, lower_df_columns_set, lower_to_original_map, condition_counts, spss_meta, file_var

    if ui is None:
        return

    path = ui.open_file_dialog(
        title="เลือกไฟล์ SPSS",
        filter_text="SPSS files (*.sav);;All files (*.*)"
    )
    if not path:
        add_log("ยกเลิกการเลือกไฟล์")
        return

    file_var.set(path)


    df, meta, error_msg_load = None, None, None # Initialize
    final_exception = None # เก็บ error สุดท้ายที่เจอ
    successful_encoding = None # เก็บ encoding ที่ใช้สำเร็จ

    # ลำดับการลอง Encoding
    encodings_to_try = ['windows-874', 'utf-8', 'tis-620', None] # None คือ default

    # --- Loop ลองหลาย Encoding ---
    for enc in encodings_to_try:
        current_encoding_name = enc if enc is not None else "default (None)"
        add_log(f"กำลังลองโหลด '{os.path.basename(path)}' ด้วย Encoding: {current_encoding_name}...") # Log การลอง
        try:
            df, meta = pyreadstat.read_sav(path, encoding=enc)
            successful_encoding = current_encoding_name
            add_log(f"โหลดไฟล์สำเร็จด้วย Encoding: {successful_encoding}", "SUCCESS") # Log ความสำเร็จ
            final_exception = None # ล้าง error ถ้าสำเร็จ
            break # !!! ออกจากลูปทันทีเมื่อโหลดสำเร็จ !!!

        except UnicodeDecodeError as e_decode:
            add_log(f"Encoding '{current_encoding_name}' ไม่สำเร็จ (Invalid byte sequence). กำลังลอง Encoding ถัดไป...", "WARNING")
            final_exception = e_decode # เก็บ error ล่าสุด (กรณี DecodeError)
            df, meta = None, None # รีเซ็ต df, meta เพื่อลองอันถัดไป
        except Exception as e_other:
             add_log(f"เกิดข้อผิดพลาดอื่นขณะลอง '{current_encoding_name}': {e_other}", "ERROR")
             final_exception = e_other # เก็บ error ล่าสุด (กรณี Error อื่น)
             df, meta = None, None # รีเซ็ต df, meta
             # อาจจะ break ที่นี่เลยถ้าไม่ต้องการลอง Encoding อื่นต่อเมื่อเจอ Error ที่ไม่ใช่ DecodeError

    # --- ตรวจสอบผลลัพธ์หลังจบ Loop ---
    if df is None or meta is None:
        # ถ้าวนจนสุดแล้วยังโหลดไม่ได้
        error_msg_load = f"ไม่สามารถโหลดไฟล์ได้ ลอง Encoding {', '.join(map(str, encodings_to_try))} แล้ว"
        if final_exception:
             error_msg_load += f"\nข้อผิดพลาดล่าสุด: {final_exception}"
    # ไม่ต้องมี else เพราะถ้าสำเร็จ successful_encoding จะถูกตั้งค่า และ df/meta จะมีค่าแล้ว

    # --- ส่วนจัดการผลลัพธ์ UI, global vars (เหมือนเดิม) ---
    if error_msg_load or df is None:
        ui.show_error("โหลดไฟล์ SPSS ล้มเหลว", error_msg_load or "ไม่สามารถโหลด DataFrame ได้")
        # Reset global state on failure
        current_df, original_df_columns_list, lower_df_columns_set, lower_to_original_map = None, [], set(), {}
        condition_counts = {}
        spss_meta = None
        file_var.set("")
        update_table()
    else:
        # โหลดสำเร็จแล้ว...
        current_df = df
        spss_meta = meta
        original_df_columns_list = current_df.columns.tolist()
        lower_df_columns_set = {c.lower() for c in original_df_columns_list}
        lower_to_original_map = {c.lower(): c for c in original_df_columns_list}
        condition_counts = {}
        ui.show_info("สำเร็จ", f"โหลดไฟล์ SPSS:\n{os.path.basename(path)}\nเรียบร้อยแล้ว (ใช้ Encoding: {successful_encoding}) ({len(df)} แถว)")

        # Recalculate counts...
        if saved_conditions:
            counts_result, error_summary = compute_counts(current_df, original_df_columns_list, lower_df_columns_set, lower_to_original_map)
            condition_counts = counts_result if counts_result else {}
            update_table(condition_counts)
            if error_summary:
                ui.show_warning("พบข้อผิดพลาด", f"บางเงื่อนไขมีปัญหาในการประมวลผลกับข้อมูลใหม่:\n{error_summary}")
        else:
            update_table()

# --- สิ้นสุดฟังก์ชัน load_file (เวอร์ชัน Log ไม่ซ้ำ) ---

def save_condition():
    """บันทึกเงื่อนไขและคำนวณ Count เฉพาะรายการใหม่"""
    global current_df, saved_conditions, original_df_columns_list, lower_df_columns_set, lower_to_original_map, condition_counts # เพิ่ม condition_counts
    cond = condition_var.get().strip()
    if not cond:
        ui.show_warning("แจ้งเตือน", "กรุณากรอกเงื่อนไขก่อนบันทึก")
        return

    # Validate condition before saving if data is loaded
    if current_df is not None:
        error_msg = validate_condition(cond, original_df_columns_list, lower_df_columns_set, lower_to_original_map)
        if error_msg:
            ui.show_error("เงื่อนไขไม่ถูกต้อง", f"ไม่สามารถบันทึกเงื่อนไข:\n'{cond}'\n\nข้อผิดพลาด: {error_msg}\n\nกรุณาแก้ไขก่อนบันทึก")
            return

    # Check for duplicates
    if cond in saved_conditions:
         ui.show_info("ข้อมูลซ้ำ", "เงื่อนไขนี้ถูกบันทึกไว้แล้ว")
         condition_var.set('')
         return

    # --- ส่วนที่แก้ไข ---
    # 1. Add condition to the list
    saved_conditions.append(cond)

    # 2. Calculate count ONLY for the new condition (if data is loaded)
    new_count = "N/A" # Default if no data
    if current_df is not None:
         new_count = calculate_single_count(cond, current_df, original_df_columns_list, lower_df_columns_set, lower_to_original_map)

    # 3. Store the new count in the dictionary
    condition_counts[cond] = new_count

    # 4. Update the table using the latest counts dictionary
    update_table(condition_counts)
    # --- สิ้นสุดส่วนที่แก้ไข ---
    condition_var.set('') # Clear input field

def delete_condition():
    """ลบเงื่อนไขที่เลือกและอัปเดต UI"""
    global current_df, saved_conditions, condition_counts # เพิ่ม condition_counts

    if ui is None:
        return

    conditions_to_delete_texts = ui.get_selected_conditions()
    if not conditions_to_delete_texts:
        ui.show_warning("ไม่ได้เลือก", "กรุณาเลือกเงื่อนไขที่ต้องการลบ")
        return

    confirm_msg = f"ต้องการลบ {len(conditions_to_delete_texts)} เงื่อนไข?\n\n" + "\n".join([f"- {c[:70]}{'...' if len(c)>70 else ''}" for c in conditions_to_delete_texts[:5]]) + ("\n..." if len(conditions_to_delete_texts)>5 else "")
    if not ui.ask_yes_no("ยืนยันการลบ", confirm_msg): return

    # --- ส่วนที่แก้ไข ---
    # 1. Remove from saved_conditions list
    original_count = len(saved_conditions)
    saved_conditions = [c for c in saved_conditions if c not in conditions_to_delete_texts]
    deleted_count = original_count - len(saved_conditions)

    # 2. Remove corresponding entries from condition_counts dictionary
    for cond_text in conditions_to_delete_texts:
         condition_counts.pop(cond_text, None) # ใช้ pop(key, None) เพื่อไม่ให้เกิด error ถ้า key ไม่มีอยู่แล้ว

    # 3. Update UI using the modified condition_counts
    update_table(condition_counts)
    # --- สิ้นสุดส่วนที่แก้ไข ---

    ui.show_info("สำเร็จ", f"ลบเงื่อนไข {deleted_count} รายการ")
    # ไม่ต้องคำนวณ count ใหม่ หรือแจ้งเตือน error การคำนวณ

def export_conditions():
    """ส่งออกเงื่อนไขที่บันทึกไว้เป็นไฟล์ Excel"""
    if not saved_conditions:
        ui.show_warning("แจ้งเตือน", "ไม่มีเงื่อนไขให้บันทึก")
        return
    path = ui.save_file_dialog(defaultextension='.xlsx', filetypes=[("Excel files", "*.xlsx")], title="บันทึกไฟล์เงื่อนไข")
    if not path: return # User cancelled

    try:
        # Create DataFrame with ID and Condition
        export_df = pd.DataFrame({
            'ID': range(1, len(saved_conditions) + 1),
            'Condition': saved_conditions
            })
        export_df.to_excel(path, index=False)
        ui.show_info("สำเร็จ", f"บันทึกเงื่อนไข {len(saved_conditions)} รายการไปยัง:\n{os.path.basename(path)}\nเรียบร้อยแล้ว")
    except Exception as e:
        ui.show_error("Error", f"บันทึกไฟล์เงื่อนไขล้มเหลว: {e}")
# --- ฟังก์ชัน import_conditions ที่อัปเดตแล้ว พร้อม Progress Bar ตอนคำนวณ Count ---
def import_conditions():
    """โหลดเงื่อนไขจากไฟล์ Excel, ล้าง Counts เดิม, และคำนวณใหม่พร้อม Progress Bar (ถ้ามีข้อมูล)"""
    # Access global variables needed
    global current_df, saved_conditions, original_df_columns_list, lower_df_columns_set, lower_to_original_map, condition_counts

    path = ui.open_file_dialog(filetypes=[("Excel files", "*.xlsx")], title="เลือกไฟล์เงื่อนไข")
    if not path: return # User cancelled

    try:
        df_import = pd.read_excel(path, header=0) # Assume header is on the first row

        # Determine the column to use ('Condition' preferably, else the first one)
        if 'Condition' in df_import.columns:
            col_name = 'Condition'
        elif len(df_import.columns) > 0:
            col_name = df_import.columns[0]
            ui.show_warning("รูปแบบไฟล์", f"ไม่พบคอลัมน์ 'Condition', กำลังใช้คอลัมน์แรก '{col_name}' แทน")
        else:
             ui.show_error("Error", "ไฟล์ Excel ไม่มีข้อมูลหรือไม่มีคอลัมน์ที่รู้จัก")
             return

        # Read conditions, convert to string, drop NaN, strip whitespace, remove empty
        conds_from_file = [str(c).strip() for c in df_import[col_name].dropna() if str(c).strip()]

        if not conds_from_file:
            ui.show_info("สำเร็จ", "ไฟล์ที่เลือกไม่มีเงื่อนไขที่สามารถนำเข้าได้")
            return

        # --- Validate imported conditions if data is loaded ---
        valid_conds = []
        import_errors = []
        if current_df is not None:
            # Validate each condition read from the file
            for c in conds_from_file:
                 error_msg = validate_condition(c, original_df_columns_list, lower_df_columns_set, lower_to_original_map)
                 if error_msg:
                     # Collect errors for invalid conditions
                     import_errors.append(f"- '{c}': {error_msg}")
                 else:
                     # Add valid conditions to the list
                     valid_conds.append(c)
        else:
            # If no data loaded, assume all conditions are potentially valid for now
            # Validation will happen later when data is loaded or check_conditions is run
            valid_conds = conds_from_file

        # --- Show errors for invalid conditions found during validation ---
        if import_errors:
            ui.show_warning("เงื่อนไขบางรายการไม่ถูกต้อง",
                                   f"เงื่อนไขต่อไปนี้จากไฟล์นำเข้าไม่ถูกต้องและจะถูกข้าม:\n" +
                                   "\n".join(import_errors) +
                                   f"\n\nดำเนินการต่อด้วยเงื่อนไขที่ถูกต้อง {len(valid_conds)} รายการ")

        # Stop if no valid conditions remain after validation
        if not valid_conds:
            ui.show_error("นำเข้าล้มเหลว", "ไม่มีเงื่อนไขที่ถูกต้องในไฟล์ที่เลือกหลังจากตรวจสอบ (ถ้าทำได้)")
            return

        # --- Ask user whether to replace or append conditions ---
        choice = ui.ask_yes_no_cancel("นำเข้าเงื่อนไข",
                                          f"พบ {len(valid_conds)} เงื่อนไขที่ถูกต้องในไฟล์\n"
                                          "ต้องการแทนที่เงื่อนไขที่มีอยู่ทั้งหมดหรือไม่?\n"
                                          "(Yes=แทนที่, No=ต่อท้ายรายการเดิม, Cancel=ยกเลิก)")

        if choice is None: return # User cancelled

        # --- Clear existing counts because conditions list is changing ---
        condition_counts = {}

        # --- Update saved_conditions list based on user choice ---
        if choice: # Yes = Replace
             saved_conditions = valid_conds
             op_text = "แทนที่"
        else: # No = Append
             added_count = 0
             for vc in valid_conds:
                 if vc not in saved_conditions: # Avoid adding duplicates
                     saved_conditions.append(vc)
                     added_count +=1
             op_text = f"เพิ่มเงื่อนไขใหม่ {added_count} รายการ"

        # --- Recalculate all counts WITH PROGRESS BAR (if data loaded) ---
        counts_result = {} # Dictionary to store results of this calculation
        error_summary_list = [] # List to collect error messages

        # Only calculate counts if data is actually loaded and there are conditions
        if current_df is not None and saved_conditions:
            total_conditions_to_calc = len(saved_conditions)
            # --- Setup Progress Bar ---
            ui.progress_bar.setMaximum(total_conditions_to_calc)
            ui.progress_bar.setValue(0)
            progress_status_var.set(f"Calculating counts (0/{total_conditions_to_calc})...")
            # Optional: Disable relevant buttons (Need references to buttons)
            # e.g., import_button.config(state=tk.DISABLED)
            # e.g., check_button_widget.config(state=tk.DISABLED)
            process_events() # Show initial progress state

            try: # Use finally to ensure UI elements are reset
                # --- Loop to calculate counts for all conditions ---
                for i, cond in enumerate(saved_conditions):
                    current_status = "Error" # Default status for progress bar
                    try:
                        # Perform steps: Validate -> Expand -> Convert -> Query -> Get Count
                        error_msg = validate_condition(cond, original_df_columns_list, lower_df_columns_set, lower_to_original_map)
                        if error_msg: raise ValueError(f"รูปแบบผิด: {error_msg}")

                        expanded = expand_wildcard(cond, original_df_columns_list, lower_to_original_map)
                        converted = auto_convert(expanded, lower_to_original_map)
                        count = len(current_df.query(converted))
                        counts_result[cond] = count
                        current_status = count # Update status for progress display

                    except Exception as e:
                        # Store error if calculation fails for this condition
                        counts_result[cond] = "Error"
                        error_summary_list.append(f"- '{cond}': {e}")
                        current_status = "Error"

                    # --- Update Progress Bar and Status Label ---
                    ui.progress_bar.setValue(i + 1)
                    # Simplified status message as requested previously
                    progress_status_var.set(f"Calculating counts {i + 1}/{total_conditions_to_calc}")
                    process_events() # IMPORTANT: Update UI within the loop
                # --- End Calculation Loop ---

            finally:
                # --- Reset Progress Bar and Status ---
                ui.progress_bar.setValue(0)
                progress_status_var.set("Idle")
                # Optional: Re-enable buttons
                # e.g., import_button.config(state=tk.NORMAL)
                # e.g., check_button_widget.config(state=tk.NORMAL)
                process_events() # Ensure final UI state is shown

            # Update global counts dictionary with the fresh results
            condition_counts = counts_result

        # --- Update the UI table ---
        # Pass the newly calculated counts (or the empty dict if no data/no conditions)
        update_table(condition_counts)
        ui.show_info("สำเร็จ", f"{op_text}เงื่อนไขเรียบร้อยแล้ว")

        # Show errors encountered during recalculation (if any)
        if error_summary_list:
             error_summary_str = "\n".join(error_summary_list)
             ui.show_warning("พบข้อผิดพลาด", f"พบข้อผิดพลาดในการคำนวณ Count หลังนำเข้า:\n{error_summary_str}")

    # --- Exception Handling for File Operations ---
    except FileNotFoundError:
         ui.show_error("Error", f"ไม่พบไฟล์ที่ระบุ: {path}")
    except pd.errors.EmptyDataError:
         ui.show_error("Error", f"ไฟล์ Excel '{os.path.basename(path)}' ว่างเปล่า หรือไม่มีข้อมูลในคอลัมน์ที่ต้องการ")
    except Exception as e:
        # Catch other potential errors during file reading or processing
        # import traceback
        # ui.show_error("Error", f"โหลดไฟล์เงื่อนไขล้มเหลว: {e}\n{traceback.format_exc()}") # Detailed error for debug
        ui.show_error("Error", f"โหลดไฟล์เงื่อนไขล้มเหลว: {e}")
    except BaseException as be:
        add_log(f"❌ เกิดข้อผิดพลาดระดับระบบระหว่างนำเข้าเงื่อนไข: {be}", "ERROR")
        add_log(traceback.format_exc(), "ERROR")
        _report_fatal_error("Error", f"โหลดไฟล์เงื่อนไขล้มเหลว (BaseException): {be}")
# --- สิ้นสุดฟังก์ชัน import_conditions ---



# --- ฟังก์ชัน check_conditions ที่ปรับปรุงใหม่ทั้งหมดเพื่อความเร็วสูงสุด ---
# --- ฟังก์ชัน check_conditions ที่ปรับปรุงใหม่ทั้งหมดเพื่อความเร็วสูงสุด ---
def check_conditions():
    """
    (เวอร์ชันปรับปรุงใหม่) Export ข้อมูลโดยเชื่อค่า Count ที่มีอยู่แล้ว
    และ Query ข้อมูลเฉพาะเงื่อนไขที่มี Count > 0 เท่านั้น
    """
    # Access global variables
    global current_df, saved_conditions, original_df_columns_list, lower_df_columns_set, lower_to_original_map, condition_counts

    # --- Pre-checks ---
    if current_df is None:
        ui.show_warning("แจ้งเตือน", "กรุณาเลือกไฟล์ SPSS ก่อน")
        return
    if not saved_conditions:
        ui.show_warning("แจ้งเตือน", "ไม่มีเงื่อนไขให้ตรวจสอบ")
        return

    # --- 1. ค้นหาเงื่อนไขที่ต้อง Export (Count > 0) จากค่าที่คำนวณไว้แล้ว ---
    conditions_to_export = [
        cond for cond, count in condition_counts.items()
        if isinstance(count, (int, np.integer)) and count > 0
    ]

    if not conditions_to_export:
        ui.show_info("ไม่มีข้อมูลสำหรับ Export", "ไม่พบเงื่อนไขใดๆ ที่มีผลลัพธ์ (Count > 0)")
        return

    total_to_process = len(conditions_to_export)
    add_log(f"พบ {total_to_process} เงื่อนไขที่มีข้อมูล จะเริ่มเตรียมการ Export...")

    # --- 2. Progress Bar Setup (สำหรับขั้นตอนการเตรียม Export) ---
    ui.progress_bar.setMaximum(total_to_process)
    ui.progress_bar.setValue(0)
    progress_status_var.set(f"Preparing to export {total_to_process} conditions...")
    
    check_button = check_button_widget # ใช้ global widget ที่ประกาศไว้
    set_button_enabled(check_button, False)
    process_events()

    # --- 3. Initialize results (สำหรับ Export) ---
    result_dict = {}  # Dictionary สำหรับเก็บ DataFrame ที่จะ Export
    errors_dict = {}  # Dictionary สำหรับเก็บข้อผิดพลาดที่เกิดระหว่างการ Query

    # --- 4. Use finally block for cleanup ---
    try:
        # --- Main Processing Loop (วนลูปเฉพาะเงื่อนไขที่จะ Export) ---
        for i, cond in enumerate(conditions_to_export):
            current_progress = i + 1
            progress_status_var.set(f"Preparing Export {current_progress}/{total_to_process}: {cond[:50]}...")
            ui.progress_bar.setValue(current_progress)
            process_events()

            try:
                # ทำการ Validate -> Expand -> Convert เพื่อสร้าง query string
                # ขั้นตอนนี้ยังจำเป็นเพื่อให้ได้ query string ที่ถูกต้องสำหรับดึงข้อมูล
                error_msg = validate_condition(cond, original_df_columns_list, lower_df_columns_set, lower_to_original_map)
                if error_msg:
                    raise ValueError(f"รูปแบบไม่ถูกต้อง (ตอนเตรียม Export): {error_msg}")

                expanded = expand_wildcard(cond, original_df_columns_list, lower_to_original_map)
                converted = auto_convert(expanded, lower_to_original_map)

                # Query DataFrame เพื่อดึงผลลัพธ์
                sub_df = current_df.query(converted)

                # ดึงรายชื่อคอลัมน์ที่ต้องการสำหรับเงื่อนไขนี้
                cols_extract = extract_cols_from_raw_condition(cond, original_df_columns_list, lower_df_columns_set, lower_to_original_map)

                # เก็บ DataFrame ที่เลือกคอลัมน์แล้ว สำหรับการ Export
                result_dict[cond] = sub_df[cols_extract]
                add_log(f"  ✓ เตรียมข้อมูล '{cond}' สำเร็จ ({len(sub_df)} รายการ)")

            except Exception as e:
                # หากเกิดข้อผิดพลาดระหว่างการ Query ของเงื่อนไขนี้
                error_message_str = f"เกิดข้อผิดพลาดขณะเตรียมข้อมูล '{cond}' เพื่อ Export: {e}"
                errors_dict[cond] = error_message_str
                add_log(f"  ✗ {error_message_str}", "ERROR")
                # อัปเดตตาราง UI ให้เห็นว่าเงื่อนไขนี้มีปัญหา
                condition_counts[cond] = "Export Error"
                update_table(condition_counts) # อัปเดตทันที
        # --- End of Main Loop ---

        # --- 5. สรุปผลและเริ่ม Export ---
        if errors_dict:
            error_summary = "\n".join([f"- {m}" for c, m in errors_dict.items()])
            ui.show_warning("พบข้อผิดพลาด", f"เกิดข้อผิดพลาดในการเตรียมข้อมูลบางรายการ:\n{error_summary}\n\nจะทำการ Export เฉพาะรายการที่สำเร็จ")

        if not result_dict:
             ui.show_error("Export ล้มเหลว", "ไม่สามารถเตรียมข้อมูลสำหรับ Export ได้เลยแม้แต่รายการเดียว")
             return

        # --- เริ่ม Export ไปยัง Excel ---
        add_log("เตรียมข้อมูลเสร็จสิ้น กำลังส่งออกไปยัง Excel...")
        progress_status_var.set("Exporting to Excel...")
        process_events()
        filename_base = os.path.splitext(os.path.basename(file_var.get()))[0] if file_var.get() else "SPSS_Check"
        # ส่ง result_dict (ที่เก็บ DataFrame) และ condition_counts (สำหรับแสดงค่า Count ใน Index)
        open_multi_excel(result_dict, condition_counts, filename_base=f"{filename_base}_CheckResult")

    finally:
        # --- Reset Progress Bar & Button ---
        ui.progress_bar.setValue(0)
        progress_status_var.set("Idle")
        if check_button:
            set_button_enabled(check_button, True)
        process_events()

# --- สิ้นสุดฟังก์ชัน check_conditions (เวอร์ชันปรับปรุง) ---

def calculate_single_count(condition, df, original_cols, lower_cols_set, lower_to_orig_map):
    """คำนวณ Count สำหรับเงื่อนไขเดียว (สำหรับใช้ตอน Save)"""
    if df is None:
        return "N/A"
    try:
        # Validate ก่อน (ถึงแม้ save_condition จะ validate แล้วก็ตาม เพื่อความสมบูรณ์)
        error_msg = validate_condition(condition, original_cols, lower_cols_set, lower_to_orig_map)
        if error_msg:
            print(f"Validation Error (single): '{condition}' -> {error_msg}") # Log validation error
            return "Error"

        expanded = expand_wildcard(condition, original_cols, lower_to_orig_map)
        converted = auto_convert(expanded, lower_to_orig_map)
        count = len(df.query(converted))
        return count

    except Exception as e:
        print(f"Error calculating single count for '{condition}': {e}") # Log processing error
        # import traceback
        # print(traceback.format_exc()) # Optional detailed traceback
        return "Error"


def select_variables_for_frequency():
    """
    เปิดหน้าต่างใหม่สำหรับเลือกตัวแปรที่ต้องการทำ Frequency Analysis
    แสดงรายการทั้งหมดตามลำดับในไฟล์ SPSS ต้นฉบับ
    """
    if current_df is None:
        ui.show_warning("ไม่มีข้อมูล", "กรุณาเลือกไฟล์ SPSS ก่อน")
        return None

    class VarTable(QtWidgets.QTableWidget):
        def keyPressEvent(self, event):
            if event.key() == QtCore.Qt.Key.Key_Delete:
                rows = sorted({i.row() for i in self.selectedItems()}, reverse=True)
                for r in rows:
                    self.removeRow(r)
            else:
                super().keyPressEvent(event)

    dialog = QtWidgets.QDialog(ui)
    dialog.setWindowTitle("เลือกตัวแปรที่ต้องการวิเคราะห์ Frequency Analysis")
    dialog.resize(720, 640)
    layout = QtWidgets.QVBoxLayout(dialog)

    title = QtWidgets.QLabel("เลือกตัวแปรที่ต้องการวิเคราะห์:")
    title.setStyleSheet("font-weight:700;")
    layout.addWidget(title)

    search_row = QtWidgets.QHBoxLayout()
    search_row.addWidget(QtWidgets.QLabel("ค้นหา:"))
    search_input = QtWidgets.QLineEdit()
    search_input.setPlaceholderText("พิมพ์เพื่อค้นหา...")
    search_row.addWidget(search_input, 1)
    layout.addLayout(search_row)

    table = VarTable()
    table.setColumnCount(2)
    table.setHorizontalHeaderLabels(["ชื่อตัวแปร", "จำนวนคอลัมน์"])
    table.horizontalHeader().setStretchLastSection(True)
    table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
    table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)
    table.verticalHeader().setVisible(False)
    layout.addWidget(table, 1)

    hint = QtWidgets.QLabel("กด Delete เพื่อตัดรายการที่ไม่ต้องการออก")
    hint.setStyleSheet("color:#6b7280;font-style:italic;")
    layout.addWidget(hint)

    btn_row = QtWidgets.QHBoxLayout()
    restore_btn = QtWidgets.QPushButton("เลือกทั้งหมด")
    clear_btn = QtWidgets.QPushButton("ยกเลิกทั้งหมด")
    btn_row.addWidget(restore_btn)
    btn_row.addWidget(clear_btn)
    btn_row.addStretch(1)
    ok_btn = QtWidgets.QPushButton("OK")
    cancel_btn = QtWidgets.QPushButton("Cancel")
    ok_btn.setStyleSheet("background:#16a34a;color:white;font-weight:700;")
    cancel_btn.setStyleSheet("background:#ef4444;color:white;font-weight:700;")
    btn_row.addWidget(cancel_btn)
    btn_row.addWidget(ok_btn)
    layout.addLayout(btn_row)

    # เตรียมข้อมูลสำหรับแสดง
    all_columns = current_df.columns.tolist()
    ma_sets = {}
    ma_pattern = re.compile(r'(.+)_O(\d+)$', flags=re.IGNORECASE)
    for col_name in all_columns:
        match = ma_pattern.match(col_name)
        if match:
            base_name = match.group(1).upper() + "_O"
            ma_sets.setdefault(base_name, []).append(col_name)

    display_items = []
    processed_cols = set()
    for col_name in all_columns:
        if col_name in processed_cols:
            continue
        match = ma_pattern.match(col_name)
        if match:
            base_name = match.group(1).upper() + "_O"
            if base_name in ma_sets and len(ma_sets[base_name]) > 1:
                display_items.append({
                    "name": base_name,
                    "count": f"({len(ma_sets[base_name])} columns)",
                    "type": "ma_set"
                })
                processed_cols.update(ma_sets[base_name])
            else:
                display_items.append({"name": col_name, "count": "", "type": "single_col"})
                processed_cols.add(col_name)
        else:
            display_items.append({"name": col_name, "count": "", "type": "single_col"})
            processed_cols.add(col_name)

    def fill_table(items):
        table.setRowCount(0)
        for item in items:
            row = table.rowCount()
            table.insertRow(row)
            name_item = QtWidgets.QTableWidgetItem(item["name"])
            count_item = QtWidgets.QTableWidgetItem(item["count"])
            if item["type"] == "ma_set":
                name_item.setBackground(QtGui.QColor("#f3f4f6"))
                count_item.setBackground(QtGui.QColor("#f3f4f6"))
            table.setItem(row, 0, name_item)
            table.setItem(row, 1, count_item)
        table.resizeColumnsToContents()

    fill_table(display_items)

    def filter_items():
        text = search_input.text().lower().strip()
        for row in range(table.rowCount()):
            name = table.item(row, 0).text().lower()
            table.setRowHidden(row, text not in name)

    def restore_all():
        fill_table(display_items)
        filter_items()

    def clear_all():
        table.setRowCount(0)

    def on_ok():
        dialog.accept()

    def on_cancel():
        dialog.reject()

    search_input.textChanged.connect(filter_items)
    restore_btn.clicked.connect(restore_all)
    clear_btn.clicked.connect(clear_all)
    ok_btn.clicked.connect(on_ok)
    cancel_btn.clicked.connect(on_cancel)

    result = dialog.exec()
    if result != QtWidgets.QDialog.DialogCode.Accepted:
        return None

    result_vars = []
    for row in range(table.rowCount()):
        if table.isRowHidden(row):
            continue
        item = table.item(row, 0)
        if item:
            result_vars.append(item.text())

    return result_vars


def run_all_frequencies():
    """
    คำนวณ Frequency Table, รวม MA Sets, ไม่แสดง Missing/Empty,
    สร้าง Index Sheet ที่มี **ครบทุกรายการที่เลือก**, และ Export Excel
    (เวอร์ชันแก้ไข: จัดการ TypeError ตอนกรองข้อมูล MA Set)
    """
    # Access global variables
    global current_df, spss_meta, file_var, progress_status_var
    # Make sure button widgets are accessible if declared globally or passed as args
    global check_button_widget, freq_button_widget

    # --- การตรวจสอบเบื้องต้น ---
    if current_df is None:
        ui.show_warning("ไม่มีข้อมูล", "กรุณาเลือกไฟล์ SPSS ก่อน")
        return

    # --- เปิดหน้าต่างเลือกคอลัมน์ ---
    selected_columns = select_variables_for_frequency()
    if selected_columns is None:
        add_log("ยกเลิกการทำงาน: ผู้ใช้ยกเลิกการเลือกคอลัมน์")
        return

    if not selected_columns:
        ui.show_warning("ไม่มีคอลัมน์ที่เลือก", "กรุณาเลือกอย่างน้อย 1 ตัวแปร หรือ MA Set")
        return

    # --- ยืนยันการทำงาน ---
    confirm_msg = f"คุณกำลังจะทำ Frequency Analysis บนตัวแปร/MA Set ที่เลือก {len(selected_columns)} รายการ"
    if len(selected_columns) > 50:
          confirm_msg += f"\n\nคำเตือน: การเลือก {len(selected_columns)} รายการอาจใช้เวลานาน"
    confirm_msg += "\n\nต้องการดำเนินการต่อหรือไม่?"

    if not ui.ask_yes_no("ยืนยันการทำงาน", confirm_msg):
        add_log("ยกเลิกการทำงาน: ผู้ใช้ไม่ยืนยันการดำเนินการ")
        return

    ui.show_info("เริ่มคำนวณ", f"กำลังเตรียมคำนวณ Frequency Tables สำหรับ {len(selected_columns)} รายการที่เลือก...");
    process_events()

    # --- ล้าง Log ---
    # Check if log_text exists and is a valid widget before clearing
    if ui and ui.log_text:
        ui.log_text.clear()
    else:
        print("Warning: log_text widget not found or not initialized.")


    # --- 1. ระบุ MA Sets & Single Columns จากข้อมูลทั้งหมด ---
    add_log("===== เริ่มการทำงานของฟังก์ชัน Frequency Analysis =====")
    add_log(f"เวลาเริ่มต้น: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
    all_columns = current_df.columns.tolist(); ma_sets = {}; processed_ma_cols = set()
    ma_pattern = re.compile(r'(.+)_O(\d+)$', flags=re.IGNORECASE); potential_ma_bases = {}
    add_log(f"จำนวนคอลัมน์ทั้งหมดในข้อมูล: {len(all_columns)}")
    add_log("กำลังระบุคอลัมน์ Multiple Answer Sets จากข้อมูลทั้งหมด...")

    for col_name in all_columns:
        match = ma_pattern.match(col_name)
        if match:
            base_name = match.group(1).upper() + "_O"
            potential_ma_bases.setdefault(base_name, []).append(col_name)

    for base_name, cols in potential_ma_bases.items():
        if len(cols) > 1:
            def get_suffix_num(c): m = ma_pattern.match(c); return int(m.group(2)) if m else float('inf')
            ma_sets[base_name] = sorted(cols, key=get_suffix_num)
            processed_ma_cols.update(ma_sets[base_name])

    single_cols = [col for col in all_columns if col not in processed_ma_cols]
    add_log(f"ระบุ MA Sets ทั้งหมด: {len(ma_sets)} sets")
    add_log(f"ระบุคอลัมน์เดี่ยวทั้งหมด: {len(single_cols)}")

    # --- 2. ตั้งค่า Progress Bar และ Disable Buttons ---
    total_items_to_process = len(selected_columns)
    if ui and ui.progress_bar:
        ui.progress_bar.setMaximum(total_items_to_process)
        ui.progress_bar.setValue(0)
    if progress_status_var: progress_status_var.set(f"Running Frequencies (0/{total_items_to_process})...")

    # Disable buttons - Use try-except in case widgets haven't been assigned yet
    buttons_to_disable = []
    if 'check_button_widget' in globals() and check_button_widget: buttons_to_disable.append(check_button_widget)
    if 'freq_button_widget' in globals() and freq_button_widget: buttons_to_disable.append(freq_button_widget)

    try:
        for btn in buttons_to_disable:
            set_button_enabled(btn, False)
    except Exception as e_disable:
        add_log(f"Warning: Error disabling buttons: {e_disable}", "WARNING")
        print(f"Warning: Error disabling buttons: {e_disable}")

    process_events()

    # --- 3. คำนวณ Frequency (วนลูปตาม selected_columns) ---
    frequency_results = {}; errors_freq = {};
    calculation_progress_count = 0
    add_log("\n===== เริ่มกระบวนการคำนวณ Frequency Tables (ตามรายการที่เลือก) =====")

    # Use a finally block to ensure buttons are re-enabled
    try:
        for item_name in selected_columns:
            calculation_progress_count += 1
            if progress_status_var: progress_status_var.set(f"Calculating ({calculation_progress_count}/{total_items_to_process}): {item_name[:40]}...")
            process_events()

            is_ma_set_calc = item_name in ma_sets

            if is_ma_set_calc:
                # ---------------------------------
                # --- คำนวณ MA Set ---
                # ---------------------------------
                cols_for_this_item = ma_sets[item_name]
                add_log(f"[{calculation_progress_count}/{total_items_to_process}] คำนวณ MA Set: {item_name} ({len(cols_for_this_item)} คอลัมน์)")
                try:
                    # ดึง Variable Label และ Value Labels (เหมือนเดิม)
                    var_label_ma = item_name; value_labels_dict_ma = {}
                    if spss_meta and cols_for_this_item:
                        first_col_label = spss_meta.column_names_to_labels.get(cols_for_this_item[0], item_name)
                        # Remove trailing option numbers/suffixes for a cleaner MA set label
                        var_label_ma = re.sub(r'[:\s]*[Oo]ption\s*\d+$|\s*\d+$', '', first_col_label).strip()
                        # Aggregate value labels from all columns in the set
                        for col in cols_for_this_item: value_labels_dict_ma.update(spss_meta.variable_value_labels.get(col, {}))
                    defined_codes_list = list(value_labels_dict_ma.keys())

                    # เตรียมข้อมูล
                    ma_data_subset = current_df[cols_for_this_item]
                    # นับ Base N (จำนวนผู้ตอบที่ตอบอย่างน้อย 1 ข้อในชุดนี้)
                    valid_rows_ma = ma_data_subset.apply(lambda col: col.notna() & (col.astype(str).str.strip() != ''), axis=0).any(axis=1)
                    valid_respondent_count = valid_rows_ma.sum()

                    # Melt ข้อมูล
                    melted_data = ma_data_subset.melt(var_name='_col_origin', value_name='Code')

                    # --- vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv ---
                    # --- ส่วนแก้ไขการกรองข้อมูล MA Set (ป้องกัน TypeError) ---
                    # --- vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv ---
                    codes_initial = melted_data['Code'] # เริ่มจากคอลัมน์ Code ดิบ

                    # 1. สร้าง Mask กรองค่าที่ไม่ใช่ NaN และไม่ใช่ None ออกไปก่อน
                    valid_mask = codes_initial.notna()
                    valid_codes_series = codes_initial[valid_mask]

                    # 2. เตรียม Series สำหรับการดำเนินการกับ String
                    #    แปลงค่าที่เหลือ (ที่ไม่ใช่ NaN/None) เป็น String
                    #    ใช้ try-except เผื่อกรณีมี Data Type แปลกๆ ที่แปลงไม่ได้
                    observed_codes_series = pd.Series(dtype='object') # สร้าง Series ว่างรอไว้
                    try:
                        string_series = valid_codes_series.astype(str)

                        # 3. สร้าง Mask กรอง String ที่ไม่ว่างเปล่า (หลังจาก strip)
                        non_empty_mask = string_series.str.strip() != ''

                        # 4. ใช้ Mask นี้กรอง Series ต้นฉบับ (valid_codes_series)
                        #    เพื่อรักษา Data Type เดิมไว้ ถ้าเป็นไปได้
                        #    reindex เพื่อให้แน่ใจว่า index ตรงกันก่อน filter
                        final_mask = non_empty_mask.reindex(valid_codes_series.index, fill_value=False)
                        observed_codes_series = valid_codes_series[final_mask]

                    except Exception as e_str_conv_filter:
                        # หากเกิด Error ตอนแปลงหรือกรอง String ให้ Log เตือนและใช้ Series ว่างแทน
                        add_log(f"  ⚠ Warning: Error during string conversion/filtering for {item_name}: {e_str_conv_filter}", "WARNING")
                        print(f"ERROR during string conversion/filtering for {item_name}: {e_str_conv_filter}")
                        # Fallback to empty series is already done by initializing observed_codes_series above
                    # --- ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ ---
                    # --- จบส่วนแก้ไขการกรองข้อมูล MA Set                      ---
                    # --- ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ ---

                    # นับจำนวนของแต่ละ Code ที่เหลืออยู่
                    observed_counts = observed_codes_series.value_counts()

                    # สร้างตารางผลลัพธ์ (ส่วนนี้เหมือนเดิม)
                    results_data = []
                    # เรียงลำดับ Code ที่กำหนดไว้ใน Value Labels
                    try:
                        numeric_defined = sorted([c for c in defined_codes_list if isinstance(c,(int,float))], key=float)
                        string_defined = sorted([str(c) for c in defined_codes_list if not isinstance(c,(int,float))], key=str)
                        sorted_defined_codes = numeric_defined + string_defined
                    except: sorted_defined_codes = sorted(defined_codes_list, key=lambda x: str(x))

                    # เพิ่มข้อมูลแต่ละ Code ลงในตาราง
                    for code in sorted_defined_codes:
                        label = value_labels_dict_ma.get(code, str(code))
                        count = observed_counts.get(code, 0)
                        # *** เราจะแสดงทุก Code ที่ Define ไว้ แม้ Count เป็น 0 ***
                        results_data.append({'Code': code, 'ข้อ': label, 'Base N': count})

                    # สร้าง DataFrame และคำนวณ %
                    if not results_data and valid_respondent_count == 0:
                        # กรณีไม่มีทั้ง defined codes และไม่มีผู้ตอบ
                        freq_table_ma = pd.DataFrame(columns=['Code','ข้อ','Base N','Base %'])
                    elif not results_data and valid_respondent_count > 0:
                         # กรณีมีผู้ตอบ แต่ไม่มี defined code หรือ observed code เลย
                        freq_table_ma = pd.DataFrame(columns=['Code','ข้อ','Base N','Base %'])
                        add_log(f"  ⚠ Warning: MA Set {item_name} มีผู้ตอบ {valid_respondent_count} แต่ไม่พบรหัสที่กำหนด/สังเกตได้", "WARNING")
                    else:
                        # กรณีมีข้อมูล
                        freq_table_ma = pd.DataFrame(results_data)
                        # คำนวณ % เทียบกับจำนวนผู้ตอบ (valid_respondent_count)
                        freq_table_ma['Base %'] = (freq_table_ma['Base N'] / valid_respondent_count * 100) if valid_respondent_count > 0 else 0.0

                    # เพิ่มแถว Total (Base)
                    total_row_ma = pd.DataFrame({'Code': ['Base Total'], 'ข้อ': [''], 'Base N': [valid_respondent_count], 'Base %': [100.0 if valid_respondent_count > 0 else 0.0]})
                    freq_table_ma = total_row_ma if freq_table_ma.empty else pd.concat([total_row_ma, freq_table_ma], ignore_index=True)

                    # สร้าง Label สุดท้ายสำหรับตาราง
                    var_label_ma_final = f"{var_label_ma} (Base: {valid_respondent_count} respondents who answered)"

                    # เก็บผลลัพธ์
                    frequency_results[item_name] = (var_label_ma_final, freq_table_ma)
                    add_log(f"    ✓ คำนวณ MA Set สำเร็จ: {item_name} (Base N: {valid_respondent_count})", "SUCCESS")

                except Exception as e:
                    # จัดการ Error ที่อาจเกิดขึ้นระหว่างคำนวณ MA Set นี้
                    error_message = f"Error calculating MA set '{item_name}': {e}"
                    add_log(f"    ✗ ข้อผิดพลาด MA Set: {error_message}", "ERROR"); print(error_message)
                    errors_freq[item_name] = str(e)
                    # import traceback # Optional for detailed debug
                    # print(traceback.format_exc()) # Optional

            elif item_name in single_cols:
                # ---------------------------------
                # --- คำนวณ Single Column ---
                # ---------------------------------
                add_log(f"[{calculation_progress_count}/{total_items_to_process}] คำนวณคอลัมน์เดี่ยว: {item_name}")
                try:
                    # ดึง Label และ Value Labels (เหมือนเดิม)
                    variable_label = item_name; value_labels_dict = {}; defined_codes = []
                    if spss_meta:
                        variable_label = spss_meta.column_names_to_labels.get(item_name, item_name)
                        value_labels_dict = spss_meta.variable_value_labels.get(item_name, {})
                        defined_codes = list(value_labels_dict.keys())

                    # กรองข้อมูลที่ไม่ใช่ Missing และไม่ใช่ String ว่าง
                    valid_data = current_df[item_name].dropna()
                    if pd.api.types.is_object_dtype(valid_data.dtype) or pd.api.types.is_string_dtype(valid_data.dtype):
                        valid_data = valid_data[valid_data.astype(str).str.strip() != '']

                    # นับจำนวนข้อมูลแต่ละค่า
                    observed_counts_series = valid_data.value_counts().sort_index()

                    # รวบรวม Code ทั้งหมดที่ควรแสดง (ทั้งที่ Define และที่เจอจริง)
                    all_valid_codes_set = set(observed_counts_series.index)
                    for code in defined_codes:
                        if pd.notna(code) and str(code).strip() != '': all_valid_codes_set.add(code)

                    # เรียงลำดับ Code ที่จะแสดง
                    try:
                        numeric_codes = sorted([c for c in all_valid_codes_set if isinstance(c,(int,float))], key=float)
                        string_codes = sorted([str(c) for c in all_valid_codes_set if not isinstance(c,(int,float))], key=str)
                        sorted_unique_codes = numeric_codes + string_codes
                    except: sorted_unique_codes = sorted(list(all_valid_codes_set), key=lambda x: str(x))

                    # สร้าง Series ที่มี Index ครบทุก Code (รวม Code ที่มี Count=0)
                    try:
                        if observed_counts_series.empty and not sorted_unique_codes:
                            freq_series_full = pd.Series(dtype='int64') # ไม่มีข้อมูลเลย
                        elif observed_counts_series.empty and sorted_unique_codes:
                            # มี defined code แต่ไม่มีข้อมูลจริง
                            example_code = sorted_unique_codes[0]; dtype = 'float64' if isinstance(example_code,(int,float)) else 'object'
                            idx = pd.Index(sorted_unique_codes, dtype=dtype); freq_series_full = pd.Series(0, index=idx, dtype='int64')
                        else:
                            # มีข้อมูลจริง, ทำ reindex
                            try: reindex_idx = pd.Index(sorted_unique_codes, dtype=observed_counts_series.index.dtype)
                            except TypeError: reindex_idx = pd.Index(sorted_unique_codes) # Fallback if dtype mismatch
                            freq_series_full = observed_counts_series.reindex(reindex_idx, fill_value=0)
                    except Exception as e_reindex:
                        add_log(f"    ⚠ Warning: Reindex failed for {item_name}. Using observed counts only. Error: {e_reindex}", "WARNING"); print(f"Warning: Reindex failed for {item_name}...")
                        freq_series_full = observed_counts_series # ใช้ข้อมูลเท่าที่นับได้

                    # สร้าง DataFrame และคำนวณ %
                    if freq_series_full.empty:
                        freq_table = pd.DataFrame(columns=['Code','ข้อ','Base N','Base %']); total_n = 0
                    else:
                        freq_table = freq_series_full.reset_index(); freq_table.columns = ['Code', 'Base N']
                        # ฟังก์ชันหา Label (เหมือนเดิม)
                        def get_label(code):
                            label = None; value_dict = value_labels_dict
                            if pd.notna(code):
                                code_str = str(code); label = value_dict.get(code)
                                if label is None: label = value_dict.get(code_str)
                                if label is None:
                                    try:
                                        code_as_float = float(code)
                                        if code_as_float.is_integer(): label = value_dict.get(int(code_as_float))
                                        if label is None: label = value_dict.get(code_as_float)
                                    except (ValueError, TypeError): pass
                                return str(label) if label is not None else code_str
                            else: return "(Missing)"

                        freq_table.insert(1, 'ข้อ', freq_table['Code'].apply(get_label))
                        total_n = freq_table['Base N'].sum()
                        freq_table['Base %'] = (freq_table['Base N'] / total_n * 100) if total_n > 0 else 0.0

                    # เพิ่มแถว Total
                    total_row = pd.DataFrame({'Code': ['Base Total'], 'ข้อ': [''], 'Base N': [total_n], 'Base %': [100.0 if total_n > 0 else 0.0]})
                    freq_table = total_row if freq_table.empty else pd.concat([total_row, freq_table], ignore_index=True)

                    # เก็บผลลัพธ์
                    frequency_results[item_name] = (variable_label, freq_table)
                    add_log(f"    ✓ คำนวณคอลัมน์สำเร็จ: {item_name} (Base N: {total_n})", "SUCCESS")

                except Exception as e:
                    # จัดการ Error ที่อาจเกิดระหว่างคำนวณ Single Column นี้
                    error_message = f"Error calculating Single Column '{item_name}': {e}"
                    add_log(f"    ✗ ข้อผิดพลาด Single Col: {error_message}", "ERROR"); print(error_message)
                    errors_freq[item_name] = str(e)
                    # import traceback # Optional
                    # print(traceback.format_exc()) # Optional

            else:
                # กรณีที่ item_name ที่เลือกมา ไม่พบใน ma_sets หรือ single_cols (ไม่ควรเกิดถ้า select_variables_for_frequency ทำงานถูกต้อง)
                warning_msg = f"Item '{item_name}' from selection list was not found as a valid MA Set key or Single Column name. Skipping calculation."
                add_log(f"    ⚠ {warning_msg}", "WARNING"); print(warning_msg)
                errors_freq[item_name] = "Item not found in identified MA sets or single columns"

        # --- สิ้นสุดลูปคำนวณ for item_name in selected_columns ---

    finally:
        # --- ไม่ว่าการคำนวณจะสำเร็จหรือล้มเหลว ต้อง Re-enable buttons ---
        if progress_status_var: progress_status_var.set("Calculation Complete. Preparing Export...") # Or set based on errors
        try:
            for btn in buttons_to_disable:
                 set_button_enabled(btn, True)
        except Exception as e_enable:
            add_log(f"Warning: Error re-enabling buttons: {e_enable}", "WARNING")
            print(f"Warning: Error re-enabling buttons: {e_enable}")
        process_events()

    # --- 4. สรุปผลการคำนวณ และ แจ้งเตือน Error (ถ้ามี) ---
    add_log("\n===== สรุปผลการคำนวณ =====")
    add_log(f"• จำนวนรายการที่เลือก: {len(selected_columns)}")
    add_log(f"• จำนวนตารางที่สร้างสำเร็จ: {len(frequency_results)}")
    add_log(f"• จำนวนรายการที่เกิดข้อผิดพลาดในการคำนวณ: {len(errors_freq)}")

    if not frequency_results:
        final_error_message = "ไม่สามารถสร้างตาราง Frequency ได้สำหรับรายการใดๆ ที่เลือก"
        if errors_freq:
            final_error_message += " เนื่องจากเกิดข้อผิดพลาดในการคำนวณ"
            error_list_short = "\n".join([f"- {col}: {msg}" for col, msg in list(errors_freq.items())[:10]])
            if len(errors_freq) > 10: error_list_short += "\n..."
            final_error_message += f"\n\nข้อผิดพลาดตัวอย่าง:\n{error_list_short}"
        add_log(f"❌ {final_error_message}", "ERROR"); ui.show_error("คำนวณล้มเหลวทั้งหมด", final_error_message)
        # Reset progress bar status
        if progress_status_var: progress_status_var.set("Idle (Failed)")
        return # จบการทำงานถ้าไม่มีผลลัพธ์เลย

    if errors_freq:
          error_list_short = "\n".join([f"- {col}: {msg}" for col, msg in list(errors_freq.items())[:10]])
          if len(errors_freq) > 10: error_list_short += "\n..."
          ui.show_warning("Frequency Calculation Errors", f"เกิดข้อผิดพลาดระหว่างการคำนวณ Frequency สำหรับบางรายการ:\n{error_list_short}\n\n(ดูรายละเอียดเพิ่มเติมใน Log)\n\nจะส่งออกเฉพาะตารางที่สร้างสำเร็จเท่านั้น")


    # --- 5. ถามผู้ใช้เรื่องรูปแบบ Export ---
    export_mode_single_sheet = ui.ask_yes_no(
        "เลือกรูปแบบการ Export",
        "ต้องการ Export ตาราง Frequency ทั้งหมดลงในชีทเดียว ('Table') หรือไม่?\n\n"
        "• กด 'Yes' -> รวมทุกตารางในชีท 'Table'\n"
        "• กด 'No'  -> แยกแต่ละตารางเป็นชีทของตัวเอง",
        icon='question'
    )
    add_log(f"\n===== เตรียมการ Export ไปยัง Excel =====")
    add_log(f"โหมดที่เลือก: {'ชีทเดียว (Single Sheet)' if export_mode_single_sheet else 'หลายชีท (Multiple Sheets)'}")

    # --- 6. เตรียมข้อมูลสำหรับ Index Sheet (ให้ครบทุกรายการที่เลือก) ---
    add_log("กำลังเตรียมข้อมูลสำหรับ Index Sheet...")
    index_sheet_info = []
    generated_safe_sheet_names = {} # เก็บชื่อชีทที่จะใช้จริง: {item_key: safe_name}
    SINGLE_SHEET_NAME = "Table" # ชื่อชีทตายตัวสำหรับโหมดชีทเดียว
    ROW_SPACING = 2  # ระยะห่างระหว่างตารางใน single sheet

    keys_to_export_list = [key for key in frequency_results if key in selected_columns]
    total_items_to_export = len(keys_to_export_list)
    add_log(f"จำนวนตารางที่จะ Export (ตามที่เลือกและสำเร็จ): {total_items_to_export}")

    # Pre-generate safe sheet names for items that WILL be exported (for multi-sheet mode link consistency)
    if not export_mode_single_sheet:
        temp_sheet_name_map = {}
        for item_key in keys_to_export_list:
            safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', item_key)[:31] # Limit sheet name length
            name_count = 1; original_safe_sheet_name = safe_sheet_name
            while safe_sheet_name in temp_sheet_name_map:
                suffix = f"_{name_count}"; max_base_len = 31 - len(suffix)
                safe_sheet_name = f"{original_safe_sheet_name[:max_base_len]}{suffix}"; name_count += 1
            temp_sheet_name_map[safe_sheet_name] = True
            generated_safe_sheet_names[item_key] = safe_sheet_name

    # สร้างข้อมูล Index โดยวนตามรายการที่ผู้ใช้เลือกทั้งหมด
    for idx, item_name in enumerate(selected_columns):
        item_id = idx + 1
        variable_name = item_name
        sheet_name = "N/A"
        base_n_count = np.nan
        link_target_url = None
        status = "Skipped" # Default status

        if item_name in frequency_results: # คำนวณสำเร็จและอยู่ในรายการที่จะ export
            status = "Exported"
            _, freq_table = frequency_results[item_name]
            # ดึง Base N count จากแถวแรก (Base Total)
            if not freq_table.empty and freq_table.iloc[0]['Code'] == 'Base Total':
                base_n_val = freq_table.iloc[0]['Base N']
                if pd.notna(base_n_val):
                    try: base_n_count = int(base_n_val)
                    except (ValueError, TypeError): base_n_count = base_n_val # Keep as is if cannot convert
                else: base_n_count = 0 # If Base N is NaN in total row
            else: base_n_count = 0 # If table is empty or no total row

            # กำหนด Sheet Name และ Link Target
            if export_mode_single_sheet:
                sheet_name = SINGLE_SHEET_NAME
                # Link ไปยังตำแหน่งคร่าวๆ (อาจจะไม่แม่นยำถ้ามี Error ก่อนหน้าเยอะ)
                # คำนวณ start row คร่าวๆ จากลำดับที่ export สำเร็จ
                export_index = keys_to_export_list.index(item_name) if item_name in keys_to_export_list else 0
                # การคำนวณ start row ที่แม่นยำต้องทำตอนเขียนจริง, อันนี้เป็นแค่การประมาณคร่าวๆ
                approx_start_row = export_index * 15 # ประมาณว่าแต่ละตารางใช้ 15 แถว
                link_target_url = f"internal:'{SINGLE_SHEET_NAME}'!A{approx_start_row + 1}"
            else:
                sheet_name = generated_safe_sheet_names.get(item_name, "Error_Name")
                if sheet_name != "Error_Name":
                    link_target_url = f"internal:'{sheet_name}'!A1"
                else:
                    status = "Export Error (Sheet Name)"
                    add_log(f"Error generating safe sheet name for: {item_name}", "ERROR")

        elif item_name in errors_freq: # คำนวณไม่สำเร็จ
            status = f"Error: {errors_freq[item_name][:60]}" # แสดง Error สั้นๆ
            sheet_name = "N/A"; base_n_count = np.nan; link_target_url = None
        else: # ไม่ได้ถูกคำนวณ (ไม่ควรเกิด แต่ใส่ไว้เผื่อ)
            status = "Not Found/Processed"
            sheet_name = "N/A"; base_n_count = np.nan; link_target_url = None

        index_sheet_info.append({
            'id': item_id, 'variable': variable_name, 'sheet_name': sheet_name,
            'count': base_n_count, 'link_target': link_target_url, 'status': status
        })
    add_log("เตรียมข้อมูล Index Sheet เสร็จสิ้น")

    # --- 7. เตรียม Export ไปยัง Excel ---
    base_filename = 'SPSS_Data'
    if file_var and file_var.get() and os.path.exists(file_var.get()): base_filename = os.path.splitext(os.path.basename(file_var.get()))[0]
    timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M'); suggested_filename = f"{base_filename}_Frequencies_{timestamp}.xlsx"
    freq_save_path = ui.save_file_dialog(defaultextension='.xlsx', filetypes=[("Excel files", "*.xlsx")], initialfile=suggested_filename, title="บันทึก Frequency Tables เป็น Excel")

    if not freq_save_path:
        add_log("❌ ยกเลิกการ Export โดยผู้ใช้", "WARNING")
        if progress_status_var: progress_status_var.set("Idle (Export Cancelled)")
        return # ไม่ต้อง re-enable buttons เพราะทำใน finally แล้ว

    add_log(f"ตำแหน่งไฟล์ที่เลือก: {freq_save_path}")
    if progress_status_var: progress_status_var.set("Exporting to Excel...")
    process_events()


    # --- 8. เขียนไฟล์ Excel ---
    writer = None # Initialize writer to None
    try:
        with pd.ExcelWriter(freq_save_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            add_log(f"กำลังเขียนไฟล์ Excel...")

            # --- Define Excel Formats ---
            title_format = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'left', 'valign': 'vcenter'})
            subtitle_format = workbook.add_format({'bold': True, 'align': 'left', 'valign': 'top', 'text_wrap': True})
            table_header_format = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            total_row_format = workbook.add_format({'bg_color': '#F0F0F0', 'bold': True, 'border': 1})
            data_format_text = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'text_wrap': True})
            data_format_num_table = workbook.add_format({'border': 1, 'num_format': '#,##0'})
            code_format_num_right = workbook.add_format({'border': 1, 'num_format': '0', 'align': 'right'})
            code_format_float_right = workbook.add_format({'border': 1, 'num_format': '0.0#######', 'align': 'right'})
            code_format_str_left = workbook.add_format({'border': 1, 'align': 'left'})
            data_format_pct_int = workbook.add_format({'border': 1, 'num_format': '0%'}) # Display % with no decimals
            index_header_format = workbook.add_format({'bold': True, 'bottom': 1, 'align': 'center', 'valign': 'vcenter','bg_color': '#D9D9D9'})
            url_format = workbook.add_format({'font_color': 'blue', 'underline': 1, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            total_row_num_format = workbook.add_format({'bg_color': '#F0F0F0', 'bold': True, 'border': 1, 'num_format': '#,##0'})
            total_row_pct_format = workbook.add_format({'bg_color': '#F0F0F0', 'bold': True, 'border': 1, 'num_format': '0%'})
            index_text_format = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
            index_id_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
            index_num_format = workbook.add_format({'border': 1, 'num_format': '#,##0', 'align': 'right', 'valign': 'vcenter'})
            index_status_exported_format = workbook.add_format({'font_color': 'green', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
            index_status_error_format = workbook.add_format({'font_color': 'red', 'border': 1, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
            index_status_other_format = workbook.add_format({'font_color': '#808080', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
            index_blank_format = workbook.add_format({'border': 1})


            # --- เขียน Index Sheet ก่อน ---
            add_log("กำลังเขียน Index Sheet...")
            index_sheet = workbook.add_worksheet('Index')
            # Headers
            index_sheet.write('A1', 'ID', index_header_format)
            index_sheet.write('B1', 'Variable / MA Set', index_header_format)
            index_sheet.write('C1', 'Sheet Name', index_header_format)
            index_sheet.write('D1', 'Base N Count', index_header_format)
            index_sheet.write('E1', 'Status', index_header_format)
            index_sheet.write('F1', 'Link to Table', index_header_format)
            max_var_len = 15; max_sheet_len = 12; max_status_len = 10

            # Write Index Data
            if index_sheet_info:
                add_log(f"เขียนข้อมูล Index จำนวน {len(index_sheet_info)} รายการ...")
                for info in index_sheet_info:
                    row_num = info['id'] # ID เริ่มจาก 1 -> แถว Excel ที่ 2
                    index_sheet.write_number(row_num, 0, info['id'], index_id_format)
                    index_sheet.write_string(row_num, 1, str(info['variable']), index_text_format)
                    index_sheet.write_string(row_num, 2, str(info['sheet_name']), index_text_format)
                    if pd.notna(info['count']) and isinstance(info['count'], (int, float)):
                        index_sheet.write_number(row_num, 3, info['count'], index_num_format)
                    else: # Handle NaN, Error strings etc.
                        index_sheet.write_string(row_num, 3, str(info['count']), index_text_format)

                    status_text = str(info['status'])
                    if status_text == "Exported": fmt_status = index_status_exported_format
                    elif status_text.startswith("Error"): fmt_status = index_status_error_format
                    else: fmt_status = index_status_other_format
                    index_sheet.write_string(row_num, 4, status_text, fmt_status)

                    if info['link_target']:
                        index_sheet.write_url(row_num, 5, info['link_target'], url_format, string='Go to Table')
                    else:
                        index_sheet.write_blank(row_num, 5, None, index_blank_format)

                    # Track max lengths for column width adjustment
                    if len(str(info['variable'])) > max_var_len: max_var_len = len(str(info['variable']))
                    if len(str(info['sheet_name'])) > max_sheet_len: max_sheet_len = len(str(info['sheet_name']))
                    if len(status_text) > max_status_len: max_status_len = len(status_text)

                # Set Index Column Widths
                index_sheet.set_column('A:A', 5) # ID
                index_sheet.set_column('B:B', min(max_var_len + 2, 50)) # Variable Name
                index_sheet.set_column('C:C', min(max_sheet_len + 2, 35)) # Sheet Name
                index_sheet.set_column('D:D', 15) # Base N
                index_sheet.set_column('E:E', min(max_status_len + 2, 40)) # Status
                index_sheet.set_column('F:F', 15) # Link
                index_sheet.freeze_panes(1, 1) # Freeze top row and first column
                index_sheet.activate() # Make Index the active sheet on open
                add_log("    ✓ เขียน Index Sheet สำเร็จ", "SUCCESS")
            else:
                 index_sheet.write('A1', 'No items were selected or processed.')
                 add_log("    ⚠ ไม่มีข้อมูลสำหรับเขียนลง Index Sheet", "WARNING")


            # --- เขียน Data Sheets/Table ---
            export_progress_count = 0
            items_actually_exported = 0
            processed_for_export = set()
            table_sheet = None
            current_row_on_table_sheet = 0
            max_col_widths = {} # Track max widths for single sheet mode
            single_sheet_links = {} # Store start row for links in single sheet mode {item_key: start_row}


            add_log("\n--- เริ่มการเขียนตาราง Frequency ลงชีท ---")
            add_log(f"จำนวนตารางที่จะ Export (ตามที่เลือกและสำเร็จ): {total_items_to_export}")

            if export_mode_single_sheet:
                table_sheet = workbook.add_worksheet(SINGLE_SHEET_NAME)

            if not keys_to_export_list:
                add_log("⚠ ไม่มีตารางที่คำนวณสำเร็จและถูกเลือก ที่จะเขียนลงชีทข้อมูล", "WARNING")
            else:
                # วนลูปตามลำดับคอลัมน์เดิม เพื่อ Export ตาม Order ที่ต้องการ
                for col_name_in_order in all_columns:
                    item_key_to_consider = None; is_ma_in_order = False
                    cols_to_mark_done = [col_name_in_order]

                    # Check if this column is the start of an MA set or a single column
                    match_in_order = ma_pattern.match(col_name_in_order)
                    if match_in_order:
                        base_name_in_order = match_in_order.group(1).upper() + "_O"
                        if base_name_in_order in ma_sets:
                            is_ma_in_order = True; item_key_to_consider = base_name_in_order
                            cols_to_mark_done = ma_sets.get(base_name_in_order, [col_name_in_order])
                    # If not MA set start, check if it's a single column
                    elif col_name_in_order in single_cols:
                        item_key_to_consider = col_name_in_order

                    # --- ตรวจสอบว่าต้อง Export รายการนี้หรือไม่ ---
                    if (item_key_to_consider is not None and
                        item_key_to_consider in keys_to_export_list and # Must be successfully calculated
                        item_key_to_consider not in processed_for_export): # And not already processed

                        export_progress_count += 1
                        if progress_status_var: progress_status_var.set(f"Writing Export ({export_progress_count}/{total_items_to_export}): {item_key_to_consider[:40]}...")
                        process_events()

                        items_actually_exported += 1
                        (variable_label, freq_table) = frequency_results[item_key_to_consider]
                        worksheet_to_write = None; start_row_offset = 0
                        current_export_log = f"[{export_progress_count}/{total_items_to_export}] กำลังเขียนตาราง: '{item_key_to_consider}'"

                        # กำหนด Worksheet ปลายทาง และ Start Row
                        if export_mode_single_sheet:
                            worksheet_to_write = table_sheet
                            start_row_offset = current_row_on_table_sheet
                            single_sheet_links[item_key_to_consider] = start_row_offset # Store start row for index link
                            add_log(f"{current_export_log} ไปยังชีท '{SINGLE_SHEET_NAME}' แถว {start_row_offset + 1}")
                        else:
                            safe_sheet_name = generated_safe_sheet_names.get(item_key_to_consider, f"Err_{item_key_to_consider}"[:31])
                            worksheet_to_write = workbook.add_worksheet(safe_sheet_name)
                            start_row_offset = 0
                            add_log(f"{current_export_log} ไปยังชีทใหม่ '{safe_sheet_name}'")

                        # --- เขียน Title, Subtitle, Headers ---
                        title_text = f"ข้อมูลความถี่สำหรับ: {item_key_to_consider}"
                        subtitle_text = f"โจทย์: {str(variable_label)}"
                        num_data_cols = freq_table.shape[1] # Should be 4 (Code, ข้อ, Base N, Base %)
                        merge_cols_idx = max(0, num_data_cols - 1)
                        worksheet_to_write.merge_range(start_row_offset + 0, 0, start_row_offset + 0, merge_cols_idx, title_text, title_format)
                        worksheet_to_write.merge_range(start_row_offset + 1, 0, start_row_offset + 1, merge_cols_idx, subtitle_text, subtitle_format)
                        worksheet_to_write.set_row(start_row_offset + 1, None) # Adjust height for subtitle

                        # Write Headers
                        start_row_table_headers = start_row_offset + 3
                        start_row_data = start_row_table_headers + 1
                        for c_idx, col_header in enumerate(freq_table.columns):
                            if col_header == 'Base %_Excel': continue # Skip internal column
                            worksheet_to_write.write(start_row_table_headers, c_idx, col_header, table_header_format)


                        # --- เขียน Data Rows with Formatting ---
                        # Prepare data (convert N/%) for writing
                        freq_table_export = freq_table.copy()
                        freq_table_export['Base N'] = pd.to_numeric(freq_table_export['Base N'], errors='coerce')
                        # Use the original '%' column for formatting logic, but write numeric value / 100
                        freq_table_export['Base %_Numeric'] = pd.to_numeric(freq_table_export['Base %'], errors='coerce') / 100.0

                        rows_written_count = 0
                        for r_idx, row_data in freq_table_export.iterrows():
                            excel_row_idx = start_row_data + r_idx
                            # Check if this is the 'Base Total' row using original freq_table
                            is_total_row = (freq_table.iloc[r_idx]['Code'] == 'Base Total')

                            # Column 0: Code
                            code_val = freq_table.iloc[r_idx]['Code'] # Use original for comparison/display logic
                            fmt_code = code_format_str_left # Default
                            if is_total_row:
                                code_display = "Base Total"; fmt_code = total_row_format
                                worksheet_to_write.write_string(excel_row_idx, 0, code_display, fmt_code)
                            else:
                                # Try converting to number for display, otherwise write as string
                                try:
                                    num_val = float(code_val)
                                    if num_val.is_integer(): worksheet_to_write.write_number(excel_row_idx, 0, int(num_val), code_format_num_right)
                                    else: worksheet_to_write.write_number(excel_row_idx, 0, num_val, code_format_float_right)
                                except (ValueError, TypeError): worksheet_to_write.write_string(excel_row_idx, 0, str(code_val), code_format_str_left)

                            # Column 1: Label ('ข้อ')
                            label_text = str(row_data['ข้อ'])
                            fmt_label = total_row_format if is_total_row else data_format_text
                            worksheet_to_write.write_string(excel_row_idx, 1, label_text, fmt_label)

                            # Column 2: Base N
                            base_n_val = row_data['Base N'] # Use numeric coerced value
                            fmt_n = total_row_num_format if is_total_row else data_format_num_table
                            if pd.notna(base_n_val): worksheet_to_write.write_number(excel_row_idx, 2, base_n_val, fmt_n)
                            else: worksheet_to_write.write_blank(excel_row_idx, 2, None, fmt_n)

                            # Column 3: Base %
                            base_pct_val = row_data['Base %_Numeric'] # Use numeric value / 100
                            fmt_pct = total_row_pct_format if is_total_row else data_format_pct_int
                            if pd.notna(base_pct_val): worksheet_to_write.write_number(excel_row_idx, 3, base_pct_val, fmt_pct) # Write as fraction for % format
                            else: worksheet_to_write.write_blank(excel_row_idx, 3, None, fmt_pct)

                            rows_written_count += 1


                        # --- Auto-adjust column widths ---
                        col_widths_this_table = {}
                        # Calculate widths based on original data representation for accuracy
                        for idx, col in enumerate(freq_table.columns):
                            if col == 'Base %_Excel' or col == 'Base %_Numeric': continue # Skip internal/temp cols

                            header_width = len(str(col)); max_len_data = 10 # Min width
                            try:
                                col_data_for_width = freq_table[col]
                                if col == 'Code': # Use the display logic for width calc
                                    display_codes = []
                                    for r_idx_w in range(len(freq_table)):
                                        code_val_w = freq_table.iloc[r_idx_w]['Code']
                                        if code_val_w == 'Base Total': display_codes.append("Base Total")
                                        else:
                                             try:
                                                 num_val_w = float(code_val_w)
                                                 if num_val_w.is_integer(): display_codes.append(str(int(num_val_w)))
                                                 else: display_codes.append(str(num_val_w))
                                             except (ValueError, TypeError): display_codes.append(str(code_val_w))
                                    col_data_for_width = pd.Series(display_codes)
                                elif col == 'Base %': # Use formatted string for width calc
                                    col_data_for_width = freq_table_export['Base %_Numeric'].apply(lambda x: f"{x*100:.0f}%" if pd.notna(x) else "")
                                elif col == 'Base N': # Use formatted string
                                     col_data_for_width = freq_table_export['Base N'].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "")

                                lengths = pd.concat([col_data_for_width.astype(str).map(len), pd.Series([header_width])])
                                max_len_data_calc = lengths.max(skipna=True)
                                if pd.notna(max_len_data_calc): max_len_data = int(max_len_data_calc)
                            except Exception as e_width: print(f"Width calc error Col '{col}' Table '{item_key_to_consider}': {e_width}")

                            current_col_width = max(max_len_data, header_width) + 2 # Add padding
                            # Get the correct column index in Excel (0-based)
                            excel_col_idx = idx
                            col_widths_this_table[excel_col_idx] = current_col_width

                            # Track max width for single sheet mode
                            if export_mode_single_sheet:
                                max_col_widths[excel_col_idx] = max(max_col_widths.get(excel_col_idx, 0), current_col_width)
                            # Set width immediately for multi-sheet mode
                            else:
                                final_width = current_col_width
                                # Apply constraints for specific columns
                                if excel_col_idx == 1: final_width = min(max(final_width, 25), 60) # ข้อ (Label)
                                elif excel_col_idx == 0: final_width = min(max(final_width, 8), 20)  # Code
                                elif excel_col_idx == 3: final_width = min(max(final_width, 10), 15) # Base %
                                else: final_width = min(final_width, 40) # Base N
                                worksheet_to_write.set_column(excel_col_idx, excel_col_idx, final_width)


                        # --- Add Back to Index link & Freeze Panes (Multi-sheet only) ---
                        if not export_mode_single_sheet:
                            try:
                                # Add link in the column after the last data column
                                link_col_idx = num_data_cols
                                worksheet_to_write.write_url(start_row_table_headers, link_col_idx, "internal:'Index'!A1", url_format, string='<- Back to Index')
                                worksheet_to_write.set_column(link_col_idx, link_col_idx, 15) # Width for link column
                            except Exception as e_link: print(f"Link write error sheet {safe_sheet_name}: {e_link}")
                            # Freeze panes below headers and right of the first column (Code)
                            worksheet_to_write.freeze_panes(start_row_data, 1)

                        # --- Update row counter for Single sheet mode ---
                        if export_mode_single_sheet:
                            # Title(1) + Subtitle(1) + Blank(1) + Header(1) + Data Rows + Blank Rows Below
                            rows_used_by_table = 1 + 1 + 1 + 1 + rows_written_count
                            current_row_on_table_sheet += rows_used_by_table + ROW_SPACING

                        # --- Mark this item and its MA columns as processed ---
                        processed_for_export.add(item_key_to_consider)
                        if is_ma_in_order:
                            processed_for_export.update(cols_to_mark_done) # Avoid processing sub-columns individually

            # --- ตั้งค่า Column Widths และ Freeze Panes (Single Sheet Mode - ทำตอนท้าย) ---
            if export_mode_single_sheet and table_sheet and items_actually_exported > 0:
                add_log("กำลังปรับความกว้างคอลัมน์และตั้งค่า Freeze Panes สำหรับชีทเดียว...")
                for col_idx, width in max_col_widths.items():
                    final_width = width
                    # Apply constraints for specific columns (adjust indices if table structure changes)
                    if col_idx == 1: final_width = min(max(width, 25), 60) # ข้อ (Label)
                    elif col_idx == 0: final_width = min(max(width, 8), 20)  # Code
                    elif col_idx == 3: final_width = min(max(width, 10), 15) # Base %
                    else: final_width = min(width, 40) # Base N
                    table_sheet.set_column(col_idx, col_idx, final_width)

                # Freeze panes below first table's header and right of first column
                first_table_header_row = 3 # Header row of the very first table
                first_table_data_row = first_table_header_row + 1
                freeze_col_index = 1 # Freeze after 'Code' column
                table_sheet.freeze_panes(first_table_data_row, freeze_col_index)
                add_log(f"    ✓ ตั้งค่า Freeze Panes ที่คอลัมน์ {freeze_col_index+1} แถว {first_table_data_row+1}")
            elif export_mode_single_sheet and items_actually_exported == 0:
                add_log("    ⚠ ไม่มีการตั้งค่า Freeze Panes เนื่องจากไม่มีตารางถูก Export ลงชีท Table", "WARNING")


            # --- แก้ไข Link ใน Index Sheet สำหรับ Single Sheet Mode ---
            if export_mode_single_sheet and index_sheet_info and items_actually_exported > 0:
                 add_log("กำลังอัปเดต Link ใน Index Sheet สำหรับโหมดชีทเดียว...")
                 for info in index_sheet_info:
                     # หา start row ที่เก็บไว้ตอนเขียนตาราง
                     actual_start_row = single_sheet_links.get(info['variable'])
                     if actual_start_row is not None and info['status'] == "Exported":
                         # สร้าง link ไปยังเซลล์แรกของ Title ของตารางนั้นๆ
                         link_target_url = f"internal:'{SINGLE_SHEET_NAME}'!A{actual_start_row + 1}"
                         row_num = info['id'] # ID เริ่มจาก 1 -> แถว Excel ที่ 2
                         index_sheet.write_url(row_num, 5, link_target_url, url_format, string='Go to Table')
                 add_log("    ✓ อัปเดต Link ใน Index Sheet สำเร็จ")


        # --- สิ้นสุด `with pd.ExcelWriter(...)` ไฟล์จะถูกบันทึก ---
        add_log("\n===== สรุปการทำงาน =====")
        add_log(f"• การส่งออกเสร็จสมบูรณ์: {os.path.basename(freq_save_path)}", "SUCCESS")
        add_log(f"• จำนวนตารางที่ส่งออก (ตามที่เลือกและสำเร็จ): {items_actually_exported}")
        if writer: # Check if writer was successfully created
             total_sheets_created = len(writer.sheets)
             add_log(f"• ชีทที่สร้างทั้งหมด: {total_sheets_created} ชีท {'(รวม Index และ Table)' if export_mode_single_sheet else '(รวม Index และ Data Sheets)'}")
        add_log(f"• เวลาสิ้นสุด: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
        ui.show_info("สำเร็จ", f"สร้างไฟล์ Frequency Tables:\n{os.path.basename(freq_save_path)}\nเรียบร้อยแล้ว ({items_actually_exported} ตารางตามที่เลือก)")

        # --- ถามผู้ใช้ว่าต้องการเปิดไฟล์หรือไม่ ---
        if ui.ask_yes_no("เปิดไฟล์", "ต้องการเปิดไฟล์ Frequency Tables ที่สร้างเสร็จหรือไม่?"):
              add_log("กำลังเปิดไฟล์ Excel...")
              try:
                  if sys.platform.startswith('win'): os.startfile(freq_save_path); add_log("  ✓ เปิดไฟล์ (Windows)")
                  elif sys.platform.startswith('darwin'): subprocess.run(["open", freq_save_path], check=True); add_log("  ✓ เปิดไฟล์ (macOS)")
                  else: subprocess.run(["xdg-open", freq_save_path], check=True); add_log("  ✓ เปิดไฟล์ (Linux)")
              except FileNotFoundError: error_msg = f"ไม่พบคำสั่งสำหรับเปิดไฟล์..."; add_log(f"  ❌ {error_msg}", "ERROR"); ui.show_error("เปิดไฟล์ไม่ได้", error_msg)
              except Exception as e: error_msg = f"เกิดข้อผิดพลาดในการเปิดไฟล์: {e}"; add_log(f"  ❌ {error_msg}", "ERROR"); ui.show_error("เปิดไฟล์ไม่ได้", f"{error_msg}\n\nกรุณาเปิดไฟล์ด้วยตนเองที่:\n{freq_save_path}")

    except PermissionError as pe:
        error_msg = f"ไม่สามารถเขียนไฟล์ได้ อาจเป็นเพราะไฟล์กำลังเปิดอยู่ หรือไม่มีสิทธิ์เขียนในตำแหน่งที่เลือก"
        add_log(f"\n❌ เกิดข้อผิดพลาด PermissionError: {error_msg}", "ERROR"); add_log(f"รายละเอียด: {pe}")
        ui.show_error("Export ล้มเหลว", f"{error_msg}\n\nกรุณาปิดไฟล์ (ถ้าเปิดอยู่) หรือเลือกตำแหน่งบันทึกอื่น")
    except Exception as e:
        # Catch-all for other unexpected errors during Excel writing
        error_msg = f"เกิดข้อผิดพลาดที่ไม่คาดคิดระหว่างการเขียนไฟล์ Excel"
        add_log(f"\n❌ เกิดข้อผิดพลาดร้ายแรงในการ Export: {error_msg}", "ERROR"); add_log(f"รายละเอียด: {e}")
        add_log(traceback.format_exc(), "ERROR") # Log full traceback for debugging
        ui.show_error("Export ล้มเหลว", f"{error_msg}: {e}")
    except BaseException as be:
        error_msg = "พบข้อผิดพลาดระดับระบบระหว่าง Export (BaseException)"
        add_log(f"\n❌ {error_msg}", "ERROR")
        add_log(f"รายละเอียด: {be}", "ERROR")
        add_log(traceback.format_exc(), "ERROR")
        ui.show_error("Export ล้มเหลว", f"{error_msg}: {be}")
    finally:
        # --- Reset Progress Bar and Status ---
        # (ทำไปแล้วใน finally block ของการคำนวณ แต่ทำซ้ำเผื่อกรณี error ตอน export)
        if ui and ui.progress_bar:
            ui.progress_bar.setValue(0)
        if progress_status_var: progress_status_var.set("Idle")
        process_events()

# --- สิ้นสุดฟังก์ชัน run_all_frequencies (เวอร์ชันแก้ไข TypeError) ---





# (ฟังก์ชัน add_log เหมือนเดิมจากครั้งก่อน)
def add_log(message, level="INFO"):
    """
    เพิ่มข้อความลงใน Log พร้อมระบุระดับความสำคัญ
    level: "INFO" (ปกติ), "SUCCESS" (สำเร็จ), "WARNING" (คำเตือน), "ERROR" (ข้อผิดพลาด)
    """
    global ui
    if ui is None:
        print(f"LOG [{level}] (UI not ready): {message}")
        return

    color_map = {
        "SUCCESS": "#16a34a",
        "WARNING": "#f59e0b",
        "ERROR": "#ef4444",
        "INFO": "#374151",
    }
    color = color_map.get(level, "#374151")

    cursor = ui.log_text.textCursor()
    fmt = QtGui.QTextCharFormat()
    fmt.setForeground(QtGui.QBrush(QtGui.QColor(color)))
    cursor.movePosition(QtGui.QTextCursor.MoveOperation.End)
    cursor.insertText(f"{message}\n", fmt)
    ui.log_text.setTextCursor(cursor)
    ui.log_text.ensureCursorVisible()

# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---

class QuotaSamplerWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self._allow_close = False
        self.setWindowTitle("โปรแกรมตรวจสอบเงื่อนไข SPSS V1.4 (Modern)")
        self.resize(1100, 720)
        try:
            icon_path = resource_path("Clean.ico")
            if os.path.exists(icon_path):
                self.setWindowIcon(QtGui.QIcon(icon_path))
        except Exception:
            pass


        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        main = QtWidgets.QVBoxLayout(central)
        main.setContentsMargins(16, 16, 16, 16)
        main.setSpacing(12)

        top_row = QtWidgets.QHBoxLayout()
        top_row.addWidget(QtWidgets.QLabel("ไฟล์:"))
        self.file_path_edit = QtWidgets.QLineEdit()
        self.file_path_edit.setReadOnly(True)
        top_row.addWidget(self.file_path_edit, 1)
        self.btn_select_file = QtWidgets.QPushButton("เลือกไฟล์ SPSS")
        self.btn_help = QtWidgets.QPushButton("วิธีใช้")
        top_row.addWidget(self.btn_select_file)
        top_row.addWidget(self.btn_help)
        main.addLayout(top_row)

        cond_row = QtWidgets.QHBoxLayout()
        cond_row.addWidget(QtWidgets.QLabel("เงื่อนไข:"))
        self.condition_input = QtWidgets.QLineEdit()
        cond_row.addWidget(self.condition_input, 1)
        self.btn_save = QtWidgets.QPushButton("บันทึก")
        self.btn_delete = QtWidgets.QPushButton("ลบ")
        self.btn_load = QtWidgets.QPushButton("โหลด")
        self.btn_export = QtWidgets.QPushButton("บันทึก")
        cond_row.addWidget(self.btn_save)
        cond_row.addWidget(self.btn_delete)
        cond_row.addWidget(self.btn_load)
        cond_row.addWidget(self.btn_export)
        main.addLayout(cond_row)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Vertical)
        main.addWidget(splitter, 1)

        self.conditions_table = QtWidgets.QTableWidget()
        self.conditions_table.setColumnCount(3)
        self.conditions_table.setHorizontalHeaderLabels(["ID", "เงื่อนไขที่บันทึกไว้", "Count"])
        self.conditions_table.horizontalHeader().setStretchLastSection(True)
        self.conditions_table.verticalHeader().setVisible(False)
        self.conditions_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.conditions_table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)
        self.conditions_table.setAlternatingRowColors(True)
        header = self.conditions_table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.conditions_table.setColumnWidth(0, 60)
        self.conditions_table.setColumnWidth(2, 90)
        self.conditions_table.setMouseTracking(True)
        self.conditions_table.cellEntered.connect(self._hover_select_row)
        splitter.addWidget(self.conditions_table)

        log_container = QtWidgets.QWidget()
        log_layout = QtWidgets.QVBoxLayout(log_container)
        log_layout.setContentsMargins(0, 0, 0, 0)
        log_header = QtWidgets.QHBoxLayout()
        log_header.addWidget(QtWidgets.QLabel("Log:"))
        self.btn_clear_log = QtWidgets.QPushButton("ล้าง Log")
        log_header.addStretch(1)
        log_header.addWidget(self.btn_clear_log)
        log_layout.addLayout(log_header)
        self.log_text = QtWidgets.QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text, 1)
        splitter.addWidget(log_container)
        splitter.setSizes([420, 180])

        progress_row = QtWidgets.QHBoxLayout()
        self.progress_label = QtWidgets.QLabel("สถานะ")
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setMinimum(0)
        progress_row.addWidget(self.progress_label)
        progress_row.addWidget(self.progress_bar, 1)
        main.addLayout(progress_row)

        action_row = QtWidgets.QHBoxLayout()
        self.check_button = QtWidgets.QPushButton("ตรวจสอบเงื่อนไข & ส่งออก")
        self.freq_button = QtWidgets.QPushButton("ทำ Frequency")
        action_row.addWidget(self.check_button, 1)
        action_row.addWidget(self.freq_button, 1)
        main.addLayout(action_row)

        self.apply_styles()
        self.bind_actions()

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background: #f3f4f6; }
            QLabel { color: #111827; }
            QLineEdit {
                background: #ffffff; border: 1px solid #d1d5db; border-radius: 8px;
                padding: 6px 10px;
            }
            QTextEdit {
                background: #ffffff; border: 1px solid #e5e7eb; border-radius: 8px;
                padding: 8px;
            }
            QTableWidget {
                background: #ffffff; border: 1px solid #e5e7eb; border-radius: 8px;
                gridline-color: #e5e7eb;
            }
            QHeaderView::section {
                background: #eef2ff; padding: 6px; border: none; font-weight: 700;
            }
            QPushButton {
                background: #111827; color: white; border: none; padding: 8px 14px;
                border-radius: 8px; font-weight: 600;
            }
            QPushButton:hover { background: #0b1220; }
        """)
        self.btn_select_file.setStyleSheet("background:#10b981;color:#0f172a;")
        self.btn_help.setStyleSheet("background:#f97316;color:white;")
        self.btn_save.setStyleSheet("background:#f59e0b;color:#111827;")
        self.btn_delete.setStyleSheet("background:#ef4444;color:white;")
        self.btn_load.setStyleSheet("background:#60a5fa;color:#0b1220;")
        self.btn_export.setStyleSheet("background:#60a5fa;color:#0b1220;")
        self.check_button.setStyleSheet("background:#2563eb;color:white;font-weight:700;")
        self.freq_button.setStyleSheet("background:#16a34a;color:white;font-weight:700;")

    def bind_actions(self):
        self.btn_help.clicked.connect(self._guard_action(show_help, "แสดงวิธีใช้"))
        self.btn_select_file.clicked.connect(self._guard_action(load_file, "เลือกไฟล์ SPSS"))
        self.btn_save.clicked.connect(self._guard_action(save_condition, "บันทึกเงื่อนไข"))
        self.btn_delete.clicked.connect(self._guard_action(delete_condition, "ลบเงื่อนไข"))
        self.btn_load.clicked.connect(self._guard_action(import_conditions, "โหลดเงื่อนไข"))
        self.btn_export.clicked.connect(self._guard_action(export_conditions, "ส่งออกเงื่อนไข"))
        self.check_button.clicked.connect(self._guard_action(check_conditions, "ตรวจสอบเงื่อนไข"))
        self.freq_button.clicked.connect(self._guard_action(run_all_frequencies, "ทำ Frequency"))
        self.btn_clear_log.clicked.connect(self.log_text.clear)

    def _guard_action(self, fn, action_name):
        def wrapped():
            try:
                fn()
            except SystemExit as e:
                add_log(f"❌ Action '{action_name}' ถูกยับยั้ง SystemExit: {e}", "ERROR")
                self.show_error("โปรแกรมถูกยับยั้งไม่ให้ปิด", f"พบ SystemExit ระหว่าง '{action_name}'\nรายละเอียด: {e}")
            except Exception as e:
                add_log(f"❌ Action '{action_name}' ล้มเหลว: {e}", "ERROR")
                add_log(traceback.format_exc(), "ERROR")
                self.show_error("เกิดข้อผิดพลาด", f"เกิดข้อผิดพลาดระหว่าง '{action_name}':\n{e}")
            except BaseException as be:
                add_log(f"❌ Action '{action_name}' ล้มเหลวระดับระบบ: {be}", "ERROR")
                add_log(traceback.format_exc(), "ERROR")
                _report_fatal_error("เกิดข้อผิดพลาดร้ายแรง", f"เกิด BaseException ระหว่าง '{action_name}':\n{be}")
        return wrapped

    def closeEvent(self, event):
        _append_runtime_log("Main window closeEvent received")
        if self._allow_close:
            event.accept()
            return
        try:
            should_close = self.ask_yes_no("ยืนยันการปิดโปรแกรม", "ต้องการปิดโปรแกรมหรือไม่?")
            if should_close:
                self._allow_close = True
                _append_runtime_log("Main window closeEvent accepted by user")
                event.accept()
            else:
                _append_runtime_log("Main window closeEvent ignored by user")
                event.ignore()
        except Exception as e:
            _append_runtime_log(f"closeEvent fallback accept due to error: {e}")
            event.accept()

    def _hover_select_row(self, row, _col):
        # ไฮไลต์ทั้งแถวเมื่อเอาเมาส์ชี้ (กัน recursion จาก cellEntered -> selectRow -> cellEntered)
        if row < 0:
            return
        selection_model = self.conditions_table.selectionModel()
        if selection_model and selection_model.isRowSelected(row, QtCore.QModelIndex()):
            return
        blocker = QtCore.QSignalBlocker(self.conditions_table)
        self.conditions_table.selectRow(row)

    def _dialog(self, title, message, kind="info", buttons=("OK",)):
        dlg = SweetAlert(title, message, kind=kind, buttons=buttons, parent=self)
        dlg.exec()
        return dlg.result

    def show_info(self, title, message):
        self._dialog(title, message, kind="info")

    def show_warning(self, title, message):
        self._dialog(title, message, kind="warning")

    def show_error(self, title, message):
        self._dialog(title, message, kind="error")

    def ask_yes_no(self, title, message, **_kwargs):
        # _kwargs รองรับพารามิเตอร์เดิม (เช่น icon=) ที่ส่งมาจาก Tkinter
        result = self._dialog(title, message, kind="question", buttons=("Yes", "No"))
        return result == "Yes"

    def ask_yes_no_cancel(self, title, message):
        result = self._dialog(title, message, kind="question", buttons=("Yes", "No", "Cancel"))
        if result == "Yes":
            return True
        if result == "No":
            return False
        return None

    def _filetypes_to_filter(self, filetypes):
        if not filetypes:
            return "All files (*.*)"
        parts = []
        for label, pattern in filetypes:
            parts.append(f"{label} ({pattern})")
        return ";;".join(parts)

    def open_file_dialog(self, title="Open File", filetypes=None, filter_text=None):
        file_filter = filter_text or self._filetypes_to_filter(filetypes)
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, title, "", file_filter)
        return path

    def save_file_dialog(self, defaultextension=None, filetypes=None, title="Save File", initialfile=None, filter_text=None):
        file_filter = filter_text or self._filetypes_to_filter(filetypes)
        default_name = initialfile or ""
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, title, default_name, file_filter)
        if path and defaultextension and not path.lower().endswith(defaultextension):
            path = f"{path}{defaultextension}"
        return path

    def get_selected_conditions(self):
        rows = sorted({i.row() for i in self.conditions_table.selectedItems()})
        conds = []
        for r in rows:
            item = self.conditions_table.item(r, 1)
            if item:
                conds.append(item.text())
        return conds


def run_this_app(working_dir=None):
    global ui, file_var, condition_var, progress_status_var, check_button_widget, freq_button_widget

    def _global_excepthook(exc_type, exc_value, exc_tb):
        try:
            err_text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
            add_log("❌ Unhandled exception in application event loop", "ERROR")
            add_log(err_text, "ERROR")
            _append_runtime_log(f"Unhandled exception: {exc_value}")
            _append_runtime_log(err_text)
            _report_fatal_error("เกิดข้อผิดพลาดที่ไม่ได้ถูกดักจับ", str(exc_value))
        except Exception:
            traceback.print_exception(exc_type, exc_value, exc_tb)

    sys.excepthook = _global_excepthook

    app = QtWidgets.QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)
    app.setFont(QtGui.QFont("Tahoma", 10))

    # Capture low-level crashes (segfault/access violation) to runtime_debug.log
    try:
        fault_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "runtime_debug.log")
        fault_log_fp = open(fault_log_path, "a", encoding="utf-8")
        faulthandler.enable(file=fault_log_fp, all_threads=True)
        _append_runtime_log("faulthandler enabled")
    except Exception as e:
        _append_runtime_log(f"faulthandler enable failed: {e}")

    ui = QuotaSamplerWindow()
    file_var = UiVar(getter=ui.file_path_edit.text, setter=ui.file_path_edit.setText)
    condition_var = UiVar(getter=ui.condition_input.text, setter=ui.condition_input.setText)
    progress_status_var = UiVar(getter=ui.progress_label.text, setter=ui.progress_label.setText)
    check_button_widget = ui.check_button
    freq_button_widget = ui.freq_button

    app.aboutToQuit.connect(lambda: _append_runtime_log("QApplication.aboutToQuit fired"))
    ui.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    run_this_app()
