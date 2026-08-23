import sys
import os
import re
import uuid  # เพิ่มบรรทัดนี้
from xml.sax.saxutils import escape
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QPushButton, QLineEdit, QTextEdit, QFileDialog,
                             QMessageBox, QListWidget, QTableWidget, QTableWidgetItem,
                             QHeaderView, QGroupBox, QDialog, QSplitter, QFrame, QTreeWidget, QTreeWidgetItem,
                             QAbstractItemView, QInputDialog)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QIcon, QColor, QPalette

# Libraries check
try:
    import pyreadstat
    HAS_PYREADSTAT = True
except ImportError:
    HAS_PYREADSTAT = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

class ImportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("นำเข้าข้อมูล (Import Data)")
        self.resize(500, 400)
        self.layout = QVBoxLayout(self)

        # Header
        header = QLabel("วางรายชื่อตัวแปร (Format: s3 [tab/space] Description)")
        header.setStyleSheet("font-weight: bold; color: #d97706;")
        self.layout.addWidget(header)

        # Text Area
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("ตัวอย่าง:\ns3\ns4    อายุ\ns5    เพศ")
        self.text_edit.setStyleSheet("border: 1px solid #ccc; border-radius: 5px; padding: 5px;")
        self.layout.addWidget(self.text_edit)

        # Buttons
        btn_layout = QHBoxLayout()
        cancel_btn = QPushButton("ยกเลิก")
        cancel_btn.clicked.connect(self.reject)
        import_btn = QPushButton("นำเข้าสู่ Pool")
        import_btn.setStyleSheet("""
            QPushButton { background-color: #d97706; color: white; font-weight: bold; border-radius: 5px; padding: 8px; }
            QPushButton:hover { background-color: #b45309; }
        """)
        import_btn.clicked.connect(self.accept)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(import_btn)
        self.layout.addLayout(btn_layout)

    def get_data(self):
        return self.text_edit.toPlainText()

class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("วิธีการใช้")
        self.resize(520, 420)
        layout = QVBoxLayout(self)

        title = QLabel("วิธีการใช้ SPSS Table Generator")
        title.setStyleSheet("font-weight: bold; color: #1f2937;")
        layout.addWidget(title)

        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setStyleSheet("border: 1px solid #e5e7eb; border-radius: 6px; padding: 8px;")
        help_text.setPlainText(
            "1) อัปโหลดไฟล์ Project (.mtd) หรือไฟล์ Reporter เพื่ออ่านโครงสร้างตารางต้นแบบ ควรมี Table ต้นแบบ 1 Table\n"
            "2) โหลดตัวแปรเข้ารายการ Pool ได้ 2 วิธี:\n"
            "   - จากไฟล์ SPSS (.sav)\n"
            "   - นำเข้า Text (วางรายชื่อตัวแปร + คำอธิบาย)\n"
            "3) เลือกตัวแปรจาก Pool แล้วย้ายไป Target ด้วยปุ่ม → (ดับเบิลคลิกก็ได้)\n"
            "4) จัดลำดับใน Target ด้วยปุ่ม ↑ ↓ หรือย้ายกลับด้วยปุ่ม ←\n"
            "5) รวมหลายตัวแปรใน Table เดียว: เลือกแถวใน Target แล้วกด “🔗 รวม Table” ตั้งชื่อ\n"
            "   (หรือพิมพ์ชื่อในคอลัมน์ One Table / คอลัมน์ 'One Table' ในไฟล์ Setting) แถวที่ชื่อเหมือนกันจะรวมกัน\n"
            "6) กดปุ่ม “สร้างไฟล์ (Generate)” แล้วเลือกที่บันทึกไฟล์ .mtd ใหม่\n"
        )
        layout.addWidget(help_text)

        close_btn = QPushButton("ปิด")
        close_btn.clicked.connect(self.accept)
        close_btn.setStyleSheet("background-color: #e5e7eb; border-radius: 5px; padding: 6px 12px;")
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignRight)

class SPSSTableClonerApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Table Reporter_Generator v1")
        self.resize(1000,720)

        # State Variables
        self.file_content = ""
        self.pool_candidates = []
        # [NEW] ตัวนับลำดับ เพื่อใช้เรียงคืนตำแหน่งเดิม
        self.global_index_counter = 0

        # Setup UI
        self.setup_styles()
        self.init_ui()

    def setup_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f3f4f6; }
            QLabel { font-family: 'Segoe UI', sans-serif; font-size: 14px; color: #374151; }
            QGroupBox {
                font-weight: bold; border: 1px solid #e5e7eb; border-radius: 8px;
                margin-top: 20px; background-color: white;
            }
            QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; left: 10px; padding: 0 5px; color: #4b5563; }
            QWidget { color: #111827; }
            QLineEdit { border: 1px solid #d1d5db; border-radius: 4px; padding: 5px; background-color: white; color: #111827; }
            QLineEdit:focus { border: 2px solid #3b82f6; }
            QTextEdit { background-color: white; color: #111827; }
            QTreeWidget, QTableWidget, QListWidget {
                border: 1px solid #e5e7eb; border-radius: 6px; background-color: white; color: #111827;
                alternate-background-color: #f9fafb; gridline-color: #f3f4f6; padding: 5px;
            }
            QTreeWidget::item:selected, QTableWidget::item:selected { background-color: #dbeafe; }
            QHeaderView::section { background-color: #f9fafb; padding: 5px; border: none; font-weight: bold; color: #6b7280; }
            QTableCornerButton::section { background-color: #f9fafb; border: none; }
            QPushButton { color: #111827; }
            QMessageBox, QDialog { background-color: #f3f4f6; color: #111827; }
            QMessageBox QLabel { color: #111827; }
            QToolTip { background-color: #fffbe6; color: #111827; border: 1px solid #d1d5db; }
        """)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # --- Header ---
        header_frame = QFrame()
        header_frame.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #7c3aed, stop:1 #4f46e5); border-radius: 10px;")
        header_layout = QHBoxLayout(header_frame)
        title_label = QLabel("📊  Table Reporter_Generator v1")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: white; background: transparent;")
        subtitle_label = QLabel("Restores Position on Remove")
        subtitle_label.setStyleSheet("font-size: 14px; color: #e9d5ff; background: transparent;")

        help_btn = QPushButton("วิธีการใช้")
        help_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        help_btn.setStyleSheet("""
            QPushButton { background-color: rgba(255,255,255,0.15); color: white; border: 1px solid rgba(255,255,255,0.4);
                         border-radius: 6px; padding: 6px 12px; }
            QPushButton:hover { background-color: rgba(255,255,255,0.25); }
        """)
        help_btn.clicked.connect(self.open_help_dialog)

        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(subtitle_label)
        header_layout.addWidget(help_btn)
        main_layout.addWidget(header_frame)

        # --- Step 1: Upload ---
        group1 = QGroupBox("1. อัปโหลดไฟล์ Project (.mtd)")
        group1.setFixedHeight(100)
        layout1 = QHBoxLayout(group1)

        self.btn_browse = QPushButton("📂 เลือกไฟล์ .mtd...")
        self.btn_browse.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_browse.setStyleSheet("""
            QPushButton { background-color: #eff6ff; color: #2563eb; border: 1px solid #bfdbfe; border-radius: 6px; padding: 8px 16px; font-weight: bold; }
            QPushButton:hover { background-color: #dbeafe; }
        """)
        self.btn_browse.clicked.connect(self.browse_file)

        self.lbl_filename = QLabel("ยังไม่ได้เลือกไฟล์")
        self.lbl_filename.setStyleSheet("color: #6b7280; font-style: italic;")

        layout1.addWidget(self.btn_browse)
        layout1.addWidget(self.lbl_filename)
        layout1.addStretch()
        main_layout.addWidget(group1)

        # --- Step 2: Variables Manager ---
        group3 = QGroupBox("2. จัดการรายการตัวแปร")
        layout3 = QVBoxLayout(group3)

        # Toolbar
        toolbar3 = QHBoxLayout()
        btn_save_excel = QPushButton("💾 Save Settings")
        btn_save_excel.setStyleSheet("color: white; background-color: #059669; border: 1px solid #047857; border-radius: 5px; padding: 6px 12px; font-weight: bold;")
        btn_save_excel.clicked.connect(self.save_settings_to_excel)

        btn_load_excel = QPushButton("📂 Load Settings")
        btn_load_excel.setStyleSheet("color: white; background-color: #0891b2; border: 1px solid #0e7490; border-radius: 5px; padding: 6px 12px; font-weight: bold;")
        btn_load_excel.clicked.connect(self.load_settings_from_excel)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.VLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        sep.setStyleSheet("color: #ccc;")

        btn_load_spss = QPushButton("📥 SPSS (.sav)")
        btn_load_spss.setStyleSheet("color: #4338ca; background-color: #e0e7ff; border: 1px solid #c7d2fe; border-radius: 5px; padding: 6px 12px;")
        btn_load_spss.clicked.connect(self.load_spss_variables)

        btn_import_text = QPushButton("📝 นำเข้า Text")
        btn_import_text.setStyleSheet("color: #d97706; background-color: #fff7ed; border: 1px solid #fed7aa; border-radius: 5px; padding: 6px 12px;")
        btn_import_text.clicked.connect(self.open_import_dialog)

        btn_add_row = QPushButton("➕ แถวว่าง")
        btn_add_row.setStyleSheet("color: #2563eb; background-color: #eff6ff; border: 1px solid #bfdbfe; border-radius: 5px; padding: 6px 12px;")
        btn_add_row.clicked.connect(self.add_empty_row)

        btn_clear = QPushButton("🗑️ ล้าง")
        btn_clear.setStyleSheet("color: #dc2626; background-color: #fef2f2; border: 1px solid #fecaca; border-radius: 5px; padding: 6px 12px;")
        btn_clear.clicked.connect(self.clear_table)

        toolbar3.addWidget(btn_save_excel)
        toolbar3.addWidget(btn_load_excel)
        toolbar3.addWidget(sep)
        toolbar3.addWidget(btn_load_spss)
        toolbar3.addWidget(btn_import_text)
        toolbar3.addStretch()
        toolbar3.addWidget(btn_add_row)
        toolbar3.addWidget(btn_clear)
        layout3.addLayout(toolbar3)

        # --- Main 3-Column Layout ---
        h_split = QHBoxLayout()

        # 1. LEFT PANEL: POOL
        pool_layout = QVBoxLayout()
        pool_header = QLabel("🔻 ตัวแปรที่พบ (Pool)")
        pool_header.setStyleSheet("color: #d97706; font-weight: bold;")

        self.txt_search_pool = QLineEdit()
        self.txt_search_pool.setPlaceholderText("🔍 ค้นหา...")
        self.txt_search_pool.textChanged.connect(self.filter_pool)
        self.txt_search_pool.setStyleSheet("border: 1px solid #f59e0b; border-radius: 4px; padding: 6px;")

        self.tree_pool = QTreeWidget()
        self.tree_pool.setColumnCount(3)
        self.tree_pool.setHeaderLabels(["Variable Name", "Type", "Description"])
        self.tree_pool.header().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.tree_pool.header().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tree_pool.header().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.tree_pool.setAlternatingRowColors(True)
        self.tree_pool.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection) # Multi-select
        self.tree_pool.itemDoubleClicked.connect(self.on_pool_double_click)

        btn_select_all = QPushButton("เลือกทั้งหมด >>")
        btn_select_all.clicked.connect(self.move_all_from_pool)

        pool_layout.addWidget(pool_header)
        pool_layout.addWidget(self.txt_search_pool)
        pool_layout.addWidget(self.tree_pool)
        pool_layout.addWidget(btn_select_all)

        pool_widget = QWidget()
        pool_widget.setLayout(pool_layout)

        # 2. MIDDLE PANEL: ACTION BUTTONS (Square Style)
        mid_layout = QVBoxLayout()
        mid_layout.addStretch()

        # [MODIFIED] Square buttons with rounded corners (border-radius: 8px)
        self.btn_move_right = QPushButton("→")
        self.btn_move_right.setFixedSize(50, 50)
        self.btn_move_right.setStyleSheet("""
            QPushButton {
                font-size: 24px; font-weight: bold; color: white;
                background-color: #2563eb;
                border-radius: 8px; /* สี่เหลี่ยมมน */
            }
            QPushButton:hover { background-color: #1d4ed8; }
        """)
        self.btn_move_right.clicked.connect(self.move_selected_from_pool)

        self.btn_move_left = QPushButton("←")
        self.btn_move_left.setFixedSize(50, 50)
        self.btn_move_left.setStyleSheet("""
            QPushButton {
                font-size: 24px; font-weight: bold; color: white;
                background-color: #ef4444;
                border-radius: 8px; /* สี่เหลี่ยมมน */
            }
            QPushButton:hover { background-color: #dc2626; }
        """)
        self.btn_move_left.clicked.connect(self.remove_selected_from_target)

        mid_layout.addWidget(self.btn_move_right)
        mid_layout.addWidget(self.btn_move_left)
        mid_layout.addStretch()

        mid_widget = QWidget()
        mid_widget.setLayout(mid_layout)

        # 3. RIGHT PANEL: TARGET TABLE
        target_main_layout = QHBoxLayout() # To hold Table + Reorder Buttons

        # 3a. Table
        table_layout = QVBoxLayout()

        # Header + Generate Button
        target_header_layout = QHBoxLayout()
        target_header = QLabel("✅ ตัวแปรที่จะสร้าง (Target)")
        target_header.setStyleSheet("color: #059669; font-weight: bold;")

        self.btn_generate = QPushButton("💾 สร้างไฟล์ (Generate)")
        self.btn_generate.setEnabled(False)
        self.btn_generate.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_generate.setStyleSheet("""
            QPushButton { background-color: #e5e7eb; color: #9ca3af; border-radius: 5px; font-weight: bold; padding: 6px 15px; }
            QPushButton:enabled { background-color: #7c3aed; color: white; }
            QPushButton:enabled:hover { background-color: #6d28d9; }
        """)
        self.btn_generate.clicked.connect(self.process_file)

        self.btn_merge = QPushButton("🔗 รวม Table")
        self.btn_merge.setToolTip("เลือกหลายแถวแล้วกด เพื่อให้ตัวแปรเหล่านั้นอยู่ใน Table เดียวกัน\n(ใส่ชื่อว่าง = ยกเลิกการรวม)")
        self.btn_merge.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_merge.setStyleSheet("color: #065f46; background-color: #d1fae5; border: 1px solid #6ee7b7; border-radius: 5px; padding: 6px 12px; font-weight: bold;")
        self.btn_merge.clicked.connect(self.merge_selected_rows)

        target_header_layout.addWidget(target_header)
        target_header_layout.addStretch()
        target_header_layout.addWidget(self.btn_merge)
        target_header_layout.addWidget(self.btn_generate)

        self.table_target = QTableWidget()
        self.table_target.setColumnCount(3)
        self.table_target.setHorizontalHeaderLabels(["Variable Name", "Description", "One Table"])
        self.table_target.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table_target.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table_target.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table_target.setAlternatingRowColors(True)
        self.table_target.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows) # Select full row
        self.table_target.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection) # Multi-select

        table_layout.addLayout(target_header_layout)
        table_layout.addWidget(self.table_target)

        # 3b. Reorder Buttons (Right of Table)
        reorder_layout = QVBoxLayout()
        reorder_layout.setContentsMargins(5, 40, 0, 0) # Offset from top

        btn_up = QPushButton("↑")
        btn_up.setFixedSize(40, 40)
        btn_up.setStyleSheet("font-size: 20px; font-weight: bold; background-color: #f3f4f6; border: 1px solid #d1d5db; border-radius: 5px;")
        btn_up.clicked.connect(self.move_row_up)

        btn_down = QPushButton("↓")
        btn_down.setFixedSize(40, 40)
        btn_down.setStyleSheet("font-size: 20px; font-weight: bold; background-color: #f3f4f6; border: 1px solid #d1d5db; border-radius: 5px;")
        btn_down.clicked.connect(self.move_row_down)

        reorder_layout.addWidget(btn_up)
        reorder_layout.addWidget(btn_down)
        reorder_layout.addStretch()

        target_main_layout.addLayout(table_layout)
        target_main_layout.addLayout(reorder_layout)

        target_widget = QWidget()
        target_widget.setLayout(target_main_layout)

        # Combine 3 columns
        h_split.addWidget(pool_widget, stretch=4)
        h_split.addWidget(mid_widget, stretch=1)
        h_split.addWidget(target_widget, stretch=6)

        layout3.addLayout(h_split)
        main_layout.addWidget(group3, stretch=1)

    # --- Logic Functions ---

    def browse_file(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Select Project File", "", "SPSS Metadata (*.mtd);;XML Files (*.xml)")
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    self.file_content = f.read()
                self.lbl_filename.setText(f"ไฟล์ที่เลือก: {os.path.basename(filename)}")
                self.lbl_filename.setStyleSheet("color: #059669; font-weight: bold;")
                self.btn_generate.setEnabled(True)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"อ่านไฟล์ไม่ได้: {str(e)}")

    def save_settings_to_excel(self):
        if not HAS_PANDAS:
            QMessageBox.critical(self, "Missing Library", "กรุณาติดตั้ง pandas และ openpyxl\n(pip install pandas openpyxl)")
            return
        rows = self.table_target.rowCount()
        data = []
        for i in range(rows):
            name = self.table_target.item(i, 0).text().strip()
            desc = self.table_target.item(i, 1).text().strip()
            one = self.table_target.item(i, 2).text().strip() if self.table_target.item(i, 2) else ""
            if name: data.append({'Variable Name': name, 'Description': desc, 'One Table': one})

        if not data:
            QMessageBox.warning(self, "Warning", "ไม่มีข้อมูลให้บันทึก")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "Save Settings", "SPSS_Cloner_Settings.xlsx", "Excel Files (*.xlsx)")
        if save_path:
            try:
                df = pd.DataFrame(data)
                df.to_excel(save_path, index=False)
                QMessageBox.information(self, "Success", f"บันทึกแล้วที่:\n{save_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def load_settings_from_excel(self):
        if not HAS_PANDAS:
            QMessageBox.critical(self, "Missing Library", "กรุณาติดตั้ง pandas และ openpyxl")
            return
        filename, _ = QFileDialog.getOpenFileName(self, "Load Settings", "", "Excel Files (*.xlsx);;All Files (*.*)")
        if not filename: return
        try:
            df = pd.read_excel(filename)
            reply = QMessageBox.question(self, "Confirm", "ล้างรายการเดิมทิ้งก่อนหรือไม่?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes)
            if reply == QMessageBox.StandardButton.Yes: self.clear_table()

            count = 0
            def _cell(val):
                # เซลล์ว่างใน Excel จะอ่านมาเป็น NaN -> ต้องแปลงเป็นค่าว่าง ไม่ใช่ "nan"
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    return ""
                s = str(val).strip()
                return "" if s.lower() in ("nan", "none") else s

            for index, row in df.iterrows():
                name = _cell(row.get('Variable Name', ''))
                desc = _cell(row.get('Description', ''))
                one = _cell(row.get('One Table', ''))
                if name:
                    # For Excel imports, we might not have original order.
                    # If duplicate with Pool, remove from Pool.
                    self.add_row_to_table(name, desc, order=999999, one_table=one) # Assign generic high order

                    # Remove from pool if exists
                    self.pool_candidates = [c for c in self.pool_candidates if c['name'] != name]
                    count += 1

            self.refresh_pool()
            QMessageBox.information(self, "Success", f"โหลดข้อมูลสำเร็จ {count} รายการ")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def filter_pool(self, text):
        search_text = text.lower().strip()
        root = self.tree_pool.invisibleRootItem()
        child_count = root.childCount()
        for i in range(child_count):
            item = root.child(i)
            name = item.text(0).lower()
            desc = item.text(2).lower() # Description is now column 2
            if not search_text or (search_text in name or search_text in desc):
                item.setHidden(False)
            else:
                item.setHidden(True)

    def load_spss_variables(self):
        if not HAS_PYREADSTAT:
            QMessageBox.critical(self, "Missing Library", "กรุณาติดตั้ง pyreadstat\n(pip install pyreadstat)")
            return
        filename, _ = QFileDialog.getOpenFileName(self, "Select SPSS Data File", "", "SPSS Data (*.sav)")
        if not filename: return
        try:
            _, meta = pyreadstat.read_sav(filename, metadataonly=True)
            var_to_group_map = {}
            group_info_map = {}

            if hasattr(meta, 'multresp_defs'):
                for mdef in meta.multresp_defs:
                    name_match = re.search(r'^\s*(\$[a-zA-Z0-9_#@]+)', mdef)
                    label_match = re.search(r"'([^']*)'", mdef)
                    if name_match:
                        set_name = name_match.group(1).replace('$', '')
                        set_label = label_match.group(1) if label_match else ""
                        group_info_map[set_name] = {'desc': set_label, 'added': False}
                        if label_match:
                            vars_in_set = mdef[label_match.end():].strip().split()
                            for v in vars_in_set: var_to_group_map[v] = {'name': set_name, 'desc': set_label}

            pattern = re.compile(r'^(.+)(_O\d+|_C\d+)$', re.IGNORECASE)
            potential_groups = {}
            for col_name, col_label in zip(meta.column_names, meta.column_labels):
                if col_name in var_to_group_map: continue
                match = pattern.match(col_name)
                if match:
                    prefix = match.group(1)
                    if prefix not in potential_groups: potential_groups[prefix] = []
                    potential_groups[prefix].append({'name': col_name, 'desc': col_label})

            for prefix, members in potential_groups.items():
                if len(members) > 1:
                    set_name = prefix
                    set_label = members[0]['desc']
                    if set_name not in group_info_map:
                        group_info_map[set_name] = {'desc': set_label, 'added': False}
                    for m in members:
                        var_to_group_map[m['name']] = {'name': set_name, 'desc': set_label}

            existing_pool_names = set(item['name'] for item in self.pool_candidates)
            existing_target_names = set()
            for i in range(self.table_target.rowCount()):
                item = self.table_target.item(i, 0)
                if item: existing_target_names.add(item.text().strip())

            added_count = 0

            # [MODIFIED] Assign unique Order ID
            for col_name, col_label in zip(meta.column_names, meta.column_labels):

                # Check Group
                if col_name in var_to_group_map:
                    g_data = var_to_group_map[col_name]
                    g_name = g_data['name']
                    if not group_info_map[g_name]['added']:
                        if g_name not in existing_pool_names and g_name not in existing_target_names:
                            # Add Group
                            self.pool_candidates.append({
                                'name': g_name,
                                'desc': g_data['desc'],
                                'type': 'MA',
                                'order': self.global_index_counter # Store order
                            })
                            self.global_index_counter += 1
                            existing_pool_names.add(g_name)
                            added_count += 1
                        group_info_map[g_name]['added'] = True
                    continue

                if col_name in existing_pool_names or col_name in existing_target_names: continue

                # Determine Type for Single Variable
                # SA: Has labels
                # NA: Numeric, No labels
                # OA: String, No labels

                var_type = 'NA' # Default

                # Check for labels
                has_labels = False
                if hasattr(meta, 'variable_value_labels') and col_name in meta.variable_value_labels:
                    if meta.variable_value_labels[col_name]: # Check if dict is not empty
                        has_labels = True

                if has_labels:
                    var_type = 'SA'
                else:
                    # Check storage type
                    # pyreadstat meta.readstat_variable_types is a dict {var: type_name}
                    # type_name usually 'String', 'Double', etc.
                    v_type_str = meta.readstat_variable_types.get(col_name, 'Double')
                    print(f"Var: {col_name}, Type: {v_type_str}") # Debug
                    if 'string' in v_type_str.lower():
                        var_type = 'OA'
                    else:
                        var_type = 'NA'

                # Add Single Variable
                self.pool_candidates.append({
                    'name': col_name,
                    'desc': col_label,
                    'type': var_type,
                    'order': self.global_index_counter # Store order
                })
                self.global_index_counter += 1
                existing_pool_names.add(col_name)
                added_count += 1

            self.refresh_pool()
            QMessageBox.information(self, "Status", f"เพิ่มรายการใหม่: {added_count}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def open_import_dialog(self):
        dlg = ImportDialog(self)
        if dlg.exec():
            self.process_import_data(dlg.get_data())

    def process_import_data(self, text):
        lines = text.split('\n')
        existing_target_names = set()
        for i in range(self.table_target.rowCount()):
            item = self.table_target.item(i, 0)
            if item: existing_target_names.add(item.text().strip())

        new_items = []
        for line in lines:
            clean = line.replace('\ufeff', '').strip()
            if not clean: continue
            parts = clean.split('\t')
            if len(parts) == 1: parts = clean.split(',')
            if len(parts) == 1 and ' ' in clean:
                sp = clean.find(' ')
                parts = [clean[:sp], clean[sp+1:]]

            name = parts[0].strip()
            desc = parts[1].strip() if len(parts) > 1 else ""
            if not name: continue
            if name.lower() in ["name", "variable name", "var_name"]: continue
            if name in existing_target_names: continue

            new_items.append({
                'name': name,
                'desc': desc,
                'type': '', # Default empty for text import
                'order': self.global_index_counter # Assign order
            })
            self.global_index_counter += 1

        self.pool_candidates.extend(new_items)
        # Deduplicate
        u_map = {c['name']: c for c in self.pool_candidates}
        self.pool_candidates = list(u_map.values())
        self.refresh_pool()
        QMessageBox.information(self, "Status", f"นำเข้าสำเร็จ Total in Pool: {len(self.pool_candidates)}")

    def open_help_dialog(self):
        dlg = HelpDialog(self)
        dlg.exec()

    def refresh_pool(self):
        # [MODIFIED] Sort by original order before display
        self.pool_candidates.sort(key=lambda x: x.get('order', 999999))

        self.tree_pool.clear()
        for cand in self.pool_candidates:
            t_type = cand.get('type', '')
            item = QTreeWidgetItem([cand['name'], t_type, cand['desc']])

            # Color Logic
            color = None
            if t_type == 'SA': color = QColor("#1d4ed8") # Blue
            elif t_type == 'MA': color = QColor("#15803d") # Green
            elif t_type == 'NA': color = QColor("#c2410c") # Orange
            elif t_type == 'OA': color = QColor("#b91c1c") # Red

            if color:
                item.setForeground(1, color)
                font = item.font(1)
                font.setBold(True)
                item.setFont(1, font)

            # Store order in hidden data if needed, but we rely on pool_candidates list order now
            self.tree_pool.addTopLevelItem(item)
        self.txt_search_pool.clear()

    # --- ACTION BUTTON LOGIC ---

    def on_pool_double_click(self, item, column):
        self.move_selected_from_pool()

    def move_selected_from_pool(self):
        selected_items = self.tree_pool.selectedItems()
        if not selected_items: return

        names_to_move = set()

        # We need to find the full candidate object (with order) from the selected text
        # Since tree might be filtered, we search in pool_candidates

        # Optimization: Create map for fast lookup
        cand_map = {c['name']: c for c in self.pool_candidates}

        for item in selected_items:
            name = item.text(0)
            if name in cand_map:
                cand = cand_map[name]
                # Pass order to target
                self.add_row_to_table(cand['name'], cand['desc'], cand.get('order', 999999))
                names_to_move.add(name)

        # Remove from pool list
        self.pool_candidates = [c for c in self.pool_candidates if c['name'] not in names_to_move]

        curr_search = self.txt_search_pool.text()
        self.refresh_pool()
        if curr_search: self.txt_search_pool.setText(curr_search)

    def remove_selected_from_target(self):
        selected_rows = sorted(set(index.row() for index in self.table_target.selectedIndexes()), reverse=True)
        if not selected_rows: return

        for row in selected_rows:
            name = self.table_target.item(row, 0).text()
            desc = self.table_target.item(row, 1).text()

            # [MODIFIED] Retrieve original order
            order_data = self.table_target.item(row, 0).data(Qt.ItemDataRole.UserRole)
            original_order = order_data if order_data is not None else 999999

            # Return to Pool if not exists
            if not any(c['name'] == name for c in self.pool_candidates):
                self.pool_candidates.append({
                    'name': name,
                    'desc': desc,
                    'type': '', # Unknown type when returning from target (unless we store it)
                    'order': original_order
                })

            self.table_target.removeRow(row)

        self.refresh_pool() # This will re-sort based on 'order'

    def move_all_from_pool(self):
        search = self.txt_search_pool.text().lower().strip()
        to_move = []

        for cand in self.pool_candidates:
            if not search or (search in cand['name'].lower() or search in cand['desc'].lower()):
                to_move.append(cand)

        for cand in to_move:
            self.add_row_to_table(cand['name'], cand['desc'], cand.get('order', 999999))

        move_names = set(c['name'] for c in to_move)
        self.pool_candidates = [c for c in self.pool_candidates if c['name'] not in move_names]

        self.refresh_pool()
        self.txt_search_pool.setText(search)

    def move_row_up(self):
        row = self.table_target.currentRow()
        if row > 0:
            self.swap_rows(row, row - 1)
            self.table_target.selectRow(row - 1)

    def move_row_down(self):
        row = self.table_target.currentRow()
        if row < self.table_target.rowCount() - 1:
            self.swap_rows(row, row + 1)
            self.table_target.selectRow(row + 1)

    def swap_rows(self, row1, row2):
        # Must swap Items to preserve UserRole (Data/Order)
        items1 = [self.table_target.takeItem(row1, c) for c in range(3)]
        items2 = [self.table_target.takeItem(row2, c) for c in range(3)]
        for c in range(3):
            self.table_target.setItem(row1, c, items2[c])
            self.table_target.setItem(row2, c, items1[c])

    def add_empty_row(self):
        self.add_row_to_table("", "", order=999999)

    def add_row_to_table(self, name, desc, order=None, one_table=""):
        row = self.table_target.rowCount()
        self.table_target.insertRow(row)

        item_name = QTableWidgetItem(name)
        if order is not None:
            # Store original order in UserRole
            item_name.setData(Qt.ItemDataRole.UserRole, order)

        item_desc = QTableWidgetItem(desc)
        item_one = QTableWidgetItem(one_table or "")

        self.table_target.setItem(row, 0, item_name)
        self.table_target.setItem(row, 1, item_desc)
        self.table_target.setItem(row, 2, item_one)

    def merge_selected_rows(self):
        """กำหนดชื่อ One Table ให้แถวที่เลือก เพื่อรวมตัวแปรไว้ใน Table เดียวกัน"""
        rows = sorted(set(idx.row() for idx in self.table_target.selectedIndexes()))
        if not rows:
            QMessageBox.warning(self, "Warning", "กรุณาเลือกแถวใน Target ก่อน")
            return
        names = [self.table_target.item(r, 0).text().strip() for r in rows]
        # เดาชื่อเริ่มต้น: ตัดส่วนหลัง '#' / '$' ออก แล้วหา prefix ร่วม
        bases = [re.split(r'[#$]', n)[0] for n in names if n]
        default = os.path.commonprefix(bases).rstrip('_') if bases else ""
        current = self.table_target.item(rows[0], 2).text().strip() if self.table_target.item(rows[0], 2) else ""
        text, ok = QInputDialog.getText(self, "รวม Table",
                                        f"ตั้งชื่อ Table สำหรับ {len(rows)} ตัวแปร\n(เว้นว่าง = ยกเลิกการรวม)",
                                        text=current or default)
        if not ok: return
        text = text.strip()
        for r in rows:
            item = self.table_target.item(r, 2)
            if item is None:
                self.table_target.setItem(r, 2, QTableWidgetItem(text))
            else:
                item.setText(text)

    def clear_table(self):
        # [MODIFIED] Return all to pool before clearing?
        # Usually "Clear" implies delete. If user wants to return, they should select all -> move left.
        # But for safety, let's just clear target. Items lost from pool can be re-imported if needed.
        # Or better: Just clear.
        self.table_target.setRowCount(0)

    def process_file(self):
            try:
                content = self.file_content
                if not content:
                    QMessageBox.warning(self, "Warning", "กรุณาอัปโหลดไฟล์ Project (.mtd) ก่อน")
                    return

                # หาโครงสร้าง <Table> ต้นแบบ
                match_table = re.search(r'(<Table [^>]*?>.*?</Table>)', content, re.DOTALL)
                if not match_table: raise Exception("ไม่พบโครงสร้างตารางในไฟล์")
                original_xml = match_table.group(1)

                # หาชื่อตัวแปรต้นแบบ
                mdm = re.findall(r'MdmName="([^"]+)"', original_xml)
                v_mdm = [m for m in mdm if m.strip()]
                if v_mdm: source_v = v_mdm[0]
                else:
                    nms = re.findall(r'<Axis[^>]+Name="([^"]+)"', original_xml)
                    rsv = ["Side", "Top", "Bottom", "Dim", "Base", "Mean", "Median"]
                    vn = [n for n in nms if n not in rsv]
                    if vn: source_v = vn[0]
                    else: raise Exception("ระบุตัวแปรต้นแบบไม่ได้")

                # หาชื่อตาราง (Name attribute) ต้นแบบ
                nm_match = re.search(r'Name="([^"]+)"', original_xml)
                if not nm_match: raise Exception("หา attribute Name ไม่เจอ")
                orig_name = nm_match.group(1)

                # หาโครงสร้าง <Node> ต้นแบบ
                node_patt = re.compile(r'(<Node [^>]*?Table="' + re.escape(orig_name) + r'"[^>]*?/>)')
                match_node = node_patt.search(content)
                orig_node = match_node.group(1) if match_node else '<Node Name="xx" Description="xx" Table="xx"/>'

                rows_data = []
                for i in range(self.table_target.rowCount()):
                    n = self.table_target.item(i, 0).text().strip()
                    d = self.table_target.item(i, 1).text().strip()
                    o_item = self.table_target.item(i, 2)
                    o = o_item.text().strip() if o_item else ""
                    if n: rows_data.append({'name': n, 'desc': d, 'one': o})

                if not rows_data:
                    QMessageBox.warning(self, "Warning", "ไม่มีรายการตัวแปร")
                    return

                # จัดกลุ่ม: แถวที่มี One Table ชื่อเดียวกัน -> รวมเป็น Table เดียว (ตำแหน่งตามแถวแรกของกลุ่ม)
                final_vars = []
                group_index = {}
                for r in rows_data:
                    if r['one']:
                        if r['one'] in group_index:
                            final_vars[group_index[r['one']]]['members'].append(r)
                        else:
                            group_index[r['one']] = len(final_vars)
                            final_vars.append({'name': r['one'], 'desc': r['desc'] or r['one'],
                                               'members': [r], 'is_group': True})
                    else:
                        final_vars.append({'name': r['name'], 'desc': r['desc'], 'members': [r], 'is_group': False})

                gen_tabs = ""
                gen_nodes = ""

                # ฟังก์ชันสำหรับ Generate GUID ใหม่ทุกครั้งที่เจอแท็ก Id
                def generate_new_guid(match):
                    return f'Id="{{{str(uuid.uuid4()).upper()}}}"'

                for v in final_vars:
                    members = v['members']
                    new_n = members[0]['name']          # ตัวแปรแรก ใช้แทนที่ตัวแปรต้นแบบ
                    clean_n = re.sub(r'[^A-Za-z0-9_]', '_', v['name'])
                    new_tb_name = f"Table_{clean_n}"

                    # เตรียมข้อความ Description (ก่อนนำไปคลีน)
                    if v['is_group']:
                        raw_t_desc = v['desc']
                        raw_n_desc = f"{v['name']} : {v['desc']}" if v['desc'] != v['name'] else v['name']
                    else:
                        raw_t_desc = v['desc'] if v['desc'] else new_n
                        raw_n_desc = f"{new_n} : {v['desc']}" if v['desc'] else new_n

                    # [สำคัญ!] แปลงอักขระพิเศษให้ปลอดภัยกับ XML (เช่น & เป็น &amp;)
                    t_desc = escape(raw_t_desc, {'"': '&quot;', "'": '&apos;'})
                    n_desc = escape(raw_n_desc, {'"': '&quot;', "'": '&apos;'})

                    # --- จัดการ Table XML ---
                    tmp = original_xml
                    tmp = tmp.replace(f'"{source_v}"', f'"{new_n}"')
                    tmp = tmp.replace(f'>{source_v}<', f'>{new_n}<')
                    tmp = tmp.replace(f'Name="{orig_name}"', f'Name="{new_tb_name}"')
                    tmp = tmp.replace('IsPopulated="true"', 'IsPopulated="false"')

                    # --- รวมหลายตัวแปรใน Table เดียว: โคลน <Axis> ของตัวแปรแรกให้ตัวแปรที่เหลือ ---
                    if len(members) > 1:
                        esc_n = re.escape(new_n)
                        m_axis = re.search(r'<Axis Name="' + esc_n + r'" MdmName="' + esc_n + r'"[^>]*>.*?</Axis>', tmp, re.DOTALL)
                        m_head = re.search(r'<Heading Name="' + esc_n + r'"\s*/>', tmp)
                        if not m_axis:
                            raise Exception(f"ไม่พบ Axis ของ '{new_n}' ในต้นแบบ จึงรวม Table '{v['name']}' ไม่ได้")
                        axis_block = m_axis.group(0)
                        extra_axes, extra_heads = "", ""
                        for mem in members[1:]:
                            mn = mem['name']
                            blk = axis_block.replace(f'"{new_n}"', f'"{mn}"').replace(f'>{new_n}<', f'>{mn}<')
                            if mem['desc']:
                                blk = re.sub(r'^(<Axis [^>]*?Label=")[^"]*(")', lambda mm: mm.group(1) + escape(mem['desc'], {'"': '&quot;'}) + mm.group(2), blk, count=1)
                            extra_axes += blk
                            extra_heads += f'<Heading Name="{mn}"/>'
                        # แทรก Heading ก่อน (อยู่หลัง Axis ใน string) แล้วค่อยแทรก Axis
                        if m_head:
                            tmp = tmp[:m_head.end()] + extra_heads + tmp[m_head.end():]
                        tmp = tmp[:m_axis.end()] + extra_axes + tmp[m_axis.end():]

                    # [สำคัญ!] ใช้ Regex กวาดเปลี่ยน Id "ทุกจุด" ที่อยู่ในโครงสร้าง Table นี้ให้เป็นของใหม่ทั้งหมด
                    tmp = re.sub(r'Id="\{?[A-Fa-f0-9\-]+\}?"', generate_new_guid, tmp)

                    if 'Description="' in tmp:
                        tmp = re.sub(r'Description="[^"]*"', f'Description="{t_desc}"', tmp, count=1)

                    gen_tabs += "\n" + tmp

                    # --- จัดการ Node XML ---
                    tmp_n = orig_node
                    tmp_n = tmp_n.replace(f'Table="{orig_name}"', f'Table="{new_tb_name}"')
                    tmp_n = tmp_n.replace(f'Name="{orig_name}"', f'Name="{new_tb_name}"')

                    # เปลี่ยน Id ของ Node
                    tmp_n = re.sub(r'Id="\{?[A-Fa-f0-9\-]+\}?"', generate_new_guid, tmp_n)

                    if 'Description="' in tmp_n:
                        tmp_n = re.sub(r'Description="[^"]*"', f'Description="{n_desc}"', tmp_n, count=1)
                    else:
                        tmp_n = tmp_n.replace('<Node ', f'<Node Description="{n_desc}" ')

                    # กันเหนียว: ถ้า orig_node เดิมไม่มี Id ให้บังคับสร้างใส่เข้าไป
                    if 'Id="' not in tmp_n:
                        tmp_n = tmp_n.replace('<Node ', f'<Node Id="{{{str(uuid.uuid4()).upper()}}}" ')

                    gen_nodes += "\n" + tmp_n

                # แทนที่ข้อมูลกลับเข้าไป
                if '</Tables>' in content:
                    content = content.replace('</Tables>', gen_tabs + '\n</Tables>')
                if '</GroupedTables>' in content:
                    last_g = content.rfind('</GroupedTables>')
                    content = content[:last_g] + gen_nodes + content[last_g:]

                # บันทึกไฟล์
                save_path, _ = QFileDialog.getSaveFileName(self, "Save File", "Fixed.mtd", "SPSS Metadata (*.mtd)")
                if save_path:
                    with open(save_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                    QMessageBox.information(self, "Success", f"เสร็จสิ้น! (Template: {source_v})")

            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))




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
        app = QApplication(sys.argv)
        # ใช้ Fusion + palette สว่างแบบตายตัว เพื่อไม่ให้ Windows Dark Mode ทำให้ตัวหนังสือกลืนกับพื้นขาว
        app.setStyle("Fusion")
        pal = QPalette()
        pal.setColor(QPalette.ColorRole.Window, QColor("#f3f4f6"))
        pal.setColor(QPalette.ColorRole.WindowText, QColor("#111827"))
        pal.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
        pal.setColor(QPalette.ColorRole.AlternateBase, QColor("#f9fafb"))
        pal.setColor(QPalette.ColorRole.Text, QColor("#111827"))
        pal.setColor(QPalette.ColorRole.Button, QColor("#e5e7eb"))
        pal.setColor(QPalette.ColorRole.ButtonText, QColor("#111827"))
        pal.setColor(QPalette.ColorRole.ToolTipBase, QColor("#fffbe6"))
        pal.setColor(QPalette.ColorRole.ToolTipText, QColor("#111827"))
        pal.setColor(QPalette.ColorRole.Highlight, QColor("#3b82f6"))
        pal.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
        pal.setColor(QPalette.ColorRole.PlaceholderText, QColor("#9ca3af"))
        app.setPalette(pal)
        font = QFont("Segoe UI", 10)
        app.setFont(font)
        window = SPSSTableClonerApp()
        window.show()
        sys.exit(app.exec())


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