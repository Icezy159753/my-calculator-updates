import sys
import os
import ast
import random
import itertools
from collections import defaultdict

import pandas as pd
import pyreadstat  # For reading SPSS files
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QListWidget, QListWidgetItem,
    QTableWidget, QTableWidgetItem, QHeaderView, QTextEdit, QTabWidget,
    QFrame, QFileDialog, QMessageBox, QCheckBox, QGroupBox, QSplitter,
    QProgressBar, QStatusBar, QAbstractItemView, QSizePolicy, QScrollArea,
    QGridLayout, QSpacerItem, QDialog, QDialogButtonBox, QStackedWidget
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QColor, QIcon, QPalette

# ==================== MODERN CLEAN STYLING ====================
MODERN_STYLE = """
* {
    font-family: 'Segoe UI', 'Tahoma', sans-serif;
    font-size: 13px;
}

QMainWindow, QDialog {
    background-color: #f0f4f8;
}

QWidget {
    color: #1a202c;
}

QDialog {
    background-color: #f0f4f8;
}

QScrollArea {
    border: none;
    background-color: transparent;
}

QFrame#cardFrame {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
}

QFrame#quotaCard {
    background-color: #ffffff;
    border: 2px solid #e2e8f0;
    border-radius: 10px;
    padding: 15px;
}

QFrame#quotaCard:hover {
    border-color: #3182ce;
}

QPushButton {
    background-color: #3182ce;
    color: #ffffff;
    border: none;
    border-radius: 8px;
    padding: 12px 24px;
    font-size: 13px;
    font-weight: 600;
    min-width: 120px;
}

QPushButton:hover {
    background-color: #2b6cb0;
}

QPushButton:pressed {
    background-color: #1a4b7c;
}

QPushButton:disabled {
    background-color: #a0aec0;
    color: #e2e8f0;
}

QPushButton#primaryBtn {
    background-color: #38a169;
    color: #ffffff;
    font-size: 16px;
    padding: 16px 40px;
    min-width: 200px;
}

QPushButton#primaryBtn:hover {
    background-color: #2f855a;
}

QPushButton#primaryBtn:disabled {
    background-color: #9ae6b4;
}

QPushButton#dangerBtn {
    background-color: #e53e3e;
    color: #ffffff;
}

QPushButton#dangerBtn:hover {
    background-color: #c53030;
}

QPushButton#exportBtn {
    background-color: #805ad5;
    color: #ffffff;
    font-size: 16px;
    padding: 16px 40px;
    min-width: 200px;
}

QPushButton#exportBtn:hover {
    background-color: #6b46c1;
}

QPushButton#exportBtn:disabled {
    background-color: #b794f4;
}

QPushButton#quotaBtn {
    background-color: #edf2f7;
    color: #2d3748;
    border: 2px solid #e2e8f0;
    padding: 10px;
    min-height: 55px;
    font-size: 12px;
}

QPushButton#quotaBtn:hover {
    background-color: #e2e8f0;
    border-color: #3182ce;
}

QPushButton#quotaBtnActive {
    background-color: #c6f6d5;
    color: #276749;
    border: 2px solid #38a169;
    padding: 10px;
    min-height: 55px;
    font-size: 12px;
}

QPushButton#quotaBtnActive:hover {
    background-color: #9ae6b4;
}

QPushButton#quotaBtnMain {
    background-color: #bee3f8;
    color: #2b6cb0;
    border: 2px solid #3182ce;
    padding: 10px;
    min-height: 55px;
    font-size: 12px;
    font-weight: bold;
}

QPushButton#quotaBtnMain:hover {
    background-color: #90cdf4;
}

QPushButton#confirmBtn {
    background-color: #38a169;
    color: #ffffff;
    font-size: 14px;
    padding: 14px 30px;
}

QPushButton#cancelBtn {
    background-color: #718096;
    color: #ffffff;
    font-size: 14px;
    padding: 14px 30px;
}

QLineEdit {
    background-color: #ffffff;
    border: 2px solid #cbd5e0;
    border-radius: 6px;
    padding: 10px 14px;
    font-size: 13px;
    color: #1a202c;
}

QLineEdit:focus {
    border: 2px solid #3182ce;
}

QComboBox {
    background-color: #ffffff;
    border: 2px solid #cbd5e0;
    border-radius: 6px;
    padding: 10px 14px;
    font-size: 13px;
    color: #1a202c;
    min-width: 150px;
}

QComboBox:hover {
    border-color: #3182ce;
}

QComboBox::drop-down {
    border: none;
    width: 30px;
}

QComboBox::down-arrow {
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #4a5568;
    margin-right: 10px;
}

QComboBox QAbstractItemView {
    background-color: #ffffff;
    border: 2px solid #cbd5e0;
    selection-background-color: #ebf8ff;
    selection-color: #1a202c;
    padding: 5px;
}

QListWidget {
    background-color: #ffffff;
    border: 2px solid #cbd5e0;
    border-radius: 6px;
    padding: 5px;
    font-size: 12px;
    color: #1a202c;
}

QListWidget::item {
    padding: 8px 12px;
    border-radius: 4px;
    margin: 2px 0;
    color: #1a202c;
}

QListWidget::item:selected {
    background-color: #bee3f8;
    color: #1a202c;
}

QListWidget::item:hover:!selected {
    background-color: #edf2f7;
}

QTableWidget {
    background-color: #ffffff;
    border: 2px solid #cbd5e0;
    border-radius: 6px;
    gridline-color: #e2e8f0;
    font-size: 12px;
    color: #1a202c;
}

QTableWidget::item {
    padding: 10px;
    color: #1a202c;
}

QTableWidget::item:selected {
    background-color: #bee3f8;
    color: #1a202c;
}

QHeaderView::section {
    background-color: #4a5568;
    color: #ffffff;
    padding: 12px;
    border: none;
    font-weight: bold;
    font-size: 12px;
}

QTextEdit {
    background-color: #1a202c;
    color: #68d391;
    border: none;
    border-radius: 8px;
    padding: 15px;
    font-family: 'Consolas', 'Monaco', monospace;
    font-size: 12px;
}

QCheckBox {
    spacing: 10px;
    font-size: 14px;
    color: #1a202c;
}

QCheckBox::indicator {
    width: 24px;
    height: 24px;
    border-radius: 6px;
    border: 2px solid #cbd5e0;
    background-color: #ffffff;
}

QCheckBox::indicator:checked {
    background-color: #38a169;
    border-color: #38a169;
}

QStatusBar {
    background-color: #ffffff;
    color: #4a5568;
    font-size: 12px;
    padding: 8px;
    border-top: 1px solid #e2e8f0;
}

QLabel {
    font-size: 13px;
    color: #1a202c;
}

QLabel#titleLabel {
    font-size: 26px;
    font-weight: bold;
    color: #1a202c;
}

QLabel#sectionTitle {
    font-size: 16px;
    font-weight: bold;
    color: #1a202c;
}

QLabel#stepTitle {
    font-size: 15px;
    font-weight: bold;
    color: #ffffff;
    background-color: #3182ce;
    padding: 10px 20px;
    border-radius: 8px;
}

QLabel#stepTitleGreen {
    font-size: 15px;
    font-weight: bold;
    color: #ffffff;
    background-color: #38a169;
    padding: 10px 20px;
    border-radius: 8px;
}

QLabel#stepTitlePurple {
    font-size: 15px;
    font-weight: bold;
    color: #ffffff;
    background-color: #805ad5;
    padding: 10px 20px;
    border-radius: 8px;
}

QScrollBar:vertical {
    background-color: #edf2f7;
    width: 12px;
    border-radius: 6px;
}

QScrollBar::handle:vertical {
    background-color: #a0aec0;
    border-radius: 6px;
    min-height: 30px;
}

QScrollBar::handle:vertical:hover {
    background-color: #718096;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
"""


# ==================== SAMPLING ALGORITHM (UNCHANGED) ====================
def flexible_quota_sampling(population_df, quota_definitions, id_col='id'):
    """ทำการสุ่มตัวอย่างแบบ Greedy (Logic unchanged)"""
    if id_col not in population_df.columns:
        raise ValueError(f"population_df must contain an '{id_col}' column.")

    num_quotas = len(quota_definitions)
    if num_quotas == 0:
        raise ValueError("No quota definitions provided.")

    all_dims = set()
    for i, (dims, targets) in enumerate(quota_definitions):
        if not isinstance(dims, list): 
            raise ValueError(f"Quota Set {i+1} Dimensions must be a list")
        if not isinstance(targets, dict): 
            raise ValueError(f"Quota Set {i+1} Targets must be a dict")
        all_dims.update(dims)

    for dim in all_dims:
        if dim not in population_df.columns:
            raise ValueError(f"Column '{dim}' not found")

    available_candidates_df = population_df.copy()
    available_candidates_df[id_col] = available_candidates_df[id_col].astype(str)
    available_ids = set(available_candidates_df[id_col])
    selected_ids = set()
    current_counts = [defaultdict(int) for _ in range(num_quotas)]
    quota_keys_cols = [f'quota_key_{i}' for i in range(num_quotas)]

    def get_quota_key(row, dimensions):
        try:
            values = [str(row[dim]) for dim in dimensions]
            if any(pd.isna(row[dim]) for dim in dimensions): 
                return None
            return tuple(values)
        except:
            return None

    for i, (dims, _) in enumerate(quota_definitions):
        available_candidates_df[f'quota_key_{i}'] = available_candidates_df.apply(
            lambda row: get_quota_key(row, dims), axis=1
        )

    available_candidates_df = available_candidates_df.dropna(subset=quota_keys_cols)
    available_ids = set(available_candidates_df[id_col])

    if len(available_candidates_df) == 0:
        return set(), [defaultdict(int) for _ in range(num_quotas)], [{} for _ in range(num_quotas)]

    total_target_size = sum(quota_definitions[0][1].values())
    max_iterations = total_target_size * 3
    iteration_count = 0

    while len(selected_ids) < total_target_size and iteration_count < max_iterations:
        iteration_count += 1
        
        needs_list = []
        for i, (_, targets) in enumerate(quota_definitions):
            stringified = {tuple(map(str, k)): v for k, v in targets.items()}
            needs = {k: t - current_counts[i].get(k, 0) for k, t in stringified.items() 
                    if t - current_counts[i].get(k, 0) > 0}
            needs_list.append(needs)

        if all(sum(n.values()) == 0 for n in needs_list):
            break
        if not available_ids:
            break

        current_df = available_candidates_df[available_candidates_df[id_col].isin(available_ids)].copy()
        if current_df.empty:
            break

        current_df['score'] = 0
        for i in range(num_quotas):
            current_df['score'] += current_df[f'quota_key_{i}'].map(needs_list[i]).fillna(0)

        candidates = current_df[current_df['score'] > 0]
        if candidates.empty:
            break

        top = candidates[candidates['score'] == candidates['score'].max()]
        selected_row = top.sample(n=1).iloc[0]
        best_id = selected_row[id_col]

        selected_ids.add(best_id)
        available_ids.remove(best_id)
        for i in range(num_quotas):
            key = selected_row[f'quota_key_{i}']
            if key:
                current_counts[i][key] += 1

    # Calculate results
    unmet_list = []
    final_counts = []
    for i, (_, targets) in enumerate(quota_definitions):
        stringified = {tuple(map(str, k)): v for k, v in targets.items()}
        unmet = {k: t - current_counts[i].get(k, 0) for k, t in stringified.items() 
                if t - current_counts[i].get(k, 0) > 0}
        unmet_list.append(unmet)
        final_counts.append(dict(current_counts[i]))

    return selected_ids, final_counts, unmet_list


# ==================== QUOTA CONFIGURATION DIALOG (WIZARD STYLE) ====================
class QuotaConfigDialog(QDialog):
    """Popup dialog for configuring a Quota Set - Wizard Style"""
    
    def __init__(self, index, parent_app, existing_data=None):
        super().__init__(parent_app)
        self.index = index
        self.parent_app = parent_app
        self.existing_data = existing_data or {'dimensions': [], 'targets': {}}
        self.dim_comboboxes = []
        self.current_dims = []
        self.result_data = None
        self.current_step = 0
        
        self.setWindowTitle(f"กำหนด Quota Set {index + 1}")
        self.setMinimumSize(950, 650)
        self.setModal(True)
        
        self.setup_ui()
        self.load_existing_data()
        self.show_step(0)
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(25, 20, 25, 20)
        
        # === Title ===
        if self.index == 0:
            title = QLabel(f"⭐ Quota Set {self.index + 1} (หลัก)")
            title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2b6cb0;")
        else:
            title = QLabel(f"📋 Quota Set {self.index + 1}")
            title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2d3748;")
        layout.addWidget(title)
        
        # === Step Indicator ===
        self.step_indicator = QLabel("")
        self.step_indicator.setStyleSheet("font-size: 14px; color: #718096; padding: 5px;")
        layout.addWidget(self.step_indicator)
        
        # === Stacked Widget for Steps ===
        self.stack = QStackedWidget()
        
        # --- Step 1: Dimension Selection ---
        step1_widget = QWidget()
        step1_layout = QVBoxLayout(step1_widget)
        step1_layout.setSpacing(15)
        
        step1_title = QLabel("🔍 STEP 1: เลือกข้อที่ใช้ตัด Quota")
        step1_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #3182ce; padding: 10px; background-color: #ebf8ff; border-radius: 8px;")
        step1_layout.addWidget(step1_title)
        
        hint = QLabel("💡 คลิกที่แถวใดก็ได้เพื่อเลือก/ยกเลิกเลือก (สามารถเลือกได้หลายข้อ)")
        hint.setStyleSheet("color: #4a5568; font-size: 13px; padding: 5px;")
        step1_layout.addWidget(hint)
        
        # Dimension table
        self.dim_table = QTableWidget()
        self.dim_table.setColumnCount(2)
        self.dim_table.setHorizontalHeaderLabels(["✓", "ชื่อคอลัมน์"])
        self.dim_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.dim_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.dim_table.setColumnWidth(0, 60)
        self.dim_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.dim_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.dim_table.verticalHeader().setVisible(False)
        self.dim_table.setStyleSheet("""
            QTableWidget { background-color: white; border: 2px solid #3182ce; border-radius: 8px; font-size: 14px; }
            QTableWidget::item { padding: 10px; color: #1a202c; }
            QTableWidget::item:selected { background-color: #bee3f8; }
            QHeaderView::section { background-color: #3182ce; color: white; padding: 10px; font-weight: bold; }
        """)
        
        self.dim_checkboxes = []
        if self.parent_app.cleaned_df is not None:
            cols = [c for c in self.parent_app.cleaned_df.columns if c != self.parent_app.id_col_target]
            self.dim_table.setRowCount(len(cols))
            for row, col in enumerate(cols):
                self.dim_table.setRowHeight(row, 40)
                
                cb = QCheckBox()
                cb.setStyleSheet("""
                    QCheckBox::indicator { width: 24px; height: 24px; border-radius: 4px; border: 2px solid #3182ce; background-color: white; }
                    QCheckBox::indicator:checked { background-color: #38a169; border-color: #38a169; }
                """)
                cb_container = QWidget()
                cb_layout_inner = QHBoxLayout(cb_container)
                cb_layout_inner.addWidget(cb)
                cb_layout_inner.setAlignment(Qt.AlignmentFlag.AlignCenter)
                cb_layout_inner.setContentsMargins(0, 0, 0, 0)
                
                self.dim_checkboxes.append(cb)
                self.dim_table.setCellWidget(row, 0, cb_container)
                
                item = QTableWidgetItem(col)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.dim_table.setItem(row, 1, item)
        
        self.dim_table.cellClicked.connect(self.toggle_dim_checkbox)
        step1_layout.addWidget(self.dim_table)
        
        self.stack.addWidget(step1_widget)
        
        # --- Step 2: Condition Selection ---
        step2_widget = QWidget()
        step2_layout = QVBoxLayout(step2_widget)
        step2_layout.setSpacing(15)
        
        step2_title = QLabel("📋 STEP 2: เลือกเงื่อนไข + กำหนดจำนวน N")
        step2_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #276749; padding: 10px; background-color: #f0fff4; border-radius: 8px;")
        step2_layout.addWidget(step2_title)
        
        self.selected_dims_label = QLabel("")
        self.selected_dims_label.setStyleSheet("color: #4a5568; font-size: 13px; padding: 5px;")
        step2_layout.addWidget(self.selected_dims_label)
        
        # Main content area - split into left (conditions) and right (quota table)
        step2_content = QHBoxLayout()
        
        # Left side - Condition selection
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 10, 0)
        
        # Scroll area for conditions
        condition_scroll = QScrollArea()
        condition_scroll.setWidgetResizable(True)
        condition_scroll.setStyleSheet("QScrollArea { border: 1px solid #e2e8f0; border-radius: 8px; background-color: white; }")
        
        self.condition_container = QWidget()
        self.condition_layout = QVBoxLayout(self.condition_container)
        self.condition_layout.setSpacing(12)
        self.condition_layout.setContentsMargins(15, 15, 15, 15)
        
        condition_scroll.setWidget(self.condition_container)
        left_layout.addWidget(condition_scroll, stretch=1)
        
        # Add button row
        add_frame = QFrame()
        add_frame.setStyleSheet("background-color: #f0fff4; border-radius: 8px; padding: 10px;")
        add_layout = QHBoxLayout(add_frame)
        
        add_layout.addWidget(QLabel("จำนวน N ="))
        self.count_input = QLineEdit()
        self.count_input.setPlaceholderText("เช่น 100")
        self.count_input.setFixedWidth(100)
        add_layout.addWidget(self.count_input)
        
        self.btn_add = QPushButton("➕ เพิ่ม Quota")
        self.btn_add.setStyleSheet("background-color: #38a169; color: white; font-weight: bold; padding: 10px 20px;")
        self.btn_add.clicked.connect(self.add_target)
        add_layout.addWidget(self.btn_add)
        add_layout.addStretch()
        
        left_layout.addWidget(add_frame)
        step2_content.addWidget(left_panel, stretch=1)
        
        # Right side - Quota table showing defined quotas
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(10, 0, 0, 0)
        
        quota_table_title = QLabel("📊 Quota ที่กำหนดแล้ว")
        quota_table_title.setStyleSheet("font-size: 14px; font-weight: bold; color: #805ad5; padding: 5px;")
        right_layout.addWidget(quota_table_title)
        
        self.step2_total_label = QLabel("รวม: 0 รายการ | N = 0")
        self.step2_total_label.setStyleSheet("font-size: 12px; color: #6b46c1; padding: 3px;")
        right_layout.addWidget(self.step2_total_label)
        
        self.step2_quota_table = QTableWidget()
        self.step2_quota_table.setColumnCount(4)
        self.step2_quota_table.setHorizontalHeaderLabels(["#", "เงื่อนไข", "N", "จัดการ"])
        self.step2_quota_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.step2_quota_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.step2_quota_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.step2_quota_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.step2_quota_table.setColumnWidth(0, 35)
        self.step2_quota_table.setColumnWidth(2, 60)
        self.step2_quota_table.setColumnWidth(3, 80)
        self.step2_quota_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.step2_quota_table.setStyleSheet("""
            QTableWidget { background-color: white; border: 2px solid #805ad5; border-radius: 8px; font-size: 11px; }
            QTableWidget::item { padding: 5px; color: #1a202c; }
            QHeaderView::section { background-color: #805ad5; color: white; padding: 6px; font-weight: bold; font-size: 11px; }
        """)
        self.step2_quota_table.setMinimumWidth(280)
        right_layout.addWidget(self.step2_quota_table, stretch=1)
        
        # Quick action buttons
        quick_btn_layout = QHBoxLayout()
        self.step2_btn_clear = QPushButton("🗑️ ล้างทั้งหมด")
        self.step2_btn_clear.setStyleSheet("background-color: #718096; color: white; font-weight: bold; padding: 6px 10px; font-size: 11px;")
        self.step2_btn_clear.clicked.connect(self.clear_all)
        quick_btn_layout.addWidget(self.step2_btn_clear)
        quick_btn_layout.addStretch()
        right_layout.addLayout(quick_btn_layout)
        
        step2_content.addWidget(right_panel, stretch=1)
        step2_layout.addLayout(step2_content, stretch=1)
        
        self.stack.addWidget(step2_widget)
        
        # --- Step 3: Target Review ---
        step3_widget = QWidget()
        step3_layout = QVBoxLayout(step3_widget)
        step3_layout.setSpacing(15)
        
        step3_title = QLabel("✅ STEP 3: รายการ Quota ที่กำหนด")
        step3_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #6b46c1; padding: 10px; background-color: #faf5ff; border-radius: 8px;")
        step3_layout.addWidget(step3_title)
        
        self.total_label = QLabel("📊 รวม: 0 รายการ | Total N = 0")
        self.total_label.setStyleSheet("font-weight: bold; color: #6b46c1; font-size: 15px; padding: 5px;")
        step3_layout.addWidget(self.total_label)
        
        self.target_table = QTableWidget()
        self.target_table.setColumnCount(3)
        self.target_table.setHorizontalHeaderLabels(["#", "เงื่อนไข (Key)", "จำนวน N"])
        self.target_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.target_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.target_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.target_table.setColumnWidth(0, 50)
        self.target_table.setColumnWidth(2, 100)
        self.target_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.target_table.setStyleSheet("""
            QTableWidget { background-color: white; border: 2px solid #805ad5; border-radius: 8px; }
            QTableWidget::item { padding: 8px; color: #1a202c; }
            QHeaderView::section { background-color: #805ad5; color: white; padding: 10px; font-weight: bold; }
        """)
        step3_layout.addWidget(self.target_table)
        
        btn_row = QHBoxLayout()
        self.btn_remove = QPushButton("🗑️ ลบที่เลือก")
        self.btn_remove.setStyleSheet("background-color: #e53e3e; color: white; font-weight: bold; padding: 10px 15px;")
        self.btn_remove.clicked.connect(self.remove_selected)
        btn_row.addWidget(self.btn_remove)
        
        self.btn_clear = QPushButton("🗑️ ล้างทั้งหมด")
        self.btn_clear.setStyleSheet("background-color: #718096; color: white; font-weight: bold; padding: 10px 15px;")
        self.btn_clear.clicked.connect(self.clear_all)
        btn_row.addWidget(self.btn_clear)
        btn_row.addStretch()
        step3_layout.addLayout(btn_row)
        
        self.stack.addWidget(step3_widget)
        layout.addWidget(self.stack, stretch=1)
        
        # === Navigation Buttons ===
        nav_layout = QHBoxLayout()
        
        self.btn_back = QPushButton("⬅️ ย้อนกลับ")
        self.btn_back.setStyleSheet("background-color: #718096; color: white; font-weight: bold; padding: 12px 25px;")
        self.btn_back.clicked.connect(self.go_back)
        nav_layout.addWidget(self.btn_back)
        
        nav_layout.addStretch()
        
        self.btn_cancel = QPushButton("❌ ยกเลิก")
        self.btn_cancel.setStyleSheet("background-color: #e53e3e; color: white; padding: 12px 25px;")
        self.btn_cancel.clicked.connect(self.reject)
        nav_layout.addWidget(self.btn_cancel)
        
        self.btn_next = QPushButton("ถัดไป ➡️")
        self.btn_next.setStyleSheet("background-color: #3182ce; color: white; font-weight: bold; padding: 12px 25px;")
        self.btn_next.clicked.connect(self.go_next)
        nav_layout.addWidget(self.btn_next)
        
        self.btn_save = QPushButton("✅ บันทึก")
        self.btn_save.setStyleSheet("background-color: #38a169; color: white; font-weight: bold; padding: 12px 30px;")
        self.btn_save.clicked.connect(self.confirm_and_save)
        nav_layout.addWidget(self.btn_save)
        
        layout.addLayout(nav_layout)
    
    def show_step(self, step):
        """Show specific step"""
        self.current_step = step
        self.stack.setCurrentIndex(step)
        
        # Update step indicator
        steps = ["1️⃣ เลือกข้อ", "2️⃣ กำหนดเงื่อนไข", "3️⃣ ตรวจสอบ"]
        indicator = " → ".join([f"**{s}**" if i == step else s for i, s in enumerate(steps)])
        self.step_indicator.setText(f"ขั้นตอน: {steps[step]}")
        
        # Update navigation buttons
        self.btn_back.setVisible(step > 0)
        self.btn_next.setVisible(step < 2)
        self.btn_save.setVisible(step == 2)
    
    def go_next(self):
        """Go to next step"""
        if self.current_step == 0:
            # Validate step 1 - must select dimensions
            dims = [self.dim_table.item(i, 1).text() for i, cb in enumerate(self.dim_checkboxes) if cb.isChecked()]
            if not dims:
                QMessageBox.warning(self, "Error", "กรุณาเลือกอย่างน้อย 1 ข้อ")
                return
            self.current_dims = dims
            self.build_condition_ui()
            self.show_step(1)
        elif self.current_step == 1:
            self.update_total()
            self.show_step(2)
    
    def go_back(self):
        """Go to previous step"""
        if self.current_step > 0:
            self.show_step(self.current_step - 1)
    
    def build_condition_ui(self):
        """Build condition selection UI for Step 2"""
        # Clear existing
        while self.condition_layout.count():
            item = self.condition_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        self.dim_comboboxes = []
        self.selected_dims_label.setText(f"📌 ข้อที่เลือก: {', '.join(self.current_dims)}")
        
        df = self.parent_app.cleaned_df
        
        for dim in self.current_dims:
            unique_values = df[dim].astype(str).dropna().unique()
            unique_values = sorted([v for v in unique_values if v and v.lower() not in ['nan', 'none', '<na>']])
            
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0, 0, 0, 0)
            
            lbl = QLabel(f"{dim}:")
            lbl.setStyleSheet("font-weight: bold; font-size: 14px; color: #276749; min-width: 150px;")
            lbl.setMinimumWidth(150)
            row_layout.addWidget(lbl)
            
            combo = QComboBox()
            combo.addItems(unique_values)
            combo.setMinimumWidth(300)
            combo.setStyleSheet("font-size: 14px; padding: 8px; min-height: 30px;")
            row_layout.addWidget(combo, stretch=1)
            self.dim_comboboxes.append(combo)
            
            self.condition_layout.addWidget(row_widget)
        
        self.condition_layout.addStretch()
    
    def toggle_dim_checkbox(self, row, col):
        """Toggle checkbox when clicking row"""
        if row < len(self.dim_checkboxes):
            cb = self.dim_checkboxes[row]
            cb.setChecked(not cb.isChecked())
    
    def load_existing_data(self):
        """Load existing configuration"""
        dims = self.existing_data.get('dimensions', [])
        targets = self.existing_data.get('targets', {})
        
        # Select dimensions
        for i, cb in enumerate(self.dim_checkboxes):
            item = self.dim_table.item(i, 1)
            if item and item.text() in dims:
                cb.setChecked(True)
        
        if dims:
            self.current_dims = dims
            self.build_condition_ui()
        
        for key, count in targets.items():
            self.add_target_row(key, count)
        self.sync_step2_table()
        self.update_total()
        
        # If already has data, jump to step 3
        if targets:
            self.show_step(2)
    
    def add_target(self):
        """Add target"""
        if not self.dim_comboboxes:
            QMessageBox.warning(self, "Error", "กรุณายืนยันข้อที่เลือกก่อน")
            return
        
        values = [combo.currentText() for combo in self.dim_comboboxes]
        if not all(values):
            QMessageBox.warning(self, "Error", "กรุณาเลือกค่าทุกช่อง")
            return
        
        count_str = self.count_input.text().strip()
        if not count_str:
            QMessageBox.warning(self, "Error", "กรุณากรอกจำนวน N")
            return
        
        try:
            count = int(count_str)
            if count < 0:
                raise ValueError()
        except:
            QMessageBox.warning(self, "Error", "จำนวน N ต้องเป็นตัวเลข >= 0")
            return
        
        self.add_target_row(tuple(values), count)
        self.count_input.clear()
        self.update_total()
    
    def add_target_row(self, key, count):
        """Add row to both tables (target_table and step2_quota_table)"""
        # Update existing row if key exists
        for row in range(self.target_table.rowCount()):
            if self.target_table.item(row, 1).text() == str(key):
                self.target_table.item(row, 2).setText(str(count))
                self.sync_step2_table()
                return
        
        # Add new row to target_table (Step 3)
        row = self.target_table.rowCount()
        self.target_table.insertRow(row)
        self.target_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        self.target_table.setItem(row, 1, QTableWidgetItem(str(key)))
        self.target_table.setItem(row, 2, QTableWidgetItem(str(count)))
        
        # Sync to step2_quota_table
        self.sync_step2_table()
    
    def sync_step2_table(self):
        """Sync step2_quota_table with target_table"""
        self.step2_quota_table.setRowCount(0)
        total_n = 0
        
        for row in range(self.target_table.rowCount()):
            key_item = self.target_table.item(row, 1)
            count_item = self.target_table.item(row, 2)
            
            if key_item and count_item:
                key_str = key_item.text()
                count = int(count_item.text())
                total_n += count
                
                # Add row to step2_quota_table
                s2_row = self.step2_quota_table.rowCount()
                self.step2_quota_table.insertRow(s2_row)
                self.step2_quota_table.setRowHeight(s2_row, 35)
                
                self.step2_quota_table.setItem(s2_row, 0, QTableWidgetItem(str(s2_row + 1)))
                
                # Format key for display (shorter)
                try:
                    key_tuple = ast.literal_eval(key_str)
                    display_key = ', '.join(str(k) for k in key_tuple)
                except:
                    display_key = key_str
                
                key_display_item = QTableWidgetItem(display_key)
                key_display_item.setToolTip(key_str)  # Show full key on hover
                self.step2_quota_table.setItem(s2_row, 1, key_display_item)
                
                self.step2_quota_table.setItem(s2_row, 2, QTableWidgetItem(str(count)))
                
                # Create action buttons widget
                action_widget = QWidget()
                action_layout = QHBoxLayout(action_widget)
                action_layout.setContentsMargins(2, 2, 2, 2)
                action_layout.setSpacing(2)
                
                edit_btn = QPushButton("✏️")
                edit_btn.setFixedSize(30, 25)
                edit_btn.setStyleSheet("background-color: #3182ce; color: white; border-radius: 4px; font-size: 12px; padding: 0px;")
                edit_btn.setToolTip("แก้ไขจำนวน N")
                edit_btn.clicked.connect(lambda checked, r=s2_row: self.edit_quota_row(r))
                
                del_btn = QPushButton("🗑️")
                del_btn.setFixedSize(30, 25)
                del_btn.setStyleSheet("background-color: #e53e3e; color: white; border-radius: 4px; font-size: 12px; padding: 0px;")
                del_btn.setToolTip("ลบรายการนี้")
                del_btn.clicked.connect(lambda checked, r=s2_row: self.delete_quota_row(r))
                
                action_layout.addWidget(edit_btn)
                action_layout.addWidget(del_btn)
                
                self.step2_quota_table.setCellWidget(s2_row, 3, action_widget)
        
        # Update step2 total label
        count_rows = self.step2_quota_table.rowCount()
        self.step2_total_label.setText(f"รวม: {count_rows} รายการ | N = {total_n}")
    
    def edit_quota_row(self, row):
        """Edit quota row - change N value"""
        if row >= self.target_table.rowCount():
            return
        
        key_item = self.target_table.item(row, 1)
        count_item = self.target_table.item(row, 2)
        
        if not key_item or not count_item:
            return
        
        current_count = count_item.text()
        
        # Simple input dialog
        from PyQt6.QtWidgets import QInputDialog
        new_count, ok = QInputDialog.getInt(
            self, "แก้ไขจำนวน N", 
            f"เงื่อนไข: {key_item.text()}\n\nจำนวน N ใหม่:",
            int(current_count), 0, 999999, 1
        )
        
        if ok:
            self.target_table.item(row, 2).setText(str(new_count))
            self.sync_step2_table()
            self.update_total()
    
    def delete_quota_row(self, row):
        """Delete quota row"""
        if row >= self.target_table.rowCount():
            return
        
        reply = QMessageBox.question(
            self, "ยืนยันการลบ", 
            f"ลบ Quota รายการที่ {row + 1}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.target_table.removeRow(row)
            self.renumber_rows()
            self.sync_step2_table()
            self.update_total()
    
    def remove_selected(self):
        """Remove selected"""
        rows = set(item.row() for item in self.target_table.selectedItems())
        for row in sorted(rows, reverse=True):
            self.target_table.removeRow(row)
        self.renumber_rows()
        self.sync_step2_table()
        self.update_total()
    
    def clear_all(self):
        """Clear all"""
        if self.target_table.rowCount() == 0:
            return
        reply = QMessageBox.question(self, "ยืนยัน", "ล้างทั้งหมด?")
        if reply == QMessageBox.StandardButton.Yes:
            self.target_table.setRowCount(0)
            self.sync_step2_table()
            self.update_total()
    
    def renumber_rows(self):
        """Renumber rows"""
        for row in range(self.target_table.rowCount()):
            self.target_table.item(row, 0).setText(str(row + 1))
    
    def update_total(self):
        """Update totals"""
        count = self.target_table.rowCount()
        total_n = sum(int(self.target_table.item(row, 2).text()) 
                     for row in range(count) if self.target_table.item(row, 2))
        self.total_label.setText(f"📊 รวม: {count} รายการ | Total N = {total_n}")
    
    def get_targets_data(self):
        """Get targets as dict"""
        targets = {}
        for row in range(self.target_table.rowCount()):
            try:
                key_str = self.target_table.item(row, 1).text()
                count = int(self.target_table.item(row, 2).text())
                key = ast.literal_eval(key_str)
                targets[tuple(map(str, key))] = count
            except:
                pass
        return targets
    
    def confirm_and_save(self):
        """Confirm and save data"""
        targets = self.get_targets_data()
        
        if self.index == 0:
            if not self.current_dims:
                QMessageBox.warning(self, "Error", "Quota Set หลัก ต้องเลือกอย่างน้อย 1 ข้อ")
                return
            if not targets:
                QMessageBox.warning(self, "Error", "Quota Set หลัก ต้องมีอย่างน้อย 1 target")
                return
        
        self.result_data = {
            'dimensions': self.current_dims,
            'targets': targets
        }
        self.accept()


# ==================== QUOTA PREVIEW DIALOG ====================
class QuotaPreviewDialog(QDialog):
    """Preview all configured quotas"""
    
    def __init__(self, parent_app):
        super().__init__(parent_app)
        self.parent_app = parent_app
        
        self.setWindowTitle("👁️ พรีวิว Quota ทั้งหมด")
        self.setMinimumSize(800, 600)
        self.setModal(True)
        
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Title
        title = QLabel("📊 พรีวิว Quota ที่กำหนดไว้ทั้งหมด")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #2d3748;")
        layout.addWidget(title)
        
        hint = QLabel("💡 คลิกที่แต่ละ Set เพื่อแก้ไข")
        hint.setStyleSheet("color: #718096; font-size: 13px;")
        layout.addWidget(hint)
        
        # Tabs for each quota set
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane { border: 2px solid #e2e8f0; border-radius: 8px; background: white; }
            QTabBar::tab { padding: 10px 20px; font-weight: bold; }
            QTabBar::tab:selected { background-color: #3182ce; color: white; }
        """)
        
        for i in range(self.parent_app.num_quota_sets):
            data = self.parent_app.quota_data[i]
            dims = data.get('dimensions', [])
            targets = data.get('targets', {})
            
            tab = QWidget()
            tab_layout = QVBoxLayout(tab)
            tab_layout.setSpacing(10)
            
            if dims or targets:
                # Has configuration
                dims_label = QLabel(f"📌 ข้อที่เลือก: {', '.join(dims) if dims else 'ไม่มี'}")
                dims_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #3182ce; padding: 10px; background-color: #ebf8ff; border-radius: 5px;")
                tab_layout.addWidget(dims_label)
                
                if targets:
                    table = QTableWidget()
                    table.setColumnCount(2)
                    table.setHorizontalHeaderLabels(["เงื่อนไข", "จำนวน N"])
                    table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
                    table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
                    table.setColumnWidth(1, 100)
                    table.setRowCount(len(targets))
                    table.setStyleSheet("""
                        QTableWidget { border: 1px solid #e2e8f0; }
                        QHeaderView::section { background-color: #4a5568; color: white; padding: 8px; }
                    """)
                    
                    total_n = 0
                    for row, (key, count) in enumerate(targets.items()):
                        key_str = ', '.join(str(k) for k in key)
                        table.setItem(row, 0, QTableWidgetItem(key_str))
                        table.setItem(row, 1, QTableWidgetItem(str(count)))
                        total_n += count
                    
                    tab_layout.addWidget(table)
                    
                    total_label = QLabel(f"📊 รวม: {len(targets)} เงื่อนไข | Total N = {total_n}")
                    total_label.setStyleSheet("font-weight: bold; color: #38a169; font-size: 14px; padding: 10px; background-color: #f0fff4; border-radius: 5px;")
                    tab_layout.addWidget(total_label)
                else:
                    no_target = QLabel("⚠️ มีข้อที่เลือกแต่ยังไม่มี Target")
                    no_target.setStyleSheet("color: #ed8936; font-size: 14px; padding: 20px;")
                    tab_layout.addWidget(no_target)
                
                # Edit button
                btn_edit = QPushButton(f"✏️ แก้ไข Set {i+1}")
                btn_edit.setStyleSheet("background-color: #3182ce; color: white; font-weight: bold; padding: 10px;")
                btn_edit.clicked.connect(lambda checked, idx=i: self.edit_set(idx))
                tab_layout.addWidget(btn_edit)
            else:
                # No configuration
                empty_label = QLabel("❌ ยังไม่ได้กำหนด Quota")
                empty_label.setStyleSheet("color: #718096; font-size: 16px; padding: 30px;")
                empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                tab_layout.addWidget(empty_label)
                
                btn_config = QPushButton(f"➕ กำหนด Set {i+1}")
                btn_config.setStyleSheet("background-color: #38a169; color: white; font-weight: bold; padding: 10px;")
                btn_config.clicked.connect(lambda checked, idx=i: self.edit_set(idx))
                tab_layout.addWidget(btn_config)
            
            tab_layout.addStretch()
            
            tab_title = f"⭐ Set {i+1}" if i == 0 else f"Set {i+1}"
            if targets:
                tab_title += f" ({len(targets)})"
            self.tabs.addTab(tab, tab_title)
        
        layout.addWidget(self.tabs)
        
        # Close button
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        btn_close = QPushButton("✅ ปิด")
        btn_close.setStyleSheet("background-color: #38a169; color: white; font-weight: bold; padding: 12px 30px;")
        btn_close.clicked.connect(self.accept)
        btn_layout.addWidget(btn_close)
        
        layout.addLayout(btn_layout)
    
    def edit_set(self, index):
        """Open editor for specific set"""
        dialog = QuotaConfigDialog(index, self.parent_app, self.parent_app.quota_data[index])
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.parent_app.quota_data[index] = dialog.result_data
            self.parent_app.update_quota_button(index)
            self.parent_app.log(f"✅ แก้ไข Quota Set {index+1} แล้ว")
            # Refresh preview
            self.accept()
            QMessageBox.information(self, "บันทึกแล้ว", f"บันทึก Set {index+1} เรียบร้อย!\n\nกดพรีวิวอีกครั้งเพื่อดูผลลัพธ์")


# ==================== MAIN APPLICATION ====================
class QuotaSamplerApp(QMainWindow):
    """Main Application Window"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("โปรแกรมตัด Quota")
        self.setMinimumSize(1100, 750)
        
        self.loaded_df = None
        self.cleaned_df = None
        self.sampling_results = None
        self.id_col_target = 'id'
        self.num_quota_sets = 7
        self.quota_data = [{} for _ in range(self.num_quota_sets)]  # Store quota configs
        self.quota_buttons = []
        self.current_file_path = ""
        
        self.setup_ui()
        self.center_window()
        
    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 15, 20, 15)
        
        # === HEADER ===
        header = QFrame()
        header.setObjectName("cardFrame")
        header.setStyleSheet("QFrame#cardFrame { background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 15px; }")
        header_layout = QVBoxLayout(header)
        header_layout.setSpacing(10)
        
        # Title
        title_row = QHBoxLayout()
        title = QLabel("📊 โปรแกรมตัด Quota Pro V1")
        title.setObjectName("titleLabel")
        title_row.addWidget(title)
        title_row.addStretch()
        header_layout.addLayout(title_row)
        
        # Subtitle
        subtitle = QLabel("โปรแกรมตัด Quota สำหรับงานวิจัย | เลือกไฟล์ Excel/SPSS → กำหนด Quota → ตัดข้อมูล → Export")
        subtitle.setStyleSheet("color: #718096; font-size: 13px;")
        header_layout.addWidget(subtitle)
        
        # Controls row
        controls = QHBoxLayout()
        
        self.btn_open = QPushButton("📂 เปิดไฟล์ข้อมูล")
        self.btn_open.setFixedWidth(180)
        self.btn_open.clicked.connect(self.load_dataset)
        controls.addWidget(self.btn_open)
        
        self.file_label = QLabel("ยังไม่ได้โหลดไฟล์")
        self.file_label.setStyleSheet("color: #718096; padding: 8px 15px; background-color: #edf2f7; border-radius: 6px;")
        controls.addWidget(self.file_label)
        
        controls.addStretch()
        
        self.btn_save_settings = QPushButton("💾 Save Settings")
        self.btn_save_settings.setFixedWidth(150)
        self.btn_save_settings.clicked.connect(self.save_settings)
        controls.addWidget(self.btn_save_settings)
        
        self.btn_load_settings = QPushButton("📥 Load Settings")
        self.btn_load_settings.setFixedWidth(150)
        self.btn_load_settings.clicked.connect(self.load_settings)
        controls.addWidget(self.btn_load_settings)
        
        header_layout.addLayout(controls)
        main_layout.addWidget(header)
        
        # === QUOTA SETS GRID ===
        quota_frame = QFrame()
        quota_frame.setObjectName("cardFrame")
        quota_frame.setStyleSheet("QFrame#cardFrame { background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; }")
        quota_layout = QVBoxLayout(quota_frame)
        quota_layout.setContentsMargins(20, 20, 20, 20)
        
        section_title = QLabel("📋 กำหนด Quota Sets (คลิกเพื่อกำหนดแต่ละ Set)")
        section_title.setObjectName("sectionTitle")
        quota_layout.addWidget(section_title)
        
        hint = QLabel("💡 คลิกที่ปุ่มด้านล่างเพื่อเปิดหน้าต่างกำหนด Quota แต่ละชุด")
        hint.setStyleSheet("color: #718096; font-size: 12px; margin-bottom: 10px;")
        quota_layout.addWidget(hint)
        
        # Button grid (2 rows)
        grid = QGridLayout()
        grid.setSpacing(15)
        
        for i in range(self.num_quota_sets):
            btn = QPushButton()
            if i == 0:
                btn.setText(f"⭐ Set {i+1} (หลัก)\n❌ ยังไม่ได้กำหนด")
                btn.setObjectName("quotaBtnMain")
            else:
                btn.setText(f"Set {i+1}\n❌ ยังไม่ได้กำหนด")
                btn.setObjectName("quotaBtn")
            
            btn.clicked.connect(lambda checked, idx=i: self.open_quota_dialog(idx))
            
            row = i // 4
            col = i % 4
            grid.addWidget(btn, row, col)
            self.quota_buttons.append(btn)
        
        quota_layout.addLayout(grid)
        main_layout.addWidget(quota_frame)
        
        # === ACTIONS + LOG ===
        bottom = QFrame()
        bottom.setObjectName("cardFrame")
        bottom.setStyleSheet("QFrame#cardFrame { background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; }")
        bottom_layout = QVBoxLayout(bottom)
        bottom_layout.setContentsMargins(20, 15, 20, 15)
        
        # Action buttons
        action_row = QHBoxLayout()
        
        self.btn_run = QPushButton("🚀 เริ่มตัด Quota!!")
        self.btn_run.setObjectName("primaryBtn")
        self.btn_run.clicked.connect(self.run_sampling)
        self.btn_run.setEnabled(False)
        action_row.addWidget(self.btn_run)
        
        self.btn_preview = QPushButton("👁️ พรีวิว Quota")
        self.btn_preview.setStyleSheet("background-color: #ed8936; color: white; font-weight: bold; padding: 14px 30px; min-width: 160px;")
        self.btn_preview.clicked.connect(self.preview_quota)
        action_row.addWidget(self.btn_preview)
        
        action_row.addStretch()
        
        self.btn_export = QPushButton("📤 Export Excel")
        self.btn_export.setObjectName("exportBtn")
        self.btn_export.clicked.connect(self.export_results)
        self.btn_export.setEnabled(False)
        action_row.addWidget(self.btn_export)
        
        bottom_layout.addLayout(action_row)
        
        # Log
        log_label = QLabel("📝 Log Output")
        log_label.setStyleSheet("font-weight: bold; font-size: 14px; margin-top: 10px;")
        bottom_layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(100)
        self.log_text.setMaximumHeight(130)
        bottom_layout.addWidget(self.log_text)
        
        main_layout.addWidget(bottom)
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("พร้อมใช้งาน")
        
        self.log("🚀 พร้อมใช้งาน - กรุณาโหลด Rawdata Excel")
    
    def center_window(self):
        screen = QApplication.primaryScreen().geometry()
        self.move((screen.width() - self.width()) // 2, (screen.height() - self.height()) // 2)
    
    def log(self, msg):
        self.log_text.append(msg)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())
    
    def clear_log(self):
        self.log_text.clear()
    
    def update_quota_button(self, index):
        """Update button display based on quota data"""
        data = self.quota_data[index]
        btn = self.quota_buttons[index]
        
        dims = data.get('dimensions', [])
        targets = data.get('targets', {})
        total_n = sum(targets.values()) if targets else 0
        
        if targets:
            if index == 0:
                btn.setText(f"⭐ Set {index+1} (หลัก)\n✅ N={total_n}")
                btn.setObjectName("quotaBtnMain")
            else:
                btn.setText(f"Set {index+1}\n✅ N={total_n}")
                btn.setObjectName("quotaBtnActive")
        else:
            if index == 0:
                btn.setText(f"⭐ Set {index+1} (หลัก)\n❌ ยังไม่ได้กำหนด")
                btn.setObjectName("quotaBtnMain")
            else:
                btn.setText(f"Set {index+1}\n❌ ยังไม่ได้กำหนด")
                btn.setObjectName("quotaBtn")
        
        # Re-apply style
        btn.style().unpolish(btn)
        btn.style().polish(btn)
    
    def open_quota_dialog(self, index):
        """Open quota configuration dialog"""
        if self.cleaned_df is None:
            QMessageBox.warning(self, "Error", "กรุณาโหลดไฟล์ข้อมูลก่อน")
            return
        
        dialog = QuotaConfigDialog(index, self, self.quota_data[index])
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.quota_data[index] = dialog.result_data
            self.update_quota_button(index)
            self.log(f"✅ บันทึก Quota Set {index+1} แล้ว")
    
    def load_dataset(self):
        """Load Excel or SPSS file"""
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Open Data File", "", 
            "Data Files (*.xlsx *.xls *.sav);;Excel (*.xlsx *.xls);;SPSS (*.sav);;All (*.*)"
        )
        if not filepath:
            return
        
        self.log(f"📂 Loading: {os.path.basename(filepath)}")
        
        try:
            # Determine file type and load accordingly
            file_ext = os.path.splitext(filepath)[1].lower()
            
            if file_ext == '.sav':
                # Load SPSS file with value labels
                self.log("📊 Detected SPSS file (.sav)")
                temp_df, meta = pyreadstat.read_sav(filepath)
                
                # Convert codes to labels using value_labels from metadata
                value_labels = meta.variable_value_labels  # Dict of {column: {code: label}}
                for col, labels_dict in value_labels.items():
                    if col in temp_df.columns and labels_dict:
                        # Map codes to labels
                        temp_df[col] = temp_df[col].map(labels_dict).fillna(temp_df[col])
                
                self.loaded_df = temp_df
                self.log(f"📝 Applied value labels to {len(value_labels)} columns")
            else:
                # Load Excel file
                self.log("📊 Detected Excel file")
                self.loaded_df = pd.read_excel(filepath, sheet_name=0)
                temp_df = self.loaded_df.copy()
            
            temp_df = self.loaded_df.copy()
            
            if temp_df.empty:
                raise ValueError("ไฟล์ว่าง")
            
            first_col = temp_df.columns[0]
            temp_df[first_col] = temp_df[first_col].astype(str)
            temp_df.rename(columns={first_col: self.id_col_target}, inplace=True)
            
            self.cleaned_df = temp_df
            self.current_file_path = filepath
            
            self.file_label.setText(f"✅ {os.path.basename(filepath)}")
            self.file_label.setStyleSheet("color: #276749; font-weight: bold; padding: 8px 15px; background-color: #c6f6d5; border-radius: 6px;")
            
            self.btn_run.setEnabled(True)
            self.sampling_results = None
            
            cols = len([c for c in temp_df.columns if c != self.id_col_target])
            file_type = "SPSS" if file_ext == '.sav' else "Excel"
            self.status_bar.showMessage(f"โหลด {file_type} สำเร็จ: {len(temp_df)} records | {cols} columns")
            self.log(f"✅ โหลด {file_type} สำเร็จ: {len(temp_df)} records, {cols} columns")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error:\n{e}")
            self.log(f"❌ Error: {e}")
    
    def run_sampling(self):
        """Run sampling"""
        if self.cleaned_df is None:
            return
        
        self.clear_log()
        self.log("🚀 เริ่มประมวลผล...")
        self.btn_run.setEnabled(False)
        self.btn_export.setEnabled(False)
        QApplication.processEvents()
        
        try:
            quota_definitions = []
            
            for i, data in enumerate(self.quota_data):
                dims = data.get('dimensions', [])
                targets = data.get('targets', {})
                
                if i == 0:
                    if not dims or not targets:
                        raise ValueError("Quota Set 1 ต้องมีข้อและ target")
                else:
                    if not dims or not targets:
                        continue
                
                quota_definitions.append((dims, targets))
                self.log(f"📊 Set {i+1}: {len(targets)} targets")
            
            if not quota_definitions:
                raise ValueError("ไม่มี Quota definition")
            
            results = flexible_quota_sampling(self.cleaned_df, quota_definitions, id_col=self.id_col_target)
            
            self.sampling_results = results
            self.quota_definitions_used = quota_definitions  # Store for export summary
            selected_ids, final_counts, unmet = results
            
            q1_sum = sum(quota_definitions[0][1].values())
            self.log(f"\n✅ สำเร็จ! Selected: {len(selected_ids)} / Target: {q1_sum}")
            
            # Build detailed summary message
            summary_lines = []
            all_met = True
            
            for i, (dims, targets) in enumerate(quota_definitions):
                set_label = f"Set {i+1}" + (" (หลัก)" if i == 0 else "")
                summary_lines.append(f"\n📊 Quota {set_label}:")
                summary_lines.append(f"   ข้อ: {', '.join(dims)}")
                
                for key, target_n in targets.items():
                    actual_n = final_counts[i].get(key, 0)
                    key_str = ', '.join(str(k) for k in key) if isinstance(key, tuple) else str(key)
                    
                    if actual_n >= target_n:
                        status = "✅"
                    else:
                        status = "⚠️"
                        all_met = False
                    
                    summary_lines.append(f"   {status} {key_str}: Target={target_n}, Actual={actual_n}")
            
            # Show summary in log
            for line in summary_lines:
                self.log(line)
            
            self.btn_export.setEnabled(True)
            self.status_bar.showMessage(f"สำเร็จ! Selected: {len(selected_ids)}")
            
            # Show result message
            if all_met:
                result_msg = f"✅ ตัด Quota เรียบร้อย!\n\n✅ ได้ครบทุก Target!\n\nSelected: {len(selected_ids)} records\n\nกด Export Excel ได้เลย!"
                QMessageBox.information(self, "สำเร็จ - ได้ครบ!", result_msg)
            else:
                result_msg = f"⚠️ ตัด Quota เสร็จแล้ว แต่บาง Target ได้ไม่ครบ\n\n"
                result_msg += f"Selected: {len(selected_ids)} records\n"
                result_msg += f"Target: {q1_sum}\n\n"
                result_msg += "ดูรายละเอียดใน Log หรือ Export Excel เพื่อดู Summary"
                QMessageBox.warning(self, "ไม่ครบ Target", result_msg)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error:\n{e}")
            self.log(f"❌ Error: {e}")
        finally:
            self.btn_run.setEnabled(True)
    
    def export_results(self):
        """Export results with Summary sheet"""
        if self.sampling_results is None:
            return
        
        filepath, _ = QFileDialog.getSaveFileName(self, "Save", "", "Excel (*.xlsx)")
        if not filepath:
            return
        
        if not filepath.endswith('.xlsx'):
            filepath += '.xlsx'
        
        self.log(f"📤 Exporting...")
        
        try:
            selected_ids, final_counts, unmet = self.sampling_results
            quota_definitions = getattr(self, 'quota_definitions_used', [])
            
            df_export = self.cleaned_df.copy()
            df_export[self.id_col_target] = df_export[self.id_col_target].astype(str)
            selected_df = df_export[df_export[self.id_col_target].isin(selected_ids)]
            
            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                # Sheet 1: Selected data
                selected_df.to_excel(writer, sheet_name='Selected', index=False)
                self.log(f"✅ Sheet 'Selected': {len(selected_df)} rows")
                
                # Sheet 2: Summary
                if quota_definitions:
                    summary_data = []
                    
                    for i, (dims, targets) in enumerate(quota_definitions):
                        set_label = f"Set {i+1}" + (" (หลัก)" if i == 0 else "")
                        dims_str = ', '.join(dims)
                        
                        for key, target_n in targets.items():
                            actual_n = final_counts[i].get(key, 0)
                            key_str = ', '.join(str(k) for k in key) if isinstance(key, tuple) else str(key)
                            diff = actual_n - target_n
                            status = "✅ ครบ" if actual_n >= target_n else "⚠️ ไม่ครบ"
                            
                            summary_data.append({
                                'Quota Set': set_label,
                                'ข้อที่ใช้ตัด': dims_str,
                                'เงื่อนไข': key_str,
                                'Target': target_n,
                                'Actual': actual_n,
                                'Diff': diff,
                                'สถานะ': status
                            })
                    
                    if summary_data:
                        summary_df = pd.DataFrame(summary_data)
                        summary_df.to_excel(writer, sheet_name='Summary', index=False)
                        
                        # Format Summary sheet
                        workbook = writer.book
                        worksheet = writer.sheets['Summary']
                        
                        # Header format
                        header_format = workbook.add_format({
                            'bold': True,
                            'bg_color': '#4a5568',
                            'font_color': 'white',
                            'border': 1,
                            'align': 'center',
                            'valign': 'vcenter'
                        })
                        
                        # Write headers with format
                        for col_num, value in enumerate(summary_df.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                        
                        # Conditional format for status
                        green_format = workbook.add_format({'bg_color': '#c6f6d5', 'font_color': '#276749'})
                        red_format = workbook.add_format({'bg_color': '#fed7d7', 'font_color': '#c53030'})
                        
                        # Set column widths
                        worksheet.set_column('A:A', 15)  # Quota Set
                        worksheet.set_column('B:B', 25)  # ข้อที่ใช้ตัด
                        worksheet.set_column('C:C', 30)  # เงื่อนไข
                        worksheet.set_column('D:E', 10)  # Target, Actual
                        worksheet.set_column('F:F', 8)   # Diff
                        worksheet.set_column('G:G', 12)  # สถานะ
                        
                        self.log(f"✅ Sheet 'Summary': {len(summary_data)} rows")
            
            self.status_bar.showMessage("Export สำเร็จ!")
            QMessageBox.information(self, "สำเร็จ", f"Export สำเร็จ!\n\n📊 Sheet 'Selected': {len(selected_df)} records\n📋 Sheet 'Summary': สรุป Target vs Actual\n\n{filepath}")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error:\n{e}")
    
    def save_settings(self):
        """Save settings - with separate Sample Size column"""
        filepath, _ = QFileDialog.getSaveFileName(self, "Save Settings", "settings.xlsx", "Excel (*.xlsx)")
        if not filepath:
            return
        
        try:
            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                for i, data in enumerate(self.quota_data):
                    rows = [{'Setting': 'Dimensions', 'Value': ','.join(data.get('dimensions', [])), 'Sample Size': ''}]
                    for key, count in data.get('targets', {}).items():
                        # Format: value1,value2,value3 in Value column, count in Sample Size column
                        key_str = ','.join(str(k) for k in key)
                        rows.append({'Setting': 'Target', 'Value': key_str, 'Sample Size': count})
                    pd.DataFrame(rows).to_excel(writer, sheet_name=f'Set_{i+1}', index=False)
            
            self.log("✅ Settings saved")
            QMessageBox.information(self, "สำเร็จ", f"Save สำเร็จ!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error:\n{e}")
    
    def load_settings(self):
        """Load settings - supports new format with Sample Size column"""
        filepath, _ = QFileDialog.getOpenFileName(self, "Load Settings", "", "Excel (*.xlsx)")
        if not filepath:
            return
        
        try:
            excel = pd.ExcelFile(filepath)
            
            for i in range(self.num_quota_sets):
                sheet = f'Set_{i+1}'
                if sheet not in excel.sheet_names:
                    continue
                
                df = pd.read_excel(excel, sheet_name=sheet)
                data = {'dimensions': [], 'targets': {}}
                
                # Check if new format (has Sample Size column)
                has_sample_size_col = 'Sample Size' in df.columns
                
                for _, row in df.iterrows():
                    setting = row['Setting']
                    value = str(row['Value'])
                    
                    if setting == 'Dimensions':
                        data['dimensions'] = [d.strip() for d in value.split(',') if d.strip()]
                    elif setting == 'Target':
                        try:
                            if has_sample_size_col and pd.notna(row.get('Sample Size')):
                                # New format: Value = conditions, Sample Size = count
                                key_parts = [k.strip() for k in value.split(',')]
                                count = int(row['Sample Size'])
                                data['targets'][tuple(key_parts)] = count
                            elif '=' in value:
                                # Old format: value1,value2,value3=count
                                parts = value.rsplit('=', 1)
                                key_parts = [k.strip() for k in parts[0].split(',')]
                                count = int(parts[1])
                                data['targets'][tuple(key_parts)] = count
                            elif ':' in value:
                                # Legacy format: tuple:count
                                parts = value.rsplit(':', 1)
                                key = ast.literal_eval(parts[0])
                                count = int(parts[1])
                                data['targets'][tuple(map(str, key))] = count
                        except:
                            pass
                
                self.quota_data[i] = data
                self.update_quota_button(i)
            
            self.log("✅ Settings loaded")
            QMessageBox.information(self, "สำเร็จ", "Load สำเร็จ!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error:\n{e}")
    
    def preview_quota(self):
        """Preview all configured quotas"""
        if self.cleaned_df is None:
            QMessageBox.warning(self, "Error", "กรุณาโหลดไฟล์ข้อมูลก่อน")
            return
        
        # Show preview dialog
        dialog = QuotaPreviewDialog(self)
        dialog.exec()


def run_this_app(working_dir=None):
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setStyleSheet(MODERN_STYLE)
    window = QuotaSamplerApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    run_this_app()
