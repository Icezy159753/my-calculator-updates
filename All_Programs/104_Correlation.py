# -*- coding: utf-8 -*-
# SPSS Correlation Tool by Gemini (GUI v4.2) - Excel Export with Index Sheet and Freeze Panes

import sys
import pandas as pd
import numpy as np
import pyreadstat
import re  # สำหรับ Natural Sort
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog,
    QMessageBox, QListWidget, QListWidgetItem, QAbstractItemView,
    QGroupBox, QInputDialog, QTextEdit,
    QDialog, QDialogButtonBox, QSplitter, QTabWidget,
    QSpacerItem, QSizePolicy, QRadioButton, QButtonGroup, QComboBox, QFrame
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QGuiApplication, QFont, QPalette, QColor

# --- ฟังก์ชันสำหรับ Natural Sort ---
def natural_sort_key(s):
    """
    สร้าง key สำหรับการเรียงลำดับแบบธรรมชาติ (natural sort)
    เพื่อให้ตัวเลขในสตริงถูกเรียงอย่างถูกต้อง (เช่น Var1, Var2, Var10 แทนที่จะเป็น Var1, Var10, Var2)
    """
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

# --- Dialog สำหรับการรวมข้อมูล 2 กลุ่ม (Merge/Stack Data) ---
class MergeDataDialog(QDialog):
    def __init__(self, all_variables, current_global_filter_desc, parent=None):
        super().__init__(parent)
        self.setWindowTitle("สร้าง Correlation จากการรวมข้อมูล (Merge and Correlate)")
        self.all_variables = all_variables
        self.current_global_filter_desc = current_global_filter_desc
        self.setMinimumSize(900, 600); self.resize(1000, 780) # เพิ่มขนาดเริ่มต้น (ของ Dialog นี้)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMaximizeButtonHint | Qt.WindowType.WindowMinimizeButtonHint) # เพิ่มปุ่ม maximize/minimize

        # ตั้งค่า UI สไตล์สำหรับ Dialog
        self.setStyleSheet(parent.styleSheet()) # ใช้ stylesheet เดียวกับหน้าต่างหลัก

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # ส่วนหัวข้อและคำอธิบาย
        main_layout.addWidget(QLabel("<b><span style='font-size:12pt; color:#007bff;'>สร้าง Correlation จากการรวมข้อมูล (Stacking)</span></b>"))
        main_layout.addWidget(QLabel("เลือกตัวแปรสำหรับแต่ละกลุ่ม โปรแกรมจะนำข้อมูลในคอลัมน์มาต่อกันตามลำดับ"))

        # GroupBox สำหรับ Filter เฉพาะ
        filter_group = QGroupBox("Filter สำหรับการ Merge นี้เท่านั้น")
        filter_layout = QVBoxLayout(filter_group)
        self.filter_query_edit = QLineEdit()
        self.filter_query_edit.setPlaceholderText("เช่น status == 1 (ถ้าเว้นว่าง ไม่ใช้ Filter)")
        filter_layout.addWidget(QLabel(f"Filter ที่ใช้อยู่: <span style='color: #28a745;'>{self.current_global_filter_desc}</span>"))
        filter_layout.addWidget(self.filter_query_edit)
        main_layout.addWidget(filter_group)

        # ส่วนเลือกตัวแปร (ซ้าย-กลาง-ขวา)
        selection_layout = QHBoxLayout()

        # กลุ่มตัวแปรทั้งหมด
        source_group = QGroupBox("ตัวแปรทั้งหมด (Available Variables)")
        source_layout = QVBoxLayout(source_group)
        self.source_list = QListWidget()
        self.source_list.addItems(self.all_variables)
        self.source_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        source_layout.addWidget(self.source_list)
        selection_layout.addWidget(source_group, 2)

        # ปุ่มย้ายตัวแปร
        buttons_layout = QVBoxLayout()
        buttons_layout.addStretch()
        add_to_group1_btn = QPushButton(">> Group 1")
        add_to_group1_btn.setObjectName("dialogAddButton") # Custom object name for styling
        add_to_group2_btn = QPushButton(">> Group 2")
        add_to_group2_btn.setObjectName("dialogAddButton")
        buttons_layout.addWidget(add_to_group1_btn)
        buttons_layout.addWidget(add_to_group2_btn)
        buttons_layout.addStretch()
        selection_layout.addLayout(buttons_layout)

        # กลุ่มตัวแปร Group 1 และ Group 2
        groups_layout = QVBoxLayout()

        group1_box = QGroupBox("Group 1 Variables")
        group1_layout = QVBoxLayout(group1_box)
        self.name1_edit = QLineEdit()
        self.name1_edit.setPlaceholderText("ตั้งชื่อคอลัมน์ใหม่สำหรับ Group 1 (เช่น Brand_A)")
        self.group1_list = QListWidget()
        self.group1_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        remove_g1_btn = QPushButton("นำออกจาก Group 1")
        remove_g1_btn.setObjectName("dialogRemoveButton") # Custom object name for styling
        group1_layout.addWidget(self.name1_edit)
        group1_layout.addWidget(self.group1_list)
        group1_layout.addWidget(remove_g1_btn)
        groups_layout.addWidget(group1_box)

        group2_box = QGroupBox("Group 2 Variables")
        group2_layout = QVBoxLayout(group2_box)
        self.name2_edit = QLineEdit()
        self.name2_edit.setPlaceholderText("ตั้งชื่อคอลัมน์ใหม่สำหรับ Group 2 (เช่น Brand_B)")
        self.group2_list = QListWidget()
        self.group2_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        remove_g2_btn = QPushButton("นำออกจาก Group 2")
        remove_g2_btn.setObjectName("dialogRemoveButton")
        group2_layout.addWidget(self.name2_edit)
        group2_layout.addWidget(self.group2_list)
        group2_layout.addWidget(remove_g2_btn)
        groups_layout.addWidget(group2_box)

        selection_layout.addLayout(groups_layout, 3)
        main_layout.addLayout(selection_layout)

        # เชื่อมสัญญาณ
        add_to_group1_btn.clicked.connect(lambda: self.move_selected_items(self.source_list, self.group1_list))
        add_to_group2_btn.clicked.connect(lambda: self.move_selected_items(self.source_list, self.group2_list))
        remove_g1_btn.clicked.connect(lambda: self.move_selected_items(self.group1_list, self.source_list))
        remove_g2_btn.clicked.connect(lambda: self.move_selected_items(self.group2_list, self.source_list))

        # ส่วนชื่อ Sheet
        sheet_name_layout = QHBoxLayout()
        sheet_name_layout.addWidget(QLabel("<b>ชื่อ Sheet สำหรับผลลัพธ์นี้:</b>"))
        self.sheet_name_edit = QLineEdit()
        self.sheet_name_edit.setPlaceholderText("เช่น Merged_BrandA_vs_BrandB")
        sheet_name_layout.addWidget(self.sheet_name_edit)
        main_layout.addLayout(sheet_name_layout)

        # ปุ่ม OK/Cancel
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        main_layout.addWidget(self.button_box)

        # เพิ่มสไตล์สำหรับปุ่มใน Dialog โดยเฉพาะ
        self.setStyleSheet(self.styleSheet() + """
            QPushButton#dialogAddButton {
                background-color: #28a745; /* Green */
                color: white;
            }
            QPushButton#dialogAddButton:hover {
                background-color: #218838;
            }
            QPushButton#dialogRemoveButton {
                background-color: #dc3545; /* Red */
                color: white;
            }
            QPushButton#dialogRemoveButton:hover {
                background-color: #c82333;
            }
            QDialogButtonBox QPushButton {
                min-width: 100px; /* Make OK/Cancel buttons wider */
            }
        """)

    def move_selected_items(self, from_list, to_list):
        """ย้ายรายการที่เลือกจาก ListWidget หนึ่งไปยังอีกลิสต์หนึ่ง"""
        for item in from_list.selectedItems():
            to_list.addItem(from_list.takeItem(from_list.row(item)))

    def accept(self):
        """ตรวจสอบข้อมูลก่อนปิด Dialog"""
        if not self.name1_edit.text().strip() or not self.name2_edit.text().strip() or not self.sheet_name_edit.text().strip():
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณาตั้งชื่อคอลัมน์ใหม่ทั้งสองกลุ่ม และตั้งชื่อ Sheet")
            return
        if self.group1_list.count() == 0 or self.group2_list.count() == 0:
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณาเลือกตัวแปรอย่างน้อยหนึ่งตัวสำหรับแต่ละกลุ่ม")
            return
        super().accept()

    def get_data(self):
        """ส่งคืนข้อมูลที่ผู้ใช้เลือกและกรอก"""
        return {
            "group1_vars": [self.group1_list.item(i).text() for i in range(self.group1_list.count())],
            "group2_vars": [self.group2_list.item(i).text() for i in range(self.group2_list.count())],
            "name1": self.name1_edit.text().strip(),
            "name2": self.name2_edit.text().strip(),
            "sheet_name": self.sheet_name_edit.text().strip(),
            "filter_query": self.filter_query_edit.text().strip()
        }

# --- หน้าต่างสำหรับสร้างตัวแปรแบบ Stack (เดี่ยว) ---
class StackVariablesDialog(QDialog):
    def __init__(self, spss_variables, meta, current_filter_desc, parent=None):
        super().__init__(parent)
        self.setWindowTitle("สร้างตัวแปรจากการรวมข้อมูล (Stacking)")
        self.setMinimumSize(900, 600); self.resize(1000, 780) # เพิ่มขนาดเริ่มต้น (ของ Dialog นี้)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMaximizeButtonHint | Qt.WindowType.WindowMinimizeButtonHint)
        self.meta = meta

        # ตั้งค่า UI สไตล์สำหรับ Dialog
        self.setStyleSheet(parent.styleSheet())

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        main_layout.addWidget(QLabel("<b><span style='font-size:12pt; color:#007bff;'>เลือกตัวแปรจาก SPSS เพื่อนำข้อมูลมาต่อกัน (Stack)</span></b>"))
        main_layout.addWidget(QLabel(f"<b>Filter ที่จะใช้:</b> <span style='color: #28a745;'>{current_filter_desc}</span>"))

        selection_layout = QHBoxLayout()

        # กลุ่มตัวแปรทั้งหมด
        source_group = QGroupBox("ตัวแปรทั้งหมดจาก SPSS")
        source_v_layout = QVBoxLayout(source_group)
        self.source_list = QListWidget()
        self.source_list.addItems(spss_variables)
        self.source_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        source_v_layout.addWidget(self.source_list)
        selection_layout.addWidget(source_group)

        # ปุ่มย้าย
        button_layout = QVBoxLayout()
        button_layout.addStretch()
        self.add_button = QPushButton("เพิ่ม  >>")
        self.add_button.setObjectName("dialogAddButton")
        self.add_button.setFixedWidth(125)
        self.add_button.setToolTip("ย้ายตัวแปรที่เลือกไปทางขวา (หรือดับเบิลคลิกที่ตัวแปร)")
        self.remove_button = QPushButton("<<  นำออก")
        self.remove_button.setObjectName("dialogRemoveButton")
        self.remove_button.setFixedWidth(125)
        self.remove_button.setToolTip("ย้ายตัวแปรที่เลือกกลับไปทางซ้าย (หรือดับเบิลคลิกที่ตัวแปร)")
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.remove_button)
        button_layout.addStretch()
        selection_layout.addLayout(button_layout)

        # กลุ่มตัวแปรที่จะ Stack
        stack_group = QGroupBox("ตัวแปรที่จะนำไป Stack")
        stack_v_layout = QVBoxLayout(stack_group)
        self.stack_list = QListWidget()
        self.stack_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        stack_v_layout.addWidget(self.stack_list)
        selection_layout.addWidget(stack_group)
        main_layout.addLayout(selection_layout)

        # กลุ่มการตั้งชื่อและโจทย์
        naming_group = QGroupBox("ตั้งชื่อและโจทย์")
        naming_layout = QVBoxLayout(naming_group)

        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("<b>ตั้งชื่อตัวแปรใหม่:</b>"))
        self.new_var_name_edit = QLineEdit()
        self.new_var_name_edit.setPlaceholderText("เช่น Merged_BL_All")
        name_layout.addWidget(self.new_var_name_edit)
        naming_layout.addLayout(name_layout)

        label_layout = QHBoxLayout()
        label_layout.addWidget(QLabel("<b>โจทย์ (Label):</b>"))
        self.new_var_label_edit = QLineEdit()
        self.new_var_label_edit.setPlaceholderText("ดึงมาจาก Label ของตัวแปรแรกที่เลือก")
        label_layout.addWidget(self.new_var_label_edit)
        naming_layout.addLayout(label_layout)

        main_layout.addWidget(naming_group)

        # ปุ่ม OK/Cancel
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # เชื่อมสัญญาณ
        self.add_button.clicked.connect(lambda: self.move_items(self.source_list, self.stack_list))
        self.remove_button.clicked.connect(lambda: self.move_items(self.stack_list, self.source_list))
        self.source_list.itemDoubleClicked.connect(lambda _: self.move_items(self.source_list, self.stack_list))
        self.stack_list.itemDoubleClicked.connect(lambda _: self.move_items(self.stack_list, self.source_list))
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # เพิ่มสไตล์สำหรับปุ่มใน Dialog โดยเฉพาะ
        self.setStyleSheet(self.styleSheet() + """
            QPushButton#dialogAddButton {
                background-color: #28a745; /* Green */
                color: white;
            }
            QPushButton#dialogAddButton:hover {
                background-color: #218838;
            }
            QPushButton#dialogRemoveButton {
                background-color: #dc3545; /* Red */
                color: white;
            }
            QPushButton#dialogRemoveButton:hover {
                background-color: #c82333;
            }
            QDialogButtonBox QPushButton {
                min-width: 100px;
            }
        """)

    def move_items(self, from_list, to_list):
        """ย้ายรายการที่เลือกจาก ListWidget หนึ่งไปยังอีกลิสต์หนึ่ง และตั้งค่า Label อัตโนมัติ"""
        was_empty = to_list.count() == 0
        selected_items = from_list.selectedItems()
        if not selected_items:
            return

        for item in selected_items:
            to_list.addItem(from_list.takeItem(from_list.row(item)))

        if was_empty and to_list == self.stack_list:
            # ถ้าเป็นครั้งแรกที่ย้ายเข้า stack_list ให้ดึง label ของตัวแปรแรกมาใส่
            first_var_name = to_list.item(0).text()
            if self.meta and self.meta.column_names_to_labels:
                self.new_var_label_edit.setText(self.meta.column_names_to_labels.get(first_var_name, first_var_name))
            else:
                self.new_var_label_edit.setText(first_var_name)

    def accept(self):
        """ตรวจสอบข้อมูลก่อนปิด Dialog"""
        if self.stack_list.count() < 2:
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณาเลือกตัวแปรที่จะ Stack อย่างน้อย 2 ตัว")
            return
        if not self.new_var_name_edit.text().strip():
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณาตั้งชื่อตัวแปรใหม่")
            return
        if not self.new_var_label_edit.text().strip():
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณาใส่โจทย์ (Label) สำหรับตัวแปรใหม่")
            return
        super().accept()

    def get_data(self):
        """ส่งคืนข้อมูลที่ผู้ใช้เลือกและกรอก"""
        return {
            "new_name": self.new_var_name_edit.text().strip(),
            "label": self.new_var_label_edit.text().strip(),
            "source_vars": [self.stack_list.item(i).text() for i in range(self.stack_list.count())]
        }

# --- หน้าต่างสำหรับสร้าง Stack แบบกลุ่ม ---
class PatternStackDialog(QDialog):
    def __init__(self, spss_variables, meta, parent=None):
        super().__init__(parent)
        self.setWindowTitle("สร้างตัวแปรแบบกลุ่มตามแพทเทิร์น (Pattern Stack)")
        self.setMinimumSize(900, 600); self.resize(1000, 780) # เพิ่มขนาดเริ่มต้น (ของ Dialog นี้)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMaximizeButtonHint | Qt.WindowType.WindowMinimizeButtonHint)
        self.meta = meta
        self.final_grouping = {}
        self.base_name = ""

        # ตั้งค่า UI สไตล์สำหรับ Dialog
        self.setStyleSheet(parent.styleSheet())

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        main_layout.addWidget(QLabel("<b><span style='font-size:12pt; color:#007bff;'>สร้างตัวแปรแบบกลุ่มตามแพทเทิร์น (Pattern Stack)</span></b>"))
        main_layout.addWidget(QLabel("<b>1. เลือกตัวแปรทั้งหมดที่ต้องการจัดกลุ่ม</b>"))

        selection_layout = QHBoxLayout()

        # กลุ่มตัวแปรทั้งหมด
        source_group = QGroupBox("ตัวแปรทั้งหมดจาก SPSS")
        source_v_layout = QVBoxLayout(source_group)
        self.source_list = QListWidget()
        self.source_list.addItems(spss_variables)
        self.source_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.source_list.setMinimumHeight(180)
        source_v_layout.addWidget(self.source_list)
        selection_layout.addWidget(source_group)

        # ปุ่มย้าย
        button_layout = QVBoxLayout()
        button_layout.addStretch()
        self.add_button = QPushButton("เพิ่ม  >>")
        self.add_button.setObjectName("dialogAddButton")
        self.add_button.setFixedWidth(125)
        self.add_button.setToolTip("ย้ายตัวแปรที่เลือกไปทางขวา (หรือดับเบิลคลิกที่ตัวแปร)")
        self.remove_button = QPushButton("<<  นำออก")
        self.remove_button.setObjectName("dialogRemoveButton")
        self.remove_button.setFixedWidth(125)
        self.remove_button.setToolTip("ย้ายตัวแปรที่เลือกกลับไปทางซ้าย (หรือดับเบิลคลิกที่ตัวแปร)")
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.remove_button)
        button_layout.addStretch()
        selection_layout.addLayout(button_layout)

        # กลุ่มตัวแปรที่เลือกเพื่อจัดกลุ่ม
        stack_group = QGroupBox("ตัวแปรที่เลือกเพื่อจัดกลุ่ม")
        stack_v_layout = QVBoxLayout(stack_group)
        self.selected_list = QListWidget()
        self.selected_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.selected_list.setMinimumHeight(180)
        stack_v_layout.addWidget(self.selected_list)
        selection_layout.addWidget(stack_group)
        main_layout.addLayout(selection_layout)

        # กลุ่มการตั้งค่าแพทเทิร์นและแสดงตัวอย่าง
        pattern_group = QGroupBox("2. ตั้งค่าแพทเทิร์นและแสดงตัวอย่าง")
        pattern_layout = QVBoxLayout(pattern_group)

        base_name_layout = QHBoxLayout()
        base_name_layout.addWidget(QLabel("ชื่อพื้นฐานสำหรับตัวแปรใหม่:"))
        self.base_name_edit = QLineEdit()
        self.base_name_edit.setPlaceholderText("เช่น NBL4 (โปรแกรมจะสร้างเป็น NBL4_1, NBL4_2, ...)")
        base_name_layout.addWidget(self.base_name_edit)
        pattern_layout.addLayout(base_name_layout)

        # เพิ่มตัวเลือกรูปแบบการจัดกลุ่ม
        grouping_mode_layout = QVBoxLayout()
        grouping_mode_layout.addWidget(QLabel("<b>รูปแบบการจัดกลุ่ม:</b>"))

        self.grouping_button_group = QButtonGroup(self)

        self.group_by_full_suffix = QRadioButton("จัดกลุ่มตามส่วนท้ายทั้งหมด (เช่น NA8#1_1, NA8#2_1, ... → กลุ่ม '_1')")
        details_font = QFont("Tahoma", 10)
        self.group_by_full_suffix.setFont(details_font)
        self.group_by_full_suffix.setChecked(True)  # ค่าเริ่มต้น
        self.grouping_button_group.addButton(self.group_by_full_suffix, 0)
        grouping_mode_layout.addWidget(self.group_by_full_suffix)

        self.group_by_first_num = QRadioButton("จัดกลุ่มตามส่วนหน้า (เช่น NA8#1_1, NA8#1_2, ... → กลุ่ม '#1_')")
        self.group_by_first_num.setFont(details_font)
        self.grouping_button_group.addButton(self.group_by_first_num, 1)
        grouping_mode_layout.addWidget(self.group_by_first_num)

        self.group_by_custom = QRadioButton("กำหนดเอง (ระบุสัญลักษณ์แยกกลุ่ม)")
        self.group_by_custom.setFont(details_font)
        self.grouping_button_group.addButton(self.group_by_custom, 2)
        grouping_mode_layout.addWidget(self.group_by_custom)

        custom_inline_layout = QHBoxLayout()
        custom_inline_layout.addSpacing(18)
        custom_inline_layout.addWidget(QLabel("สัญลักษณ์:"))
        self.delimiter_edit = QLineEdit("#")
        self.delimiter_edit.setMaximumWidth(80)
        self.delimiter_edit.setPlaceholderText("เช่น _")
        custom_inline_layout.addWidget(self.delimiter_edit)
        custom_inline_layout.addWidget(QLabel("แยกส่วนที่:"))
        self.split_index_edit = QLineEdit("1")
        self.split_index_edit.setMaximumWidth(80)
        self.split_index_edit.setPlaceholderText("เช่น 1 หรือ -1")
        custom_inline_layout.addWidget(self.split_index_edit)
        custom_inline_layout.addStretch()
        grouping_mode_layout.addLayout(custom_inline_layout)

        self.group_by_custom.toggled.connect(self._toggle_custom_inputs)
        self._toggle_custom_inputs(False)

        pattern_layout.addLayout(grouping_mode_layout)

        self.preview_button = QPushButton("จัดกลุ่มและแสดงตัวอย่าง")
        self.preview_button.setStyleSheet("background-color: #007bff; color: white;") # Blue button
        self.preview_button.clicked.connect(self.preview_grouping)
        pattern_layout.addWidget(self.preview_button)

        pattern_layout.addWidget(QLabel("<b>3. ตรวจสอบผลการจัดกลุ่ม:</b> คลิกปุ่มด้านบนเพื่อดูตัวอย่างในหน้าต่างใหม่"))
        main_layout.addWidget(pattern_group)

        # ปุ่ม OK/Cancel
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # เชื่อมสัญญาณ
        self.add_button.clicked.connect(lambda: self.move_items(self.source_list, self.selected_list))
        self.remove_button.clicked.connect(lambda: self.move_items(self.selected_list, self.source_list))
        self.source_list.itemDoubleClicked.connect(lambda _: self.move_items(self.source_list, self.selected_list))
        self.selected_list.itemDoubleClicked.connect(lambda _: self.move_items(self.selected_list, self.source_list))
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # เพิ่มสไตล์สำหรับปุ่มใน Dialog โดยเฉพาะ
        self.setStyleSheet(self.styleSheet() + """
            QPushButton#dialogAddButton {
                background-color: #28a745; /* Green */
                color: white;
            }
            QPushButton#dialogAddButton:hover {
                background-color: #218838;
            }
            QPushButton#dialogRemoveButton {
                background-color: #dc3545; /* Red */
                color: white;
            }
            QPushButton#dialogRemoveButton:hover {
                background-color: #c82333;
            }
            QDialogButtonBox QPushButton {
                min-width: 100px;
            }
        """)

    def move_items(self, from_list, to_list):
        """ย้ายรายการที่เลือกจาก ListWidget หนึ่งไปยังอีกลิสต์หนึ่ง"""
        for item in from_list.selectedItems():
            to_list.addItem(from_list.takeItem(from_list.row(item)))

    def _toggle_custom_inputs(self, enabled):
        """เปิด/ปิดช่องกำหนดเองให้ใช้งานได้เฉพาะตอนเลือกโหมดกำหนดเอง"""
        self.delimiter_edit.setEnabled(enabled)
        self.split_index_edit.setEnabled(enabled)

    def preview_grouping(self):
        """แสดงตัวอย่างการจัดกลุ่มตัวแปรตามแพทเทิร์น"""
        self.base_name = self.base_name_edit.text().strip()
        if not self.base_name:
            QMessageBox.warning(self, "ข้อมูลไม่ครบ", "กรุณากรอก 'ชื่อพื้นฐาน' สำหรับตัวแปรใหม่")
            return

        selected_vars = [self.selected_list.item(i).text() for i in range(self.selected_list.count())]
        if not selected_vars:
            QMessageBox.warning(self, "ไม่มีตัวแปร", "กรุณาเลือกตัวแปรที่จะจัดกลุ่มก่อน")
            return

        grouped_vars = {}

        # ตรวจสอบว่าผู้ใช้เลือกรูปแบบการจัดกลุ่มแบบไหน
        if self.group_by_custom.isChecked():
            delimiter = self.delimiter_edit.text().strip()
            if not delimiter:
                QMessageBox.warning(self, "ข้อมูลไม่ครบ", "กรุณากรอกสัญลักษณ์สำหรับการแยกกลุ่ม")
                return

            try:
                split_index = int(self.split_index_edit.text().strip())
            except ValueError:
                QMessageBox.warning(self, "ข้อมูลไม่ถูกต้อง", "กรุณากรอกตัวเลขสำหรับ 'แยกส่วนที่'")
                return

            for var in selected_vars:
                if delimiter in var:
                    parts = var.split(delimiter)
                    try:
                        pattern_suffix = parts[split_index]
                        grouped_vars.setdefault(pattern_suffix, []).append(var)
                    except IndexError:
                        continue

            if not grouped_vars:
                QMessageBox.critical(
                    self,
                    "ไม่พบแพทเทิร์น",
                    f"ไม่สามารถแยกกลุ่มได้ด้วยสัญลักษณ์ '{delimiter}' และส่วนที่ {split_index}\n"
                    "โปรดตรวจสอบว่าตัวแปรมีสัญลักษณ์นี้และมีส่วนที่ระบุ"
                )
                self.final_grouping = {}
                return

        elif self.group_by_first_num.isChecked():
            # จัดกลุ่มตามตัวเลขหน้า (เช่น #1_, #2_, #3_)
            for var in selected_vars:
                # ใช้ regex เพื่อแยกส่วนตัวเลขหน้าหลัง '#'
                match = re.search(r'#(\d+)_', var)
                if match:
                    pattern_suffix = '#' + match.group(1) + '_'
                    grouped_vars.setdefault(pattern_suffix, []).append(var)

            if not grouped_vars:
                QMessageBox.critical(self, "ไม่พบแพทเทิร์น", "ไม่สามารถหาแพทเทิร์น '#ตัวเลข_' ในตัวแปรที่เลือกได้")
                self.final_grouping = {}
                return

        else:
            # จัดกลุ่มตามส่วนท้ายทั้งหมด (เช่น _1, _2, _3) - รูปแบบเดิม
            for var in selected_vars:
                # ใช้ regex เพื่อแยกส่วนของสตริงที่ตามหลัง '#'
                match = re.search(r'#(.*)', var)
                if match:
                    pattern_suffix = '#' + match.group(1)
                    grouped_vars.setdefault(pattern_suffix, []).append(var)

            if not grouped_vars:
                QMessageBox.critical(self, "ไม่พบแพทเทิร์น", "ไม่สามารถหาแพทเทิร์น '#' ในตัวแปรที่เลือกได้")
                self.final_grouping = {}
                return

        # เรียงลำดับตัวแปรภายในแต่ละกลุ่ม
        for key in grouped_vars:
            grouped_vars[key].sort(key=natural_sort_key)

        self.final_grouping = grouped_vars
        preview_content = ""
        # เรียงลำดับกลุ่มตาม key ด้วย natural sort
        for item_index in sorted(grouped_vars.keys(), key=natural_sort_key):
            new_var_name = f"{self.base_name}_{item_index.replace('#', '')}"
            vars_to_stack = grouped_vars[item_index]
            
            preview_content += f"จะสร้างตัวแปรใหม่: {new_var_name}\n"
            preview_content += f"  (Label จะดึงจาก: {self.meta.column_names_to_labels.get(vars_to_stack[0], vars_to_stack[0])})\n"
            preview_content += f"  - จากการรวม:\n"
            for v in vars_to_stack:
                preview_content += f"    - {v}\n"
            preview_content += "-" * 50 + "\n"
        self._show_preview_dialog(preview_content)

    def _show_preview_dialog(self, preview_content):
        """เปิดหน้าต่างใหม่เพื่อแสดงตัวอย่างการจัดกลุ่ม"""
        dialog = QDialog(self)
        dialog.setWindowTitle("ตัวอย่างการจัดกลุ่ม (Pattern Stack Preview)")
        dialog.setMinimumSize(700, 600)
        dialog.setStyleSheet(self.styleSheet())

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        preview_text = QTextEdit()
        preview_text.setReadOnly(True)
        preview_text.setFont(QFont("Courier New", 10))
        preview_text.setText(preview_content)
        layout.addWidget(preview_text)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        button_box.rejected.connect(dialog.reject)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)

        dialog.exec()

    def accept(self):
        """ตรวจสอบข้อมูลก่อนปิด Dialog"""
        if not self.final_grouping:
            QMessageBox.warning(self, "ยังไม่ได้จัดกลุ่ม", "กรุณากดปุ่ม 'จัดกลุ่มและแสดงตัวอย่าง' ก่อนกด OK")
            return
        super().accept()

    def get_data(self):
        """ส่งคืนข้อมูลการจัดกลุ่มที่ผู้ใช้ตั้งค่า"""
        return {"groups": self.final_grouping, "base_name": self.base_name}

# --- หน้าต่างสำหรับ "ตั้งค่า" ตัวแปรเพื่อเพิ่มลงคิว ---

# --- หน้าต่างสำหรับสร้างตัวแปร Recode Justright ---
RECODE_JR_MAPS = {
    5: {1: 1, 5: 1, 2: 2, 4: 2, 3: 3},
    7: {1: 1, 7: 1, 2: 2, 6: 2, 3: 3, 5: 3, 4: 4},
}

def detect_justright_scale(series):
    """ตรวจจับว่าเป็น 5 หรือ 7 Scale จากค่าสูงสุดของข้อมูล"""
    try:
        max_val = pd.to_numeric(series, errors='coerce').max()
        if pd.isna(max_val):
            return 5
        return 7 if max_val > 5 else 5
    except Exception:
        return 5

def recode_justright_series(series, scale):
    """Recode ค่าตาม Justright: 5 Scale -> 1,5=1 / 2,4=2 / 3=3 ; 7 Scale -> 1,7=1 / 2,6=2 / 3,5=3 / 4=4"""
    numeric = pd.to_numeric(series, errors='coerce')
    return numeric.map(RECODE_JR_MAPS.get(int(scale), RECODE_JR_MAPS[5]))

class RecodeJustrightDialog(QDialog):
    def __init__(self, spss_variables, meta, df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("สร้างตัวแปร Recode Justright")
        self.setMinimumSize(900, 600); self.resize(1000, 780)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMaximizeButtonHint | Qt.WindowType.WindowMinimizeButtonHint)
        self.meta = meta
        self.df = df
        self.setStyleSheet(parent.styleSheet())

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        main_layout.addWidget(QLabel("<b><span style='font-size:12pt; color:#007bff;'>เลือกตัวแปรที่จะสร้างเป็นตัวแปรใหม่แบบ Recode Justright</span></b>"))
        hint = QLabel(
            "<table cellspacing='0' cellpadding='2'>"
            "<tr><td><b>5 Scale</b></td><td>&nbsp;&nbsp;1,5 → 1 &nbsp;|&nbsp; 2,4 → 2 &nbsp;|&nbsp; 3 → 3</td></tr>"
            "<tr><td><b>7 Scale</b></td><td>&nbsp;&nbsp;1,7 → 1 &nbsp;|&nbsp; 2,6 → 2 &nbsp;|&nbsp; 3,5 → 3 &nbsp;|&nbsp; 4 → 4</td></tr>"
            "</table>"
        )
        hint.setTextFormat(Qt.TextFormat.RichText)
        hint.setStyleSheet("background-color: #ffffff; border: 1px solid #d3d3d3; border-radius: 8px; padding: 8px 12px; color: #2c3e50;")
        main_layout.addWidget(hint)
        note = QLabel("ตัวแปรใหม่จะถูกสร้างแยกจากตัวแปรเดิม (ไม่ต่อ Data) และคำนวณตอน Export — ค่าที่อยู่นอกช่วง Scale จะกลายเป็นค่าว่าง")
        note.setWordWrap(True)
        note.setStyleSheet("color: #6c757d;")
        main_layout.addWidget(note)

        selection_layout = QHBoxLayout()
        source_group = QGroupBox("ตัวแปรทั้งหมดจาก SPSS")
        source_v_layout = QVBoxLayout(source_group)
        self.source_list = QListWidget()
        self.source_list.addItems(spss_variables)
        self.source_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        source_v_layout.addWidget(self.source_list)
        selection_layout.addWidget(source_group)

        button_layout = QVBoxLayout()
        button_layout.addStretch()
        self.add_button = QPushButton("เพิ่ม  >>")
        self.add_button.setObjectName("dialogAddButton")
        self.add_button.setFixedWidth(125)
        self.add_button.setToolTip("ย้ายตัวแปรที่เลือกไปทางขวา (หรือดับเบิลคลิกที่ตัวแปร)")
        self.remove_button = QPushButton("<<  นำออก")
        self.remove_button.setObjectName("dialogRemoveButton")
        self.remove_button.setFixedWidth(125)
        self.remove_button.setToolTip("ย้ายตัวแปรที่เลือกกลับไปทางซ้าย (หรือดับเบิลคลิกที่ตัวแปร)")
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.remove_button)
        button_layout.addStretch()
        selection_layout.addLayout(button_layout)

        self.target_group = QGroupBox("ตัวแปรที่จะ Recode Justright")
        target_group = self.target_group
        target_v_layout = QVBoxLayout(target_group)
        self.target_list = QListWidget()
        self.target_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        target_v_layout.addWidget(self.target_list)
        selection_layout.addWidget(target_group)
        main_layout.addLayout(selection_layout)

        option_group = QGroupBox("ตั้งค่าการสร้างตัวแปรใหม่")
        option_layout = QVBoxLayout(option_group)

        row = QHBoxLayout()
        row.setSpacing(10)
        row.addWidget(QLabel("<b>ต่อท้ายชื่อตัวแปรใหม่:</b>"))
        self.suffix_edit = QLineEdit("_JR")
        self.suffix_edit.setPlaceholderText("เช่น _JR")
        self.suffix_edit.setMaximumWidth(160)
        row.addWidget(self.suffix_edit)
        row.addSpacing(20)
        row.addWidget(QLabel("<b>Scale:</b>"))
        self.scale_combo = QComboBox()
        self.scale_combo.addItems(["5 Scale", "7 Scale"])
        self.scale_combo.setMinimumWidth(220)
        row.addWidget(self.scale_combo)
        row.addStretch()
        option_layout.addLayout(row)

        self.example_label = QLabel()
        self.example_label.setStyleSheet("color: #6c757d; font-style: italic;")
        option_layout.addWidget(self.example_label)
        self.suffix_edit.textChanged.connect(self._update_example)
        self.scale_combo.currentIndexChanged.connect(self._update_example)
        self._update_example()
        main_layout.addWidget(option_group)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        self.add_button.clicked.connect(lambda: self.move_items(self.source_list, self.target_list))
        self.remove_button.clicked.connect(lambda: self.move_items(self.target_list, self.source_list))
        self.source_list.itemDoubleClicked.connect(lambda _: self.move_items(self.source_list, self.target_list))
        self.target_list.itemDoubleClicked.connect(lambda _: self.move_items(self.target_list, self.source_list))
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setText("สร้างตัวแปร")
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText("ยกเลิก")
        self._update_target_title()

        self.setStyleSheet(self.styleSheet() + """
            QPushButton#dialogAddButton { background-color: #28a745; color: white; }
            QPushButton#dialogAddButton:hover { background-color: #218838; }
            QPushButton#dialogRemoveButton { background-color: #dc3545; color: white; }
            QPushButton#dialogRemoveButton:hover { background-color: #c82333; }
            QDialogButtonBox QPushButton { min-width: 100px; }
        """)

    def move_items(self, from_list, to_list):
        for item in from_list.selectedItems():
            to_list.addItem(from_list.takeItem(from_list.row(item)))
        self._update_target_title()

    def _update_example(self):
        suffix = self.suffix_edit.text().strip() or "_JR"
        scale = 7 if self.scale_combo.currentIndex() == 1 else 5
        self.example_label.setText(f"ตัวอย่าง: Q1  →  Q1{suffix}   (Label: <Label เดิม> (Justright {scale} Scale))")

    def _update_target_title(self):
        n = self.target_list.count()
        self.target_group.setTitle(f"ตัวแปรที่จะ Recode Justright ({n} ตัว)" if n else "ตัวแปรที่จะ Recode Justright")

    def accept(self):
        if self.target_list.count() < 1:
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณาเลือกตัวแปรที่จะ Recode อย่างน้อย 1 ตัว")
            return
        if not self.suffix_edit.text().strip():
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณาระบุคำต่อท้ายชื่อตัวแปรใหม่")
            return
        super().accept()

    def get_data(self):
        """ส่งคืนรายการ (ตัวแปรเดิม, ตัวแปรใหม่, scale)"""
        suffix = self.suffix_edit.text().strip()
        idx = self.scale_combo.currentIndex()
        items = []
        for i in range(self.target_list.count()):
            src = self.target_list.item(i).text()
            scale = 7 if idx == 1 else 5
            items.append({"source": src, "new_name": f"{src}{suffix}", "scale": scale})
        return items

class SetQueueVarsDialog(QDialog):
    def __init__(self, all_vars, pre_selected_vars, parent=None):
        super().__init__(parent)
        self.all_vars = all_vars # เก็บไว้สำหรับเรียงลำดับ
        self.setWindowTitle("ตั้งค่าตัวแปรสำหรับเพิ่มลงคิว")
        self.setMinimumSize(900, 600); self.resize(1000, 780) # เพิ่มขนาดเริ่มต้น (ของ Dialog นี้)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMaximizeButtonHint | Qt.WindowType.WindowMinimizeButtonHint)

        # ตั้งค่า UI สไตล์สำหรับ Dialog
        self.setStyleSheet(parent.styleSheet())

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        main_layout.addWidget(QLabel("<b><span style='font-size:12pt; color:#007bff;'>เลือกตัวแปรสำหรับตั้งค่าที่จะเพิ่มลงคิว</span></b>"))

        selection_layout = QHBoxLayout()

        # ตัวแปรที่ใช้งานได้ทั้งหมด
        available_group = QGroupBox("Available Variables")
        available_v_layout = QVBoxLayout(available_group)
        self.available_list = QListWidget()
        self.available_list.addItems(all_vars) # ใส่ all_vars ทั้งหมดใน available_list ตั้งแต่แรก
        self.available_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        available_v_layout.addWidget(self.available_list)
        selection_layout.addWidget(available_group)

        # ปุ่มย้าย
        button_layout = QVBoxLayout()
        button_layout.addStretch()
        self.add_button = QPushButton("เพิ่ม  >>")
        self.add_button.setObjectName("dialogAddButton")
        self.add_button.setFixedWidth(125)
        self.add_button.setToolTip("ย้ายตัวแปรที่เลือกไปทางขวา (หรือดับเบิลคลิกที่ตัวแปร)")
        self.remove_button = QPushButton("<<  นำออก")
        self.remove_button.setObjectName("dialogRemoveButton")
        self.remove_button.setFixedWidth(125)
        self.remove_button.setToolTip("ย้ายตัวแปรที่เลือกกลับไปทางซ้าย (หรือดับเบิลคลิกที่ตัวแปร)")
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.remove_button)
        button_layout.addStretch()
        selection_layout.addLayout(button_layout)

        # ตัวแปรที่เลือกแล้ว
        selected_group = QGroupBox("Selected Variables")
        selected_v_layout = QVBoxLayout(selected_group)
        self.selected_list = QListWidget()
        self.selected_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        selected_v_layout.addWidget(self.selected_list)
        selection_layout.addWidget(selected_group)
        main_layout.addLayout(selection_layout)

        # เคลื่อนย้าย pre_selected_vars จาก available ไปยัง selected
        # ทำหลังจากที่ available_list ถูก populate แล้ว
        for var in pre_selected_vars:
            items = self.available_list.findItems(var, Qt.MatchFlag.MatchExactly)
            if items:
                row = self.available_list.row(items[0])
                self.selected_list.addItem(self.available_list.takeItem(row))

        # ปุ่ม OK/Cancel
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # เชื่อมสัญญาณ
        self.add_button.clicked.connect(lambda: self.move_items(self.available_list, self.selected_list, sort_destination=False))
        self.remove_button.clicked.connect(lambda: self.move_items(self.selected_list, self.available_list, sort_destination=True))
        self.available_list.itemDoubleClicked.connect(lambda _: self.move_items(self.available_list, self.selected_list, sort_destination=False))
        self.selected_list.itemDoubleClicked.connect(lambda _: self.move_items(self.selected_list, self.available_list, sort_destination=True))
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # เพิ่มสไตล์สำหรับปุ่มใน Dialog โดยเฉพาะ
        self.setStyleSheet(self.styleSheet() + """
            QPushButton#dialogAddButton {
                background-color: #28a745; /* Green */
                color: white;
            }
            QPushButton#dialogAddButton:hover {
                background-color: #218838;
            }
            QPushButton#dialogRemoveButton {
                background-color: #dc3545; /* Red */
                color: white;
            }
            QPushButton#dialogRemoveButton:hover {
                background-color: #c82333;
            }
            QDialogButtonBox QPushButton {
                min-width: 100px;
            }
        """)

    def move_items(self, from_list, to_list, sort_destination=True):
        """ย้ายรายการที่เลือกจาก ListWidget หนึ่งไปยังอีกลิสต์หนึ่ง พร้อมตัวเลือกในการเรียงลำดับปลายทาง"""
        for item in from_list.selectedItems():
            to_list.addItem(from_list.takeItem(from_list.row(item)))
        if sort_destination:
            items_text = [to_list.item(i).text() for i in range(to_list.count())]
            # เรียงลำดับตามลำดับเดิมของ all_vars
            sorted_items_text = sorted(items_text, key=lambda x: self.all_vars.index(x) if x in self.all_vars else float('inf'))
            to_list.clear()
            to_list.addItems(sorted_items_text)

    def get_selected_vars(self):
        """ส่งคืนรายการตัวแปรที่เลือก"""
        return [self.selected_list.item(i).text() for i in range(self.selected_list.count())]

# --- หน้าต่างหลักของโปรแกรม ---
class SPSSCorrelationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("โปรแกรมรัน Correlation จาก SPSS Data V1")
        self.df = None
        self.meta = None
        self.template_setups_for_queue = []
        self.virtual_variables = {}
        self.current_queue_setup_vars = []
        self.active_filter_query = ""
        self.active_filter_description = "ไม่มี" # เริ่มต้นด้วย "ไม่มี"

        # กำหนดขนาดต่ำสุดของหน้าต่างหลัก เพื่อให้สามารถปรับขนาดได้ แต่ไม่เล็กเกินไป
        self.setMinimumSize(QSize(700, 550)) 

        # ตั้งค่า Font ทั่วไป
        default_font = QFont("Tahoma", 9)
        self.setFont(default_font)

        # ตั้งค่า StyleSheet ทั่วทั้งแอปพลิเคชัน
        self.setStyleSheet("""
            QMainWindow, QDialog {
                background-color: #eef1f5; /* พื้นหลังสีเทาอ่อน (ใช้กับ Dialog ด้วย) */
            }
            QFrame#app_header {
                background-color: #1f3b5b;
                border-radius: 12px;
            }
            QLabel#app_title { color: #ffffff; font-size: 15pt; font-weight: bold; }
            QLabel#app_subtitle { color: #b9c9dc; font-size: 9pt; }
            QLabel#file_chip {
                background-color: rgba(255,255,255,0.12);
                color: #ffffff; border-radius: 12px; padding: 5px 12px; font-size: 9pt;
            }
            QLabel#file_chip[loaded="true"] {
                background-color: #28a745; color: #ffffff;
            }
            QLabel#section_hint { color: #6c757d; font-size: 9pt; }
            QLabel#field_label { color: #495057; font-weight: bold; font-size: 9pt; margin-top: 2px; }
            QLabel#queue_empty_hint { color: #9aa3ad; font-style: italic; }
            QListWidget#queue_list::item {
                padding: 0px; margin: 3px 2px; border-radius: 10px;
                border: 1px solid #e3e8ee; background-color: #ffffff;
            }
            QListWidget#queue_list::item:hover { border-color: #b7c2cf; background-color: #fbfcfe; }
            QListWidget#queue_list::item:selected {
                border: 2px solid #6f42c1; background-color: #f6f1fd;
            }
            QWidget#queue_card { background: transparent; }
            QLabel#queue_badge {
                background-color: #1f3b5b; color: #ffffff; border-radius: 15px;
                font-weight: bold; font-size: 10pt;
            }
            QLabel#queue_title { font-size: 10.5pt; font-weight: bold; color: #1f2d3d; }
            QLabel#queue_subtitle { color: #6c757d; font-size: 9pt; }
            QLabel#queue_chip {
                border-radius: 9px; padding: 2px 9px; font-size: 8.5pt; font-weight: 600;
            }
            QLabel#queue_chip[kind="vars"] { background-color: #e7f1ff; color: #0b3d91; }
            QLabel#queue_chip[kind="virtual"] { background-color: #e9faf5; color: #0f8a6a; }
            QLabel#queue_chip[kind="filter"] { background-color: #fff4e5; color: #b35c00; }
            QToolTip {
                background-color: #2c3e50;
                color: #ffffff;
                border: none;
                padding: 6px 8px;
                font-size: 9pt;
            }
            QComboBox {
                border: 1px solid #ced4da;
                border-radius: 5px;
                padding: 6px 10px;
                background-color: #ffffff;
                min-height: 24px;
            }
            QComboBox:hover { border-color: #007bff; }
            QComboBox::drop-down { border: none; width: 24px; }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                border: 1px solid #ced4da;
                selection-background-color: #e7f1ff;
                selection-color: #000000;
                padding: 4px;
            }
            QRadioButton { spacing: 6px; }
            QRadioButton::indicator { width: 16px; height: 16px; }
            QWidget {
                font-family: 'Tahoma';
                font-size: 9.5pt;
                color: #333333;
            }
            QGroupBox {
                background-color: #ffffff; /* พื้นหลังสีขาวสำหรับ Group Box */
                border: 1px solid #dde3ea; /* เส้นขอบสีเทาอ่อน */
                border-radius: 12px; /* มุมโค้ง */
                margin-top: 1.4em; /* เว้นที่ว่างสำหรับ Title */
                padding: 16px; /* Padding ภายใน */
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 3px 12px;
                margin-left: 12px;
                color: #ffffff;
                background-color: #1f3b5b; /* ป้ายหัวข้อสีน้ำเงินเข้ม */
                border-radius: 9px;
                font-size: 10pt;
                font-weight: bold;
            }
            QPushButton {
                background-color: #007bff; /* สีน้ำเงินสำหรับปุ่มหลัก */
                color: white;
                border: none;
                border-radius: 8px;
                padding: 7px 12px;
                font-size: 9pt;
                font-weight: 600;
                min-height: 28px;
            }
            QPushButton:hover {
                background-color: #0056b3; /* สีน้ำเงินเข้มขึ้นเมื่อ hover */
            }
            QPushButton:pressed { padding-top: 8px; padding-bottom: 6px; }
            /* ปุ่มรอง: พื้นขาว เส้นขอบ */
            QPushButton[secondary="true"] {
                background-color: #ffffff;
                color: #344767;
                border: 1px solid #cfd6df;
                font-weight: 500;
            }
            QPushButton[secondary="true"]:hover {
                background-color: #f4f7fb;
                border-color: #1f3b5b;
                color: #1f3b5b;
            }
            QPushButton[secondary="true"]:disabled {
                background-color: #f6f7f9; color: #b8c0ca; border-color: #e3e7ec;
            }

            /* สไตล์เฉพาะสำหรับปุ่มต่างๆ */
            QPushButton#load_button {
                background-color: #28a745; /* สีเขียวสำหรับโหลดไฟล์ */
            }
            QPushButton#load_button:hover {
                background-color: #218838;
            }
            QPushButton#set_queue_vars_button {
                background-color: #17a2b8; /* สีฟ้า-เขียวสำหรับเลือกตัวแปร */
                min-height: 34px;
                font-size: 9.5pt;
            }
            QPushButton#set_queue_vars_button:hover {
                background-color: #138496;
            }
            QPushButton#stack_vars_button {
                background-color: #5c7cfa; /* น้ำเงินอินดิโก้: ต่อ Data ปกติ */
            }
            QPushButton#stack_vars_button:hover { background-color: #4c6ef5; }
            QPushButton#pattern_stack_button {
                background-color: #fd7e14; /* ส้ม: ต่อ Data แบบกลุ่ม */
            }
            QPushButton#pattern_stack_button:hover { background-color: #e8700d; }
            QPushButton#recode_jr_button {
                background-color: #20c997; /* เขียวมิ้นต์: Recode Justright */
            }
            QPushButton#recode_jr_button:hover { background-color: #1aa179; }
            QPushButton:disabled {
                background-color: #c8ced6;
                color: #f4f6f8;
            }
            QPushButton#add_to_queue_button {
                background-color: #ffc107; /* สีส้มสำหรับเพิ่มลงคิว */
                color: #333333; /* ตัวอักษรสีดำ */
                min-height: 36px;
                font-size: 10pt;
            }
            QPushButton#add_to_queue_button:hover {
                background-color: #e0a800;
            }
            QPushButton#apply_filter_button {
                background-color: #fd7e14; /* สีส้มเข้มสำหรับใช้ Filter */
            }
            QPushButton#apply_filter_button:hover {
                background-color: #e66b0d;
            }
            QPushButton#clear_filter_button {
                background-color: #dc3545; /* สีแดงสำหรับล้าง Filter */
            }
            QPushButton#clear_filter_button:hover {
                background-color: #c82333;
            }
            QPushButton#execute_template_button {
                background-color: #6f42c1; /* สีม่วง: ปุ่มหลักของโปรแกรม */
                font-weight: bold;
                font-size: 10.5pt;
                min-height: 44px;
            }
            QPushButton#execute_template_button:hover {
                background-color: #5d34a4;
            }
            QPushButton#remove_from_queue_button {
                background-color: #ffffff; color: #c0392b; border: 1px solid #e6a8a1;
            }
            QPushButton#remove_from_queue_button:hover {
                background-color: #fdf0ee;
            }
            QPushButton#remove_from_queue_button:disabled, QPushButton#clear_queue_button:disabled {
                background-color: #f6f7f9; color: #b8c0ca; border-color: #e3e7ec;
            }
            QPushButton#clear_queue_button {
                background-color: #ffffff; color: #c0392b; border: 1px solid #e6a8a1;
            }
            QPushButton#clear_queue_button:hover {
                background-color: #fdf0ee;
            }

            QLabel {
                font-size: 9.5pt;
                color: #333333;
            }
            QLabel#file_label, QLabel#current_filter_label {
                font-style: italic;
                color: #555555;
            }
            QLabel#status_label {
                color: #1e7e34; /* สีเขียวสำหรับ Status */
                font-weight: 600;
                font-size: 9.5pt;
                padding: 10px 14px;
                border: 1px solid #cde9d4;
                border-radius: 10px;
                background-color: #edf8f0;
            }
            QLineEdit, QTextEdit, QListWidget {
                border: 1px solid #d5dbe3; /* สีเทาอมฟ้าอ่อน */
                border-radius: 8px;
                padding: 8px 10px;
                font-size: 9.5pt;
                background-color: #fbfcfd;
            }
            QLineEdit:hover, QListWidget:hover { border-color: #b7c2cf; }
            QLineEdit:focus, QTextEdit:focus, QListWidget:focus {
                border: 1px solid #1f3b5b;
                background-color: #ffffff;
            }
            QListWidget { padding: 4px; }
            QListWidget::item {
                padding: 5px 8px; /* เพิ่ม padding ของแต่ละ item ใน list */
                border-radius: 4px;
            }
            QListWidget::item:hover {
                background-color: #f1f3f5;
            }
            QListWidget::item:selected {
                background-color: #e7f1ff; /* สีฟ้าอ่อนสำหรับ item ที่เลือก */
                color: #0b3d91;
            }
            QScrollBar:vertical {
                background: transparent; width: 10px; margin: 2px;
            }
            QScrollBar::handle:vertical {
                background: #c5cbd3; border-radius: 5px; min-height: 30px;
            }
            QScrollBar::handle:vertical:hover { background: #9aa3ad; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
            QSplitter::handle {
                background-color: #ced4da; /* สีเทาอ่อนสำหรับ handle ของ splitter */
                border-radius: 4px;
                width: 8px; /* ความกว้างของ handle */
            }
            QSplitter::handle:hover {
                background-color: #aebacd; /* สีเข้มขึ้นเมื่อ hover */
            }
            QSplitter::handle:pressed {
                background-color: #8c9bb3; /* สีเข้มขึ้นเมื่อกดค้าง */
            }
        """)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        outer_layout = QVBoxLayout(central_widget)
        outer_layout.setContentsMargins(15, 15, 15, 15)
        outer_layout.setSpacing(14)

        outer_layout.addWidget(self._create_app_header())

        # ใช้ QSplitter เพื่อให้ปรับขนาดแต่ละ Panel ได้
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self._create_left_panel())
        splitter.addWidget(self._create_right_panel())
        splitter.setChildrenCollapsible(False)
        outer_layout.addWidget(splitter, 1)

        self.update_ui_state()
        self._center_on_screen() # ถูกเรียกใช้ที่นี่เป็นอันดับสุดท้าย เพื่อกำหนดขนาดและจัดกลาง

        # กำหนดขนาดเริ่มต้นของ Panel ซ้าย-ขวา
        # ใช้ int() ครอบเพื่อให้แน่ใจว่าเป็นจำนวนเต็มตามที่ setSizes ต้องการ
        # ค่า self.width() และ self.height() จะเป็นค่าหลังจากที่ _center_on_screen() ตั้งค่าให้แล้ว
        splitter.setSizes([int(self.width() * 0.6), int(self.width() * 0.4)]) # กำหนดขนาดเริ่มต้น (ประมาณ 60:40)


    def _create_app_header(self):
        """แถบหัวโปรแกรมด้านบน: ชื่อโปรแกรม + สถานะไฟล์"""
        header = QFrame()
        header.setObjectName("app_header")
        h = QHBoxLayout(header)
        h.setContentsMargins(20, 14, 20, 14)
        h.setSpacing(12)

        text_box = QVBoxLayout()
        text_box.setSpacing(2)
        title = QLabel("Correlation จาก SPSS Data")
        title.setObjectName("app_title")
        subtitle = QLabel("โหลดไฟล์ .sav  →  ตั้งค่าการรัน  →  เพิ่มลงคิว  →  Export เป็น Excel")
        subtitle.setObjectName("app_subtitle")
        text_box.addWidget(title)
        text_box.addWidget(subtitle)
        h.addLayout(text_box, 1)

        self.file_chip = QLabel("ยังไม่ได้โหลดไฟล์")
        self.file_chip.setObjectName("file_chip")
        self.file_chip.setProperty("loaded", "false")
        h.addWidget(self.file_chip, 0, Qt.AlignmentFlag.AlignVCenter)
        return header

    def _set_file_chip(self, text, loaded):
        """อัปเดตป้ายสถานะไฟล์บนแถบหัวโปรแกรม"""
        self.file_chip.setText(text)
        self.file_chip.setProperty("loaded", "true" if loaded else "false")
        self.file_chip.style().unpolish(self.file_chip)
        self.file_chip.style().polish(self.file_chip)

    def _create_left_panel(self):
        """สร้าง Panel ด้านซ้าย: โหลดไฟล์, ตั้งค่ารัน"""
        left_widget = QWidget()
        layout = QVBoxLayout(left_widget)
        layout.setSpacing(20) # เพิ่มระยะห่างระหว่าง GroupBox

        layout.addWidget(self._create_file_loading_section())
        layout.addWidget(self._create_run_setup_section())
        layout.addStretch() # ดันส่วน Status ไปด้านล่าง

        self.status_label = QLabel("สถานะ: พร้อมใช้งาน")
        self.status_label.setObjectName("status_label") # กำหนด Object Name เพื่อใช้ใน CSS
        layout.addWidget(self.status_label)

        return left_widget

    def _create_right_panel(self):
        """สร้าง Panel ด้านขวา: คิวการรันและจัดการ Template"""
        queue_widget = QWidget()
        queue_layout = QVBoxLayout(queue_widget)
        queue_layout.setSpacing(15)

        queue_layout.addWidget(self._create_template_queue_tab(), 1)

        return queue_widget

    def _create_file_loading_section(self):
        """ส่วนสำหรับโหลดไฟล์ SPSS"""
        group_box = QGroupBox("1  โหลดไฟล์ SPSS")
        layout = QVBoxLayout(group_box)
        layout.setSpacing(10)

        file_layout = QHBoxLayout()
        self.load_button = QPushButton("📂  เลือกไฟล์ SPSS (.sav)")
        self.load_button.setObjectName("load_button")
        self.load_button.setMinimumHeight(34)
        self.load_button.clicked.connect(self.load_spss_file)

        self.file_label = QLabel("ยังไม่ได้โหลดไฟล์ — รองรับไฟล์ .sav จาก SPSS")
        self.file_label.setObjectName("file_label")
        self.file_label.setWordWrap(True) # ให้ข้อความขึ้นบรรทัดใหม่ได้

        file_layout.addWidget(self.load_button)
        file_layout.addWidget(self.file_label, 1) # ให้ label ขยายได้

        layout.addLayout(file_layout)
        return group_box

    def _create_run_setup_section(self):
        """ส่วนสำหรับตั้งค่าการรันเดี่ยวหรือเพิ่มลงคิว"""
        setup_group_box = QGroupBox("2  ตั้งค่าการรัน")
        setup_layout = QVBoxLayout(setup_group_box)
        setup_layout.setSpacing(8)

        def field_label(text):
            lbl = QLabel(text)
            lbl.setObjectName("field_label")
            return lbl

        setup_layout.addWidget(field_label("ชื่อ Sheet ที่จะ Export"))
        self.single_sheet_name_edit = QLineEdit()
        self.single_sheet_name_edit.setPlaceholderText("เช่น Analysis_Q1_vs_Q2")
        setup_layout.addWidget(self.single_sheet_name_edit)

        setup_layout.addWidget(field_label("ข้อความหัวข้อใน Excel (Cell A1)"))
        self.single_header_a1_edit = QLineEdit()
        self.single_header_a1_edit.setPlaceholderText("ถ้าว่างจะใช้ชื่อ Sheet เป็นหัวข้อ")
        setup_layout.addWidget(self.single_header_a1_edit)

        setup_layout.addWidget(field_label("Filter เฉพาะสำหรับ Run นี้"))
        self.single_filter_edit = QLineEdit()
        self.single_filter_edit.setPlaceholderText("เช่น gender == 1 (ถ้าว่างไม่ใช้ Filter)")
        setup_layout.addWidget(self.single_filter_edit)

        setup_layout.addSpacing(6)
        setup_layout.addWidget(field_label("ตัวแปรสำหรับ Correlation"))
        self.set_queue_vars_button = QPushButton("คลิกเพื่อเลือกตัวแปรที่จะรัน Correlation")
        self.set_queue_vars_button.setObjectName("set_queue_vars_button")
        self.set_queue_vars_button.clicked.connect(self.open_set_queue_vars_dialog)
        setup_layout.addWidget(self.set_queue_vars_button)

        # ปุ่มสร้างตัวแปร Stack
        setup_layout.addSpacing(6)
        tools_label = QLabel("เครื่องมือสร้างตัวแปรใหม่ — ใช้ก่อนเลือกตัวแปร")
        tools_label.setObjectName("section_hint")
        setup_layout.addWidget(tools_label)
        button_layout = QHBoxLayout()
        self.stack_vars_button = QPushButton("ต่อ Data ปกติ")
        self.stack_vars_button.setObjectName("stack_vars_button")
        self.stack_vars_button.setToolTip("เปิดหน้าต่างเพื่อเลือกตัวแปรมา Stack เป็นตัวแปรใหม่ 1 ตัว")
        self.stack_vars_button.clicked.connect(self.open_stack_variables_dialog)
        
        self.pattern_stack_button = QPushButton("ต่อ Data แบบกลุ่ม")
        self.pattern_stack_button.setObjectName("pattern_stack_button")
        self.pattern_stack_button.setToolTip("สร้างตัวแปรใหม่หลายตัวโดยจัดกลุ่มจากแพทเทิร์นในชื่อตัวแปร")
        self.pattern_stack_button.clicked.connect(self.open_pattern_stack_dialog)
        
        self.recode_jr_button = QPushButton("Recode Justright")
        self.recode_jr_button.setObjectName("recode_jr_button")
        self.recode_jr_button.setToolTip("สร้างตัวแปรใหม่จากตัวแปรเดิมโดย Recode แบบ Justright (5 Scale: 1,5=1 / 2,4=2 / 3=3 | 7 Scale: 1,7=1 / 2,6=2 / 3,5=3 / 4=4)")
        self.recode_jr_button.clicked.connect(self.open_recode_justright_dialog)

        button_layout.addWidget(self.stack_vars_button)
        button_layout.addWidget(self.pattern_stack_button)
        button_layout.addWidget(self.recode_jr_button)
        setup_layout.addLayout(button_layout)

        setup_layout.addSpacing(8)
        self.add_to_queue_button = QPushButton("➕  เพิ่มรายการนี้ลงในคิวการประมวลผล")
        self.add_to_queue_button.setObjectName("add_to_queue_button")
        self.add_to_queue_button.clicked.connect(self.add_current_setup_to_template_queue)
        setup_layout.addWidget(self.add_to_queue_button)

        return setup_group_box

    def _create_global_filter_section(self):
        """ส่วนสำหรับกำหนด Global Filter"""
        filter_group_box = QGroupBox("3. Global Filter (ใช้กับทุก Run)")
        filter_layout_main = QVBoxLayout(filter_group_box)
        filter_layout_main.setSpacing(8)

        label_desc = QLabel("Filter นี้จะถูกใช้กับทุก Run หากใน Run นั้นไม่ได้ระบุ Filter เอง")
        label_desc.setWordWrap(True)
        label_desc.setStyleSheet("font-weight: normal; color: #34495e;") # สีข้อความคำอธิบาย

        filter_layout_main.addWidget(label_desc)

        self.filter_query_edit = QLineEdit()
        self.filter_query_edit.setPlaceholderText("เช่น s7_check==1 & check_quota==1")
        self.filter_query_edit.textChanged.connect(self.update_ui_state) # อัปเดตสถานะปุ่มเมื่อข้อความเปลี่ยน
        filter_layout_main.addWidget(self.filter_query_edit)

        filter_action_layout = QHBoxLayout()
        self.apply_filter_button = QPushButton("ใช้ Filter")
        self.apply_filter_button.setObjectName("apply_filter_button")
        self.apply_filter_button.clicked.connect(self.apply_global_filter_query)
        
        self.clear_filter_button = QPushButton("ล้าง Filter")
        self.clear_filter_button.setObjectName("clear_filter_button")
        self.clear_filter_button.clicked.connect(self.clear_global_filter_query)
        
        filter_action_layout.addWidget(self.apply_filter_button)
        filter_action_layout.addWidget(self.clear_filter_button)
        filter_layout_main.addLayout(filter_action_layout)

        self.current_filter_label = QLabel(f"Filter ที่ใช้อยู่: <span style='color: #28a745; font-weight: bold;'>{self.active_filter_description}</span>")
        self.current_filter_label.setObjectName("current_filter_label")
        self.current_filter_label.setWordWrap(True)
        filter_layout_main.addWidget(self.current_filter_label)

        return filter_group_box

    def _create_template_queue_tab(self):
        """ส่วนสำหรับจัดการคิวและ Template"""
        group_box = QGroupBox("3  คิวการประมวลผล & Template")
        layout = QVBoxLayout(group_box)
        layout.setSpacing(10)

        # ปุ่มจัดการคิว
        template_action_layout = QVBoxLayout()
        self.save_template_button = QPushButton("💾  บันทึกคิวเป็น Template")
        self.save_template_button.setObjectName("save_template_button")
        self.save_template_button.setProperty("secondary", "true")
        self.save_template_button.clicked.connect(self.save_setups_as_template)

        self.load_template_button = QPushButton("📄  โหลด Template")
        self.load_template_button.setObjectName("load_template_button")
        self.load_template_button.setProperty("secondary", "true")
        self.load_template_button.clicked.connect(self.load_setups_from_template)

        self.execute_template_button = QPushButton("▶  รันทุกรายการในคิวและ Export Excel")
        self.execute_template_button.setObjectName("execute_template_button")
        self.execute_template_button.clicked.connect(self.execute_all_queued_setups)

        template_row = QHBoxLayout()
        template_row.addWidget(self.save_template_button)
        template_row.addWidget(self.load_template_button)
        template_action_layout.addLayout(template_row)
        layout.addLayout(template_action_layout)

        # รายการในคิว
        queue_header = QHBoxLayout()
        queue_title = QLabel("รายการในคิวการประมวลผล")
        queue_title.setObjectName("field_label")
        self.queue_count_label = QLabel("0 รายการ")
        self.queue_count_label.setObjectName("section_hint")
        queue_header.addWidget(queue_title)
        queue_header.addStretch()
        queue_header.addWidget(self.queue_count_label)
        layout.addLayout(queue_header)
        hint = QLabel("ดับเบิลคลิกรายการเพื่อดึงกลับมาแก้ไข")
        hint.setObjectName("section_hint")
        layout.addWidget(hint)
        self.template_queue_list_widget = QListWidget()
        self.template_queue_list_widget.setObjectName("queue_list")
        self.template_queue_list_widget.setSpacing(2)
        self.template_queue_list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.template_queue_list_widget.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.template_queue_list_widget.itemDoubleClicked.connect(self.load_queued_item_to_setup_ui)
        self.template_queue_list_widget.itemSelectionChanged.connect(self.update_ui_state) # อัปเดตสถานะปุ่มเมื่อเลือกรายการ
        layout.addWidget(self.template_queue_list_widget, 1)

        # ปุ่มจัดการรายการในคิว
        queue_buttons_layout = QHBoxLayout()
        self.remove_from_queue_button = QPushButton("ลบรายการที่เลือก")
        self.remove_from_queue_button.setObjectName("remove_from_queue_button")
        self.remove_from_queue_button.clicked.connect(self.remove_selected_from_template_queue)

        self.clear_queue_button = QPushButton("ล้างคิวทั้งหมด")
        self.clear_queue_button.setObjectName("clear_queue_button")
        self.clear_queue_button.clicked.connect(self.clear_template_queue)

        queue_buttons_layout.addWidget(self.remove_from_queue_button)
        queue_buttons_layout.addWidget(self.clear_queue_button)
        layout.addLayout(queue_buttons_layout)

        layout.addSpacing(6)
        layout.addWidget(self.execute_template_button)

        return group_box

    def _center_on_screen(self):
        """จัดตำแหน่งหน้าต่างให้อยู่กึ่งกลางหน้าจอและกำหนดขนาดเริ่มต้น"""
        try:
            screen = QGuiApplication.primaryScreen().availableGeometry()

            # กำหนดขนาดเริ่มต้นของหน้าต่าง
            target_width = min(1000, screen.width() - 80)  # ความกว้างเริ่มต้น
            target_height = min(900, screen.height() - 80) # ความสูงเริ่มต้น
            self.resize(target_width, target_height) # กำหนดขนาดเริ่มต้น

            # จัดตำแหน่งหน้าต่างให้อยู่กึ่งกลาง โดยใช้ขนาดที่เพิ่งตั้งค่าไป
            self.move(screen.center().x() - self.width() // 2, screen.center().y() - self.height() // 2)
        except Exception:
            # Fallback หากข้อมูลหน้าจอไม่พร้อมใช้งาน
            self.resize(800, 600) # Fallback with new initial size
            self.move(100, 100)

    def update_ui_state(self):
        """อัปเดตสถานะของ UI elements (เปิด/ปิดใช้งาน) ตามเงื่อนไข"""
        file_loaded = self.df is not None
        
        # ส่วนตั้งค่าการรัน
        for widget in [self.single_sheet_name_edit, self.single_header_a1_edit, self.single_filter_edit,
                        self.set_queue_vars_button, self.add_to_queue_button,
                        self.stack_vars_button, self.pattern_stack_button, self.recode_jr_button]:
            widget.setEnabled(file_loaded)

        # จัดการคิวและ Template
        queue_has_items = len(self.template_setups_for_queue) > 0
        self.save_template_button.setEnabled(queue_has_items)
        self.load_template_button.setEnabled(file_loaded) # สามารถโหลด Template ได้เมื่อโหลดไฟล์ SPSS แล้ว
        self.execute_template_button.setEnabled(queue_has_items and file_loaded)
        self.remove_from_queue_button.setEnabled(bool(self.template_queue_list_widget.selectedItems()))
        self.clear_queue_button.setEnabled(queue_has_items)

    def load_spss_file(self):
        """โหลดไฟล์ SPSS (.sav) และอ่านข้อมูล"""
        file_name, _ = QFileDialog.getOpenFileName(self, "โหลดไฟล์ SPSS", "", "SPSS Files (*.sav)")
        if file_name:
            try:
                # ใช้ pyreadstat ในการอ่านไฟล์ .sav
                self.df, self.meta = pyreadstat.read_sav(
                    file_name,
                    apply_value_formats=False, # ไม่แปลง value labels เป็น category โดยตรง
                    formats_as_category=False, # ไม่แปลงตัวแปรเป็น category type
                    user_missing=True # จัดการ User-defined missing values
                )
                short_name = file_name.split('/')[-1]
                self.file_label.setText(f"ไฟล์ที่โหลด: {short_name}   •  {len(self.df):,} แถว  •  {len(self.df.columns):,} ตัวแปร")
                self.file_label.setStyleSheet("color: #1e7e34; font-weight: 600; font-size: 9.5pt; font-style: normal;")
                self._set_file_chip(f"●  {short_name}", True)
                
                self.reset_setup_form_for_next_run()
                self.active_filter_query = ""
                self.active_filter_description = "ไม่มี"
                self.clear_template_queue(show_message=False) # ล้างคิวเก่า
                self.virtual_variables = {} # ล้างตัวแปรเสมือนที่สร้างไว้
                
                self.status_label.setText("สถานะ: โหลดไฟล์ SPSS สำเร็จ. กรุณาตั้งค่าการรัน หรือโหลด Template")
                QMessageBox.information(self, "สำเร็จ", "โหลดไฟล์ SPSS เรียบร้อยแล้ว")
            except Exception as e:
                QMessageBox.critical(self, "เกิดข้อผิดพลาด", f"ไม่สามารถโหลดไฟล์ SPSS ได้: {e}\nโปรดตรวจสอบว่าไฟล์ไม่เสียหายและเป็นรูปแบบ .sav ที่ถูกต้อง")
                self.df = None
                self.meta = None
                self.file_label.setText("ยังไม่ได้โหลดไฟล์ — รองรับไฟล์ .sav จาก SPSS")
                self.file_label.setStyleSheet("font-style: italic; color: gray; font-size: 9.5pt;")
                self._set_file_chip("ยังไม่ได้โหลดไฟล์", False)
            finally:
                self.update_ui_state()

    def apply_global_filter_query(self):
        """ใช้ Global Filter กับข้อมูล"""
        if not hasattr(self, "filter_query_edit") or not hasattr(self, "current_filter_label"):
            return
        if self.df is None:
            QMessageBox.warning(self, "ไม่มีข้อมูล", "กรุณาโหลดไฟล์ SPSS ก่อน")
            return
        
        query_str = self.filter_query_edit.text().strip()
        if not query_str:
            self.clear_global_filter_query()
            return
        
        try:
            # ลอง query เพื่อตรวจสอบ syntax ก่อนที่จะบันทึก
            _ = self.df.query(query_str)
            self.active_filter_query = query_str
            self.active_filter_description = query_str
            QMessageBox.information(self, "สำเร็จ", f"Global Filter '{query_str}' ถูกตั้งค่าแล้ว")
        except Exception as e:
            QMessageBox.critical(self, "Filter ไม่ถูกต้อง", f"Query ไม่ถูกต้อง: {e}\nโปรดตรวจสอบไวยากรณ์ (Syntax) ของ Filter")
            self.active_filter_description = "ไม่ถูกต้อง" # อัปเดตสถานะให้เป็นไม่ถูกต้อง
        finally:
            self.current_filter_label.setText(f"Filter ที่ใช้อยู่: <span style='color: #28a745; font-weight: bold;'>{self.active_filter_description}</span>")
            self.update_ui_state()

    def clear_global_filter_query(self, show_message=True):
        """ล้าง Global Filter ที่ตั้งค่าไว้"""
        self.active_filter_query = ""
        self.active_filter_description = "ไม่มี"
        if hasattr(self, "filter_query_edit"):
            self.filter_query_edit.clear()
        if hasattr(self, "current_filter_label"):
            self.current_filter_label.setText(f"Filter ที่ใช้อยู่: <span style='color: #28a745; font-weight: bold;'>{self.active_filter_description}</span>")
        if show_message:
            QMessageBox.information(self, "สำเร็จ", "ล้าง Global Filter แล้ว")
        self.update_ui_state()

    def reset_setup_form_for_next_run(self):
        """เคลียร์ฟอร์มการตั้งค่า Run เพื่อเตรียมรับข้อมูลสำหรับรายการต่อไป"""
        self.single_sheet_name_edit.clear()
        self.single_header_a1_edit.clear()
        self.single_filter_edit.clear()
        self.current_queue_setup_vars = []
        self.set_queue_vars_button.setText("คลิกเพื่อเลือกตัวแปรที่จะรัน Correlation")
        self.set_queue_vars_button.setStyleSheet("background-color: #17a2b8; color: white;") # กลับไปใช้สีเริ่มต้น

    def open_stack_variables_dialog(self):
        """เปิด Dialog สำหรับสร้างตัวแปรแบบ Stack (เดี่ยว)"""
        if self.df is None:
            QMessageBox.warning(self, "ไม่มีข้อมูล", "กรุณาโหลดไฟล์ SPSS ก่อน")
            return
        dialog = StackVariablesDialog(self.df.columns.tolist(), self.meta, self.active_filter_description, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._create_stacked_variable(dialog.get_data())

    def _create_stacked_variable(self, data):
        """สร้าง 'สูตร' สำหรับตัวแปรเสมือนที่ได้จากการ Stack"""
        new_name, new_label, source_vars = data["new_name"], data["label"], data["source_vars"]
        
        # ตรวจสอบว่าชื่อตัวแปรใหม่ซ้ำกับตัวแปรที่มีอยู่จริงหรือตัวแปรเสมือนที่สร้างไปแล้วหรือไม่
        if new_name in (self.df.columns.tolist() + list(self.virtual_variables.keys())):
            QMessageBox.critical(self, "ชื่อซ้ำ", f"ชื่อตัวแปร '{new_name}' มีอยู่แล้ว กรุณาใช้ชื่ออื่น")
            return
        
        # ตรวจสอบว่า source_vars ทั้งหมดมีอยู่ใน DataFrame หรือไม่
        invalid_sources = [v for v in source_vars if v not in self.df.columns]
        if invalid_sources:
            QMessageBox.critical(self, "ตัวแปรต้นฉบับไม่ถูกต้อง", f"ตัวแปรต้นฉบับเหล่านี้ไม่พบในไฟล์ SPSS: {', '.join(invalid_sources)}")
            return

        # บันทึกแค่ "สูตร" สำหรับตัวแปรเสมือน (จะไม่คำนวณจริงจนกว่าจะ Export)
        self.virtual_variables[new_name] = {"label": new_label, "source_vars": source_vars, "type": "single_stack"}
        QMessageBox.information(self, "สำเร็จ", f"สร้างตัวแปรต่อ Data>> '{new_name}' เรียบร้อยแล้ว\nตัวแปรนี้จะถูกคำนวณเมื่อทำการ Export")

    def open_pattern_stack_dialog(self):
        """เปิด Dialog สำหรับสร้างตัวแปร Stack แบบกลุ่มตามแพทเทิร์น"""
        if self.df is None:
            QMessageBox.warning(self, "ไม่มีข้อมูล", "กรุณาโหลดไฟล์ SPSS ก่อน")
            return
        dialog = PatternStackDialog(self.df.columns.tolist(), self.meta, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._create_pattern_stacked_variables(dialog.get_data())

    def _create_pattern_stacked_variables(self, data):
        """สร้าง 'สูตร' สำหรับตัวแปรเสมือนแบบกลุ่มที่ได้จากแพทเทิร์น"""
        groups, base_name = data["groups"], data["base_name"]
        
        created_count = 0
        all_available_vars = self.df.columns.tolist() + list(self.virtual_variables.keys()) # ตัวแปรที่มีอยู่ทั้งหมด (จริง + เสมือน)
        
        for item_index, vars_to_stack in groups.items():
            try:
                new_var_name = f"{base_name}_{item_index.replace('#', '')}"
                
                if new_var_name in all_available_vars:
                    print(f"ข้ามการสร้าง '{new_var_name}' เนื่องจากมีชื่อซ้ำ")
                    continue # ข้ามถ้าชื่อซ้ำ

                # ตรวจสอบว่า source_vars ทั้งหมดมีอยู่ใน DataFrame หรือไม่
                invalid_sources = [v for v in vars_to_stack if v not in self.df.columns]
                if invalid_sources:
                    QMessageBox.critical(self, "ตัวแปรต้นฉบับไม่ถูกต้อง", f"ไม่สามารถสร้างตัวแปร '{new_var_name}' ได้: ตัวแปรต้นฉบับเหล่านี้ไม่พบในไฟล์ SPSS: {', '.join(invalid_sources)}")
                    continue # ข้ามถ้า source ไม่ถูกต้อง

                new_var_label = self.meta.column_names_to_labels.get(vars_to_stack[0], vars_to_stack[0])
                
                # บันทึกแค่ "สูตร" สำหรับตัวแปรเสมือน
                # --- จุดที่แก้ไข: เปลี่ยนจาก new_name เป็น new_var_name ---
                self.virtual_variables[new_var_name] = {"label": new_var_label, "source_vars": vars_to_stack, "type": "pattern_stack"}
                all_available_vars.append(new_var_name) # เพิ่มชื่อใหม่ลงในรายการที่มีอยู่เพื่อตรวจสอบการซ้ำ
                created_count += 1
            except Exception as e:
                QMessageBox.critical(self, "เกิดข้อผิดพลาดระหว่างสร้างกลุ่ม", f"ไม่สามารถสร้างตัวแปรสำหรับกลุ่ม {item_index} ได้: {e}")
                continue
        
        if created_count > 0:
            QMessageBox.information(self, "สำเร็จ", f"สร้าง 'สูตร' สำหรับตัวแปรเสมือนแบบกลุ่มสำเร็จ {created_count} รายการ")
        else:
            QMessageBox.warning(self, "ไม่สำเร็จ", "ไม่สามารถสร้างตัวแปรเสมือนใหม่ได้ (อาจมีชื่อซ้ำทั้งหมดหรือตัวแปรต้นฉบับไม่ถูกต้อง)")

    def open_recode_justright_dialog(self):
        """เปิด Dialog สำหรับสร้างตัวแปร Recode Justright"""
        if self.df is None:
            QMessageBox.warning(self, "ไม่มีข้อมูล", "กรุณาโหลดไฟล์ SPSS ก่อน")
            return
        dialog = RecodeJustrightDialog(self.df.columns.tolist(), self.meta, self.df, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._create_recode_justright_variables(dialog.get_data())

    def _create_recode_justright_variables(self, items):
        """สร้าง 'สูตร' ตัวแปรเสมือนแบบ Recode Justright (คำนวณจริงตอน Export)"""
        all_available_vars = set(self.df.columns.tolist() + list(self.virtual_variables.keys()))
        created, skipped = [], []
        for it in items:
            src, new_name, scale = it["source"], it["new_name"], int(it["scale"])
            if new_name in all_available_vars:
                skipped.append(f"{new_name} (ชื่อซ้ำ)")
                continue
            if src not in self.df.columns:
                skipped.append(f"{new_name} (ไม่พบ {src})")
                continue
            base_label = self.meta.column_names_to_labels.get(src, src) if self.meta and self.meta.column_names_to_labels else src
            self.virtual_variables[new_name] = {
                "label": f"{base_label} (Justright {scale} Scale)",
                "source_vars": [src],
                "type": "recode_justright",
                "scale": scale,
            }
            all_available_vars.add(new_name)
            created.append(f"{src} -> {new_name} ({scale} Scale)")
        msg = ""
        if created:
            msg += f"สร้างตัวแปร Recode Justright สำเร็จ {len(created)} ตัวแปร:\n" + "\n".join(created)
        if skipped:
            msg += ("\n\n" if msg else "") + f"ข้าม {len(skipped)} รายการ:\n" + "\n".join(skipped)
        if created:
            QMessageBox.information(self, "สำเร็จ", msg)
        else:
            QMessageBox.warning(self, "ไม่สำเร็จ", msg or "ไม่มีตัวแปรที่สร้างได้")

    def open_set_queue_vars_dialog(self):
        """เปิด Dialog สำหรับเลือกตัวแปรที่จะรัน Correlation"""
        if self.df is None:
            QMessageBox.warning(self, "ไม่มีข้อมูล", "กรุณาโหลดไฟล์ SPSS ก่อน")
            return
        # รวมตัวแปรจริงและตัวแปรเสมือนเข้าด้วยกันสำหรับให้เลือก
        all_displayable_vars = self.df.columns.tolist() + list(self.virtual_variables.keys())
        
        dialog = SetQueueVarsDialog(all_displayable_vars, self.current_queue_setup_vars, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.current_queue_setup_vars = dialog.get_selected_vars()
            count = len(self.current_queue_setup_vars)
            if count > 0:
                self.set_queue_vars_button.setText(f"ตั้งค่า/แก้ไขตัวแปร... (เลือกแล้ว {count} ตัวแปร)")
                self.set_queue_vars_button.setStyleSheet("background-color: #007bff; color: white;") # เปลี่ยนสีเมื่อมีการเลือก
            else:
                self.set_queue_vars_button.setText("คลิกเพื่อเลือกตัวแปรที่จะรัน Correlation")
                self.set_queue_vars_button.setStyleSheet("background-color: #17a2b8; color: white;") # กลับไปใช้สีเดิม
            self.update_ui_state()

    def add_current_setup_to_template_queue(self):
        """เพิ่มการตั้งค่า Run ปัจจุบันลงในคิว"""
        sheet_name = self.single_sheet_name_edit.text().strip()
        if not sheet_name:
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณาใส่ชื่อ Sheet")
            return
        
        # ตรวจสอบว่าชื่อ Sheet ซ้ำในคิวหรือไม่
        if any(s['sheet_name'] == sheet_name for s in self.template_setups_for_queue):
            QMessageBox.warning(self, "ชื่อ Sheet ซ้ำ", f"ชื่อ Sheet '{sheet_name}' มีอยู่ในคิวแล้ว กรุณาใช้ชื่ออื่น")
            return

        selected_vars = self.current_queue_setup_vars
        if len(selected_vars) < 2:
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณาตั้งค่าตัวแปรอย่างน้อย 2 ตัวสำหรับ Correlation")
            return
        
        # ตรวจสอบว่าตัวแปรที่เลือกมานั้น มีอยู่ใน DataFrame หรือเป็น Virtual Variable ที่สร้างไว้แล้วหรือไม่
        all_valid_vars = set(self.df.columns.tolist() + list(self.virtual_variables.keys()))
        invalid_selected_vars = [v for v in selected_vars if v not in all_valid_vars]
        if invalid_selected_vars:
            QMessageBox.critical(self, "ตัวแปรไม่ถูกต้อง", f"ตัวแปรเหล่านี้ไม่พบในข้อมูลหรือตัวแปรเสมือนที่สร้างไว้: {', '.join(invalid_selected_vars)}\nโปรดตรวจสอบหรือสร้างตัวแปรเสมือนก่อน")
            return

        header_a1 = self.single_header_a1_edit.text().strip() or sheet_name
        filter_query = self.single_filter_edit.text().strip()

        # ตรวจสอบ Filter syntax ก่อนเพิ่มลงคิว หาก Filter ถูกระบุ
        if filter_query:
            try:
                # ลอง query บน DataFrame ต้นฉบับเพื่อตรวจสอบ syntax
                _ = self.df.query(filter_query)
            except Exception as e:
                QMessageBox.critical(self, "Filter ไม่ถูกต้อง", f"Filter สำหรับ Run นี้ '{filter_query}' มีไวยากรณ์ไม่ถูกต้อง: {e}\nโปรดแก้ไขก่อนเพิ่มลงคิว")
                return

        self.template_setups_for_queue.append({
            'sheet_name': sheet_name,
            'header_a1': header_a1,
            'filter_query': filter_query,
            'variables_list': selected_vars
        })
        self.update_template_queue_list_widget()
        self.status_label.setText(f"สถานะ: เพิ่ม '{sheet_name}' ลงในคิวแล้ว. จำนวนในคิว: {len(self.template_setups_for_queue)}")
        self.update_ui_state()
        self.reset_setup_form_for_next_run() # เคลียร์ฟอร์มหลังเพิ่มลงคิว

    def _make_queue_item_widget(self, index, setup):
        """สร้างการ์ดแสดงรายละเอียด 1 รายการในคิว"""
        card = QWidget()
        card.setObjectName("queue_card")
        h = QHBoxLayout(card)
        h.setContentsMargins(10, 8, 10, 8)
        h.setSpacing(12)

        badge = QLabel(str(index))
        badge.setObjectName("queue_badge")
        badge.setFixedSize(30, 30)
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        h.addWidget(badge, 0, Qt.AlignmentFlag.AlignTop)

        body = QVBoxLayout()
        body.setSpacing(4)
        title = QLabel(setup.get('sheet_name', ''))
        title.setObjectName("queue_title")
        body.addWidget(title)

        header_a1 = setup.get('header_a1', '')
        if header_a1 and header_a1 != setup.get('sheet_name', ''):
            sub = QLabel(header_a1)
            sub.setObjectName("queue_subtitle")
            sub.setWordWrap(True)
            body.addWidget(sub)

        chips = QHBoxLayout()
        chips.setSpacing(6)
        var_chip = QLabel(f"{len(setup.get('variables_list', []))} ตัวแปร")
        var_chip.setObjectName("queue_chip")
        var_chip.setProperty("kind", "vars")
        chips.addWidget(var_chip)
        n_virtual = sum(1 for v in setup.get('variables_list', []) if v in self.virtual_variables)
        if n_virtual:
            v_chip = QLabel(f"ตัวแปรสร้างใหม่ {n_virtual}")
            v_chip.setObjectName("queue_chip")
            v_chip.setProperty("kind", "virtual")
            chips.addWidget(v_chip)
        if setup.get('filter_query'):
            fq = setup['filter_query']
            f_chip = QLabel(f"Filter: {fq if len(fq) <= 40 else fq[:37] + '...'}")
            f_chip.setToolTip(fq)
            f_chip.setObjectName("queue_chip")
            f_chip.setProperty("kind", "filter")
            chips.addWidget(f_chip)
        chips.addStretch()
        body.addLayout(chips)
        h.addLayout(body, 1)
        return card

    def update_template_queue_list_widget(self):
        """อัปเดตรายการใน QListWidget ที่แสดงคิว (แสดงเป็นการ์ด)"""
        self.template_queue_list_widget.clear()
        for i, setup in enumerate(self.template_setups_for_queue):
            item = QListWidgetItem()
            widget = self._make_queue_item_widget(i + 1, setup)
            hint = widget.sizeHint()
            hint.setWidth(0)  # ให้กว้างตาม List ไม่ดัน scrollbar แนวนอน
            item.setSizeHint(hint)
            self.template_queue_list_widget.addItem(item)
            self.template_queue_list_widget.setItemWidget(item, widget)
        if not self.template_setups_for_queue:
            placeholder = QListWidgetItem("ยังไม่มีรายการในคิว — ตั้งค่าด้านซ้ายแล้วกด \"เพิ่มรายการนี้ลงในคิว\" หรือโหลด Template")
            placeholder.setFlags(Qt.ItemFlag.NoItemFlags)
            placeholder.setForeground(QColor("#9aa3ad"))
            self.template_queue_list_widget.addItem(placeholder)
        if hasattr(self, "queue_count_label"):
            self.queue_count_label.setText(f"{len(self.template_setups_for_queue)} รายการ")

    def remove_selected_from_template_queue(self):
        """ลบรายการที่เลือกออกจากคิว"""
        selected_items = self.template_queue_list_widget.selectedItems()
        if not selected_items:
            return # ไม่มีอะไรเลือก

        # ถามยืนยันก่อนลบ
        reply = QMessageBox.question(self, 'ยืนยันการลบ', 
                                     f'คุณต้องการลบ {len(selected_items)} รายการที่เลือกออกจากคิวหรือไม่?', 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.No:
            return

        indices_to_delete = [self.template_queue_list_widget.row(item) for item in selected_items]
        indices_to_delete.sort(reverse=True) # ลบจากหลังมาหน้าเพื่อไม่ให้ index เปลี่ยน

        for index in indices_to_delete:
            if 0 <= index < len(self.template_setups_for_queue):
                del self.template_setups_for_queue[index]
        
        self.update_template_queue_list_widget()
        self.status_label.setText(f"สถานะ: ลบรายการออกจากคิวแล้ว. จำนวนในคิว: {len(self.template_setups_for_queue)}")
        self.update_ui_state()

    def clear_template_queue(self, show_message=True):
        """ล้างคิวทั้งหมด"""
        if show_message and self.template_setups_for_queue:
            if QMessageBox.question(self, "ยืนยันการล้างคิว", "คุณต้องการล้างทุกรายการในคิวใช่หรือไม่?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) == QMessageBox.StandardButton.No:
                return
        self.template_setups_for_queue = []
        self.update_template_queue_list_widget()
        if show_message:
            self.status_label.setText("สถานะ: ล้างคิวเรียบร้อยแล้ว")
        self.update_ui_state()

    def load_queued_item_to_setup_ui(self, item):
        """โหลดข้อมูลของรายการที่ดับเบิลคลิกในคิวไปยังฟอร์มตั้งค่า Run"""
        row_index = self.template_queue_list_widget.row(item)
        if 0 <= row_index < len(self.template_setups_for_queue):
            setup_data = self.template_setups_for_queue[row_index]
            self.single_sheet_name_edit.setText(setup_data.get('sheet_name', ''))
            self.single_header_a1_edit.setText(setup_data.get('header_a1', ''))
            self.single_filter_edit.setText(setup_data.get('filter_query', ''))
            
            selected_vars = setup_data.get('variables_list', [])
            self.current_queue_setup_vars = selected_vars
            count = len(selected_vars)
            if count > 0:
                self.set_queue_vars_button.setText(f"ตั้งค่า/แก้ไขตัวแปร... (เลือกแล้ว {count} ตัวแปร)")
                self.set_queue_vars_button.setStyleSheet("background-color: #007bff; color: white;")
            else:
                self.set_queue_vars_button.setText("คลิกเพื่อเลือกตัวแปรที่จะรัน Correlation")
                self.set_queue_vars_button.setStyleSheet("background-color: #17a2b8; color: white;")
            
            QMessageBox.information(self, "โหลดการตั้งค่า", f"โหลดการตั้งค่า '{setup_data['sheet_name']}' มายังหน้าจอหลักเพื่อแก้ไขหรือเพิ่มลงคิว")
            self.update_ui_state()

    def save_setups_as_template(self):
        """บันทึกการตั้งค่าคิวทั้งหมดและตัวแปรเสมือนลงในไฟล์ Excel Template"""
        if not self.template_setups_for_queue:
            QMessageBox.warning(self, "คิวว่างเปล่า", "กรุณาเพิ่มรายการลงในคิวก่อนที่จะบันทึกเป็น Template")
            return
        
        # เตรียมข้อมูลสำหรับ sheet "RunSetups"
        run_setups_data = []
        for s in self.template_setups_for_queue:
            run_setups_data.append({
                'ชื่อ Sheetname': s.get('sheet_name', ''),
                'หัวคอลัม A1': s.get('header_a1', ''),
                'Filter': s.get('filter_query', ''),
                'ตัวแปรที่จะรัน Correlation': "\n".join(s.get('variables_list', []))
            })
        df_runs = pd.DataFrame(run_setups_data)

        # เตรียมข้อมูลสำหรับ sheet "StackingInstructions" (เฉพาะตัวแปรเสมือนที่ถูกใช้ในคิว)
        stacking_instructions = []
        recode_jr_rows = [] # สำหรับ sheet "Recode Justright"
        saved_virtual_vars = set() # เก็บชื่อตัวแปรเสมือนที่ถูกบันทึกไปแล้ว เพื่อป้องกันการซ้ำ

        for setup in self.template_setups_for_queue:
            for var_name in setup.get('variables_list', []):
                if var_name in self.virtual_variables and var_name not in saved_virtual_vars:
                    recipe = self.virtual_variables[var_name]
                    if recipe.get('type') == 'recode_justright':
                        recode_jr_rows.append({
                            'ข้อเดิม': (recipe.get('source_vars') or [''])[0],
                            'ข้อใหม่': var_name,
                            'Recode Scale': recipe.get('scale', 5),
                        })
                        saved_virtual_vars.add(var_name)
                        continue
                    stacking_instructions.append({
                        'NewVariableName': var_name,
                        'Label': recipe.get('label', ''),
                        'SourceVariables': "\n".join(recipe.get('source_vars', [])),
                        'Type': recipe.get('type', 'unknown') # เพิ่ม Type เพื่อระบุว่าเป็น single หรือ pattern stack
                    })
                    saved_virtual_vars.add(var_name)
        df_stacks = pd.DataFrame(stacking_instructions)
        df_recode_jr = pd.DataFrame(recode_jr_rows)

        file_name, _ = QFileDialog.getSaveFileName(self, "บันทึก Template", "", "Excel Files (*.xlsx)")
        if file_name:
            try:
                with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
                    df_runs.to_excel(writer, index=False, sheet_name="RunSetups")
                    if not df_stacks.empty:
                        df_stacks.to_excel(writer, index=False, sheet_name="StackingInstructions")
                    if not df_recode_jr.empty:
                        df_recode_jr.to_excel(writer, index=False, sheet_name="Recode Justright")
                QMessageBox.information(self, "บันทึกสำเร็จ", f"บันทึก Template ไปยัง {file_name} เรียบร้อยแล้ว")
            except Exception as e:
                QMessageBox.critical(self, "เกิดข้อผิดพลาด", f"ไม่สามารถบันทึก Template ได้: {e}\nโปรดตรวจสอบว่าไฟล์ไม่ได้ถูกเปิดอยู่หรือมีสิทธิ์ในการเขียน")
        self.update_ui_state()

    def load_setups_from_template(self):
        """โหลด Template จากไฟล์ Excel และนำเข้าสู่คิวและตัวแปรเสมือน"""
        if self.df is None:
            QMessageBox.warning(self, "ไม่มีข้อมูล", "กรุณาโหลดไฟล์ SPSS ก่อนที่จะโหลด Template")
            return
        
        file_name, _ = QFileDialog.getOpenFileName(self, "โหลด Template", "", "Excel Files (*.xlsx)")
        if file_name:
            try:
                all_sheets = pd.read_excel(file_name, sheet_name=None)
                loaded_configs = []
                loaded_virtual_vars_count = 0

                # 1. โหลด Stacking Instructions ก่อน (ถ้ามี)
                if 'StackingInstructions' in all_sheets:
                    df_stacks = all_sheets['StackingInstructions']
                    for _, row in df_stacks.iterrows():
                        new_name = str(row.get('NewVariableName', '')).strip()
                        label = str(row.get('Label', '')).strip()
                        source_vars_text = str(row.get('SourceVariables', ''))
                        source_vars = [v.strip() for v in source_vars_text.splitlines() if v.strip()]
                        var_type = str(row.get('Type', 'single_stack')).strip() # ดึง Type มาด้วย

                        if not new_name or not source_vars:
                            continue # ข้ามถ้าข้อมูลไม่ครบ

                        # ตรวจสอบว่าชื่อตัวแปรเสมือนนี้ยังไม่ถูกสร้าง
                        if new_name not in self.virtual_variables:
                            # ตรวจสอบว่า source_vars ทั้งหมดมีอยู่ใน DataFrame หรือไม่
                            invalid_sources = [col for col in source_vars if col not in self.df.columns]
                            if not invalid_sources: # ถ้า source variables ถูกต้อง
                                self.virtual_variables[new_name] = {"label": label, "source_vars": source_vars, "type": var_type}
                                loaded_virtual_vars_count += 1
                            else:
                                print(f"Warning: ข้ามการสร้างตัวแปรเสมือน '{new_name}' เนื่องจาก source variable(s) ไม่ถูกต้อง: {', '.join(invalid_sources)}")
                        else:
                            print(f"Info: ตัวแปรเสมือน '{new_name}' มีอยู่แล้ว จึงไม่ได้โหลดซ้ำจาก Template")
                
                if loaded_virtual_vars_count > 0:
                    QMessageBox.information(self, "สร้างตัวแปรต่อData", f"สร้างตัวแปร ต่อData จาก Template จำนวน {loaded_virtual_vars_count} ตัวแปรเรียบร้อยแล้ว")

                # 1.5 โหลด Recode Justright (ถ้ามี)
                loaded_recode_jr_count = 0
                if 'Recode Justright' in all_sheets:
                    df_rjr = all_sheets['Recode Justright']
                    for _, row in df_rjr.iterrows():
                        src = str(row.get('ข้อเดิม', '')).strip()
                        new_name = str(row.get('ข้อใหม่', '')).strip()
                        try:
                            scale = int(float(row.get('Recode Scale', 5)))
                        except Exception:
                            scale = 5
                        if scale not in (5, 7):
                            scale = 5
                        if not src or not new_name or src == 'nan' or new_name == 'nan':
                            continue
                        if new_name in self.virtual_variables or new_name in self.df.columns:
                            print(f"Info: ตัวแปร '{new_name}' มีอยู่แล้ว จึงไม่ได้โหลดซ้ำจาก Template")
                            continue
                        if src not in self.df.columns:
                            print(f"Warning: ข้าม Recode Justright '{new_name}' เนื่องจากไม่พบตัวแปรเดิม '{src}'")
                            continue
                        base_label = self.meta.column_names_to_labels.get(src, src) if self.meta and self.meta.column_names_to_labels else src
                        self.virtual_variables[new_name] = {
                            "label": f"{base_label} (Justright {scale} Scale)",
                            "source_vars": [src],
                            "type": "recode_justright",
                            "scale": scale,
                        }
                        loaded_recode_jr_count += 1
                if loaded_recode_jr_count > 0:
                    QMessageBox.information(self, "สร้างตัวแปร Recode Justright", f"สร้างตัวแปร Recode Justright จาก Template จำนวน {loaded_recode_jr_count} ตัวแปรเรียบร้อยแล้ว")


                # 2. โหลด Run Setups
                if 'RunSetups' not in all_sheets:
                    QMessageBox.critical(self, "ไฟล์ไม่ถูกต้อง", "ไม่พบ Sheet 'RunSetups' ในไฟล์ Template")
                    return
                
                df_runs = all_sheets['RunSetups']
                for _, row in df_runs.iterrows():
                    sheet_name = str(row.get('ชื่อ Sheetname', '')).strip()
                    header_a1 = str(row.get('หัวคอลัม A1', '')).strip()
                    raw_filter = row.get('Filter')
                    filter_query = '' if pd.isna(raw_filter) else str(raw_filter).strip()
                    vars_text = str(row.get('ตัวแปรที่จะรัน Correlation', ''))
                    variables_list = [v.strip() for v in vars_text.splitlines() if v.strip()]

                    # ตรวจสอบว่าตัวแปรทั้งหมด (จริงหรือเสมือน) มีอยู่ในระบบ
                    all_current_vars = set(self.df.columns.tolist() + list(self.virtual_variables.keys()))
                    invalid_vars_in_setup = [v for v in variables_list if v not in all_current_vars]
                    if invalid_vars_in_setup:
                        print(f"Warning: ข้าม Run '{sheet_name}' เนื่องจากตัวแปรไม่ถูกต้อง: {', '.join(invalid_vars_in_setup)}")
                        continue # ข้าม Run นี้ถ้ามีตัวแปรที่ไม่ถูกต้อง

                    loaded_configs.append({
                        'sheet_name': sheet_name,
                        'header_a1': header_a1,
                        'filter_query': filter_query,
                        'variables_list': variables_list
                    })
                
                if not loaded_configs:
                    QMessageBox.information(self, "Template ว่างเปล่า", "ไม่พบการตั้งค่า Run ในไฟล์ Excel ที่เลือก")
                    return
                
                # ถามผู้ใช้ว่าจะเพิ่มต่อท้ายหรือแทนที่
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("ยืนยันการโหลด")
                msg_box.setText("คุณต้องการเพิ่มรายการที่โหลดมาต่อท้ายคิวเดิม หรือแทนที่คิวเดิมทั้งหมด?")
                add_button = msg_box.addButton("เพิ่มต่อท้าย (Add)", QMessageBox.ButtonRole.AcceptRole)
                replace_button = msg_box.addButton("แทนที่ (Replace)", QMessageBox.ButtonRole.DestructiveRole)
                msg_box.addButton("ยกเลิก (Cancel)", QMessageBox.ButtonRole.RejectRole)
                msg_box.exec()
                
                if msg_box.clickedButton() == add_button:
                    self.template_setups_for_queue.extend(loaded_configs)
                    QMessageBox.information(self, "โหลดสำเร็จ", f"เพิ่มการตั้งค่า Run จำนวน {len(loaded_configs)} รายการต่อท้ายคิวเรียบร้อยแล้ว")
                elif msg_box.clickedButton() == replace_button:
                    self.template_setups_for_queue = loaded_configs
                    QMessageBox.information(self, "โหลดสำเร็จ", f"แทนที่คิวด้วยการตั้งค่า Run จำนวน {len(loaded_configs)} รายการเรียบร้อยแล้ว")
                else:
                    return # ผู้ใช้ยกเลิก

                self.update_template_queue_list_widget()
                self.status_label.setText(f"สถานะ: โหลด {len(loaded_configs)} รายการ. จำนวนในคิวทั้งหมด: {len(self.template_setups_for_queue)}")
            except Exception as e:
                QMessageBox.critical(self, "เกิดข้อผิดพลาด", f"ไม่สามารถโหลด Template ได้: {e}\nโปรดตรวจสอบว่าไฟล์เป็นรูปแบบ Excel ที่ถูกต้องและไม่เสียหาย")
            finally:
                self.update_ui_state()

    def execute_all_queued_setups(self):
        """รันทุกรายการในคิวและ Export ผลลัพธ์ไปยังไฟล์ Excel"""
        if not self.template_setups_for_queue:
            QMessageBox.warning(self, "คิวว่างเปล่า", "ไม่มีรายการในคิวให้รัน")
            return
        if self.df is None:
            QMessageBox.warning(self, "ไม่มีข้อมูล", "กรุณาโหลดไฟล์ SPSS ก่อนที่จะรัน")
            return

        reply = QMessageBox.question(self, 'ยืนยันการรันและ Export',
                                     f"มี {len(self.template_setups_for_queue)} รายการในคิว\nคุณต้องการรันและ Export ผลลัพธ์เป็นไฟล์ Excel เลยใช่หรือไม่?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.No:
            self.status_label.setText("สถานะ: ผู้ใช้ยกเลิกการรัน")
            return

        file_name, _ = QFileDialog.getSaveFileName(self, "บันทึกผลลัพธ์ทั้งหมด", "", "Excel Files (*.xlsx)")
        if not file_name:
            self.status_label.setText("สถานะ: ผู้ใช้ยกเลิกการเลือกไฟล์สำหรับบันทึก")
            return

        total_runs = len(self.template_setups_for_queue)
        executed_count = 0
        skipped_count = 0
        error_count = 0

        # เก็บข้อมูลสำหรับ Index Sheet
        index_sheet_data = []

        try:
            with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
                workbook = writer.book

                # 1. สร้าง Index Sheet ก่อน
                index_ws = workbook.add_worksheet('Index')
                # *** ส่วนที่เพิ่มเข้ามา: Freeze Panes ที่ B3 สำหรับ Index Sheet ***
                index_ws.freeze_panes(2, 1) # Freeze ที่ Row 3, Column B (0-indexed: row=2, col=1)
                
                # กำหนด Format สำหรับ Index Sheet
                fmt_index_title = workbook.add_format({'bold': True, 'font_size': 14, 'font_name': 'Tahoma', 'color': '#2c3e50'})
                fmt_index_header = workbook.add_format({'bold': True, 'font_size': 10, 'font_name': 'Tahoma', 'bg_color': '#e9ecef', 'border': 1, 'align': 'center'})
                fmt_index_no = workbook.add_format({'font_size': 9, 'font_name': 'Tahoma', 'align': 'center', 'border': 1})
                fmt_index_link = workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_name': 'Tahoma', 'font_size': 9, 'border': 1})

                index_ws.write_string(0, 0, "Index Page for Correlation Results", fmt_index_title)
                index_ws.write_string(1, 0, "No.", fmt_index_header)
                index_ws.write_string(1, 1, "Sheet Description", fmt_index_header)
                index_ws.set_column(0, 0, 5) # ความกว้างสำหรับคอลัมน์ "No."
                index_ws.set_column(1, 1, 60) # ความกว้างสำหรับคอลัมน์ "Sheet Description"
                
                # 2. วนลูปเพื่อสร้างแต่ละ Sheet Correlation
                for i, config in enumerate(self.template_setups_for_queue):
                    sheet_name = config.get('sheet_name', f"Run_{i+1}")
                    header_a1_text = config.get('header_a1', sheet_name)

                    # เพิ่มข้อมูลสำหรับ Index Sheet
                    index_sheet_data.append({'sheet_name': sheet_name, 'header_a1': header_a1_text})

                    self.status_label.setText(f"กำลังประมวลผล {i+1}/{total_runs}: {sheet_name}...")
                    QApplication.processEvents() # ทำให้ UI อัปเดตสถานะระหว่างประมวลผล

                    selected_vars_list = config.get('variables_list', [])
                    if len(selected_vars_list) < 2:
                        print(f"ข้าม '{sheet_name}' เนื่องจากมีตัวแปรน้อยกว่า 2 ตัว")
                        skipped_count += 1
                        continue

                    # 1. กำหนด Filter สำหรับ Run นี้ (ถ้ามี filter เฉพาะจะใช้ตัวนั้น ถ้าไม่มีจะใช้ Global Filter)
                    filter_query_for_this_run = config.get('filter_query', '')
                    actual_filter = filter_query_for_this_run or self.active_filter_query
                    eff_filter_desc = actual_filter or "ไม่มี"

                    # 2. สร้าง DataFrame สำหรับ Run นี้โดยการ apply filter
                    df_for_run = self.df.copy()
                    if actual_filter:
                        try:
                            df_for_run = df_for_run.query(actual_filter)
                        except Exception as e_filter:
                            print(f"ข้าม '{sheet_name}' เนื่องจาก Filter error: {e_filter}")
                            error_count += 1
                            continue
                        
                    if df_for_run.empty:
                        print(f"ข้าม '{sheet_name}' เนื่องจาก Filter ทำให้ไม่มีข้อมูลเหลืออยู่")
                        skipped_count += 1
                        continue

                    # 3. สร้างข้อมูลสำหรับ Correlation โดยรวบรวมตัวแปรจริงและตัวแปรเสมือน
                    data_for_corr = {}
                    has_missing_var_in_source = False # Flag เพื่อตรวจสอบว่ามี source variable ที่ไม่พบหรือไม่
                    for var_name in selected_vars_list:
                        if var_name in self.virtual_variables:
                            # ถ้าเป็นตัวแปรเสมือน ให้สร้าง stacked series ตอนนี้ โดยใช้ df_for_run ที่ถูก filter แล้ว
                            recipe = self.virtual_variables[var_name]
                            source_vars = recipe.get('source_vars', [])
                            
                            # ตรวจสอบว่า source_vars ทั้งหมดมีอยู่ใน df_for_run.columns หรือไม่
                            missing_source_cols = [col for col in source_vars if col not in df_for_run.columns]
                            if missing_source_cols:
                                print(f"ข้ามตัวแปร '{var_name}' ในชีท '{sheet_name}' เนื่องจากไม่พบ source column(s): {', '.join(missing_source_cols)}")
                                has_missing_var_in_source = True
                                break # หยุดการประมวลผลตัวแปรใน run นี้
                            
                            try:
                                if recipe.get('type') == 'recode_justright':
                                    # Recode Justright จาก DataFrame ที่ถูก filter แล้ว (ไม่ต่อ Data)
                                    data_for_corr[var_name] = recode_justright_series(df_for_run[source_vars[0]], recipe.get('scale', 5)).reset_index(drop=True)
                                else:
                                    # Stack columns จาก DataFrame ที่ถูก filter แล้ว
                                    stacked_series = pd.concat([df_for_run[col] for col in source_vars], ignore_index=True)
                                    data_for_corr[var_name] = stacked_series
                            except Exception as e_stack:
                                # --- จุดที่แก้ไข: เปลี่ยน "empowers" เป็น "ตัวแปร" และ new_name เป็น var_name ---
                                print(f"ข้ามตัวแปร '{var_name}' ในชีท '{sheet_name}' เนื่องจากสร้าง stacked variable ไม่สำเร็จ: {e_stack}")
                                has_missing_var_in_source = True
                                break
                        elif var_name in df_for_run.columns:
                            # ถ้าเป็นตัวแปรปกติ ให้ดึงจาก DataFrame ที่ถูก filter แล้ว
                            data_for_corr[var_name] = df_for_run[var_name]
                        else:
                            print(f"ข้ามตัวแแปร '{var_name}' ในชีท '{sheet_name}' เนื่องจากไม่พบในข้อมูล (หลังจาก filter แล้ว)")
                            has_missing_var_in_source = True
                            break # หยุดการประมวลผลตัวแปรใน run นี้
                    
                    if has_missing_var_in_source:
                        skipped_count += 1
                        continue # ข้าม Run นี้ไป

                    # สร้าง DataFrame สำหรับการแปลงประเภทข้อมูลและการคำนวณ
                    subset_df = pd.DataFrame(data_for_corr)
                    
                    # แปลงข้อมูลเป็นตัวเลข (ถ้าทำได้) โดยที่ค่าที่ไม่สามารถแปลงได้จะเป็น NaN
                    for col in subset_df.columns:
                        subset_df[col] = pd.to_numeric(subset_df[col], errors='coerce')
                    
                    # สร้าง label map สำหรับ Excel header (สำหรับตัวแปรทั้งหมดที่เลือก)
                    label_map = {}
                    for var in selected_vars_list:
                        if var in self.virtual_variables:
                            label_map[var] = self.virtual_variables[var]["label"]
                        elif self.meta and self.meta.column_names_to_labels and var in self.meta.column_names_to_labels:
                            label_map[var] = self.meta.column_names_to_labels[var]
                        else:
                            label_map[var] = var # ใช้ชื่อตัวแปรหากไม่มี label

                    run_excel_labels = [label_map[var] for var in selected_vars_list]

                    # Initialize correlation results and N counts for ALL selected variables with zeros
                    current_corr_results = pd.DataFrame(0.0, index=selected_vars_list, columns=selected_vars_list)
                    current_n_obs_per_variable = pd.Series(0, index=selected_vars_list, dtype=int)

                    # คำนวณ N สำหรับทุกตัวแปรที่เลือก (แม้จะไม่มีข้อมูลก็ตาม)
                    for var_name in selected_vars_list:
                        if var_name in subset_df.columns:
                            current_n_obs_per_variable.loc[var_name] = subset_df[var_name].notna().sum()

                    # ระบุคอลัมน์ที่เป็นตัวเลขและมีข้อมูล (ไม่ใช่ NaN ทั้งหมด) เพื่อใช้คำนวณ Correlation
                    numeric_and_not_all_nan_cols = [col for col in subset_df.columns 
                                                    if subset_df[col].dtype in [np.int64, np.float64] # ตรวจสอบว่าเป็นชนิดตัวเลข
                                                    and subset_df[col].notna().any()] # และมีค่าที่ไม่ใช่ NaN อย่างน้อยหนึ่งค่า

                    if len(numeric_and_not_all_nan_cols) >= 2:
                        # ถ้ามีตัวแปรที่เป็นตัวเลขและมีข้อมูลอย่างน้อย 2 ตัว ให้คำนวณ Correlation จริง
                        numeric_df_for_corr_calc = subset_df[numeric_and_not_all_nan_cols]
                        
                        # คำนวณค่า Correlation จริงสำหรับชุดตัวแปรที่ใช้งานได้
                        temp_actual_corr = numeric_df_for_corr_calc.corr(method='pearson')
                        
                        # นำค่า Correlation ที่คำนวณได้จริงมาใส่ใน current_corr_results ที่ initialized เป็น 0.0 ไว้แล้ว
                        for row_var in temp_actual_corr.index:
                            for col_var in temp_actual_corr.columns:
                                current_corr_results.loc[row_var, col_var] = temp_actual_corr.loc[row_var, col_var]
                                
                        # ตั้งค่า Diagonal (Correlation ตัวเอง) ให้เป็น 1.0 หากมีข้อมูล ไม่เช่นนั้นเป็น 0.0 (ตามค่า initialized)
                        for var in selected_vars_list:
                            if current_n_obs_per_variable.loc[var] > 0:
                                current_corr_results.loc[var, var] = 1.0
                            # ถ้า N เป็น 0 มันจะยังคงเป็น 0.0 ตามที่ initialized ไว้แล้ว
                    else:
                        # ถ้ามีตัวแปรที่เป็นตัวเลขและมีข้อมูลน้อยกว่า 2 ตัว ไม่สามารถคำนวณ Correlation ได้
                        # Correlation Matrix จะยังคงเป็น 0.0 ทั้งหมดตามค่า initialized ไว้
                        print(f"ข้ามการคำนวณ Correlation ที่แท้จริงสำหรับ '{sheet_name}' เนื่องจากมีตัวแปรที่เป็นตัวเลขและมีข้อมูลน้อยกว่า 2 ตัว")
                    
                    # เขียนผลลัพธ์ลง Excel Sheet
                    ws = writer.book.add_worksheet(sheet_name)
                    
                    # *** ส่วนที่เพิ่มเข้ามา: Freeze Panes ที่ C5 ***
                    ws.freeze_panes(4, 2) # Freeze ที่ Row 5, Column C (0-indexed: row=4, col=2)
                    
                    # กำหนด Format ของ Cell
                    fmt_back_link = workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_name': 'Tahoma', 'font_size': 9, 'align': 'left'})
                    fmt_title = workbook.add_format({'bold': True, 'font_size': 12, 'font_name': 'Tahoma', 'color': '#2c3e50'})
                    fmt_filter = workbook.add_format({'italic': True, 'font_size': 9, 'font_name': 'Tahoma', 'color': '#555555'})
                    fmt_col_h = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1, 'bold': True, 'bg_color': '#dee2e6', 'font_name': 'Tahoma'}) # สีพื้นหลัง header
                    fmt_row_h = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#f8f9fa', 'font_name': 'Tahoma'}) # สีพื้นหลัง row header
                    fmt_n = workbook.add_format({'align': 'center', 'border': 1, 'font_name': 'Tahoma', 'font_size': 8, 'color': '#6c757d'}) # ขนาด font เล็กลงและสีเทาสำหรับ N
                    fmt_corr = workbook.add_format({'num_format': '0.000', 'align': 'right', 'border': 1, 'font_name': 'Tahoma'})
                    fmt_diag = workbook.add_format({'num_format': '0.000', 'align': 'right', 'border': 1, 'bold': True, 'bg_color': '#e9ecef', 'font_name': 'Tahoma'}) # สีพื้นหลังสำหรับ Diagonal

                    # เพิ่ม Hyperlink กลับไปยัง Index Sheet ที่ A1
                    ws.write_url(0, 0, "internal:'Index'!A1", fmt_back_link, "<< กลับไปยัง Index")

                    # เขียนหัวข้อและ Filter (เลื่อนลงมา 1 แถว)
                    ws.write_string(1, 0, header_a1_text, fmt_title)
                    ws.write_string(1, 1, f"Filter: {eff_filter_desc}", fmt_filter)
                    
                    # กำหนดความสูงของแถว Header และความกว้างคอลัมน์ (เลื่อนลงมา 1 แถว)
                    ws.set_row(2, 60) # แถวหัวคอลัมน์สูง 60 (เดิมคือแถวที่ 1)
                    ws.set_column(0, 0, 35) # คอลัมน์ Variable Label (ยังคงเป็นคอลัมน์แรก)
                    ws.set_column(1, 1, 12) # คอลัมน์ N (ยังคงเป็นคอลัมน์ที่สอง)

                    # เขียน Column Headers (Label ของตัวแปร)
                    for c_idx, header_text in enumerate(run_excel_labels):
                        ws.write_string(2, c_idx + 2, str(header_text), fmt_col_h) # แถวที่ 2 (เดิมคือแถวที่ 1)
                    
                    # เขียน Row Headers (Label ของตัวแปร) และ N
                    for r_idx, header_text in enumerate(run_excel_labels):
                        ws.write_string(r_idx + 3, 0, str(header_text), fmt_row_h) # แถวที่ r_idx + 3 (เดิมคือ r_idx + 2)
                        n_val = current_n_obs_per_variable.get(selected_vars_list[r_idx], 'N/A') # ดึง N จากชื่อตัวแปรโดยตรง
                        ws.write_string(r_idx + 3, 1, f"(N={n_val})", fmt_n) # แถวที่ r_idx + 3 (เดิมคือ r_idx + 2)

                    # เขียนค่า Correlation
                    for r_idx in range(len(run_excel_labels)):
                        for c_idx in range(len(run_excel_labels)):
                            cell_value = current_corr_results.iloc[r_idx, c_idx]
                            current_format = fmt_diag if r_idx == c_idx else fmt_corr
                            
                            # เขียน 0.0 หากเป็น NaN หรือค่าอื่นที่ต้องการให้แสดงเป็น 0
                            if pd.isna(cell_value):
                                ws.write_number(r_idx + 3, c_idx + 2, 0.0, current_format) # แถวที่ r_idx + 3 (เดิมคือ r_idx + 2)
                            else:
                                ws.write_number(r_idx + 3, c_idx + 2, cell_value, current_format)
                    
                    # กำหนดความกว้างของคอลัมน์ข้อมูล Correlation
                    for c_idx_data in range(len(run_excel_labels)):
                        ws.set_column(c_idx_data + 2, c_idx_data + 2, 12) # ความกว้าง 12 สำหรับคอลัมน์ข้อมูล

                    executed_count += 1
                
                # 3. ใส่ข้อมูลลงใน Index Sheet หลังจากสร้างทุก Correlation Sheet แล้ว
                for row_num, info in enumerate(index_sheet_data):
                    index_ws.write_number(row_num + 2, 0, row_num + 1, fmt_index_no)
                    # ลิงก์ไปยัง A2 ของแต่ละชีท (เนื่องจาก A1 เป็นลิงก์กลับไป Index)
                    index_ws.write_url(row_num + 2, 1, f"internal:'{info['sheet_name']}'!A2", fmt_index_link, info['header_a1'])

            QMessageBox.information(self, "สำเร็จ", f"Export ผลลัพธ์จำนวน {executed_count} รายการ (จากทั้งหมด {total_runs} รายการ) ไปยังไฟล์ Excel เรียบร้อยแล้ว\nข้ามไป {skipped_count} รายการ, พบข้อผิดพลาด {error_count} รายการ")
            self.status_label.setText(f"สถานะ: Export สำเร็จ {executed_count}/{total_runs} รายการ. คิวถูกล้างแล้ว")
            
            self.clear_template_queue(show_message=False) # ล้างคิวหลังจากรันสำเร็จ

        except ImportError:
            QMessageBox.warning(self, "ต้องการ XlsxWriter", "กรุณาติดตั้ง 'xlsxwriter' โดยใช้คำสั่ง: pip install xlsxwriter")
        except Exception as e:
            QMessageBox.critical(self, "เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกผลลัพธ์ได้: {e}\nโปรดตรวจสอบว่าไฟล์ไม่ได้ถูกเปิดอยู่หรือมีสิทธิ์ในการเขียน")
        finally:
            self.update_ui_state()

def main():
    app = QApplication(sys.argv)
    window = SPSSCorrelationApp()
    window.show() # เรียก show() ที่นี่ เพื่อให้หน้าต่างแสดงผล
    sys.exit(app.exec())




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
        main()

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





