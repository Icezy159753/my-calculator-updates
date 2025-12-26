import os
import sys
from datetime import datetime

import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QFont, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


def normalize_key(value):
    return str(value).strip().lower()


def normalize_code(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip().lower()


def normalize_label(value):
    if value is None:
        return ""
    return " ".join(str(value).split()).strip().lower()


def format_value_labels(value_labels):
    if not value_labels:
        return "-"
    items = []
    for code, label in value_labels.items():
        if isinstance(code, float) and code.is_integer():
            code_text = str(int(code))
        else:
            code_text = str(code)
        label_text = "" if label is None else str(label)
        items.append((normalize_code(code), f"{code_text}={label_text}"))
    items.sort(key=lambda entry: entry[0])
    return "; ".join(item[1] for item in items)


class MapDataWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Smart Map Data")
        self.resize(1100, 720)

        icon_path = os.path.join(os.path.dirname(__file__), "I_Main.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.excel_source_path = ""
        self.excel_target_path = ""
        self.excel_last_result = None

        self.spss_source_path = ""
        self.spss_target_path = ""
        self.spss_last_result = None

        self._build_ui()
        self._apply_styles()

    def _build_ui(self):
        container = QWidget()
        root_layout = QVBoxLayout(container)
        root_layout.setContentsMargins(18, 18, 18, 18)
        root_layout.setSpacing(14)

        header = QFrame()
        header.setObjectName("HeaderFrame")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(18, 16, 18, 16)
        header_title = QLabel("Map Data อัจฉริยะ")
        header_title.setFont(QFont("Leelawadee UI", 18, QFont.Weight.Bold))
        header_sub = QLabel("Map Excel -> Excel และ SPSS -> SPSS พร้อมแจ้งเตือน + Log")
        header_sub.setFont(QFont("Leelawadee UI", 10))
        header_layout.addWidget(header_title)
        header_layout.addWidget(header_sub)
        root_layout.addWidget(header)

        tabs = QTabWidget()
        tabs.addTab(self._build_excel_tab(), "Excel -> Excel")
        tabs.addTab(self._build_spss_tab(), "SPSS -> SPSS")
        root_layout.addWidget(tabs)

        self.setCentralWidget(container)

    def _build_excel_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(12)

        file_group = QGroupBox("โหลดไฟล์ Excel")
        file_layout = QVBoxLayout(file_group)

        self.excel_source_input = QLineEdit()
        self.excel_source_input.setReadOnly(True)
        self.excel_source_sheet = QComboBox()
        source_row = self._build_file_row(
            "ไฟล์ Excel ต้นทาง", self.excel_source_input, "เลือกไฟล์", self._pick_excel_source
        )

        self.excel_target_input = QLineEdit()
        self.excel_target_input.setReadOnly(True)
        self.excel_target_sheet = QComboBox()
        target_row = self._build_file_row(
            "ไฟล์ Excel ปลายทาง", self.excel_target_input, "เลือกไฟล์", self._pick_excel_target
        )

        file_layout.addLayout(source_row)
        file_layout.addLayout(self._build_sheet_row("Sheet ต้นทาง", self.excel_source_sheet))
        file_layout.addSpacing(6)
        file_layout.addLayout(target_row)
        file_layout.addLayout(self._build_sheet_row("Sheet ปลายทาง", self.excel_target_sheet))

        layout.addWidget(file_group)

        action_group = QGroupBox("ดำเนินการ")
        action_layout = QHBoxLayout(action_group)
        self.excel_map_button = QPushButton("เริ่ม Map คอลัมน์")
        self.excel_map_button.clicked.connect(self.analyze_excel)
        self.excel_save_log_button = QPushButton("บันทึก Log")
        self.excel_save_log_button.clicked.connect(self.save_excel_log)
        self.excel_save_mismatch_button = QPushButton("บันทึกที่ไม่ตรง")
        self.excel_save_mismatch_button.clicked.connect(self.save_excel_mismatch)
        self.excel_clear_button = QPushButton("ล้างผลลัพธ์")
        self.excel_clear_button.clicked.connect(self.clear_excel_results)

        action_layout.addWidget(self.excel_map_button)
        action_layout.addWidget(self.excel_save_log_button)
        action_layout.addWidget(self.excel_save_mismatch_button)
        action_layout.addStretch(1)
        action_layout.addWidget(self.excel_clear_button)

        layout.addWidget(action_group)

        self.excel_summary = QLabel("พร้อมใช้งาน")
        layout.addWidget(self.excel_summary)

        self.excel_table = QTableWidget(0, 4)
        self.excel_table.setHorizontalHeaderLabels([
            "สถานะ",
            "คอลัมน์ต้นทาง",
            "คอลัมน์ปลายทาง",
            "หมายเหตุ",
        ])
        self.excel_table.horizontalHeader().setStretchLastSection(True)
        self.excel_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(self.excel_table)

        self.excel_log = QTextEdit()
        self.excel_log.setReadOnly(True)
        self.excel_log.setPlaceholderText("Log จะถูกแสดงที่นี่")
        layout.addWidget(self.excel_log)

        return tab

    def _build_spss_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(12)

        file_group = QGroupBox("โหลดไฟล์ SPSS")
        file_layout = QVBoxLayout(file_group)

        self.spss_source_input = QLineEdit()
        self.spss_source_input.setReadOnly(True)
        source_row = self._build_file_row(
            "ไฟล์ SPSS ต้นทาง", self.spss_source_input, "เลือกไฟล์", self._pick_spss_source
        )

        self.spss_target_input = QLineEdit()
        self.spss_target_input.setReadOnly(True)
        target_row = self._build_file_row(
            "ไฟล์ SPSS ปลายทาง", self.spss_target_input, "เลือกไฟล์", self._pick_spss_target
        )

        file_layout.addLayout(source_row)
        file_layout.addLayout(target_row)
        file_layout.addLayout(self._build_spss_encoding_row())
        layout.addWidget(file_group)

        action_group = QGroupBox("ดำเนินการ")
        action_layout = QHBoxLayout(action_group)
        self.spss_map_button = QPushButton("เริ่ม Map ตัวแปร")
        self.spss_map_button.clicked.connect(self.analyze_spss)
        self.spss_save_log_button = QPushButton("บันทึก Log")
        self.spss_save_log_button.clicked.connect(self.save_spss_log)
        self.spss_save_mismatch_button = QPushButton("บันทึกที่ไม่ตรง")
        self.spss_save_mismatch_button.clicked.connect(self.save_spss_mismatch)
        self.spss_clear_button = QPushButton("ล้างผลลัพธ์")
        self.spss_clear_button.clicked.connect(self.clear_spss_results)

        action_layout.addWidget(self.spss_map_button)
        action_layout.addWidget(self.spss_save_log_button)
        action_layout.addWidget(self.spss_save_mismatch_button)
        action_layout.addStretch(1)
        action_layout.addWidget(self.spss_clear_button)

        layout.addWidget(action_group)

        self.spss_summary = QLabel("พร้อมใช้งาน")
        layout.addWidget(self.spss_summary)

        self.spss_table = QTableWidget(0, 7)
        self.spss_table.setHorizontalHeaderLabels([
            "สถานะ",
            "ตัวแปรต้นทาง",
            "Label ต้นทาง",
            "Label ปลายทาง",
            "Value Label ต้นทาง",
            "Value Label ปลายทาง",
            "หมายเหตุ",
        ])
        self.spss_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.spss_table)

        self.spss_log = QTextEdit()
        self.spss_log.setReadOnly(True)
        self.spss_log.setPlaceholderText("Log จะถูกแสดงที่นี่")
        layout.addWidget(self.spss_log)

        return tab

    def _build_file_row(self, label_text, line_edit, button_text, callback):
        row = QHBoxLayout()
        label = QLabel(label_text)
        label.setMinimumWidth(140)
        button = QPushButton(button_text)
        button.clicked.connect(callback)
        row.addWidget(label)
        row.addWidget(line_edit, 1)
        row.addWidget(button)
        return row

    def _build_sheet_row(self, label_text, combo_box):
        row = QHBoxLayout()
        label = QLabel(label_text)
        label.setMinimumWidth(140)
        row.addWidget(label)
        row.addWidget(combo_box, 1)
        return row

    def _build_spss_encoding_row(self):
        row = QHBoxLayout()
        label = QLabel("Encoding SPSS")
        label.setMinimumWidth(140)
        self.spss_encoding = QComboBox()
        self.spss_encoding.addItems(["UTF-8", "CP874"])
        self.spss_encoding.setCurrentText("UTF-8")
        row.addWidget(label)
        row.addWidget(self.spss_encoding, 1)
        return row

    def _current_spss_encoding(self):
        value = self.spss_encoding.currentText().strip()
        return value or "UTF-8"

    def _apply_styles(self):
        self.setStyleSheet(
            """
            QWidget {
                background: #f4f7fb;
                color: #1f2933;
                font-family: "Leelawadee UI";
                font-size: 10pt;
            }
            #HeaderFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #e0ecff, stop:1 #f8fbff);
                border: 1px solid #d0ddee;
                border-radius: 14px;
            }
            QGroupBox {
                background: #ffffff;
                border: 1px solid #dfe5ef;
                border-radius: 12px;
                margin-top: 10px;
                padding: 12px;
            }
            QGroupBox:title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                font-weight: 600;
                color: #2b2b2b;
            }
            QLineEdit, QComboBox {
                background: #f8fafc;
                border: 1px solid #d9e2ec;
                border-radius: 6px;
                padding: 4px 8px;
            }
            QPushButton {
                background: #2b6cb0;
                color: #ffffff;
                border-radius: 8px;
                padding: 6px 14px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #2c5282;
            }
            QPushButton:disabled {
                background: #a0aec0;
            }
            QTableWidget {
                background: #ffffff;
                border: 1px solid #dfe5ef;
                border-radius: 8px;
            }
            QHeaderView::section {
                background: #edf2f7;
                padding: 6px;
                border: none;
                font-weight: 600;
            }
            QTextEdit {
                background: #0f172a;
                color: #e2e8f0;
                border-radius: 8px;
                padding: 6px;
                font-family: "Leelawadee UI";
                font-size: 9pt;
            }
            """
        )

    def _pick_excel_source(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "เลือกไฟล์ Excel ต้นทาง", "", "Excel Files (*.xlsx *.xls *.xlsm)"
        )
        if not path:
            return
        self.excel_source_path = path
        self.excel_source_input.setText(path)
        self._load_excel_sheets(path, self.excel_source_sheet)

    def _pick_excel_target(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "เลือกไฟล์ Excel ปลายทาง", "", "Excel Files (*.xlsx *.xls *.xlsm)"
        )
        if not path:
            return
        self.excel_target_path = path
        self.excel_target_input.setText(path)
        self._load_excel_sheets(path, self.excel_target_sheet)

    def _load_excel_sheets(self, path, combo_box):
        try:
            excel_file = pd.ExcelFile(path)
            combo_box.clear()
            combo_box.addItems(excel_file.sheet_names)
        except Exception as exc:
            QMessageBox.critical(self, "โหลดไฟล์ไม่สำเร็จ", f"ไม่สามารถอ่านไฟล์ Excel:\n{exc}")

    def analyze_excel(self):
        if not self.excel_source_path or not self.excel_target_path:
            QMessageBox.warning(self, "ข้อมูลไม่ครบ", "กรุณาเลือกไฟล์ Excel ทั้งสองไฟล์")
            return

        source_sheet = self.excel_source_sheet.currentText()
        target_sheet = self.excel_target_sheet.currentText()
        if not source_sheet or not target_sheet:
            QMessageBox.warning(self, "ข้อมูลไม่ครบ", "กรุณาเลือก Sheet ให้ครบทั้งสองไฟล์")
            return

        try:
            source_df = pd.read_excel(self.excel_source_path, sheet_name=source_sheet)
            target_df = pd.read_excel(self.excel_target_path, sheet_name=target_sheet)
        except Exception as exc:
            QMessageBox.critical(self, "เกิดข้อผิดพลาด", f"ไม่สามารถอ่านไฟล์ Excel:\n{exc}")
            return

        matched, missing, extra = self._compare_columns(
            list(source_df.columns), list(target_df.columns)
        )

        total_matched = len(matched)
        total_missing = len(missing)
        total_extra = len(extra)
        mismatch_case = sum(1 for _, _, note in matched if note != "ตรงกันสมบูรณ์")

        self._fill_excel_table(matched, missing, extra)

        self.excel_summary.setText(
            f"Matched: {total_matched} | Missing: {total_missing} | Extra: {total_extra} | "
            f"Case-Insensitive: {mismatch_case}"
        )

        log_lines = [
            f"[{self._now()}] วิเคราะห์ Excel สำเร็จ",
            f"ไฟล์ต้นทาง: {self.excel_source_path} | Sheet: {source_sheet}",
            f"ไฟล์ปลายทาง: {self.excel_target_path} | Sheet: {target_sheet}",
            f"Matched: {total_matched}",
            f"Missing in Target: {total_missing}",
            f"Extra in Target: {total_extra}",
        ]
        if missing:
            log_lines.append("คอลัมน์ที่ไม่พบในปลายทาง: " + ", ".join(map(str, missing)))
        if extra:
            log_lines.append("คอลัมน์เกินในปลายทาง: " + ", ".join(map(str, extra)))
        if mismatch_case:
            log_lines.append("มีคอลัมน์ที่ match แบบไม่สนตัวพิมพ์ใหญ่-เล็ก")

        self.excel_log.setPlainText("\n".join(log_lines))
        self.excel_last_result = {
            "matched": matched,
            "missing": missing,
            "extra": extra,
            "source": self.excel_source_path,
            "target": self.excel_target_path,
            "source_sheet": source_sheet,
            "target_sheet": target_sheet,
            "log": log_lines,
        }

    def _compare_columns(self, source_cols, target_cols):
        target_map = {}
        for col in target_cols:
            key = normalize_key(col)
            if key not in target_map:
                target_map[key] = col

        matched = []
        missing = []
        used_target = set()

        for col in source_cols:
            if col in target_cols and col not in used_target:
                matched.append((col, col, "ตรงกันสมบูรณ์"))
                used_target.add(col)
                continue

            key = normalize_key(col)
            if key in target_map and target_map[key] not in used_target:
                matched.append((col, target_map[key], "ตรงกันแบบไม่สนตัวพิมพ์ใหญ่-เล็ก"))
                used_target.add(target_map[key])
            else:
                missing.append(col)

        extra = [col for col in target_cols if col not in used_target]
        return matched, missing, extra

    def _fill_excel_table(self, matched, missing, extra):
        rows = []
        for source_col, target_col, note in matched:
            rows.append(("Matched", source_col, target_col, note))
        for col in missing:
            rows.append(("Missing in Target", col, "-", "ไม่พบคอลัมน์ปลายทาง"))
        for col in extra:
            rows.append(("Extra in Target", "-", col, "คอลัมน์เกินในปลายทาง"))

        self._fill_table(self.excel_table, rows, {
            "Matched": QColor("#16a34a"),
            "Missing in Target": QColor("#dc2626"),
            "Extra in Target": QColor("#d97706"),
        })

    def save_excel_log(self):
        if not self.excel_last_result:
            QMessageBox.information(self, "ไม่มีข้อมูล", "กรุณา Map คอลัมน์ก่อนบันทึก Log")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "บันทึก Log Excel",
            "Map_Log_Excel.xlsx",
            "Excel Workbook (*.xlsx);;CSV (*.csv);;Text (*.txt)",
        )
        if not path:
            return

        try:
            if path.lower().endswith(".xlsx"):
                with pd.ExcelWriter(path) as writer:
                    pd.DataFrame(
                        self.excel_last_result["matched"],
                        columns=["Source", "Target", "Note"],
                    ).to_excel(writer, sheet_name="Matched", index=False)
                    pd.DataFrame(
                        self.excel_last_result["missing"], columns=["Missing in Target"]
                    ).to_excel(writer, sheet_name="Missing", index=False)
                    pd.DataFrame(
                        self.excel_last_result["extra"], columns=["Extra in Target"]
                    ).to_excel(writer, sheet_name="Extra", index=False)
            elif path.lower().endswith(".csv"):
                rows = self._build_combined_rows(self.excel_last_result)
                pd.DataFrame(rows).to_csv(path, index=False)
            else:
                with open(path, "w", encoding="utf-8") as handle:
                    handle.write("\n".join(self.excel_last_result["log"]))

            QMessageBox.information(self, "บันทึกสำเร็จ", f"บันทึก Log แล้วที่:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "บันทึกไม่สำเร็จ", f"ไม่สามารถบันทึก Log:\n{exc}")

    def save_excel_mismatch(self):
        if not self.excel_last_result:
            QMessageBox.information(self, "ไม่มีข้อมูล", "กรุณา Map คอลัมน์ก่อนบันทึก")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "บันทึกรายการที่ไม่ตรง (Excel)",
            "Map_Excel_Mismatch.xlsx",
            "Excel Workbook (*.xlsx)",
        )
        if not path:
            return

        try:
            rows = []
            for source_col, target_col, note in self.excel_last_result.get("matched", []):
                if note != "ตรงกันสมบูรณ์":
                    rows.append(("Matched (Case)", source_col, target_col, note))
            for col in self.excel_last_result.get("missing", []):
                rows.append(("Missing in Target", col, "-", "ไม่พบคอลัมน์ปลายทาง"))
            for col in self.excel_last_result.get("extra", []):
                rows.append(("Extra in Target", "-", col, "คอลัมน์เกินในปลายทาง"))

            df = pd.DataFrame(
                rows, columns=["Status", "Source Column", "Target Column", "Note"]
            )
            df.to_excel(path, index=False)
            QMessageBox.information(self, "บันทึกสำเร็จ", f"บันทึกไฟล์แล้วที่:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "บันทึกไม่สำเร็จ", f"ไม่สามารถบันทึกไฟล์:\n{exc}")

    def clear_excel_results(self):
        self.excel_table.setRowCount(0)
        self.excel_summary.setText("พร้อมใช้งาน")
        self.excel_log.clear()
        self.excel_last_result = None

    def _pick_spss_source(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "เลือกไฟล์ SPSS ต้นทาง", "", "SPSS Files (*.sav *.zsav)"
        )
        if not path:
            return
        self.spss_source_path = path
        self.spss_source_input.setText(path)

    def _pick_spss_target(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "เลือกไฟล์ SPSS ปลายทาง", "", "SPSS Files (*.sav *.zsav)"
        )
        if not path:
            return
        self.spss_target_path = path
        self.spss_target_input.setText(path)

    def analyze_spss(self):
        if not self.spss_source_path or not self.spss_target_path:
            QMessageBox.warning(self, "ข้อมูลไม่ครบ", "กรุณาเลือกไฟล์ SPSS ทั้งสองไฟล์")
            return

        try:
            import pyreadstat
        except ImportError:
            QMessageBox.critical(
                self, "ขาดไลบรารี", "ไม่พบ pyreadstat กรุณาติดตั้งก่อนใช้งาน"
            )
            return

        try:
            encoding = self._current_spss_encoding()
            _, source_meta = pyreadstat.read_sav(
                self.spss_source_path, metadataonly=True, encoding=encoding
            )
            _, target_meta = pyreadstat.read_sav(
                self.spss_target_path, metadataonly=True, encoding=encoding
            )
        except Exception as exc:
            fallback = "CP874" if encoding.upper() == "UTF-8" else "UTF-8"
            try:
                _, source_meta = pyreadstat.read_sav(
                    self.spss_source_path, metadataonly=True, encoding=fallback
                )
                _, target_meta = pyreadstat.read_sav(
                    self.spss_target_path, metadataonly=True, encoding=fallback
                )
                QMessageBox.information(
                    self,
                    "เปลี่ยน Encoding อัตโนมัติ",
                    f"อ่านไฟล์ไม่สำเร็จด้วย {encoding} จึงลอง {fallback} ให้โดยอัตโนมัติ",
                )
            except Exception as exc2:
                QMessageBox.critical(
                    self,
                    "เกิดข้อผิดพลาด",
                    f"ไม่สามารถอ่านไฟล์ SPSS:\n{exc2}",
                )
                return

        source_vars = list(source_meta.column_names)
        target_vars = list(target_meta.column_names)
        source_labels = source_meta.column_names_to_labels or {}
        target_labels = target_meta.column_names_to_labels or {}
        source_value_labels = self._get_value_labels(source_meta)
        target_value_labels = self._get_value_labels(target_meta)

        matched, missing, extra, label_mismatch = self._compare_spss(
            source_vars,
            target_vars,
            source_labels,
            target_labels,
            source_value_labels,
            target_value_labels,
        )

        self._fill_spss_table(matched, missing, extra)

        total_matched = len(matched)
        total_missing = len(missing)
        total_extra = len(extra)
        total_label_mismatch = sum(1 for row in matched if row[0] == "Label mismatch")
        total_value_mismatch = sum(1 for row in matched if "Value" in row[0])

        self.spss_summary.setText(
            f"Matched: {total_matched} | Label mismatch: {total_label_mismatch} | "
            f"Value mismatch: {total_value_mismatch} | Missing: {total_missing} | Extra: {total_extra}"
        )

        log_lines = [
            f"[{self._now()}] วิเคราะห์ SPSS สำเร็จ",
            f"ไฟล์ต้นทาง: {self.spss_source_path}",
            f"ไฟล์ปลายทาง: {self.spss_target_path}",
            f"Matched: {total_matched}",
            f"Label mismatch: {total_label_mismatch}",
            f"Value label mismatch: {total_value_mismatch}",
            f"Missing in Target: {total_missing}",
            f"Extra in Target: {total_extra}",
        ]
        if missing:
            log_lines.append("ตัวแปรที่ไม่พบในปลายทาง: " + ", ".join(missing))
        if extra:
            log_lines.append("ตัวแปรเกินในปลายทาง: " + ", ".join(extra))
        if label_mismatch:
            log_lines.append("ตรวจพบ Label ไม่ตรงกันในบางตัวแปร")

        self.spss_log.setPlainText("\n".join(log_lines))
        self.spss_last_result = {
            "matched": matched,
            "missing": missing,
            "extra": extra,
            "source": self.spss_source_path,
            "target": self.spss_target_path,
            "log": log_lines,
        }

    def _get_value_labels(self, meta):
        variable_value_labels = getattr(meta, "variable_value_labels", None)
        if variable_value_labels:
            return variable_value_labels

        value_labels = getattr(meta, "value_labels", None) or {}
        variable_to_label = getattr(meta, "variable_to_label", None) or {}
        result = {}
        for var, label_set in variable_to_label.items():
            result[var] = value_labels.get(label_set, {})
        return result

    def _compare_value_labels(self, source_map, target_map):
        source_norm = {
            normalize_code(code): normalize_label(label)
            for code, label in (source_map or {}).items()
        }
        target_norm = {
            normalize_code(code): normalize_label(label)
            for code, label in (target_map or {}).items()
        }
        missing = sorted(code for code in source_norm if code not in target_norm)
        extra = sorted(code for code in target_norm if code not in source_norm)
        diff = sorted(
            code
            for code in source_norm
            if code in target_norm and source_norm[code] != target_norm[code]
        )
        if not missing and not extra and not diff:
            return False, ""

        parts = []
        if missing:
            parts.append("ขาดรหัส: " + ", ".join(missing))
        if extra:
            parts.append("รหัสเกิน: " + ", ".join(extra))
        if diff:
            parts.append("รหัสแต่ Label ต่าง: " + ", ".join(diff))
        return True, "; ".join(parts)

    def _compare_spss(
        self,
        source_vars,
        target_vars,
        source_labels,
        target_labels,
        source_value_labels,
        target_value_labels,
    ):
        target_map = {}
        for var in target_vars:
            key = normalize_key(var)
            if key not in target_map:
                target_map[key] = var

        matched = []
        missing = []
        label_mismatch = []
        used_target = set()

        for var in source_vars:
            target_var = None
            if var in target_vars and var not in used_target:
                target_var = var
            else:
                key = normalize_key(var)
                if key in target_map and target_map[key] not in used_target:
                    target_var = target_map[key]

            if not target_var:
                missing.append(var)
                continue

            used_target.add(target_var)
            source_label = source_labels.get(var, "")
            target_label = target_labels.get(target_var, "")
            source_values = source_value_labels.get(var, {})
            target_values = target_value_labels.get(target_var, {})
            value_mismatch, value_note = self._compare_value_labels(
                source_values, target_values
            )

            label_ok = normalize_label(source_label) == normalize_label(target_label)
            value_ok = not value_mismatch

            if label_ok and value_ok:
                status = "Matched"
                note = "Label และ Value Label ตรงกัน"
            elif not label_ok and value_ok:
                status = "Label mismatch"
                note = "Label ไม่ตรงกัน"
                label_mismatch.append(var)
            elif label_ok and not value_ok:
                status = "Value label mismatch"
                note = "Value Label ไม่ตรงกัน"
            else:
                status = "Label + Value mismatch"
                note = "Label และ Value Label ไม่ตรงกัน"
                label_mismatch.append(var)

            if value_note:
                note = f"{note} ({value_note})"

            matched.append((
                status,
                var,
                source_label,
                target_label,
                format_value_labels(source_values),
                format_value_labels(target_values),
                note,
            ))

        extra = [var for var in target_vars if var not in used_target]
        return matched, missing, extra, label_mismatch

    def _fill_spss_table(self, matched, missing, extra):
        rows = []
        for (
            status,
            var,
            source_label,
            target_label,
            source_values,
            target_values,
            note,
        ) in matched:
            rows.append((
                status,
                var,
                source_label,
                target_label,
                source_values,
                target_values,
                note,
            ))
        for var in missing:
            rows.append(("Missing in Target", var, "-", "-", "-", "-", "ไม่พบตัวแปรปลายทาง"))
        for var in extra:
            rows.append(("Extra in Target", "-", "-", "-", "-", "-", "ตัวแปรเกินในปลายทาง"))

        self._fill_table(self.spss_table, rows, {
            "Matched": QColor("#16a34a"),
            "Label mismatch": QColor("#f59e0b"),
            "Value label mismatch": QColor("#0ea5e9"),
            "Label + Value mismatch": QColor("#ef4444"),
            "Missing in Target": QColor("#dc2626"),
            "Extra in Target": QColor("#d97706"),
        })

    def save_spss_log(self):
        if not self.spss_last_result:
            QMessageBox.information(self, "ไม่มีข้อมูล", "กรุณา Map ตัวแปรก่อนบันทึก Log")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "บันทึก Log SPSS",
            "Map_Log_SPSS.xlsx",
            "Excel Workbook (*.xlsx);;CSV (*.csv);;Text (*.txt)",
        )
        if not path:
            return

        try:
            if path.lower().endswith(".xlsx"):
                with pd.ExcelWriter(path) as writer:
                    pd.DataFrame(
                        self.spss_last_result["matched"],
                        columns=[
                            "Status",
                            "Variable",
                            "Source Label",
                            "Target Label",
                            "Source Value Labels",
                            "Target Value Labels",
                            "Note",
                        ],
                    ).to_excel(writer, sheet_name="Matched", index=False)
                    pd.DataFrame(
                        self.spss_last_result["missing"], columns=["Missing in Target"]
                    ).to_excel(writer, sheet_name="Missing", index=False)
                    pd.DataFrame(
                        self.spss_last_result["extra"], columns=["Extra in Target"]
                    ).to_excel(writer, sheet_name="Extra", index=False)
            elif path.lower().endswith(".csv"):
                rows = self._build_combined_rows(self.spss_last_result)
                pd.DataFrame(rows).to_csv(path, index=False)
            else:
                with open(path, "w", encoding="utf-8") as handle:
                    handle.write("\n".join(self.spss_last_result["log"]))

            QMessageBox.information(self, "บันทึกสำเร็จ", f"บันทึก Log แล้วที่:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "บันทึกไม่สำเร็จ", f"ไม่สามารถบันทึก Log:\n{exc}")

    def save_spss_mismatch(self):
        if not self.spss_last_result:
            QMessageBox.information(self, "ไม่มีข้อมูล", "กรุณา Map ตัวแปรก่อนบันทึก")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "บันทึกรายการที่ไม่ตรง (SPSS)",
            "Map_SPSS_Mismatch.xlsx",
            "Excel Workbook (*.xlsx)",
        )
        if not path:
            return

        try:
            rows = [
                entry
                for entry in self.spss_last_result.get("matched", [])
                if entry[0] != "Matched"
            ]
            missing_rows = [
                ("Missing in Target", var, "-", "-", "-", "-", "ไม่พบตัวแปรปลายทาง")
                for var in self.spss_last_result.get("missing", [])
            ]
            extra_rows = [
                ("Extra in Target", "-", "-", "-", "-", "-", "ตัวแปรเกินในปลายทาง")
                for var in self.spss_last_result.get("extra", [])
            ]
            data = rows + missing_rows + extra_rows
            df = pd.DataFrame(
                data,
                columns=[
                    "Status",
                    "Variable",
                    "Source Label",
                    "Target Label",
                    "Source Value Labels",
                    "Target Value Labels",
                    "Note",
                ],
            )
            df.to_excel(path, index=False)
            QMessageBox.information(self, "บันทึกสำเร็จ", f"บันทึกไฟล์แล้วที่:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "บันทึกไม่สำเร็จ", f"ไม่สามารถบันทึกไฟล์:\n{exc}")

    def clear_spss_results(self):
        self.spss_table.setRowCount(0)
        self.spss_summary.setText("พร้อมใช้งาน")
        self.spss_log.clear()
        self.spss_last_result = None

    def _build_combined_rows(self, result):
        rows = []
        if "matched" in result:
            for entry in result["matched"]:
                if len(entry) == 3:
                    rows.append({
                        "status": "Matched",
                        "source": entry[0],
                        "target": entry[1],
                        "note": entry[2],
                    })
                elif len(entry) >= 7:
                    rows.append({
                        "status": entry[0],
                        "variable": entry[1],
                        "source_label": entry[2],
                        "target_label": entry[3],
                        "source_value_labels": entry[4],
                        "target_value_labels": entry[5],
                        "note": entry[6],
                    })
                else:
                    rows.append({
                        "status": entry[0],
                        "source": entry[1],
                        "target": entry[3],
                        "note": entry[4],
                    })
        for entry in result.get("missing", []):
            rows.append({"status": "Missing", "source": entry, "target": "", "note": ""})
        for entry in result.get("extra", []):
            rows.append({"status": "Extra", "source": "", "target": entry, "note": ""})
        return rows

    def _fill_table(self, table, rows, status_colors):
        table.setRowCount(len(rows))
        for row_idx, row_data in enumerate(rows):
            status_value = row_data[0]
            color = status_colors.get(status_value, QColor("#1f2933"))
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if col_idx == 0:
                    item.setForeground(color)
                table.setItem(row_idx, col_idx, item)

        table.resizeColumnsToContents()

    def _now(self):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def run_this_app(working_dir=None):
    app = QApplication.instance()
    created_app = False
    if app is None:
        app = QApplication(sys.argv)
        created_app = True

    window = MapDataWindow()
    window.show()

    if created_app:
        sys.exit(app.exec())


if __name__ == "__main__":
    run_this_app()
