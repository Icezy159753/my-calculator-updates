#!/usr/bin/env python
"""TRF SPSS Excel 01 — portable single-file edition.

Copy this .py file alone. No local modules, icons, or templates are required.
Install dependencies once on a new computer:
    python -m pip install PyQt6 pyreadstat pandas openpyxl
Run:
    python "TRF SPSS Excel 01.py"

Embedded control images are unpacked to an automatically cleaned temporary
directory, never alongside this file. User data is chosen in the application.
"""
from __future__ import annotations


# Embedded control artwork; no external asset files need to be carried.
import tempfile
import atexit
from pathlib import Path

_control_assets = tempfile.TemporaryDirectory(prefix='trf_controls_')
atexit.register(_control_assets.cleanup)
(Path(_control_assets.name) / 'check.svg').write_text('<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16"><path d="M3.5 8 6.5 11 12.5 5" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>\n', encoding='utf-8')
(Path(_control_assets.name) / 'chevron.svg').write_text('<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16"><path d="m4 6 4 4 4-4" fill="none" stroke="#606571" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/></svg>\n', encoding='utf-8')


# ---- app_identity.py ----
import base64
from PyQt6.QtGui import QIcon, QPixmap

APP_NAME = 'TRF SPSS Excel 01'
WINDOWS_APP_ID = 'TRF.SPSS.Excel.01'
ICON_PNG = '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'


def app_icon():
    pixmap = QPixmap()
    pixmap.loadFromData(base64.b64decode(ICON_PNG), 'PNG')
    return QIcon(pixmap)


# ---- excel_export.py ----
import textwrap
from openpyxl.styles import Font, PatternFill, Alignment


def questionnaire_blocks(result):
    """Every exported variable, in Data order, including untransformed questions."""
    # Combined MA needs the original option codebook as well as combination labels.
    combined_options = {}
    for record in result.dictionary.to_dict('records'):
        variable = record['Variable']
        if record['MA Group'] and result.formats.get(variable, '').startswith('A'):
            code = record['Option Code']
            if code != '':
                combined_options.setdefault(variable, {})[code] = record['Option Label'] or record['Label']
    for variable in result.data.columns:
        codes = dict(combined_options.get(variable, {}))
        codes.update(result.values.get(variable, {}))
        yield variable, result.labels.get(variable) or variable, codes


def write_qnr(workbook, result):
    blocks = list(questionnaire_blocks(result))
    if 1 + sum(1 + len(codes) for _, _, codes in blocks) > 1048576:
        raise ValueError('ชีท QNR เกินจำนวนแถวสูงสุดของ Excel')
    sheet = workbook.create_sheet('QNR')
    sheet.append(['Q.No', 'QNR'])
    sheet.column_dimensions['A'].width = 15
    sheet.column_dimensions['B'].width = 80
    sheet.freeze_panes = 'B2'
    header_fill = PatternFill('solid', fgColor='FFD9D9D9')
    question_fill = PatternFill('solid', fgColor='FFDDEBF7')
    normal = Font(name='Tahoma', size=11, color='20242E')
    bold = Font(name='Tahoma', size=11, bold=True, color='20242E')
    for cell in sheet[1]:
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
    sheet.row_dimensions[1].height = 25

    def add_row(left, right, question=False):
        sheet.append([left, right])
        row = sheet.max_row
        for cell in sheet[row]:
            if isinstance(cell.value, str):
                cell.data_type = 's'  # Labels beginning with '=' are text, never formulas.
            cell.font = bold if question else normal
            cell.alignment = Alignment(vertical='top', wrap_text=True)
            if question:
                cell.fill = question_fill
        # Reserve enough height for long questions and multi-line option labels.
        def lines(value, width):
            return sum(max(1, len(textwrap.wrap(part, width=width))) for part in str(value).split('\n'))
        count = max(lines(left, 13), lines(right, 72))
        sheet.row_dimensions[row].height = min(409, max(21, count * 16 + 6))

    for variable, question, codes in blocks:
        add_row(variable, question, question=True)
        for code, caption in codes.items():
            add_row(code, caption)
    return sheet


# ---- transform_core.py ----

from dataclasses import dataclass
from pathlib import Path
import re
import os
import tempfile

import pandas as pd
import pyreadstat


PATTERN = re.compile(r"^(.+?)(\$|_O)(\d+)$", re.I)


@dataclass
class Group:
    key: str
    base: str
    columns: list[str]
    kind: str
    question: str


@dataclass
class Result:
    data: pd.DataFrame
    labels: dict
    values: dict
    missing: dict
    formats: dict
    measures: dict
    dictionary: pd.DataFrame


def read_source(path):
    # Preserve SPSS numeric dates + their display formats and user-missing codes.
    return pyreadstat.read_sav(str(path), user_missing=True, disable_datetime_conversion=True)


def clean(series, meta):
    result = series.copy()
    for interval in (getattr(meta, 'missing_ranges', {}) or {}).get(series.name, []):
        result = result.mask(result.between(interval['lo'], interval['hi']))
    return result


def detect_groups(data, meta):
    groups = {}
    for column in data:
        match = PATTERN.match(column)
        if match:
            base, separator, _ = match.groups()
            key = base + separator.upper()
            groups.setdefault(key, []).append(column)
    result = []
    for key, columns in groups.items():
        columns.sort(key=lambda c: int(PATTERN.match(c)[3]))
        base = PATTERN.match(columns[0])[1]
        observed = set()
        numeric = all(pd.api.types.is_numeric_dtype(data[c]) for c in columns)
        for c in columns:
            observed.update(clean(data[c], meta).dropna().unique())
        kind = 'binary' if numeric and observed and observed <= {0, 1} else 'codes'
        for mr in (getattr(meta, 'mr_sets', {}) or {}).values():
            if set(mr['variable_list']) == set(columns):
                kind = 'binary' if mr['is_dichotomy'] and mr['counted_value'] == 1 else 'codes'
        if not numeric:
            kind = 'unsupported'
        question = meta.column_names_to_labels.get(columns[0]) or base
        result.append(Group(key, base, columns, kind, question))
    return result


def code_text(value):
    return str(int(value)) if float(value).is_integer() else str(value)


def convert(data, meta, groups, mode='binary', keep_original=False, zero_empty=False):
    if not groups:
        raise ValueError('กรุณาเลือกกลุ่ม MA อย่างน้อย 1 กลุ่ม')
    labels = dict(meta.column_names_to_labels)
    values = {k: dict(v) for k, v in meta.variable_value_labels.items()}
    missing = dict(getattr(meta, 'missing_ranges', {}) or {})
    formats = dict(getattr(meta, 'original_variable_types', {}) or {})
    measures = dict(getattr(meta, 'variable_measure', {}) or {})
    selected = {c for g in groups for c in g.columns}
    retained = list(data.columns) if keep_original else [c for c in data if c not in selected]
    used = {c.lower() for c in data.columns}
    generated = {}
    records = []

    def name_for(proposed):
        # SPSS variable names are at most 64 bytes; retain UTF-8 boundaries.
        def trim(s, length):
            return s.encode('utf-8')[:length].decode('utf-8', errors='ignore')
        candidate = trim(proposed, 64)
        i = 2
        while candidate.lower() in used:
            suffix = f'_{i}'
            candidate = trim(proposed, 64 - len(suffix)) + suffix
            i += 1
        used.add(candidate.lower())
        return candidate

    for group in groups:
        block = pd.DataFrame({c: clean(data[c], meta) for c in group.columns})
        if group.kind not in ('binary', 'codes') or not all(pd.api.types.is_numeric_dtype(block[c]) for c in block):
            raise ValueError(f'{group.base}: รองรับเฉพาะรหัสคำตอบตัวเลข')
        option_labels = {}
        if group.kind == 'binary':
            observed = set(block.stack().dropna().unique())
            if not observed <= {0, 1}:
                raise ValueError(f'{group.base}: พบค่าที่ไม่ใช่ 0/1 กรุณาเลือกชนิดต้นทางเป็นช่องรหัส')
            codes = [int(PATTERN.match(c)[3]) for c in group.columns]
            if len(codes) != len(set(codes)):
                raise ValueError(f'{group.base}: รหัสท้ายชื่อตัวแปรซ้ำกัน')
            binary = block.copy()
            binary.columns = codes
            for c, code in zip(group.columns, codes):
                option_labels[code] = meta.column_names_to_labels.get(c) or f'Code {code}'
        else:
            if zero_empty:
                block = block.mask(block == 0)
            for c in group.columns:
                for code, label in values.get(c, {}).items():
                    if not isinstance(code, (int, float)) or (zero_empty and code == 0):
                        continue
                    if any(r['lo'] <= code <= r['hi'] for r in missing.get(c, [])):
                        continue
                    if code in option_labels and option_labels[code] != label:
                        raise ValueError(f'{group.base}: Label ของรหัส {code} ไม่ตรงกันระหว่างช่อง')
                    option_labels[code] = label
            observed = set(block.stack().dropna().unique())
            codes = sorted(observed | set(option_labels))
            if any(not float(code).is_integer() or code < 0 for code in codes):
                raise ValueError(f'{group.base}: รหัส MA ต้องเป็นจำนวนเต็มตั้งแต่ 0 ขึ้นไป')
            binary = pd.DataFrame({code: block.eq(code).any(axis=1).astype(float) for code in codes}, index=data.index)
            binary.loc[block.isna().all(axis=1), :] = float('nan')
        if not codes:
            raise ValueError(f'{group.base}: ไม่มีรหัสคำตอบหรือ Value Label ให้แปลง')
        codes = sorted(codes)
        binary = binary[codes]
        out = {}

        def register(name, series, label, mapping, fmt):
            out[name] = series
            labels[name] = label
            values[name] = mapping
            formats[name] = fmt
            measures[name] = 'nominal'

        if mode == 'binary':
            for code in codes:
                name = name_for(f'{group.base}_MA{code_text(code)}')
                label = f'{group.question} | {code_text(code)}: {option_labels.get(code, code_text(code))}'
                register(name, binary[code], label, {0: 'No', 1: 'Yes'}, 'F1.0')
                for value, caption in ((0, 'No'), (1, 'Yes')):
                    records.append([name, group.question, value, caption, group.key, code, option_labels.get(code, ''), ', '.join(group.columns)])
        elif mode == 'combined':
            name = name_for(f'{group.base}_MA')
            def join_row(row):
                if row.isna().all():
                    return '<MISSING>'
                picked = ','.join(code_text(c) for c in codes if row[c] == 1)
                if row.isna().any():
                    return picked + (';' if picked else '') + '?=' + ','.join(code_text(c) for c in codes if pd.isna(row[c]))
                return picked or '<NONE>'
            combined = binary.apply(join_row, axis=1)
            mapping = {}
            for raw in combined.unique():
                if raw == '<MISSING>':
                    mapping[raw] = 'Missing / ไม่ได้ตอบ'
                elif raw == '<NONE>':
                    mapping[raw] = 'No selection / ไม่เลือก'
                else:
                    first = raw.split(';')[0]
                    captions = [option_labels.get(float(c), c) for c in first.split(',') if c and not c.startswith('?=')]
                    mapping[raw] = ' | '.join(captions) + (' [มีช่อง Missing]' if '?=' in raw else '')
            register(name, combined, group.question, mapping, f'A{max(8, combined.str.encode("utf-8").str.len().max())}')
            for code in codes:
                records.append([name, group.question, code, option_labels.get(code, code_text(code)), group.key, code, option_labels.get(code, ''), ', '.join(group.columns)])
            for raw, caption in mapping.items():
                records.append([name, group.question, raw, caption, group.key, '', '', 'ค่าที่พบในข้อมูลรวมช่อง'])
        else:
            raise ValueError(f'Unknown mode: {mode}')
        generated[group.columns[0]] = pd.DataFrame(out, index=data.index)

    pieces = []
    for c in data:
        if c in retained:
            pieces.append(data[[c]])
        if c in generated:
            pieces.append(generated[c])
    output = pd.concat(pieces, axis=1)
    for c in retained:
        for code, caption in (values.get(c) or {'': ''}).items():
            records.append([c, labels.get(c, ''), code, caption, '', '', '', c])
    def subset(mapping):
        return {k: v for k, v in mapping.items() if k in output}
    dictionary = pd.DataFrame(records, columns=['Variable', 'Question', 'Code', 'Label', 'MA Group', 'Option Code', 'Option Label', 'Source'])
    return Result(output, subset(labels), subset(values), subset(missing), subset(formats), subset(measures), dictionary)


def save_result(result, meta, path, source=None):
    path = Path(path).resolve()
    if source and path == Path(source).resolve():
        raise ValueError('กรุณาบันทึกเป็นไฟล์ใหม่ ไม่เขียนทับไฟล์ต้นฉบับ')
    if path.suffix.lower() not in ('.sav', '.xlsx'):
        raise ValueError('รองรับ .sav หรือ .xlsx เท่านั้น')
    fd, temporary = tempfile.mkstemp(prefix='.ma_export_', suffix=path.suffix, dir=path.parent)
    os.close(fd)
    try:
        if path.suffix.lower() == '.sav':
            pyreadstat.write_sav(result.data, temporary, column_labels=result.labels,
                                variable_value_labels=result.values, missing_ranges=result.missing,
                                variable_format=result.formats, variable_measure=result.measures,
                                file_label=getattr(meta, 'file_label', '') or '',
                                note=getattr(meta, 'notes', None), row_compress=True)
        else:
            if len(result.data) > 1048575 or len(result.data.columns) > 16384 or len(result.dictionary) > 1048575:
                raise ValueError('ข้อมูลเกินจำนวนแถวหรือคอลัมน์สูงสุดของ Excel')
            with pd.ExcelWriter(temporary, engine='openpyxl') as writer:
                result.data.to_excel(writer, sheet_name='Data', index=False)
                from openpyxl.styles import Font, PatternFill, Alignment
                for ws in writer.book:
                    ws.freeze_panes = 'A2'
                    ws.auto_filter.ref = ws.dimensions
                    for row in ws:
                        for cell in row:
                            # Questionnaire text must never be interpreted as a formula.
                            if cell.data_type == 'f':
                                cell.data_type = 's'
                    for cell in ws[1]:
                        cell.font = Font(name='Aptos', bold=True, color='FFFFFF')
                        cell.fill = PatternFill('solid', fgColor='346BDE')
                    ws.row_dimensions[1].height = 28
                    for column in ws.columns:
                        letter = column[0].column_letter
                        ws.column_dimensions[letter].width = 20
                write_qnr(writer.book, result)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


# ---- ui_components.py ----
from pathlib import Path
import pandas as pd
from PyQt6.QtCore import Qt, pyqtSignal, QAbstractTableModel
from PyQt6.QtWidgets import QFrame, QHBoxLayout, QPushButton, QButtonGroup, QLabel, QSizePolicy

STYLE = '''
QWidget { color: #20242e; font-family: 'Leelawadee UI'; font-size: 13px; }
QMainWindow, QDialog { background: #f5f5f7; }
QLabel { background: transparent; }
QLabel#brand { font-size: 19px; font-weight: 700; }
QLabel#title { font-size: 24px; font-weight: 700; }
QLabel#section { font-size: 16px; font-weight: 700; }
QLabel#muted { color: #606571; }
QLabel#badge { background: #e8effc; color: #285bb9; padding: 5px 10px; border-radius: 6px; }
QFrame#topbar { background: #ffffff; border-bottom: 1px solid #e4e5e9; }
QFrame#panel { background: #ffffff; border: 1px solid #e0e2e7; border-radius: 12px; }
QFrame#settings { background: #eeeff2; border-radius: 12px; }
QFrame#footer { background: #ffffff; border-top: 1px solid #e4e5e9; }
QFrame#divider { background: #dbdde3; max-height: 1px; }
QPushButton { background: #e3edff; color: #2c4f86; border: 1px solid #bfd2f3; border-radius: 7px; padding: 8px 15px; font-weight: 600; min-height: 20px; }
QPushButton:hover { background: #d3e3ff; border-color: #8fafe4; }
QPushButton:pressed { background: #c1d6fa; }
QPushButton:focus { border: 1px solid #3574e6; }
QPushButton:disabled { color: #8b909b; background: #eceef2; border-color: #e5e6ea; }
QPushButton#primary { background: #ccebdd; color: #24563f; border: 1px solid #9acdb5; padding: 10px 22px; }
QPushButton#primary:hover { background: #b5e2cc; border-color: #74b795; }
QPushButton#primary:pressed { background: #9ed4b9; }
QPushButton#primary:disabled { background: #eceef2; color: #8b909b; border-color: #e5e6ea; }
QPushButton#quiet { background: #eee7fb; border: 1px solid #d9cbed; color: #604787; padding: 6px 9px; }
QPushButton#quiet:hover { background: #e2d5f5; border-color: #bea7dc; }
QPushButton#quiet:pressed { background: #d6c4ed; }
QPushButton#quiet:disabled { background: #eceef2; color: #8b909b; border-color: #e5e6ea; }
QPushButton#quiet:focus { border-color: #3574e6; }
QFrame#segment { background: #e7edf8; border-radius: 9px; }
QFrame#segment QPushButton { background: #f0f5ff; border: 1px solid #d5e1f6; color: #49618b; border-radius: 6px; padding: 7px 13px; font-weight: 600; }
QFrame#segment QPushButton:checked { background: #cadcfc; color: #203f73; border-color: #8cabdf; }
QFrame#segment QPushButton:hover:!checked { background: #dce9ff; border-color: #aec6ee; }
QFrame#segment QPushButton:pressed { background: #b8d0f6; }
QFrame#segment QPushButton:focus { border-color: #3574e6; }
QFrame#segment QPushButton:disabled { background: #eceef2; color: #9298a3; border-color: #e1e4e9; }
QLineEdit { background: #ffffff; border: 1px solid #d9dde4; border-radius: 7px; padding: 8px 12px; selection-background-color: #d7e5ff; }
QLineEdit:focus { border-color: #3574e6; }
QComboBox { background: #eaf1ff; color: #2c4f86; border: 1px solid #bfd2f3; border-radius: 6px; padding: 7px 28px 7px 10px; }
QComboBox::drop-down { border: 0; width: 25px; }
QComboBox::down-arrow { image: url(ASSET_ROOT/chevron.svg); width: 12px; height: 12px; }
QComboBox QAbstractItemView { background: #ffffff; selection-background-color: #e4edff; selection-color: #20242e; padding: 5px; border: 1px solid #d9dde4; }
QTableView { background: #ffffff; alternate-background-color: #fafbfc; border: 0; gridline-color: #eceef2; selection-background-color: #e7efff; selection-color: #20242e; }
QTableView::item { padding: 7px; border-bottom: 1px solid #eff0f3; }
QTableView::item:focus { border: 1px solid #90b1ef; }
QHeaderView::section { background: #f6f7f9; color: #5c626e; padding: 10px 9px; border: 0; border-bottom: 1px solid #e6e8ed; font-weight: 600; }
QTableCornerButton::section { background: #f6f7f9; border: 0; }
QScrollBar:vertical { background: #f3f4f6; width: 12px; margin: 0; }
QScrollBar:horizontal { background: #f3f4f6; height: 12px; margin: 0; }
QScrollBar::handle:vertical { background: #bec4ce; min-height: 28px; border: 3px solid #f3f4f6; border-radius: 5px; }
QScrollBar::handle:horizontal { background: #bec4ce; min-width: 28px; border: 3px solid #f3f4f6; border-radius: 5px; }
QScrollBar::handle:hover { background: #8793a5; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; border: 0; background: transparent; }
QScrollBar::add-page, QScrollBar::sub-page { background: transparent; }
QCheckBox { spacing: 9px; background: transparent; }
QCheckBox::indicator, QTableView::indicator { width: 17px; height: 17px; border: 1px solid #a6bce0; border-radius: 4px; background: #eaf1ff; }
QCheckBox::indicator:checked, QTableView::indicator:checked { background: #2869dc; border-color: #2869dc; image: url(ASSET_ROOT/check.svg); }
QCheckBox::indicator:hover { border-color: #2869dc; }
QCheckBox:focus { color: #2058b8; }
QProgressBar { background: #e7ebf3; border: 0; }
QProgressBar::chunk { background: #2869dc; }
QToolTip { background: #20242e; color: #ffffff; border: 0; padding: 7px; }
'''.replace('ASSET_ROOT', Path(_control_assets.name).as_posix())


def tint_action(widget, tone):
    """Give each action a pastel identity while retaining clear interaction states."""
    palettes = {
        'blue': ('#edf3ff', '#cbdcfc', '#aec8f0', '#31558b'),
        'violet': ('#f1ebfc', '#e0d2f6', '#c9b3e8', '#614287'),
        'mint': ('#e7f5ed', '#c7e9d6', '#a6d4bc', '#285d43'),
        'peach': ('#fff0e4', '#f9d9bf', '#e9bb96', '#81502c'),
        'rose': ('#fcecf0', '#f3d0dc', '#e3adc0', '#813e57'),
        'teal': ('#e4f4f3', '#c0e5e1', '#99cec8', '#285c59'),
    }
    light, selected, border, ink = palettes[tone]
    widget.setStyleSheet(f'''
        QPushButton {{ background: {light}; color: {ink}; border: 1px solid {border}; }}
        QPushButton:hover {{ background: {selected}; border-color: {ink}; }}
        QPushButton:checked {{ background: {selected}; color: {ink}; border: 1px solid {ink}; }}
        QPushButton:pressed {{ background: {border}; }}
        QPushButton:focus {{ border: 2px solid {ink}; }}
        QPushButton:disabled {{ background: #eceef2; color: #8b909b; border: 1px solid #e1e4e9; }}
    ''')


def label(text, name=None):
    widget = QLabel(text)
    widget.setTextFormat(Qt.TextFormat.PlainText)
    if name:
        widget.setObjectName(name)
    return widget


class Segment(QFrame):
    currentIndexChanged = pyqtSignal(int)

    def __init__(self, items, tones=None):
        super().__init__()
        self.setObjectName('segment')
        self.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
        self.items, self._index, self.buttons = items, 0, []
        row = QHBoxLayout(self)
        row.setContentsMargins(3, 3, 3, 3)
        row.setSpacing(2)
        self.group = QButtonGroup(self)
        for index, (text, _) in enumerate(items):
            button = QPushButton(text)
            button.setCheckable(True)
            if tones:
                tint_action(button, tones[index])
            self.group.addButton(button, index)
            self.buttons.append(button)
            row.addWidget(button, 1)
        self.buttons[0].setChecked(True)
        self.group.idClicked.connect(self.setCurrentIndex)

    def currentIndex(self):
        return self._index

    def currentData(self):
        return self.items[self._index][1]

    def setCurrentIndex(self, index):
        self.buttons[index].setChecked(True)
        if index != self._index:
            self._index = index
            self.currentIndexChanged.emit(index)


class PreviewModel(QAbstractTableModel):
    """Format visible cells on demand while retaining every row."""
    def __init__(self, frame, labels=None):
        super().__init__()
        self.frame, self.labels = frame, labels or {}

    def rowCount(self, parent=None):
        return 0 if parent is not None and parent.isValid() else len(self.frame)

    def columnCount(self, parent=None):
        return 0 if parent is not None and parent.isValid() else len(self.frame.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if index.isValid() and role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole):
            value = self.frame.iat[index.row(), index.column()]
            if pd.isna(value):
                return ''
            return str(int(value)) if isinstance(value, float) and value.is_integer() else str(value)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal:
            column = str(self.frame.columns[section])
            if role == Qt.ItemDataRole.DisplayRole:
                return column
            if role == Qt.ItemDataRole.ToolTipRole:
                return f'{column}\n{self.labels.get(column, "")}'
        elif role == Qt.ItemDataRole.DisplayRole:
            return str(section + 1)


# ---- app.py ----
import sys
from pathlib import Path
import pandas as pd
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QColor, QShortcut, QKeySequence, QPalette, QCursor
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QFrame, QTableWidget, QTableWidgetItem,
    QHeaderView, QComboBox, QCheckBox, QTableView, QMessageBox, QProgressBar,
    QAbstractItemView, QStackedWidget, QLineEdit, QDialog, QDialogButtonBox,
)


class Worker(QThread):
    succeeded = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, job):
        super().__init__()
        self.job = job

    def run(self):
        try:
            self.succeeded.emit(self.job())
        except Exception as exc:
            self.failed.emit(str(exc))


def button(text, callback, style=None):
    item = QPushButton(text)
    if style:
        item.setObjectName(style)
    item.clicked.connect(callback)
    return item


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(app_icon())
        # Keep native controls consistent with this light UI on dark Windows themes.
        palette = QPalette()
        for role, color in {
            QPalette.ColorRole.Window: '#f5f5f7',
            QPalette.ColorRole.WindowText: '#20242e',
            QPalette.ColorRole.Base: '#ffffff',
            QPalette.ColorRole.AlternateBase: '#fafbfc',
            QPalette.ColorRole.Text: '#20242e',
            QPalette.ColorRole.Button: '#f0f1f4',
            QPalette.ColorRole.ButtonText: '#20242e',
            QPalette.ColorRole.Highlight: '#2869dc',
            QPalette.ColorRole.HighlightedText: '#ffffff',
            QPalette.ColorRole.PlaceholderText: '#606571',
            QPalette.ColorRole.Light: '#ffffff',
            QPalette.ColorRole.Midlight: '#f4f5f7',
            QPalette.ColorRole.Mid: '#d6d9df',
            QPalette.ColorRole.Dark: '#a8aeb9',
            QPalette.ColorRole.Shadow: '#878d99',
        }.items():
            palette.setColor(role, QColor(color))
        self.setPalette(palette)
        self.resize(1320, 860)
        self.setMinimumSize(1060, 730)
        self._initial_position_set = False
        self.data = self.meta = self.source = self.result = None
        self.groups, self.source_kinds = [], []
        self.worker = None
        self.setAcceptDrops(True)
        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        topbar = QFrame()
        topbar.setObjectName('topbar')
        top = QHBoxLayout(topbar)
        top.setContentsMargins(26, 13, 26, 13)
        mark = label('')
        mark.setPixmap(self.windowIcon().pixmap(40, 40))
        mark.setFixedSize(40, 40)
        top.addWidget(mark)
        top.addSpacing(3)
        top.addWidget(label(APP_NAME, 'brand'))
        top.addSpacing(22)
        self.file_info = label('ยังไม่ได้เปิดไฟล์', 'muted')
        self.file_info.setMaximumWidth(480)
        top.addWidget(self.file_info)
        top.addStretch()
        self.open_btn = button('เปิดไฟล์ SPSS', self.open_file)
        tint_action(self.open_btn, 'peach')
        self.open_btn.setToolTip('เปิดไฟล์ .sav หรือ .zsav (Ctrl+O)')
        top.addWidget(self.open_btn)
        outer.addWidget(topbar)
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(26, 22, 26, 20)
        layout.setSpacing(17)
        heading = QHBoxLayout()
        titles = QVBoxLayout()
        titles.setSpacing(5)
        titles.addWidget(label('จัดรูปแบบคำตอบ MA', 'title'))
        self.stats = label('เปิดไฟล์หรือลากไฟล์ SPSS มาวางเพื่อเริ่มต้น', 'muted')
        titles.addWidget(self.stats)
        heading.addLayout(titles)
        heading.addStretch()
        self.tabs = Segment([('ตั้งค่าการแปลง', 'setup'), ('พรีวิวข้อมูล', 'preview')], ['blue', 'violet'])
        self.tabs.setMinimumWidth(300)
        self.tabs.currentIndexChanged.connect(self.change_tab)
        heading.addWidget(self.tabs)
        layout.addLayout(heading)
        self.pages = QStackedWidget()
        self.pages.addWidget(self.build_setup())
        self.pages.addWidget(self.build_preview())
        layout.addWidget(self.pages, 1)
        outer.addWidget(body, 1)
        footer = QFrame()
        footer.setObjectName('footer')
        bottom = QHBoxLayout(footer)
        bottom.setContentsMargins(26, 14, 26, 14)
        self.status = label('พร้อมเริ่มงาน', 'muted')
        self.status.setWordWrap(True)
        bottom.addWidget(self.status, 1)
        self.preview_btn = button('พรีวิวผลลัพธ์', self.make_preview)
        tint_action(self.preview_btn, 'violet')
        bottom.addWidget(self.preview_btn)
        self.save_btn = button('บันทึก SPSS…', self.export, 'primary')
        bottom.addWidget(self.save_btn)
        outer.addWidget(footer)
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(3)
        self.progress.hide()
        outer.addWidget(self.progress)
        self.mode.currentIndexChanged.connect(self.invalidate)
        self.format.currentIndexChanged.connect(self.update_format)
        self.keep.toggled.connect(self.invalidate)
        self.zero.toggled.connect(self.invalidate)
        QShortcut(QKeySequence('Ctrl+O'), self, activated=self.open_btn.click)
        QShortcut(QKeySequence('Ctrl+S'), self, activated=self.save_btn.click)
        self.set_busy(False)

    def build_setup(self):
        page = QWidget()
        row = QHBoxLayout(page)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(20)
        panel = QFrame()
        panel.setObjectName('panel')
        content = QVBoxLayout(panel)
        content.setContentsMargins(16, 15, 16, 10)
        content.setSpacing(12)
        toolbar = QHBoxLayout()
        toolbar.addWidget(label('กลุ่ม MA', 'section'))
        self.selection_count = label('0 กลุ่ม', 'badge')
        toolbar.addWidget(self.selection_count)
        toolbar.addStretch()
        self.all_btn = button('เลือกทั้งหมด', lambda: self.select_all(True), 'quiet')
        self.none_btn = button('ล้างการเลือก', lambda: self.select_all(False), 'quiet')
        tint_action(self.all_btn, 'mint')
        tint_action(self.none_btn, 'rose')
        toolbar.addWidget(self.all_btn)
        toolbar.addWidget(self.none_btn)
        content.addLayout(toolbar)
        self.group_search = QLineEdit()
        self.group_search.setPlaceholderText('ค้นหากลุ่ม MA หรือโจทย์…')
        self.group_search.setClearButtonEnabled(True)
        self.group_search.textChanged.connect(self.filter_groups)
        content.addWidget(self.group_search)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(['', 'กลุ่ม', 'ช่องเดิม', 'ผลลัพธ์', 'โจทย์ / Label'])
        self.table.verticalHeader().hide()
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setWordWrap(False)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        for index, width in enumerate([35, 94, 70, 128]):
            self.table.setColumnWidth(index, width)
        self.table.itemChanged.connect(self.invalidate)
        self.table.setMinimumHeight(240)
        self.group_stack = QStackedWidget()
        empty = QWidget()
        empty_layout = QVBoxLayout(empty)
        empty_layout.addStretch()
        empty_layout.addWidget(label('เริ่มจากไฟล์ SPSS ของคุณ', 'section'), alignment=Qt.AlignmentFlag.AlignCenter)
        empty_layout.addWidget(label('ลากไฟล์ .sav มาวาง หรือกดเปิดไฟล์ด้านบน', 'muted'), alignment=Qt.AlignmentFlag.AlignCenter)
        empty_layout.addStretch()
        self.group_stack.addWidget(empty)
        self.group_stack.addWidget(self.table)
        content.addWidget(self.group_stack, 1)
        self.group_note = label('รองรับชื่อตัวแปร $1, $2 และ _O1, _O2', 'muted')
        content.addWidget(self.group_note)
        row.addWidget(panel, 1)
        settings = QFrame()
        settings.setObjectName('settings')
        settings.setFixedWidth(292)
        options = QVBoxLayout(settings)
        options.setContentsMargins(18, 19, 18, 19)
        options.setSpacing(10)
        options.addWidget(label('รูปแบบผลลัพธ์', 'section'))
        self.mode = Segment([('0/1 แยกช่อง', 'binary'), ('รวมช่องเดียว', 'combined')], ['teal', 'peach'])
        options.addWidget(self.mode)
        self.example = label('S10_MA1     S10_MA2     S10_MA3\n       1                   0                   1')
        self.example.setStyleSheet('background: #ffffff; border-radius: 8px; padding: 12px 9px; font-size: 12px; color: #355581;')
        self.example.setMinimumHeight(64)
        options.addWidget(self.example)
        self.hint = label('หนึ่งตัวเลือกต่อหนึ่งคอลัมน์\n0 = No  ·  1 = Yes', 'muted')
        self.hint.setWordWrap(True)
        options.addWidget(self.hint)
        options.addSpacing(3)
        self.keep = QCheckBox('เก็บตัวแปร MA ต้นฉบับด้วย')
        options.addWidget(self.keep)
        self.zero = QCheckBox('ถือว่า 0 เป็นช่องว่าง')
        self.zero.hide()
        divider = QFrame()
        divider.setObjectName('divider')
        divider.setFixedHeight(1)
        options.addWidget(divider)
        options.addWidget(label('ไฟล์ที่บันทึก', 'section'))
        self.format = Segment([('SPSS', 'sav'), ('Excel', 'xlsx')], ['violet', 'mint'])
        options.addWidget(self.format)
        self.format_hint = label('.sav พร้อมโจทย์และ Value Labels\nMA แบบ 0/1 ใช้ No / Yes', 'muted')
        self.format_hint.setWordWrap(True)
        options.addWidget(self.format_hint)
        options.addStretch()
        self.advanced_btn = button('ตั้งค่าต้นทางขั้นสูง…', self.advanced, 'quiet')
        options.addWidget(self.advanced_btn)
        note = label('คงค่า Missing และโจทย์จากไฟล์ต้นฉบับ', 'muted')
        note.setWordWrap(True)
        options.addWidget(note)
        row.addWidget(settings)
        return page

    def build_preview(self):
        panel = QFrame()
        panel.setObjectName('panel')
        content = QVBoxLayout(panel)
        content.setContentsMargins(16, 15, 16, 10)
        content.setSpacing(12)
        row = QHBoxLayout()
        self.preview_caption = label('พรีวิวข้อมูล', 'section')
        row.addWidget(self.preview_caption)
        row.addStretch()
        self.preview_source = Segment([('ผลลัพธ์', 'result'), ('ต้นฉบับ', 'source')], ['teal', 'rose'])
        self.preview_source.currentIndexChanged.connect(self.refresh_preview)
        row.addWidget(self.preview_source)
        content.addLayout(row)
        filters = QHBoxLayout()
        self.column_search = QLineEdit()
        self.column_search.setPlaceholderText('ค้นหาชื่อตัวแปรหรือ Label เช่น S10…')
        self.column_search.setClearButtonEnabled(True)
        self.column_search.textChanged.connect(self.filter_preview)
        filters.addWidget(self.column_search, 1)
        self.ma_only = QCheckBox('เฉพาะ MA ที่เลือก')
        self.ma_only.toggled.connect(self.filter_preview)
        filters.addWidget(self.ma_only)
        content.addLayout(filters)
        self.preview_stack = QStackedWidget()
        empty = QWidget()
        center = QVBoxLayout(empty)
        center.addStretch()
        self.empty_title = label('ยังไม่มีข้อมูลพรีวิว', 'section')
        self.empty_detail = label('เปิดไฟล์และเลือกกลุ่ม MA เพื่อดูผลลัพธ์', 'muted')
        center.addWidget(self.empty_title, alignment=Qt.AlignmentFlag.AlignCenter)
        center.addWidget(self.empty_detail, alignment=Qt.AlignmentFlag.AlignCenter)
        center.addStretch()
        self.preview = QTableView()
        self.preview.setAlternatingRowColors(True)
        self.preview.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.preview.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.preview.horizontalHeader().setDefaultSectionSize(135)
        self.preview.verticalHeader().setDefaultSectionSize(36)
        self.preview_stack.addWidget(empty)
        self.preview_stack.addWidget(self.preview)
        content.addWidget(self.preview_stack, 1)
        bottom = QHBoxLayout()
        self.preview_count = label('แสดงข้อมูลครบทุกแถว', 'muted')
        bottom.addWidget(self.preview_count)
        bottom.addStretch()
        bottom.addWidget(label('วางเมาส์บนชื่อคอลัมน์เพื่อดูโจทย์', 'muted'))
        content.addLayout(bottom)
        return panel

    def update_format(self, *args):
        excel = self.format.currentIndex() == 1
        self.save_btn.setText('บันทึก Excel…' if excel else 'บันทึก SPSS…')
        self.format_hint.setText('.xlsx มีชีท Data และ QNR\nโจทย์และ Code / Label ครบทุกข้อ' if excel else '.sav พร้อมโจทย์และ Value Labels\nMA แบบ 0/1 ใช้ No / Yes')

    def change_tab(self, index):
        self.pages.setCurrentIndex(index)
        if index == 1:
            if self.data is not None and self.result is None and self.settings()[0] and self.preview_source.currentIndex() == 0:
                self.make_preview()
            else:
                self.refresh_preview()

    def invalidate(self, *args):
        self.result = None
        combined = self.mode.currentData() == 'combined'
        output = 'รวมช่องเดียว' if combined else '0/1 แยกช่อง'
        self.table.blockSignals(True)
        selected = 0
        for row in range(self.table.rowCount()):
            checked = self.table.item(row, 0).checkState() == Qt.CheckState.Checked
            selected += int(checked)
            self.table.item(row, 3).setText(output if checked else 'คงเดิม')
            self.table.item(row, 3).setForeground(QColor('#285bb9' if checked else '#737986'))
        self.table.blockSignals(False)
        self.selection_count.setText(f'เลือก {selected} / {len(self.groups)}')
        self.example.setText('S10_MA\n1,3,5' if combined else 'S10_MA1     S10_MA2     S10_MA3\n       1                   0                   1')
        self.hint.setText('รวมรหัสที่เลือก เช่น 1,3,5\nคงสถานะไม่เลือกและ Missing' if combined else 'หนึ่งตัวเลือกต่อหนึ่งคอลัมน์\n0 = No  ·  1 = Yes')
        self.status.setText(f'เลือก {selected} กลุ่ม · {output}')
        self.refresh_preview()
        self.set_busy(bool(self.worker and self.worker.isRunning()))

    def refresh_preview(self, *args):
        source = self.preview_source.currentIndex() == 1
        frame = self.data if source else (self.result.data if self.result is not None else None)
        labels = (self.meta.column_names_to_labels if self.meta else {}) if source else (self.result.labels if self.result else {})
        if frame is None:
            self.preview.setModel(PreviewModel(pd.DataFrame()))
            self.preview_stack.setCurrentIndex(0)
            self.empty_title.setText('ยังไม่มีข้อมูลพรีวิว' if self.data is None else 'ผลลัพธ์รออัปเดต')
            self.empty_detail.setText('เปิดไฟล์และเลือกกลุ่ม MA เพื่อดูผลลัพธ์' if self.data is None else 'กด “พรีวิวผลลัพธ์” เพื่อแปลงตามการตั้งค่าล่าสุด')
            self.preview_count.setText('แสดงข้อมูลครบทุกแถว')
            return
        self.preview_stack.setCurrentIndex(1)
        self.preview.setModel(PreviewModel(frame, labels))
        self.preview_caption.setText('ข้อมูลต้นฉบับ' if source else 'ผลลัพธ์การแปลง')
        self.filter_preview()

    def filter_preview(self, *args):
        model = self.preview.model()
        if model is None:
            return
        term = self.column_search.text().strip().casefold()
        source = self.preview_source.currentIndex() == 1
        selected = self.settings()[0]
        original_ma = {c for g in selected for c in g.columns}
        result_ma = set()
        if self.result is not None:
            dictionary = self.result.dictionary
            result_ma = set(dictionary.loc[dictionary['MA Group'].isin([g.key for g in selected]), 'Variable'])
        visible = 0
        for index, column in enumerate(model.frame.columns):
            text = f'{column} {model.labels.get(column, "")}'.casefold()
            show = term in text and (not self.ma_only.isChecked() or column in (original_ma if source else result_ma))
            self.preview.setColumnHidden(index, not show)
            visible += int(show)
        self.preview_count.setText(f'{len(model.frame):,} แถวทั้งหมด · แสดง {visible:,} / {len(model.frame.columns):,} คอลัมน์')

    def filter_groups(self, text):
        term = text.strip().casefold()
        visible = 0
        for row, group in enumerate(self.groups):
            show = term in f'{group.key} {group.question}'.casefold()
            self.table.setRowHidden(row, not show)
            visible += int(show)
        self.group_note.setText(f'แสดง {visible} / {len(self.groups)} กลุ่ม · เลื่อนตารางเพื่อดูเพิ่มเติม' if visible else 'ไม่พบกลุ่มที่ตรงกับคำค้น')

    def set_busy(self, busy):
        loaded = self.data is not None
        for widget in (self.open_btn, self.table, self.mode, self.format, self.keep, self.all_btn, self.none_btn, self.tabs, self.advanced_btn):
            widget.setEnabled(not busy)
        self.advanced_btn.setEnabled(not busy and loaded)
        self.all_btn.setEnabled(not busy and loaded)
        self.none_btn.setEnabled(not busy and loaded)
        selected = any(self.table.item(r, 0).checkState() == Qt.CheckState.Checked for r in range(self.table.rowCount()))
        self.preview_btn.setEnabled(not busy and loaded and selected)
        self.save_btn.setEnabled(not busy and loaded and selected)
        self.progress.setVisible(busy)
        self.progress.setRange(0, 0 if busy else 100)

    def run_job(self, job, done, message):
        self.set_busy(True)
        self.status.setText(message)
        self.worker = Worker(job)
        self.worker.succeeded.connect(done)
        self.worker.failed.connect(self.error)
        self.worker.finished.connect(lambda: self.set_busy(False))
        self.worker.start()

    def error(self, message):
        self.status.setText('ดำเนินการไม่สำเร็จ')
        QMessageBox.critical(self, APP_NAME, message)

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, 'เลือกไฟล์ SPSS', str(Path.cwd()), 'SPSS (*.sav *.zsav)')
        if path:
            self.load(path)

    def load(self, path):
        def done(payload):
            self.data, self.meta, self.groups = payload
            self.source_kinds = [g.kind for g in self.groups]
            self.source = path
            self.result = None
            self.file_info.setText(Path(path).name)
            self.file_info.setToolTip(str(path))
            self.stats.setText(f'{len(self.data):,} ผู้ตอบ   ·   {len(self.data.columns):,} ตัวแปร   ·   {len(self.groups)} กลุ่ม MA')
            self.table.blockSignals(True)
            self.table.setRowCount(len(self.groups))
            for row, group in enumerate(self.groups):
                check = QTableWidgetItem()
                check.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
                if group.kind == 'unsupported':
                    check.setFlags(Qt.ItemFlag.NoItemFlags)
                check.setCheckState(Qt.CheckState.Unchecked if group.kind == 'unsupported' else Qt.CheckState.Checked)
                self.table.setItem(row, 0, check)
                for index, text in ((1, group.base), (2, str(len(group.columns))), (3, ''), (4, group.question)):
                    self.table.setItem(row, index, QTableWidgetItem(text))
                self.table.item(row, 1).setToolTip(', '.join(group.columns))
                self.table.item(row, 4).setToolTip(group.question)
                self.table.item(row, 2).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setRowHeight(row, 49)
            self.table.blockSignals(False)
            self.group_stack.setCurrentIndex(1)
            self.group_search.clear()
            self.filter_groups('')
            self.tabs.setCurrentIndex(0)
            self.column_search.clear()
            self.invalidate()
        def job():
            data, meta = read_source(path)
            return data, meta, detect_groups(data, meta)
        self.run_job(job, done, 'กำลังอ่านไฟล์และตรวจกลุ่ม MA…')

    def select_all(self, selected):
        self.table.blockSignals(True)
        for row, group in enumerate(self.groups):
            self.table.item(row, 0).setCheckState(Qt.CheckState.Checked if selected and group.kind != 'unsupported' else Qt.CheckState.Unchecked)
        self.table.blockSignals(False)
        self.invalidate()

    def settings(self):
        groups = []
        for row, group in enumerate(self.groups):
            if self.table.item(row, 0).checkState() == Qt.CheckState.Checked:
                groups.append(Group(group.key, group.base, group.columns, self.source_kinds[row], group.question))
        return groups, self.mode.currentData(), self.keep.isChecked(), self.zero.isChecked()

    def advanced(self):
        dialog = QDialog(self)
        dialog.setWindowTitle('ตั้งค่าการอ่านข้อมูลต้นทาง')
        dialog.resize(650, 540)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(22, 22, 22, 22)
        layout.addWidget(label('รูปแบบข้อมูลที่ตรวจพบ', 'section'))
        help_text = label('ใช้ค่าที่ตรวจพบได้ตามปกติ เปลี่ยนเมื่อข้อมูลต้นฉบับใช้รูปแบบอื่น\nผลลัพธ์ยังเป็นรูปแบบที่เลือกในหน้าตั้งค่าการแปลง', 'muted')
        help_text.setWordWrap(True)
        layout.addWidget(help_text)
        table = QTableWidget(len(self.groups), 2)
        table.setHorizontalHeaderLabels(['กลุ่ม', 'ข้อมูลต้นฉบับ'])
        table.verticalHeader().hide()
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        for row, group in enumerate(self.groups):
            table.setItem(row, 0, QTableWidgetItem(group.key))
            combo = QComboBox()
            combo.addItem('รหัสที่เลือก เช่น 1, 3, 5', 'codes')
            combo.addItem('0/1 แยกช่อง', 'binary')
            if group.kind == 'unsupported':
                combo.addItem('ข้อความ (ไม่แปลง)', 'unsupported')
                combo.setEnabled(False)
            combo.setCurrentIndex(combo.findData(self.source_kinds[row]))
            table.setCellWidget(row, 1, combo)
            table.setRowHeight(row, 44)
        layout.addWidget(table)
        zero = QCheckBox('รหัสต้นทาง 0 หมายถึงช่องว่าง (เฉพาะแบบรหัสที่เลือก)')
        zero.setChecked(self.zero.isChecked())
        layout.addWidget(zero)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Save).setText('ใช้การตั้งค่า')
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText('ยกเลิก')
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.source_kinds = [table.cellWidget(r, 1).currentData() for r in range(len(self.groups))]
            self.zero.setChecked(zero.isChecked())
            self.invalidate()

    def show_result(self, result):
        self.result = result
        self.preview_source.setCurrentIndex(0)
        self.tabs.setCurrentIndex(1)
        self.refresh_preview()
        self.status.setText(f'พร้อมบันทึก · {len(result.data):,} แถว · {len(result.data.columns):,} ตัวแปร')

    def make_preview(self):
        args = self.settings()
        if self.result is not None:
            self.show_result(self.result)
            return
        self.run_job(lambda: convert(self.data, self.meta, *args), self.show_result, 'กำลังเตรียมพรีวิวผลลัพธ์…')

    def export(self):
        extension = '.sav' if self.format.currentIndex() == 0 else '.xlsx'
        default = Path(self.source).with_name(Path(self.source).stem + '_MA' + extension)
        path, _ = QFileDialog.getSaveFileName(self, 'บันทึกผลลัพธ์', str(default), f'Output (*{extension})')
        if not path:
            return
        if Path(path).suffix.lower() != extension:
            path += extension
            if Path(path).exists() and QMessageBox.question(self, 'ไฟล์มีอยู่แล้ว', f'เขียนทับ {path}?') != QMessageBox.StandardButton.Yes:
                return
        args = self.settings()
        cached = self.result
        def job():
            result = cached if cached is not None else convert(self.data, self.meta, *args)
            save_result(result, self.meta, path, self.source)
            return result
        def done(result):
            self.show_result(result)
            self.status.setText(f'บันทึกแล้ว: {Path(path).name}')
            self.status.setToolTip(path)
            QMessageBox.information(self, 'บันทึกสำเร็จ', f'{path}\n\n{len(result.data):,} แถว · {len(result.data.columns):,} ตัวแปร')
        self.run_job(job, done, 'กำลังบันทึกไฟล์…')

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and not (self.worker and self.worker.isRunning()):
            urls = event.mimeData().urls()
            if len(urls) == 1 and Path(urls[0].toLocalFile()).suffix.lower() in ('.sav', '.zsav'):
                event.acceptProposedAction()

    def dropEvent(self, event):
        self.load(event.mimeData().urls()[0].toLocalFile())

    def showEvent(self, event):
        super().showEvent(event)
        if not self._initial_position_set:
            self._initial_position_set = True
            QTimer.singleShot(0, self.center_on_screen)

    def center_on_screen(self):
        screen = QApplication.screenAt(QCursor.pos()) or self.screen()
        if screen is None:
            return
        available = screen.availableGeometry()
        frame = self.frameGeometry()
        extra_width = frame.width() - self.width()
        extra_height = frame.height() - self.height()
        self.resize(min(self.width(), max(self.minimumWidth(), available.width() - extra_width - 24)),
                    min(self.height(), max(self.minimumHeight(), available.height() - extra_height - 24)))
        frame = self.frameGeometry()
        frame.moveCenter(available.center())
        self.move(frame.topLeft())

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, 'กำลังทำงาน', 'กรุณารอให้การอ่านหรือบันทึกไฟล์เสร็จก่อนปิดโปรแกรม')
            event.ignore()
        else:
            event.accept()




 
# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None, argv=None): # ชื่อฟังก์ชันนี้จะถูกใช้ใน Launcher
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน QuotaSamplerApp.
    """
    print(f"--- QUOTA_SAMPLER_INFO: Starting 'QuotaSamplerApp' via run_this_app() ---")
    try:
    # --- ส่วนที่ใช้รันโปรแกรม ---
    #if __name__ == '__main__':
        if sys.platform == 'win32':
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(WINDOWS_APP_ID)
        # Launcher arguments belong to Main_Program, not to the SPSS reader.
        # Direct execution passes its file arguments explicitly below.
        app_args = list(argv) if argv is not None else []
        app = QApplication([sys.argv[0], *app_args])
        app.setApplicationName(APP_NAME)
        app.setApplicationDisplayName(APP_NAME)
        app.setWindowIcon(app_icon())
        app.setStyle('Fusion')
        app.setStyleSheet(STYLE)
        window = Window()
        window.show()
        if app_args:
            window.load(app_args[0])
        sys.exit(app.exec())

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
    run_this_app(argv=sys.argv[1:])

    print("--- Finished direct execution of QuotaSamplerApp.py ---")
# <<< END OF CHANGES >>>
