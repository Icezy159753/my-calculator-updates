# Cut MKT — single-file editable source
# Python 3.12+ on Windows. Install dependencies once:
# python -m pip install "customtkinter>=5.2,<6" "openpyxl>=3.1,<4" "Pillow>=10,<13"
# Run: python "Cut MKT.py"
# Contains the complete Excel engine, GUI and embedded icon.
# No table_engine.py, assets folder, or EXE is required.
# Excel source files and user Settings are selected at runtime.
# Generated from table_engine.py and table_trimmer.py by build_onefile.py.

"""Desktop batch table trimmer. Run with Python 3.12+ on Windows."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import importlib.util
import subprocess
import tempfile
import hashlib
import shutil


def ensure_dependencies():
    """Allow an IDE's previously selected Python to hand off to this project."""
    if getattr(sys, 'frozen', False):
        return
    missing = [name for name in ('customtkinter', 'openpyxl', 'PIL') if importlib.util.find_spec(name) is None]
    if not missing:
        return
    local_python = Path(__file__).resolve().parent / '.venv' / 'Scripts' / 'python.exe'
    if local_python.is_file() and local_python.resolve() != Path(sys.executable).resolve():
        raise SystemExit(subprocess.call([str(local_python), str(Path(__file__).resolve()), *sys.argv[1:]]))
    raise SystemExit(
        'Missing packages: ' + ', '.join(missing) + '\n'
        'Run: python -m pip install "customtkinter>=5.2,<6" "openpyxl>=3.1,<4" "Pillow>=10,<13"'
    )


if __name__ == '__main__':
    ensure_dependencies()

import customtkinter as ctk
from PIL import Image
# ===== Excel engine (editable source) =====
"""Strict batch validation and worksheet deletion with compact Contents rows."""

from dataclasses import dataclass, field
from pathlib import Path
from zipfile import ZipFile
from xml.dom import minidom
import xml.etree.ElementTree as ET
import posixpath
import re
import os
import tempfile
import unicodedata
from bisect import bisect_left, bisect_right
import hashlib

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
N = {'m': MAIN}


def key(text: str) -> str:
    return ' '.join(unicodedata.normalize('NFC', str(text)).split()).casefold()


def contents_key(text: str) -> str:
    """Ignore only export decorations, never question wording or question codes."""
    text = key(text).lstrip('*').strip()
    return re.sub(r'\s*\[(?:sa|ma|oa|na)\]\s*$', '', text).strip()


def question_row_key(text: str) -> tuple[str, int, str] | None:
    """Match a matrix item by question, R number and its complete item wording."""
    text = contents_key(text)
    head = re.match(r'^(q\d+[a-z]?)\b', text)
    rows = list(re.finditer(r'\(\s*r\s*(\d+)\s*\)', text))
    if head is None or len(rows) != 1:
        return None
    row = rows[0]
    codes = re.findall(r'\bq\d+[a-z]?\b', text[:row.start()])
    if not codes or codes[-1] != head[1]:
        return None
    item = text[row.end():].strip()
    return (head[1], int(row[1]), item) if item else None


def link_sheet(location: str) -> str:
    location = location.lstrip('#')
    if '!' not in location:
        return ''
    name = location.rsplit('!', 1)[0]
    if name.startswith("'") and name.endswith("'"):
        name = name[1:-1].replace("''", "'")
    return '' if '[' in name else name


@dataclass(frozen=True)
class Entry:
    title: str
    sheet: str = ''
    row: int = 0


@dataclass
class Catalog:
    path: Path
    contents: str
    contents_part: str
    entries: list[Entry]
    sheets: dict[str, tuple[str, str]]


class ContentsMatcher:
    """One matching implementation for GUI counts, preview and actual removal."""
    def __init__(self, catalog: Catalog):
        self.exact, self.decorated, self.matrix, self.by_sheet = {}, {}, {}, {}
        for e in catalog.entries:
            self.exact.setdefault(key(e.title), []).append(e)
            self.decorated.setdefault(contents_key(e.title), []).append(e)
            row_key = question_row_key(e.title)
            if row_key:
                self.matrix.setdefault(row_key, []).append(e)
            if e.sheet:
                self.by_sheet.setdefault(key(e.sheet), []).append(e)

    def find(self, wanted: Entry) -> tuple[list[Entry], str]:
        matches, mode = self.exact.get(key(wanted.title), []), 'exact'
        if not matches:
            matches, mode = self.decorated.get(contents_key(wanted.title), []), 'decorated'
        if not matches:
            matches, mode = self.matrix.get(question_row_key(wanted.title), []), 'matrix'
        if not matches and wanted.sheet:
            matches, mode = self.by_sheet.get(key(wanted.sheet), []), 'sheet'
        if len({key(e.sheet) for e in matches}) > 1:
            matches = [e for e in matches if wanted.sheet and key(e.sheet) == key(wanted.sheet)]
            if not matches:
                return [], 'ambiguous'
        return matches, mode


def read_catalog(path: Path, setting: bool = False) -> Catalog:
    with ZipFile(path) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        targets = {r.get('Id'): posixpath.normpath(posixpath.join('xl', r.get('Target', ''))) if not r.get('Target', '').startswith('/') else r.get('Target')[1:] for r in rels}
        sheets = {s.get('name'): (s.get(f'{{{REL}}}id'), targets[s.get(f'{{{REL}}}id')]) for s in wb.find('m:sheets', N)}
        contents = next((s for s in sheets if key(s) in ('contents', 'contests')), '')
        if setting and not contents:
            contents = next(iter(sheets))
        if not contents:
            raise ValueError('ไม่พบชีท Contents หรือ Contests')
        part = sheets[contents][1]
        root = ET.fromstring(z.read(part))
        strings = [''.join(t.text or '' for t in si.iter(f'{{{MAIN}}}t')) for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', N)] if 'xl/sharedStrings.xml' in z.namelist() else []
        links = {h.get('ref'): link_sheet(h.get('location', '')) for h in root.findall('m:hyperlinks/m:hyperlink', N)}
        entries = []
        for row in root.findall('m:sheetData/m:row', N):
            cells = {re.sub(r'\d', '', c.get('r', '')): c for c in row}
            def value(col):
                c = cells.get(col)
                if c is None:
                    return ''
                v = c.find('m:v', N)
                text = v.text or '' if v is not None else ''
                if c.get('t') == 's':
                    return strings[int(text)]
                if c.get('t') == 'inlineStr':
                    return ''.join(t.text or '' for t in c.iter(f'{{{MAIN}}}t'))
                return text
            title = value('C').strip()
            if not title or key(title) in ('question', 'contents text', 'ข้อความ contents', 'รายการคำถาม'):
                continue
            r = int(row.get('r'))
            sheet = links.get(f'B{r}') or links.get(f'C{r}') or ''
            if not sheet and 'B' in cells:
                f = cells['B'].find('m:f', N)
                if f is not None and f.text:
                    match = re.search(r'HYPERLINK\s*\(\s*"([^"]+)"', f.text, re.I)
                    if match:
                        sheet = link_sheet(match[1])
            # New settings store the explicit worksheet name in B.
            if setting and not sheet and value('B') not in ('Table', ''):
                sheet = value('B').strip()
            if not sheet and title in sheets:
                sheet = title
            entries.append(Entry(title, sheet, r))
        if not entries:
            raise ValueError('ไม่พบรายการในคอลัมน์ C')
        return Catalog(Path(path), contents, part, entries, sheets)


def save_settings(path: Path, entries: list[Entry]):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = 'Settings'
    ws.append(['ลำดับ', 'Sheet', 'ข้อความ Contents'])
    for i, e in enumerate(entries, 1):
        ws.append([i, e.sheet, e.title])
        # Treat all user text as literal, including strings starting with '='.
        ws.cell(i + 1, 2).data_type = 's'
        ws.cell(i + 1, 3).data_type = 's'
    for c in ws[1]:
        c.font = Font(name='Tahoma', bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='163D48')
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical='top', wrap_text=True)
        ws.row_dimensions[row[0].row].height = 42
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 26
    ws.column_dimensions['C'].width = 120
    ws.freeze_panes = 'C2'
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


@dataclass
class Plan:
    catalog: Catalog
    removed: set[str] = field(default_factory=set)
    rows: set[int] = field(default_factory=set)
    warnings: list[str] = field(default_factory=list)
    infos: list[str] = field(default_factory=list)
    matched_items: int = 0
    unmatched_items: int = 0
    source_digest: str = ''


def plan_removal(catalog: Catalog, selected: list[Entry]) -> Plan:
    plan = Plan(catalog)
    matcher = ContentsMatcher(catalog)
    for wanted in selected:
        matches, mode = matcher.find(wanted)
        if not matches:
            if mode == 'ambiguous':
                plan.warnings.append(f'รายการซ้ำระบุชีทไม่ได้: {wanted.title}')
            else:
                plan.infos.append(f'ข้าม: ไม่พบข้อความ / ลิงก์ชีทใน Contents (ชีทที่ระบุ: {wanted.sheet or "ค้นจากข้อความ"}): {wanted.title}')
            plan.unmatched_items += 1
            continue
        found = False
        for e in matches:
            actual = next((s for s in catalog.sheets if key(s) == key(e.sheet)), '')
            if not actual:
                plan.infos.append(f'ข้าม: ไม่พบชีท / ลิงก์ Table: {e.sheet or "(ไม่มีลิงก์)"} — {e.title}')
            elif actual == catalog.contents:
                plan.warnings.append(f'ไม่อนุญาตให้ตัดชีทสารบัญ: {actual}')
            else:
                plan.removed.add(actual)
                found = True
        if found:
            plan.matched_items += 1
            if mode == 'matrix':
                row_key = question_row_key(wanted.title)
                plan.infos.append(f'จับคู่ข้อความย่อสำเร็จ {row_key[0].upper()} R{row_key[1]} → {", ".join(sorted({e.sheet for e in matches}))}')
            elif mode == 'decorated':
                plan.infos.append(f'จับคู่สำเร็จ (ละ * / ชนิดคำถามท้ายข้อความ): {wanted.title}')
            elif mode == 'sheet':
                plan.warnings.append(f'ข้อความต่างจาก Setting ใช้ลิงก์ชีท {wanted.sheet} ใน Contents: {wanted.title}')
        else:
            plan.unmatched_items += 1
    # Remove every Contents entry pointing to a removed sheet, including duplicates.
    plan.rows = {e.row for e in catalog.entries if key(e.sheet) in {key(s) for s in plan.removed}}
    return plan


def nodes(doc, name):
    return list(doc.getElementsByTagNameNS(MAIN, name))


def xml_bytes(doc):
    return doc.toxml(encoding='utf-8')


class RowDeletion:
    def __init__(self, rows):
        self.rows = sorted(rows)
        self.removed = set(rows)

    def shift(self, row):
        return row - bisect_left(self.rows, row)

    def ref(self, reference):
        """Map cell/range coordinates through actual row deletions."""
        match = re.fullmatch(r'(\$?[A-Z]+)(\$?)(\d+)(?::(\$?[A-Z]+)(\$?)(\d+))?', reference, re.I)
        if not match:
            return reference
        start = int(match[3])
        if match[6] is None:
            return '' if start in self.removed else f'{match[1]}{match[2]}{self.shift(start)}'
        end = int(match[6])
        while start <= end and start in self.removed: start += 1
        while end >= start and end in self.removed: end -= 1
        if start > end:
            return ''
        return f'{match[1]}{match[2]}{self.shift(start)}:{match[4]}{match[5]}{self.shift(end)}'

    def formula(self, formula, contents_name, local=False):
        from openpyxl.formula.tokenizer import Tokenizer
        had_equals = formula.startswith('=')
        tokenizer = Tokenizer(formula if had_equals else '=' + formula)
        for token in tokenizer.items:
            if token.type != 'OPERAND' or token.subtype != 'RANGE':
                continue
            prefix, separator, reference = token.value.rpartition('!')
            if separator:
                if key(prefix.strip("'").replace("''", "'")) != key(contents_name):
                    continue
            elif local:
                reference = token.value
            else:
                continue
            mapped = self.ref(reference)
            token.value = (prefix + '!' if separator else '') + mapped if mapped else '#REF!'
        result = ''.join(t.value for t in tokenizer.items)
        return ('=' if had_equals else '') + result


def compact_contents(doc, deletion, contents_name):
    for row in nodes(doc, 'row'):
        old = int(row.getAttribute('r'))
        if old in deletion.removed:
            row.parentNode.removeChild(row)
        else:
            row.setAttribute('r', str(deletion.shift(old)))
            for cell in row.getElementsByTagNameNS(MAIN, 'c'):
                cell.setAttribute('r', deletion.ref(cell.getAttribute('r')))
    for element in list(doc.getElementsByTagNameNS(MAIN, '*')):
        if element.localName in ('c', 'row'):
            continue
        for attribute in ('ref', 'sqref', 'activeCell', 'topLeftCell'):
            if not element.hasAttribute(attribute):
                continue
            mapped = ' '.join(filter(None, (deletion.ref(r) for r in element.getAttribute(attribute).split())))
            if mapped:
                element.setAttribute(attribute, mapped)
            elif element.localName == 'selection':
                element.setAttribute(attribute, 'A1')
            elif element.parentNode is not None:
                element.parentNode.removeChild(element)
                break
        if element.localName == 'pane' and element.hasAttribute('ySplit'):
            split = float(element.getAttribute('ySplit'))
            if split.is_integer():
                element.setAttribute('ySplit', str(int(split) - bisect_right(deletion.rows, split)))
        if element.localName == 'brk' and element.parentNode.localName == 'rowBreaks':
            element.setAttribute('id', str(deletion.shift(int(element.getAttribute('id')))))
    for name in ('mergeCells', 'dataValidations', 'rowBreaks'):
        for group in nodes(doc, name):
            group.setAttribute('count', str(sum(child.nodeType == child.ELEMENT_NODE for child in group.childNodes)))
    for f in nodes(doc, 'f') + nodes(doc, 'formula') + nodes(doc, 'formula1') + nodes(doc, 'formula2'):
        if f.firstChild and f.firstChild.nodeType == f.TEXT_NODE:
            f.firstChild.data = deletion.formula(f.firstChild.data, contents_name, local=True)
    shift_contents_links(doc, deletion, contents_name)
    return doc


def shift_contents_links(doc, deletion, contents_name):
    changed = False
    for link in nodes(doc, 'hyperlink'):
        location = link.getAttribute('location')
        if key(link_sheet(location)) == key(contents_name):
            prefix, _, reference = location.rpartition('!')
            mapped = deletion.ref(reference)
            if mapped != reference:
                changed = True
                if mapped: link.setAttribute('location', prefix + '!' + mapped)
                else: link.parentNode.removeChild(link)
    return changed


def validate_plan(plan):
    """All issues block deletion, including formulas in sheets being kept."""
    issues = list(plan.warnings)
    c = plan.catalog
    with ZipFile(c.path) as z:
        bad = z.testzip()
        if bad: issues.append(f'ไฟล์เสีย: {bad}')
        root = ET.fromstring(z.read('xl/workbook.xml'))
        if not any(s.get('name') not in plan.removed and s.get('state', 'visible') == 'visible' for s in root.find('m:sheets', N)):
            issues.append('ต้องเหลือชีทที่มองเห็นอย่างน้อย 1 ชีท')
        patterns = [re.compile(r"(?:'" + re.escape(s.replace("'", "''")) + r"'|(?<![\w.'])" + re.escape(s) + r")!", re.I) for s in plan.removed]
        deletion = RowDeletion(plan.rows)
        for name, (_, part) in c.sheets.items():
            sheet = ET.fromstring(z.read(part))
            if name in plan.removed: continue
            if name == c.contents:
                if plan.rows and any(sheet.find('m:' + tag, N) is not None for tag in ('drawing', 'legacyDrawing', 'tableParts')):
                    issues.append(f'ชีท {name}: มีรูปภาพหรือ Excel Table ที่ยังไม่รองรับการเลื่อนแถว')
                deleted_cells = {cell.get('r') for row in sheet.findall('m:sheetData/m:row', N) if int(row.get('r')) in plan.rows for cell in row}
            else:
                deleted_cells = set()
            for cell in sheet.findall('.//m:c', N):
                if cell.get('r') in deleted_cells: continue
                for formula in cell.findall('m:f', N):
                    if any(p.search(formula.text or '') for p in patterns):
                        issues.append(f'ชีท {name} เซลล์ {cell.get("r")}: สูตรอ้างอิงชีทที่จะตัด')
                    shifted = deletion.formula(formula.text or '', c.contents, local=name == c.contents)
                    if '#REF!' in shifted and '#REF!' not in (formula.text or ''):
                        issues.append(f'ชีท {name} เซลล์ {cell.get("r")}: สูตรอ้างอิงแถว Contents ที่จะลบ')
    return issues


@dataclass(frozen=True)
class BatchIssue:
    path: Path
    detail: str


def preflight_batch(paths, selected):
    plans, issues = [], []
    for path in paths:
        path = Path(path)
        try:
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            plan = plan_removal(read_catalog(path), selected)
            issues.extend(BatchIssue(path, detail) for detail in validate_plan(plan))
            if hashlib.sha256(path.read_bytes()).hexdigest() != digest:
                issues.append(BatchIssue(path, 'ไฟล์เปลี่ยนระหว่างตรวจสอบ กรุณาลองใหม่'))
            plan.source_digest = digest
            plans.append(plan)
        except Exception as exc:
            issues.append(BatchIssue(path, str(exc)))
    return plans, issues


def write_trimmed(plan: Plan, destination: Path) -> list[str]:
    """Keep untouched ZIP members byte-identical after decompression.

    Contents rows, ranges and links move up like Excel Delete Sheet Rows.
    """
    c = plan.catalog
    if not plan.removed:
        raise ValueError('ไม่มีชีทที่ตัดได้ จึงไม่สร้างสำเนา')
    if destination.resolve() == c.path.resolve():
        raise ValueError('ห้ามบันทึกทับไฟล์ต้นฉบับ')
    if destination.exists():
        raise FileExistsError(f'มีไฟล์ปลายทางแล้ว: {destination}')
    if plan.source_digest and hashlib.sha256(c.path.read_bytes()).hexdigest() != plan.source_digest:
        raise ValueError('ไฟล์ต้นทางเปลี่ยนหลังตรวจสอบ กรุณาลองใหม่')
    issues = validate_plan(plan)
    if issues:
        raise ValueError('\n'.join(issues))
    warnings = []
    with ZipFile(c.path) as src:
        wb = minidom.parseString(src.read('xl/workbook.xml'))
        original = nodes(wb, 'sheet')
        kept = [s for s in original if s.getAttribute('name') not in plan.removed]
        if not any(s.getAttribute('state') in ('', 'visible') for s in kept):
            raise ValueError('ต้องเหลือชีทที่มองเห็นอย่างน้อย 1 ชีท')
        removed_ids = {c.sheets[s][0] for s in plan.removed}
        removed_parts = {c.sheets[s][1] for s in plan.removed}
        skip = set(removed_parts)
        skip.update(posixpath.join(posixpath.dirname(p), '_rels', posixpath.basename(p) + '.rels') for p in removed_parts)
        for s in original:
            if s not in kept:
                s.parentNode.removeChild(s)
        index_map = {original.index(s): i for i, s in enumerate(kept)}
        for dn in nodes(wb, 'definedName'):
            local = dn.getAttribute('localSheetId')
            text = ''.join(n.data for n in dn.childNodes if n.nodeType == n.TEXT_NODE)
            refers_removed = any(re.search(r"(?:'" + re.escape(s.replace("'", "''")) + r"'|(?<![\w.'])" + re.escape(s) + r")!", text, re.I) for s in plan.removed)
            if (local and int(local) not in index_map) or refers_removed:
                dn.parentNode.removeChild(dn)
            elif local:
                dn.setAttribute('localSheetId', str(index_map[int(local)]))
        first_visible = next(i for i, s in enumerate(kept) if s.getAttribute('state') in ('', 'visible'))
        for view in nodes(wb, 'workbookView'):
            view.setAttribute('activeTab', str(first_visible))
            view.setAttribute('firstSheet', '0')
        for calc in nodes(wb, 'calcPr'):
            calc.setAttribute('fullCalcOnLoad', '1')
        rels = minidom.parseString(src.read('xl/_rels/workbook.xml.rels'))
        for r in list(rels.documentElement.childNodes):
            if r.nodeType != r.ELEMENT_NODE:
                continue
            if r.getAttribute('Id') in removed_ids or r.getAttribute('Type').endswith('/calcChain'):
                if r.getAttribute('Type').endswith('/calcChain'):
                    target = r.getAttribute('Target')
                    skip.add(target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target)))
                r.parentNode.removeChild(r)
        skip.add('xl/calcChain.xml')
        ct = minidom.parseString(src.read('[Content_Types].xml'))
        for r in list(ct.documentElement.childNodes):
            if r.nodeType == r.ELEMENT_NODE and r.getAttribute('PartName').lstrip('/') in skip:
                r.parentNode.removeChild(r)
        deletion = RowDeletion(plan.rows)
        contents = compact_contents(minidom.parseString(src.read(c.contents_part)), deletion, c.contents)
        for dn in nodes(wb, 'definedName'):
            if dn.firstChild:
                dn.firstChild.data = deletion.formula(dn.firstChild.data, c.contents)
        changes = {'xl/workbook.xml': xml_bytes(wb), 'xl/_rels/workbook.xml.rels': xml_bytes(rels), '[Content_Types].xml': xml_bytes(ct), c.contents_part: xml_bytes(contents)}
        # Adjust only references to moved Contents rows in surviving worksheets.
        for name, (_, part) in c.sheets.items():
            if name in plan.removed or name == c.contents:
                continue
            data = src.read(part)
            if b'<f' not in data and b'hyperlink' not in data:
                continue
            doc = minidom.parseString(data)
            changed = shift_contents_links(doc, deletion, c.contents)
            for f in nodes(doc, 'f'):
                if f.firstChild:
                    before = f.firstChild.data
                    after = deletion.formula(before, c.contents)
                    if before != after:
                        f.firstChild.data = after
                        changed = True
            if changed: changes[part] = xml_bytes(doc)
        destination.parent.mkdir(parents=True, exist_ok=True)
        fd, temp = tempfile.mkstemp(suffix='.xlsx', dir=destination.parent)
        os.close(fd)
        try:
            with ZipFile(temp, 'w') as out:
                for info in src.infolist():
                    if info.filename not in skip:
                        out.writestr(info, changes.get(info.filename, src.read(info.filename)))
            with ZipFile(temp) as check:
                if check.testzip():
                    raise ValueError('ตรวจสอบไฟล์ผลลัพธ์ไม่ผ่าน')
            # Windows rename refuses to overwrite existing files.
            os.rename(temp, destination)
        finally:
            if os.path.exists(temp):
                os.unlink(temp)
    return warnings

# ===== Desktop GUI (editable source) =====

BASE = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
INPUT = BASE / '2_Table Test' if (BASE / '2_Table Test').is_dir() else BASE
APP_ICON: Path  # Initialized below from the embedded icon before main runs.
INK = '#173D48'
TEAL = '#137F76'
BG = '#F1F5F5'


def ask_native_dialog_centered(dialog, **options):
    """Center a Windows file/folder dialog after its initial layout."""
    if sys.platform != 'win32':
        return dialog(**options)
    import ctypes
    from ctypes import wintypes

    class MonitorInfo(ctypes.Structure):
        _fields_ = [('cbSize', wintypes.DWORD), ('rcMonitor', wintypes.RECT),
                    ('rcWork', wintypes.RECT), ('dwFlags', wintypes.DWORD)]

    user32 = ctypes.WinDLL('user32', use_last_error=True)
    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
    hook_proc = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
    user32.SetWindowsHookExW.argtypes = [ctypes.c_int, hook_proc, wintypes.HINSTANCE, wintypes.DWORD]
    user32.SetWindowsHookExW.restype = wintypes.HANDLE
    user32.CallNextHookEx.argtypes = [wintypes.HANDLE, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM]
    user32.CallNextHookEx.restype = ctypes.c_ssize_t
    user32.UnhookWindowsHookEx.argtypes = [wintypes.HANDLE]
    user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
    user32.MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
    user32.MonitorFromWindow.restype = wintypes.HANDLE
    user32.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.POINTER(MonitorInfo)]
    user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                  ctypes.c_int, ctypes.c_int, wintypes.UINT]
    kernel32.GetCurrentThreadId.restype = wintypes.DWORD
    positioned = set()
    timers = {}
    timer_proc = ctypes.WINFUNCTYPE(None, wintypes.HWND, wintypes.UINT, ctypes.c_size_t, wintypes.DWORD)
    user32.SetTimer.argtypes = [wintypes.HWND, ctypes.c_size_t, wintypes.UINT, timer_proc]
    user32.SetTimer.restype = ctypes.c_size_t
    user32.KillTimer.argtypes = [wintypes.HWND, ctypes.c_size_t]

    def position(hwnd):
        rect = wintypes.RECT()
        info = MonitorInfo()
        info.cbSize = ctypes.sizeof(info)
        monitor = user32.MonitorFromWindow(hwnd, 2)
        if user32.GetWindowRect(hwnd, ctypes.byref(rect)) and user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
            work = info.rcWork
            x = work.left + ((work.right - work.left) - (rect.right - rect.left)) // 2
            y = work.top + ((work.bottom - work.top) - (rect.bottom - rect.top)) // 2
            user32.SetWindowPos(hwnd, None, max(work.left, x), max(work.top, y), 0, 0, 0x0015)

    @timer_proc
    def after_layout(_hwnd, _message, timer, _time):
        user32.KillTimer(None, timer)
        hwnd = timers.pop(timer, None)
        if hwnd is not None:
            position(hwnd)

    @hook_proc
    def on_activate(code, hwnd, parameter):
        if code == 5 and hwnd not in positioned:  # HCBT_ACTIVATE
            name = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, name, len(name))
            if name.value == '#32770':
                position(hwnd)
                positioned.add(hwnd)
                # Explorer restores its own position during initial layout.
                # A native timer runs inside the common dialog's modal loop.
                timer = user32.SetTimer(None, 0, 150, after_layout)
                if timer:
                    timers[timer] = hwnd
        return user32.CallNextHookEx(None, code, hwnd, parameter)

    hook = user32.SetWindowsHookExW(5, on_activate, None, kernel32.GetCurrentThreadId())  # WH_CBT
    if not hook:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return dialog(**options)
    finally:
        user32.UnhookWindowsHookEx(hook)
        for timer in timers:
            user32.KillTimer(None, timer)


def ask_directory_centered(**options):
    return ask_native_dialog_centered(filedialog.askdirectory, **options)


def output_paths(plans, folder, suffix):
    suffix = suffix.strip()
    if any(c in '<>:"/\\|?*' or ord(c) < 32 for c in suffix) or suffix.endswith('.'):
        raise ValueError('ข้อความต่อท้ายมีอักขระที่ใช้ในชื่อไฟล์ไม่ได้: < > : " / \\ | ? * หรือจุดท้ายข้อความ')
    sources = {str(p.catalog.path.resolve()).casefold() for p in plans}
    destinations, seen = [], set()
    for plan in plans:
        source = plan.catalog.path
        name = source.stem + (' ' + suffix if suffix else '') + source.suffix
        target = folder / name
        identity = str(target.resolve()).casefold()
        if len(name) > 255:
            raise ValueError(f'ชื่อไฟล์ยาวเกิน 255 ตัวอักษร: {name}')
        if identity in sources:
            raise ValueError(f'ห้ามเขียนทับไฟล์ต้นฉบับ: {target}\nเลือกโฟลเดอร์อื่นหรือใส่ข้อความต่อท้าย')
        if identity in seen:
            raise ValueError(f'ไฟล์ต้นทางมีชื่อผลลัพธ์ซ้ำกัน: {name}\nกรุณาแยกรันไฟล์เหล่านี้')
        if target.exists():
            raise ValueError(f'มีไฟล์ชื่อนี้อยู่แล้ว จะไม่เขียนทับ: {target}')
        seen.add(identity)
        destinations.append(target)
    return destinations


def save_batch(plans, destinations, notify=lambda *_: None):
    """Stage all work, then publish with exclusive creation and rollback."""
    folder = destinations[0].parent
    published = []
    with tempfile.TemporaryDirectory(prefix='.table_stage_', dir=folder) as staging:
        stage = Path(staging)
        for j, (plan, destination) in enumerate(zip(plans, destinations)):
            try:
                write_trimmed(plan, stage / destination.name)
            except Exception as exc:
                raise ValueError(f'ไฟล์ {plan.catalog.path}: {exc}\nยกเลิกทั้งชุด ไม่มีการบันทึกผลลัพธ์') from exc
            notify(j, plan)
        for plan in plans:
            if hashlib.sha256(plan.catalog.path.read_bytes()).hexdigest() != plan.source_digest:
                raise ValueError(f'ไฟล์ {plan.catalog.path} เปลี่ยนระหว่างทำงาน ยกเลิกทั้งชุด')
        try:
            for destination in destinations:
                # Exclusive creation also protects files created after the dialogs.
                with destination.open('xb') as output:
                    stat = os.fstat(output.fileno())
                    published.append((destination, (stat.st_dev, stat.st_ino)))
                    with (stage / destination.name).open('rb') as source:
                        shutil.copyfileobj(source, output)
                    output.flush()
                    os.fsync(output.fileno())
        except Exception as exc:
            cleanup_errors = []
            for target, identity in published:
                try:
                    if target.exists():
                        stat = target.stat()
                        if (stat.st_dev, stat.st_ino) == identity:
                            target.unlink()
                except OSError as cleanup:
                    cleanup_errors.append(f'{target}: {cleanup}')
            detail = '\nลบผลลัพธ์ที่บันทึกค้างไม่สำเร็จ:\n' + '\n'.join(cleanup_errors) if cleanup_errors else '\nยกเลิกผลลัพธ์ทั้งชุด'
            raise ValueError(f'บันทึกไฟล์ {destination} ไม่สำเร็จ: {exc}{detail}') from exc


def center_window(window, parent=None):
    """Position using physical pixels, so Windows DPI scaling stays correct."""
    window.update_idletasks()
    if sys.platform == 'win32':
        import ctypes
        from ctypes import wintypes

        class MonitorInfo(ctypes.Structure):
            _fields_ = [('cbSize', wintypes.DWORD), ('rcMonitor', wintypes.RECT),
                        ('rcWork', wintypes.RECT), ('dwFlags', wintypes.DWORD)]

        user32 = ctypes.WinDLL('user32', use_last_error=True)
        user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
        user32.GetAncestor.restype = wintypes.HWND
        user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
        user32.GetWindowRect.restype = wintypes.BOOL
        user32.MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
        user32.MonitorFromWindow.restype = wintypes.HANDLE
        user32.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.POINTER(MonitorInfo)]
        user32.GetMonitorInfoW.restype = wintypes.BOOL
        user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int,
                                      ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.UINT]
        user32.SetWindowPos.restype = wintypes.BOOL
        hwnd = user32.GetAncestor(window.winfo_id(), 2)  # GA_ROOT
        rect = wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            raise ctypes.WinError(ctypes.get_last_error())
        if parent is not None:
            parent.update_idletasks()
            target = wintypes.RECT()
            parent_hwnd = user32.GetAncestor(parent.winfo_id(), 2)
            if not user32.GetWindowRect(parent_hwnd, ctypes.byref(target)):
                raise ctypes.WinError(ctypes.get_last_error())
        else:
            info = MonitorInfo()
            info.cbSize = ctypes.sizeof(info)
            monitor = user32.MonitorFromWindow(hwnd, 2)  # nearest monitor
            if not user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
                raise ctypes.WinError(ctypes.get_last_error())
            target = info.rcWork  # Exclude the taskbar; support non-primary monitors.
        x = target.left + ((target.right - target.left) - (rect.right - rect.left)) // 2
        y = target.top + ((target.bottom - target.top) - (rect.bottom - rect.top)) // 2
        if not user32.SetWindowPos(hwnd, None, x, y, 0, 0, 0x0015):
            raise ctypes.WinError(ctypes.get_last_error())
        return
    width, height = window.winfo_width(), window.winfo_height()
    border = max(0, window.winfo_rootx() - window.winfo_x())
    title = max(0, window.winfo_rooty() - window.winfo_y())
    width += 2 * border
    height += title + border
    if parent is None:
        x = (window.winfo_screenwidth() - width) // 2
        y = (window.winfo_screenheight() - height) // 2
    else:
        parent.update_idletasks()
        parent_border = max(0, parent.winfo_rootx() - parent.winfo_x())
        parent_title = max(0, parent.winfo_rooty() - parent.winfo_y())
        x = parent.winfo_x() + (parent.winfo_width() + 2 * parent_border - width) // 2
        y = parent.winfo_y() + (parent.winfo_height() + parent_title + parent_border - height) // 2
    window.geometry(f'+{max(0, x)}+{max(0, y)}')


class StudioIcon:
    def iconbitmap(self, bitmap=None, default=None):
        # CTk schedules its own icon after 200 ms. Keep our icon on those calls too.
        if sys.platform == 'win32' and APP_ICON.is_file():
            bitmap = str(APP_ICON)
            if default is not None:
                default = str(APP_ICON)
        return super().iconbitmap(bitmap=bitmap, default=default)

    def _windows_set_titlebar_icon(self):
        if sys.platform == 'win32' and APP_ICON.is_file():
            self.iconbitmap(str(APP_ICON))


class StudioToplevel(StudioIcon, ctk.CTkToplevel):
    pass


class App(StudioIcon, ctk.CTk):
    def __init__(self, prompt_source=False):
        super().__init__()
        if sys.platform == 'win32' and APP_ICON.is_file():
            self.iconbitmap(default=str(APP_ICON))
        # Keep initial mapping and DPI/layout changes invisible until positioned.
        self.attributes('-alpha', 0.0)
        self.title('Cut MKT • ตัดชีท Excel')
        self.geometry('1160x720')
        self.minsize(1000, 660)
        self.configure(fg_color=BG)
        self.catalogs = []
        self.source_folder = None
        self.source_paths = []
        self.requested_paths = []
        self.entries: dict[tuple[str, str], Entry] = {}
        self.selected: set[tuple[str, str]] = set()
        self.events = queue.Queue()
        self.busy = False
        self.actions = []
        self.last_output = None
        self.search = tk.StringVar()
        self.count = tk.StringVar(value='ยังไม่ได้เลือกชีท')
        self.status = tk.StringVar(value='พร้อมทำงาน')
        self.build_ui()
        self.search.trace_add('write', lambda *_: self.render_entries())
        self.after(100, self.pump)
        if prompt_source:
            self.after(250, self.import_files)
        self.protocol('WM_DELETE_WINDOW', self.close)
        # CTk applies the initial DPI-adjusted size after the first UI events.
        self.after(150, self.show_startup)

    def show_startup(self):
        center_window(self)
        self.attributes('-alpha', 1.0)

    @staticmethod
    def ident(e):
        return key(e.title), key(e.sheet)

    def button(self, parent, text, command, primary=False):
        b = ctk.CTkButton(parent, text=text, command=command, height=38, corner_radius=8,
                          font=('Tahoma', 13), fg_color=TEAL if primary else '#E5EEEE',
                          text_color='white' if primary else INK,
                          hover_color='#0D6861' if primary else '#D4E3E3')
        self.actions.append(b)
        return b

    def build_ui(self):
        header = ctk.CTkFrame(self, fg_color=INK, corner_radius=0)
        header.pack(fill='x')
        brand = ctk.CTkFrame(header, fg_color='transparent')
        brand.pack(fill='x', padx=28, pady=18)
        if APP_ICON.is_file():
            with Image.open(APP_ICON) as icon:
                logo = icon.convert('RGBA')
            self.header_logo_image = ctk.CTkImage(light_image=logo, dark_image=logo, size=(72, 72))
            self.header_logo = ctk.CTkLabel(brand, text='', image=self.header_logo_image, width=72, height=72)
            self.header_logo.pack(side='left', padx=(0, 18))
        brand_text = ctk.CTkFrame(brand, fg_color='transparent')
        brand_text.pack(side='left', fill='x', expand=True)
        ctk.CTkLabel(brand_text, text='CUT MKT', text_color='#86D5C6', font=('Segoe UI', 13, 'bold')).pack(anchor='w')
        ctk.CTkLabel(brand_text, text='เลือกชีทที่ต้องการตัดออก', text_color='white', font=('Tahoma', 27, 'bold')).pack(anchor='w', pady=(2, 5))
        ctk.CTkLabel(brand_text, text='อ่าน Contents คอลัมน์ C  •  บันทึก Setting  •  ทำงานพร้อมกันหลายไฟล์', text_color='#C4D8DC', font=('Tahoma', 13)).pack(anchor='w')
        body = ctk.CTkFrame(self, fg_color='transparent')
        body.pack(fill='both', expand=True, padx=22, pady=18)
        body.grid_columnconfigure(0, weight=0, minsize=290)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)
        left = ctk.CTkFrame(body, fg_color='white', corner_radius=12, width=290)
        left.grid(row=0, column=0, sticky='nsew', padx=(0, 16))
        ctk.CTkLabel(left, text='01   ไฟล์ Table', text_color=INK, font=('Tahoma', 17, 'bold')).pack(anchor='w', padx=18, pady=(18, 8))
        self.source_label = ctk.CTkLabel(left, text='เลือก Excel ที่ต้องการตัดชีท\nรองรับ .xlsx / .xlsm', text_color='#526D74', justify='left', wraplength=245, font=('Tahoma', 12))
        self.source_label.pack(anchor='w', padx=18)
        self.source_button = self.button(left, 'เลือกไฟล์ Excel', self.import_files, True)
        self.source_button.configure(height=44, font=('Tahoma', 13, 'bold'))
        self.source_button.pack(fill='x', padx=18, pady=(12, 5))
        ctk.CTkLabel(left, text='เลือกหลายไฟล์พร้อมกันได้', height=18, text_color='#526D74', font=('Tahoma', 11)).pack(anchor='w', padx=18, pady=(0, 8))
        self.file_list = tk.Listbox(left, bg='#F5F8F8', fg=INK, selectbackground='#CDE6E1', relief='flat', borderwidth=0, font=('Tahoma', 10), height=5, exportselection=False)
        self.file_list.pack(fill='both', expand=True, padx=18)
        self.file_list.insert('end', 'ยังไม่ได้เลือกไฟล์')
        self.file_list.bind('<<ListboxSelect>>', self.file_detail)
        self.file_label = ctk.CTkLabel(left, text='0 ไฟล์', text_color='#647A80', wraplength=245, justify='left', font=('Tahoma', 11))
        settings_panel = ctk.CTkFrame(left, fg_color='#F0F7F5', corner_radius=12)
        settings_panel.pack(side='bottom', fill='x', padx=12, pady=(12, 12), before=self.file_list)
        ctk.CTkLabel(settings_panel, text='02   รายการตัด (Setting)', text_color=INK, font=('Tahoma', 15, 'bold')).pack(anchor='w', padx=12, pady=(10, 4))
        self.load_setting_button = self.button(settings_panel, 'โหลด Setting จาก Excel', self.load_setting, True)
        self.load_setting_button.configure(height=44, font=('Tahoma', 13, 'bold'))
        self.load_setting_button.pack(fill='x', padx=12, pady=(6, 0))
        self.setting_file_label = ctk.CTkLabel(settings_panel, text='นำรายการที่บันทึกไว้มาใช้', height=18, text_color='#526D74', wraplength=220, justify='left', font=('Tahoma', 11))
        self.setting_file_label.pack(anchor='w', padx=12, pady=(3, 8))
        self.save_setting_button = self.button(settings_panel, 'บันทึก Setting เป็น Excel', self.save_setting)
        self.save_setting_button.configure(height=44, font=('Tahoma', 13, 'bold'), fg_color=INK, text_color='white', text_color_disabled='#9FB7BE', hover_color='#245461', border_width=0)
        self.save_setting_button.pack(fill='x', padx=12)
        self.setting_hint = ctk.CTkLabel(settings_panel, text='เลือกชีทในรายการก่อนบันทึก', height=18, text_color='#526D74', wraplength=220, justify='left', font=('Tahoma', 11))
        self.setting_hint.pack(anchor='w', padx=12, pady=(3, 10))
        self.save_setting_button.configure(state='disabled')
        self.file_label.pack(side='bottom', anchor='w', padx=18, pady=8, before=self.file_list)
        right = ctk.CTkFrame(body, fg_color='white', corner_radius=12)
        right.grid(row=0, column=1, sticky='nsew')
        top = ctk.CTkFrame(right, fg_color='transparent')
        top.pack(fill='x', padx=18, pady=(16, 10))
        ctk.CTkLabel(top, text='03   รายการจาก Contents', text_color=INK, font=('Tahoma', 17, 'bold')).pack(side='left')
        ctk.CTkLabel(top, textvariable=self.count, text_color=TEAL, font=('Tahoma', 12, 'bold')).pack(side='right')
        ctk.CTkLabel(right, text='ค้นหาคำถามหรือชื่อชีท', text_color='#647A80', font=('Tahoma', 11)).pack(anchor='w', padx=18)
        ctk.CTkEntry(right, textvariable=self.search, height=38, font=('Tahoma', 13)).pack(fill='x', padx=18)
        tools = ctk.CTkFrame(right, fg_color='transparent')
        tools.pack(fill='x', padx=18, pady=10)
        self.button(tools, 'เลือกที่แสดงทั้งหมด', lambda: self.select_visible(True)).pack(side='left', padx=(0, 8))
        self.button(tools, 'ยกเลิกที่แสดง', lambda: self.select_visible(False)).pack(side='left')
        ctk.CTkLabel(tools, text='คลิกแถว / Space เพื่อเลือก', text_color='#647A80', font=('Tahoma', 11)).pack(side='right')
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('Treeview', font=('Tahoma', 10), rowheight=22, background='white', fieldbackground='white', foreground=INK, borderwidth=0)
        style.configure('Treeview.Heading', font=('Tahoma', 10, 'bold'), background='#EAF1F1', foreground=INK, relief='flat')
        style.map('Treeview', background=[('selected', '#DCEFEA')], foreground=[('selected', INK)])
        self.detail_panes = tk.PanedWindow(right, orient='horizontal', sashwidth=8, sashrelief='flat', bg='#D3E2E2', borderwidth=0, opaqueresize=True, cursor='sb_h_double_arrow')
        self.detail_panes.pack(fill='both', expand=True, padx=18, pady=(0, 16))
        table = ctk.CTkFrame(self.detail_panes, fg_color='transparent')
        self.detail_panes.add(table, minsize=330, width=560, stretch='always')
        self.tree = ttk.Treeview(table, columns=('check', 'sheet', 'title', 'files'), show='headings', selectmode='browse')
        for name, label, width in [('check', 'ตัด', 42), ('sheet', 'ชีทจริง', 105), ('title', 'ข้อความ Contents', 320), ('files', 'ไฟล์', 45)]:
            self.tree.heading(name, text=label)
            self.tree.column(name, width=width, minwidth=width if name != 'title' else 160, stretch=name == 'title', anchor='center' if name in ('check', 'files') else 'w')
        sy = ttk.Scrollbar(table, orient='vertical', command=self.tree.yview)
        sx = ttk.Scrollbar(table, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        sy.grid(row=0, column=1, sticky='ns')
        sx.grid(row=1, column=0, sticky='ew')
        table.grid_rowconfigure(0, weight=1)
        table.grid_columnconfigure(0, weight=1)
        self.tree.tag_configure('missing', foreground='#B65F27')
        self.tree.bind('<ButtonRelease-1>', self.toggle)
        self.tree.bind('<space>', self.toggle)
        self.tree.bind('<Double-Button-1>', self.show_entry)
        log_panel = ctk.CTkFrame(self.detail_panes, fg_color='#F5F8F8', corner_radius=8)
        self.detail_panes.add(log_panel, minsize=260, width=440, stretch='always')
        ctk.CTkLabel(log_panel, text='Log การทำงาน', text_color=INK, font=('Tahoma', 12, 'bold')).pack(anchor='w', padx=12, pady=(6, 0))
        ctk.CTkLabel(log_panel, text='ลากเส้นแบ่งซ้าย–ขวาเพื่อปรับพื้นที่', text_color='#647A80', font=('Tahoma', 10)).pack(anchor='w', padx=12, pady=(0, 4))
        self.log_box = ctk.CTkTextbox(log_panel, height=210, fg_color='#F5F8F8', text_color=INK, font=('Tahoma', 12), wrap='word')
        self.log_box.pack(fill='both', expand=True, padx=6, pady=(0, 6))
        for tag, color in [('info', INK), ('warning', '#A85C00'), ('error', '#BA2839'), ('success', '#087363')]:
            self.log_box.tag_config(tag, foreground=color)
        self.log_box.configure(state='disabled')
        footer = ctk.CTkFrame(self, fg_color='white', corner_radius=0)
        footer.pack(side='bottom', fill='x', before=body)
        self.progress = ctk.CTkProgressBar(footer, width=180, progress_color=TEAL)
        self.progress.pack(side='left', padx=(24, 12), pady=20)
        self.progress.set(0)
        ctk.CTkLabel(footer, textvariable=self.status, text_color=INK, font=('Tahoma', 12)).pack(side='left')
        self.button(footer, 'ตัดชีทและบันทึกไฟล์ใหม่', self.run_batch, True).pack(side='right', padx=(10, 24), pady=16)
        self.button(footer, 'เปิดผลลัพธ์', self.open_output).pack(side='right', padx=5)

    def log(self, text):
        tag = 'info'
        if text.lstrip().startswith('Error') or any(word in text for word in ('ผิดพลาด', 'อ่านไม่ได้')): tag = 'error'
        elif any(word in text for word in ('คำเตือน', 'ไม่พบ', 'ข้าม', 'ข้อความต่าง', 'รายการซ้ำ')): tag = 'warning'
        elif any(word in text for word in ('สำเร็จ', 'บันทึกแล้ว')): tag = 'success'
        self.log_box.configure(state='normal')
        self.log_box.insert('end', f'[{datetime.now():%H:%M:%S}] {text}\n', tag)
        self.log_box.see('end')
        self.log_box.configure(state='disabled')

    def background(self, job):
        if self.busy:
            return
        self.busy = True
        for b in self.actions:
            b.configure(state='disabled')
        self.status.set('กำลังทำงาน…')
        self.progress.set(0)
        def worker():
            try:
                job()
            except Exception as exc:
                self.events.put(('error', str(exc)))
            finally:
                self.events.put(('idle', None))
        threading.Thread(target=worker, daemon=True).start()

    def pump(self):
        try:
            while True:
                kind, data = self.events.get_nowait()
                if kind == 'log': self.log(data)
                elif kind == 'progress': self.progress.set(data)
                elif kind == 'status': self.status.set(data)
                elif kind == 'sources': self.requested_paths = data
                elif kind == 'setting_name':
                    self.setting_file_label.configure(text=f'โหลดแล้ว: {Path(data).name}')
                elif kind == 'batch_errors': self.show_batch_errors(data)
                elif kind == 'save_options':
                    plans, response = data
                    options = None
                    try:
                        options = self.ask_save_options(plans)
                    except Exception as exc:
                        messagebox.showerror('บันทึกไม่ได้', str(exc), parent=self)
                    finally:
                        response.put(options)
                elif kind == 'catalogs':
                    self.catalogs = data
                    # Keep selected Setting entries even when missing from current files.
                    self.entries = {i: e for i, e in self.entries.items() if i in self.selected}
                    self.file_list.delete(0, 'end')
                    for c in data:
                        self.file_list.insert('end', c.path.name)
                        for e in c.entries:
                            self.entries[self.ident(e)] = e
                    self.file_label.configure(text=f'{len(data)} ไฟล์พร้อมทำงาน')
                    self.render_entries()
                elif kind == 'setting':
                    for e in data: self.entries[self.ident(e)] = e
                    self.selected = {self.ident(e) for e in data}
                    self.render_entries()
                    self.log(f'โหลด Setting แล้ว {len(self.selected)} รายการ')
                    text_only = sum(not e.sheet for e in data)
                    if text_only:
                        self.log(f'{text_only} รายการใส่เฉพาะข้อความ: โปรแกรมจะหาชีทจาก Contents ของแต่ละไฟล์ให้อัตโนมัติ')
                elif kind == 'complete':
                    folder, success, warns, failures = data
                    self.last_output = folder
                    message = f'บันทึกสำเร็จ {success} ไฟล์\n\n{folder}'
                    (messagebox.showwarning if warns or failures else messagebox.showinfo)('ผลการตัดชีท', message, parent=self)
                elif kind == 'warning': messagebox.showwarning('ผลการตรวจสอบ', data, parent=self)
                elif kind == 'error':
                    self.log('ผิดพลาด: ' + data)
                    messagebox.showerror('ดำเนินการไม่สำเร็จ', data, parent=self)
                elif kind == 'idle':
                    self.busy = False
                    for b in self.actions: b.configure(state='normal')
                    self.update_setting_action()
                    self.status.set('พร้อมทำงาน')
        except queue.Empty:
            pass
        self.after(100, self.pump)

    def choose_source(self):
        self.import_files()

    def show_batch_errors(self, issues):
        dialog = StudioToplevel(self)
        dialog.title('พบ Error — ยกเลิกการตัดทั้งชุด')
        dialog.geometry('800x440')
        dialog.transient(self)
        ctk.CTkLabel(dialog, text=f'พบ Error {len(issues)} รายการ — ยังไม่มีการตัดหรือบันทึกไฟล์ใด',
                     text_color='#BA2839', font=('Tahoma', 16, 'bold')).pack(padx=18, pady=16)
        details = ctk.CTkTextbox(dialog, wrap='word', font=('Tahoma', 13))
        details.pack(fill='both', expand=True, padx=18, pady=(0, 12))
        for issue in issues:
            details.insert('end', f'ไฟล์: {issue.path}\nError: {issue.detail}\n\n')
        details.configure(state='disabled')
        ctk.CTkButton(dialog, text='ปิด', command=dialog.destroy, fg_color=INK, height=36).pack(pady=(0, 16))
        def show():
            center_window(dialog, self)
            dialog.grab_set()
        dialog.after(100, show)

    def scan_folder(self):
        if self.busy: return
        if self.source_folder is None and not self.source_paths:
            self.choose_source()
            return
        folder = self.source_folder
        selected_paths = list(self.source_paths)
        # Do not leave the previous source runnable if the new source cannot be read.
        self.catalogs = []
        self.requested_paths = []
        self.entries = {i: e for i, e in self.entries.items() if i in self.selected}
        self.file_list.delete(0, 'end')
        self.file_label.configure(text='กำลังอ่านต้นทางที่เลือก…')
        self.render_entries()
        def job():
            paths = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in ('.xlsx', '.xlsm') and not p.name.startswith('~$')) if folder else selected_paths
            self.events.put(('sources', paths))
            self.events.put(('log', f'อ่านต้นทาง: {folder or str(len(paths)) + " ไฟล์ที่เลือก"}'))
            catalogs = []
            failed = []
            for i, p in enumerate(paths):
                try: catalogs.append(read_catalog(p))
                except Exception as exc:
                    failed.append(f'{p.name}: {exc}')
                    self.events.put(('log', f'อ่านไม่ได้: {p.name}: {exc}'))
                self.events.put(('progress', (i + 1) / len(paths)))
            self.events.put(('catalogs', catalogs))
            self.events.put(('log', f'อ่าน Contents สำเร็จ {len(catalogs)} / {len(paths)} ไฟล์'))
            if not paths: self.events.put(('warning', 'ไม่พบไฟล์ .xlsx / .xlsm ในโฟลเดอร์ที่เลือก'))
            if failed: self.events.put(('warning', '\n'.join(failed)))
        self.background(job)

    def import_files(self):
        if self.busy: return
        names = ask_native_dialog_centered(filedialog.askopenfilenames, parent=self, title='เลือกไฟล์ Excel ที่ต้องการตัดชีท (เลือกได้หลายไฟล์)', initialdir=self.source_folder or INPUT, filetypes=[('Excel', '*.xlsx *.xlsm')])
        if not names: return
        self.source_folder = None
        self.source_paths = list(dict.fromkeys(Path(name).resolve() for name in names if not Path(name).name.startswith('~$')))
        self.source_label.configure(text=f'ต้นทาง: {len(self.source_paths)} ไฟล์ที่เลือก')
        self.scan_folder()

    def select_folder(self):
        if self.busy: return
        name = ask_directory_centered(parent=self, title='เลือกโฟลเดอร์ต้นทางที่จะตัดชีท', initialdir=self.source_folder or INPUT)
        if not name: return
        self.source_folder = Path(name)
        self.source_paths = []
        self.source_label.configure(text=f'ต้นทาง: {name}')
        self.scan_folder()

    def file_detail(self, _event):
        selection = self.file_list.curselection()
        if selection and selection[0] < len(self.catalogs):
            c = self.catalogs[selection[0]]
            self.file_label.configure(text=f'{c.path.name}\n{len(c.entries)} รายการ Contents')

    def render_entries(self):
        query = key(self.search.get())
        self.tree.delete(*self.tree.get_children())
        self.row_keys = {}
        matchers = [(c, ContentsMatcher(c), {key(s): s for s in c.sheets}) for c in self.catalogs]
        for i, e in self.entries.items():
            if query and query not in key(e.title + ' ' + e.sheet): continue
            iid = str(len(self.row_keys))
            self.row_keys[iid] = i
            count, resolved = 0, set()
            for c, matcher, sheets in matchers:
                matches, _ = matcher.find(e)
                actual = {sheets[key(m.sheet)] for m in matches if key(m.sheet) in sheets and sheets[key(m.sheet)] != c.contents}
                if actual:
                    count += 1
                    resolved.update(actual)
            sheet_label = ', '.join(sorted(resolved)) or e.sheet or 'ค้นจากข้อความ'
            self.tree.insert('', 'end', iid=iid, values=('☑' if i in self.selected else '☐', sheet_label, e.title, count), tags=('missing',) if not count else ())
        self.count.set(f'เลือกตัด {len(self.selected)} / {len(self.entries)} รายการ')
        self.update_setting_action()

    def update_setting_action(self):
        count = len(self.selected)
        self.save_setting_button.configure(state='normal' if count and not self.busy else 'disabled')
        self.setting_hint.configure(text=f'เก็บ {count} รายการไว้ใช้ครั้งถัดไป' if count else 'เลือกชีทในรายการก่อนบันทึก')

    def toggle(self, event):
        if self.busy: return 'break'
        iid = self.tree.focus() if event.keysym == 'space' else self.tree.identify_row(event.y)
        if iid in self.row_keys:
            i = self.row_keys[iid]
            if i in self.selected: self.selected.remove(i)
            else: self.selected.add(i)
            values = list(self.tree.item(iid, 'values'))
            values[0] = '☑' if i in self.selected else '☐'
            self.tree.item(iid, values=values)
            self.count.set(f'เลือกตัด {len(self.selected)} / {len(self.entries)} รายการ')
            self.update_setting_action()
        return 'break' if event.keysym == 'space' else None

    def show_entry(self, event):
        iid = self.tree.identify_row(event.y)
        if iid in self.row_keys:
            e = self.entries[self.row_keys[iid]]
            messagebox.showinfo(e.sheet or 'Contents', e.title, parent=self)

    def select_visible(self, checked):
        if self.busy: return
        for i in self.row_keys.values():
            if checked: self.selected.add(i)
            else: self.selected.discard(i)
        self.render_entries()

    def chosen(self):
        return [e for i, e in self.entries.items() if i in self.selected]

    def save_setting(self):
        selected = self.chosen()
        if not selected:
            messagebox.showwarning('ยังไม่ได้เลือก', 'เลือกอย่างน้อย 1 รายการก่อนบันทึก Setting', parent=self)
            return
        name = ask_native_dialog_centered(filedialog.asksaveasfilename, parent=self, title='บันทึก Setting เป็น Excel', initialdir=BASE, initialfile='Sheet_Removal_Setting.xlsx', defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if name:
            if Path(name).resolve() in {c.path.resolve() for c in self.catalogs}:
                messagebox.showerror('บันทึกไม่ได้', 'กรุณาเลือกชื่อ Setting ที่ไม่ใช่ไฟล์ Table ต้นฉบับ', parent=self)
                return
            def job():
                save_settings(Path(name), selected)
                self.events.put(('log', f'บันทึก Setting {len(selected)} รายการ: {name}'))
            self.background(job)

    def load_setting(self):
        if self.busy: return
        name = ask_native_dialog_centered(filedialog.askopenfilename, parent=self, title='โหลด Setting จาก Excel', initialdir=BASE, filetypes=[('Excel', '*.xlsx')])
        if name:
            def job():
                entries = read_catalog(Path(name), setting=True).entries
                self.events.put(('setting', entries))
                self.events.put(('setting_name', name))
            self.background(job)

    def ready(self):
        if not self.requested_paths or not self.selected:
            messagebox.showwarning('ยังไม่พร้อม', 'เลือกไฟล์ Excel ต้นทาง และเลือกอย่างน้อย 1 รายการ', parent=self)
            return False
        return True

    def ask_suffix(self, first_source):
        dialog = StudioToplevel(self)
        dialog.title('ข้อความต่อท้ายชื่อไฟล์ใหม่')
        dialog.geometry('700x300')
        dialog.transient(self)
        result = None
        suffix = tk.StringVar()
        preview = tk.StringVar(value=first_source.name)
        ctk.CTkLabel(dialog, text='ใส่ข้อความต่อท้ายทุกไฟล์ เช่น MKT', font=('Tahoma', 17, 'bold'), text_color=INK).pack(pady=(20, 12))
        entry = ctk.CTkEntry(dialog, textvariable=suffix, height=38, font=('Tahoma', 14))
        entry.pack(fill='x', padx=24)
        ctk.CTkLabel(dialog, text='ตัวอย่างชื่อไฟล์แรก', text_color=TEAL, font=('Tahoma', 12, 'bold')).pack(anchor='w', padx=24, pady=(14, 2))
        ctk.CTkLabel(dialog, textvariable=preview, wraplength=640, justify='left', font=('Tahoma', 13)).pack(fill='x', padx=24)
        def update_preview(*_):
            text = suffix.get().strip()
            preview.set(first_source.stem + (' ' + text if text else '') + first_source.suffix)
        suffix.trace_add('write', update_preview)
        def accept():
            nonlocal result
            text = suffix.get().strip()
            if not text:
                messagebox.showwarning('กรุณากรอกข้อความ', 'ใส่ข้อความต่อท้าย หรือกดยกเลิก', parent=dialog)
                return
            if any(c in '<>:"/\\|?*' or ord(c) < 32 for c in text) or text.endswith('.'):
                messagebox.showerror('ชื่อไฟล์ไม่ถูกต้อง', 'ห้ามใช้ < > : " / \\ | ? * หรือจุดท้ายข้อความ', parent=dialog)
                return
            result = text
            dialog.destroy()
        buttons = ctk.CTkFrame(dialog, fg_color='transparent')
        buttons.pack(side='bottom', pady=18)
        ctk.CTkButton(buttons, text='ตกลง', command=accept, fg_color=TEAL).pack(side='left', padx=6)
        ctk.CTkButton(buttons, text='ยกเลิก', command=dialog.destroy, fg_color=INK).pack(side='left', padx=6)
        dialog.bind('<Return>', lambda _: accept())
        dialog.bind('<Escape>', lambda _: dialog.destroy())
        def show():
            center_window(dialog, self)
            dialog.grab_set()
            entry.focus_set()
        dialog.after(100, show)
        self.wait_window(dialog)
        return result

    def ask_save_options(self, plans):
        self.status.set('ตรวจสอบผ่านแล้ว — เลือกที่บันทึก')
        name = ask_directory_centered(parent=self, title='เลือกโฟลเดอร์บันทึกไฟล์ใหม่ (ไม่เขียนทับต้นฉบับ)',
                                       initialdir=self.last_output if self.last_output and self.last_output.exists() else plans[0].catalog.path.parent,
                                       mustexist=True)
        if not name:
            return None
        append = messagebox.askyesnocancel('ชื่อไฟล์ใหม่', 'ต้องการใส่ข้อความต่อท้ายชื่อไฟล์ทุกไฟล์หรือไม่?\n\nYes: กรอกข้อความและดูตัวอย่าง\nNo: ใช้ชื่อเดิม\nCancel: ยกเลิกการตัด', parent=self)
        if append is None:
            return None
        suffix = self.ask_suffix(plans[0].catalog.path) if append else ''
        if suffix is None:
            return None
        destinations = output_paths(plans, Path(name), suffix)
        self.log(f'บันทึกที่: {name}\nตัวอย่าง: {destinations[0].name}')
        self.status.set('กำลังตัดและบันทึกไฟล์ใหม่…')
        return destinations

    def run_batch(self):
        if self.busy or not self.ready(): return
        paths, selected = list(self.requested_paths), self.chosen()
        def job():
            self.events.put(('status', 'กำลังตรวจสอบทุกไฟล์ก่อนตัด…'))
            self.events.put(('log', f'ตรวจสอบอัตโนมัติ {len(paths)} ไฟล์ / {len(selected)} รายการต่อไฟล์'))
            plans, issues = preflight_batch(paths, selected)
            for plan in plans:
                self.events.put(('log', f'{plan.catalog.path.name}: จับคู่ได้ {plan.matched_items}/{len(selected)} รายการ / ตัดได้ {len(plan.removed)} ชีท'))
                for info in plan.infos: self.events.put(('log', '  ' + info))
                if not plan.removed:
                    self.events.put(('log', f'ข้ามไฟล์ {plan.catalog.path.name}: ไม่มีชีทที่ตรงกับ Setting จึงไม่สร้างไฟล์ใหม่'))
            if issues:
                for issue in issues:
                    self.events.put(('log', f'Error ไฟล์ {issue.path}: {issue.detail}'))
                self.events.put(('log', 'ผิดพลาด: ยกเลิกทั้งชุด ไม่มีการตัดหรือบันทึกไฟล์ใด'))
                self.events.put(('batch_errors', issues))
                return
            plans = [plan for plan in plans if plan.removed]
            if not plans:
                self.events.put(('log', 'ไม่มีชีทที่ตรงกับ Setting ในทุกไฟล์ ไม่ต้องตัดและไม่สร้างไฟล์ใหม่'))
                self.events.put(('warning', 'ไม่พบชีทที่ตรงกับ Setting ในไฟล์ที่เลือก จึงข้ามทั้งหมด ไม่มีการสร้างไฟล์ใหม่'))
                return
            self.events.put(('log', 'ตรวจสอบสำเร็จทุกไฟล์ ไม่มี Error กรุณาเลือกที่บันทึกและชื่อไฟล์'))
            response = queue.Queue(maxsize=1)
            self.events.put(('save_options', (plans, response)))
            destinations = response.get()
            if destinations is None:
                self.events.put(('log', 'ยกเลิกการบันทึก ไม่มีการตัดหรือสร้างไฟล์ใหม่'))
                return
            folder = destinations[0].parent
            def notify(j, plan):
                self.events.put(('log', f'เตรียมผลลัพธ์: {plan.catalog.path.name} → {destinations[j].name}'))
                self.events.put(('progress', (j + 1) / len(plans)))
            save_batch(plans, destinations, notify)
            self.events.put(('log', f'บันทึกสำเร็จ {len(plans)} ไฟล์ → {folder}'))
            self.events.put(('complete', (folder, len(plans), 0, 0)))
        self.background(job)

    def open_output(self):
        if self.last_output and self.last_output.is_dir():
            os.startfile(self.last_output)
        else:
            messagebox.showinfo('เปิดผลลัพธ์', 'ยังไม่มีผลลัพธ์ที่บันทึกสำเร็จ', parent=self)

    def close(self):
        if self.busy:
            messagebox.showinfo('กำลังทำงาน', 'กรุณารอให้งานเสร็จก่อนปิดโปรแกรม', parent=self)
        else: self.destroy()



# ===== Embedded Windows icon (Base64 data only; no external assets needed) =====
import atexit
import base64

_ICON_BASE64 = (
    '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'
    '1cDWx3WENgGyNovPbL7UxpUzXuYcTaKRVlIrEtOfd+YSRiO/nCaoAMW/pSJAAkmcgk3ooHUDybhKU7K6ZMdSw6e2fpBqMeh1wQTI'
    'aI/aAnUdBRmKppydIntFBhLwvYceY0LYyaYdZfdOj2FC13HbLTfh1t96D0KKQu4/SImNCKiHAjXS8lrkicsnXk4riVCvTuOxROOy'
    'M1ZjZ22EKSZEF/4h8R+WStixay82P/AYVE+PsfYlbS2P/fvrKcf/JV5jI/jWRnkRlUFkAogGyMPB7V0uO0z6Nui3zTAw1zO5itIs'
    'lwdPBRqt0TGM9vYiNtulCUhk48rU64bdLB/gt0qOCUskpY0oa1NTqlWpMD7Qzsi2FIETJCTKQB21CoSupQLgBAG1wEXabAG9PVh3'
    '30O45/yz8cYLz8NUvS6YqAowHTfVdHwQr7vkQpx6/DJ8/PNfws7nnucW8tQGX5JJqIW77FFMDBAlERC18PCBQSw8eSVQAcammrz9'
    'MRHeqiCWdr0F9eT3HkFr/zDChQtpPwC/pVp7GHg2A38WKXvMDMDwhsmI4b9ss0Pjdjkfw3fJ3MmGNz1PxKlwQgCZOyP8t3PPxvDU'
    'JJKQNgcwCRm2OoY2RTHp1ILOiFqhat4ozXDOkkU4Y+F8ZFmMWy+9BH/b040iXYdFoASc7bVsoITBUzPGOI6xcuFCnH/ccWi1Gvjw'
    'JRdgQamAZquFsaefwUPf/BbOX70KC8MQNQPxkiSi/zdNDWLZkrkctfvU3/49Rjc8h4+84mLUI9LqspULPTWpJbL8D4yN4CeRxuuW'
    'H4fN08MszyzxbUCpWAgxODKOZ+5+ACiJ8eeSWDyh66J9Lujl5yC2p5Tr4MVUBtmMXt/gc2FgMx6vMYHBy3ILe6YUcKOm1UQTc96y'
    'Jfint/4/s939iHg5iatI4hauWXUarll1Do7sMM2aQKI1QaM+xcDP61edhutWnoqhsSHsvexc/J9HHsN9//4g3v7G1yrUqYJY8UaO'
    'JIsKKlB7qxOYU67gY+9+J5557lkUxkcwPTGOYncvtKYuuwpUzpY167h3zyBOP+c8jGZ1NKmekKSatZVoSGmqwt4uPPmv38XUxh0I'
    'BgbMlBvidghyhwIyS+evvTO8T48eCpZooKadiWyFj+/12VJlu7qdjO3IFbTeg/eeyyLLFGXSjNUa+NauTaglZg9d1/lC8hEtvCpw'
    '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'
    'yJfyYm35ywTwrBTIlbj1qaXrpyyhnDoipc3erFKXIU0mDKZs5ILtBxuIlS2dx2TrTb/xosNh8sJHT3J49gnbOWa0eW5i3oHOZrU6'
    'iWAydrlEXoZt90TkfY1jSvYnn2cjtP62zrJvYNum9bPO838gBrAHx668ARZwylnnBQV1sVb6fCA4CQrLoEBVk2UEvG80mUemGZhF'
    'VdqaDxkfiN1AbyVRkx2+h+/6WDewHeHJoWW3Hw2IaNbqEv9dChBd/Nhd2hni+dtidChqzuWKGducfLp2HtKj+3G7ImsU2B2uFZX3'
    'NAGMQ+tB6GyrBp5CkvwIO59/1lvldieAF73qf9EMYI8Aa9YEhwhFhlixogdZVmHgPKXmfjb+SQflhWfUQSkDWpQqY0bbsol+iqes'
    'VLT902ULEP4qE1UhCE2P+sycQy1kyqRqAvrF+9Br155Cvu++S/c1R3sTCsz6vqsvtOPnnj/BrNdXdG3zfoEiTt0tqHoT27dPz7qq'
    'BYG1e9z83I9fJAP491DMDHT8Ah/mP8ERdMzTUYE6/1EZ4IXu+//VGP6jHD6Bf6HEful46XjpeOl46XjpeOl46XjpeOkAHf8vL7V8'
    'HS5RJBcAAAAASUVORK5CYIKJUE5HDQoaCgAAAA1JSERSAAABAAAAAQAIBgAAAFxyqGYAAQAASURBVHic7P0HvGXXVR+Of8+59742'
    'b7pm1IulUZct25ItW+4FF9zBkjsugE0vP0gCJOAYSAL/JEB6iENISEwCxoABuclNlqxiVav3NhpNnzev33rO/7PqXvvcJwiyPTJE'
    'V3rz7rv3lH323mut7+rA06+nX0+/nn49/Xr69fTr6dfTr6dfT7+efj39evr19Ovp19Ovp19Pv55+Pf16+vX06+nX06+nX0+/nn49'
    '/Xr69fTr6dfTr6dfT7+efj39evr19Ovp19Ovp19Pv55+Pf16+vX06+/Sq3iqB/Bd8np6Hv7ffNX4f/z1933jF/joRwvceac85759'
    '8nv79hrnnFPjY/+0Bor/5zfB/+OvwvdI3B/0+uQ5NfAxev/3do/8/WIAH/1o6Qv58pdX+NjHqv+70z5afuLQoc5yb0urWlkqqt5q'
    'MVqcL0eT7RLDQTEfDx5MFxj107y1uro5NgBY0H/tb3stAJ2OHNdtlRgO5fzJYcWf2zVbE7Wd7a/RIKzRev29GA5Yn8aQHRtf64Hh'
    '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'
    '7vOQkeIb9wAAAABJRU5ErkJggg=='
)
_icon_directory = tempfile.TemporaryDirectory(prefix='cut_mkt_icon_')
atexit.register(_icon_directory.cleanup)
APP_ICON = Path(_icon_directory.name) / 'cut_mkt.ico'
APP_ICON.write_bytes(base64.b64decode(_ICON_BASE64, validate=True))





 
# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None): # ชื่อฟังก์ชันนี้จะถูกใช้ใน Launcher
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน QuotaSamplerApp.
    """
    print(f"--- QUOTA_SAMPLER_INFO: Starting 'QuotaSamplerApp' via run_this_app() ---")
    try:
    # --- ส่วนที่ใช้รันโปรแกรม --
    #if __name__ == '__main__':
        if sys.platform == 'win32':
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('CutMKT.ExcelTrimmer')
        ctk.set_appearance_mode('light')
        ctk.set_default_color_theme('green')
        App().mainloop()

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

