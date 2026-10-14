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
import io
import math


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
from PIL import Image, ImageTk
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
    mkt_entries: list[Entry] = field(default_factory=list)


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
        mkt_entries = []
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
            entry = Entry(title, sheet, r)
            entries.append(entry)
            if 'mkt issue' in key(value('D')):
                mkt_entries.append(entry)
        if not entries:
            raise ValueError('ไม่พบรายการในคอลัมน์ C')
        return Catalog(Path(path), contents, part, entries, sheets, mkt_entries)


def read_mkt_settings(path: Path) -> list[Entry]:
    """Read only Contents rows marked MKT Issue and verify their linked sheets."""
    catalog = read_catalog(path)
    if not catalog.mkt_entries:
        raise ValueError('ไม่พบ MKT Issue ใน Contents คอลัมน์ D')
    sheets = {key(name): name for name in catalog.sheets}
    entries = []
    for entry in catalog.mkt_entries:
        actual = sheets.get(key(entry.sheet))
        if not actual or actual == catalog.contents:
            raise ValueError(f'Contents แถว {entry.row}: ไม่พบชีทปลายทางของ MKT Issue ({entry.sheet or entry.title})')
        entries.append(Entry(entry.title, actual, entry.row))
    return entries


def read_setting_entries(path: Path) -> list[Entry]:
    catalog = read_catalog(path, setting=True)
    if catalog.mkt_entries:
        return read_mkt_settings(path)
    return catalog.entries


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
DATA = BASE / 'data' if (BASE / 'data').is_dir() else BASE.parent / 'data'
DEFAULT_MKT_SOURCE = DATA / 'J69229 Penetration survey for Health Care (ADL1-5) Y2026 Wave3 (HEADER G GBKK & NE) 8 Oct N_Nan.xlsx'
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
        self.geometry('1160x780')
        self.minsize(1000, 720)
        self.configure(fg_color=BG)
        self.catalogs = []
        self.source_folder = None
        self.source_paths = []
        self.requested_paths = []
        self.entries: dict[tuple[str, str], Entry] = {}
        self.selected: set[tuple[str, str]] = set()
        self.mkt_entries: list[Entry] = []
        self.mkt_source = None
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
        ctk.CTkLabel(brand_text, text='ลิสต์ MKT Issue คอลัมน์ D  •  บันทึก Setting  •  ทำงานพร้อมกันหลายไฟล์', text_color='#C4D8DC', font=('Tahoma', 13)).pack(anchor='w')
        mkt_tools = ctk.CTkFrame(self, fg_color='white', corner_radius=0)
        mkt_tools.pack(fill='x')
        self.mkt_help_button = ctk.CTkButton(mkt_tools, text='?', command=self.show_mkt_example,
            width=40, height=40, corner_radius=20, font=('Segoe UI', 22, 'bold'),
            fg_color='#F59E0B', text_color='#422006', hover_color='#FBBF24',
            border_width=2, border_color='#B45309')
        self.mkt_help_button.pack(side='left', padx=(22, 10), pady=8)
        self.mkt_help_hovered = False
        self.mkt_help_phase = 0
        self.mkt_help_tooltip = None
        self.mkt_help_button.bind('<Enter>', self.show_mkt_help_tooltip, add='+')
        self.mkt_help_button.bind('<Leave>', self.hide_mkt_help_tooltip, add='+')
        self.after(1000, self.animate_mkt_help)
        self.button(mkt_tools, 'ลิสต์ MKT Issue จาก Contents', self.list_mkt_issue, True).pack(side='left', padx=(0, 8), pady=8)
        self.save_mkt_button = self.button(mkt_tools, 'Save ข้อ MKT Issue เป็น Excel', self.save_mkt_issue)
        self.save_mkt_button.pack(side='left', padx=(0, 12), pady=8)
        self.mkt_label = ctk.CTkLabel(mkt_tools, text='ยังไม่ได้โหลด MKT Issue', text_color=TEAL, font=('Tahoma', 12))
        self.mkt_label.pack(side='left')
        self.save_mkt_button.configure(state='disabled')
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
                elif kind == 'mkt':
                    source, entries = data
                    self.mkt_source = source
                    self.mkt_entries = entries
                    for e in entries:
                        self.entries[self.ident(e)] = e
                    self.render_entries()
                    self.mkt_label.configure(text=f'MKT Issue {len(entries)} รายการ')
                    self.log(f'ลิสต์ MKT Issue จาก Contents คอลัมน์ D: {source.name} / {len(entries)} รายการ พร้อมชื่อชีท')
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
        self.save_mkt_button.configure(state='normal' if self.mkt_entries and not self.busy else 'disabled')
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

    def animate_mkt_help(self):
        """A slow colour pulse keeps the help icon visible and clickable."""
        if not self.mkt_help_button.winfo_exists():
            return
        if not self.mkt_help_hovered:
            amount = (1 - math.cos(self.mkt_help_phase * math.pi / 15)) / 2
            start, end = (245, 158, 11), (253, 210, 68)
            color = '#' + ''.join(f'{round(a + (b - a) * amount):02x}' for a, b in zip(start, end))
            self.mkt_help_button.configure(fg_color=color)
            self.mkt_help_phase = (self.mkt_help_phase + 1) % 30
        self.after(80, self.animate_mkt_help)

    def show_mkt_help_tooltip(self, _event=None):
        self.mkt_help_hovered = True
        self.mkt_help_button.configure(fg_color='#F59E0B')
        if self.mkt_help_tooltip is not None:
            return
        tip = tk.Toplevel(self)
        self.mkt_help_tooltip = tip
        tip.withdraw()
        tip.overrideredirect(True)
        tk.Label(tip, text='ดูตัวอย่างไฟล์ MKT พี่แนน', bg=INK, fg='white',
            font=('Tahoma', 11), padx=10, pady=6).pack()
        tip.geometry(f'+{self.mkt_help_button.winfo_rootx()}+{self.mkt_help_button.winfo_rooty() + self.mkt_help_button.winfo_height() + 5}')
        tip.deiconify()

    def hide_mkt_help_tooltip(self, _event=None):
        self.mkt_help_hovered = False
        if self.mkt_help_tooltip is not None:
            self.mkt_help_tooltip.destroy()
            self.mkt_help_tooltip = None

    def show_mkt_example(self):
        """Show the bundled reference without importing any Table or Setting."""
        self.hide_mkt_help_tooltip()
        existing = getattr(self, 'mkt_example_window', None)
        if existing is not None and existing.winfo_exists():
            existing.lift()
            existing.focus_set()
            return
        try:
            if MKT_EXAMPLE_IMAGE_BASE64:
                source = io.BytesIO(base64.b64decode(MKT_EXAMPLE_IMAGE_BASE64, validate=True))
            else:
                source = DATA / 'MKT_Nan_Example.png'
            with Image.open(source) as image:
                original = image.convert('RGB')
        except (OSError, ValueError) as exc:
            messagebox.showerror('ภาพตัวอย่าง MKT พี่แนน',
                f'ยังไม่มีภาพตัวอย่างแนบในโปรแกรม\nกรุณาวางภาพที่ {DATA / "MKT_Nan_Example.png"}\n\n{exc}', parent=self)
            return
        dialog = StudioToplevel(self)
        self.mkt_example_window = dialog
        dialog.title('ตัวอย่างไฟล์ MKT พี่แนน — Contents คอลัมน์ D')
        width = min(1200, max(600, self.winfo_screenwidth() - 100))
        height = min(820, max(400, self.winfo_screenheight() - 120))
        dialog.geometry(f'{width}x{height}')
        dialog.transient(self)
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(1, weight=1)
        toolbar = ctk.CTkFrame(dialog, fg_color='transparent')
        toolbar.grid(row=0, column=0, sticky='ew', padx=12, pady=10)
        ctk.CTkLabel(toolbar, text='ตัวอย่าง: Contents คอลัมน์ D ระบุ MKT Issue',
            font=('Tahoma', 14, 'bold')).pack(side='left')
        frame = ttk.Frame(dialog)
        frame.grid(row=1, column=0, sticky='nsew', padx=12, pady=(0, 12))
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        canvas = tk.Canvas(frame, background='white', highlightthickness=0)
        canvas.grid(row=0, column=0, sticky='nsew')
        vertical = ttk.Scrollbar(frame, orient='vertical', command=canvas.yview)
        vertical.grid(row=0, column=1, sticky='ns')
        horizontal = ttk.Scrollbar(frame, orient='horizontal', command=canvas.xview)
        horizontal.grid(row=1, column=0, sticky='ew')
        canvas.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        def display(fit):
            scale = min(1.0, (width - 48) / original.width, (height - 100) / original.height) if fit else 1.0
            size = (max(1, round(original.width * scale)), max(1, round(original.height * scale)))
            dialog.preview_image = ImageTk.PhotoImage(original.resize(size, Image.Resampling.LANCZOS), master=dialog)
            canvas.delete('all')
            canvas.create_image(0, 0, anchor='nw', image=dialog.preview_image)
            canvas.configure(scrollregion=(0, 0, *size))
            canvas.xview_moveto(0)
            canvas.yview_moveto(0)
        ctk.CTkButton(toolbar, text='ขนาดจริง 100%', width=125, command=lambda: display(False)).pack(side='right', padx=(8, 0))
        ctk.CTkButton(toolbar, text='พอดีหน้าต่าง', width=120, command=lambda: display(True)).pack(side='right')
        display(True)
        dialog.bind('<Escape>', lambda _: dialog.destroy())
        dialog.after(100, lambda: center_window(dialog, self) if dialog.winfo_exists() else None)

    def load_mkt_source(self, path):
        def job():
            entries = read_mkt_settings(Path(path))
            self.events.put(('mkt', (Path(path), entries)))
        self.background(job)

    def list_mkt_issue(self):
        if self.busy: return
        name = ask_native_dialog_centered(filedialog.askopenfilename, parent=self,
            title='ลิสต์ MKT Issue จาก Contents คอลัมน์ D',
            initialdir=DATA, initialfile=DEFAULT_MKT_SOURCE.name,
            filetypes=[('Excel', '*.xlsx *.xlsm')])
        if name:
            self.load_mkt_source(Path(name))

    def save_mkt_issue(self):
        if self.busy or not self.mkt_entries: return
        name = ask_native_dialog_centered(filedialog.asksaveasfilename, parent=self,
            title='Save ข้อ MKT Issue เป็น Excel พร้อมชื่อชีท', initialdir=DATA,
            initialfile='MKT_Issue_Setting.xlsx', defaultextension='.xlsx',
            filetypes=[('Excel', '*.xlsx')])
        if not name: return
        protected = {c.path.resolve() for c in self.catalogs}
        if self.mkt_source:
            protected.add(self.mkt_source.resolve())
        if Path(name).resolve() in protected:
            messagebox.showerror('บันทึกไม่ได้', 'กรุณาเลือกชื่อไฟล์ใหม่ที่ไม่ใช่ต้นฉบับ', parent=self)
            return
        entries = list(self.mkt_entries)
        def job():
            save_settings(Path(name), entries)
            self.events.put(('log', f'บันทึกแล้ว MKT Issue {len(entries)} รายการ พร้อมชื่อชีท: {name}'))
        self.background(job)

    def save_setting(self):
        selected = self.chosen()
        if not selected:
            messagebox.showwarning('ยังไม่ได้เลือก', 'เลือกอย่างน้อย 1 รายการก่อนบันทึก Setting', parent=self)
            return
        name = ask_native_dialog_centered(filedialog.asksaveasfilename, parent=self, title='บันทึก Setting เป็น Excel', initialdir=BASE, initialfile='Sheet_Removal_Setting.xlsx', defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if name:
            if Path(name).resolve() in {c.path.resolve() for c in self.catalogs} | ({self.mkt_source.resolve()} if self.mkt_source else set()):
                messagebox.showerror('บันทึกไม่ได้', 'กรุณาเลือกชื่อ Setting ที่ไม่ใช่ไฟล์ Table ต้นฉบับ', parent=self)
                return
            def job():
                save_settings(Path(name), selected)
                self.events.put(('log', f'บันทึก Setting {len(selected)} รายการ: {name}'))
            self.background(job)

    def load_setting(self):
        if self.busy: return
        name = ask_native_dialog_centered(filedialog.askopenfilename, parent=self, title='โหลด Setting จาก Excel', initialdir=DATA, filetypes=[('Excel', '*.xlsx *.xlsm')])
        if name:
            def job():
                entries = read_setting_entries(Path(name))
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

# Embedded MKT reference image; no external image file is required.
MKT_EXAMPLE_IMAGE_BASE64 = (
    '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'
    'YdCTiIiIiIiIiIiINhUGPYmIiIiIiIiIiGhTYdCTiIiIiIiIiIiINhUGPYmIiIiIiIiIiGhTYdCTiIiIiIiIiIiINhUGPYmIiIiI'
    'iIiIiGhTYdCTiIiIiIiIiIiINhUGPYmIiIiIiIiIiGhTYdCTiIiIiIiIiIiINhUGPYmIiIiIiIiIiGhTYdCTiIiIiIiIiIiINhUG'
    'PYmIiIiIiIiIiGhTYdCTiIiIiIiIiIiINhUGPYmIiIiIiIiIiGhT2Xb16lW89dZbuHLlipxGREREREREREREtOFs27t3Lz784Q9j'
    '3759choRERERERERERHRhsPh7URERERERERERLSpMOhJREREREREREREmwqDnkRERERERERERLSpMOhJREREREREREREmwqDnkRE'
    'RERERERERLSpMOhJREREREREREREmwqDnkRERERERERERLSpMOhJREREREREREREmwqDnkRERERERERERLSpMOhJRERERERERERE'
    'mwqDnkRERERERERERLSpMOhJREREREREREREmwqDnkRERERERERERLSpMOhJREREREREREREmwqDnkRERERERERERLSpMOhJRERE'
    'REREREREm4px/vx5x3EcOI6DVquFT3/60/IyW97zzz+PX/zFX5RnExERERERERHRiHMcR55FW4Bx8+bN4Mi/9NJLOHLku+ElevC1'
    '3/+wPCu1f3bkSThOVZ6d2iDKcut71zA9/Z48m4iIiIiIiIiIRphhFBn03KKGFvT8Jz//j+TZif3u7/8h/unP/yMANTkptUGUhUFP'
    'IiIiIiIiIqKNh0HPrWt4z/R0nN4/AAzDkHPsnZx/mg8RERERERERERFtKEMLesqxwzQfd/3BBRzl/NN8iIiIiIiIiIiIaGMZYtDT'
    'fTlSLx8MuKennH+aDxEREREREREREW0sQwx69v5x1x9cwFHOP82HiIiIiIiIiIiINpbhBT37+A+D7unZx39ERERERERERES0sQzt'
    '7e2f+uQToXlfOvGbwf+/+847uHPnDt5efRurb7+NEye+FFr2xT84M9C3t/dbFr69nYiIiIiIiIho4+Hb27eu4fX0lIaJHz/2LN59'
    '91386Ic/xH+69Z/wt3/7A/ztD36AL33pS13LYtA9PfssCxEREREREREREW0cwwt6vud0fX7l+HH86Ec/xA++/3384Ps/QPl/KXct'
    '47znRhoHGYWX809bFiIiIiIiIiIiIto4hhf0jHgpkOMA//J/msPf/uAH+Ff/6n/tShN7Vw6zp2fashAREREREREREdHGMbSg53uO'
    'o/w8P/+/dc0TPxhwT085/7RlISIiIiIiIiIioo1jaEFP+S3oaf7DoHt69vEfERERERERERERbSzDC3pGDBVP+nHXH1zAUc4/zYeI'
    'iIiIiIiIiIg2liEGPZ2ePxh0T8+IbST9EBERERERERER0cYyxKBn7x93/cEFHOX803yIiIiIiIiIiIhoYzFu3rwZhPZeeuklHDny'
    '3fASPfja739YnpXaPzvyJBynKs9ObRBlufW9a5iefk+eTUREREREREREI8wwigPtWEcbh/H888/zyBMREREREUXY9t572Luygr/6'
    'z/9zvLN9u5xMREREIyqip+eR8BLrxDCMkYnEz8/PY3p6Wp5NRERERESb3Nt/9mf47j/9pzDuvRf3/uzP4n2FAt73cz8HY+dOeVEi'
    'IhpBoxRforU1tGd69osNkoiIiIiI1tvbL70EAHDefhs/OnMGN48exfVPfAI3jx7F7a99De9cvCivQkRERCNggEHPG/jM9MvYPv0y'
    '7pm/ISemNpC3t1+7hMe/eAm2PB8xaUREREREtOU5P/gBVs+elWfj3atX8cMXX8T3jx/Hdz/9aQZAiYiIRpA66HnuPLb7Qczp8zgj'
    'p0vOzLew8qmHsDp/GHemd8nJqYV6el67hMe9ssifJGXbOBooGjlUNmwkNk350yy7gdgV5AwDhmEg17Vz67XP67XdQdno5R9FDRS9'
    'dmoYBooNOb0PjWKQr2EUMcisiYiI1poD4H3/9X+Nuw8elJMCYgD0xn/5X+J709P44Ysvwrl1S16UiIiI1pA66PnA/VidO4iHsRvf'
    'mL8fT8jpIbdx4cpOTH1sh5zQs1BPzz0H8Mr8YazOHw7K9Afe9zuxZRuS0I29JmjQKMKQAzZCYKz3oEA4aNFfXmthWIGrYeXbCxuV'
    'p0oYrztwHAfLMxl5AUpklI6pSp9l1F4/5HNb2I60XvBRXoB0LNQdt61W8xC2230daRSlcgBue88ZMHKVcI/5fBWO48Bx6rDE+URE'
    'RBvQtg9+EO8/fhy7/viPcd///r/j/c88g+0f+5i8WCBqCPwPf/d32QOUiIhoHaiDnjJvOPjJeb+H5Rs4ec1NOjP/Oo5dvoVjs8J8'
    'oXemuGxSaZ7padffCPX+fLx+O5Rejyhz2G2c/GLcMoJGEUYBQcDAaZexUpADF15ws1CTZjZQzHYCY3WrhoIcNEjMRLntlcFx0C6v'
    'oBARsKC10karaWIsK8+nzSePqrOMnuLacdcP+wJWrLpwXgOlp7xrRBBQ9D9tlE3AHGijq2FRvIjYFczJlzEAsJewABNmcwFLvV3A'
    'iIiINgzj3ntx7yc+gQ/86q/ivtOnEwVA/R6gN48dC4bAMwBKRES0dpIHPQHg8kW0HvR6WD6zA89W3WdiPjH9EJ7bvxPPzR3GnfkH'
    'cXTPbZysXsT4M93LJmWkeKZnpvCg2wvU6wmKF893ApdCmc9/CpHlODP/OhY+7g3Njy2rjcpcDVa9irw/KzOD5bqF2pwYvGygmF3A'
    'VFvq7dRYRA0WJr2V88fL3UGDpU5vru4h0mqZiSmYoTm6IaxSbzI/MdQLVezZ5fVsa/jpfnBVzKeAqNiI2yOsgBqaKGWlnmTKfdWV'
    '3Sfnm0MuJ+XTKHo90bzyV/rZnieyjuSyaILPkfscVb9yjz+xXGn2R3VcfFFtQe7JKH5XlVWe1513fPl19ej2LIw+vjJd+b31uvbZ'
    'Xy6qvLKI+ujalygJrh+ZGSy7XS/d5EPjQLOFdjBH0DiBUtPCrDL6mnR/fCbK5fC1zF5aAMrlrl6b9tICMHUKp8rAAqOeRES0hTAA'
    'SkREtDGkC3ruP4hjD3j/v3cnHpGSA9duYAHCsg/sxtOXr6Me14NSkKanZ8ieXZjaL3wXypz52G48cvmWFCC5gcWzuzFb8Ibmx5XV'
    'XsJCVG++7BjMIDDhBm5Q76UnWBOl1qTbi6tuoVl6KvEQ2saJEprWpBdMccuwUm4Lvcn8AJIXWAp6k7VRxgU3MCj0QnXq4yhlxaBT'
    'E6U54JTjwHGqyMvbUA5nzWBmuQ4r6JnqB3zkfT0RKl902UVyvss4NWWiubAUHOPGYg3m1ATcwyBvz6/bpNuDG0SKrCO5LEJQK0RV'
    'Bi8tVL95VMVefV3lEvJqlwFV/SmPCzRtIY5c1qh5cfUaVX5dPWYwoT2+CdkV5AorQi9pv37iyqsTtS8REl0/wuwLK4A5BnkVP4Bq'
    'lo8r2lqP+zMxCSv4Q0wDJ0rA1MQhaSEbSwvA1EQGmYkpQDgmREREW0lXAPSrX8X7n3kG9yQMgP6NEAB998035UWJiIioD9u+fkWe'
    'NQBXb+GboRk7MLb/FlpXQzO1jBQ9PQFxiPvrePaynOrZswPj8rxrt7GC6/hkMDy+ha/Iy3QZxyE5ypI5hHGs4IINNIpZlMbr3nPy'
    'JNkxmPLw0RAT5eN+N9BJWGiiFRUJAdxAS7bTi2tx0oHjb9RewoLYAywzgSnT264XeAm2gwxmqjPISL1Q3e2LZTVRPjXTCTDJ20hN'
    '3le3/rryFcseIzMzGwrYLNZMTE345ZO359Vtmu3F1lEcRRn8NLF+gXBPvWwJTb+O/OX9vIT217U/Oqq2IC3WLaqsMe2jq14V5dfQ'
    'H980xHrPo1rNJyivTpp90V8/QuwKnio1Yc3KdZ2gl2fP+5PH8TJQOtFw27s12/3Hm8YJlMa9+ZkJTKGEE7H5EhERbW7Gvffi3p/5'
    'GXzgV38VH0oYAH1PCIDeOHLEfQs8A6BEREQDsW35tSFEPffulHpU3kbr8k6M7Q0tpeWk6el57jzuf3GH93Kjh/CbYk9P0bXbWJHn'
    '7dmBceHFSO7LkR7E0T3ygqKo4MQFrGAchzI2LqwAqBW8YGRnqG6xgc5Q1oIYyFLJYiw8Xl3i94Zze6qtiIVqt9BEDYVgaGsWJX9D'
    '7RaaUYEXQOpRFrP9dktT9j7oyh4rj0mr6Q63bSyiZk4hOiYm7Fva7aWpI624dRsoGkl7bArSHBddW+hX2npNJOnx1ZDPQX/Y91DK'
    'G0V3/RBnur2KUW5H/AGlgWJBGiYv62N/MhNTMGuLKC7WYAUR/o7GYk24xrn51uKjqURERFtGPwHQ74sB0K99jQFQIiKiHqUb3p7U'
    'nl2Y2n8di+e87+eu4yv7d6OgDSSGpenpaV+9DTy6232L+7UbWFD09DzzhxfxTX+5wC5MPnodc8LLj87Mn8eZ0DKCzASmzIjel+0W'
    'muYYsshgZlkYkux0huoGgQvxZSR1C1AGbtpoJQpS5FFtl4FSVnpmX+ftzP6nmvd7myqEhtjGbF+XT98UZU8gP2mhubCEinbos7xv'
    'KbaXpo60Ytb1epUqe/KppDkuaZbtSYp6TSjZ8Y0hnYO1YNj34MsbEnv98HUC3ssRx9+uzKFmlhF00FXqcX8yM5i1aqitRG2jgcWa'
    'nG8dVm0xfug8ERHRFiQGQHe98EK6AOjx4wyAEhER9WjbgX375Hmuc+exffYiXsN1PDmtCQJG2oGjxYNY+bL3RvQvA9/4/IFUwQkn'
    'RU/PTOEjePpsy91W9RbGxZ6ely9izBu6/uSVgzg/vUtIdD0x/RCmvvV68Pb3JyEHRkUZzMyKQRJ3CGquUIsegqoTt548lFon6L3m'
    'vVwlfxxls4Y54YGgdqXopnnDXAud10WjUqzA9oZqB+vEbT9zCOPi0O7GYswLcyKCPVF0ZY8k5ZufhNUsoVTTBAzFfUuzvbR1pBO3'
    'bnYMptgrMLZ+PWmOi6otIIsxIThnV+bUeaikqdcumraS5Pjqyt8ohl80lB2DaY4h21d5k0py/dAHPN3nbCqGvIv63J/8pBUdVG4s'
    'ohY8O9iXxViiofNERERb3PbtDIASERGtlX9786Zz0/ucPn3aGRVu3HM0PP/88/Isx6lbDoDgY5bb8hKeumPBdDrJdccK1hPny2lR'
    '6SI5X1e7bArrSfmZZaezuJRm1f0MHDNy+9HbC9WDZUUv43HLBgewnHpXflHfVWUPC+frqlvCPrlzYuo2+fZS11FAV4bodTv7Jtev'
    'vLz0PcVx6SqXX29CHma5rNm2ap4/P6pe5eXD3/VtJer4RtCWX3ccosorE8skl0/+HkF3/ZDS5GXqlq5cMtX+1B1LOF86y6rK7ae1'
    'nbIZfb1rl82Ia4y8DSIiIor09tvOj/7kT5wf/NqvOdeffNJ5a9++2M+1hx5yvvdLv+Tc+p3fcd65eFHOkYiIPBih+BKtLePmzZuO'
    'HwB96aWXcOTIEQDA9umXxdjoQK3OH5ZnaQ2zLHFW5w9jfn4e09PTcpLLriCXXcBUu5c3tdNw2KjksmjNikN5GygacxjjcdoEoo7v'
    'BrVu148GisYiJh3NM0H7thbbICIi2oRWV/H2q69itdnE26++ijt//ufyEl227d2Lex97DNsfewz3mibu+shH5EWIiLYswzBSjSam'
    'zUMZ9Fxvo9QotUFPAHYlh2yp6b1YaK2DF9TFriCXbWE2FGxh0HPTiDy+G9f6XD/cIfT+kH+rPsAAcqMIoxDkjPomOU5ERETrwVld'
    'xeqrr+LtZhN3zp7F6htvyIt0uevDH8b2Rx9lAJSIyDNK8SVaWyMb9BwlcUFPGi2NooEC6nBCURwGPTeL6ONLREREtLkxAEpE1BsG'
    'PbeukQ16jlKjZNCTiIiIiIhGhbO62hn+zgAoEZHWKMWXaG1tu3r1Kt566y1cuXIlmGkYxrpPHceJnL8eU588n1NOOeWUU0455ZRT'
    'TjnldK2n2+69F/d+4hP4z/6H/wH3nT6Nf3bjBt7/z/853lhdhcq7b70VvAX+Lx59FDePHsU/27kT7775Zlf+nHLKKaebbUpbE3t6'
    'JsCenkRERERENOr8HqCrZ89i9dVX2QOUiGjE4ku0tkY26DlKGPQkIiIiIqKNhAFQIiIXg55b1zZ5Ru9u4DPTL2P79Mu4Z/6GnJja'
    'QLogX7uEx794CbY8HzFpREREREREG5ixfTvu/cQn8IF/8S9w3+nT+NBXv4r3//N/ju0PPigvGhCHwP/Npz+N78/M4PbXvoZ333xT'
    'XpSIiGjkbTty8jV5nuvceWz3g5jT53FGTpecmW9h5VMPYXX+MO5M75KTUwtF4a9dwuNeWeRPkrJtHQ0UjRwqw4rkNoowDAPFhpyg'
    'kHb5oRtU/Qwqn03KriBnGDAMA7lRqaSRa4sbxCgey34NsC3YldzmqZeBSnmN3IztjDaZlG2aEhpUvQ4qn42ggaJ3vUz7b5mxfTu2'
    '/8zPBAHQ+7wA6D0xAdDbv//7DIASEdGGpe7p+cD9WJ07iIexG9+Yvx9PyOkht3Hhyk5MfWyHnNCzUE/PPQfwyvxhrM4fDsr0B973'
    'O7FlGxLvxjn6R4eNSq6TFl4m/GPF0PxIaxQNGNKvmUbRgJGrhHqo2pVc17x0kvxYbKBYWEG57aCa977HrZOvwmmXsVIoIsVvsgFK'
    'UMaRk6TMSZZZLzYqT5UwXnfgOA6WZzLyAutAbrs6o1y3a20EjqX2OuvTHbOItIFdl2wsLQBTE2tRLxH7sWbEbQ+jHCPQzgaq1zrq'
    'db1e9bO9ftbt1XpsU2WUyrKWRmm/+ylLP+sOUKJ/31Qs1B33min+rrEruVCe0X9Ecu9Rtv3s/4rLXgD0Q6EA6N+TVwjIAdCbXgD0'
    'HQZAiYhohG37hZ9/WJ4XzRsOfnLe72H5Bk5ec5POzL+OY5dv4disMF/onSkum1Sa5y3Y9TdCvT8fr98Opdcjyhx2Gye/GLeMoFGE'
    'UUDwg8O9ge7+wWJ5N3H+p5oHYF/AilUP5rXLQOmp6IBlftICVi4IaQ0s1kyYWMCSsEK71YQ5NYFh3iralTnUrFl07kfzqDrLwneF'
    'zAxmrRrmIn94UW8S1v26aKPVNDGWleevn+62qzPKdbvW1vlYJrjONooGDKOAmrhegrSBXJfsJSxgCmsS81xXwz4n1rmdEaU27HNi'
    'VG3V/R6e4D6hbqHW5x/i7EoO2dJ4599Mp47xUrY78GkvYQEmzGbnXsLYvh33egHQXS8cx7+58TD+KkEP0B96AdDvMgBKRESj7ObN'
    'm47/OX36tBNy9U0n93TLaQT//6dO8c+9tD9vOXd/4U2n7TiO49xyKl943alc9Ve85VS+oFo2GTfuGUEsU5Srbzq5p72ySGVuL73e'
    'KcfVN52c9/+N5//UyS3dcheKKOvzzz8vfGs7ZROOVRdmOY7j1C0HZtlbT7FMlLrlAJYTuWi77JgwnbJfmHbZMa26U7fEvOuOFSzj'
    '/X/ZcgA4ABwzWLnuWN48/+Pm4Za1Mz+qLG2nbArlcBxpu/5+CPmIO98uO2ZQNzJVufw00ynXy44ZKlt4neh6jtovXf04CfN1pH33'
    'thPsX1Qe7jKhbYXai09T5lAdSHXf9tP8T8K0yLL68+Vtxhxjx1GUXy6DvP2I7QyUpu1GtgOxbnXL+elR9SdL0MZV2+iq83JXW6pb'
    '4e/tsikcG1UZ4+o+6lgq1ok7vvK+CfsUrk+R4hoaed5I50OIJk17XYrXLpte+WPO77j6UX73RR0LXb4R+mpHfrkU5ZDbt7iOfOy7'
    'ROW5lu1MsW6QJpYt4vzpWi9qf6LqX27Y8nqmYyrblGrbwYKKc14kby+ujkVR6/ZSV2J61HoieZvDrh+fvF/Svw3B/w9637315HNA'
    'pjxeSdYX9yHie2Sblfdbs2/a3x6iNPn4ecjtQbhuqPIJyOsOqS1F1p+G/Htfq+5YXcfULUvXZiLuMfx/uzr/hsmE/FdXnR/9yZ84'
    'P/j1X3euf/KTzlv79sV+/urjH3e+d/Soc+t3fse5c/GinDkR0bqBKr5Em166oKcYDAx9l4Ke8rLOdafoByL7FRf0FMsilyNyf647'
    'xVB+3WUNBT1VP0zaZccMfljIP6oifoh42mUz4ibeF77xD36g1K3OD6jQdr0fYH5axI+dQNd6EfvkCy3rE9bpqpO6Y4V2OCrwpNBV'
    'LjGg6AR1EvxQiyybT94vXf30lm84UKDOQz7OcoChQ1HmUB2Iy0jb7DUtrt5jj7EwX15ObP+hOo/YzqBFHseoduCXWSy/bjld/cWI'
    'quuo81lR591BTdMxrXDbcpN1ZUxS99HHsrstxhzf0L7J3xV11rXv4nx5HbmcIl1aiutSl/C66vM7rn7k+lWVR07T5Svpqsu07Ujc'
    'tlwO79+6IJ+2U7b8YIF87ijK15XnGrYzed3QOS6JOn8ityHtj6L+u4XXi2tT0WXWnfOy6HqPrmOZvK6kqz4GX97h14/u30/N/g9k'
    '3731Yq/PquOVdH35+HvflW1WrgPVOaCrO5mujuL2US6/nI/q2IbXHXhbUtafhra8snp30FPZluV6Ev7taqv+8BeRv8MAKBFtfAx6'
    'bl3qZ3r24+otfDM0YwfG9t9C62poppaR8u3tnSHur+PZy3KqZ88OjMvzrt3GCq7jk8Hw+Ba+Ii/TZRyH5OE9mUMYxwou2ACQwcxy'
    'Z1h7u2yiVoh4fpBdwVOlJqzZGcXQ9AwmpkzUFt0BL+2W9/y47BhMb9i7vbSApjWJziN9TJSPe9/yk7CCMsF7jpH3rJ9sCc1Qmka7'
    'JW0jShOttv//eVRDD0/M4NC4mC7TlctE+ZRQP/YSFpoWZv3xVZkJTJk1eFWUgKJ+esi3dSKHubF25xl0mjwyM7OwgqFE7mMKkj8L'
    'UKoDUeMESuI2k6ZpyuqK2qbuGCs0FlGDhUl/0fwkLMRtZ4CUbVduB6r2qVgutv5kMW3c30boOoLIOs8cGu889qKxiJo1i1NjrVDb'
    'GssmKWMvdS+tk+T4hvZN/q67BsVdZ/sVd13SkIa2K8/v2PrpUep8+2hHOvYSFprCMUUGM1W/fcjnTprjtpbtTC6nWFe9nrey7vqP'
    'o2xTgLrMsee8Rmwdx0lYVwMq79DrR15Waxj7HnN9jj1eMevHStJmFeeA7rdHJEUdxe6jTM5Hd052DKctJak/T6MIo1CDVa9G/FZJ'
    'qN1C0xxD9CVbKEvjBErj3uN+MhOYQgknlPUpueceYQj8C4lfgsQh8EREtN6GE/TcuxOPXL4lPIvyNlqXd2Jsb2gprTTP9MS587j/'
    'xR3ey40ewm/ulxfwXLuNFXnenh0YF16M5L4c6UEc3SMvKIr4IWVfwErUTbr/g6orqNJAMVsCym3ty1U6N6YNLNa8/DMTmPKe69lu'
    'NWHG3pnC+1FewEq5HTzrx5IX6VVmBst1C7VCLw9kT1mudgtN1FAIHtSeRakpL9SD1Pk2UatJC2jzyGPSamJhyXZ/yJsDfBag8oeu'
    'Jk1b1gj9HONQGbIYM0OpIyBpmYTlUtVfyjbuU9V5fhJWswX3nrAGazKPzCG4bcu+gBW/baUqYx+GdnzTXWfXku2+wUgIKGjO72HV'
    'T9J8+21HOu0Wmmt1PJLub1/EfHs8b2Wq+o+laVMhvV6XIvRcx2nqalDlHXL9tFtQJYWtx757ej5eMXpuswLVb49Y0n4Max9DBtyW'
    '0tSfXUEu8YsWNbJjML3rebfOH7AaizWgVgiV2e9YkYoiALr9oYfkJQNRAdAf/u7vMgBKRERDN5yg555dmNp/HYvnvO/nruMr+3ej'
    'oA0khhkpenraV28Dj+523+J+7QYWFD09z/zhRXzTXy6wC5OPXsec8PKjM/PncSa0jCAzgSlTDmDG/ZVV1vmRHPum2vyk+xfoxgWs'
    'BD3W/B5K8l+kNby/mCf/y3tK+WrQszXVA9l7KlfnrZX+p68fi4E0+Zoot9uYWshKP2bVeeQnLTQXllBZrA32xVPKH7oxaZqyRur1'
    'GIfK0EYr7Y3d0CUtk7xcwvrrqY17Ius8j0mrhsVGA4s1rxdMfhLjpRNouN3BhbaVsIz9GMbxHch1dlhsLC00MS5F+pTn9zDqBynz'
    '7bsdKWTHMJQYRJQ0+9szId9+zltZZP3HU7apELku+jjne63jVHUl59t7eYdaP0nb9jrtO9DH8UqixzYb0P720JH2Y5j7KBh4W0pV'
    'fwP4w1HmEMajesE2FlEL8nev9eEy12HVFjVlS0AIgIpvgU8SAL157BgDoERENHTqoOe589g+exGv4TqenNYEASPtwNHiQax82Xsj'
    '+peBb3z+gOJHRDQnRU/PTOEjePpsy91W9RbGxZ6ely9izBu6/uSVgzg/vUtIdD0x/RCmvvV68Pb3JyEHRkUZzMxKP2DsCnKFWmeY'
    'ul1BToiE2ZU5YYhOioAn0PkL9NwCxLGG+UkLtbm5ZL1x4P+AF3pONRaltxpHBBh8wnD6SI1i+O2Q2TGYocCEjQsriqGSseWS5I+j'
    'bIbfumxXit2PDgho9kuUOl+4beFUGSv+owvi8shPwmqWUKrF3SAlLDO8PMUfumL9adNiyiqLPcYKXhmC7Xg3iMFwtVg2KjkDRnA+'
    'xX2XxLVdpCiTuFya+kvbxn2aOs+OmViZm0Mt+EOIG8Cam1vpBOPSlLFXfR9flQTX2b6J1yW3HQX1bVeQMzqPJLErORi5ituOvGGN'
    'XfsYdX5r6yeLMSGw6/47oSNcF7T5SvptR12EcnhDOgvi+Vj06mmQ0uxvP8R8ez1vgXAdaeq/m3Ttj2pTsl6vS0DvbQoIr5umrgZV'
    'Xgy5fuQgkmqf1mzfJamPl0xz/UnVZiPofnvECdVR3D6m+K3UZYhtKU39ZWaw7MjD2mN+10TK43jZlIKrDRQLNZjl427+jUXhWu/L'
    'Ykz1iIdeCAHQ+7wA6E4GQImIaL1pX2S0jkbpQbPht7d76lboRUWdB7Y7nYefBx/hgeDSetHrh7XLZvfLFdrl8MsaHCfigeURD2v3'
    't2lZirSIh5dHvvAj4oH2wSeirJEPS3epyyXvj0/aXqK8re63nkd+T5JvVL0q6kLKo27Jx6ybvsxO1/ZDbUo6rto0ZVlV29Qc40DE'
    'un5b7VovYtkuES9K0X6Xqdqual/EMumWi0hXtpc0bTzh9r06DZ/+US9CUJVR3m4UeRn5uyfx8Y37HkF7nfXOp8j906c5jnxdinop'
    'RadsoRdd1IUXyUkiz29l/YT3zyyXtfXRdY3W5RvSbzsKH6eucsj5d73wpDuPMDlN/u5R7q+8fNx3kaZuej5v5TrSb0PUXbdRbSou'
    'PyldbveCru0p67ibvK6+roZU3iHXj/rfT9Xxjlqul7LI7UtBebwSrq+8/qjKLeYrb0P6rqw7mWpbHuU+yu0hpjyS4baluPUE0r81'
    '3syY3zX16BcNyW0x9G+m/HKpjtC/b46jzb9X7/kvQfqN33Bu/Df/TdcLj6I+/kuQbn/ta3wJEhENBEYovkRry7h586bjB0Bfeukl'
    'HDlyBACwffplMTY6UKvzh+VZWsMsS5zV+cOYn5/H9PS0nOSyK8hlFzDVXnYfDL5J2ZUcsq1ZOKnGXrkaRSP8wp8ty0Yll0VrNuUQ'
    'tlQaKBpzGItsj7q0zau77Sath6TL0dAN4Trb63WpUTSwOBl1Dq/F+U2DNerneFSbGvUyy4ZZ3lGvn1Eqy3raCPUw6m1J1EDRWMRk'
    'Vw/RQRlu/s6dO1htNrF69ixWX30Vq6+/Li/S5a4PfxjbH3sM9z72GO557DHc/ZGPyIsQEcUyDCPVaGLaRNjTM15kT09B5y+r8l9r'
    'NxP9X82V2mXHHPBfjDesIdVF3RLyrFtSbzd12tYht135u0rS5WgtDPQ62/O5qOkB03OetH5G/ByPbFMjXuYuQyzvyNfPKJVlbW24'
    '3x4j35ZE4Z6kyg6haYVGVsh1MRzsAUpEa2mU4ku0ttTP9FxnGykKn5lZ9h4IPmp/DR6kPKr1cZSymrdQyhpFGNkSxuvD+WvxRtM4'
    'UUKz63lKgyC8QbQA1JfFZx7q0raKHtoujZyBXWf7ui7lUVX0fhne+U1bFduUHutnlG2s3x4bqy3lUY16eVK/xJcvKf6dGzTDfwbo'
    'r/wK7jt9Gh/6nd/B+3/pl3p6Bui7ly7JixIREQEAlMPb19sodT/WDm8nIiIiIiKivjl37uDOq6/i7Vdf7WkI/HbTxF0HDsiLENEW'
    'N0rxJVpb265evYq33noLV65cCWYahrHuU8dxIuevx9Qnz+eUU0455ZRTTjnllFNOOeV0MNNt27dj++HD+KD3Fvh/euMG3v9Lv4TX'
    'V1ehIvYA/YtHHsHNmRn8s5078e6lS135c8opp1t3SlsTe3omwJ6eRERERERE68O5c8ft+en1Ar3DHqBElMIoxZdobY1s0HOUMOhJ'
    'REREREQ0At55B283m6kCoNs+/GE3+PnYY7iXAVCiLYdBz61rgC8yuoHPTL+M7dMv4575G3JiagPpgnztEh7/4iXY8nzEpBERERER'
    'EdHouftu3Hv4MD7wK7+CXS+8gPu8lyDdo3kJ0nveEPjvHzuGG0eO4ObMDG7zJUhERJueOuh57jy2+0HM6fM4I6dLzsy3sPKph7A6'
    'fxh3pnfJyamFovDXLuFxryzyJ0nZto4GikYOlWFFchtFGEaKN2CnXX7oBlU/g8pnk7IryHlvbc2NSiWNXFvcIEbxWPZrgG3BruQ2'
    'T70MVMpr5GZsZ7TJpGzTlNCg6nVQ+dCGxAAoERFpbHtNnuN74H6szh3Ew9iNb8zfjyfk9JDbuHBlJ6Y+tkNO6Fmop+eeA3hl/jBW'
    '5w8HZfoD7/ud2LINiXfj7H/CN9A2KrlOWngZOU39I61RNGBId+aNogEjVwn1ULUrua556ST5sdhAsbCCcttBNe99j1snX4XTLmOl'
    'UMQA4gs9SFDGkZOkzEmWWS82Kk+VMF534DgOlmcy8gLrQG67OqNct2ttBI6l9jrr0x2ziLSBXZdsLC0AUxNrUS8R+7FmxG0Poxzr'
    '2c6GsT9x1nqba729OL2Wp9f1hmGUyrKWRmm/+ylLP+v2IdG/Z9Q3RQB0+8c/Li8ZEAOgf/PpTzMASkS0yWw7q4x6Srzh4Cfn/R6W'
    'b+DkNTfpzPzrOHb5Fo7NCvOF3pniskmled6CXX8j1Pvz8frtUHo9osxht3Hyi3HLCBpFGAWg7rg3ae4NdPcPGMu7ifM/1TwAewkL'
    '4/VgXrsMlJ6KDljmJy1g5YKQ1sBizYSJBSwJK7RbTZhTExjmraJdmUPNmkXnfjSPqrMsfFfIzGDWqmFuzX9dbmYJ635dtNFqmhjL'
    'yvPXT3fb1Rnlul1r63wsE1xnG0UDhlFATVwvQdpArkv2EhYwhTWJea6rYZ8T69zOiFIb9jkxqrbqfg9OcF9Qt1Dr+w9vFEsIgH7o'
    '9OlEAdB3r1xhAJSIaLP5139507l50/2cPn3aCbn6ppN7uuU0gv//U6f4517an7ecu7/wptN2HMdxbjmVL7zuVK76K95yKl9QLZuM'
    'G/eMIJYpytU3ndzTXlmkMreXXu+U4+qbTs77/8bzf+rklm65C0WU9fnnnxe+tZ2yCceqC7Mcx3HqlgOz7K2nWCZKu+yYsJzIRdtl'
    'x4TplP3CtMuOadWduiXmXXesYBnv/8uWA8AB4JjBynXH8ub5HzcPt6yd+VFlaTtlUyiH40jb9fZfzF/c+XbZMYO6kanK5aeZTrle'
    'dsxQ2cLrRNdz1H7p6sdJmK8j7bu3nWD/ovJwlwltK9RefJoyh+pAqvu2n+Z/EqZFltWfL28z5hg7jqL8chnk7UdsZ6A0bTeyHYh1'
    'q1vOT4+qP1mCNq7aRledl7vaUt0Kf2+XTeHYqMoYV/dRx1KxTtzxlfdN2KdwfYoU19DI80Y6H0I0adrrUrx22fTKH3N+x9WP8rsv'
    '6ljo8o3QVzvyy6Uoh9y+xXXkY98lKs+1ameqbYvzIs4ZeTsB1bkmkrdpOqay7Qxje3F1KdOVoZe6UpSnq33KOyOvN+x688n7KP3b'
    'EPy/uEySeojOv2s9+RyQKY9jkvXFfYj4HnlM5P3W7Jv2t4coTT5+HnJ7EK4bqnwC8rpDakuR9SeQf9/T2rpzx/nRn/6p84Pf+A3n'
    'xuSk89a+fbGfv3r4Yed7R486t772NeedN9+UcySiDQCq+BJtevADnomCnmIwMPRdCnrKyzrXnaIfiOxXXNBTLItcjsj9ue4UQ/l1'
    'lzUU9FT9UGmXheCl/KNK/cNavtEMC9/4BzfZdavzAyq0Xe8HmJ9WtxQ/dqPWi9gnX2hZn7BOV53UHSu0w1GBJ4WucokBRSeok6DO'
    'Isvmk/dLVz+95Rs+fuo82mUztB/q464oc6gOxGWkbfaaFlfvscdYmC8vJ7b/UJ1HbGfQIo9jVDvwyyyWX7ecrv5iRNV11PmsqPPu'
    'oKbpmFa4bbnJujImqfvoY9ndFmOOb2jf5O+KOuvad3G+vI5cTpEuLcV1qUt4XfX5HVc/cv2qyiOn6fKVdNVl2nYkblsuh/dvXZBP'
    '2ylbfrBAPncU5evKcw3bWde2JVHnTN/Xg/A249rOoLenr0uZrgwSXV11pQl5KNpnt7WuN92/n3KdCrr2tZeyeOvFXp9VxzHp+nK7'
    '8L4rj4lcB6pjrKs7ma6O4vZRLr+cj+rYhtcdeFtS1p9AWz5aS+/5AdD/+X9mAJRok2PQc+tSv8ioH1dv4ZuhGTswtv8WWldDM7VC'
    'z/RMoDPE/XU8e1lO9ezZgXF53rXbWMF1fDIYHt/CV+RluozjkDy8J3MI41jBBRsAMphZ7gxrb5dN1Arh5wfZlRwMw0ChZmqeCZfB'
    'xJSJ2qI7AKbd8p4flx2D6Q17t5cW0LQm0XlUoYnyce9bfhJWUCZ4zzHynieULaEZStNot6RtRGmi1fb/P49q6OGJGRwaF9NlunKZ'
    'KJ+a6Qzdt5ew0LQw64+vykxgyqzBq6IEFPXTQ76tEznMjbU7z6DT5JGZmYXV9B9L4D6mQH3cZVIdiBonUBK3mTRNU1ZX1DZ1x1ih'
    'sYgaLEz6i+YnYSFuOwOkbLtyO1C1T8VysfUni2nj/jZC1xFE1nnm0HjnsReNRdSsWZwaa4Xa1lg2SRl7qXtpnSTHN7Rv8nfdNSju'
    'OtuvuOuShjS0XXl+x9ZPj1Ln20c70rGXsNAUjikymKn67UM+d9Ict7VsZ7KE52pf14MOZdsBhrK9+LqUKcoAJK+r2GPQ3T7jDL3e'
    '5GW1EtZDqrLEXJ9jj2PM+rGSHBPFMdb99oikqKPYfZTJ+ejaXMdw2pKm/hpFGIUarHo14rcJrTXDHwL/uc/hPn8I/NGjHAJPRLSJ'
    'bPv6FXnWAOzdiUcu3xKeRXkbrcs7MbY3tJRWmmd64tx53P/iDu/lRg/hN/fLC3iu3caKPG/PDowLL0ZyX470II7ukRcURfyQsi9g'
    'Jeom3f9BJQVVMjPL3nN9xlHKqp/r07kxbWCx5uWfmcCU91zPdqsJM/bOFN6P8gJWym0vGFuHJS/Sq8wMlusWaoVeHtCeslztFpqo'
    'oRA8DD6LUlNeqAep822iVpMW0OaRx6TVxMKS7f6QNwf4LEBzDMoWoErTljVCP8c4VIYsxsxQ6ghIWiZhuVT1l7KN+1R1np+E1WzB'
    'vSeswZrMI3MIbtuyL2DFb1upytiHoR3fdNfZtWS7bzASAgqa83tY9ZM0337bkU67heZaHY+k+9uXNOdqr9cDmabthAxqe/3Upbhs'
    'mrrSULXPWEOut3YLqqSwNPXQY1lUej6OMXo+JgLVb49Y0n4Max9DBtyWdPVnV5BL/GJFWmsMgBIRbU7bll8bQtRzzy5M7b+OxXPe'
    '93PX8ZX9u1HQBhLDjBQ9Pe2rt4FHd7tvcb92AwuKnp5n/vAivukvF9iFyUevY054+dGZ+fM4E1pGkJnAlBnRO6jdQrOXH3nZMZhR'
    'N/e+/KT7F+jGBawEPdb8HkryX6Q1vL+YJ//Le0r5atCzNdUD2nsql9V5uYn3GcyPxzT5mii325hayEo3A+o88pMWmgtLqCzWBvvi'
    'KS9wEUmXpilrpF6PcagMbbTS3tgNXdIyycslrL+e2rgnss7zmLRqWGw0sFjzesHkJzFeOoGG2x1caFsJy9iPYRzfQV9nB8rG0kIT'
    '41KkT3l+D6N+kDLfvtuRQnYMQ4lBREmzv71Kda7KZej9XFO2nZDBba/3uhSWTVVXMSLbZ7yh1lvStp2qHnosi0rPxzGBHo9JQPvb'
    'Q0faj2Huo2DgbUlbf2v0hyLqyyACoN8vlRgAJSIaAerh7efOY/vsRbyG63hyWhMEjLQDR4sHsfJl743oXwa+8fkDih8R0ZwUPT0z'
    'hY/g6bMtd1vVWxgXe3pevogxb+j6k1cO4vz0LiHR9cT0Q5j61uvB29+fhBwYFWUwMyv9gLEryBVqsGa9oUR2BTkhEmZX5jpDdBrF'
    'UJDMXlpAU/lXZXT+Aj23AHGsYX7SQm1uLllvHEQEVxuL0luNIwIMPmE4faRGETlx7H52DGYoMGHjwopiqGRsuST54yib4bcu25Vi'
    '6NEBYZr9EqXOF25bOFXGiv/ogrg88pOwmiWUanE3SAnLDC9PcbiXWH/atJiyymKPsYJXhmA73g1iMFwtlo1KzoDR6SYR810S13aR'
    'okzicmnqL20b92nqPDtmYmVuDrXgDyFuAGtubqUTjEtTxl71fXxVElxn+yZel9x2FNS3XUHO6DySxK7kYOQqbjvyhjV27WPU+a2t'
    'nyzGhMCu+++EjnBd0OYr6bcddRHK4Q3pLIjnY9Grp0FKs7+pCfuT5lzt9XoAdF/jo9qObFDb66cuxWXT1FUXoTya9tltDestcwjj'
    'qn8/RWnqodeyROnnOAL660+qYxJB99sjTqiO4vYxxW+lLkNsS7r6y8xg2eGw9o0mFABdWMCHvva1RAHQ2y+8wAAoEdEo+Le6Fxmt'
    'o1F60Gz47e2euhV6UVHnge1O5+HnwUd4WLn/koWoNIV22ex+eYD8sgbH6Xo4u/zdzcd/k6SlSIsqT9QLPyIeaB98IsqqeaC+ulzy'
    '/vik7SXK2+p+63nk9yT5RtWroi6kPOqWfMy66cvsdG0/1Bal46pNU5ZVtU3NMQ5ErBtq8/L2Vfn4Il6Uov0uU7Vd1b6IZdItF5Gu'
    'bC9p2njC7Xt1Gj79o16MoCqjvN0o8jLyd0/i4xv3PYL2OuudT5H7p09zHPm6FPVSik7ZQi+6qAsvkpNEnt/K+gnvn1kua+uj6xqt'
    'yzek33YUPk5d5ZDz73rhSXceYXKa/N2j3F95+bjvYfL+6M9VRT06EelyexPI23Qi284Qt6esS5m+DPq6Uh+D7n/j1NsQde3HkOtN'
    '/e+nan+iluulLHL9KSiPY8L1ldcfVbnFfOVtSN+VdSdTbcuj3Ee5PcSURzLctqRZr+slR7SRvbe66rz98supX4J0c2aGL0EiWgcY'
    'ofgSrS3j5s2bjh8Afemll3DkyBEAwPbpl8XY6ECtzh+WZ2kNsyxxVucPY35+HtPT03KSy64gl13AVHsZqj8ObwZ2JYdsaxZOqrFX'
    'rkbRCL/wZ8uyUcll0ZpNOYQtlQaKxhzGItujLm3z6m67Sesh6XI0dEO4zvZ6XWoUDSxORp3Da3F+0/oZ5vUgqu0Mc3tJjUIZdEa1'
    '3nyjVJb1tBHqYdTbEo06584d3Dl7Fm+/+ipWX30Vq9/6lrxIl7v27cO9jz2Gex59FPeaJu46cEBehIgGyDCMVKOJafNQBj3X2yg1'
    'Sm3Q0w+qlJrecx4364+jHn/82RXksi3McjjP0OqiUSwCVS/PRhHG3Bjay+4QYF3a1iG3Xfm7StLlaC0M9Drb87nYQNFYxGTUej3n'
    'SRvDEK8HkW1niNtLbBTKoDGy9eYbpbKsrQ3322Pk2xJtKO+8g9VXX2UAlGjEjFJ8idaW+pme62wjNcjgTezOZv5hlEe1Po5SNsVb'
    'PBtFGNkSxusMAgBA40QJTct/dt4gCW8QLQD10I2FLm2r6KHt0sgZ2HW2r+tSHlVFUHN45zdtdmw7vWG9jbKN9duDbYkG6u67sf2n'
    'fxof+Nzn8KHTp3Gf9wzQe1I8A/Sm9wzQd/gMUCKivrGnZwJxPT2JiIiIiIiIIr3zTtD78+1XX8WdhD1Atz/2GLY/+ii2mybuZg9Q'
    'op6NUnyJ1ta2q1ev4q233sKVK1eCmYZhrPvUcZzI+esx9cnzOeWUU0455ZRTTjnllFNOOeVUO73nHtz70z+ND/7Kr2DX6dP4f924'
    'gfcfPYrXVleh8u6VK/ih1wP0Lx55BDdLJfzTnTvxzqVL3flzyimnsVPamtjTMwH29CQiIiIiIqKBEnqArp49i9XXXpOX6OL3AL33'
    'scdwz2OPsQcoUQKjFF+itTWyQc9RwqAnERERERERDQ0DoERDw6Dn1jXAFxndwGemX8b26Zdxz/wNOTG1gXRBvnYJj3/xEmx5PmLS'
    'iIiIiIiIiNbK3XfjXv8lSAsLuO9rX8MHZmaw/eGH5SUD/hD4m88+i+96L0H64e/9Hl+CRETkUQc9z53Hdj+IOX0eZ+R0yZn5FlY+'
    '9RBW5w/jzvQuOTm1UBT+2iU87pVF/iQp29bRQNHIobKhI7mbYR+GhXWTzCjXk65sujSdXteT9ZJPL+ush1Ev56iXL4pY5jTlT7Ps'
    'euuxrHYFOe/N1bnUKw9Sj+Vfc2nK2UDRfyu4YaDYkNP70CgG+RpGEYPMmogoNS8A+v7jxxkAJSLqgzro+cD9WJ07iIexG9+Yvx9P'
    'yOkht3Hhyk5MfWyHnNCzUE/PPQfwyvxhrM4fDsr0B973O7FlG5LQj2P5h7eNSq6TFl5GlxbWKBowpIRG0YCRq4R6qNqVXNe8dNLc'
    'cAzaem571A2qbgaVz0a0lfd9o5KOmfZaKwoHQwyj+1pJ62FQ5+Cg8hk2G5WnShivO3AcB8szGXmBIUpTR2mWHTUW6o5bv9U8hHO/'
    'O1DZKBowuvbT+x0mXx/yVTiOA8epwxLnExGtNyEAep8XAH0/A6BERIls+3rnpe163nDwk/N+D8s3cPKam3Rm/nUcu3wLx2aF+ULv'
    'THHZpNI8b8GuvxHq/fl4/XYovR5R5rDbOPnFuGUEjSKMAoIf3U67jJVC98245d30+B/3x3l8mi8/aQErF4Qf5Q0s1kyYWMCS8Eu9'
    '3WrCnJrAWt5aEW0MeVSdZaxp3IEGLrhe1i3UCt2BjQ4T5XbnutqeWkBWvjATDVUbraaJsaw8f9RsxmtjDYvi6W5XMFcTvgfzl7AA'
    'E2Yz/FuKiGgjMPwh8AyAEhElsu1y0qAnAFy+iNaDXg/LZ3bg2ar7TMwnph/Cc/t34rm5w7gz/yCO7rmNk9WLGH+me9mkjBTP9MwU'
    'HnR7gXo9QfHi+U7gUijz+U8hshxn5l/Hwse9ofmxZbVRmavBqlcRxCkzM1iuW6jNDbhXUXYs/KPcvoAVaxaz40202v5CbiB0akK4'
    'c1nq9IzqDK3r7gXV6XlaQA1NlLLRPSXU6/ppOVQa/nA+f/0kQ9A0247cByTMFz2US7ePXqBbTPcThWGM4d4k3vYriv1Q5RdIWzeq'
    '8mvygZ8eUTa/B0zc/im/i+Q06buyLlTHKoq4bAGd+9wk24o5Vso60OUnzIsqe+QxlKXZf12eunxU9RYhcl89fW9b017z1c4fhbJj'
    'MLGCC6oqk2QmpmCG/nDki9p29/yuMivrQJVfFGkbSepR1wZ1aX2XN6Y+YkUcTz8pcl9V29PkE1Ct66epznFdmi/mOhmIKqeXv/xv'
    'kfK4KcojHMvu8kGxbY+yruOujVGi2m+Cupf3P7ROzLUnERPlcvh3mL20AJTLXb027aUFYOoUTpWBBUY9iWgDkwOgH/rd3+05APru'
    '5cvyokREm4J6eHuU/Qdx7AHv//fuxCNScuDaDSxAWPaB3Xj68nXU43pQCtL09AzZswtT+4XvQpkzH9uNRy7fkm5SbmDx7G7MFryh'
    '+XFltZewENWLIzsGs9lCEIsEUCtE3QDEpwUyE5gyOwFOe2kBGMsiP2mh5ndnsC9gBeM4FMQ8myi1JoNeUc3SCe8GI4+q0LPU7Z1a'
    'RAMZzCzXYQU9pIRgbkC1rq+J0hxwyvHXd2++VsptxfI+1bZV+5A0X1+acmn20a4gV1gRepD5Q98aKGY7wxid+jhKWale/P1olwF/'
    'P5T5idLWjar8qnw625mYMtFcWArOi8Zizes5HLd/A6CsC92xkknLRtanbluIqNenvGCApg6U+amORdy2RGn2H5o8dfkkrDfo9hUR'
    '205yvqrSYtpru4Vm6Hqn1zhRQnP8kNQLXrVt6I+dsg50+cm8wJRV9/Jooww/KKuqR00b1KUNpLya+khEdTxV+6raniofkWpdn7xN'
    '8bzTpSHmOhleLrqc8r9FmuPmLx/6tyMLYzHi35IQzbYj61qgbCsyVftNUPe6f4uV20tpYhJW8IfiBk6UgKmJQ9JCNpYWgKmJDDIT'
    'U4BwTImINjLj7rux/fHHew6A/s2RIwyAEtGmlC7omdTVW/hmaMYOjO2/hdbV0EwtI0VPT0Ac4v46nlVdp/fswLg879ptrOA6PhkM'
    'j2/hK/IyXSJuujOHMB70QMpgZrlzA9Aum6gV/B4VujSZe6PlBzjbLfeHOrJjQe8le2kBTWtSuAE0UT7ufctPwgr1ihJ6VmRLaKbo'
    'MaVf10T51Ezn5s9ewkLTwqw/bi4zgSlTGnampdiH1PmmLZduH8XetXlUq3mgsYgaLEz6lZ+fhBUaXifsR6h9IDq/RBR1A8SUXy0z'
    'Mxu6UQx6Dsfu36BE1EXssRLIy2pFbAuIqFdvudg6UOWnOxaKbYnkfdLtP6DOU5ePnBZLta/ythOcr7o0lUYRRkHqZd/F7+Xmfgqo'
    'w5HPrdht645dRB3E5ifw/mgW1BcymKn61yhFPeraoC4N6L+8QEx99Eqxr0Cf29OtK29TrBtdmkt5nUxE+rco9rjJ/3bI39PUibxv'
    'qnUj2opM235j6l73b/HA5HG8DJRONNw6tma7h+83TqA07s3PTGAKJZxQtn0ioo2JAVAioo7hBD337pR6VN5G6/JOjO0NLaXlpOnp'
    'ee487n9xh/dyo4fwm2JPT9G121iR5+3ZgXHhxUjuy5EexNE98oKiiJuGrh6XHZmZ2cibqLg0AMgcGvee69nAYs3LPzOBKe+5nu1W'
    'E2ZXt9MoDRSNXntWpFy33UITNRT8GyAji1JTXqgH/earXV+zj/7jC6J655pj6NR+FmNm8EVNl1/PNOWPlcek1XSH+TUWUTOnENzL'
    '97J/aajqQnusJO0WVEkhqm11kfZTVQfK/NIcC0Wdptn/LkKeunyS1ht0+6oRu21FWhShJ1pULKYj/EzProAn4ratOXaqOtDmJ0nZ'
    'UzWgaoO6tEGUV1cfQ9HP9tKsqzjvAE2a5jrZC9VxWw+qtiJTtt80dZ/y2pNSZmIKZm0RxcUarCCq3NFYrAG1QqjtByNniIg2ITEA'
    '+qHTpxkAJaItZ9uj6utd7/bswtT+61g8530/dx1f2b8bBW0gMcxI0dPTvnobeHS3+xb3azewoLgen/nDi/imv1xgFyYfvY454eVH'
    'Z+bP40xoGYE05DzQbqEZuokZkLw3XKtxASuW36Mzg0PjTbTaKXqbeD1LeupZ0dO6nber+p+o+EN6/earWD9uH4O3urpDBIOXqYQe'
    'adBGK+mdnCq/XsWVP0Z+0kJzYQkVechmr/uXhrIuFMdKlh1D4niBclsiaT91dRCVX6pjoavThPvfRc5TkU+aeoNiX+Vluii2HZsW'
    'JSrY0ivFtuOOnbIOFPnJ0ta5T9cGdWn9ljeuPgatn+2lWlc+R0TqNOV1she647YelG1FoGq/qepek88gZGYwa9VQWykj6JAaaGCx'
    'Jrf9OqzaYve+EhFtRnfdxQAoEW0525SXuHPnsX32Il7DdTw5rQkCRtqBo8WDWPmy90b0LwPf+PyBVDcIToqenpnCR/D02Za7reot'
    'jIs9PS9fxJg3dP3JKwdxfnqXkOh6YvohTH3r9eDt709CDoyKMpiZlW4I7ApyhRqsWW/4ll1BTugmYVfmOkPZdGmRvN4lc+7zPIO5'
    'kxZqc3NYSdrbRH4BSGNRenFARCDXF7uuJH8cZbOGOWHMvl0pKobwQ79tUep8Jbr1dfvYKIZfHJEdg2mOIesNSQzy82781MfSo8pP'
    'XCaQsG505Qfi88lPwmqWUKoJN67a/ctiTHzebGVO0yY0y6rqQnesZJlDGBeHhnbtu0e1LXEZSPupqwNVfrHHQqBqM2n2XxYqvyaf'
    'pPUGzb6Ky8h029alAd3tNTOD5cjnOPZAt23dsVPVgS4/mTeUvBD8G2CjUpRfhiPRtUFd2iDKq6uPVGKuP77Y7WnyiV1XoDrv4tKi'
    'rpO90B23vmnqSEXVVsRloGm/aeoecdceG5Wc7kVK8fKTVnRQurGIWvDHY18WY9rHOxARbVJSAPQ+LwB6DwOgRLTZ3Lx50/E/p0+f'
    'dkaFG/ccDc8//7w8y3HqlgMg+JjltpjoWEIaYDn1RGnR2mXTAUwntIl22TEBB5a4dt2xQsuFv7v5eNu1LEVadHnU68rb9En7aZad'
    'rkU84W3L+UV9T5KvvJ44P3p9/T6Kx0zI1z8O8vyu7YvfNflJ0tSNuvxyPtHqltyedPvnhM4Bs1xW1LdHuayuLtTHqot4PnYdu7ht'
    'qeZ7lHWgXk99LNTryMeza1nl/uvyjEgX81HWm0y1jagyy98V29akdbXXdtkxlWXzydsWyWlJti3XiaoOItKUx8rpXtaqK8onfFe2'
    'QV3aYMqrr4+o/++2Vtcx9bop6qIrLbxfkdfJEHkd+btHe9zSfA9LXtfy/6vqQCYt69WFvu4j8lNee9pO2dTVcd2xuo6/YhuOI6S5'
    '+YZ/q7naZTPi2iRvg4hoi3jnHedHf/Znzg++9CXnxqc+5by1b1/s568eecT53syMc+t3f9d559IlOUeikYARii/R2jJu3rzp+AHQ'
    'l156CUeOHAEAbJ9+WYyNDtTq/GF5ltYwyxJndf4w5ufnMT09LSe57Apy2QVMtZe7H5hPtGHYqOSyaM0qhrluag0UjTmM8Rwm2qR0'
    '57guTbaVr5OjooGisYjJQfX8jrQW2yAi2gDefRdvv/oqVl99Fatnz2L1m+FXFUe56+/+XWx/9FFsf+wx3GuauGu/6mUbRGvLMIxU'
    'o4lp81AGPdfbKDVKbdATgF3JIVtqei/RSHLjRDRi7Apy2RZmt+RNXpqgBxFtPLpzXJcm2dLXyVHhvjTJHw5v1QcYgG4UYRSCnFHn'
    'cSYi6mAAlDa4UYov0doa2aDnKIkLehJtdI2igQLq0W+73vRSBD2IaAPSneO6tLCtfZ0kIiLyMABKa8T5wQ/wzne+I8/uycMPP4zX'
    'XntNnk0j5K49e7Dt7/wdeXbfRjboOUqReAY9iYiIiIiIiAReAPTO2bNuIJQBUBqgdy9dwl+bpjybNqkfn5/H+55Qv1K8V9uuXr2K'
    't956C1euXAlmGoax7lPHcSLnr8fUJ8/nlFNOOeWUU0455ZRTTjnllNMtOb37btzrvwV+YQFTN27gA6USvrm6CpV3/+N/xA9feAHf'
    'f/ZZ/PnDD+NmqYR/snMn3r10qTt/Trf09KMf/SiI+sWengmwpycRERERERFRAv32AH3sMdx14IC8CG0x7Om5tQyrp+fIBj1HCYOe'
    'REREREREROk4777bCX6mCIDe+9hjuOfRRxkA3cLkoKfxwQ/i7p/8ydAySX3rW9/Cxz/+cXk2raN3r13De3/1V8H3DRD0vIHPTLfw'
    'vwFwHh3Dneld8gKpDKSn57VLeLwK/P8+fwBd7yfQpUkY9CQiIiIiIiLqHQOglIYc9LzrwAH8RLMZWiapgcSXaKC+Nz2NH505E3wf'
    'VtBzmzwjcO48tk+/jO3TL+Oe6fPoFCXamfkWVj71EFbnD/cd8AQQbpDXLuFxryzyJ0nZto4GikYOFVueT7QG7ApyhgHDMJAb2UY4'
    'qHOkl3x6WWcUDKrcunx0aXH6WTdOA0WvTRuGgWJDTu9BoxjkZxhFDCJLIiIioo3AuOsubM/l8IFjx3DfwgI+9Hu/h/eXStj+yCPy'
    'ooF3/+N/xO3Tp/H9Z5/F30xN4fu//Mu4/Xu/h3cvXZIXJSLqog56PnA/VucO4mHsxjfm74c+3nobF67sxNTHdsgJPfMfXgsA2HMA'
    'r8wfxur84aBMf+B9vxNbtiEJ3bgqbobFZXIVBPfkQnBId9PbKBowpIwbRSkvAHYl1zUvnbigQVz6ZrQV97kfNipPlTBed+A4DpZn'
    '4vpPJ8FjEG9QdSTmM6g8R1iS63fAQt1x23U1788LB0MNo/taDdio5Lqv18hX4TgOHKcOS5xPREREtIWIAdAPnT7NACgRDcW2r3de'
    '2q537RIe/+IlnJz3e1i+gZPX3KQz86/j2OVbODYrzBd6Z4rLJpWm67FdfyPU+/Px+u1Qej2izGG3cfKLccsIGkUYBQQ3wk67jJWC'
    'dOMsL7M84w2jDweHnDpQkG+KPflJC1i5IKQ1sFgzYWIBS8IK7VYT5tRE7DD93uVRdZYxkDjWhrEV97kfbbSaJsay8vx+8Bisna1X'
    '11ZwDbZQK6j/+NTNRiVXAPz1/QBm6FoNwF7CAkyYzfD1moiIiIgk27YxAEpEQ7HtctKgJwBcvojWg14Py2d24NnqJdgAnph+CM/t'
    '34nn5g7jzvyDOLrnNk5WL2L8me5lkwr19IyRKTzo9gL1eoLixfOdwKVQ5vOfQmQ5zsy/joWPe0PzY8tqozJXg1WvIuj0k5nBct1C'
    'bc4PXkYs47OXsCAGh7JjMJsttKXFAD9NuGG2L2DFmsXseBOtYAU3EDo1IUQrljq9mDrDjLt7JrlBWvcGvoYmSllVz1Op55fUS0rV'
    'wyk0xLlR7PR4CvV0FXuUyT3MInqfNfx1NeWsSPsvlDe+PsT9E8oSV+ZeyiWmyfkp91mXl66cuuMWrouuwymKzL+HNiR/jyxbRF0k'
    '2u+IfZdFniOIqQsxrYCamCSL3B9P39tOeu52n4eNYvi7XclJda3JU1luWVRZo+jqU5VH1Dnh697f0HVHlK92em1mx2BiBRe6FlKw'
    'l7DQtDAZurjnUQ3+sOWylxaAqVM4VQYWGPUkIiIiSkYVAH34YXnJAAOgRKSy7dKVFFHP/Qdx7AHv//fuhPLvLtduYAHCsg/sxtOX'
    'r6Me14NSkKanZ8ieXZjaL3wXypz52G48cvmWdAN8A4tnd2O24A3NjyurHLT0icFLb5mVuaibdgAYx6FQjyrFDXdmAlNmJ8BpLy0A'
    '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'
    'sYbtlNJT/k3PWWauKDyDEBmatO16m1z4FlPQomn0FWhxSBDtprgtjZL7PdbuoGVzjZA4coW+Bk5ZDwnNcwuKbnzOCIIgCIIgCIKY'
    'H5ClZwD8LD0JgiAIgiAIgiAIgiCI1kCWnt1CYw3bMUvP6elpHDlyBIcPH3Y9NU2bddc0TaH/bLgOvD+55JJLLrnkkksuueSSSy65'
    '5JJLLrmtddeuXQuWt99+WxgvqEssTBb19PRg1apVWL16tevpWFjOpqsxlp6z7Trw/uSSSy655JJLLrnkkksuueSSSy655LbW3b9/'
    'P1guu+wyYbygLrEwoW96EgRBEARBEARBEARBEAQxr2ih0vME7hp8FUsGX8X5oyf4wNC0xAT56EFc//BBVHl/+IQRBEEQBEEQBEEQ'
    'BEEQBNG1yJWeb+3FEkeJObgXtc+Litk5WsbUJzbizOhmvDe4nA8OjcfS8+hBXG/Lwl9BZOsaiiloWjMnJxeR0mLIdpUml5V5NuSf'
    'jTIJYo7S9BwUgKbLUD2zqjA/wqWtZmOIBY08bwnXZvOHbq53N8tOEARBEATRGlIaoGkQrkmyMSsslq35pbh7AEDRiqfFgGzK/ltx'
    'Sd+/nHycKwYyzmshi17nfRyuuhxnRi7FNViBF0cvh/oMpdPYd3gZBtYv5QMaxmPpufJifH90M86MbnZl+o59/56vbG3CXrg7l+hh'
    '8cSJZWsDt5pFzE3LLD7iOZiVDKYSKYiysygixZRbl/ecgBZVcuZC27RKBr98wii0/cJFNJKmC2nHXCOkiFRiCpmKiVzcvufnGk00'
    'N4n6oYpsjE9ry97oPKdpc0jJWMXEGDDQH+EDCCIAomemnXS6PIIgCIIgiO4gP8IpGKvAWJAD4qtALAHoGcCcBIZygGnWLgN2GOM3'
    'JFs6xLk0k4AsKhGeRa9JtZ4c9nbwJ0cdC8s9ePKoFbRzdDfuO3QK9w0z/ox1Jhs3KGG+6Vkt7PFYf15fOO0JLwhk9nIaTz7sF4eh'
    'mIKWAAqmCdM07QU8p4zg40wO2QO3iFQ0jb6C5V/JAOkos/iPDGHYyGNEuTrRkanY+ZomKgNjiAo1IcTcI46cOSmf8OYVYeoaJu4C'
    'gp9HWjnXcFSzI8gbw1wfcHNNZgoJRvFZTGnQtATybBIGwy7buSxlamPznGmamJwrA6Q6gTEMgHSeRHdA8ytBEARBEASPngGMEjDB'
    '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'
    'w8b1+tWNj4Kh7mMH6Tjh69WorCK49LacjT7b9Yjy59N47xsuu6E+5fFra0FdAz77dTJJ+9usl6OuXwTPhmnWp2vq/zdb3/q6141z'
    'giAIgiCIOcDZd94xj6xe7V7HrruOjxIYWHZ17mUApp7x+qnC+XvThLX+A0znHExV+iBXXZpC/T0AM1PxCePzqsB9x3XCnXUcK3vB'
    'ENfF/2qMn27Z4unfXxaLfJSWoJ08edJ0FKCvvPIKbrvtNgDAksFXWd1oSzkzupn3UtJOWfw4M7oZo6OjGBwc5IMsqlnEomMYqEwG'
    '/mXAj2JKw0hvBZOtypCYBxSR0kbQ28JxRlSRjUVRHg66PXaWacNcw1LNxhAtD4u/EdwG5v48Vz8+iinrMKcONREBNDH3NZpuoVE/'
    'zgmCIAiCIOYCMwcP4l299rGgxRdfjItKft8rEqNpGtjNxCkNmMoAk0NsrBp8OH/PxsvbW8XZd05ZfCFFQEsw9zpQmbQOGXLydzAK'
    'tS3oqjBUgZj9nU4AKFSAkSgwwMhZzQLRtJMagNHoqfRMw4bgZ4OD+NXO2p7yC0dHccGN8j3XjSJVes421qBsrPFajVLp6SgL0iX7'
    '4I0mF1jVLGLRMobNXOCTt4iFAC3gW04XPmstnWvq6OAY64a2r5OxiJQ2juRclnmeUEylgJzdzsUUtJFeVCYDnHruoYPjuZupG+cE'
    'QRAEQRBzg3YqPYlW0ljDdkrpKf+m5ywzVxSeQYgMTdp2vU0uroopaNE0+gq0+CCIdlPclkbJ/R5rd9CyuUZIHLlCH9LRMN8dbYAu'
    'mefqx0ccOVIMdQjm5PIEUAit8CSCUj/OCYIgCIIgCGL+QJaeAfCz9CQIgiAIgiAIgiAIgiBaA1l6dguNNWzHLD2np6dx5MgRHD58'
    '2PXUNG3WXdM0hf6z4Trw/uSSSy655JJLLrnkkksuueSSSy655LbWXbt2LVjefvttYbygLrEwWdTT04NVq1Zh9erVrqdjYTmbrsZY'
    'es6268D7k0suueSSSy655JJLLrnkkksuueSS21p3//79YLnsssuE8YK6xMKEvulJEARBEARBEARBEARBEMS8ooVKzxO4a/BVLBl8'
    'FeePnuADQ9MSE+SjB3H9wwdR5f3hE0YQBEEQBEEQBEEQBEEQRNciV3q+tRdLHCXm4F7UPi8qZudoGVOf2Igzo5vx3uByPjg0HkvP'
    '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'
    '9JxthafDfFV88taN7OVYVO79n//fujDnevvHBwAAqz+8yvW7cv3vAABO/eKUazH5gWXL3PALf8NSSJ47d85TjnnOxMzMDKqVHwMA'
    'XhyfwMdv/EN8/MY/xFeyljXFuXPn3PgXXHCBm+dv/uZvAEx9zHPWVxKcMq5c/zv4d3/8h9j1g9fxH1NpZJ74Mk6c+Ke6+rBXu3B+'
    '/XV+gdy06WqPBcWmTVdjenoajz32KDYx267vvvsuN/6mTVd37KXBoVQqYXp6WvryEI8nOv6iLaMZWZ95Zjs+9rHf573bhq+sCbGs'
    '+/fvx5t79jS0Va9Rwsh6xRW/41GmvfLKdzuuDAwjbxhYRRDLI4883HElNFpQz04pAsLKuev1NwKNb0eJ1glaWYcrr7zSE7dYKHTF'
    '3ANbySyTVfZ8ODyzfbu0zFYTto6y/tq/fz+mp6c9dVq5cmVHFYHHjh3DAw8+WLdtneXll1+WPgvJ5K2Ynp5Gybb2TKXuavjTFu3E'
    't8+4/9+7don7jMdvXLaaVtdj//79KBYLUgvpl19+GVdccQWSyVvx5pud+yGr0/V02L9/P958s3PvQK2sZ6d/3OIJU5dmxpGu6545'
    'Z3z8BaxfX/sxvJ341dGPJ574Am6+OY4nnrB+KPK7Xwi8+GKx6e31Ip3n97//fTz33HP4p3/6Jz5ISCfeVwmi3XiUnnNF4ekwHxWf'
    'MzNnpdfvXPXbWHHRchRe3InyVLkufGbmLC6+1Pq+5YG3D7h+5akyAOCC978PMzOWhaUJ0w0/d85SKJ47N+O9N89hZuYsLrHzvHrT'
    'Rjz40P2e65JL19TFn5k5i3OmpQjl750y/unECQz8+z/EX//N0/jTT/57/Pfv/QMKLxbr6sNe7ebZZ7/u2W7hbJPbtesN9PT04IEH'
    'HnTD7r77LvT09LjxO71FBLblnuqX6Q996ENSa85nn/064vEE7902wsr6yCMPu8rkXR3eZhRIVsbi0OEb33gW8UTn2hQhZV27di2e'
    '++a3cPsf/xE2XXM1fvjDH9ZZBLSbMPLyPPvs1+vad9M1V2PTNVfjRz/6UZ2yYHz8BUxPT0sXOe0kTD3Xb9jgsfZy5O4EYeQMwzPb'
    't+P3f/8PeO+20Mo6pFLWt56dcfXnn/zknJ57nK3Qm+zvP/Kyqp4Ph1zuqY7++BG2jjIciyUWZ+HejFIgLMnkrei/6SY88sjDfBAA'
    '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'
)

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

