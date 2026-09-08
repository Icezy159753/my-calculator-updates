"""Delete Total + NA Batch ? standalone Python application.

Copy this file alone to another folder or Windows computer.
Requires Python 3.10+, Microsoft Excel Desktop, PyQt6 and pywin32.
Install dependencies once: python -m pip install "PyQt6>=6.6,<7" "pywin32>=306"
Run: python DeleteTotalNA.py

Includes processing, GUI, styles, and the application icon. No local imports,
assets folder, or requirements.txt is needed to run after installing dependencies.
"""
from __future__ import annotations

import os
import re
import shutil
import posixpath
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}

XL_UP = -4162
XL_TO_LEFT = -4159
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
XL_EDGE_LEFT = 7
XL_INSIDE_VERTICAL = 11
XL_CONTINUOUS = 1
XL_THIN = 2
GRAY_235 = 235 + (235 * 256) + (235 * 65536)

# Range() rejects address strings longer than 255 characters.
ADDRESS_LIMIT = 240
# Below this width a mixed block is cheaper to probe cell by cell than to keep
# bisecting, which bounds the scan cost when merges are scattered.
BISECT_FLOOR = 8


@dataclass(slots=True)
class ProcessResult:
    sheets_scanned: int = 0
    merged_areas_changed: int = 0
    total_columns_deleted: int = 0
    na_columns_deleted: int = 0

    def summary(self) -> str:
        return (
            f"{self.sheets_scanned} ชีต · ยกเลิก Merge {self.merged_areas_changed} จุด · "
            f"ลบ TOTAL {self.total_columns_deleted} คอลัมน์ · "
            f"ลบ NA {self.na_columns_deleted} คอลัมน์"
        )


def validate_input_file(path: Path) -> None:
    if not path.is_file():
        raise FileNotFoundError(f"ไม่พบไฟล์: {path}")
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"ไม่รองรับไฟล์ชนิด {path.suffix or '(ไม่มีนามสกุล)'}")


def make_working_copy(source: Path, work_dir: Path, index: int) -> Path:
    validate_input_file(source)
    work_dir.mkdir(parents=True, exist_ok=True)
    destination = work_dir / f"{index:04d}_{source.name}"
    shutil.copy2(source, destination)
    return destination


def _cell_text(value: Any) -> str:
    return "" if value is None else str(value)


def _column_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _column_index(letters: str) -> int:
    index = 0
    for letter in letters:
        index = index * 26 + (ord(letter) - 64)
    return index


# Building the address in Python and calling ws.Range() costs a fraction of
# ws.Cells(row, col): measured at 1.8 ms against 11.8 ms per access. Every hot
# loop below goes through these helpers rather than Cells().
def _cell_ref(row: int, col: int) -> str:
    return f"{_column_letter(col)}{row}"


def _row_span_ref(row: int, start: int, end: int) -> str:
    if start == end:
        return _cell_ref(row, start)
    return f"{_column_letter(start)}{row}:{_column_letter(end)}{row}"


_ADDRESS_PART = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$")


def _address_span(address: str) -> tuple[str, int, int, int]:
    """An area address as (head reference, head row, head column, last column).

    Reading MergeArea.Address once replaces separate Cells(1, 1), .Row and
    .Column round trips, and the last column tells the caller which columns the
    area covered without asking Excel again.
    """
    parts = address.split(":", 1)
    head = _ADDRESS_PART.match(parts[0])
    if head is None:
        raise ValueError(f"ไม่รู้จักตำแหน่งเซลล์: {address}")
    letters, digits = head.groups()
    head_col = _column_index(letters)
    last_col = head_col
    if len(parts) == 2:
        tail = _ADDRESS_PART.match(parts[1])
        if tail is None:
            raise ValueError(f"ไม่รู้จักตำแหน่งเซลล์: {address}")
        last_col = _column_index(tail.group(1))
    return f"{letters}{digits}", int(digits), head_col, last_col


def _contiguous_runs(columns: list[int]) -> list[tuple[int, int]]:
    runs: list[tuple[int, int]] = []
    for col in sorted(columns):
        if runs and col == runs[-1][1] + 1:
            runs[-1] = (runs[-1][0], col)
        else:
            runs.append((col, col))
    return runs


def _chunk_addresses(pieces: list[str]) -> list[str]:
    """Join area references into multi-area addresses that stay under the limit."""
    chunks: list[str] = []
    current: list[str] = []
    length = 0
    for piece in pieces:
        if current and length + len(piece) + 1 > ADDRESS_LIMIT:
            chunks.append(",".join(current))
            current, length = [], 0
        current.append(piece)
        length += len(piece) + 1
    if current:
        chunks.append(",".join(current))
    return chunks


def _column_address_chunks(runs: list[tuple[int, int]]) -> list[str]:
    return _chunk_addresses(
        [f"{_column_letter(start)}:{_column_letter(end)}" for start, end in runs]
    )


def _last_value_column(ws: Any, row: int) -> int:
    edge = _cell_ref(row, int(ws.Columns.Count))
    return int(ws.Range(edge).End(XL_TO_LEFT).Column)


def _merged_columns_in_row(ws: Any, row: int, last_col: int) -> list[int]:
    """Columns of `row` that are merged, located by bisection.

    A whole block reports MergeCells False in a single call, so an unmerged
    stretch costs one round trip instead of one per cell.
    """
    found: list[int] = []
    pending = [(1, last_col)]
    while pending:
        start, end = pending.pop()
        if start > end:
            continue
        if end - start + 1 <= BISECT_FLOOR:
            found.extend(
                col
                for col in range(start, end + 1)
                if bool(ws.Range(_cell_ref(row, col)).MergeCells)
            )
            continue
        state = ws.Range(_row_span_ref(row, start, end)).MergeCells
        if state is False:
            continue
        if state is True:
            found.extend(range(start, end + 1))
            continue
        middle = (start + end) // 2
        pending.append((middle + 1, end))
        pending.append((start, middle))
    found.sort()
    return found


def _workbook_merge_index(path: Path) -> dict[str, list[str]]:
    """Read merge geometry once, without thousands of Excel COM queries.

    Excel still performs all edits, including column/reference adjustments.
    Binary/encrypted workbooks retain the existing COM discovery path.
    """
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return {}
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    try:
        with zipfile.ZipFile(path) as archive:
            relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            targets = {
                rel.attrib["Id"]: posixpath.normpath(
                    posixpath.join("xl", rel.attrib["Target"])
                ).lstrip("/")
                for rel in relationships
                if rel.attrib.get("TargetMode") != "External"
            }
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            result = {}
            for sheet in workbook.findall("s:sheets/s:sheet", ns):
                target = targets[sheet.attrib[f"{{{rel_ns}}}id"]]
                root = ET.fromstring(archive.read(target))
                result[sheet.attrib["name"]] = [
                    merge.attrib["ref"]
                    for merge in root.findall("s:mergeCells/s:mergeCell", ns)
                ]
            return result
    except (OSError, zipfile.BadZipFile, KeyError, ET.ParseError):
        return {}


def _unmerge_from_index(ws: Any, addresses: list[str]) -> int:
    if not addresses:
        return 0
    areas = []
    for address in addresses:
        head, row, col, end_col = _address_span(address)
        end_row = int(_ADDRESS_PART.fullmatch(address.split(":")[-1]).group(2))
        areas.append((row, end_row, col, end_col, head, address))
    last_row = int(ws.Range(f"C{int(ws.Rows.Count)}").End(XL_UP).Row)
    first = min((max(2, r) for r, end, c, right, _, _ in areas
                 if c <= 3 <= right and end >= 2 and max(2, r) <= last_row),
                default=0)
    if not first:
        return 0
    last_col = _last_value_column(ws, first - 1)
    count = 0
    selected = []
    for area in sorted(areas, key=lambda a: a[2]):
        row, end, col, right, head, address = area
        if not (row <= first <= end and col <= last_col):
            continue
        count += min(3 - count, min(right, last_col) - col + 1) if count < 3 else 1
        if count >= 3:
            # Vertical/single-column merges can shift into the next merge.
            # Keep sequential Excel behavior for that less common layout.
            if row != first or end != first or right == col:
                return _unmerge_and_shift_first_merged_row(ws)
            selected.append(area)
    span = ws.Range(_row_span_ref(first, 1, last_col))
    raw = span.Value
    values = (raw,) if last_col == 1 else raw[0]
    for address in _chunk_addresses([a[5] for a in selected]):
        ws.Range(address).UnMerge()
    if selected:
        left, right = selected[0][2], selected[-1][2] + 1
        heads = {a[2] for a in selected}
        # A typical exported header contains only merged labels and blanks.
        # Write those labels together; do not rewrite unrelated populated cells
        # (which could contain formulas, rich text, or literal '=' strings).
        only_labels = all(
            col in heads or col > len(values) or values[col - 1] is None
            for col in range(left, right + 1)
        )
        if only_labels:
            shifted = [None] * (right - left + 1)
            for col in heads:
                shifted[col - left] = ""
                shifted[col - left + 1] = values[col - 1]
            ws.Range(_row_span_ref(first, left, right)).Value = (tuple(shifted),)
        else:
            for row, end, col, right, head, address in selected:
                ws.Range(f"{head}:{_cell_ref(row, col + 1)}").Value = (("", values[col - 1]),)
    span.Interior.Color = GRAY_235
    span.Font.Bold = True
    return len(selected)


def _unmerge_and_shift_first_merged_row(ws: Any) -> int:
    last_row = int(ws.Range(f"C{int(ws.Rows.Count)}").End(XL_UP).Row)
    first_merged_row = 0

    for start in range(2, last_row + 1, 256):
        end = min(start + 255, last_row)
        if ws.Range(f"C{start}:C{end}").MergeCells is False:
            continue
        for row in range(start, end + 1):
            if bool(ws.Range(f"C{row}").MergeCells):
                first_merged_row = row
                break
        if first_merged_row:
            break

    if first_merged_row == 0:
        return 0

    reference_last_col = _last_value_column(ws, first_merged_row - 1)
    row_range = ws.Range(_row_span_ref(first_merged_row, 1, reference_last_col))

    merge_count = 0
    changed = 0
    # This deliberately follows the original VBA: merged cells are counted while
    # walking the row, and processing begins once that counter reaches three.
    # Unmerging only clears merge flags, so a column the pre-scan reports as
    # unmerged stays unmerged; re-checking each candidate as it is reached keeps
    # the left-to-right counting identical while skipping the untouched columns.
    candidates = _merged_columns_in_row(ws, first_merged_row, reference_last_col)
    index = 0
    while index < len(candidates):
        col = candidates[index]
        index += 1
        cell = ws.Range(_cell_ref(first_merged_row, col))
        if not bool(cell.MergeCells):
            continue
        merge_count += 1
        if merge_count < 3:
            continue
        merged_area = cell.MergeArea
        head_ref, head_row, head_col, area_last_col = _address_span(
            str(merged_area.Address)
        )
        value = ws.Range(head_ref).Value
        merged_area.UnMerge()
        # Clearing the head and filling the cell to its right is one write.
        ws.Range(
            f"{head_ref}:{_column_letter(head_col + 1)}{head_row}"
        ).Value = (("", value),)
        changed += 1
        # Every cell of the area just unmerged now reports MergeCells False, so
        # asking Excel about the rest of them would only confirm what is known.
        while index < len(candidates) and candidates[index] <= area_last_col:
            index += 1

    row_range.Interior.Color = GRAY_235
    row_range.Font.Bold = True
    return changed


def _find_total_row(ws: Any) -> int:
    values = ws.Range("C1:C1000").Value
    for row, (value,) in enumerate(values, start=1):
        if _cell_text(value).upper() == "TOTAL":
            return row
    return 0


def _add_left_borders_to_previous_row(ws: Any, target_row: int) -> None:
    if target_row <= 1:
        return
    previous_row = target_row - 1
    last_col = _last_value_column(ws, previous_row)
    values = ws.Range(_row_span_ref(previous_row, 1, last_col)).Value
    values = (values,) if last_col == 1 else values[0]
    columns = [
        col
        for col, value in enumerate(values, start=1)
        if value is not None and _cell_text(value).strip() != ""
    ]
    # Giving every cell of a run its own left border is the same as drawing the
    # run's left edge plus the vertical lines inside it. Borders() applies to
    # every area of a multi-area range, so a whole header row of runs — real
    # ones average around a dozen — takes two calls rather than two per run.
    runs = _contiguous_runs(columns)
    if not runs:
        return
    # Single-cell runs have no inside edge, so each chunk asks only for the
    # edges its own areas can carry.
    singles = [_row_span_ref(previous_row, s, e) for s, e in runs if s == e]
    spans = [_row_span_ref(previous_row, s, e) for s, e in runs if s != e]
    for pieces, edges in ((singles, (XL_EDGE_LEFT,)),
                          (spans, (XL_EDGE_LEFT, XL_INSIDE_VERTICAL))):
        for address in _chunk_addresses(pieces):
            span = ws.Range(address)
            for edge in edges:
                border = span.Borders(edge)
                border.LineStyle = XL_CONTINUOUS
                border.Weight = XL_THIN


def _delete_extra_total_and_na_columns_legacy(ws: Any, target_row: int) -> tuple[int, int]:
    last_col = _last_value_column(ws, target_row)
    found_total = False
    total_deleted = 0
    col = 1

    while col <= last_col:
        if _cell_text(ws.Range(_cell_ref(target_row, col)).Value).upper() == "TOTAL":
            if not found_total:
                found_total = True
                col += 1
            else:
                ws.Columns(col).Delete()
                last_col -= 1
                total_deleted += 1
        else:
            col += 1

    last_col = _last_value_column(ws, target_row)
    na_deleted = 0
    col = 1
    while col <= last_col:
        if _cell_text(ws.Range(_cell_ref(target_row, col)).Value).upper() == "NA":
            ws.Columns(col).Delete()
            last_col -= 1
            na_deleted += 1
        else:
            col += 1

    last_col = _last_value_column(ws, target_row)
    if target_row < 3:
        raise ValueError(
            f"พบ TOTAL ที่แถว {target_row}; Macro เดิมต้องมีอย่างน้อย 2 แถวก่อนหน้า"
        )
    letter = _column_letter(last_col)
    final_range = ws.Range(f"{letter}{target_row - 2}:{letter}{target_row - 1}")
    final_range.Borders.LineStyle = XL_CONTINUOUS
    return total_deleted, na_deleted


def _delete_extra_total_and_na_columns(ws: Any, target_row: int) -> tuple[int, int]:
    last_col = _last_value_column(ws, target_row)
    headers = ws.Range(_row_span_ref(target_row, 1, last_col))
    # Formula headers may change after each deletion: retain the original path.
    if headers.HasFormula is not False:
        ws.Application.Calculation = XL_CALCULATION_AUTOMATIC
        try:
            return _delete_extra_total_and_na_columns_legacy(ws, target_row)
        finally:
            ws.Application.Calculation = XL_CALCULATION_MANUAL
    if target_row < 3:
        raise ValueError(f"พบ TOTAL ที่แถว {target_row}; ต้องมีอย่างน้อย 2 แถวก่อนหน้า")
    raw = headers.Value
    values = (raw,) if last_col == 1 else raw[0]
    seen_total = False
    delete: list[int] = []
    total_deleted = na_deleted = 0
    for col, value in enumerate(values, 1):
        text = _cell_text(value).upper()
        if text == "TOTAL":
            if seen_total:
                delete.append(col)
                total_deleted += 1
            seen_total = True
        elif text == "NA":
            delete.append(col)
            na_deleted += 1
    # One multi-area delete per address chunk, and the chunks run right to left
    # so the column indices still queued stay valid.
    for address in reversed(_column_address_chunks(_contiguous_runs(delete))):
        ws.Range(address).Delete()
    last_col = _last_value_column(ws, target_row)
    letter = _column_letter(last_col)
    ws.Range(
        f"{letter}{target_row - 2}:{letter}{target_row - 1}"
    ).Borders.LineStyle = XL_CONTINUOUS
    return total_deleted, na_deleted


def process_workbook(
    excel: Any,
    workbook_path: Path,
    on_sheet: Callable[[float, str], None] | None = None,
) -> ProcessResult:
    result = ProcessResult()
    book = None
    try:
        merge_index = _workbook_merge_index(workbook_path)
        book = excel.Workbooks.Open(
            str(workbook_path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALCULATION_MANUAL

        worksheets = book.Worksheets
        sheets = [worksheets.Item(index) for index in range(1, int(worksheets.Count) + 1)]
        sheets = [ws for ws in sheets if ws.Name not in {"Contents", "Info"}]
        total_sheets = len(sheets)
        # Two passes over every sheet plus the save: counting all of them keeps
        # the bar from reaching the end while work is still outstanding.
        steps = total_sheets * 2 + 1
        done = 0

        def step(detail: str) -> None:
            nonlocal done
            done += 1
            if on_sheet is not None:
                on_sheet(done / steps, detail)

        for position, ws in enumerate(sheets, start=1):
            result.sheets_scanned += 1
            addresses = merge_index.get(ws.Name)
            result.merged_areas_changed += (
                _unmerge_from_index(ws, addresses) if addresses is not None
                else _unmerge_and_shift_first_merged_row(ws)
            )
            step(f"ยกเลิก Merge · ชีต {position}/{total_sheets}")

        excel.Calculate()

        for position, ws in enumerate(sheets, start=1):
            target_row = _find_total_row(ws)
            if target_row:
                _add_left_borders_to_previous_row(ws, target_row)
                total_deleted, na_deleted = _delete_extra_total_and_na_columns(
                    ws, target_row
                )
                result.total_columns_deleted += total_deleted
                result.na_columns_deleted += na_deleted
                # Refresh cross-sheet formulas before inspecting the next sheet.
                excel.Calculate()
            step(f"ลบคอลัมน์ TOTAL/NA · ชีต {position}/{total_sheets}")

        excel.Calculation = XL_CALCULATION_AUTOMATIC
        book.Save()
        return result
    finally:
        try:
            excel.Calculation = XL_CALCULATION_AUTOMATIC
            excel.ScreenUpdating = True
        except Exception:
            pass
        if book is not None:
            book.Close(SaveChanges=False)


def _excel_dispatch(dispatch: Any) -> Any:
    """Share interface metadata within this Excel session, never COM objects.

    Dynamic pywin32 normally rediscovers the same Range/Border/etc. methods for
    each new object. Cache by interface IID so Excel versions still supply
    their own metadata, without a generated gen_py cache in the packaged EXE.
    The closure belongs to one worker thread and one Excel process only.
    """
    import pythoncom
    from win32com.client import dynamic

    interfaces: dict[str, tuple[Any, Any]] = {}

    class SessionDispatch(dynamic.CDispatch):
        def _wrap_dispatch_(self, ob, userName=None, returnCLSID=None):
            return wrap(ob, userName)

        def _make_method_(self, name):
            method = super()._make_method_(name)
            if method is not None:
                # pywin32's generated method bodies use a local Dispatch
                # global instead of _wrap_dispatch_ for return objects.
                method.__func__.__globals__["Dispatch"] = (
                    lambda ob, userName=None, resultCLSID=None: wrap(ob, userName)
                )
            return method

    def wrap(ob, name=None):
        try:
            info = ob.GetTypeInfo()
            iid = str(info.GetTypeAttr()[0])
        except pythoncom.com_error:
            return dynamic.Dispatch(ob, name, createClass=SessionDispatch)
        if iid == "{00000000-0000-0000-0000-000000000000}":
            return dynamic.Dispatch(ob, name, createClass=SessionDispatch, typeinfo=info)
        cached = interfaces.get(iid)
        if cached is not None:
            representation, lazy = cached
            return SessionDispatch(ob, representation, name, lazydata=lazy)
        obj = dynamic.Dispatch(ob, name, createClass=SessionDispatch, typeinfo=info)
        interfaces[iid] = (obj._olerepr_, obj._lazydata_)
        return obj

    return wrap(dispatch)


def process_files(
    jobs: list[dict[str, Any]],
    work_dir: Path,
    progress: Callable[[float, str, str, str], None],
) -> None:
    import pythoncom
    import pywintypes
    import win32com.client.gencache as gencache

    original_get_class = gencache.GetClassForCLSID
    gencache.GetClassForCLSID = lambda clsid: None
    pythoncom.CoInitialize()
    excel = None
    try:
        excel_clsid = pywintypes.IID("{00024500-0000-0000-C000-000000000046}")
        dispatch = pythoncom.CoCreateInstance(
            excel_clsid,
            None,
            pythoncom.CLSCTX_LOCAL_SERVER,
            pythoncom.IID_IDispatch,
        )
        excel = _excel_dispatch(dispatch)
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3

        for index, job in enumerate(jobs, start=1):
            source = Path(job["path"])
            progress(index - 1, str(source), "processing", "กำลังประมวลผล")
            try:
                working_copy = make_working_copy(source, work_dir, index)

                # Sheet progress lands between this file's start and its finish.
                def on_sheet(
                    fraction: float,
                    detail: str,
                    base: int = index - 1,
                    origin: str = str(source),
                ) -> None:
                    progress(base + fraction, origin, "processing", detail)

                result = process_workbook(excel, working_copy, on_sheet)
                job["processed_path"] = str(working_copy)
                job["result"] = result.summary()
                job["status"] = "ready"
                progress(index, str(source), "ready", result.summary())
            except Exception as exc:
                job["status"] = "error"
                job["error"] = str(exc)
                progress(index, str(source), "error", str(exc))
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        gencache.GetClassForCLSID = original_get_class


def save_processed_files(
    jobs: list[dict[str, Any]],
    progress: Callable[[float, str, str, str], None],
) -> None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    ready_jobs = [job for job in jobs if job.get("status") == "ready"]

    for index, job in enumerate(ready_jobs, start=1):
        original = Path(job["path"])
        processed = Path(job["processed_path"])
        progress(index - 1, str(original), "saving", "กำลังสร้าง Backup และบันทึก")

        backup_dir = original.parent / f"Delete Total+NA Backup {timestamp}"
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_path = backup_dir / original.name
        shutil.copy2(original, backup_path)

        pending = original.with_name(
            f".{original.stem}.delete_total_na_pending{original.suffix}"
        )
        try:
            shutil.copy2(processed, pending)
            os.replace(pending, original)
        finally:
            if pending.exists():
                pending.unlink()

        job["status"] = "saved"
        job["backup_path"] = str(backup_path)
        progress(index, str(original), "saved", f"บันทึกแล้ว · Backup: {backup_path}")


# ---------------- User interface ----------------

import shutil
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QObject, QThread, QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QColor, QDragEnterEvent, QDropEvent, QIcon, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)



APP_STYLE = """
QWidget {
    background: #F4F7FB;
    color: #172033;
    font-family: "Leelawadee UI";
    font-size: 10.5pt;
}
QFrame#hero {
    background: #13213C;
    border-radius: 18px;
}
QLabel#title {
    background: transparent;
    color: white;
    font-size: 24pt;
    font-weight: 700;
}
QLabel#subtitle {
    background: transparent;
    color: #B9C7E4;
    font-size: 10.5pt;
}
QFrame#dropCard, QFrame#statusCard {
    background: white;
    border: 1px solid #DDE5F1;
    border-radius: 16px;
}
QLabel#dropTitle {
    background: transparent;
    color: #24324A;
    font-size: 15pt;
    font-weight: 650;
}
QLabel#dropHint, QLabel#muted {
    background: transparent;
    color: #6D7890;
}
QPushButton {
    min-height: 40px;
    padding: 0 18px;
    border-radius: 10px;
    font-weight: 600;
}
QPushButton#primary {
    background: #3478F6;
    color: white;
    border: none;
}
QPushButton#primary:hover { background: #2368E5; }
QPushButton#primary:pressed { background: #1856C6; }
QPushButton#save {
    background: #0E9F6E;
    color: white;
    border: none;
}
QPushButton#save:hover { background: #087F5B; }
QPushButton#save:disabled, QPushButton#primary:disabled {
    background: #C7CFDC;
    color: #F7F9FC;
}
QPushButton#secondary {
    background: white;
    color: #33415C;
    border: 1px solid #CCD6E5;
}
QPushButton#secondary:hover { background: #F1F5FA; }
QPushButton#danger {
    background: #FFF4F4;
    color: #C74444;
    border: 1px solid #F3CCCC;
}
QTableWidget {
    background: white;
    alternate-background-color: #F8FAFD;
    border: 1px solid #DDE5F1;
    border-radius: 12px;
    gridline-color: transparent;
    selection-background-color: #E7F0FF;
    selection-color: #172033;
}
QHeaderView::section {
    background: #EAF0F8;
    color: #536078;
    border: none;
    border-bottom: 1px solid #D8E1EE;
    padding: 10px;
    font-weight: 650;
}
QTableWidget::item { padding: 8px; border-bottom: 1px solid #EDF1F6; }
QProgressBar {
    background: #E5EAF2;
    border: none;
    border-radius: 6px;
    height: 12px;
    text-align: center;
    color: transparent;
}
QProgressBar::chunk { background: #3478F6; border-radius: 6px; }
"""


STATUS_TEXT = {
    "pending": "รอประมวลผล",
    "processing": "กำลังประมวลผล",
    "ready": "พร้อมบันทึก",
    "error": "ไม่สำเร็จ",
    "saving": "กำลังบันทึก",
    "saved": "บันทึกแล้ว",
}

STATUS_COLOR = {
    "pending": "#6D7890",
    "processing": "#3478F6",
    "ready": "#0E9F6E",
    "error": "#D64545",
    "saving": "#B7791F",
    "saved": "#087F5B",
}


def format_duration(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:.1f} วินาที"
    minutes, remainder = divmod(int(round(seconds)), 60)
    if minutes < 60:
        return f"{minutes} นาที {remainder} วินาที"
    hours, minutes = divmod(minutes, 60)
    return f"{hours} ชั่วโมง {minutes} นาที {remainder} วินาที"


class BatchWorker(QObject):
    # Float, not int: a single large workbook reports fractional per-sheet steps.
    progress = pyqtSignal(float, str, str, str)
    finished = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, mode: str, jobs: list[dict[str, Any]], work_dir: Path):
        super().__init__()
        self.mode = mode
        self.jobs = jobs
        self.work_dir = work_dir

    def run(self) -> None:
        try:
            callback = lambda count, path, status, detail: self.progress.emit(
                count, path, status, detail
            )
            if self.mode == "process":
                process_files(self.jobs, self.work_dir, callback)
            else:
                save_processed_files(self.jobs, callback)
            self.finished.emit(self.mode)
        except Exception as exc:
            self.failed.emit(str(exc))


class DropCard(QFrame):
    files_dropped = pyqtSignal(list)

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("dropCard")
        self.setAcceptDrops(True)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 23, 28, 23)
        layout.setSpacing(7)
        title = QLabel("วางไฟล์ Excel ที่นี่")
        title.setObjectName("dropTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint = QLabel("รองรับหลายไฟล์พร้อมกัน: XLSX, XLSM, XLSB และ XLS")
        hint.setObjectName("dropHint")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        layout.addWidget(hint)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:
        paths = [url.toLocalFile() for url in event.mimeData().urls()]
        self.files_dropped.emit(paths)
        event.acceptProposedAction()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Delete Total + NA Batch")
        self.resize(980, 720)
        self.setMinimumSize(820, 620)
        self.jobs: list[dict[str, Any]] = []
        self.work_dir = Path(tempfile.mkdtemp(prefix="delete_total_na_"))
        self.thread: QThread | None = None
        self.worker: BatchWorker | None = None
        self.unsaved_results = False
        self._prompt_save_when_idle = False
        self._started_at = 0.0
        self._process_seconds = 0.0
        self._save_seconds = 0.0
        self._build_ui()
        self._update_controls()

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        page = QVBoxLayout(root)
        page.setContentsMargins(26, 24, 26, 24)
        page.setSpacing(16)

        hero = QFrame()
        hero.setObjectName("hero")
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(28, 24, 28, 24)
        hero_layout.setSpacing(6)
        title = QLabel("Delete Total + NA")
        title.setObjectName("title")
        subtitle = QLabel(
            "จัดการ Merge, คอลัมน์ TOTAL ซ้ำ และคอลัมน์ NA หลายไฟล์ในครั้งเดียว"
        )
        subtitle.setObjectName("subtitle")
        hero_layout.addWidget(title)
        hero_layout.addWidget(subtitle)
        page.addWidget(hero)

        self.drop_card = DropCard()
        self.drop_card.files_dropped.connect(self.add_paths)
        page.addWidget(self.drop_card)

        toolbar = QHBoxLayout()
        toolbar.setSpacing(9)
        self.add_button = QPushButton("เลือกไฟล์")
        self.add_button.setObjectName("primary")
        self.add_button.clicked.connect(self.choose_files)
        self.remove_button = QPushButton("นำรายการที่เลือกออก")
        self.remove_button.setObjectName("secondary")
        self.remove_button.clicked.connect(self.remove_selected)
        self.clear_button = QPushButton("ล้างรายการ")
        self.clear_button.setObjectName("danger")
        self.clear_button.clicked.connect(self.clear_jobs)
        toolbar.addWidget(self.add_button)
        toolbar.addWidget(self.remove_button)
        toolbar.addWidget(self.clear_button)
        toolbar.addStretch()
        page.addLayout(toolbar)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["ชื่อไฟล์", "ตำแหน่ง", "สถานะ"])
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        page.addWidget(self.table, 1)

        status_card = QFrame()
        status_card.setObjectName("statusCard")
        status_layout = QVBoxLayout(status_card)
        status_layout.setContentsMargins(20, 16, 20, 16)
        status_layout.setSpacing(9)
        self.status_label = QLabel("เพิ่มไฟล์เพื่อเริ่มต้น")
        self.status_label.setObjectName("muted")
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        status_layout.addWidget(self.status_label)
        status_layout.addWidget(self.progress)
        page.addWidget(status_card)

        actions = QHBoxLayout()
        actions.setSpacing(10)
        safety = QLabel("ไฟล์จริงจะยังไม่เปลี่ยนจนกว่าจะกดบันทึก · สร้าง Backup อัตโนมัติ")
        safety.setObjectName("muted")
        safety.setWordWrap(True)
        self.run_button = QPushButton("ประมวลผล")
        self.run_button.setObjectName("primary")
        self.run_button.clicked.connect(self.start_processing)
        self.save_button = QPushButton("บันทึกทับไฟล์เดิม")
        self.save_button.setObjectName("save")
        self.save_button.clicked.connect(self.confirm_save)
        actions.addWidget(safety, 1)
        actions.addWidget(self.run_button)
        actions.addWidget(self.save_button)
        page.addLayout(actions)

    def choose_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "เลือกไฟล์ Excel",
            "",
            "Excel files (*.xlsx *.xlsm *.xlsb *.xls)",
        )
        self.add_paths(files)

    def add_paths(self, paths: list[str]) -> None:
        if self.thread is not None:
            return
        existing = {str(Path(job["path"]).resolve()).lower() for job in self.jobs}
        rejected: list[str] = []
        for raw_path in paths:
            path = Path(raw_path)
            normalized = str(path.resolve()).lower()
            if (
                path.is_file()
                and path.suffix.lower() in SUPPORTED_EXTENSIONS
                and normalized not in existing
            ):
                self.jobs.append({"path": str(path.resolve()), "status": "pending"})
                existing.add(normalized)
            elif path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                rejected.append(path.name)
        self._refresh_table()
        self._update_controls()
        if rejected:
            QMessageBox.information(
                self,
                "มีไฟล์ที่ไม่รองรับ",
                "ข้ามไฟล์ต่อไปนี้:\n" + "\n".join(rejected[:12]),
            )

    def remove_selected(self) -> None:
        rows = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.jobs.pop(row)
        self.unsaved_results = any(job.get("status") == "ready" for job in self.jobs)
        self._refresh_table()
        self._update_controls()

    def clear_jobs(self) -> None:
        if self.unsaved_results:
            answer = QMessageBox.question(
                self,
                "ล้างผลลัพธ์ที่ยังไม่ได้บันทึก?",
                "ผลประมวลผลชั่วคราวจะถูกลบ แต่ไฟล์ต้นฉบับยังเหมือนเดิม",
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        self.jobs.clear()
        self.unsaved_results = False
        self._reset_work_dir()
        self._refresh_table()
        self._update_controls()

    def _reset_work_dir(self) -> None:
        shutil.rmtree(self.work_dir, ignore_errors=True)
        self.work_dir = Path(tempfile.mkdtemp(prefix="delete_total_na_"))

    def _refresh_table(self) -> None:
        self.table.setRowCount(len(self.jobs))
        for row, job in enumerate(self.jobs):
            path = Path(job["path"])
            name_item = QTableWidgetItem(path.name)
            folder_item = QTableWidgetItem(str(path.parent))
            status_key = job.get("status", "pending")
            detail = job.get("result") or job.get("error") or ""
            status_item = QTableWidgetItem(STATUS_TEXT.get(status_key, status_key))
            status_item.setForeground(QColor(STATUS_COLOR.get(status_key, "#6D7890")))
            status_item.setToolTip(detail)
            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, folder_item)
            self.table.setItem(row, 2, status_item)
            self.table.setRowHeight(row, 42)

    def _update_controls(self) -> None:
        busy = self.thread is not None and self.thread.isRunning()
        ready_count = sum(job.get("status") == "ready" for job in self.jobs)
        self.add_button.setEnabled(not busy)
        self.remove_button.setEnabled(bool(self.jobs) and not busy)
        self.clear_button.setEnabled(bool(self.jobs) and not busy)
        self.run_button.setEnabled(bool(self.jobs) and not busy)
        self.save_button.setEnabled(ready_count > 0 and not busy)

    def start_processing(self) -> None:
        if not self.jobs:
            return
        self._reset_work_dir()
        for job in self.jobs:
            path = job["path"]
            job.clear()
            job.update({"path": path, "status": "pending"})
        self.unsaved_results = False
        self._process_seconds = 0.0
        self._save_seconds = 0.0
        self.progress.setValue(0)
        self.status_label.setText("กำลังเปิด Microsoft Excel แบบซ่อนหน้าต่าง…")
        self._start_worker("process")

    def confirm_save(self) -> None:
        ready_count = sum(job.get("status") == "ready" for job in self.jobs)
        if ready_count == 0:
            return
        spent = (
            f"ประมวลผลเสร็จใน {format_duration(self._process_seconds)}\n\n"
            if self._process_seconds
            else ""
        )
        answer = QMessageBox.question(
            self,
            "บันทึกทับไฟล์เดิม",
            f"{spent}"
            f"บันทึกผลลัพธ์ {ready_count} ไฟล์กลับตำแหน่งเดิมหรือไม่?\n\n"
            "โปรแกรมจะสร้างโฟลเดอร์ Backup ไว้ข้างไฟล์ต้นฉบับก่อนทุกครั้ง",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer == QMessageBox.StandardButton.Yes:
            self.status_label.setText("กำลังสร้าง Backup และบันทึกไฟล์…")
            self.progress.setValue(0)
            self._start_worker("save")

    def _start_worker(self, mode: str) -> None:
        self._started_at = time.perf_counter()
        self.thread = QThread(self)
        self.worker = BatchWorker(mode, self.jobs, self.work_dir)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.failed.connect(self.worker.deleteLater)
        self.thread.finished.connect(self._worker_cleanup)
        self.thread.start()
        self._update_controls()

    def on_progress(self, count: float, path: str, status: str, detail: str) -> None:
        status_changed = False
        for job in self.jobs:
            if Path(job["path"]) == Path(path):
                status_changed = job.get("status") != status
                job["status"] = status
                if status == "error":
                    job["error"] = detail
                break
        total = max(1, len(self.jobs) if status not in {"saving", "saved"} else sum(job.get("status") in {"ready", "saving", "saved"} for job in self.jobs))
        self.progress.setValue(min(100, round((count / total) * 100)))
        self.status_label.setText(f"{Path(path).name}: {detail}")
        # Per-sheet updates arrive many times per file; the table only changes
        # when a file moves between states, so rebuilding it every tick is waste.
        if status_changed:
            self._refresh_table()

    def on_finished(self, mode: str) -> None:
        elapsed = time.perf_counter() - self._started_at if self._started_at else 0.0
        if mode == "process":
            self._process_seconds = elapsed
            ready = sum(job.get("status") == "ready" for job in self.jobs)
            errors = sum(job.get("status") == "error" for job in self.jobs)
            self.unsaved_results = ready > 0
            self._prompt_save_when_idle = ready > 0
            self.status_label.setText(
                f"ประมวลผลเสร็จใน {format_duration(elapsed)} · "
                f"พร้อมบันทึก {ready} ไฟล์ · ไม่สำเร็จ {errors} ไฟล์"
            )
            # With files ready, the save prompt follows immediately and carries
            # the time; a box of its own here would only add a click.
            if ready == 0:
                QMessageBox.information(
                    self,
                    "ประมวลผลเสร็จ",
                    f"ใช้เวลาทั้งหมด {format_duration(elapsed)}\n\n"
                    f"ไม่มีไฟล์ที่พร้อมบันทึก · ไม่สำเร็จ {errors} ไฟล์",
                )
        else:
            self._save_seconds = elapsed
            saved = sum(job.get("status") == "saved" for job in self.jobs)
            self.unsaved_results = any(job.get("status") == "ready" for job in self.jobs)
            total = self._process_seconds + elapsed
            self.status_label.setText(
                f"บันทึกเรียบร้อย {saved} ไฟล์ พร้อม Backup · "
                f"รวมเวลาทั้งหมด {format_duration(total)}"
            )
            QMessageBox.information(
                self,
                "เสร็จเรียบร้อย",
                f"บันทึกทับไฟล์เดิมแล้ว {saved} ไฟล์\n"
                f"สร้าง Backup ไว้ข้างไฟล์ต้นฉบับแล้ว\n\n"
                f"เวลาประมวลผล: {format_duration(self._process_seconds)}\n"
                f"เวลาบันทึก: {format_duration(elapsed)}\n"
                f"รวมเวลาทั้งหมด: {format_duration(total)}",
            )
        self.progress.setValue(100)
        self._refresh_table()

    def on_failed(self, message: str) -> None:
        self.status_label.setText("เกิดข้อผิดพลาด")
        QMessageBox.critical(self, "ไม่สามารถทำรายการได้", message)

    def _worker_cleanup(self) -> None:
        thread = self.thread
        self.worker = None
        self.thread = None
        if thread is not None:
            thread.deleteLater()
        self._update_controls()
        if self._prompt_save_when_idle:
            self._prompt_save_when_idle = False
            self.showNormal()
            self.raise_()
            self.activateWindow()
            QApplication.alert(self)
            QTimer.singleShot(0, self.confirm_save)

    def closeEvent(self, event) -> None:
        if self.thread is not None and self.thread.isRunning():
            QMessageBox.warning(self, "กำลังทำงาน", "กรุณารอให้การทำงานเสร็จก่อนปิดโปรแกรม")
            event.ignore()
            return
        if self.unsaved_results:
            answer = QMessageBox.question(
                self,
                "ปิดโดยยังไม่บันทึก?",
                "ผลลัพธ์ชั่วคราวจะถูกลบ แต่ไฟล์ต้นฉบับยังไม่เปลี่ยน",
            )
            if answer != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        shutil.rmtree(self.work_dir, ignore_errors=True)
        event.accept()


# Embedded application icon: no assets folder is needed.
_APP_ICON_BASE64 = (
    "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"
    "waf/Povch/782mJ02QULwlsfC+0te5ooVxtoNFtM9GazDbfdQuAT0ZWYV8RNBjFCIKQU9I9VARNDWfhmNDUBtTSYd0dKitNAdxJD"
    "3a4hvp7ZmgFk+jxaYhiGS2a1Ple8s2aSlBpSf5Llp1yL5KwiOV9ihaZ+VgYl70fMYFQNqwcHhWIehXwByxaNYM3i/mDHnvHs5q0H"
    "PCsjb4hufeVXT5YJnicDKKvU+uZbQvn+zV/szve983Pv7HN7Bvqcbz7iif3jdcxW6mi32mg0mmg2mghpxtPgsitEg6R1Okt5I+pT"
    "KoAG1YhvTSzYIPeANj1jjCLW+t3o9vTTpVWB5gimQ0w0aEbQfqeWGmwNsCWoGEI7f0r2a8s/ZsyUflG3oG0G/eqw2LR9kIy+shWY"
    "yfTvbIGa74QFx7FRLBQwONCPc85YEo6NTVqPbT5gW5b3vujbr/unk2GC58cAN97riJuvDKz3P/E3WavrN//uXUNuqac/c/sjLRye"
    "rKJSq6NRb6FZb8BzXTaiiPBk4cswUATXM4ktamWFxaRTQ6FnVupGJUK20NTXeoA7hUPMNGb2kirhY8PUzFQ7It7imZmIfXM+dcHk"
    "Gry7uedYVSTCQd2cPoCImGaUWJrpg2JpkBiYxGACNt83Mwudw7IRCYGM46Cnuxvnrl8dladm8OTmXY6dk9eF37n2m8/XMPz5GeA6"
    "aeMWEeY+uPkDbjv3+U+8acBdvHhB5o5Hm5goVzEzV0O91kCj1oDvepBMfEV0KZWxxyMYi1UlCdJkNgaT8s2U3o8HUCEremCVXtU+"
    "m7b2tZ9tKKZtAEOzFAeh87pKdcRMyCaHJmZi8icMwnIwdV/HA4djMa9Fe1onxUSft5+RbOl9LBvCtpWbKwS6ikWct2FdNDk+IXZt"
    "3+dnSvJS/zvXbTZu+C+OAfQF+v/7U+tnj7R+/KsvHcxedeEKcedPGmJsck7OVWqiXmuiNldF6Ptq5kdK5/OsJzIwIyjjiAmh3buU"
    "KZ8S69o5jwmeADza2Y/Fs2IozVwpHdxBI0MgYwtAQXV8Xx3WoNpP79GhatR5jN0/D9TucEnSBupx3g1xza58HBE/bSdo8JHcTO06"
    "CtvhISG74MILzwqfeXJnZvLgge2D6wYunal5zZ8XJ3BOmPjmMaUU4voff27V4p7iZecs9X+wpW0dmZgDEb9WqaM+W0EY+IgCT1n4"
    "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"
)


def _application_icon() -> QIcon:
    import base64

    pixmap = QPixmap()
    if not pixmap.loadFromData(base64.b64decode(_APP_ICON_BASE64), "ICO"):
        raise RuntimeError("Cannot load the embedded application icon")
    return QIcon(pixmap)


def main() -> int:
    if sys.platform == "win32":
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("DeleteTotalNA.Batch.1")
    app = QApplication(sys.argv)
    app.setWindowIcon(_application_icon())
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLE)
    window = MainWindow()
    window.show()
    return app.exec()


#if __name__ == "__main__":
    raise SystemExit(main())


 
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
        raise SystemExit(main())

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
