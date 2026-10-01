r"""Auto Lychee — the whole program in ONE readable Python file.

    python AutoLychee_OneFile.py              start Auto Lychee (asks to pip-install missing packages)
    python AutoLychee_OneFile.py --install    install / update the required packages
    python AutoLychee_OneFile.py --check      load every section and report (for developers)
    python AutoLychee_OneFile.py --build-exe  build a single-file exe from this file (PyInstaller)

HOW THIS FILE IS ORGANISED
    1. This loader (up to "END OF LOADER").
    2. One section per original module, in this order:
           core, fast_styles, history, chrome, driver, sounds, post_total_na, post_del_sig,
           post_cut_percent, worker, app, onefile_assets
       Each section starts with a line of the form  "# ====== MODULE: <name> ======"  and is the
       unchanged source of <name>.py from the multi-file project, except where it is marked
       "ONEFILE:" (paths, how the worker is started, the start-up block of app).
    3. onefile_assets: the icons (SVG as text, ICO/PNG as base64).

    The loader runs every section as its own module (its own globals), exactly like separate
    .py files: `import worker`, `from core import Job` … keep working, names never clash
    (post_total_na and post_cut_percent both define process_workbook), and modules are only
    loaded when used (the GUI never loads the Lyche driver). Tracebacks show real line numbers
    of this file. Edit a section like a normal module; do not remove the section header lines.

PROCESSES (same as the project)
    GUI            python AutoLychee_OneFile.py
    worker         python AutoLychee_OneFile.py --worker request.json   (started by the GUI)
    post child     python AutoLychee_OneFile.py --post                  (started by the worker)

DATA
    Queue, settings, logs and the unpacked icons: %LOCALAPPDATA%\AutoLychee\OneFile
    (override with the AUTOLYCHEE_DATA environment variable). Separate from the project's data\
    folder and from Auto Lychee.exe.

NEEDS
    Windows, Lyche-Epoch, Python with: PySide6, pywinauto==0.6.9, comtypes, pywin32, openpyxl,
    Pillow, lxml (optional, faster Excel files). Microsoft Excel for Step 1 (Delete Total + NA).
"""
from __future__ import annotations

import __future__ as _future
import base64 as _base64
import importlib.abc as _abc
import importlib.util as _util
import os as _os
import re as _re
import subprocess as _subprocess
import sys as _sys
from pathlib import Path as _Path

FROZEN = getattr(_sys, 'frozen', False)  # running as an exe built with --build-exe
SINGLE_FILE = _Path(__file__).resolve()
if FROZEN and not SINGLE_FILE.exists():  # the exe carries this file as data (see --build-exe)
    SINGLE_FILE = _Path(getattr(_sys, '_MEIPASS', '.')) / SINGLE_FILE.name
DATA = _Path(_os.environ.get('AUTOLYCHEE_DATA')
             or _Path(_os.environ.get('LOCALAPPDATA', _Path.home())) / 'AutoLychee' / 'OneFile')
ASSETS = DATA / 'assets'
PACKAGES = {'PySide6': 'PySide6', 'pywinauto': 'pywinauto==0.6.9', 'comtypes': 'comtypes',
            'win32api': 'pywin32', 'openpyxl': 'openpyxl', 'PIL': 'Pillow', 'lxml': 'lxml'}
OPTIONAL = {'lxml'}

_HEADER = _re.compile(r'^# ====== MODULE: ([A-Za-z_]\w*) ======$', _re.M)
_END_OF_MODULES = '\n# ====== END OF MODULES ======\n'
_FUTURE_NOTE = '\n# from __future__ import annotations  (applied to this section by the loader)\n'
_sections: dict[str, tuple[str, int]] = {}


def _read_sections() -> dict[str, tuple[str, int]]:
    """{module name: (source, first line number)} from the section headers of this file."""
    if not _sections:
        text = SINGLE_FILE.read_text(encoding='utf-8')
        end = text.index(_END_OF_MODULES)
        headers = list(_HEADER.finditer(text, 0, end))
        for index, header in enumerate(headers):
            start = header.end() + 1
            stop = headers[index + 1].start() if index + 1 < len(headers) else end + 1
            _sections[header.group(1)] = (text[start:stop], text.count('\n', 0, start) + 1)
    return _sections


def worker_command(request: str) -> tuple[str, list[str]]:
    """How the GUI starts the background worker (program, arguments)."""
    if FROZEN:
        return _sys.executable, ['--worker', request]
    return _sys.executable, ['-X', 'utf8', '-u', str(SINGLE_FILE), '--worker', request]


class _SectionImporter(_abc.MetaPathFinder, _abc.Loader):
    """Serves `import <section name>` from this file; sections run as ordinary modules."""
    def find_spec(self, fullname, path=None, target=None):
        if fullname in _read_sections() and fullname != 'app':
            return _util.spec_from_loader(fullname, self, origin=str(SINGLE_FILE))
        return None

    def create_module(self, spec):
        return None

    def exec_module(self, module):
        module.__file__ = str(SINGLE_FILE)
        exec(_compile(module.__name__), module.__dict__)


def _compile(name: str):
    source, line = _read_sections()[name]
    flags = _future.annotations.compiler_flag if _FUTURE_NOTE in '\n' + source else 0
    return compile('\n' * (line - 1) + source, str(SINGLE_FILE), 'exec', flags=flags, dont_inherit=True)


def _bind_stdio(stdin: bool = False) -> None:
    """A windowed exe may start without console streams: bind them to the pipes we were given."""
    import ctypes
    import msvcrt
    if _sys.stdout is None:
        handle = ctypes.windll.kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
        _sys.stdout = open(msvcrt.open_osfhandle(handle, _os.O_WRONLY), 'w', encoding='utf-8', buffering=1)
    else:
        _sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
    if stdin:
        if _sys.stdin is None:
            handle = ctypes.windll.kernel32.GetStdHandle(-10)  # STD_INPUT_HANDLE
            _sys.stdin = open(msvcrt.open_osfhandle(handle, _os.O_RDONLY), 'r', encoding='utf-8')
        else:
            _sys.stdin.reconfigure(encoding='utf-8')
    if _sys.stderr is None:
        _sys.stderr = _sys.stdout


def _write_assets() -> None:
    """Icons from the onefile_assets section to DATA/assets (only when they changed)."""
    import onefile_assets
    ASSETS.mkdir(parents=True, exist_ok=True)
    files = {name: text.encode('utf-8') for name, text in onefile_assets.TEXT.items()}
    files.update({name: _base64.b64decode(data) for name, data in onefile_assets.BINARY.items()})
    for name, content in files.items():
        target = ASSETS / name
        if not target.exists() or target.read_bytes() != content:
            target.write_bytes(content)


def _missing() -> list[str]:
    return [package for module, package in PACKAGES.items() if _util.find_spec(module) is None]


def _ask(text: str, question: bool = False) -> bool:
    print(text)
    try:
        import ctypes
        flags = 0x24 if question else 0x40  # YESNO|QUESTION or INFORMATION
        return ctypes.windll.user32.MessageBoxW(None, text, 'Auto Lychee', flags) in (1, 6)
    except Exception:
        return not question


def _pip(packages: list[str]) -> int:
    command = [_sys.executable, '-m', 'pip', 'install', *packages]
    print('>', ' '.join(command))
    console = _sys.stdout is None or 'pythonw' in _sys.executable.lower()
    return _subprocess.call(command, creationflags=_subprocess.CREATE_NEW_CONSOLE if console else 0)


def _ensure_packages() -> None:
    missing = _missing()
    if FROZEN or not [p for p in missing if p not in OPTIONAL]:
        return
    if not _ask('Auto Lychee ต้องใช้แพ็กเกจ: ' + ', '.join(missing) + '\n\nติดตั้งตอนนี้เลยไหม (pip install)?', True):
        raise SystemExit(1)
    if _pip(missing) != 0 or [p for p in _missing() if p not in OPTIONAL]:
        _ask('ติดตั้งไม่สำเร็จ ลองรันเอง:\n' + _sys.executable + ' -m pip install ' + ' '.join(missing))
        raise SystemExit(1)


def _check() -> None:
    """Load every section (the GUI one too, without starting it) and report."""
    for name in _read_sections():
        if name == 'app':
            namespace = {'__name__': 'app', '__file__': str(SINGLE_FILE)}
            exec(_compile('app'), namespace)
            print(f'  ok  app ({len(namespace)} names)')
        else:
            module = __import__(name)
            print(f'  ok  {name} ({len(vars(module))} names)')
    print(f'{len(_read_sections())} sections OK · data: {DATA}')


def _build_exe() -> None:
    """dist_onefile\\Auto Lychee OneFile.exe; this file goes inside as data (the loader reads it)."""
    if _util.find_spec('PyInstaller') is None:
        raise SystemExit('ต้องติดตั้ง PyInstaller ก่อน:  ' + _sys.executable + ' -m pip install pyinstaller')
    _write_assets()
    here = SINGLE_FILE.parent
    hidden = ['pythoncom', 'pywintypes', 'win32com.client', 'win32com.client.dynamic', 'win32job',
              'win32gui', 'win32process', 'win32api', 'win32con', 'openpyxl', 'lxml.etree',
              'pywinauto', 'comtypes', 'comtypes.client', 'PIL.ImageGrab']
    command = [_sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--onefile', '--windowed',
               '--name', 'Auto Lychee OneFile', '--icon', str(ASSETS / 'logo.ico'),
               '--add-data', f'{SINGLE_FILE}{_os.pathsep}.',
               '--distpath', str(here / 'dist_onefile'), '--workpath', str(here / 'build_onefile'),
               '--specpath', str(here / 'build_onefile')]
    for module in hidden:
        command += ['--hidden-import', module]
    command.append(str(SINGLE_FILE))
    print('>', ' '.join(command))
    raise SystemExit(_subprocess.call(command))


def _run_app() -> None:
    namespace = {'__name__': 'app', '__file__': str(SINGLE_FILE)}
    exec(_compile('app'), namespace)
    namespace['main']()


def _smoke_test() -> None:
    """Load all sections and construct the GUI without running survey jobs."""
    _check()
    _write_assets()
    namespace = {'__name__': 'app', '__file__': str(SINGLE_FILE)}
    exec(_compile('app'), namespace)
    application = namespace['QApplication']([str(SINGLE_FILE)])
    application.setStyle('Fusion')
    window = namespace['App']()
    application.processEvents()
    window.deleteLater()
    application.processEvents()
    print('Auto Lychee GUI smoke test OK', flush=True)


def _main() -> None:
    _sys.modules.setdefault('onefile', _sys.modules[__name__])  # sections read paths from here
    _sys.meta_path.insert(0, _SectionImporter())
    args = _sys.argv[1:]
    if len(args) >= 2 and args[0] == '--worker':
        _bind_stdio()
        import worker
        _sys.argv = [str(SINGLE_FILE), args[1]]
        worker.main()
    elif args == ['--post']:
        _bind_stdio(stdin=True)
        import worker
        worker.post_main()
    elif args[:1] == ['--install']:
        raise SystemExit(_pip(['--upgrade', *PACKAGES.values()]))
    elif args[:1] == ['--check']:
        if FROZEN:
            _bind_stdio()
        _check()
    elif args == ['--smoke-test']:
        if FROZEN:
            _bind_stdio()
        _smoke_test()
    elif args[:1] == ['--build-exe']:
        _build_exe()
    else:
        _ensure_packages()
        _write_assets()
        _run_app()


if __name__ == '__main__':
    _main()
    raise SystemExit(0)
raise ImportError('AutoLychee_OneFile.py is a program: run it, do not import it')
# ============================== END OF LOADER ==============================
# Everything below is loaded section by section by the loader above (never run top to bottom).

# ====== MODULE: core ======
"""Queue validation and storage, independent of the Windows UI."""
# from __future__ import annotations  (applied to this section by the loader)
import json
import re
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from uuid import uuid4


@dataclass
class Job:
    history: str = ''
    output: str = ''
    filter: str = '-'
    base: str = ''
    status: str = 'รอรัน'
    detail: str = ''
    id: str = field(default_factory=lambda: uuid4().hex)
    # Post-processing after the Banner is exported (Delete Total + NA, post_total_na.py). Off by default;
    # new fields go after `id` so positional Job(history, output, filter, base, status, detail) still works.
    total_na: bool = True  # on by default (user request); its empty-row option stays off
    total_na_empty_rows: bool = False
    # Del Sig (Del_Sig.py, post_del_sig.py) runs after Delete Total + NA. Off by default.
    del_sig: bool = False
    del_sig_groups: str = ''
    del_sig_mode: str = 'NORMAL'  # NORMAL = Crosstab ธรรมดา, MATRIX = Matrix
    del_sig_beside: bool = False  # False = Sig ปกติ, True = Sig ข้าง
    # Cut N / % (151_CutLychee_Persence.py, post_cut_percent.py) runs after Del Sig. Off by default.
    cut_percent: bool = False
    cut_percent_mode: str = 'both'  # both = N + % (two files), count = N Only, percent = % Only
    # The last Step (always last, user rule): how Lyche exports. sheets = separated sheets (Cross: Export
    # with Analysis Axis; Matrix: Export all → Excel(Separated Sheets)); onesheet = Export all → Excel(One Sheet).
    export_mode: str = 'sheets'
    # Banner Manual (user request): after loading the History, replace its Banner with these Lyche item
    # codes, in order (Clear all → Yes, then search each item → To Banner). The Stub stays the History's.
    banner_manual: bool = False
    banner_manual_items: str = ''  # e.g. 'QUOTA1, QUOTA6' (commas, spaces or new lines between items)


def parse_filter(value: str):
    value = value.strip()
    if value in ('', '-'):
        return None
    match = re.fullmatch(r'([A-Za-z_][A-Za-z0-9_]*)\s*(\^=|!=|<>|=)\s*(-?\d+)', value)
    if not match:
        raise ValueError('Filter ต้องเป็น - หรือ ตัวแปร = code / ตัวแปร ^= code เช่น QUOTA6 ^= 1')
    var, operator, code = match.groups()
    return var, 'Include' if operator == '=' else 'Exclude', str(int(code))


def manual_items(text: str) -> list[str]:
    """Banner Manual items as a list: split on commas / semicolons / white space, duplicates dropped."""
    items, seen = [], set()
    for item in re.split(r'[,;\s]+', text or ''):
        if item and item.casefold() not in seen:
            seen.add(item.casefold())
            items.append(item)
    return items


def output_name(value: str) -> str:
    value = value.strip()
    if value.lower().endswith('.xlsx'):
        value = value[:-5]
    if not value or re.search(r'[<>:"/\\|?*\x00-\x1f]', value) or value.endswith((' ', '.')):
        raise ValueError('ชื่อไฟล์ว่างหรือมีอักขระที่ Windows ไม่รองรับ')
    if re.match(r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)', value, re.I):
        raise ValueError('ชื่อนี้สงวนไว้สำหรับ Windows')
    return value + '.xlsx'


def validate_jobs(jobs: list[Job], folder: Path, check_files=True):
    if not jobs:
        raise ValueError('เพิ่มรายการรันอย่างน้อยหนึ่งแถว')
    if not folder.is_dir():
        raise ValueError('เลือกโฟลเดอร์บันทึกที่มีอยู่จริง')
    seen = set()
    pending = 0
    for index, job in enumerate(jobs, 1):
        try:
            if not job.history.strip():
                raise ValueError('ยังไม่ได้ระบุ Banner จาก History')
            name = output_name(job.output)
            if name.casefold() in seen:
                raise ValueError('ชื่อไฟล์ซ้ำกับแถวก่อนหน้า')
            seen.add(name.casefold())
            parse_filter(job.filter)
            if job.base and (not job.base.isdigit() or int(job.base) < 1):
                raise ValueError('Base ต้องเป็นจำนวนเต็มบวก หรือเว้นว่าง')
            if job.del_sig and not re.search(r'[A-Za-z]', job.del_sig_groups or ''):
                raise ValueError('เปิด Del Sig แล้วแต่ยังไม่ได้ใส่กลุ่ม Sig (ดับเบิลคลิกแถวเพื่อตั้งค่า)')
            if job.banner_manual and not manual_items(job.banner_manual_items):
                raise ValueError('เปิด Banner Manual แล้วแต่ยังไม่ได้ใส่ข้อ (คลิกช่อง Banner Manual เพื่อตั้งค่า)')
            if job.status != 'OK':
                pending += 1
                if check_files and (folder / name).exists():
                    raise ValueError('มีไฟล์ผลลัพธ์อยู่แล้ว กรุณาเปลี่ยนชื่อไฟล์หรือตรวจไฟล์เดิม')
        except ValueError as exc:
            raise ValueError(f'แถว {index}: {exc}') from exc
    if not pending:
        raise ValueError('ทุกแถวเป็น OK แล้ว')


def verify_xlsx(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        if 'xl/workbook.xml' not in archive.namelist() or archive.testzip() is not None:
            raise ValueError('ไฟล์ Excel ไม่สมบูรณ์')
        sheets = sum(bool(re.fullmatch(r'xl/worksheets/sheet\d+\.xml', n)) for n in archive.namelist())
        if not sheets:
            raise ValueError('ไม่พบชีทในไฟล์ผลลัพธ์')
        return sheets


def save_json(path: Path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + '.tmp')
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')
    temporary.replace(path)


ON, OFF = 'เปิด', 'ปิด'
CUT_MODE_LABELS = {'both': 'N + %', 'count': 'N Only', 'percent': '% Only'}
EXPORT_MODE_LABELS = {'sheets': 'แยกชีท', 'onesheet': 'One Sheet'}
# Queue export/import columns (header text, Job field). Import finds columns by header text, so
# older exports (Banner, ชื่อไฟล์ผลลัพธ์, Filter, Base, สถานะ, รายละเอียด) still load.
QUEUE_COLUMNS = (
    ('Banner', 'history'),
    ('Banner Manual', 'banner_manual'),
    ('Banner Manual ข้อ', 'banner_manual_items'),
    ('ชื่อไฟล์ผลลัพธ์', 'output'),
    ('Filter', 'filter'),
    ('Base', 'base'),
    ('Step 1 Delete Total + NA', 'total_na'),
    ('Step 1 ลบแถวว่าง', 'total_na_empty_rows'),
    ('Step 2 Del Sig', 'del_sig'),
    ('Step 2 กลุ่ม Sig', 'del_sig_groups'),
    ('Step 2 ประเภท Crosstab', 'del_sig_mode'),
    ('Step 2 รูปแบบ Sig', 'del_sig_beside'),
    ('Step 3 ตัด N / %', 'cut_percent'),
    ('Step 3 ผลลัพธ์', 'cut_percent_mode'),
    ('Step 4 Export', 'export_mode'),
    ('สถานะ', None),
    ('รายละเอียด', None),
)
_HEADER_ALIASES = {'history': 'history', 'banner จาก history': 'history', 'output': 'output'}


def setting_cell(job: Job, field: str):
    """A Job setting as the text written to the exported queue."""
    value = getattr(job, field)
    if field == 'del_sig_mode':
        return 'Matrix' if value == 'MATRIX' else 'Crosstab ธรรมดา'
    if field == 'del_sig_beside':
        return 'Sig ข้าง' if value else 'Sig ปกติ'
    if field == 'cut_percent_mode':
        return CUT_MODE_LABELS.get(value, 'N + %')
    if field == 'export_mode':
        return EXPORT_MODE_LABELS.get(value, 'แยกชีท')
    if isinstance(value, bool):
        return ON if value else OFF
    return value


def _setting_value(field: str, text: str):
    """Exported (or hand-typed) text back to the Job value; '' keeps the Job default."""
    if not text:
        return getattr(Job(), field)
    folded = text.strip().casefold()
    if field == 'del_sig_mode':
        return 'MATRIX' if 'matrix' in folded else 'NORMAL'
    if field == 'del_sig_beside':
        return 'ข้าง' in folded or 'beside' in folded
    if field == 'cut_percent_mode':
        compact = folded.replace(' ', '')
        if compact in CUT_MODE_LABELS:
            return compact
        if 'only' in compact or compact in ('n', '%'):
            return 'percent' if compact.startswith('%') else 'count'
        return 'both'
    if field == 'export_mode':
        return 'onesheet' if 'one' in folded.replace(' ', '') else 'sheets'
    if field == 'del_sig_groups':
        return text.strip().upper()
    if field == 'banner_manual_items':
        return ', '.join(manual_items(text))
    if isinstance(getattr(Job(), field), bool):
        return folded in (ON, 'on', 'true', '1', 'yes', 'y', '✓', 'x')
    return text


def _cell_text(value) -> str:
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():  # Excel stores 280 as 280.0
        value = int(value)
    return str(value).strip()


def read_jobs(rows, default_history='') -> list[Job]:
    """TSV/Excel/CSV rows. With a header row (as written by the queue export) columns are found
    by name, including the Step 1-3 settings; without one: history, output, filter, optional
    base (or just output, filter)."""
    jobs = []
    fields = None
    by_header = {header.casefold(): field for header, field in QUEUE_COLUMNS}
    by_header.update(_HEADER_ALIASES)
    for row in rows:
        values = [_cell_text(v) for v in row]
        if not any(values):
            continue
        if fields is None and values[0].casefold() in ('history', 'banner จาก history', 'banner'):
            fields = [by_header.get(value.casefold()) for value in values]
            continue
        if fields is not None:
            data = {field: value for field, value in zip(fields, values) if field}
            job = Job(history=data.get('history') or default_history, output=data.get('output', ''),
                      filter=data.get('filter') or '-', base=data.get('base', ''))
            for field, value in data.items():
                if field not in ('history', 'output', 'filter', 'base'):
                    setattr(job, field, _setting_value(field, value))
            if not job.total_na:
                job.total_na_empty_rows = False
            jobs.append(job)
            continue
        if len(values) == 2:
            values = [default_history, *values]
        values += [''] * max(0, 4 - len(values))
        jobs.append(Job(history=values[0] or default_history, output=values[1], filter=values[2] or '-', base=values[3]))
    return jobs


# ====== MODULE: fast_styles ======
"""Exact, faster style lookup for openpyxl (used by Del Sig and Cut N / %).

Profiling a ~200-sheet Lyche export: ~80% of openpyxl.load_workbook (and most of the Cut N / %
work) is `IndexedList.add()` on the workbook's style lists. Every border assigned to a merged
cell is hashed and deep-compared against the stored styles through openpyxl's generic
Serialisable.__hash__/__eq__ (recursive, string-converting).

`exact_style_cache()` memoises `add()` per list with a structural key built from the raw
descriptor values. Equal keys imply the objects are equal for openpyxl too (same class,
same raw values, recursively), so a memo hit returns the index `add()` itself would return
for that value. A miss simply calls the original `add()`. The memo is dropped whenever
openpyxl rebuilds the list's dict. Output files are byte-identical (verified against the
unpatched code, see AI_HANDOFF.md)."""
# from __future__ import annotations  (applied to this section by the loader)

import threading
from contextlib import contextmanager

from openpyxl.descriptors import Descriptor
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.utils.indexed_list import IndexedList

_MISSING = object()
_plain: dict[type, bool] = {}
_lock = threading.Lock()
_depth = 0
_original_add = IndexedList.add


_SCALARS = frozenset({str, int, float, bool, type(None), object})  # object = _MISSING


def _layout(cls):
    """(attributes, elements) for classes compared by openpyxl's generic Serialisable rules
    over plain descriptors (value stored in the instance __dict__); None for anything else."""
    try:
        return _plain[cls]
    except KeyError:
        ok = (cls.__eq__ is Serialisable.__eq__ and cls.__hash__ is Serialisable.__hash__
              and all(isinstance(getattr(cls, name, None), Descriptor)
                      and not hasattr(type(getattr(cls, name)), '__get__')
                      for name in cls.__attrs__ + cls.__elements__))
        layout = _plain[cls] = (tuple(cls.__attrs__), tuple(cls.__elements__)) if ok else None
        return layout


def _key(obj):
    """Structural key from raw values. Types are part of it: 1, 1.0 and True compare equal but
    may serialise differently."""
    cls = obj.__class__
    layout = _plain.get(cls, _MISSING)
    if layout is _MISSING:
        layout = _layout(cls)
    if layout is None:
        raise TypeError(cls)
    get = obj.__dict__.get
    parts = [cls]
    for name in layout[0]:
        value = get(name, _MISSING)
        kind = value.__class__
        if kind not in _SCALARS and isinstance(value, (Serialisable, list, dict)):
            raise TypeError(cls)  # attributes are expected to be plain values
        parts.append(kind)
        parts.append(value)
    for name in layout[1]:
        value = get(name, _MISSING)
        kind = value.__class__
        if kind in _SCALARS:
            parts.append((kind, value))
        elif isinstance(value, Serialisable):
            parts.append(_key(value))
        elif isinstance(value, (list, tuple)):
            parts.append((kind,) + tuple(_key(v) if isinstance(v, Serialisable) else (v.__class__, v) for v in value))
        else:
            parts.append((kind, value))
    return tuple(parts)


def _add(self, value):
    if not isinstance(value, Serialisable):
        return _original_add(self, value)
    try:
        key = _key(value)
        hash(key)
    except TypeError:
        return _original_add(self, value)
    state = self.__dict__.get('_exact_memo')
    if state is None or state[0] is not self._dict:
        state = (self._dict, {})
        self.__dict__['_exact_memo'] = state
    index = state[1].get(key)
    if index is not None:
        return index
    index = _original_add(self, value)
    if state[0] is self._dict:
        state[1][key] = index
    return index


_original_copy = Serialisable.__copy__
_copies: dict = {}


def _clone(value):
    """A fresh object with exactly the same state (nested style objects and lists cloned too)."""
    if isinstance(value, Serialisable):
        new = object.__new__(type(value))
        new.__dict__.update({name: _clone(v) for name, v in value.__dict__.items()})
        return new
    if type(value) is list:
        return [_clone(v) for v in value]
    if type(value) is tuple:
        return tuple(_clone(v) for v in value)
    return value


def _copy(self):
    """openpyxl copies a style object by writing it to XML and parsing it back. The result
    depends only on the object's class and raw values, so the first copy of each distinct value
    goes through openpyxl itself and later copies clone that (never handed out) prototype."""
    cls = type(self)
    try:
        key = _key(self)
        # openpyxl's copy also carries over non-XML state (e.g. Border.diagonal_direction)
        layout = _plain[cls]
        extras = tuple(sorted(((name, value.__class__, value) for name, value in self.__dict__.items()
                               if name not in layout[0] and name not in layout[1]), key=lambda item: item[0]))
        if any(kind not in _SCALARS for _, kind, _ in extras):
            raise TypeError(cls)  # non-scalar extra state: let openpyxl copy it
        key = (key, extras)
        hash(key)
    except TypeError:
        return _original_copy(self)
    prototype = _copies.get(key)
    if prototype is None:
        prototype = _original_copy(self)
        _copies[key] = prototype
    return _clone(prototype)


@contextmanager
def exact_style_cache():
    """Patch IndexedList.add and style copies while openpyxl loads/processes/saves; restores
    both (and frees the copy memo) afterwards."""
    global _depth
    with _lock:
        _depth += 1
        IndexedList.add = _add
        Serialisable.__copy__ = _copy
    try:
        yield
    finally:
        with _lock:
            _depth -= 1
            if not _depth:
                IndexedList.add = _original_add
                Serialisable.__copy__ = _original_copy
                _copies.clear()


# ====== MODULE: history ======
"""Read saved Personal History metadata only; no UI actions or Lyche writes."""
# from __future__ import annotations  (applied to this section by the loader)
import os
import re
from pathlib import Path
import xml.etree.ElementTree as ET


def project_code(title: str) -> str:
    # Also the project window itself ('Lyche-Epoch <code:name>'): the app no longer opens
    # Cross Tabulation at start-up just to read Personal History.
    match = re.match(r'^(?:(?:Cross Tabulation|Tabulation Result) - )?Lyche-Epoch\s+<([A-Za-z0-9_-]+):[^>]+>', title)
    if not match:
        raise ValueError('อ่านรหัสโปรเจกต์จากหน้าต่าง Lyche ไม่ได้')
    return match.group(1)


def read_personal_history(title: str, source='Personal', root: Path | None = None):
    if source != 'Personal':
        raise ValueError('การดึงเบื้องหลังรองรับ Personal History ในเครื่องนี้เท่านั้น ยังไม่มีแหล่งข้อมูล Shared ที่ยืนยันได้')
    if root is None:
        appdata = os.environ.get('APPDATA')
        if not appdata:
            raise ValueError('ไม่พบโฟลเดอร์ AppData ของผู้ใช้')
        root = Path(appdata) / 'Intage' / 'LycheEpoch' / 'result'
    path = root / project_code(title) / 'PersonalTabulationList.xml'
    if not path.is_file():
        raise FileNotFoundError('ยังไม่พบ Personal History ที่บันทึกไว้ของโปรเจกต์นี้ กรุณาบันทึก History ใน Lyche แล้วกด Get Banner ใหม่')
    try:
        tree = ET.parse(path)
    except ET.ParseError as exc:
        raise ValueError('ไฟล์ History กำลังเปลี่ยนแปลงหรืออ่านไม่ได้ กรุณากด Get Banner อีกครั้ง') from exc
    entries = []
    for index, info in enumerate(tree.getroot().iter()):
        if info.tag.rsplit('}', 1)[-1] != 'tabulation_info':
            continue
        fields = {child.tag.rsplit('}', 1)[-1]: (child.text or '').strip() for child in info}
        name = fields.get('tabulation_name', '')
        code = fields.get('tabulation_cd', '')
        if not name or not code:
            continue
        order = fields.get('display_order', '')
        entries.append((int(order) if order.isdigit() else index, index, name))
    names = list(dict.fromkeys(entry[2] for entry in sorted(entries)))
    return names, path


def from_window(handle: int, source='Personal', last_title: str | None = None):
    """`last_title` (the title shown in the app) is used when the window has since been closed —
    the app closes Cross Tabulation after each run — so Personal History still reads in the background."""
    import win32gui
    if handle and win32gui.IsWindow(handle):
        title = win32gui.GetWindowText(handle)
    elif last_title:
        title = last_title
    else:
        raise ValueError('หน้าต่าง Lyche ปิดแล้ว กรุณาค้นหาหน้าต่างใหม่')
    if not re.match(r'^((Cross Tabulation|Tabulation for Matr?ix settings|Tabulation Result) - )?Lyche-Epoch <', title):
        raise ValueError('หน้าต่างที่เลือกไม่ใช่หน้าต่างโปรเจกต์ของ Lyche')
    return read_personal_history(title, source)


# ====== MODULE: chrome ======
"""macOS-style window chrome for a frameless Qt window on Windows.

Traffic-light buttons are Qt widgets; dragging, Aero Snap, double-click maximize,
edge resizing, the drop shadow and Windows 11 rounded corners stay native by
keeping WS_THICKFRAME/WS_CAPTION and answering WM_NCCALCSIZE / WM_NCHITTEST.
"""
# from __future__ import annotations  (applied to this section by the loader)
import ctypes
import sys
from ctypes import wintypes

from PySide6.QtCore import Qt, QEvent
from PySide6.QtWidgets import QHBoxLayout, QLabel, QPushButton, QWidget

WIN = sys.platform == 'win32'
WM_NCCALCSIZE, WM_NCHITTEST = 0x0083, 0x0084
HTCLIENT, HTCAPTION = 1, 2
HTLEFT, HTRIGHT, HTTOP, HTTOPLEFT, HTTOPRIGHT, HTBOTTOM, HTBOTTOMLEFT, HTBOTTOMRIGHT = 10, 11, 12, 13, 14, 15, 16, 17
TITLE_HEIGHT = 40


class TrafficLights(QWidget):
    """Close / minimise / zoom dots; glyphs appear on hover and dots gray out when inactive."""

    def __init__(self, window, parent=None):
        super().__init__(parent)
        self.window_ = window
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        self.buttons = []
        # Right-aligned in Windows order so close stays in the far-right corner.
        for name, glyph, action, tip in (
                ('tl_min', '−', window.showMinimized, 'ย่อ'),
                ('tl_max', '+', self.toggle_maximized, 'ขยาย / คืนขนาด'),
                ('tl_close', '×', window.close, 'ปิด')):
            button = QPushButton('')
            button.setObjectName(name)
            button.setToolTip(tip)
            button.setFocusPolicy(Qt.NoFocus)
            button.setCursor(Qt.ArrowCursor)
            button.glyph = glyph
            button.clicked.connect(action)
            layout.addWidget(button)
            self.buttons.append(button)

    def toggle_maximized(self):
        self.window_.showNormal() if self.window_.isMaximized() else self.window_.showMaximized()

    def enterEvent(self, event):
        for button in self.buttons:
            button.setText(button.glyph)
        super().enterEvent(event)

    def leaveEvent(self, event):
        for button in self.buttons:
            button.setText('')
        super().leaveEvent(event)

    def set_active(self, active):
        for button in self.buttons:
            button.setProperty('inactive', not active)
            button.style().unpolish(button)
            button.style().polish(button)


class TitleBar(QWidget):
    def __init__(self, window, title, parent=None):
        super().__init__(parent)
        self.setFixedHeight(TITLE_HEIGHT)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 0, 16, 0)
        self.lights = TrafficLights(window)
        spacer = QWidget()
        spacer.setFixedWidth(self.lights.sizeHint().width())  # keeps the title centred
        layout.addWidget(spacer)
        layout.addStretch(1)
        self.icon = QLabel()
        self.icon.setFixedSize(18, 18)
        self.icon.setScaledContents(True)
        pixmap = window.windowIcon().pixmap(36, 36)
        if not pixmap.isNull():
            self.icon.setPixmap(pixmap)
        else:
            self.icon.hide()
        layout.addWidget(self.icon, 0, Qt.AlignVCenter)
        layout.addSpacing(6)
        self.label = QLabel(title)
        self.label.setObjectName('windowTitle')
        layout.addWidget(self.label, 0, Qt.AlignCenter)
        layout.addSpacing(24)  # balances the icon so the title stays centred
        layout.addStretch(1)
        layout.addWidget(self.lights, 0, Qt.AlignVCenter)
        self.spacer = spacer


class _MARGINS(ctypes.Structure):
    _fields_ = [('left', ctypes.c_int), ('right', ctypes.c_int), ('top', ctypes.c_int), ('bottom', ctypes.c_int)]


class MacWindowMixin:
    """Mix into a QMainWindow *before* QMainWindow. Call setup_chrome() in __init__."""
    RESIZE_BORDER = 6

    def setup_chrome(self, title):
        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint | Qt.Window)
        self.title_bar = TitleBar(self, title)
        self._drag_widgets = {self.title_bar, self.title_bar.label, self.title_bar.spacer, self.title_bar.icon}
        self._native_ready = False
        return self.title_bar

    def add_drag_widgets(self, *widgets):
        self._drag_widgets.update(widgets)

    def showEvent(self, event):
        super().showEvent(event)
        if WIN and not self._native_ready:
            self._native_ready = True
            self._apply_native_frame()

    def changeEvent(self, event):
        if event.type() == QEvent.ActivationChange and hasattr(self, 'title_bar'):
            self.title_bar.lights.set_active(self.isActiveWindow())
        super().changeEvent(event)

    def _apply_native_frame(self):
        user32, dwm = ctypes.windll.user32, ctypes.windll.dwmapi
        hwnd = int(self.winId())
        style = user32.GetWindowLongW(hwnd, -16)
        style |= 0x00040000 | 0x00C00000 | 0x00020000 | 0x00010000 | 0x00080000  # THICKFRAME CAPTION MIN MAX SYSMENU
        user32.SetWindowLongW(hwnd, -16, style)
        dwm.DwmExtendFrameIntoClientArea(hwnd, ctypes.byref(_MARGINS(0, 0, 1, 0)))  # restores the drop shadow
        corner = ctypes.c_int(2)  # DWMWCP_ROUND (Windows 11)
        dwm.DwmSetWindowAttribute(hwnd, 33, ctypes.byref(corner), ctypes.sizeof(corner))
        border = ctypes.c_uint(0x00EAD6C9)  # COLORREF of #c9d6ea
        dwm.DwmSetWindowAttribute(hwnd, 34, ctypes.byref(border), ctypes.sizeof(border))
        user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020)  # FRAMECHANGED

    def nativeEvent(self, event_type, message):
        if not WIN or bytes(event_type) != b'windows_generic_MSG':
            return False, 0
        msg = wintypes.MSG.from_address(int(message))
        if msg.message == WM_NCCALCSIZE and msg.wParam:
            if self.isMaximized():  # a maximised window overhangs the monitor by its frame width
                user32 = ctypes.windll.user32
                dpi = user32.GetDpiForWindow(msg.hWnd)
                frame = user32.GetSystemMetricsForDpi(32, dpi) + user32.GetSystemMetricsForDpi(92, dpi)
                rect = wintypes.RECT.from_address(msg.lParam)
                rect.left += frame
                rect.top += frame
                rect.right -= frame
                rect.bottom -= frame
            return True, 0
        if msg.message == WM_NCHITTEST:
            return True, self._hit_test(msg)
        return False, 0

    def _hit_test(self, msg):
        x = ctypes.c_short(msg.lParam & 0xFFFF).value
        y = ctypes.c_short((msg.lParam >> 16) & 0xFFFF).value
        rect = wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(msg.hWnd, ctypes.byref(rect))
        ratio = self.devicePixelRatioF() or 1
        border = int(self.RESIZE_BORDER * ratio)
        if not self.isMaximized():
            left, right = x < rect.left + border, x >= rect.right - border
            top, bottom = y < rect.top + border, y >= rect.bottom - border
            if top and left:
                return HTTOPLEFT
            if top and right:
                return HTTOPRIGHT
            if bottom and left:
                return HTBOTTOMLEFT
            if bottom and right:
                return HTBOTTOMRIGHT
            if left:
                return HTLEFT
            if right:
                return HTRIGHT
            if top:
                return HTTOP
            if bottom:
                return HTBOTTOM
        local_x, local_y = int((x - rect.left) / ratio), int((y - rect.top) / ratio)
        child = self.childAt(local_x, local_y)
        if child is None or child in self._drag_widgets:
            return HTCAPTION
        return HTCLIENT


# ====== MODULE: driver ======
"""Lyche-Epoch UIA adapter. All lookups are scoped to the selected process/window.

No absolute screen coordinates. Unsupported or ambiguous controls stop the run.
This module is loaded only in the worker process, never in the GUI thread.
"""
# from __future__ import annotations  (applied to this section by the loader)
import ctypes
import os
from ctypes import wintypes
import sys
import re
import time
import warnings
from contextlib import contextmanager, nullcontext
from pathlib import Path
sys.coinit_flags = 0  # UI Automation uses MTA; STA can stall WPF providers.
warnings.filterwarnings('ignore', message='Apply externally defined coinit_flags.*', category=UserWarning)
import comtypes.client
comtypes.client.gen_dir = None  # No generated code in the system Python installation.
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.uia_element_info import UIAElementInfo
from pywinauto.uia_defines import IUIA, NoPatternInterfaceError
import win32gui
import win32process
from core import parse_filter


class Stopped(Exception):
    pass


# The settings window Lyche shows for a History: Cross Tabulation, or — after Import History loads a
# Matrix History — the SAME window renamed 'Tabulation for Matix settings' (Lyche's spelling; the
# corrected spelling is accepted too). Tabulate turns either into 'Tabulation Result'.
MATRIX_PREFIXES = ('Tabulation for Matix settings', 'Tabulation for Matrix settings')
SETTINGS_PREFIXES = ('Cross Tabulation',) + MATRIX_PREFIXES
WINDOW_PREFIXES = SETTINGS_PREFIXES + ('Tabulation Result',)
# Export All Setting (after Export all → Excel(Separated Sheets) for Matrix, or Excel(One Sheet) for the
# One Sheet export), as the user set it (2026-10-01): Item ID + banner/stub numbers on; Export Sheet =
# Table only (no Information, no Graph).
EXPORT_ALL_TICKS = {'chkNoSort': False, 'chkDeleteTotal': False, 'chkItemId': True, 'chkColumnCode': True,
                    'chkRowCode': True, 'chkOutputInfoSheet': False, 'chkOutputTableSheet': True,
                    'chkOutputGraphSheet': False}


def windows():
    result = []
    def collect(handle, _):
        title = win32gui.GetWindowText(handle)
        if win32gui.IsWindowVisible(handle) and title.startswith(WINDOW_PREFIXES) and ' - Lyche-Epoch' in title:
            result.append({'handle': handle, 'title': title})
    win32gui.EnumWindows(collect, None)
    return result


def user_foreground(pid):
    """The window the user is working in right now, unless it is a Lyche window."""
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    if not hwnd or win32process.GetWindowThreadProcessId(hwnd)[1] == pid:
        return None
    return hwnd


def give_back_foreground(hwnd):
    """Return the foreground to the user's window right after a step that needed Lyche active."""
    if hwnd and win32gui.IsWindow(hwnd) and ctypes.windll.user32.GetForegroundWindow() != hwnd:
        force_foreground(hwnd)


def force_foreground(hwnd):
    """Give an (off-screen) Lyche dialog the foreground so posted keys reach it. Windows only lets
    the foreground thread do this, so borrow its input queue with AttachThreadInput."""
    user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
    foreground = user32.GetForegroundWindow()
    if foreground == hwnd:
        return
    current = kernel32.GetCurrentThreadId()
    threads = {user32.GetWindowThreadProcessId(foreground, None), user32.GetWindowThreadProcessId(hwnd, None)} - {current, 0}
    attached = [t for t in threads if user32.AttachThreadInput(current, t, True)]
    try:
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
    finally:
        for t in attached:
            user32.AttachThreadInput(current, t, False)
    time.sleep(.1)


def launchers():
    """Lyche-Epoch project windows ('Lyche-Epoch <code:name>')."""
    result = []
    def collect(handle, _):
        title = win32gui.GetWindowText(handle)
        if win32gui.IsWindowVisible(handle) and re.match(r'^Lyche-Epoch <[^>]+>', title):
            result.append({'handle': handle, 'title': title})
    win32gui.EnumWindows(collect, None)
    return result


def open_cross_tabulation(emit, timeout=120):
    """No Cross Tabulation window yet: on each Lyche-Epoch project window select the
    Tabulation tab (pnlTabulation) then press Cross Tabulation (btnQuestionCross).
    UIA Select/Invoke first (no mouse); fall back to clicking the control itself."""
    launchers = []
    def collect(handle, _):
        title = win32gui.GetWindowText(handle)
        if win32gui.IsWindowVisible(handle) and re.match(r'^Lyche-Epoch <[^>]+>', title):
            launchers.append((handle, title))
    win32gui.EnumWindows(collect, None)
    if not launchers:
        emit('log', text='ไม่พบหน้าต่าง Lyche-Epoch ที่เปิดโปรเจกต์ไว้ — เปิดโปรเจกต์ใน Lyche ก่อน')
        return []
    api = IUIA().iuia
    def element(root, aid):
        found = root.FindFirst(4, api.CreatePropertyCondition(30011, aid))
        return UIAWrapper(UIAElementInfo(found)) if found else None
    def wait_for(fn, seconds):
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            result = fn()
            if result:
                return result
            time.sleep(.1)  # poll fast so the new window is minimised before it draws
        return None
    pid = win32process.GetWindowThreadProcessId(launchers[0][0])[1]
    # Park the new window off-screen the moment it shows (2 ms poller), then minimise it without an
    # animation and give it back its normal on-screen restore position — nothing pops up.
    with HideLycheDialogs(pid) as hider:
        hider.burst(15)
        for handle, title in launchers:
            root = api.ElementFromHandle(handle)
            tab = element(root, 'pnlTabulation')
            if tab is None:
                emit('log', text=f'ไม่พบแท็บ Tabulation ใน {title}')
                continue
            try:
                tab.iface_selection_item.Select()
            except Exception:
                tab.click_input()
            button = wait_for(lambda: (b := element(root, 'btnQuestionCross')) and b.is_enabled() and b, 10)
            if button is None:
                emit('log', text=f'ไม่พบปุ่ม Cross Tabulation ใน {title}')
                continue
            hider.burst(15)
            try:
                button.iface_invoke.Invoke()
            except Exception:
                button.click_input()
            emit('log', text=f'เปิด Tabulation → Cross Tabulation ให้ {title} (ย่อไว้เบื้องหลัง)')
        opened = wait_for(windows, timeout) or []
        for item in opened:
            handle = item['handle']
            off = ctypes.c_int(1)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(handle, 3, ctypes.byref(off), 4)  # no minimise animation
            placement = win32gui.GetWindowPlacement(handle)
            rect = hider.original.get(handle)
            normal = rect if rect and rect[0] > -15000 else placement[4]
            if normal[0] <= -15000:
                normal = (100, 100, 100 + normal[2] - normal[0], 100 + normal[3] - normal[1])
            win32gui.SetWindowPlacement(handle, (placement[0], 7, placement[2], placement[3], normal))  # SW_SHOWMINNOACTIVE
            on = ctypes.c_int(0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(handle, 3, ctypes.byref(on), 4)
    return opened


class HideLycheDialogs:
    """Keep Lyche's new top-level windows off-screen while the bot works in the background.

    An out-of-context WinEvent hook (EVENT_OBJECT_CREATE..LOCATIONCHANGE) filtered to the Lyche
    process runs on its own thread with its own message loop, so windows are parked within
    milliseconds even while the worker thread is busy in a long UIA call. Only top-level
    windows that did not exist when the hook started are touched (never the launcher or the
    Cross Tabulation window); `prefixes` optionally narrows it to certain titles.
    Lyche's WPF windows reject WS_EX_LAYERED (verified live), so windows are moved, not faded."""
    _PROC = ctypes.WINFUNCTYPE(None, wintypes.HANDLE, wintypes.DWORD, wintypes.HWND, wintypes.LONG,
                               wintypes.LONG, wintypes.DWORD, wintypes.DWORD)

    def __init__(self, pid, prefixes=None, also=()):
        self.pid, self.prefixes = pid, tuple(prefixes) if prefixes else None
        self.also = set(also)  # pre-existing windows to park anyway (the Cross Tabulation window)
        self._callback = self._PROC(self._on_event)
        self._thread = None
        self._thread_id = None
        self._existing = set()
        self.parked = []
        self._stop = None
        self._poller = None
        self._burst_until = 0.0
        import threading
        self._paused = threading.Event()
        self._suspended = threading.Event()  # set: park nothing (see suspend)
        self.original = {}  # hwnd -> on-screen rect before it was parked
        self.touched = {}  # hwnd -> original extended style (restored on exit)

    @contextmanager
    def paused(self):
        """Stop polling (and use the normal GIL switch interval) while the worker does CPU work that
        does not touch Lyche, e.g. Delete Total + NA / Del Sig on the saved file."""
        switch = sys.getswitchinterval()
        self._paused.set()
        sys.setswitchinterval(.005)
        try:
            yield
        finally:
            sys.setswitchinterval(switch)
            self._paused.clear()

    @contextmanager
    def suspend(self):
        """Park nothing for a moment: Lyche.menu_click_on_screen needs the window and its drop-down
        on screen (under the app's screen freeze) to click a gallery item."""
        self._suspended.set()
        try:
            yield
        finally:
            self._suspended.clear()

    def burst(self, seconds=2.0):
        """Spin the poller without sleeping for a moment — call right before an action that opens
        a Lyche window, so it is parked within a fraction of a millisecond of being shown."""
        self._burst_until = time.monotonic() + seconds

    def __enter__(self):
        import threading
        def collect(handle, _):
            if win32process.GetWindowThreadProcessId(handle)[1] == self.pid:
                self._existing.add(handle)
        win32gui.EnumWindows(collect, None)
        ready = threading.Event()
        def loop():
            import win32api
            user32 = ctypes.windll.user32
            user32.SetWinEventHook.restype = wintypes.HANDLE
            user32.SetWinEventHook.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.HMODULE, self._PROC,
                                               wintypes.DWORD, wintypes.DWORD, wintypes.DWORD]
            hook = user32.SetWinEventHook(0x8000, 0x800B, None, self._callback, self.pid, 0, 0)  # OUTOFCONTEXT
            self._thread_id = win32api.GetCurrentThreadId()
            ready.set()
            win32gui.PumpMessages()  # until WM_QUIT from __exit__
            if hook:
                user32.UnhookWinEvent(wintypes.HANDLE(hook))
        self._thread = threading.Thread(target=loop, name='lyche-parker', daemon=True)
        self._thread.start()
        ready.wait(2)
        # WinEvents arrive ~100 ms after a window is shown (measured live) — long enough to see it,
        # e.g. the Save As dialog, which also clamps itself onto the monitor. So also poll every
        # top-level window (<1 ms per pass) every 2 ms, faster than a 16 ms frame.
        self._stop = threading.Event()
        def lyche_threads():
            threads = set()
            def collect(handle, _):
                thread, process = win32process.GetWindowThreadProcessId(handle)
                if process == self.pid:
                    threads.add(thread)
            win32gui.EnumWindows(collect, None)
            return threads
        def poll():
            # Only Lyche's own UI thread(s) are scanned (EnumThreadWindows): scanning every window on
            # the desktop every 2 ms starved the worker's Python work of the GIL — Del Sig took minutes.
            threads, refreshed = lyche_threads(), time.monotonic()
            visible = win32gui.IsWindowVisible
            def check(hwnd, _):
                if visible(hwnd):
                    self._consider(hwnd, 'poll')
                return True
            while not self._stop.is_set():
                if self._paused.is_set():
                    time.sleep(.1)
                    continue
                if time.monotonic() - refreshed > 2:
                    threads, refreshed = lyche_threads(), time.monotonic()
                for thread in threads:
                    try:
                        win32gui.EnumThreadWindows(thread, check, None)
                    except win32gui.error:
                        pass
                # Never a busy spin (user request: keep CPU low): 1 ms right after a press, else 5 ms.
                time.sleep(.001 if time.monotonic() < self._burst_until else .005)
        self._poller = threading.Thread(target=poll, name='lyche-poller', daemon=True)
        self._poller.start()
        return self

    def __exit__(self, *_):
        if self._stop:
            self._stop.set()
            self._poller.join(2)
        if self._thread_id:
            import win32api
            win32api.PostThreadMessage(self._thread_id, 0x0012, 0, 0)  # WM_QUIT
            self._thread.join(2)
            self._thread_id = None
        for hwnd, ex in list(self.touched.items()):  # windows that survive get their taskbar style back
            try:
                if win32gui.IsWindow(hwnd) and win32gui.GetWindowLong(hwnd, -20) != ex:
                    win32gui.SetWindowLong(hwnd, -20, ex)
            except win32gui.error:
                pass

    def _on_event(self, _hook, event, hwnd, id_object, _child, _thread, _time):
        if id_object == 0 and hwnd and event in (0x8000, 0x8002, 0x800B):
            self._consider(hwnd, hex(event))

    def _consider(self, hwnd, source):
        try:
            if self._suspended.is_set():
                return
            if hwnd in self._existing and hwnd not in self.also:
                return
            if win32gui.IsIconic(hwnd):
                return
            if win32gui.GetWindowLong(hwnd, -16) & 0x40000000:  # WS_CHILD: controls inside dialogs
                return
            if self.prefixes and not win32gui.GetWindowText(hwnd).startswith(self.prefixes):
                return
            rect = win32gui.GetWindowRect(hwnd)
            if rect[0] > -15000:
                self.original.setdefault(hwnd, rect)
            detected = time.time()
            if self.hide(hwnd):
                self.parked.append((round(detected, 3), round(time.time() - detected, 3), source,
                                    win32gui.GetWindowText(hwnd)[:30] or win32gui.GetClassName(hwnd)))
        except Exception:
            pass

    def hide(self, hwnd):
        """Park a Lyche window off-screen. The move waits for Lyche's UI thread (66-242 ms measured
        live). An empty window region at creation was tried and dropped: it waits for the same
        thread, did not shorten the flash, and a run then missed the 'Saved.' box."""
        try:
            if hwnd not in self.touched:
                ex = win32gui.GetWindowLong(hwnd, -20)  # GWL_EXSTYLE
                self.touched[hwnd] = ex
                title = win32gui.GetWindowText(hwnd)
                if not title.startswith(WINDOW_PREFIXES):
                    win32gui.SetWindowLong(hwnd, -20, ex | 0x00000080)  # TOOLWINDOW: no taskbar button
            if win32gui.GetWindowRect(hwnd)[0] <= -15000:
                return False  # already parked (avoids LOCATIONCHANGE recursion)
            win32gui.SetWindowPos(hwnd, 0, -20000, -20000, 0, 0, 0x0001 | 0x0004 | 0x0010)  # NOSIZE NOZORDER NOACTIVATE
            return True
        except win32gui.error:
            return False

    @staticmethod
    def pump(seconds):
        time.sleep(seconds)  # the hook has its own message loop; kept for callers that wait


class Lyche:
    def __init__(self, handle: int, control_dir: Path, emit, timeout=900, background=True):
        self.background = background  # UIA patterns only, Lyche kept off-screen (see background_session)
        self.hider = None
        self.matrix = False  # the last run_export was a Matrix History (worker: Del Sig in Matrix mode)
        self._freeze_depth = 0  # frozen_screen nesting (one freeze/unfreeze for the outermost only)
        self.handle = handle
        self.control_dir = control_dir
        self.emit = emit
        self.timeout = timeout
        self.desktop = Desktop(backend='uia')
        self.api = IUIA().iuia
        if not win32gui.IsWindow(handle):
            raise ValueError('หน้าต่าง Lyche ปิดแล้ว กรุณาค้นหาหน้าต่างใหม่')
        self.pid = win32process.GetWindowThreadProcessId(handle)[1]
        self.project = self.project_of(win32gui.GetWindowText(handle))
        if not self.project:
            raise ValueError('ไม่ใช่หน้าต่าง Cross Tabulation / Tabulation Result ของ Lyche')

    @staticmethod
    def project_of(title):
        match = re.search(r'<([^>]+)>', title)
        return match.group(1) if match else ''

    def checkpoint(self):
        if (self.control_dir / 'stop').exists() or ctypes.windll.user32.GetAsyncKeyState(0x77) & 0x8000:
            raise Stopped('หยุดแล้ว (F8 / Stop) — Lyche อาจยังคำนวณหรือบันทึกงานอยู่')
        while (self.control_dir / 'pause').exists():
            if (self.control_dir / 'stop').exists() or ctypes.windll.user32.GetAsyncKeyState(0x77) & 0x8000:
                raise Stopped('หยุดขณะพัก')
            time.sleep(.15)

    def lyche_error(self):
        now = time.monotonic()
        if now - getattr(self, '_error_checked', 0.0) < 1:  # cheap: at most once per second
            return None
        self._error_checked = now
        return self._lyche_error_now()

    def _lyche_error_now(self):
        """Text of a visible Lyche 'Error' message box, if any (e.g. 'Unknown Error Occurred.' when the
        Lyche server refuses a request — seen live as HTTP 403 in LycheError.log)."""
        for handle in self.visible_dialogs('Error'):
            try:
                return ' '.join(x.window_text().strip() for x in self.all(self.wrapper(handle), kind=50020)).strip() or 'Error'
            except Exception:
                return 'Error'
        return None

    def visible_dialogs(self, prefix):
        """Every visible top-level window of this Lyche whose title starts with `prefix`. Walks only
        the Lyche UI thread's windows (its dialogs live there), not every window on the desktop."""
        found = []
        def collect(handle, _):
            if win32gui.IsWindowVisible(handle) and win32gui.GetWindowText(handle).startswith(prefix):
                found.append(handle)
            return True
        try:
            win32gui.EnumThreadWindows(win32process.GetWindowThreadProcessId(self.handle)[0], collect, None)
        except win32gui.error:
            def collect_all(handle, _):
                if win32process.GetWindowThreadProcessId(handle)[1] == self.pid:
                    collect(handle, _)
            win32gui.EnumWindows(collect_all, None)
        return found

    def wait(self, fn, label, timeout=20, poll=.2):
        start = time.monotonic()
        last = None
        while time.monotonic() - start < timeout:
            self.checkpoint()
            error = self.lyche_error()
            if error:
                raise RuntimeError(f'Lyche แจ้งข้อผิดพลาด: {error} — ตรวจการเชื่อมต่อ/ล็อกอินของ Lyche แล้วลองใหม่')
            try:
                result = fn()
                if result:
                    return result
            except (LookupError, RuntimeError) as exc:
                last = exc
            time.sleep(poll)
        raise TimeoutError(f'รอ {label} เกิน {timeout} วินาที' + (f': {last}' if last else ''))

    def wrapper(self, handle):
        return UIAWrapper(UIAElementInfo(self.api.ElementFromHandle(handle)))

    def main(self):
        if not win32gui.IsWindow(self.handle):
            raise RuntimeError('หน้าต่างต้นทางปิดแล้ว')
        title = win32gui.GetWindowText(self.handle)
        if self.project_of(title) != self.project:
            raise RuntimeError('โปรเจกต์ใน Lyche เปลี่ยนไป หยุดเพื่อป้องกันรันผิดงาน')
        if win32gui.IsIconic(self.handle) and not self.background:  # auto-opened Cross Tabulation starts minimised
            win32gui.ShowWindow(self.handle, 9)  # SW_RESTORE
            time.sleep(.5)
        return self.wrapper(self.handle)

    def dialog(self, prefix, visible_only=True):
        """visible_only=False also finds dialogs Windows hides while their owner is minimised."""
        found = []
        def collect(handle, _):
            if (not visible_only or win32gui.IsWindowVisible(handle)) \
                    and win32process.GetWindowThreadProcessId(handle)[1] == self.pid:
                title = win32gui.GetWindowText(handle)
                if title.startswith(prefix):
                    project = self.project_of(title)
                    if not project or project == self.project:
                        found.append(handle)
        win32gui.EnumWindows(collect, None)
        if len(found) > 1 and not visible_only:
            found = [h for h in found if win32gui.IsWindowVisible(h)] or found
        if len(found) > 1:
            raise RuntimeError(f'พบหน้าต่าง {prefix} มากกว่าหนึ่งหน้าต่าง')
        return self.wrapper(found[0]) if found else None

    def find(self, root, *, aid=None, name=None, kind=None, required=True):
        conditions = []
        for prop, value in ((30011, aid), (30005, name), (30003, kind)):
            if value is not None:
                conditions.append(self.api.CreatePropertyCondition(prop, value))
        condition = conditions[0]
        for extra in conditions[1:]:
            condition = self.api.CreateAndCondition(condition, extra)
        element = root.element_info.element.FindFirst(4, condition)
        if not element:
            if required:
                raise LookupError(f'ไม่พบปุ่ม/ช่อง {aid or name}')
            return None
        return UIAWrapper(UIAElementInfo(element))

    def all(self, root, *, aid=None, kind=None):
        condition = self.api.CreatePropertyCondition(30011 if aid else 30003, aid or kind)
        elements = root.element_info.element.FindAll(4, condition)
        return [UIAWrapper(UIAElementInfo(elements.GetElement(i))) for i in range(elements.Length)]

    def click(self, control):
        self.checkpoint()
        if not control.is_enabled():
            raise RuntimeError(f'ปุ่มยังใช้ไม่ได้: {control.window_text()}')
        if self.background:
            self.press(control)
            return
        # Lyche ribbon Invoke can return without opening its dialog. Use the
        # current control rectangle, never a stored screen coordinate.
        control.top_level_parent().set_focus()
        control.click_input()

    def press(self, control):
        """Activate a control through UIA patterns only (no mouse, no focus change)."""
        if self.hider:
            self.hider.burst()  # any press may open a Lyche window
        info = control.element_info.element
        for pattern_id, action in ((10000, 'Invoke'), (10015, 'Toggle'), (10010, 'Select'), (10005, 'Expand')):
            try:
                pattern = info.GetCurrentPattern(pattern_id)
            except Exception:
                pattern = None
            if not pattern:
                continue
            if action == 'Invoke':
                control.iface_invoke.Invoke()
            elif action == 'Toggle':
                control.iface_toggle.Toggle()
            elif action == 'Select':
                control.iface_selection_item.Select()
            else:
                control.iface_expand_collapse.Expand()
            return
        try:
            # pywinauto 0.6.9 has no iface_legacy_iaccessible: use the raw UIA pattern (10018).
            from comtypes.gen.UIAutomationClient import IUIAutomationLegacyIAccessiblePattern
            info.GetCurrentPattern(10018).QueryInterface(IUIAutomationLegacyIAccessiblePattern).DoDefaultAction()
        except Exception as exc:
            raise RuntimeError(f'กดปุ่มเบื้องหลังไม่ได้: {control.window_text() or control.element_info.automation_id}') from exc

    @contextmanager
    def quiet(self):
        """Pause the background window hider while doing work that does not involve Lyche."""
        if self.hider is None:
            yield
            return
        with self.hider.paused():
            yield

    @contextmanager
    def background_session(self, keep_minimised=False):
        """Run with Lyche out of sight. A minimised Cross Tabulation window stays minimised (UIA and
        WPF layout keep working; restoring it would animate on screen and Windows clamps an
        off-screen restore position back to 0,0). A normal window is moved off-screen; a maximised
        one is minimised without activation. Every new Lyche window is parked off-screen by
        HideLycheDialogs. The window's original state is put back afterwards.
        keep_minimised: a minimised window is not brought back to normal for the ribbon (the Banner
        Manual check needs no ribbon); if Lyche shows it by itself it is parked, then minimised again."""
        if not self.background:
            yield
            return
        handle = self.handle
        placement = win32gui.GetWindowPlacement(handle)
        was_iconic, was_max = win32gui.IsIconic(handle), placement[1] == 3
        rect = win32gui.GetWindowRect(handle)
        def transitions(disabled):
            value = ctypes.c_int(1 if disabled else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(handle, 3, ctypes.byref(value), 4)  # DWMWA_TRANSITIONS_FORCEDISABLED
        transitions(True)
        switch = sys.getswitchinterval()
        try:
            with HideLycheDialogs(self.pid, also=(handle,)) as hider:
                self.hider = hider
                if was_iconic and keep_minimised:
                    pass
                elif was_iconic or was_max:
                    # Ribbon drop-downs (Tabulate All) do not open while minimised, so bring the
                    # window back to normal without activation; the hook parks it off-screen at once.
                    win32gui.SetWindowPlacement(handle, (placement[0], 4, placement[2], placement[3], placement[4]))
                else:
                    win32gui.SetWindowPos(handle, 0, -20000, -20000, 0, 0, 0x0001 | 0x0004 | 0x0010)
                end = time.monotonic() + 2
                while win32gui.GetWindowRect(handle)[0] > -15000 and time.monotonic() < end:
                    time.sleep(.02)
                if win32gui.GetWindowRect(handle)[0] > -15000:
                    win32gui.SetWindowPos(handle, 0, -20000, -20000, 0, 0, 0x0001 | 0x0004 | 0x0010)
                for stale in self.saved_dialogs():  # left over by an earlier run
                    self.close_saved(stale)
                try:
                    yield
                except BaseException:
                    self.close_leftover_dialogs()  # an invisible modal dialog would lock Lyche for the user
                    raise
                finally:
                    self.emit('log', text=f'ทำงานเบื้องหลัง: ซ่อนหน้าต่าง Lyche {len(hider.parked)} ครั้ง')
                    if os.environ.get('LYCHE_PARK_DEBUG'):
                        self.emit('log', text=f'parked: {hider.parked}')
        finally:
            sys.setswitchinterval(switch)
            self.hider = None
            # The hook is gone now, so putting the window back is not undone by it.
            if win32gui.IsWindow(handle):
                if was_iconic:
                    win32gui.SetWindowPlacement(handle, (placement[0], 7, placement[2], placement[3], placement[4]))
                elif was_max:
                    win32gui.SetWindowPlacement(handle, placement)
                else:
                    win32gui.SetWindowPos(handle, 0, rect[0], rect[1], 0, 0, 0x0001 | 0x0004 | 0x0010)
                transitions(False)

    def focus(self, root):
        self.checkpoint()
        root.set_focus()

    def text(self, root, aid):
        control = self.find(root, aid=aid, required=False)
        return control.window_text().strip() if control else ''

    def menu(self, aid, item):
        if self.background:
            return self.menu_background(aid, item)
        root = self.main()
        self.focus(root)
        control = self.find(root, aid=aid)
        self.checkpoint()
        if control.element_info.control_type == 'SplitButton':
            rect = control.rectangle()
            control.click_input(coords=(rect.width() - 8, rect.height() - 8))
        else:
            self.click(control)
        def lookup():
            # WPF menus may be separate popup HWNDs in the same process.
            for window in self.desktop.windows(process=self.pid, visible_only=True):
                found = self.find(window, name=item, kind=50011, required=False)
                if found:
                    return found
            return self.find(self.main(), name=item, required=False)
        self.click(self.wait(lookup, item))

    def menu_background(self, aid, item):
        """Open a ribbon (split) button's drop-down with ExpandCollapse and Invoke the item.
        A WPF ribbon drop-down closes at once unless its window is active (live failure on Beppu:
        Expand left the state at Collapsed), so the off-screen window gets the foreground first."""
        root = self.main()
        control = self.find(root, aid=aid)
        self.checkpoint()
        hold = self.control_dir / 'hold-focus'  # the app's focus guard waits meanwhile
        hold.touch()
        user_window = user_foreground(self.pid)  # handed back as soon as the item is invoked
        held_since = time.monotonic()
        found = None
        try:
            for _ in range(5):  # user activity can steal focus and close the drop-down: retry
                force_foreground(self.handle)
                try:
                    control.iface_expand_collapse.Expand()
                except Exception:
                    self.press(control)
                try:
                    found = self.wait(lambda: self._menu_item(item), item, 5, poll=.05)
                    break
                except TimeoutError:
                    foreground = ctypes.windll.user32.GetForegroundWindow()
                    try:
                        state = control.iface_expand_collapse.CurrentExpandCollapseState
                    except Exception:
                        state = '?'
                    self.emit('log', text=f'เมนู {item} ยังไม่เปิด: foreground={win32gui.GetWindowText(foreground)[:40]!r} '
                                          f'(Lyche={foreground == self.handle}) state={state} enabled={control.is_enabled()}')
                    try:
                        control.iface_expand_collapse.Collapse()
                    except Exception:
                        pass
            if found is None:
                raise TimeoutError(f'เปิดเมนู {item} ไม่สำเร็จ')
            self.press(found)
        finally:
            give_back_foreground(user_window)
            hold.unlink(missing_ok=True)
            if os.environ.get('LYCHE_FOCUS_DEBUG'):
                self.emit('log', text=f'focus hold menu {item}: {(time.monotonic() - held_since) * 1000:.0f} ms')

    def _menu_item(self, item):
        # MenuItem only: a same-named Text label has no Invoke pattern (live failure).
        # Lyche exposes each ribbon menu item several times; only some copies support Invoke.
        # One targeted UIA query per visible Lyche window (the drop-down is its own popup window):
        # scanning every MenuItem was slow and kept Lyche in the foreground ~2 s.
        condition = self.api.CreateAndCondition(self.api.CreatePropertyCondition(30005, item),
                                                self.api.CreatePropertyCondition(30003, 50011))
        handles = [self.handle]
        def collect(handle, _):
            if handle != self.handle and win32gui.IsWindowVisible(handle)                     and win32process.GetWindowThreadProcessId(handle)[1] == self.pid:
                handles.append(handle)
        win32gui.EnumWindows(collect, None)
        for handle in handles:
            try:
                found = self.api.ElementFromHandle(handle).FindAll(4, condition)
            except Exception:
                continue
            for index in range(found.Length):
                element = found.GetElement(index)
                try:
                    if element.CurrentIsEnabled and element.GetCurrentPattern(10000):
                        return UIAWrapper(UIAElementInfo(element))
                except Exception:
                    pass
        return None

    def _gallery_item(self, item):
        """The on-screen copy of a ribbon gallery item (drop-down popup window, not the main window)."""
        condition = self.api.CreatePropertyCondition(30005, item)
        handles = []
        def collect(handle, _):
            if handle != self.handle and win32gui.IsWindowVisible(handle) \
                    and win32process.GetWindowThreadProcessId(handle)[1] == self.pid:
                handles.append(handle)
        win32gui.EnumWindows(collect, None)
        for handle in handles:
            try:
                found = self.api.ElementFromHandle(handle).FindAll(4, condition)
            except Exception:
                continue
            for index in range(found.Length):
                element = found.GetElement(index)
                try:
                    rect = element.CurrentBoundingRectangle
                    if element.CurrentControlType in (50011, 50007) and element.CurrentIsEnabled \
                            and not element.CurrentIsOffscreen and rect.right > rect.left and rect.left > -15000:
                        return UIAWrapper(UIAElementInfo(element))
                except Exception:
                    pass
        return None

    @contextmanager
    def frozen_screen(self):
        """Ask the app to cover the screens with a still image (same protocol as worker.screen_freeze:
        emit 'freeze', wait ≤ 2 s for control_dir/frozen) so a moment of Lyche on screen is not seen."""
        if not self.background or self._freeze_depth:  # nested: the outer freeze already covers it
            self._freeze_depth += 1
            try:
                yield
            finally:
                self._freeze_depth -= 1
            return
        self._freeze_depth = 1
        flag = self.control_dir / 'frozen'
        flag.unlink(missing_ok=True)
        self.emit('freeze')
        end = time.monotonic() + 2
        while not flag.exists() and time.monotonic() < end:
            time.sleep(.02)
        try:
            yield
        finally:
            self._freeze_depth = 0
            self.emit('unfreeze')
            flag.unlink(missing_ok=True)

    def menu_click_on_screen(self, aid, item, expect=None):
        """BG mode, for ribbon gallery items that react to a real click only. Live (2026-10-01) Matrix
        Export all → Excel(Separated Sheets) ignored every UIA pattern (Invoke/LegacyIA/Select) and a
        posted or typed Enter. User choice: under the app's screen freeze bring the window on screen,
        open the drop-down, click the item at its current rectangle, put the cursor back and park the
        window again. `expect` = title prefix of the window the click opens (waited for, then parked)."""
        import win32api
        hold = self.control_dir / 'hold-focus'  # the app's focus guard waits meanwhile
        hold.touch()
        user_window = user_foreground(self.pid)
        cursor = win32api.GetCursorPos()
        parked = win32gui.GetWindowRect(self.handle)
        started = time.monotonic()
        try:
            with self.frozen_screen():
                with (self.hider.suspend() if self.hider else nullcontext()):
                    work = win32api.GetMonitorInfo(win32api.MonitorFromPoint(cursor, 2))['Work']
                    win32gui.SetWindowPos(self.handle, 0, work[0] + 40, work[1] + 40, 0, 0, 0x0001 | 0x0004 | 0x0010)
                    clicked = False
                    for _ in range(3):
                        self.checkpoint()
                        force_foreground(self.handle)
                        control = self.find(self.main(), aid=aid)
                        try:
                            control.iface_expand_collapse.Expand()
                        except Exception:
                            self.press(control)
                        try:
                            found = self.wait(lambda: self._gallery_item(item), item, 5, poll=.05)
                        except TimeoutError:
                            continue
                        found.click_input()
                        clicked = True
                        break
                    win32api.SetCursorPos(cursor)
                    if not clicked:
                        raise TimeoutError(f'เปิดเมนู {item} ไม่สำเร็จ')
                win32gui.SetWindowPos(self.handle, 0, parked[0], parked[1], 0, 0, 0x0001 | 0x0004 | 0x0010)
                if expect:  # keep the freeze until the window the click opened is parked by the hider
                    end = time.monotonic() + 3
                    while time.monotonic() < end and not self.dialog(expect, visible_only=False):
                        time.sleep(.05)
                    time.sleep(.15)
        finally:
            try:
                win32api.SetCursorPos(cursor)
            except Exception:
                pass
            give_back_foreground(user_window)
            hold.unlink(missing_ok=True)
        self.emit('log', text=f'คลิกเมนู {item} (แช่จอ {time.monotonic() - started:.1f} วินาที)')

    def settings(self):
        root = self.main()
        if root.window_text().startswith('Tabulation Result'):
            self.click(self.find(root, name='Back to Setting', kind=50000))
            self.wait(lambda: win32gui.GetWindowText(self.handle).startswith(SETTINGS_PREFIXES), 'Back to Setting')
        if not win32gui.GetWindowText(self.handle).startswith(SETTINGS_PREFIXES):
            raise RuntimeError('กรุณาเปิดหน้า Cross Tabulation ก่อน')
        self.wait_ready()
        return self.main()

    def wait_ready(self, timeout=60):
        """A Cross Tabulation window that was just opened has its title before its ribbon exists
        (live failure: 'ไม่พบปุ่ม/ช่อง rbnOpen'), so wait until Import History is there."""
        if not self.find(self.main(), aid='rbnOpen', required=False):
            self.wait(lambda: self.find(self.main(), aid='rbnOpen', required=False),
                      'หน้าต่าง Cross Tabulation โหลดเสร็จ', timeout)

    def history_dialog(self, source):
        dialog = self.dialog('Tabulation List -')
        if not dialog:
            if not self.dialog('Confirm -'):
                root = self.settings()
                self.click(self.find(root, aid='rbnOpen'))
            def opened():
                confirm = self.dialog('Confirm -')
                if confirm:
                    text = ' '.join(x.window_text() for x in self.all(confirm, kind=50020))
                    if text.strip() != 'Change will not be saved. Proceed?':
                        raise ValueError(f'พบคำถามที่ไม่รู้จัก: {text}')
                    yes = self.find(confirm, name='Yes', kind=50000, required=False) or self.find(confirm, name='&Yes', kind=50000)
                    self.click(yes)
                    self.emit('log', text='Change will not be saved. Proceed? → กด Yes อัตโนมัติ')
                    return None
                return self.dialog('Tabulation List -')
            dialog = self.wait(opened, 'Import History')
        tab = self.find(dialog, aid='tabPersonal' if source == 'Personal' else 'tabShare', required=False) \
            or self.find(dialog, name='Personal Tabulation' if source == 'Personal' else 'Shared Tabulation', kind=50019)
        self.checkpoint()
        tab.select()
        return dialog

    def history_rows(self, dialog):
        grid = self.find(dialog, aid='summaryList')
        rows = self.all(grid, kind=50029)
        result = []
        for row in rows:
            texts = [x.window_text().strip() for x in self.all(row, kind=50020)]
            names = [x for x in texts if x and not re.fullmatch(r'\d{4}-\d\d-\d\d.*', x)]
            if names:
                result.append((names[0], row))
        return result

    def scan_pages(self, grid, collect):
        """Visit UIA virtualized rows without guessing scrollbar coordinates."""
        try:
            scroll = grid.iface_scroll
            if scroll.CurrentVerticallyScrollable:
                scroll.SetScrollPercent(-1, 0)
        except (AttributeError, NotImplementedError, NoPatternInterfaceError):
            scroll = None
        previous = None
        for _ in range(500):
            self.checkpoint()
            result = collect()
            yield result
            if scroll is None or not scroll.CurrentVerticallyScrollable:
                return
            percent = scroll.CurrentVerticalScrollPercent
            if percent >= 100 or percent == previous:
                return
            previous = percent
            scroll.Scroll(2, 3)  # No horizontal scroll; one vertical large increment.
            time.sleep(.15)
        raise RuntimeError('รายการยาวเกินขอบเขตการอ่าน โปรดใช้ช่องพิมพ์ชื่อแทน')

    def get_banners(self, source):
        """UI read of the Import History list — used only for Shared, which lives on the Lyche
        server (no local file). Personal stays on the background XML path in history.py."""
        was_minimized = win32gui.IsIconic(self.handle)
        try:
            return self._read_history_quietly(source)
        except Stopped:
            raise
        except Exception as exc:  # fall back to the visible path rather than failing Get Banner
            self.emit('log', text=f'อ่านเบื้องหลังไม่สำเร็จ ({exc}) — เปิด Import History ตามปกติ')
            leftover = self.dialog('Tabulation List -', visible_only=False)
            if leftover:
                self.close_history(leftover)
        names = self._read_history_names(source)
        if was_minimized:
            win32gui.ShowWindow(self.handle, 7)  # SW_SHOWMINNOACTIVE: put it back out of the way
        return names

    def _read_history_quietly(self, source):
        """Import History via UIA patterns only: the Cross Tabulation window stays as it is
        (minimised or not), no mouse. A WinEvent hook makes Lyche's Tabulation List / Confirm
        windows fully transparent and off-screen the moment Windows shows them."""
        root = self.wrapper(self.handle)
        if win32gui.GetWindowText(self.handle).startswith('Tabulation Result'):
            raise RuntimeError('หน้าต่างอยู่ที่ Tabulation Result')
        with HideLycheDialogs(self.pid, ('Tabulation List -', 'Confirm -')) as hider:
            dialog = self.dialog('Tabulation List -', visible_only=False)
            if not dialog:
                self.wait_ready()
                self.find(self.wrapper(self.handle), aid='rbnOpen').iface_invoke.Invoke()
                end = time.monotonic() + 20
                while not dialog:
                    self.checkpoint()
                    if time.monotonic() > end:
                        raise TimeoutError('รอ Import History เกิน 20 วินาที')
                    hider.pump(.03)
                    confirm = self.dialog('Confirm -', visible_only=False)
                    if confirm:
                        hider.hide(confirm.handle)
                        text = ' '.join(x.window_text() for x in self.all(confirm, kind=50020)).strip()
                        if text != 'Change will not be saved. Proceed?':
                            raise ValueError(f'พบคำถามที่ไม่รู้จัก: {text}')
                        yes = self.find(confirm, name='Yes', kind=50000, required=False) or self.find(confirm, name='&Yes', kind=50000)
                        yes.iface_invoke.Invoke()
                        self.emit('log', text='Change will not be saved. Proceed? → กด Yes อัตโนมัติ')
                    dialog = self.dialog('Tabulation List -', visible_only=False)
            hider.hide(dialog.handle)
            return self._read_open_history(dialog, source)

    def _read_open_history(self, dialog, source):
        try:
            tab = self.find(dialog, aid='tabPersonal' if source == 'Personal' else 'tabShare')
            tab.iface_selection_item.Select()
            expand = self.find(dialog, aid='linkExpand', required=False)
            if expand:
                expand.iface_invoke.Invoke()
            time.sleep(.2)
            names = []
            grid = self.find(dialog, aid='summaryList')
            for rows in self.scan_pages(grid, lambda: self.history_rows(dialog)):
                names.extend(name for name, _ in rows)
            return list(dict.fromkeys(names))
        finally:
            self.close_history(dialog)

    def _read_history_names(self, source):
        dialog = self.history_dialog(source)
        # Expand folders first. A separate search field supports names outside visible pages.
        expand = self.find(dialog, aid='linkExpand', required=False)
        if expand:
            self.click(expand)
        names = []
        grid = self.find(dialog, aid='summaryList')
        for rows in self.scan_pages(grid, lambda: self.history_rows(dialog)):
            names.extend(name for name, _ in rows)
        names = list(dict.fromkeys(names))
        self.close_history(dialog)
        return names

    def close_history(self, dialog):
        """Close Tabulation List without loading anything. A mouse click on the WPF close
        button was ignored live, so try UIA Invoke, then WindowPattern.Close, then WM_CLOSE."""
        def invoke():
            self.find(dialog, aid='PART_CloseWindowButton').iface_invoke.Invoke()
        attempts = (invoke, lambda: dialog.iface_window.Close(),
                    lambda: win32gui.PostMessage(dialog.handle, 0x0010, 0, 0))
        for attempt in attempts:
            self.checkpoint()
            try:
                attempt()
            except Exception:
                continue
            end = time.monotonic() + 4
            while time.monotonic() < end:
                if not self.dialog('Tabulation List -', visible_only=False):
                    return
                time.sleep(.15)
        raise RuntimeError('ปิดหน้าต่าง Import History ไม่สำเร็จ')

    def load(self, name, source):
        dialog = self.history_dialog(source)
        search = self.find(dialog, aid='txtKeyWord')
        self.checkpoint()
        search.set_edit_text(name)
        self.click(self.find(dialog, aid='btnSearch'))
        def exact_row():
            matches = [row for text, row in self.history_rows(dialog) if text == name]
            if len(matches) > 1:
                raise ValueError(f'ชื่อ History ซ้ำ: {name}')
            return matches[0] if matches else None
        row = self.wait(exact_row, f'Banner {name}')
        self.checkpoint()
        row.select()
        self.click(self.find(dialog, aid='rbnSelect'))
        self.wait(lambda: not self.dialog('Tabulation List -'), 'โหลด Banner')
        self.wait(lambda: self.main().window_text().rstrip().endswith('> ' + name), 'ชื่อ Banner บนหน้าต่าง')
        root = self.main()
        # A Matrix History has only the Tabulation Item list (the Multiway list below stays empty).
        panels = ('grdTabulationItemUp',) if self.is_matrix() else ('grdTabulationItemUp', 'grdTabulationItemDown')
        for panel in panels:
            grid = self.find(self.find(root, aid=panel), aid='lvTabulationItem')
            if not self.all(grid, kind=50029):
                raise RuntimeError('Banner หรือ Stub ว่างหลังโหลด History' if len(panels) > 1
                                   else 'Tabulation Item ว่างหลังโหลด History (Matrix)')
        self.emit('log', text=f'โหลด {name} แล้ว' + (' (Matrix)' if self.is_matrix() else ''))

    def is_matrix(self):
        """True while the settings window is Lyche's Matrix one ('Tabulation for Matix settings')."""
        return win32gui.IsWindow(self.handle) and win32gui.GetWindowText(self.handle).startswith(MATRIX_PREFIXES)

    def banner_codes(self, root=None):
        """Item codes in the Banner list ('QUOTA1 QUOTA1.Quota: Brand' → 'QUOTA1'), top to bottom."""
        grid = self.find(self.find(root or self.main(), aid='grdTabulationItemUp'), aid='lvTabulationItem')
        codes = []
        for row in self.all(grid, kind=50029):
            cells = [x.window_text().strip() for x in self.all(row, kind=50020)]
            text = next((c for c in cells if c and not c.isdigit()), '')
            codes.append(text.split(' ', 1)[0])
        return codes

    def set_banner_items(self, items):
        """Banner Manual (user request): replace the History's Banner with `items`, in order. As a user
        would: Banner 'Clear all' → Yes, then per item: search the item list, select the item whose
        code matches exactly, 'To Banner'. UIA patterns only, so it also works off-screen."""
        if self.is_matrix():
            raise RuntimeError('Banner Manual ใช้กับ Banner แบบ Matrix ไม่ได้ — ปิด Banner Manual ของแถวนี้')
        root = self.settings()
        panel = self.find(root, aid='grdTabulationItemUp')
        if self.banner_codes(root):
            self.click(self.find(panel, aid='btnClrRows'))
            def cleared():
                for prefix in ('Confirm', 'Question'):
                    for handle in self.visible_dialogs(prefix):
                        dialog = self.wrapper(handle)
                        yes = self.find(dialog, name='Yes', kind=50000, required=False) \
                            or self.find(dialog, name='&Yes', kind=50000, required=False)
                        if yes:
                            text = ' '.join(x.window_text() for x in self.all(dialog, kind=50020)).strip()
                            self.press(yes)
                            self.emit('log', text=f'Clear all Banner: {text or win32gui.GetWindowText(handle)} → Yes')
                            return False  # wait for the list to empty
                return not self.banner_codes(root)
            self.wait(cleared, 'Clear all Banner')
        for index, item in enumerate(items, 1):
            self.checkpoint()
            row, _ = self.find_item(root, item)
            if row is None:
                raise RuntimeError(f'Banner Manual: ไม่พบข้อ {item} ในรายการ Item ของ Lyche')
            self.checkpoint()
            row.select()
            time.sleep(.2)
            self.click(self.find(root, aid='btnToColQ'))
            def added():
                codes = self.banner_codes(root)
                return codes if len(codes) >= index else None
            codes = self.wait(added, f'To Banner {item}')
            if [c.casefold() for c in codes] != [x.casefold() for x in items[:index]]:
                raise RuntimeError(f'Banner Manual: Banner ไม่ตรงหลังเพิ่ม {item} (ได้ {", ".join(codes)})')
        self.emit('log', text=f'Banner Manual: {", ".join(items)}')

    def find_item(self, root, item):
        """Search the settings window's item list for the item whose code is exactly `item` (case aside).
        Returns (row, its text such as 'QUOTA6 QUOTA6') or (None, '')."""
        search_panel = self.find(root, aid='searchPanel')
        tree = self.find(root, aid='treeListViewQ')
        self.find(search_panel, aid='txtKeyWord').set_edit_text(item)
        self.click(self.find(search_panel, aid='btnSearch'))
        def exact_row():
            for row in self.all(tree, kind=50029):
                for text in (x.window_text().strip() for x in self.all(row, kind=50020)):
                    if text.split(' ', 1)[0].casefold() == item.casefold():
                        return row, text
            return None
        # The search jumps to its first hit ('Quota1' also hits QUOTA10…); step with Next until the
        # exact item is among the rows the list has drawn.
        for _ in range(30):
            try:
                return self.wait(exact_row, f'ข้อ {item}', 3)
            except TimeoutError:
                hits = re.search(r'(\d+)\s*/\s*(\d+)', self.text(search_panel, 'txtShowinfo'))
                if not hits or int(hits.group(1)) >= int(hits.group(2)):
                    break
                self.click(self.find(search_panel, aid='btnNext'))
        return None, ''

    def check_items(self, items):
        """Banner Manual 'เช็คกับ Lyche': which items exist in the project's item list. Search only — the
        Banner is not touched; the search box is cleared afterwards."""
        if self.is_matrix():
            raise RuntimeError('หน้าต่าง Lyche ตอนนี้เป็น Matrix — Banner Manual ใช้กับ Banner แบบ Matrix ไม่ได้')
        root = self.settings()
        results = []
        try:
            for item in items:
                self.checkpoint()
                row, text = self.find_item(root, item)
                results.append({'item': item, 'found': row is not None, 'label': text})
        finally:
            try:
                panel = self.find(root, aid='searchPanel')
                self.find(panel, aid='txtKeyWord').set_edit_text('')
            except Exception:
                pass
        found = sum(r['found'] for r in results)
        self.emit('log', text=f'เช็คข้อ Banner Manual กับ Lyche: พบ {found}/{len(results)} ข้อ'
                              + ''.join(f' · ไม่พบ {r["item"]}' for r in results if not r['found']))
        return results

    def combo_value(self, combo):
        try:
            return combo.iface_value.CurrentValue.strip()
        except (NoPatternInterfaceError, AttributeError):
            return combo.window_text().strip()

    def choose_combo(self, dialog, aid, value, options=()):
        """Open the dropdown and click the option, like a user; select() cannot see WinForms items.

        If the open list exposes no matching item names, fall back to arrow keys using the
        known option order (`options`)."""
        combo = self.find(dialog, aid=aid)
        if self.combo_value(combo) == value:
            return
        if self.background and value in options:
            return self.choose_combo_background(combo, value, options)
        self.click(combo)
        def option():
            roots = [combo]
            def collect(handle, _):
                if win32gui.IsWindowVisible(handle) and not win32gui.GetWindowText(handle) \
                        and win32process.GetWindowThreadProcessId(handle)[1] == self.pid:
                    roots.append(self.wrapper(handle))
            win32gui.EnumWindows(collect, None)
            for root in roots:
                for item in self.all(root, kind=50007):
                    if item.window_text().strip().casefold() == value.casefold():
                        return item
            return None
        start = time.monotonic()
        item = None
        while time.monotonic() - start < (3 if options else 10) and not item:
            self.checkpoint()
            item = option()
            if not item:
                time.sleep(.2)
        self.checkpoint()
        if item:
            item.click_input()
        elif value in options:
            send_keys('{HOME}' + '{DOWN}' * options.index(value) + '{ENTER}', pause=0.05)
            self.emit('log', text=f'เลือก {value} ด้วยคีย์บอร์ด (ไม่พบชื่อตัวเลือกใน dropdown)')
        else:
            raise TimeoutError(f'ไม่พบตัวเลือก {value} ใน dropdown')
        time.sleep(.3)
        current = self.combo_value(combo)
        if current and current != value:
            raise RuntimeError(f'เลือก {value} ไม่สำเร็จ (ได้ "{current}")')

    def choose_combo_background(self, combo, value, options):
        """Lyche's condition items expose no names (live runs fell back to arrow keys), so pick
        by position: expand, Select() the n-th list item, collapse. The added Filter label is
        checked for not(...) afterwards, which catches a wrong pick."""
        # Live finding: the open drop-down exposes no items to UIA (and no Value/Name), so keys are
        # posted to the off-screen dialog instead of typed: focus the combo, open it, Home, Down×n, Enter.
        self.checkpoint()
        dialog_hwnd = combo.top_level_parent().handle
        hold = self.control_dir / 'hold-focus'  # tells the app's focus guard to wait (see App.keep_focus)
        hold.touch()
        user_window = user_foreground(self.pid)  # handed back right after the keys
        held_since = time.monotonic()
        try:
            # Posted keys only reach WPF once the dialog is really active and the combo holds keyboard
            # focus; posting earlier was dropped (first attempts failed live). Wait for both, then use
            # the closed combo's own arrow-key selection (no popup): Home, Down × index.
            element = combo.element_info.element
            ready = False
            for _ in range(4):  # keep the user's keyboard away only briefly
                force_foreground(dialog_hwnd)
                element.SetFocus()
                time.sleep(.1)
                foreground = ctypes.windll.user32.GetForegroundWindow()
                if foreground and win32process.GetWindowThreadProcessId(foreground)[1] == self.pid:
                    ready = True  # Lyche is active (the exact HWND can be a WPF child/owner)
                    break
            if not ready:
                self.emit('log', text='หน้าต่าง Filter ยังไม่ได้โฟกัส — ลองส่งปุ่มต่อ')
            post = ctypes.windll.user32.PostMessageW
            def key(vk):
                post(dialog_hwnd, 0x0100, vk, 0x00000001)  # WM_KEYDOWN
                post(dialog_hwnd, 0x0101, vk, 0xC0000001)  # WM_KEYUP
                time.sleep(.08)
            key(0x24)  # HOME
            for _ in range(options.index(value)):
                key(0x28)  # DOWN
            time.sleep(.2)
        finally:
            give_back_foreground(user_window)
            hold.unlink(missing_ok=True)
            if os.environ.get('LYCHE_FOCUS_DEBUG'):
                self.emit('log', text=f'focus hold Exclude: {(time.monotonic() - held_since) * 1000:.0f} ms')
        try:
            if combo.iface_expand_collapse.CurrentExpandCollapseState:
                combo.iface_expand_collapse.Collapse()
        except Exception:
            pass
        self.emit('log', text=f'เลือก {value} (เบื้องหลัง — ตรวจจากป้ายเงื่อนไขหลังเพิ่ม)')

    def set_filter(self, expression, expected_base=''):
        parsed = parse_filter(expression)
        root = self.settings()
        if self.background:
            # The split button's own Invoke opens Filter Settings (verified live); its drop-down
            # menu would not open while Lyche is minimised.
            if self.hider:
                self.hider.burst()
            self.find(root, aid='cmbFilter').iface_invoke.Invoke()
        else:
            self.menu('cmbFilter', 'Filter Settings')
        dialog = self.wait(lambda: self.dialog('Filter Condition Settings'), 'หน้าต่าง Filter')
        # Clear every existing row before selecting a new condition.
        for button in self.all(dialog, aid='btnClear'):
            if button.is_enabled():
                self.click(button)
        if any(x.is_enabled() for x in self.all(dialog, aid='btnClear')):
            raise RuntimeError('ล้าง Filter เดิมไม่สำเร็จ')
        if parsed:
            variable, condition, code = parsed
            self.checkpoint()
            self.find(dialog, aid='txtKeyWord').set_edit_text(variable)
            self.click(self.find(dialog, aid='btnSearch'))
            def variable_row():
                grid = self.find(dialog, aid='treeListViewQ')
                for row in self.all(grid, kind=50029):
                    values = [x.window_text().strip() for x in self.all(row, kind=50020)]
                    if any(v.split(' ', 1)[0].casefold() == variable.casefold() for v in values):
                        return row
                return None
            row = self.wait(variable_row, variable)
            self.checkpoint()
            row.select()
            def code_row():
                grid = self.find(dialog, aid='treeListViewC')
                for page in self.scan_pages(grid, lambda: self.all(grid, kind=50029)):
                    for item in page:
                        values = [x.window_text().strip() for x in self.all(item, kind=50020)]
                        if values and values[0] == code:
                            return item
                return None
            category = self.wait(code_row, f'code {code}')
            self.checkpoint()
            category.select()
            attempts = 2 if self.background else 1
            for attempt in range(attempts):
                if condition == 'Exclude':  # Include is Lyche's default; only ^= needs the dropdown
                    self.choose_combo(dialog, 'comboBoxCondType', condition,
                                      ('Include', 'Exclude', 'Include Not Applicable', 'Exclude Not Applicable'))
                arrow = self.find(dialog, aid='btnSelect')
                self.click(arrow)
                first = arrow.parent()
                self.wait(lambda: self.find(first, aid='btnClear').is_enabled(), 'เพิ่ม Filter')
                label = ' '.join(x.window_text() for x in self.all(first, kind=50020))
                if variable.casefold() not in label.casefold():
                    raise RuntimeError('ตรวจชื่อ Filter ที่เลือกไม่ได้')
                if ('not(' in label.replace(' ', '').lower()) == (condition == 'Exclude'):
                    break
                if attempt == attempts - 1:
                    raise RuntimeError('Include/Exclude ไม่ตรงกับรายการรัน')
                # Background key posting missed: clear the wrong condition and pick again.
                self.emit('log', text='เงื่อนไข Include/Exclude ไม่ตรง — ล้างแล้วเลือกใหม่')
                self.click(self.find(first, aid='btnClear'))
                self.wait(lambda: not self.find(first, aid='btnClear').is_enabled(), 'ล้างเงื่อนไข')
                category.select()
        self.click(self.find(dialog, aid='rbnOK'))
        self.wait(lambda: not self.dialog('Filter Condition Settings'), 'Done Filter')
        root = self.main()
        filter_text = self.text(root, 'lblFilterValue')
        if not parsed and filter_text.lower() != 'off':
            raise RuntimeError('Filter ยังไม่เป็น Off')
        if parsed and filter_text.lower() == 'off':
            raise RuntimeError('Filter ยังไม่เปิดใช้งาน')
        sample = self.text(root, 'lblSample')
        digits = re.search(r'([\d,]+)\s+respondents', sample)
        actual = int(digits.group(1).replace(',', '')) if digits else None
        if expected_base and actual != int(expected_base):
            raise RuntimeError(f'Base ไม่ตรง: ต้องการ {expected_base}, อ่านได้ {actual}')
        self.emit('log', text=f'Filter {expression or "-"} | {sample}')
        return actual  # respondents after the filter (shown as the row's Base)

    def fill_filename(self, save, target: str):
        """Type the full path like a user; the Save As dialog ignored ValuePattern-only changes."""
        host = self.find(save, aid='FileNameControlHost', required=False)
        field = (self.find(host, kind=50004, required=False) if host else None) \
            or self.find(save, name='File name:', kind=50004, required=False) \
            or self.find(save, aid='1001', kind=50004)
        value = lambda: field.iface_value.CurrentValue.strip()
        self.checkpoint()
        if not self.background:
            field.top_level_parent().set_focus()  # Preview: show the dialog; the text is set below
        # Both modes: no keyboard typing. Preview used to type the path with send_keys; live
        # (2026-10-01) a key was lost mid-path on the 2nd row and the fallback SetValue appended
        # the path to the half-typed text. Send WM_CHAR to the Win32 edit instead — WM_SETTEXT /
        # ValuePattern alone change the text but the Save As dialog ignores them and saves the
        # default name.
        hwnd = field.handle
        if not hwnd:
            raise RuntimeError('ไม่พบช่อง File name แบบ Win32 ใน Save As')
        # Per-character WM_CHAR mangled Thai names (a stray character was appended on Save), so
        # set all but the last character with WM_SETTEXT and "type" only the last one (the 'x'
        # of .xlsx) — that keystroke is what makes the dialog pick up the edited text.
        send = ctypes.windll.user32.SendMessageW
        for _ in range(3):  # the last WM_CHAR was once dropped live (".xls"): check and redo
            send(hwnd, 0x000C, 0, ctypes.c_wchar_p(target[:-1]))  # WM_SETTEXT
            send(hwnd, 0x00B1, len(target) - 1, len(target) - 1)  # EM_SETSEL caret at end
            time.sleep(.1)
            send(hwnd, 0x0102, ord(target[-1]), 0)  # WM_CHAR
            time.sleep(.3)
            if value() == target:
                break
        if value() != target:
            raise RuntimeError(f'กรอกชื่อไฟล์ใน Save As ไม่สำเร็จ (ได้ "{value()}") — ยังไม่กด Save')
        self.emit('log', text=f'Save As → {target}')

    def run_export(self, destination: Path, one_sheet=False):
        """Tabulate All, then export and save to `destination`. Export (the row's last Step):
        - separated sheets (default): Cross → Export with Analysis Axis; Matrix → Export all →
          Excel(Separated Sheets) + Export All Setting;
        - one_sheet: Cross and Matrix → Export all → Excel(One Sheet) (+ 'Some tables could not be
          generated' → Yes, if asked) + Export All Setting.
        In the background Windows' System Sounds are muted meanwhile (user request: Lyche's message
        boxes ding although they are kept off-screen) and given back right after."""
        if not self.background:
            return self._run_export(destination, one_sheet)
        from sounds import system_sounds_muted
        with system_sounds_muted(self.control_dir.parent / 'sounds-muted'):
            return self._run_export(destination, one_sheet)

    def _run_export(self, destination: Path, one_sheet=False):
        if destination.exists():
            raise FileExistsError(f'ไม่เขียนทับไฟล์เดิม: {destination.name}')
        self.matrix = self.is_matrix()  # the result window's title no longer tells Matrix from Cross
        export_all = 'Excel(One Sheet)' if one_sheet else ('Excel(Separated Sheets)' if self.matrix else None)
        self.menu('rbnRun', 'Tabulate All')
        self.wait(lambda: self.main().window_text().startswith('Tabulation Result'), 'Tabulation Result', self.timeout, poll=.5)
        self.emit('log', text='กำลังคำนวณตาราง รอปุ่ม Export พร้อม')
        def busy(progress):
            if not self.background:
                return progress.is_visible()
            # Off-screen, UIA reports every element as offscreen, so read the bar's value instead.
            try:
                pattern = progress.iface_range_value
                return 0 < pattern.CurrentValue < pattern.CurrentMaximum
            except Exception:
                return False
        def ready():
            root = self.main()
            if export_all:  # Matrix ('Export with Analysis Axis' stays disabled) or One Sheet: Export all
                export = self.find(root, aid='rbnOutputExcel', required=False)
            else:
                export = self.find(root, name='Export with Analysis Axis', kind=50000, required=False)
            progress = self.all(root, kind=50012)
            return export if export and export.is_enabled() and not any(busy(p) for p in progress) else None
        export = self.wait(ready, 'คำนวณตารางเสร็จ', self.timeout, poll=.5)  # long wait: poll gently
        def open_prompt():
            self.ack_copy_warning()
            for prefix in ('Question', 'Confirm', 'Information', 'Lyche'):
                dialog = self.dialog(prefix)
                if dialog:
                    texts = ' '.join(c.window_text() for c in self.all(dialog, kind=50020))
                    if 'Open file?' in texts:
                        return dialog
            return None
        # Export all (Matrix / One Sheet) in the background: keep the app's screen freeze on while Lyche
        # pops its boxes (user request: no flashes) — from the menu click to the Save As dialog, then
        # again around Save until 'Do not copy or paste…' is answered. Nested freezes are one freeze.
        frozen = self.frozen_screen if (export_all and self.background) else nullcontext
        with frozen():
            # In the background an Invoke that lands while Lyche is still finishing can be ignored:
            # press again (up to 3 times) until "Open file?" appears. A modal prompt blocks repeats.
            for attempt in range(3 if self.background else 1):
                try:
                    if export_all:
                        # User's flow (2026-10-01): Export all → the item, the Export All Setting ticks,
                        # OK; from 'Open file?' on it is the same as Export with Analysis Axis.
                        self.wait(ready, 'คำนวณตารางเสร็จ', self.timeout, poll=.5)
                        self.export_all(export_all)
                    else:
                        self.click(self.wait(ready, 'คำนวณตารางเสร็จ', self.timeout, poll=.5))
                    prompt = self.wait(open_prompt, 'Open file?', 15 if self.background else 20)
                    break
                except TimeoutError:
                    if attempt == (2 if self.background else 0):
                        raise
            no = self.find(prompt, name='No', kind=50000, required=False) or self.find(prompt, name='&No', kind=50000, required=False)
            if no is None:
                raise RuntimeError('ไม่พบปุ่ม No ใน Open file?')
            self.click(no)
            save = self.wait(lambda: (self.ack_copy_warning(), self.dialog('Save As'))[1], 'Save As')
        self.fill_filename(save, str(destination.resolve()))
        if destination.exists():
            raise FileExistsError('มีไฟล์ถูกสร้างขึ้นระหว่างรัน หยุดก่อนบันทึกทับ')
        save_button = self.find(save, aid='1', kind=50000, required=False)
        with frozen():
            self.click(save_button or self.find(save, name='Save', kind=50000))
            if export_all and self.background:  # One Sheet: 'Do not copy or paste…' follows within ~1 s
                end = time.monotonic() + 3
                while time.monotonic() < end and not self.ack_copy_warning() and not self.saved_dialogs():
                    time.sleep(.1)
        saved = self.wait(lambda: (self.ack_copy_warning(), (self.saved_dialogs() or [None])[0])[1],
                          'Saved.', self.timeout, poll=.5)
        self.close_saved(saved)
        self.settings()

    def export_all(self, item):
        """Export all → `item` ('Excel(Separated Sheets)' / 'Excel(One Sheet)'); answer Lyche's
        'Some tables could not be generated … export only the tables that were successfully
        generated?' with Yes when it is asked (user rule); then the Export All Setting ticks + OK."""
        opened = ('Export All Setting', 'Warning')
        if self.background:  # these gallery items react to a real click only (see menu_click_on_screen)
            self.menu_click_on_screen('rbnOutputExcel', item, expect=opened)
        else:
            self.menu('rbnOutputExcel', item)
        def next_dialog():
            if self.ack_copy_warning():
                return None
            warning = self.dialog('Warning')
            if warning:
                text = ' '.join(x.window_text() for x in self.all(warning, kind=50020)).strip()
                if 'Some tables could not be generated' not in text:
                    raise ValueError(f'Lyche ถามคำถามที่ไม่รู้จัก: {text}')
                yes = self.find(warning, name='Yes', kind=50000, required=False) \
                    or self.find(warning, name='&Yes', kind=50000)
                self.click(yes)
                self.emit('log', text='Some tables could not be generated → กด Yes (export เฉพาะตารางที่สร้างได้)')
                return None
            return self.dialog('Export All Setting')
        setting = self.wait(next_dialog, 'Export All Setting', 15 if self.background else 20)
        self.export_all_setting(setting, item)

    def ack_copy_warning(self):
        """Lyche's 'Warning - * Do not copy or paste while in process.' (seen during the One Sheet export)
        is modal and, parked off-screen, held the export until our time-out (the file was written the
        moment the box closed). User rule: press OK. True when one was dismissed."""
        for handle in self.visible_dialogs('Warning'):
            dialog = self.wrapper(handle)
            text = ' '.join(x.window_text() for x in self.all(dialog, kind=50020))
            if 'Do not copy or paste while in process' in text:
                self.press(self.find(dialog, name='OK', kind=50000, required=False)
                           or self.find(dialog, name='&OK', kind=50000))
                self.emit('log', text='Warning: Do not copy or paste while in process → กด OK')
                return True
        return False

    def export_all_setting(self, dialog, item='Excel(Separated Sheets)'):
        """'Export All Setting': set the ticks as the user's picture (EXPORT_ALL_TICKS), check them,
        then OK. Toggle pattern only, so it also works off-screen."""
        for aid, want in EXPORT_ALL_TICKS.items():
            box = self.find(dialog, aid=aid)
            for _ in range(2):
                if bool(box.get_toggle_state()) == want:
                    break
                self.checkpoint()
                box.toggle()
                time.sleep(.1)
            if bool(box.get_toggle_state()) != want:
                raise RuntimeError(f'ตั้งค่า Export All Setting ไม่สำเร็จ: {box.window_text().strip() or aid}')
        self.checkpoint()
        self.click(self.find(dialog, aid='btnOk', kind=50000, required=False) or self.find(dialog, name='OK', kind=50000))
        self.emit('log', text=f'Export all → {item} · ติ๊ก Item ID, banner/stub number, '
                              'Table (ไม่เอา Information, Graph)')

    def close_leftover_dialogs(self):
        """After a failed background step, close the Lyche dialogs we opened (they are off-screen and
        modal, so Lyche would otherwise look frozen): Lyche Error boxes by OK, Filter/Import History
        without applying, Save As by Cancel, Saved. by OK, Confirm by No (Yes for 'Changes will be
        lost', which only discards the unapplied Filter)."""
        closers = (('Error', 'OK'), ('Confirm -', 'No'), ('Information', 'OK'), ('Question', 'No'),
                   ('Save As', 'Cancel'), ('Export All Setting', 'Cancel'), ('Warning', 'No'),
                   ('Filter Condition Settings', None),
                   ('Tabulation List -', None))
        for _ in range(4):
            left = False
            for prefix, button in closers:
                for handle in self.visible_dialogs(prefix):  # visible only: owned dialogs hide while the owner is minimised
                    left = True
                    dialog = self.wrapper(handle)
                    choice = button
                    if prefix == 'Confirm -':
                        text = ' '.join(x.window_text() for x in self.all(dialog, kind=50020)).strip()
                        if text == 'Changes will be lost. Proceed?':
                            choice = 'Yes'
                    try:
                        target = self.find(dialog, name=choice, kind=50000, required=False) if choice else None
                        if target:
                            self.press(target)
                        else:
                            close = self.find(dialog, aid='PART_CloseWindowButton', required=False)
                            (self.press(close) if close else dialog.iface_window.Close())
                    except Exception:
                        win32gui.PostMessage(handle, 0x0010, 0, 0)
                    time.sleep(.5)
            if not left:
                return
        self.emit('log', text='ยังมีหน้าต่าง Lyche ค้างหลังเกิดข้อผิดพลาด กรุณาตรวจใน Lyche')

    def close_window(self):
        """Close the Cross Tabulation window once the queue is done (user request). Lyche asks
        'Change will not be saved. Proceed?' (verified live) → Yes, as already approved; any other
        prompt stops here and is reported."""
        handle = self.handle
        if not win32gui.IsWindow(handle):
            return
        if win32gui.GetWindowText(handle).startswith('Tabulation Result'):
            # Closing from the result screen asks a different question; go back first.
            self.settings()
        if self.hider:
            self.hider.burst(5)
        try:
            self.wrapper(handle).iface_window.Close()
        except Exception:
            win32gui.PostMessage(handle, 0x0010, 0, 0)  # WM_CLOSE
        end = time.monotonic() + 20
        while time.monotonic() < end:
            if not win32gui.IsWindow(handle):
                self.emit('log', text='ปิดหน้าต่าง Cross Tabulation แล้ว')
                return
            confirm = self.dialog('Confirm -', visible_only=False)
            if confirm:
                text = ' '.join(x.window_text() for x in self.all(confirm, kind=50020)).strip()
                if text != 'Change will not be saved. Proceed?':
                    raise ValueError(f'ปิด Cross Tabulation ไม่ได้ — พบคำถามที่ไม่รู้จัก: {text}')
                self.press(self.find(confirm, name='Yes', kind=50000, required=False) or self.find(confirm, name='&Yes', kind=50000))
            time.sleep(.2)
        raise RuntimeError('ปิดหน้าต่าง Cross Tabulation ไม่สำเร็จ')

    def saved_dialogs(self):
        """Every 'Saved.' message box of this Lyche process (a stale one must not block the run)."""
        found = []
        def collect(handle, _):
            if win32gui.IsWindowVisible(handle) and win32process.GetWindowThreadProcessId(handle)[1] == self.pid \
                    and win32gui.GetWindowText(handle).startswith('Information'):
                found.append(handle)
        win32gui.EnumWindows(collect, None)
        result = []
        for handle in found:
            dialog = self.wrapper(handle)
            if any(x.window_text().strip() == 'Saved.' for x in self.all(dialog, kind=50020)):
                result.append(dialog)
        return result

    def close_saved(self, dialog):
        """Press OK and make sure the box is really gone: in the background an early Invoke was
        ignored and the boxes piled up (live finding), so retry, then Close, then WM_CLOSE."""
        handle = dialog.handle
        def ok():
            self.find(dialog, name='OK', kind=50000).iface_invoke.Invoke()
        for attempt in (ok, ok, lambda: dialog.iface_window.Close(),
                        lambda: win32gui.PostMessage(handle, 0x0010, 0, 0)):
            self.checkpoint()
            try:
                attempt()
            except Exception:
                pass
            end = time.monotonic() + 3
            while time.monotonic() < end:
                if not win32gui.IsWindow(handle) or not win32gui.IsWindowVisible(handle):
                    return
                time.sleep(.1)
        raise RuntimeError('ปิดกล่อง Saved. ของ Lyche ไม่สำเร็จ')


# ====== MODULE: sounds ======
"""Silence Windows' "System Sounds" for a moment (user request, 2026-10-01).

Lyche's message boxes (Some tables could not be generated…, Do not copy or paste…, Saved.) play the
Windows warning/information sound even while the bot keeps them off-screen. During a background export
the worker mutes the System Sounds audio session (only that session: music/video keep playing) and
restores the previous state right after. The previous state is also written to DATA/sounds-muted, so
the app can restore it if the worker ends abruptly (App.finished, and at the next start).

Core Audio through comtypes only (no extra package). Every failure means "could not mute" and is
ignored: muting is a nicety, never a reason to stop a run.
"""
# from __future__ import annotations  (applied to this section by the loader)

from contextlib import contextmanager
from ctypes import POINTER, c_float, c_int, c_void_p, wintypes
from pathlib import Path

import comtypes
from comtypes import COMMETHOD, GUID, HRESULT, IUnknown

CLSID_MMDeviceEnumerator = GUID('{BCDE0395-E52F-467C-8E3D-C4579291692E}')


class ISimpleAudioVolume(IUnknown):
    _iid_ = GUID('{87CE5498-68D6-44E5-9215-6DA47EF883D8}')
    _methods_ = [
        COMMETHOD([], HRESULT, 'SetMasterVolume', (['in'], c_float, 'level'), (['in'], POINTER(GUID), 'context')),
        COMMETHOD([], HRESULT, 'GetMasterVolume', (['out', 'retval'], POINTER(c_float), 'level')),
        COMMETHOD([], HRESULT, 'SetMute', (['in'], wintypes.BOOL, 'mute'), (['in'], POINTER(GUID), 'context')),
        COMMETHOD([], HRESULT, 'GetMute', (['out', 'retval'], POINTER(wintypes.BOOL), 'mute')),
    ]


class IAudioSessionControl(IUnknown):
    _iid_ = GUID('{F4B1A599-7266-4319-A8CA-E70ACB11E8CD}')
    _methods_ = [
        COMMETHOD([], HRESULT, 'GetState', (['out', 'retval'], POINTER(c_int), 'state')),
        COMMETHOD([], HRESULT, 'GetDisplayName', (['out', 'retval'], POINTER(wintypes.LPWSTR), 'name')),
        COMMETHOD([], HRESULT, 'SetDisplayName', (['in'], wintypes.LPCWSTR, 'name'), (['in'], POINTER(GUID), 'context')),
        COMMETHOD([], HRESULT, 'GetIconPath', (['out', 'retval'], POINTER(wintypes.LPWSTR), 'path')),
        COMMETHOD([], HRESULT, 'SetIconPath', (['in'], wintypes.LPCWSTR, 'path'), (['in'], POINTER(GUID), 'context')),
        COMMETHOD([], HRESULT, 'GetGroupingParam', (['out', 'retval'], POINTER(GUID), 'grouping')),
        COMMETHOD([], HRESULT, 'SetGroupingParam', (['in'], POINTER(GUID), 'grouping'), (['in'], POINTER(GUID), 'context')),
        COMMETHOD([], HRESULT, 'RegisterAudioSessionNotification', (['in'], c_void_p, 'events')),
        COMMETHOD([], HRESULT, 'UnregisterAudioSessionNotification', (['in'], c_void_p, 'events')),
    ]


class IAudioSessionControl2(IAudioSessionControl):
    _iid_ = GUID('{BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}')
    _methods_ = [
        COMMETHOD([], HRESULT, 'GetSessionIdentifier', (['out', 'retval'], POINTER(wintypes.LPWSTR), 'id')),
        COMMETHOD([], HRESULT, 'GetSessionInstanceIdentifier', (['out', 'retval'], POINTER(wintypes.LPWSTR), 'id')),
        COMMETHOD([], HRESULT, 'GetProcessId', (['out', 'retval'], POINTER(wintypes.DWORD), 'pid')),
        COMMETHOD([], HRESULT, 'IsSystemSoundsSession'),
        COMMETHOD([], HRESULT, 'SetDuckingPreference', (['in'], wintypes.BOOL, 'opt_out')),
    ]


class IAudioSessionEnumerator(IUnknown):
    _iid_ = GUID('{E2F5BB11-0570-40CA-ACDD-3AA01277DEE8}')
    _methods_ = [
        COMMETHOD([], HRESULT, 'GetCount', (['out', 'retval'], POINTER(c_int), 'count')),
        COMMETHOD([], HRESULT, 'GetSession', (['in'], c_int, 'index'),
                  (['out', 'retval'], POINTER(POINTER(IAudioSessionControl)), 'session')),
    ]


class IAudioSessionManager2(IUnknown):
    _iid_ = GUID('{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}')
    _methods_ = [  # IAudioSessionManager, then IAudioSessionManager2
        COMMETHOD([], HRESULT, 'GetAudioSessionControl', (['in'], POINTER(GUID), 'session'), (['in'], wintypes.DWORD, 'flags'),
                  (['out', 'retval'], POINTER(POINTER(IAudioSessionControl)), 'control')),
        COMMETHOD([], HRESULT, 'GetSimpleAudioVolume', (['in'], POINTER(GUID), 'session'), (['in'], wintypes.DWORD, 'flags'),
                  (['out', 'retval'], POINTER(POINTER(ISimpleAudioVolume)), 'volume')),
        COMMETHOD([], HRESULT, 'GetSessionEnumerator',
                  (['out', 'retval'], POINTER(POINTER(IAudioSessionEnumerator)), 'sessions')),
        COMMETHOD([], HRESULT, 'RegisterSessionNotification', (['in'], c_void_p, 'notification')),
        COMMETHOD([], HRESULT, 'UnregisterSessionNotification', (['in'], c_void_p, 'notification')),
        COMMETHOD([], HRESULT, 'RegisterDuckNotification', (['in'], wintypes.LPCWSTR, 'session'), (['in'], c_void_p, 'notification')),
        COMMETHOD([], HRESULT, 'UnregisterDuckNotification', (['in'], c_void_p, 'notification')),
    ]


class IMMDevice(IUnknown):
    _iid_ = GUID('{D666063F-1587-4E43-81F1-B948E807363F}')
    _methods_ = [
        COMMETHOD([], HRESULT, 'Activate', (['in'], POINTER(GUID), 'iid'), (['in'], wintypes.DWORD, 'clsctx'),
                  (['in'], c_void_p, 'params'), (['out', 'retval'], POINTER(POINTER(IUnknown)), 'interface')),
        COMMETHOD([], HRESULT, 'OpenPropertyStore', (['in'], wintypes.DWORD, 'access'), (['out', 'retval'], POINTER(c_void_p), 'store')),
        COMMETHOD([], HRESULT, 'GetId', (['out', 'retval'], POINTER(wintypes.LPWSTR), 'id')),
        COMMETHOD([], HRESULT, 'GetState', (['out', 'retval'], POINTER(wintypes.DWORD), 'state')),
    ]


class IMMDeviceEnumerator(IUnknown):
    _iid_ = GUID('{A95664D2-9614-4F35-A746-DE8DB63617E6}')
    _methods_ = [
        COMMETHOD([], HRESULT, 'EnumAudioEndpoints', (['in'], c_int, 'flow'), (['in'], wintypes.DWORD, 'mask'),
                  (['out', 'retval'], POINTER(c_void_p), 'devices')),
        COMMETHOD([], HRESULT, 'GetDefaultAudioEndpoint', (['in'], c_int, 'flow'), (['in'], c_int, 'role'),
                  (['out', 'retval'], POINTER(POINTER(IMMDevice)), 'device')),
        COMMETHOD([], HRESULT, 'GetDevice', (['in'], wintypes.LPCWSTR, 'id'), (['out', 'retval'], POINTER(POINTER(IMMDevice)), 'device')),
        COMMETHOD([], HRESULT, 'RegisterEndpointNotificationCallback', (['in'], c_void_p, 'client')),
        COMMETHOD([], HRESULT, 'UnregisterEndpointNotificationCallback', (['in'], c_void_p, 'client')),
    ]


def _system_sounds_volumes():
    """ISimpleAudioVolume of the System Sounds session on the default playback devices."""
    try:
        comtypes.CoInitialize()  # harmless when this thread already has COM (S_FALSE / other apartment)
    except OSError:
        pass
    enumerator = comtypes.CoCreateInstance(CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER)
    volumes, seen = [], set()
    for role in (0, 1):  # eConsole, eMultimedia (usually the same device)
        try:
            device = enumerator.GetDefaultAudioEndpoint(0, role)  # eRender
        except comtypes.COMError:
            continue
        device_id = device.GetId()
        if device_id in seen:
            continue
        seen.add(device_id)
        manager = device.Activate(IAudioSessionManager2._iid_, comtypes.CLSCTX_ALL, None).QueryInterface(IAudioSessionManager2)
        sessions = manager.GetSessionEnumerator()
        for index in range(sessions.GetCount()):
            control = sessions.GetSession(index).QueryInterface(IAudioSessionControl2)
            if control.IsSystemSoundsSession() == 0:  # S_OK = the System Sounds session (S_FALSE otherwise)
                volumes.append(control.QueryInterface(ISimpleAudioVolume))
    return volumes


def set_system_sounds_mute(mute: bool):
    """Mute or unmute System Sounds; returns the previous mute state, or None if it could not be done."""
    try:
        volumes = _system_sounds_volumes()
        if not volumes:
            return None
        previous = bool(volumes[0].GetMute())
        for volume in volumes:
            volume.SetMute(bool(mute), None)
        return previous
    except Exception:
        return None


@contextmanager
def system_sounds_muted(marker: Path):
    """Mute System Sounds for the block; `marker` records the previous state for crash recovery."""
    previous = set_system_sounds_mute(True)
    if previous is False:  # was audible: remember to give it back even if this process dies
        try:
            marker.write_text('0', encoding='utf-8')
        except OSError:
            pass
    try:
        yield
    finally:
        if previous is False:
            set_system_sounds_mute(False)
            marker.unlink(missing_ok=True)


def restore_if_left_muted(marker: Path):
    """App side: a worker that died while muting left `marker` — give the sound back."""
    if marker.exists():
        if marker.read_text(encoding='utf-8', errors='ignore').strip() == '0':
            set_system_sounds_mute(False)
        marker.unlink(missing_ok=True)


# ====== MODULE: post_total_na ======
# Delete Total + NA — processing core copied unchanged from 156_DeleteTotalNA.py (lines 1-883),
# without its PyQt6 window. Auto Lychee calls process_in_place() after each exported Banner.
# Keep the copied part in sync with 156_DeleteTotalNA.py if that script changes.
"""Delete Total + NA Batch ? standalone Python application.

Copy this file alone to another folder or Windows computer.
Requires Python 3.10+, Microsoft Excel Desktop, PyQt6 and pywin32.
Install dependencies once: python -m pip install "PyQt6>=6.6,<7" "pywin32>=306"
Run: python DeleteTotalNA.py

Includes processing, GUI, styles, and the application icon. No local imports,
assets folder, or requirements.txt is needed to run after installing dependencies.
"""
# from __future__ import annotations  (applied to this section by the loader)

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
    na_rows_deleted: int = 0
    empty_rows_deleted: int = 0

    def summary(self) -> str:
        return (
            f"{self.sheets_scanned} ชีต · ยกเลิก Merge {self.merged_areas_changed} จุด · "
            f"ลบ TOTAL {self.total_columns_deleted} คอลัมน์ · "
            f"ลบ NA {self.na_columns_deleted} คอลัมน์ · {self.na_rows_deleted} แถว · "
            f"ลบแถวไม่มี Label/Frequency {self.empty_rows_deleted} แถว"
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
    header_row = _find_total_row(ws)
    if header_row and first >= header_row:
        # A merged statistic (e.g. Mean) is data, not a banner header.
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
    header_row = _find_total_row(ws)
    if header_row and first_merged_row >= header_row:
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


def _delete_na_rows(ws: Any) -> int:
    """Delete NA stub rows, including every row of a vertically merged label."""
    last_row = int(ws.Range(f"B{int(ws.Rows.Count)}").End(XL_UP).Row)
    raw = ws.Range(f"B1:B{last_row}").Value
    values = ((raw,),) if last_row == 1 else raw
    rows: set[int] = set()
    for row, (value,) in enumerate(values, start=1):
        if _cell_text(value).strip().upper() != "NA":
            continue
        cell = ws.Range(f"B{row}")
        if bool(cell.MergeCells):
            area = cell.MergeArea
            first = int(area.Row)
            end = first + int(area.Rows.Count) - 1
        else:
            first = end = row
        rows.update(range(first, end + 1))

    # Work upwards so earlier row numbers remain valid after each deletion.
    for first, end in reversed(_contiguous_runs(sorted(rows))):
        ws.Range(f"{first}:{end}").Delete()
    return len(rows)


def _delete_empty_label_frequency_rows(
    ws: Any, header_row: int, merge_addresses: list[str] | None = None,
    on_progress: Callable[[str], None] | None = None,
) -> int:
    """Remove empty data rows while preserving merged records and annotations."""
    if not header_row:
        return 0
    # UsedRange can include formatting down to row 1,048,576. Find actual
    # content (including formulas) instead of scanning those empty cells.
    find_options = dict(What="*", After=ws.Range("A1"), LookIn=-4123,
                        LookAt=2, SearchDirection=2, MatchCase=False,
                        SearchFormat=False)
    last_cell = ws.Cells.Find(SearchOrder=1, **find_options)
    if last_cell is None:
        return 0
    last_row = int(last_cell.Row)
    last_col = max(3, int(ws.Cells.Find(SearchOrder=2, **find_options).Column))
    intervals: list[tuple[int, int]] = []
    if merge_addresses is not None:
        for address in merge_addresses:
            _, first, _, _ = _address_span(address)
            end = int(_ADDRESS_PART.fullmatch(address.split(":")[-1]).group(2))
            if end > first:
                intervals.append((first, end))
        # Combine overlapping vertical merges; adjacent records stay separate.
        combined: list[tuple[int, int]] = []
        for first, end in sorted(intervals):
            if combined and first <= combined[-1][1]:
                combined[-1] = (combined[-1][0], max(end, combined[-1][1]))
            else:
                combined.append((first, end))
        intervals = combined
        for first, end in intervals:
            if first <= last_row <= end:
                last_row = end
    else:
        # Binary files use COM discovery, including the last merged tail.
        for col in _merged_columns_in_row(ws, last_row, last_col):
            area = ws.Range(_cell_ref(last_row, col)).MergeArea
            last_row = max(last_row, int(area.Row) + int(area.Rows.Count) - 1)
    rows: list[int] = []

    def empty_or_zero(value: Any) -> bool:
        if value is None or (isinstance(value, str) and not value.strip()):
            return True
        if isinstance(value, bool):
            return False
        try:
            return float(value) == 0
        except (TypeError, ValueError):
            return False

    for start in range(header_row + 1, last_row + 1, 256):
        end = min(start + 255, last_row)
        if on_progress is not None:
            on_progress(f"ตรวจ Label/Frequency แถว {start}–{end}/{last_row}")
        values = ws.Range(f"A{start}:{_column_letter(last_col)}{end}").Value
        for row, record in enumerate(values, start):
            if _cell_text(record[1]).strip():
                continue
            # Keep notes in the code column; numeric category codes are allowed.
            code = _cell_text(record[0]).strip()
            if code:
                try:
                    float(code)
                except ValueError:
                    continue
            if not all(empty_or_zero(value) for value in record[2:]):
                continue
            rows.append(row)
    candidates = set(rows)
    if merge_addresses is not None:
        # All merge geometry is already in memory: no per-cell COM calls.
        for first, end in intervals:
            if first > last_row:
                break
            if not all(row in candidates for row in range(first, end + 1)):
                candidates.difference_update(range(first, end + 1))
        rows = sorted(candidates)
        for first, end in reversed(_contiguous_runs(rows)):
            if on_progress is not None:
                on_progress(f"ลบแถว {first}–{end}")
            ws.Range(f"{first}:{end}").Delete()
        return len(rows)
    rows = []
    visited: set[int] = set()
    for row in sorted(candidates):
        if row in visited:
            continue
        # Treat all rows connected by vertical merges as one record. Inspect
        # its head AND continuation rows, including significance annotations.
        group: set[int] = set()
        pending = [row]
        seen_areas: set[str] = set()
        while pending:
            current = pending.pop()
            if current in group:
                continue
            group.add(current)
            for col in _merged_columns_in_row(ws, current, last_col):
                area = ws.Range(_cell_ref(current, col)).MergeArea
                address = str(area.Address)
                if address in seen_areas:
                    continue
                seen_areas.add(address)
                first = int(area.Row)
                end = first + int(area.Rows.Count) - 1
                pending.extend(r for r in range(first, end + 1) if r not in group)
        visited.update(group)
        if group <= candidates:
            rows.extend(group)
    for first, end in reversed(_contiguous_runs(rows)):
        ws.Range(f"{first}:{end}").Delete()
    return len(rows)


def _normalize_table_spacing(ws: Any) -> None:
    """Keep one empty worksheet row before each subsequent exported table."""
    last_row = int(ws.Range(f"A{int(ws.Rows.Count)}").End(XL_UP).Row)
    if last_row < 2:
        return
    values = ws.Range(f"A1:A{last_row}").Value
    starts = [row for row, (value,) in enumerate(values, 1)
              if _cell_text(value).strip().lower() == "contents"]
    if len(starts) < 2:
        return
    used = ws.UsedRange
    last_col = int(used.Column) + int(used.Columns.Count) - 1
    for start in reversed(starts[1:]):
        previous_start = max(row for row in starts if row < start)
        first_blank = start
        for row in range(start - 1, previous_start, -1):
            span = ws.Range(_row_span_ref(row, 1, last_col))
            raw = span.Value
            record = (raw,) if last_col == 1 else raw[0]
            # Never mistake a blank continuation of Mean/data for a spacer.
            if span.MergeCells is not False or any(
                _cell_text(value).strip() for value in record
            ):
                break
            first_blank = row
        blank_count = start - first_blank
        if blank_count > 1:
            ws.Range(f"{first_blank + 1}:{start - 1}").Delete()
        elif blank_count == 0:
            ws.Range(f"{start}:{start}").Insert()
            ws.Range(f"{start}:{start}").Clear()
            ws.Range(f"{start}:{start}").RowHeight = ws.StandardHeight


def process_workbook(
    excel: Any,
    workbook_path: Path,
    on_sheet: Callable[[float, str], None] | None = None,
    delete_empty_rows: bool = False,
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
            def detail(message: str) -> None:
                if on_sheet is not None:
                    on_sheet(done / steps, f"{ws.Name} · ชีต {position}/{total_sheets} · {message}")

            detail("กำลังตรวจตาราง")
            target_row = _find_total_row(ws)
            if delete_empty_rows:
                result.empty_rows_deleted += _delete_empty_label_frequency_rows(
                    ws, target_row, merge_index.get(ws.Name), detail
                )
            if target_row:
                _add_left_borders_to_previous_row(ws, target_row)
                total_deleted, na_deleted = _delete_extra_total_and_na_columns(
                    ws, target_row
                )
                result.total_columns_deleted += total_deleted
                result.na_columns_deleted += na_deleted
            na_rows_deleted = _delete_na_rows(ws)
            result.na_rows_deleted += na_rows_deleted
            _normalize_table_spacing(ws)
            if target_row or na_rows_deleted:
                # Refresh cross-sheet formulas before inspecting the next sheet.
                detail("กำลังคำนวณสูตร")
                excel.Calculate()
            step(f"ลบคอลัมน์ TOTAL/NA และแถว NA · ชีต {position}/{total_sheets}")

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
    """Plain pywin32 dynamic dispatch (Auto Lychee change, 2026-10-01).

    The original shared interface metadata across objects with a SessionDispatch subclass. Live, a
    second Delete Total + NA in the same process (the post-processing child runs one per queue row)
    failed at `excel.Calculate()` with "'bool' object is not callable" whenever the previous file
    needed no changes (e.g. a Matrix export): the shared metadata leaked between Excel sessions.
    Plain dynamic dispatch gave the same output (every value/style/merge/size compared on a
    200-sheet file) and was not slower (15 s vs 29 s).
    """
    from win32com.client import dynamic

    return dynamic.Dispatch(dispatch)


def process_files(
    jobs: list[dict[str, Any]],
    work_dir: Path,
    progress: Callable[[float, str, str, str], None],
    delete_empty_rows: bool = False,
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

                result = process_workbook(
                    excel, working_copy, on_sheet, delete_empty_rows=delete_empty_rows
                )
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




# ---------------- Auto Lychee integration ----------------

def process_in_place(path: Path, delete_empty_rows: bool = False,
                     progress: Callable[[str], None] | None = None) -> str:
    """Run Delete Total + NA on one exported workbook and save it back to the same path.

    Uses the original batch functions (hidden Microsoft Excel, working copy in a temp folder).
    The original file is replaced only after processing succeeded, so a failure leaves the raw
    Lyche export untouched. Runs in its own thread: the Auto Lychee worker thread already holds
    a multithreaded COM apartment for UI Automation, and Excel automation needs a fresh one.
    """
    import tempfile
    import threading

    path = Path(path)
    validate_input_file(path)
    job: dict[str, Any] = {"path": str(path)}
    failure: list[BaseException] = []

    def report(_count: float, _path: str, _status: str, detail: str) -> None:
        if progress is not None:
            progress(detail)

    with tempfile.TemporaryDirectory(prefix="autolychee_totalna_") as work:
        def run() -> None:
            try:
                process_files([job], Path(work), report, delete_empty_rows=delete_empty_rows)
            except BaseException as exc:  # surfaced to the caller below
                failure.append(exc)

        thread = threading.Thread(target=run, name="delete-total-na", daemon=True)
        thread.start()
        thread.join()
        if failure:
            raise failure[0]
        if job.get("status") != "ready":
            raise RuntimeError(f"Delete Total+NA ไม่สำเร็จ: {job.get('error', 'ไม่ทราบสาเหตุ')}")
        pending = path.with_name(f".{path.stem}.delete_total_na_pending{path.suffix}")
        try:
            shutil.copy2(job["processed_path"], pending)
            os.replace(pending, path)
        finally:
            if pending.exists():
                pending.unlink()
    return job.get("result", "")


# ====== MODULE: post_del_sig ======
# Del Sig — processing core copied unchanged from Del_Sig.py (lines 1-7 and 16-657): only its PyQt6
# imports (lines 8-15) and window are left out. Auto Lychee calls process_in_place() after each
# exported Banner (after Delete Total + NA). Keep in sync with Del_Sig.py if that script changes.
import sys
import os
import re
import bisect
from collections import defaultdict
from datetime import datetime


import openpyxl
from openpyxl.styles import Font, Border, Side
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.cell_range import CellRange


def writable_cell(ws, row, col):
    """Return a writable cell. If (row, col) falls inside a merged range,
    return the top-left anchor of that range (the only writable member)."""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
                return ws.cell(row=mr.min_row, column=mr.min_col)
    return cell


# ===================== Utilities =====================

def clean_header(text):
    if not isinstance(text, str): return ""
    return "".join(re.findall(r"[A-Za-z]", text)).upper()

def parse_sig_groups(sig_input):
    groups = [g.strip().upper() for g in sig_input.split(",") if g.strip()]
    expanded = []
    for g in groups:
        if "-" in g and len(g) == 3:
            a,b = g.split("-")
            expanded.append("".join(chr(c) for c in range(ord(a), ord(b)+1)))
        else:
            expanded.append("".join(sorted(set(re.findall(r"[A-Z]", g)))))
    out = {}
    for rng in expanded:
        for ch in rng:
            out[ch] = rng
    return out

def detect_letter_header_row(ws, start_row, end_row, min_seq=3):
    scan_until = min(end_row, start_row + 120)
    for r in range(start_row, scan_until + 1):
        letters = []
        for c in range(3, ws.max_column + 1):
            t = clean_header(str(ws.cell(row=r, column=c).value))
            if len(t) == 1 and 'A' <= t <= 'Z':
                letters.append((c, t))
        if len(letters) < min_seq:
            continue
        letters.sort()
        vals = {c:v for c,v in letters}
        for c,v in letters:
            if v == 'A' and vals.get(c+1) == 'B' and vals.get(c+2) == 'C':
                return r, c
    return -1, -1

def build_col_letter_map(ws, header_row, start_col):
    mapping = {}
    for c in range(start_col, ws.max_column + 1):
        t = clean_header(str(ws.cell(row=header_row, column=c).value))
        if len(t) == 1 and 'A' <= t <= 'Z':
            mapping[c] = t
        else:
            if mapping: break
    return mapping

def find_total_row_in_colB(ws, header_row, end_row):
    for r in range(header_row, end_row + 1):
        v = ws.cell(row=r, column=2).value
        if isinstance(v, str) and v.strip().upper() == "TOTAL":
            return r
    return -1

def find_last_data_row(ws, start_row, col_indexes):
    last = start_row
    for r in range(start_row, ws.max_row + 1):
        any_val = False
        for c in col_indexes:
            v = ws.cell(row=r, column=c).value
            if v not in (None, "") and str(v).strip() != "":
                any_val = True
                break
        if any_val: last = r
    return last

def filter_sig_text(text, allowed_letters, preserve_tokens=("ADJ",)):
    if not isinstance(text, str) or not text:
        return text
    out = text
    dynamic_preserve = tuple(
        tok for tok in preserve_tokens
        if tok and set(tok.upper()).issubset(set(allowed_letters))
    )
    placeholders = {}
    for idx, tok in enumerate(dynamic_preserve):
        if tok in out:
            ph = f"\u00a7{idx}\u00a7"
            placeholders[ph] = tok
            out = out.replace(tok, ph)
    out = out.replace("_", "")
    keep = set(allowed_letters)
    buf = []
    for ch in out:
        if 'A' <= ch.upper() <= 'Z':
            if ch.upper() in keep:
                buf.append(ch)
        else:
            buf.append(ch)
    out = "".join(buf)
    for ph, tok in placeholders.items():
        out = out.replace(ph, tok)
    return out.strip()


# ===================== Sig Beside Utilities =====================

def collect_sig_rows_and_merge(ws, start_row, end_row, label_cols, data_cols, keep_first_unlabeled=False):
    """
    Rule for Sig-beside mode:
    - If any label column has value -> value row.
    - If all label columns are empty -> paired/sig row.
      Merge only pure-letter cells (A-Z only, no digits) into previous value row,
      then delete that row (to compact layout like target output).
    - keep_first_unlabeled=True is used for N% layout (Count + Column % + Sig):
      keep first unlabeled row (Column %), then process/delete subsequent unlabeled rows.
    Return list of row indexes to delete.
    """
    data_cols = list(data_cols)
    label_cols = list(label_cols)
    rows_to_delete = []
    prev_value_row = None
    merge_target = None
    unlabeled_after_value = 0

    for r in range(start_row, end_row + 1):
        is_value_row = any(
            ws.cell(row=r, column=lc).value not in (None, "") and
            str(ws.cell(row=r, column=lc).value).strip()
            for lc in label_cols
        )

        if is_value_row:
            prev_value_row = r
            merge_target = r
            unlabeled_after_value = 0
            continue

        if prev_value_row is None:
            continue

        unlabeled_after_value += 1
        if keep_first_unlabeled and unlabeled_after_value == 1:
            # N% layout: first unlabeled row is the Column % row, which we keep.
            # The Sig must attach beside the % on THIS row, not the Count row.
            merge_target = r
            continue

        target = merge_target if merge_target is not None else prev_value_row
        for c in data_cols:
            sig_val = ws.cell(row=r, column=c).value
            if (sig_val and isinstance(sig_val, str) and sig_val.strip()
                    and re.search(r'[A-Za-z]', sig_val)
                    and not re.search(r'\d', sig_val)):
                val_cell = writable_cell(ws, target, c)
                existing = val_cell.value
                if existing is None or (isinstance(existing, str) and not str(existing).strip()):
                    val_cell.value = sig_val.strip()
                else:
                    val_cell.value = str(existing) + sig_val.strip()

        rows_to_delete.append(r)

    return rows_to_delete

def is_n_percent_layout(ws, start_row, end_row):
    scan_end = min(end_row, start_row + 30)
    texts = []

    for r in range(start_row, scan_end + 1):
        for c in (1, 2, 3):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip():
                texts.append(v.upper())

    joined = " ".join(texts)
    joined = joined.replace("\r", " ").replace("\n", " ").replace("  ", " ")
    joined = re.sub(r"\s+", " ", joined)

    has_1st = ("1ST ROW" in joined)
    has_2nd = ("2ND ROW" in joined)
    has_count = ("COUNT" in joined)
    has_col_pct = (
        ("COLUMN %" in joined) or ("COLUMN%" in joined) or
        ("COLUM %" in joined) or ("COLUM%" in joined)
    )
    return has_1st and has_2nd and has_count and has_col_pct

def unmerge_ranges_touching_rows(ws, target_rows):
    if not target_rows:
        return
    targets = set(target_rows)
    for mr in list(ws.merged_cells.ranges):
        if any(r in targets for r in range(mr.min_row, mr.max_row + 1)):
            ws.unmerge_cells(str(mr))

def renumber_col_a_by_col_b(ws, crosstab_mode="NORMAL"):
    mode = (crosstab_mode or "").upper()
    if mode != "NORMAL":
        return

    stub_starts = [r for r, cell in enumerate(ws['A'], 1) if str(cell.value).startswith('Stub')]
    if not stub_starts:
        stub_starts = [1]
    table_boundaries = stub_starts + [ws.max_row + 2]

    for i in range(len(table_boundaries) - 1):
        start_row = table_boundaries[i]
        end_row = table_boundaries[i + 1] - 2
        if end_row <= start_row:
            continue

        data_total_row = -1
        for r in range(start_row, end_row + 1):
            c = ws.cell(row=r, column=3).value
            if str(c).strip().upper() == 'TOTAL':
                data_total_row = r
                break
        if data_total_row == -1:
            continue

        idx = 0
        for r in range(data_total_row + 1, end_row + 1):
            b = ws.cell(row=r, column=2).value
            has_label = b not in (None, "") and str(b).strip() != ""
            a_cell = ws.cell(row=r, column=1)
            if isinstance(a_cell, MergedCell):
                # Non-anchor of a merged A pair (the Column % row). Leave it as
                # part of the merge but keep the index sequence consistent.
                if has_label:
                    idx += 1
                continue
            if has_label:
                a_cell.value = idx
                idx += 1
            else:
                a_cell.value = None

def transfer_bottom_borders_before_delete(ws, rows_to_delete):
    """
    Before rows are deleted, push each deleted row's bottom border up to the
    nearest surviving row above it. The group-closing grid line lives on the
    Sig row (the row we delete), so without this the bottom border of every
    group would vanish from the data columns -> broken table grid.
    """
    if not rows_to_delete:
        return
    del_set = set(rows_to_delete)
    max_col = ws.max_column

    # Pre-map every merged non-anchor cell -> its anchor, so resolving the
    # writable target is an O(1) lookup instead of rescanning all merged ranges
    # for every border we move (sheets can have thousands of merges).
    anchor_of = {}
    for mr in ws.merged_cells.ranges:
        anchor = (mr.min_row, mr.min_col)
        for rr in range(mr.min_row, mr.max_row + 1):
            for cc in range(mr.min_col, mr.max_col + 1):
                if (rr, cc) != anchor:
                    anchor_of[(rr, cc)] = anchor

    for r in sorted(del_set):
        t = r - 1
        while t in del_set and t >= 1:
            t -= 1
        if t < 1:
            continue
        for c in range(1, max_col + 1):
            src = ws.cell(row=r, column=c).border
            if src.bottom and src.bottom.style:
                ar, ac = anchor_of.get((t, c), (t, c))
                dst = ws.cell(row=ar, column=ac)
                b = dst.border
                dst.border = Border(
                    left=b.left,
                    right=b.right,
                    top=b.top,
                    bottom=src.bottom,
                    diagonal=b.diagonal,
                    diagonal_direction=b.diagonal_direction,
                    outline=b.outline,
                    vertical=b.vertical,
                    horizontal=b.horizontal,
                )

def delete_rows_preserving_merges(ws, rows_to_delete):
    """
    Delete rows AND correctly shift/clip merged ranges and row heights.

    openpyxl's ws.delete_rows() does not adjust merged ranges or row heights at
    all, leaving stale merges and mis-indexed heights after the rows above them
    are removed. Stale merges then corrupt neighbouring values when written to
    (labels disappear); stale heights make random rows too tall/short. To avoid
    this we snapshot merges + heights, unmerge all, delete the rows, then
    re-apply each remapped to its new position:
      - merges living entirely on deleted rows are dropped; merges that lost
        only some rows (e.g. a 3-row Count/%/Sig label block that loses its Sig
        row) are clipped to the surviving rows
      - row heights are re-keyed to the surviving rows' new positions
    """
    if not rows_to_delete:
        return

    del_set = set(rows_to_delete)
    sorted_dels = sorted(del_set)

    # Strip hyperlinks ONLY on the rows we are about to delete. Removing every
    # hyperlink in the sheet (as an older version did) wiped the navigation
    # links in the header rows ("Contents"/"Info"/"Next") so they stopped
    # working after a Sig-beside run. Surviving rows keep their hyperlinks;
    # openpyxl rewrites each one's ref from cell.coordinate at save time, so the
    # rebuilt (shifted) cells below still point at the right place.
    for r in del_set:
        for cell in ws[r]:
            if cell.hyperlink is not None:
                cell.hyperlink = None

    def new_index(row):
        # surviving row's new position = row minus deleted rows above it
        return row - bisect.bisect_left(sorted_dels, row)

    snapshot = [(mr.min_row, mr.max_row, mr.min_col, mr.max_col)
                for mr in ws.merged_cells.ranges]

    # Snapshot the style of EVERY cell inside a merged range before unmerging.
    # ws.unmerge_cells() drops the non-anchor MergedCell objects from ws._cells,
    # so the rebuild below loses them and only the top-left anchor keeps a border.
    # Excel still PRINTS such a merged box (it derives the outline from the anchor)
    # but on SCREEN it draws borders per-cell -> the right/bottom edges have no
    # cell to draw them and the table grid looks broken in the app. We restore
    # these cells (with their per-edge borders) at their shifted positions below.
    merged_cell_styles = {}
    for (r1, r2, c1, c2) in snapshot:
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                src = ws._cells.get((rr, cc))
                if src is not None:
                    merged_cell_styles[(rr, cc)] = src._style

    for mr in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mr))

    # snapshot explicit row heights before deleting (openpyxl won't shift them)
    old_heights = {r: dim.height for r, dim in ws.row_dimensions.items()
                   if dim.height is not None}

    # Compact the rows in a SINGLE pass by rebuilding the cell map, instead of
    # calling ws.delete_rows() once per row. delete_rows() shifts every cell
    # below the deletion point each time -> O(deleted * cells), which is tens of
    # seconds on large sheets. Rebuilding is O(cells). Safe here because the
    # sheet has no conditional formatting / data validation that reference rows
    # (merges and row heights are handled explicitly below).
    new_cells = {}
    for (row, col), cell in ws._cells.items():
        if row in del_set:
            continue
        nr = new_index(row)
        if nr != row:
            cell.row = nr
        new_cells[(nr, col)] = cell
    ws._cells = new_cells
    ws._current_row = max((r for r, _ in new_cells), default=0)

    # re-key row heights: clear whatever delete_rows left behind, then re-apply
    # each surviving row's height at its shifted position.
    for r in list(ws.row_dimensions.keys()):
        ws.row_dimensions[r].height = None
    for old_r, h in old_heights.items():
        if old_r in del_set:
            continue
        ws.row_dimensions[new_index(old_r)].height = h

    # Recreate the merged-range cells that unmerge_cells dropped, at their shifted
    # positions, restoring each cell's original style. This keeps every edge of a
    # merged box (right/bottom included) carrying its own border so the grid lines
    # render on screen, not only in print. Done before re-adding the merges so the
    # targets are still writable plain cells.
    for (rr, cc), st in merged_cell_styles.items():
        if rr in del_set:
            continue
        nr = new_index(rr)
        existing = ws._cells.get((nr, cc))
        if existing is not None:
            existing._style = st
        else:
            ws.cell(row=nr, column=cc)._style = st

    # Re-apply via merged_cells.add (not ws.merge_cells): these cells came from
    # ranges that were already style-cleaned when the file was first merged, and
    # unmerging does not undo that, so the expensive _clean_merge_range step is
    # unnecessary here and only adds cost.
    for (r1, r2, c1, c2) in snapshot:
        surviving = [r for r in range(r1, r2 + 1) if r not in del_set]
        if not surviving:
            continue
        nr1 = new_index(surviving[0])
        nr2 = new_index(surviving[-1])
        if nr1 == nr2 and c1 == c2:
            continue  # collapsed to a single cell -> nothing to merge
        ws.merged_cells.add(CellRange(min_col=c1, min_row=nr1, max_col=c2, max_row=nr2))

def add_bottom_grid_to_last_used_row(ws):
    last_used_row = 0
    last_used_col = 0
    max_col = ws.max_column

    for r in range(ws.max_row, 0, -1):
        row_has_data = False
        row_last_col = 0
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if v not in (None, "") and str(v).strip() != "":
                row_has_data = True
                row_last_col = c
        if row_has_data:
            last_used_row = r
            last_used_col = row_last_col
            break

    if last_used_row == 0 or last_used_col == 0:
        return

    thin = Side(style="thin")
    for c in range(1, last_used_col + 1):
        cell = ws.cell(row=last_used_row, column=c)
        b = cell.border
        cell.border = Border(
            left=b.left,
            right=b.right,
            top=b.top,
            bottom=thin,
            diagonal=b.diagonal,
            diagonal_direction=b.diagonal_direction,
            outline=b.outline,
            vertical=b.vertical,
            horizontal=b.horizontal,
        )

def merge_col_ab_pairs_for_n_percent(ws, crosstab_mode="NORMAL"):
    mode = (crosstab_mode or "").upper()
    max_col = ws.max_column

    stub_starts = [r for r, cell in enumerate(ws['A'], 1) if str(cell.value).startswith('Stub')]
    if not stub_starts:
        stub_starts = [1]
    table_boundaries = stub_starts + [ws.max_row + 2]

    # Rows whose column A already belongs to a merge. delete_rows_preserving_merges
    # has normally already merged the Count/Column% A-B pairs, so this lets us skip
    # the expensive unmerge+remerge for pairs that are already correct.
    premerged_a_rows = set()
    for mr in ws.merged_cells.ranges:
        if mr.min_col <= 1 <= mr.max_col:
            premerged_a_rows.update(range(mr.min_row, mr.max_row + 1))

    for i in range(len(table_boundaries) - 1):
        start_row = table_boundaries[i]
        end_row = table_boundaries[i + 1] - 2
        if end_row <= start_row:
            continue

        data_total_row = -1
        if mode == "MATRIX":
            data_total_row = find_total_row_in_colB(ws, start_row, end_row)
        else:
            for r in range(start_row, end_row + 1):
                c = ws.cell(row=r, column=3).value
                if str(c).strip().upper() == 'TOTAL':
                    data_total_row = r
                    break
        if data_total_row == -1:
            continue

        r = data_total_row + 1
        while r <= end_row:
            label = ws.cell(row=r, column=2).value
            has_label = label not in (None, "") and str(label).strip() != ""

            if not has_label:
                r += 1
                continue

            r2 = r + 1
            if r2 <= end_row:
                next_label = ws.cell(row=r2, column=2).value
                next_has_label = next_label not in (None, "") and str(next_label).strip() != ""
                if not next_has_label:
                    # Already merged by the delete step -> nothing to do.
                    if r in premerged_a_rows and r2 in premerged_a_rows:
                        r = r2 + 1
                        continue
                    row2_has_data = any(
                        ws.cell(row=r2, column=cc).value not in (None, "") and
                        str(ws.cell(row=r2, column=cc).value).strip() != ""
                        for cc in range(3, max_col + 1)
                    )
                    if row2_has_data:
                        for mr in list(ws.merged_cells.ranges):
                            touches_ab = not (mr.max_col < 1 or mr.min_col > 2)
                            overlaps_rows = not (mr.max_row < r or mr.min_row > r2)
                            if touches_ab and overlaps_rows:
                                ws.unmerge_cells(str(mr))
                        ws.cell(row=r2, column=1).value = None
                        ws.merge_cells(start_row=r, start_column=1, end_row=r2, end_column=1)
                        ws.merge_cells(start_row=r, start_column=2, end_row=r2, end_column=2)
                        r = r2 + 1
                        continue
            r += 1


# ===================== Core Processing =====================

def place_sig_stamp(ws, header_row, sig_text, col=3):
    stamp_row = header_row - 2 if header_row - 2 >= 1 else header_row
    cell = ws.cell(row=stamp_row, column=col)
    cell.value = sig_text
    cell.font = Font(name='Arial', size=9, color='FFFF0000')


def process_single_excel_file(file_path, sig_input, crosstab_mode="NORMAL", sig_beside=False):
    import openpyxl
    try:
        rules_by_char = parse_sig_groups(sig_input)
        wb = openpyxl.load_workbook(file_path)
        sheets_to_process = [n for n in wb.sheetnames if n.lower() not in ['contents', 'info']]
        cells_changed_count = 0

        for sheet_name in sheets_to_process:
            ws = wb[sheet_name]
            rows_to_delete_in_sheet = []

            stub_starts = [r for r, cell in enumerate(ws['A'], 1) if str(cell.value).startswith('Stub')]
            if not stub_starts:
                stub_starts = [1]
            table_boundaries = stub_starts + [ws.max_row + 2]

            for i in range(len(table_boundaries) - 1):
                start_row = table_boundaries[i]
                end_row   = table_boundaries[i + 1] - 2
                if end_row <= start_row:
                    continue

                sig_text = f"Sig Used: {sig_input}"
                n_percent_layout = is_n_percent_layout(ws, start_row, end_row)

                if crosstab_mode.upper() == "NORMAL":
                    header_row, data_total_row = -1, -1
                    for r in range(start_row, end_row + 1):
                        d = ws.cell(row=r, column=4).value
                        e = ws.cell(row=r, column=5).value
                        c = ws.cell(row=r, column=3).value
                        if clean_header(str(d)) == 'A' and clean_header(str(e)) == 'B':
                            header_row = r
                        if str(c).strip().upper() == 'TOTAL':
                            data_total_row = r
                    if header_row == -1 or data_total_row == -1:
                        continue

                    place_sig_stamp(ws, header_row, sig_text, col=3)

                    for col in range(3, ws.max_column + 1):
                        hdr = clean_header(str(ws.cell(row=header_row, column=col).value))
                        if hdr in rules_by_char:
                            keep = set(rules_by_char[hdr])
                            for r in range(data_total_row + 1, end_row + 1):
                                v = ws.cell(row=r, column=col).value
                                if isinstance(v, str) and v.strip():
                                    nv = filter_sig_text(v, keep, preserve_tokens=("ADJ",))
                                    if nv != v:
                                        ws.cell(row=r, column=col).value = nv
                                        cells_changed_count += 1

                    if sig_beside:
                        data_cols_n = list(range(3, ws.max_column + 1))
                        sig_rows = collect_sig_rows_and_merge(
                            ws, data_total_row + 1, end_row, label_cols=[1, 2], data_cols=data_cols_n,
                            keep_first_unlabeled=n_percent_layout
                        )
                        rows_to_delete_in_sheet.extend(sig_rows)
                    continue

                # ================= MATRIX =================
                header_row, first_col = detect_letter_header_row(ws, start_row, end_row, min_seq=3)
                if header_row == -1:
                    continue

                place_sig_stamp(ws, header_row, sig_text, col=3)

                col_map = build_col_letter_map(ws, header_row, first_col)
                if not col_map:
                    continue

                total_row = find_total_row_in_colB(ws, header_row, end_row)
                if total_row == -1:
                    continue
                data_start = total_row
                data_end   = find_last_data_row(ws, data_start, list(col_map.keys()))
                if data_end < data_start:
                    continue

                for c, letter in col_map.items():
                    allowed = set(rules_by_char.get(letter, ""))
                    for r in range(data_start, data_end + 1):
                        v = ws.cell(row=r, column=c).value
                        if isinstance(v, str) and v.strip():
                            nv = filter_sig_text(v, allowed, preserve_tokens=("ADJ",))
                            if nv != v:
                                ws.cell(row=r, column=c).value = nv
                                cells_changed_count += 1

                if sig_beside:
                    sig_rows = collect_sig_rows_and_merge(
                        ws, data_start, data_end, label_cols=[1, 2], data_cols=list(col_map.keys()),
                        keep_first_unlabeled=n_percent_layout
                    )
                    rows_to_delete_in_sheet.extend(sig_rows)

            if rows_to_delete_in_sheet:
                transfer_bottom_borders_before_delete(ws, rows_to_delete_in_sheet)
                delete_rows_preserving_merges(ws, rows_to_delete_in_sheet)

                renumber_col_a_by_col_b(ws, crosstab_mode)
                add_bottom_grid_to_last_used_row(ws)

            if sig_beside:
                merge_col_ab_pairs_for_n_percent(ws, crosstab_mode)

        wb.save(file_path)
        return cells_changed_count, "Success"

    except Exception as e:
        return 0, f"Error: {e}"



# ===================== Auto Lychee integration =====================

def process_in_place(file_path, sig_input, crosstab_mode="NORMAL", sig_beside=False):
    """Run Del Sig on one exported workbook and save it back to the same path (openpyxl, no Excel).

    Works on a copy and replaces the file only on success, so an error leaves it as it was."""
    import shutil
    import tempfile
    from pathlib import Path

    file_path = Path(file_path)
    if not str(sig_input or "").strip():
        raise ValueError("Del Sig: ยังไม่ได้ใส่กลุ่ม Sig")
    with tempfile.TemporaryDirectory(prefix="autolychee_delsig_") as work:
        working = Path(work) / file_path.name
        shutil.copy2(file_path, working)
        from fast_styles import exact_style_cache
        with exact_style_cache():  # same output bytes, faster openpyxl style handling
            cells, status = process_single_excel_file(str(working), sig_input.strip(), crosstab_mode, sig_beside)
        if status != "Success":
            raise RuntimeError(f"Del Sig ไม่สำเร็จ: {status}")
        pending = file_path.with_name(f".{file_path.stem}.del_sig_pending{file_path.suffix}")
        try:
            shutil.copy2(working, pending)
            os.replace(pending, file_path)
        finally:
            if pending.exists():
                pending.unlink()
    side = "Sig ข้าง" if sig_beside else "Sig ปกติ"
    return f"{crosstab_mode.upper()} · {side} · แก้ไข {cells} cells"


# ====== MODULE: post_cut_percent ======
"""Cut N / % rows — core copied unchanged from 151_CutLychee_Persence.py (PyQt UI and the
ProcessPoolExecutor removed: Auto Lychee runs it single-threaded to keep CPU low)."""
# from __future__ import annotations  (applied to this section by the loader)

import sys
import traceback
import os
import time
import re
from datetime import datetime
from copy import copy
from pathlib import Path
from typing import List, Tuple

import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.styles import PatternFill
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from openpyxl.worksheet.hyperlink import Hyperlink
from bisect import bisect_left



# ----------------------------
# Excel processing core
# ----------------------------
KEEP_COUNT = "count"
KEEP_PERCENT = "percent"
KEEP_BOTH = "both"
TABLE_ONE_SHEET_SIG = "one_sheet_sig"
TABLE_ONE_SHEET_NOT_SIG = "one_sheet_not_sig"
TABLE_MULTI_SHEET_SIG = "multi_sheet_sig"
TABLE_MULTI_SHEET_NOT_SIG = "multi_sheet_not_sig"
QUESTION_CODE_RE = re.compile(r"\b[A-Z]{1,3}\d+[A-Z]?(?:_[0-9]+|Z\d+)?\b", re.IGNORECASE)
SIG_TEXT_RE = re.compile(r"^[A-Z]+$")


def unique_output_path(out_dir: Path, source_path: Path, keep_mode: str = KEEP_COUNT) -> Path:
    base_stem = build_output_stem(source_path.stem, keep_mode)
    candidate = out_dir / f"{base_stem}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{base_stem}_{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def sheet_has_table_legend(sheet) -> bool:
    for row in range(1, sheet.max_row + 1):
        for col in range(1, 3):
            value = sheet.cell(row=row, column=col).value
            if value is None:
                continue
            text = str(value)
            if re.search(r"(?i)count", text) and re.search(r"(?i)column\s*%", text):
                return True
    return False


def count_stub_blocks(sheet) -> int:
    count = 0
    for row in range(1, sheet.max_row + 1):
        if str(sheet.cell(row=row, column=1).value or "").strip().lower().startswith("stub:"):
            count += 1
    return count


def detect_one_sheet_layout(wb) -> bool:
    table_like_sheets = []
    for sheet in wb.worksheets:
        if sheet.title.strip().lower() in {"contents", "itemlist"}:
            continue
        if sheet_has_table_legend(sheet) or count_stub_blocks(sheet):
            table_like_sheets.append(sheet)

    if not table_like_sheets:
        return False
    if any(count_stub_blocks(sheet) > 1 for sheet in table_like_sheets):
        return True
    return len(table_like_sheets) == 1 and table_like_sheets[0].max_row > 200


def detect_sig_layout(wb) -> bool:
    label_merge_heights: list[int] = []

    for sheet in wb.worksheets:
        if sheet.title.strip().lower() in {"contents", "itemlist"}:
            continue
        if not sheet_has_table_legend(sheet):
            continue

        for merged_range in sheet.merged_cells.ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            if min_col <= 2 and max_col <= 2 and max_row > min_row:
                text = str(sheet.cell(row=min_row, column=min_col).value or "")
                if re.search(r"(?i)1st\s*row", text):
                    continue
                label_merge_heights.append(max_row - min_row + 1)

    if not label_merge_heights:
        return False
    return max(label_merge_heights) >= 3


def detect_workbook_table_type(wb) -> str:
    is_one_sheet = detect_one_sheet_layout(wb)
    has_sig = detect_sig_layout(wb)

    if is_one_sheet and has_sig:
        return TABLE_ONE_SHEET_SIG
    if is_one_sheet and not has_sig:
        return TABLE_ONE_SHEET_NOT_SIG
    if not is_one_sheet and has_sig:
        return TABLE_MULTI_SHEET_SIG
    return TABLE_MULTI_SHEET_NOT_SIG


def build_output_stem(source_stem: str, keep_mode: str = KEEP_COUNT) -> str:
    today = datetime.now().strftime("%Y%m%d")
    stem = source_stem

    # Remove trailing processed marker if present.
    stem = re.sub(r"_processed(?:_\d+)?$", "", stem, flags=re.IGNORECASE)

    # Normalize any N% token in filename to the selected output type and keep it
    # as a clear suffix before the date.
    output_token = "%" if keep_mode == KEEP_PERCENT else "N"
    stem = re.sub(r"(?i)\bN\s*%", output_token, stem)

    date_match = re.search(r"(?:\s+)?\d{8}$", stem)
    if date_match:
        stem = stem[: date_match.start()].rstrip()

    stem = re.sub(r"(?i)(?:\s+)(?:N|%)$", "", stem.strip()).rstrip()
    stem = f"{stem} {output_token}" if stem else output_token

    stem = f"{stem} {today}"

    return stem


def excel_sheet_location(sheet_name: str, cell_ref: str = "A1") -> str:
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", sheet_name):
        return f"{sheet_name}!{cell_ref}"
    escaped = sheet_name.replace("'", "''")
    return f"'{escaped}'!{cell_ref}"


def set_internal_link(
    ws,
    coord: str,
    target_sheet: str | None,
    text: str | None = None,
    cell_ref: str = "A1",
) -> None:
    cell = ws[coord]
    if text is not None:
        cell.value = text
    if not target_sheet:
        cell.hyperlink = None
        return
    cell.hyperlink = Hyperlink(
        ref=coord,
        location=excel_sheet_location(target_sheet, cell_ref),
        display=str(cell.value) if cell.value is not None else None,
    )


def sheet_from_internal_location(location: str) -> str | None:
    quoted = re.match(r"'((?:[^']|'')+)'!", location)
    if quoted:
        return quoted.group(1).replace("''", "'")
    unquoted = re.match(r"([^!]+)!", location)
    if unquoted:
        return unquoted.group(1)
    return None


def normalize_question_code(value: object) -> str | None:
    text = str(value or "")
    matches = QUESTION_CODE_RE.findall(text)
    if not matches:
        return None
    code = matches[0].upper()
    code = re.sub(r"Z\d+$", "", code)
    return code


def collect_table_anchors(wb, contents_sheet: str = "Contents") -> tuple[list[tuple[str, str]], dict[str, list[tuple[str, str]]]]:
    ordered: list[tuple[str, str]] = []
    by_code: dict[str, list[tuple[str, str]]] = {}

    for ws in wb.worksheets:
        if ws.title == contents_sheet:
            continue
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=1).value
            text = str(value or "").strip()
            if not text.lower().startswith("stub:"):
                continue
            anchor_row = row_idx
            for candidate_row in range(row_idx - 1, max(0, row_idx - 5), -1):
                if str(ws.cell(row=candidate_row, column=1).value or "").strip().lower() == "contents":
                    anchor_row = candidate_row
                    break
            anchor = (ws.title, f"A{anchor_row}")
            ordered.append(anchor)
            code = normalize_question_code(text)
            if code:
                by_code.setdefault(code, []).append(anchor)

    return ordered, by_code


def repair_contents_table_links(wb, contents_sheet: str = "Contents") -> None:
    if contents_sheet not in wb.sheetnames:
        return

    ws = wb[contents_sheet]
    ordered_anchors, anchors_by_code = collect_table_anchors(wb, contents_sheet)
    used_by_code: dict[str, int] = {}
    sequential_idx = 0

    for row_idx in range(1, ws.max_row + 1):
        table_col = None
        for candidate_col in (1, 2):
            if str(ws.cell(row=row_idx, column=candidate_col).value or "").strip().lower() == "table":
                table_col = candidate_col
                break
        if table_col is None:
            continue

        cell = ws.cell(row=row_idx, column=table_col)
        question_text = ws.cell(row=row_idx, column=table_col + 1).value
        code = normalize_question_code(question_text)
        anchor = None

        if code and code in anchors_by_code:
            code_idx = used_by_code.get(code, 0)
            choices = anchors_by_code[code]
            anchor = choices[min(code_idx, len(choices) - 1)]
            used_by_code[code] = code_idx + 1

        if anchor is None and sequential_idx < len(ordered_anchors):
            anchor = ordered_anchors[sequential_idx]

        sequential_idx += 1
        if anchor is None:
            cell.hyperlink = None
            continue

        target_sheet, cell_ref = anchor
        set_internal_link(ws, cell.coordinate, target_sheet, cell_ref=cell_ref)


def repair_internal_links(wb, contents_sheet: str = "Contents") -> None:
    existing = set(wb.sheetnames)
    navigable_sheets: List[str] = []
    has_contents_sheet = contents_sheet in existing

    repair_contents_table_links(wb, contents_sheet)

    for name in wb.sheetnames:
        if name == contents_sheet:
            continue
        ws = wb[name]
        if str(ws["A1"].value or "").strip().lower() == "contents":
            navigable_sheets.append(name)

    for idx, name in enumerate(navigable_sheets):
        ws = wb[name]
        if has_contents_sheet:
            set_internal_link(ws, "A1", contents_sheet)
        else:
            set_internal_link(ws, "A1", name)
        if str(ws["B1"].value or "").strip().lower() == "info" and "Info" in existing:
            set_internal_link(ws, "B1", "Info")

        prev_name = navigable_sheets[idx - 1] if idx > 0 else None
        next_name = navigable_sheets[idx + 1] if idx < len(navigable_sheets) - 1 else None

        if str(ws["E1"].value or "").strip().lower() == "previous":
            set_internal_link(ws, "E1", prev_name)
        if str(ws["F1"].value or "").strip().lower() == "next":
            set_internal_link(ws, "F1", next_name)

    for ws in wb.worksheets:
        first_contents_ref = None
        for row in ws.iter_rows():
            for cell in row:
                if str(cell.value or "").strip().lower() == "contents":
                    first_contents_ref = cell.coordinate
                    break
            if first_contents_ref:
                break

        for row in ws.iter_rows():
            for cell in row:
                if str(cell.value or "").strip().lower() == "contents":
                    if has_contents_sheet:
                        set_internal_link(ws, cell.coordinate, contents_sheet)
                    elif first_contents_ref:
                        set_internal_link(ws, cell.coordinate, ws.title, cell_ref=first_contents_ref)
                    continue

                if not cell.hyperlink:
                    continue
                location = cell.hyperlink.location
                if not location and isinstance(cell.hyperlink.target, str):
                    target = cell.hyperlink.target
                    if target.startswith("#"):
                        location = target[1:]
                if not location:
                    continue
                location_text = str(location)
                target_sheet = sheet_from_internal_location(location_text)
                if not target_sheet:
                    continue
                if target_sheet not in existing:
                    cell.hyperlink = None
                elif isinstance(cell.hyperlink.target, str) and cell.hyperlink.target.startswith("#"):
                    cell.hyperlink = Hyperlink(
                        ref=cell.coordinate,
                        location=location_text,
                        display=str(cell.value) if cell.value is not None else None,
                    )


def unique_output_path_reserved(
    out_dir: Path, source_path: Path, reserved_names: set[str], keep_mode: str = KEEP_COUNT
) -> Path:
    base = build_output_stem(source_path.stem, keep_mode)
    candidate_name = f"{base}.xlsx"
    idx = 1
    while True:
        candidate = out_dir / candidate_name
        key = str(candidate.resolve()).lower()
        if key not in reserved_names and not candidate.exists():
            reserved_names.add(key)
            return candidate
        candidate_name = f"{base}_{idx}.xlsx"
        idx += 1


def process_one_file_task(src_path: str, dst_path: str, keep_mode: str = KEEP_COUNT) -> str:
    process_workbook(Path(src_path), Path(dst_path), keep_mode)
    return dst_path


def process_workbook(input_path: Path, save_path: Path, keep_mode: str = KEEP_COUNT) -> None:
    if keep_mode not in {KEEP_COUNT, KEEP_PERCENT}:
        raise ValueError(f"Unsupported keep mode: {keep_mode}")

    wb = openpyxl.load_workbook(input_path)
    table_type = detect_workbook_table_type(wb)

    def detect_effective_max_col(sheet) -> int:
        # Use worksheet dimension first, then trim based on table values.
        # Rows 1-3 can contain navigation links such as Previous/Next outside the table.
        dim = sheet.calculate_dimension()
        _, _, dim_max_col, _ = range_boundaries(dim)

        max_row = sheet.max_row
        sample_rows = [4, 5, 6, 7, 8, 9, 10, max_row]
        sample_rows.extend(range(4, min(max_row, 30) + 1))
        if max_row > 30:
            sample_rows.extend(range(max_row - 9, max_row + 1))
        sample_rows = sorted(set(r for r in sample_rows if 1 <= r <= max_row))

        for col in range(dim_max_col, 2, -1):
            for r in sample_rows:
                value = sheet.cell(row=r, column=col).value
                if value is not None and str(value).strip() != "":
                    return col
        return 3

    def delete_rows_desc(sheet, row_indexes: List[int]) -> None:
        if not row_indexes:
            return
        rows = sorted(set(row_indexes), reverse=True)
        start = rows[0]
        count = 1
        prev = rows[0]
        for r in rows[1:]:
            if r == prev - 1:
                count += 1
            else:
                sheet.delete_rows(start - count + 1, count)
                start = r
                count = 1
            prev = r
        sheet.delete_rows(start - count + 1, count)

    def compact_rows_once(sheet, row_indexes: List[int]) -> None:
        deleted_rows = sorted(set(row_indexes))
        if not deleted_rows:
            return
        deleted_set = set(deleted_rows)
        original_merged_ranges = []
        for merged_range in sheet.merged_cells.ranges:
            top_cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
            original_merged_ranges.append(
                (
                    merged_range.bounds,
                    top_cell.value,
                    copy(top_cell.border),
                    copy(top_cell.fill),
                    copy(top_cell.font),
                    copy(top_cell.alignment),
                    top_cell.number_format,
                )
            )
        new_cells = {}
        for (row_idx, col_idx), cell in list(sheet._cells.items()):
            if row_idx in deleted_set:
                continue
            new_row = row_idx - bisect_left(deleted_rows, row_idx)
            if new_row != row_idx:
                cell.row = new_row
            new_cells[(new_row, col_idx)] = cell
        sheet._cells = new_cells

        new_row_dimensions = {}
        for row_idx, dimension in list(sheet.row_dimensions.items()):
            if row_idx in deleted_set:
                continue
            new_row = row_idx - bisect_left(deleted_rows, row_idx)
            dimension.index = new_row
            new_row_dimensions[new_row] = dimension
        # Auto Lychee fix (2026-10-01): keep openpyxl's own row_dimensions object (it creates a row's
        # dimension on first access). The original assigned this plain dict, so a later
        # `sheet.row_dimensions[row].height` for a row without one raised KeyError (live: One Sheet
        # Matrix export). Files that worked before never hit that lookup, so their output is unchanged.
        sheet.row_dimensions.clear()
        sheet.row_dimensions.update(new_row_dimensions)
        sheet._current_row = max((row for row, _ in sheet._cells.keys()), default=1)

        shifted_merges = []
        for bounds, merge_label, merge_border, merge_fill, merge_font, merge_alignment, merge_number_format in original_merged_ranges:
            min_col, min_row, max_col, max_row = bounds
            is_label_merge = min_col <= 2 and max_col <= 2
            if not is_label_merge:
                continue
            remaining_rows = [row_idx for row_idx in range(min_row, max_row + 1) if row_idx not in deleted_set]
            if not remaining_rows:
                continue
            new_min_row = remaining_rows[0] - bisect_left(deleted_rows, remaining_rows[0])
            new_max_row = remaining_rows[-1] - bisect_left(deleted_rows, remaining_rows[-1])
            start = f"{get_column_letter(min_col)}{new_min_row}"
            end = f"{get_column_letter(max_col)}{new_max_row}"
            should_merge = not (new_min_row == new_max_row and min_col == max_col)
            if should_merge:
                shifted_merges.append(CellRange(f"{start}:{end}"))
            existing_cell = sheet._cells.get((new_min_row, min_col))
            if existing_cell is not None and type(existing_cell).__name__ == "MergedCell":
                del sheet._cells[(new_min_row, min_col)]
            target_cell = sheet.cell(row=new_min_row, column=min_col)
            target_cell.value = merge_label
            target_cell.border = copy(merge_border)
            target_cell.fill = copy(merge_fill)
            target_cell.font = copy(merge_font)
            target_cell.alignment = copy(merge_alignment)
            target_cell.number_format = merge_number_format
        new_merged_cells = MultiCellRange()
        new_merged_cells.ranges = set(shifted_merges)
        sheet.merged_cells = new_merged_cells

    def choose_value_and_format(upper_cell, lower_cell):
        upper_val = upper_cell.value
        lower_val = lower_cell.value

        if lower_val is None or str(lower_val).strip() == "":
            return upper_val, upper_cell.number_format

        if isinstance(upper_val, (int, float)) and isinstance(lower_val, (int, float)):
            upper_has_frac = abs(float(upper_val) - int(float(upper_val))) > 1e-12
            lower_has_frac = abs(float(lower_val) - int(float(lower_val))) > 1e-12

            if upper_has_frac and not lower_has_frac:
                return upper_val, upper_cell.number_format
            return lower_val, lower_cell.number_format

        return lower_val, lower_cell.number_format

    def append_sig_values_to_row(sheet, target_row: int, sig_row: int, max_col: int) -> None:
        appended = False
        label_cells = []
        for col in range(1, min(max_col, 2) + 1):
            target_cell = sheet.cell(row=target_row, column=col)
            sig_cell = sheet.cell(row=sig_row, column=col)
            label_cells.append(target_cell)
            target_cell.border = Border(
                left=target_cell.border.left,
                right=target_cell.border.right,
                top=target_cell.border.top,
                bottom=sig_cell.border.bottom,
            )
        for col in range(3, max_col + 1):
            target_cell = sheet.cell(row=target_row, column=col)
            sig_cell = sheet.cell(row=sig_row, column=col)
            sig_value = sig_cell.value
            if sig_value is not None and str(sig_value).strip() != "":
                appended = True
                base_value = target_cell.value
                if base_value is None or str(base_value).strip() == "":
                    target_cell.value = sig_value
                else:
                    target_cell.value = f"{base_value}\n{sig_value}"
                new_alignment = copy(target_cell.alignment)
                new_alignment.wrap_text = True
                target_cell.alignment = new_alignment
            target_cell.border = Border(
                left=target_cell.border.left,
                right=target_cell.border.right,
                top=target_cell.border.top,
                bottom=sig_cell.border.bottom,
            )
        if appended:
            target_height = sheet.row_dimensions[target_row].height or 15
            sig_height = sheet.row_dimensions[sig_row].height or 15
            sheet.row_dimensions[target_row].height = max(target_height, target_height + sig_height)
            thin_border = Side(border_style="thin", color="000000")
            for cell in label_cells:
                if cell.border.bottom.style is None:
                    cell.border = Border(
                        left=cell.border.left,
                        right=cell.border.right,
                        top=cell.border.top,
                        bottom=thin_border,
                    )

    def normalize_multiline_rows(sheet, max_col: int) -> None:
        for row_idx in range(1, sheet.max_row + 1):
            has_multiline = False
            for col_idx in range(1, max_col + 1):
                cell = sheet.cell(row=row_idx, column=col_idx)
                if "\n" not in str(cell.value or ""):
                    continue
                # Auto Lychee fix (2026-10-01): the "1st row: Count / 2nd row: Column %" legend is
                # multi-line too. In a Matrix export it sits on the column-number row (A4, merged A4:B5),
                # which this raised to 30 pt (user: make it as low as the TOTAL row again). The legend
                # still gets wrap_text, it just no longer makes its row taller.
                if not re.search(r"(?i)1st\s*row", str(cell.value)):
                    has_multiline = True
                new_alignment = copy(cell.alignment)
                new_alignment.wrap_text = True
                cell.alignment = new_alignment
            if has_multiline:
                sheet.row_dimensions[row_idx].height = max(sheet.row_dimensions[row_idx].height or 15, 30)

    def normalize_body_label_styles(sheet, max_col: int) -> None:
        thin_border = Side(border_style="thin", color="000000")
        label_fill = None
        index_fill = None
        for row_idx in range(1, sheet.max_row + 1):
            label_cell = sheet.cell(row=row_idx, column=2)
            if str(label_cell.value or "").strip() and label_cell.fill.fill_type:
                label_fill = copy(label_cell.fill)
                index_fill = copy(sheet.cell(row=row_idx, column=1).fill)
                break
        if label_fill is None:
            label_fill = PatternFill("solid", fgColor="FFF5E4")
        if index_fill is None:
            index_fill = PatternFill("solid", fgColor="FFFFFF")
        mean_index_fill = PatternFill("solid", fgColor="FFFF0000")

        def is_body_index(text: str) -> bool:
            return bool(re.fullmatch(r"\d+(?:\.0+)?", text))

        for row_idx in range(1, sheet.max_row + 1):
            code_text = str(sheet.cell(row=row_idx, column=1).value or "").strip()
            label_text = str(sheet.cell(row=row_idx, column=2).value or "").strip()
            if re.search(r"(?i)1st\s*row", code_text) or re.search(r"(?i)1st\s*row", label_text):
                continue
            if not is_body_index(code_text):
                continue
            if not label_text and not code_text:
                continue
            has_table_values = any(
                str(sheet.cell(row=row_idx, column=col_idx).value or "").strip()
                for col_idx in range(3, max_col + 1)
            )
            if not has_table_values:
                continue

            row_dimension = sheet.row_dimensions.get(row_idx)
            if row_dimension is not None and row_dimension.height is None:
                row_dimension.height = 13.5
            left_cell = sheet.cell(row=row_idx, column=1)
            label_cell = sheet.cell(row=row_idx, column=2)
            if code_text:
                left_cell.fill = copy(mean_index_fill if label_text.lower() == "mean" else index_fill)
            if label_text:
                label_cell.fill = copy(label_fill)
            for cell in (left_cell, label_cell):
                cell.border = Border(
                    left=cell.border.left if cell.border.left.style else thin_border,
                    right=cell.border.right if cell.border.right.style else thin_border,
                    top=cell.border.top if cell.border.top.style else thin_border,
                    bottom=cell.border.bottom if cell.border.bottom.style else thin_border,
                )

    def normalize_single_row_table_borders(sheet, max_col: int) -> None:
        thin_border = Side(border_style="thin", color="000000")
        merged_body_rows = set()
        for merged_range in sheet.merged_cells.ranges:
            min_col, min_row, merge_max_col, max_row = merged_range.bounds
            if min_col <= 2 and merge_max_col <= 2 and max_row > min_row:
                for row_idx in range(min_row, max_row + 1):
                    merged_body_rows.add(row_idx)

        for row_idx in range(1, sheet.max_row + 1):
            if row_idx in merged_body_rows:
                continue
            label_text = str(sheet.cell(row=row_idx, column=2).value or "").strip()
            code_text = str(sheet.cell(row=row_idx, column=1).value or "").strip()
            if not re.fullmatch(r"\d+(?:\.0+)?", code_text):
                continue
            if not label_text and not code_text:
                continue
            has_table_values = any(
                str(sheet.cell(row=row_idx, column=col_idx).value or "").strip()
                for col_idx in range(3, max_col + 1)
            )
            if not has_table_values:
                continue
            for col_idx in range(1, max_col + 1):
                cell = sheet.cell(row=row_idx, column=col_idx)
                cell.border = Border(
                    left=cell.border.left if cell.border.left.style else thin_border,
                    right=cell.border.right if cell.border.right.style else thin_border,
                    top=cell.border.top if cell.border.top.style else thin_border,
                    bottom=cell.border.bottom if cell.border.bottom.style else thin_border,
                )

    def remove_consecutive_duplicate_labels(sheet) -> None:
        rows_to_delete: List[int] = []
        row_idx = 7
        while row_idx < sheet.max_row:
            code_text = str(sheet.cell(row=row_idx, column=1).value or "").strip()
            label_text = str(sheet.cell(row=row_idx, column=2).value or "").strip()
            next_code_text = str(sheet.cell(row=row_idx + 1, column=1).value or "").strip()
            next_label_text = str(sheet.cell(row=row_idx + 1, column=2).value or "").strip()
            if (
                label_text
                and label_text == next_label_text
                and re.fullmatch(r"\d+(?:\.0+)?", code_text)
                and re.fullmatch(r"\d+(?:\.0+)?", next_code_text)
            ):
                for col_idx in range(3, sheet.max_column + 1):
                    target_cell = sheet.cell(row=row_idx, column=col_idx)
                    source_cell = sheet.cell(row=row_idx + 1, column=col_idx)
                    target_cell.value = source_cell.value
                    target_cell.number_format = source_cell.number_format
                rows_to_delete.append(row_idx + 1)
            row_idx += 1
        delete_rows_desc(sheet, rows_to_delete)

    def normalize_multi_sheet_sig_body_borders(sheet, max_col: int) -> None:
        thin_border = Side(border_style="thin", color="000000")
        merged_rows_by_start: set[int] = set()
        for merged_range in sheet.merged_cells.ranges:
            min_col, min_row, merge_max_col, max_row = merged_range.bounds
            if min_col <= 2 and merge_max_col <= 2 and max_row == min_row + 1:
                merged_rows_by_start.add(min_row)

        for row_idx in sorted(merged_rows_by_start):
            if not re.fullmatch(r"\d+(?:\.0+)?", str(sheet.cell(row=row_idx, column=1).value or "").strip()):
                continue
            sig_row = row_idx + 1
            sheet.row_dimensions[row_idx].height = 13.5
            for col_idx in range(1, max_col + 1):
                value_cell = sheet.cell(row=row_idx, column=col_idx)
                sig_cell = sheet.cell(row=sig_row, column=col_idx)
                value_cell.border = Border(
                    left=value_cell.border.left if value_cell.border.left.style else thin_border,
                    right=value_cell.border.right if value_cell.border.right.style else thin_border,
                    top=value_cell.border.top if value_cell.border.top.style else thin_border,
                    bottom=value_cell.border.bottom,
                )
                sig_cell.border = Border(
                    left=sig_cell.border.left if sig_cell.border.left.style else thin_border,
                    right=sig_cell.border.right if sig_cell.border.right.style else thin_border,
                    top=sig_cell.border.top,
                    bottom=sig_cell.border.bottom if sig_cell.border.bottom.style else thin_border,
                )

    def normalize_one_sheet_header_fill(sheet, max_col: int) -> None:
        header_fill = None
        for row_idx in range(1, sheet.max_row + 1):
            for col_idx in range(3, max_col + 1):
                cell = sheet.cell(row=row_idx, column=col_idx)
                if cell.fill.fill_type and cell.fill.fgColor.rgb not in {None, "00000000", "FFFFFFFF"}:
                    header_fill = copy(cell.fill)
                    break
            if header_fill is not None:
                break
        if header_fill is None:
            header_fill = PatternFill("solid", fgColor="FFEBEBEB")

        for row_idx in range(1, sheet.max_row + 1):
            for col_idx in range(1, 3):
                cell = sheet.cell(row=row_idx, column=col_idx)
                if not re.search(r"(?i)1st\s*row", str(cell.value or "")):
                    continue

                applied = False
                for merged_range in sheet.merged_cells.ranges:
                    min_col, min_row, merge_max_col, max_row = merged_range.bounds
                    if min_col <= col_idx <= merge_max_col and min_row <= row_idx <= max_row:
                        if min_col <= 2 and merge_max_col <= 2:
                            top_left = sheet.cell(row=min_row, column=min_col)
                            top_left.fill = copy(header_fill)
                            applied = True
                        break
                if not applied:
                    cell.fill = copy(header_fill)

    def unmerge_body_label_ranges(sheet, max_col: int) -> None:
        for merged_range in list(sheet.merged_cells.ranges):
            min_col, min_row, merge_max_col, max_row = merged_range.bounds
            if min_col > 2 or merge_max_col > 2:
                continue
            top_left = sheet.cell(row=min_row, column=min_col)
            top_text = str(top_left.value or "")
            if re.search(r"(?i)1st\s*row", top_text):
                continue
            has_body_label = any(
                str(sheet.cell(row=row_idx, column=2).value or "").strip()
                for row_idx in range(min_row, max_row + 1)
            )
            if not has_body_label:
                continue
            value = top_left.value
            fill = copy(top_left.fill)
            font = copy(top_left.font)
            alignment = copy(top_left.alignment)
            number_format = top_left.number_format
            border = copy(top_left.border)
            sheet.unmerge_cells(str(merged_range))
            for row_idx in range(min_row, max_row + 1):
                cell = sheet.cell(row=row_idx, column=min_col)
                if row_idx == min_row and cell.value is None:
                    cell.value = value
                cell.fill = copy(fill)
                cell.font = copy(font)
                cell.alignment = copy(alignment)
                cell.number_format = number_format
                cell.border = copy(border)

    def is_sig_only_row(sheet, row_idx: int, max_col: int) -> bool:
        if str(sheet.cell(row=row_idx, column=1).value or "").strip():
            return False
        if str(sheet.cell(row=row_idx, column=2).value or "").strip():
            return False

        has_sig = False
        for col_idx in range(3, max_col + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            text = str(value or "").strip()
            if not text:
                continue
            if not re.fullmatch(r"[A-Z]+", text):
                return False
            has_sig = True
        return has_sig

    def merge_orphan_sig_rows(sheet, max_col: int) -> None:
        rows_to_delete: List[int] = []
        for row_idx in range(2, sheet.max_row + 1):
            if not is_sig_only_row(sheet, row_idx, max_col):
                continue
            previous_label = str(sheet.cell(row=row_idx - 1, column=2).value or "").strip()
            if not previous_label:
                continue
            append_sig_values_to_row(sheet, row_idx - 1, row_idx, max_col)
            rows_to_delete.append(row_idx)
        delete_rows_desc(sheet, rows_to_delete)
        normalize_multiline_rows(sheet, max_col)

    def normalize_row_legends(sheet) -> List[int]:
        legend_rows: List[int] = []
        for row in range(1, sheet.max_row + 1):
            for col in range(1, 3):
                cell = sheet.cell(row=row, column=col)
                if type(cell).__name__ == "MergedCell":
                    continue
                value = cell.value
                if value is None:
                    continue
                text = str(value)
                if re.search(r"(?i)count", text) and re.search(r"(?i)column\s*%", text):
                    if table_type != TABLE_MULTI_SHEET_NOT_SIG:
                        cell.value = "1st row:  Column %" if keep_mode == KEEP_PERCENT else "1st row:  Count"
                    legend_rows.append(row)
                    break
        return legend_rows

    def process_stacked_table_sheet(sheet, max_col: int) -> None:
        rows_to_delete: List[int] = []
        vertically_merged_following_rows: set[int] = set()
        for merged_range in sheet.merged_cells.ranges:
            min_col, min_row, merge_max_col, max_row = merged_range.bounds
            if min_col <= 2 and merge_max_col <= 2 and max_row > min_row:
                vertically_merged_following_rows.update(range(min_row + 1, max_row + 1))

        def has_data_values(row_idx: int) -> bool:
            for col_idx in range(3, max_col + 1):
                value = sheet.cell(row=row_idx, column=col_idx).value
                if value is not None and str(value).strip() != "":
                    return True
            return False

        # Auto Lychee speed-up (2026-10-01): openpyxl's sheet.max_row scans every cell on each call and
        # this loop asked for it several times per row (91% of the time on a One Sheet export). Inside the
        # loop it cannot change — cells are only read/created on rows <= max_row and rows are deleted
        # after the loop — so it is read once; the assert below proves that on every run.
        sheet_max_row = sheet.max_row
        row = 1
        while row < sheet_max_row:
            label = sheet.cell(row=row, column=2).value
            label_text = str(label).strip() if label is not None else ""
            if label_text == "":
                row += 1
                continue

            next_label = sheet.cell(row=row + 1, column=2).value
            next_label_text = str(next_label).strip() if next_label is not None else ""
            possible_sig_row = row + 2
            has_following_sig_row = (
                possible_sig_row <= sheet_max_row
                and str(sheet.cell(row=possible_sig_row, column=1).value or "").strip() == ""
                and str(sheet.cell(row=possible_sig_row, column=2).value or "").strip() == ""
                and possible_sig_row in vertically_merged_following_rows
            )
            if (
                row + 1 in vertically_merged_following_rows
                and next_label_text == ""
                and has_data_values(row)
                and not has_data_values(row + 1)
                and not has_following_sig_row
            ):
                for col in range(1, max_col + 1):
                    target_cell = sheet.cell(row=row, column=col)
                    source_cell = sheet.cell(row=row + 1, column=col)
                    target_cell.border = Border(
                        left=target_cell.border.left,
                        right=target_cell.border.right,
                        top=target_cell.border.top,
                        bottom=source_cell.border.bottom,
                    )
                rows_to_delete.append(row + 1)
                row += 2
                continue

            if next_label_text not in {"", label_text} or (
                not has_data_values(row + 1) and not has_following_sig_row
            ):
                row += 1
                continue

            pct_row = row + 1
            is_total_row = label_text.upper() == "TOTAL"
            sig_row = pct_row + 1
            has_sig_row = (
                sig_row <= sheet_max_row
                and str(sheet.cell(row=sig_row, column=1).value or "").strip() == ""
                and str(sheet.cell(row=sig_row, column=2).value or "").strip() == ""
            )
            last_group_row = sig_row if has_sig_row else pct_row

            keep_row = row
            pct_has_values = has_data_values(pct_row)
            if keep_mode == KEEP_PERCENT and not is_total_row and pct_has_values:
                keep_row = pct_row

            for col in range(1, max_col + 1):
                target_cell = sheet.cell(row=keep_row, column=col)
                bottom_source = sheet.cell(row=last_group_row, column=col)
                bottom_border = bottom_source.border.bottom if col <= 2 else target_cell.border.bottom
                target_cell.border = Border(
                    left=target_cell.border.left,
                    right=target_cell.border.right,
                    top=target_cell.border.top,
                    bottom=bottom_border,
                )

            if keep_mode == KEEP_PERCENT and not is_total_row and pct_has_values:
                rows_to_delete.append(row)
            else:
                rows_to_delete.append(pct_row)

            row = last_group_row + 1

        assert sheet.max_row == sheet_max_row, 'max_row changed inside the stacked-table loop'
        compact_rows_once(sheet, rows_to_delete)
        normalize_body_label_styles(sheet, max_col)
        if table_type in {
            TABLE_ONE_SHEET_SIG,
            TABLE_ONE_SHEET_NOT_SIG,
            TABLE_MULTI_SHEET_SIG,
            TABLE_MULTI_SHEET_NOT_SIG,
        }:
            normalize_one_sheet_header_fill(sheet, max_col)
        if table_type == TABLE_MULTI_SHEET_SIG:
            normalize_multi_sheet_sig_body_borders(sheet, max_col)
        if table_type in {TABLE_ONE_SHEET_NOT_SIG, TABLE_MULTI_SHEET_NOT_SIG}:
            normalize_single_row_table_borders(sheet, max_col)
        normalize_multiline_rows(sheet, max_col)

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        if sheet_name.strip().lower() == "contents":
            continue
        if sheet.max_row < 7:
            continue
        legend_rows = normalize_row_legends(sheet)
        if not legend_rows:
            continue

        max_col = detect_effective_max_col(sheet)

        if len(legend_rows) > 1:
            process_stacked_table_sheet(sheet, max_col)
            continue

        # Auto Lychee fix (2026-10-01): the original always began the table body at row 7 (Cross: the
        # "1st row: Count" legend is row 6). A Matrix export has a shorter header — legend row 4, TOTAL
        # merged over rows 6-7 — so that TOTAL was neither unmerged nor grouped and its 100 row stayed in
        # both the N and the % file. The body now starts right after the legend, never later than row 7
        # (Cross: unchanged, byte-identical to the original).
        first_row = min(7, min(legend_rows) + 1)
        merged_ranges = list(sheet.merged_cells.ranges)
        for m_range in merged_ranges:
            min_col, min_row, m_max_col, m_max_row = m_range.bounds
            if min_row < first_row:
                continue

            top_left_cell = sheet.cell(row=min_row, column=min_col)
            tl_val = top_left_cell.value
            tl_border = copy(top_left_cell.border)
            tl_fill = copy(top_left_cell.fill)
            tl_font = copy(top_left_cell.font)
            tl_alignment = copy(top_left_cell.alignment)
            tl_number_format = top_left_cell.number_format

            sheet.unmerge_cells(str(m_range))

            for r in range(min_row, m_max_row + 1):
                for c in range(min_col, m_max_col + 1):
                    cell = sheet.cell(row=r, column=c)
                    cell.value = tl_val
                    if tl_border:
                        cell.border = copy(tl_border)
                    if tl_fill:
                        cell.fill = copy(tl_fill)
                    if tl_font:
                        cell.font = copy(tl_font)
                    if tl_alignment:
                        cell.alignment = copy(tl_alignment)
                    cell.number_format = tl_number_format

        rows_to_delete_regular: List[int] = []
        pending_label_merges: List[Tuple[int, int]] = []

        def row_has_values(row_idx: int, start_col: int = 3) -> bool:
            for col_idx in range(start_col, max_col + 1):
                value = sheet.cell(row=row_idx, column=col_idx).value
                if value is not None and str(value).strip() != "":
                    return True
            return False

        row = first_row
        while row <= sheet.max_row:
            label = sheet.cell(row=row, column=2).value
            label_text = str(label).strip() if label is not None else ""
            if label_text == "":
                row += 1
                continue

            group_start = row
            group_end = row
            code_value = sheet.cell(row=row, column=1).value
            while group_end + 1 <= sheet.max_row:
                next_label = sheet.cell(row=group_end + 1, column=2).value
                next_code = sheet.cell(row=group_end + 1, column=1).value
                if next_label != label or next_code != code_value:
                    break
                group_end += 1

            group_len = group_end - group_start + 1
            is_total_group = label_text.upper() == "TOTAL"

            if is_total_group:
                if table_type == TABLE_MULTI_SHEET_SIG and group_len >= 3:
                    for remove_row in range(group_start + 1, group_end):
                        rows_to_delete_regular.append(remove_row)
                    pending_label_merges.append((group_start, group_end))
                    row = group_end + 1
                    continue

                for remove_row in range(group_start + 1, group_end + 1):
                    for col in range(1, max_col + 1):
                        target_cell = sheet.cell(row=group_start, column=col)
                        source_cell = sheet.cell(row=remove_row, column=col)
                        target_cell.border = Border(
                            left=target_cell.border.left,
                            right=target_cell.border.right,
                            top=target_cell.border.top,
                            bottom=source_cell.border.bottom,
                        )
                    rows_to_delete_regular.append(remove_row)
            elif group_len >= 3:
                # N/%/Sig tables: remove the row that does not match the selected output type.
                remove_row = group_start if keep_mode == KEEP_PERCENT else group_start + 1
                rows_to_delete_regular.append(remove_row)
                keep_row = group_start + 1 if keep_mode == KEEP_PERCENT else group_start
                if table_type == TABLE_MULTI_SHEET_SIG:
                    pending_label_merges.append((keep_row, group_end))
                    row = group_end + 1
                    continue

                for sig_row in range(group_start + 2, group_end + 1):
                    if row_has_values(sig_row):
                        append_sig_values_to_row(sheet, keep_row, sig_row, max_col)
                        rows_to_delete_regular.append(sig_row)
                    else:
                        for col in range(1, max_col + 1):
                            target_cell = sheet.cell(row=keep_row, column=col)
                            source_cell = sheet.cell(row=sig_row, column=col)
                            target_cell.border = Border(
                                left=target_cell.border.left,
                                right=target_cell.border.right,
                                top=target_cell.border.top,
                                bottom=source_cell.border.bottom,
                            )
                        rows_to_delete_regular.append(sig_row)
            elif group_len == 2:
                pct_row = group_start + 1
                for col in range(1, max_col + 1):
                    target_cell = sheet.cell(row=group_start, column=col)
                    source_cell = sheet.cell(row=pct_row, column=col)
                    target_cell.border = Border(
                        left=target_cell.border.left,
                        right=target_cell.border.right,
                        top=target_cell.border.top,
                        bottom=source_cell.border.bottom,
                    )

                if keep_mode == KEEP_PERCENT:
                    for col in range(3, max_col + 1):
                        target_cell = sheet.cell(row=group_start, column=col)
                        source_cell = sheet.cell(row=pct_row, column=col)
                        chosen_val, chosen_fmt = choose_value_and_format(target_cell, source_cell)
                        target_cell.value = chosen_val
                        target_cell.number_format = chosen_fmt

                rows_to_delete_regular.append(group_start if keep_mode == KEEP_PERCENT else pct_row)

            row = group_end + 1

        delete_rows_desc(sheet, rows_to_delete_regular)
        if table_type == TABLE_MULTI_SHEET_NOT_SIG:
            remove_consecutive_duplicate_labels(sheet)
        merge_orphan_sig_rows(sheet, max_col)
        unmerge_body_label_ranges(sheet, max_col)
        normalize_body_label_styles(sheet, max_col)
        if table_type in {
            TABLE_ONE_SHEET_SIG,
            TABLE_ONE_SHEET_NOT_SIG,
            TABLE_MULTI_SHEET_SIG,
            TABLE_MULTI_SHEET_NOT_SIG,
        }:
            normalize_one_sheet_header_fill(sheet, max_col)
        if table_type == TABLE_MULTI_SHEET_SIG:
            normalize_multi_sheet_sig_body_borders(sheet, max_col)
        if table_type in {TABLE_ONE_SHEET_NOT_SIG, TABLE_MULTI_SHEET_NOT_SIG}:
            normalize_single_row_table_borders(sheet, max_col)
        normalize_multiline_rows(sheet, max_col)

        deleted_rows = sorted(set(rows_to_delete_regular))

        def shifted_row(original_row: int) -> int:
            deleted_before = sum(1 for deleted_row in deleted_rows if deleted_row < original_row)
            return original_row - deleted_before

        for start_row, end_row in pending_label_merges:
            shifted_start = shifted_row(start_row)
            shifted_end = shifted_row(end_row)
            if shifted_end <= shifted_start:
                continue
            for col in range(1, 3):
                top_cell = sheet.cell(row=shifted_start, column=col)
                bottom_cell = sheet.cell(row=shifted_end, column=col)
                top_cell.border = Border(
                    left=top_cell.border.left,
                    right=top_cell.border.right,
                    top=top_cell.border.top,
                    bottom=bottom_cell.border.bottom,
                )
                sheet.merge_cells(
                    start_row=shifted_start,
                    start_column=col,
                    end_row=shifted_end,
                    end_column=col,
                )

        if table_type == TABLE_MULTI_SHEET_SIG:
            normalize_multi_sheet_sig_body_borders(sheet, max_col)

        last_row = 7
        for r in range(sheet.max_row, 6, -1):
            row_has_data = False
            for c in range(1, max_col + 1):
                v = sheet.cell(row=r, column=c).value
                if v is not None and str(v).strip() != "":
                    row_has_data = True
                    break
            if row_has_data:
                last_row = r
                break
        if table_type == TABLE_MULTI_SHEET_SIG:
            last_row = sheet.max_row

        thin_border = Side(border_style="thin", color="000000")
        for col in range(1, max_col + 1):
            cell = sheet.cell(row=last_row, column=col)
            cell.border = Border(
                left=cell.border.left,
                right=cell.border.right,
                top=cell.border.top,
                bottom=thin_border,
            )

    repair_internal_links(wb)
    wb.save(save_path)


def resolve_input_items(items: List[str]) -> Tuple[List[Path], List[str]]:
    valid_paths: List[Path] = []
    invalid_items: List[str] = []

    for raw in items:
        p = raw.strip().strip('"').strip("'")
        if not p:
            continue
        path_obj = Path(p)

        if path_obj.is_file() and path_obj.suffix.lower() == ".xlsx":
            valid_paths.append(path_obj)
        elif path_obj.is_dir():
            files = sorted(path_obj.glob("*.xlsx"))
            if files:
                valid_paths.extend(files)
            else:
                invalid_items.append(f"{p} (no .xlsx inside folder)")
        else:
            invalid_items.append(p)

    seen = set()
    deduped: List[Path] = []
    for p in valid_paths:
        key = str(p.resolve()).lower()
        if key not in seen:
            seen.add(key)
            deduped.append(p)

    return deduped, invalid_items



# ===================== Auto Lychee integration =====================

CUT_MODES = {"both": [KEEP_COUNT, KEEP_PERCENT], "count": [KEEP_COUNT], "percent": [KEEP_PERCENT]}


def output_stem_no_date(source_stem: str, keep_mode: str = KEEP_COUNT) -> str:
    """build_output_stem() without the date (user request): '<name> N' / '<name> %'.
    Same N% / trailing N-% normalisation as the original; the source name is otherwise kept."""
    stem = re.sub(r"_processed(?:_\d+)?$", "", source_stem, flags=re.IGNORECASE)
    output_token = "%" if keep_mode == KEEP_PERCENT else "N"
    stem = re.sub(r"(?i)\bN\s*%", output_token, stem)
    stem = re.sub(r"(?i)(?:\s+)(?:N|%)$", "", stem.strip()).rstrip()
    return f"{stem} {output_token}" if stem else output_token


def output_path_no_date(out_dir: Path, source_path: Path, reserved_names: set[str], keep_mode: str) -> Path:
    """unique_output_path_reserved() with output_stem_no_date(): '_1', '_2' … when the name exists."""
    base = output_stem_no_date(source_path.stem, keep_mode)
    candidate_name = f"{base}.xlsx"
    idx = 1
    while True:
        candidate = out_dir / candidate_name
        key = str(candidate.resolve()).lower()
        if key not in reserved_names and not candidate.exists():
            reserved_names.add(key)
            return candidate
        candidate_name = f"{base}_{idx}.xlsx"
        idx += 1


def process_to_files(file_path, keep_mode="both"):
    """Like the original tool's "Process All Files" with the output folder = the Banner's folder:
    the source file stays, and one new file per selected output (N / %) is written beside it,
    named '<name> N.xlsx' / '<name> %.xlsx' (no date, user request). One workbook at a time
    (no process pool) to keep CPU low. Each file is written to a temporary name first and
    renamed only when complete."""
    file_path = Path(file_path)
    if keep_mode not in CUT_MODES:
        raise ValueError(f"ตัด N/%: ไม่รู้จักโหมด {keep_mode}")
    reserved: set[str] = set()
    saved: list[Path] = []
    for mode in CUT_MODES[keep_mode]:
        target = output_path_no_date(file_path.parent, file_path, reserved, mode)
        pending = target.with_name(f".{target.stem}.cut_pending.xlsx")
        try:
            from fast_styles import exact_style_cache
            with exact_style_cache():  # same output bytes, faster openpyxl style handling
                process_workbook(file_path, pending, mode)
            os.replace(pending, target)
        finally:
            if pending.exists():
                pending.unlink()
        saved.append(target)
    return saved


# ====== MODULE: worker ======
# from __future__ import annotations  (applied to this section by the loader)
import json
import os
import sys
import threading
import time
from contextlib import contextmanager
import traceback
from datetime import datetime
from pathlib import Path
from core import Job, manual_items, output_name, validate_jobs


_emit_lock = threading.Lock()
POST_CHILD = False  # True in the `--post` child: its row updates must not move the table selection


def emit(event, **values):
    if POST_CHILD and event == 'status':
        values.setdefault('background', True)
    line = json.dumps({'event': event, **values}, ensure_ascii=False)
    with _emit_lock:  # the post-processing relay thread prints too
        print(line, flush=True)


@contextmanager
def screen_freeze(config):
    """Lyche shows a newly opened window before any other process can move it (12-70 ms measured),
    so for the few short steps that open one, ask the app to cover every screen with a still image
    of itself and wait for its acknowledgement (control_dir/frozen, max 2 s)."""
    if not config.get('background', True):  # Preview mode: the user wants to see Lyche
        yield
        return
    flag = Path(config['control_dir']) / 'frozen'
    flag.unlink(missing_ok=True)
    emit('freeze')
    end = time.monotonic() + 2
    while not flag.exists() and time.monotonic() < end:
        time.sleep(.02)
    try:
        yield
    finally:
        emit('unfreeze')
        flag.unlink(missing_ok=True)


def post_process(job, path):
    """After Lyche saved the Banner, in order: 1) Delete Total + NA (156_DeleteTotalNA.py, hidden
    Excel), 2) Del Sig (Del_Sig.py, openpyxl). Each saves back to the same file before the next
    queue row runs. 3) Cut N / % (151_CutLychee_Persence.py) then writes new N / % file(s) beside
    it, as the original tool does. Progress shows in the row (max 1/s)."""
    notes = []
    if job.total_na:
        from post_total_na import process_in_place as delete_total_na
        emit('status', id=job.id, status='กำลังรัน', detail='Delete Total + NA · เปิด Excel เบื้องหลัง')
        last = [0.0]

        def progress(detail):
            if time.monotonic() - last[0] >= 1:
                last[0] = time.monotonic()
                emit('status', id=job.id, status='กำลังรัน', detail=f'Delete Total + NA · {detail}')

        summary = delete_total_na(path, delete_empty_rows=job.total_na_empty_rows, progress=progress)
        emit('log', text=f'Delete Total + NA {path.name}: {summary}')
        notes.append(f'Total+NA: {summary}')
    if job.del_sig:
        from post_del_sig import process_in_place as del_sig
        emit('status', id=job.id, status='กำลังรัน', detail=f'Del Sig · {job.del_sig_groups}')
        summary = del_sig(path, job.del_sig_groups, job.del_sig_mode, job.del_sig_beside)
        emit('log', text=f'Del Sig {path.name}: {job.del_sig_groups} · {summary}')
        notes.append(f'Del Sig: {job.del_sig_groups} · {summary}')
    if job.cut_percent:
        from post_cut_percent import process_to_files as cut_percent
        label = CUT_LABELS.get(job.cut_percent_mode, job.cut_percent_mode)
        emit('status', id=job.id, status='กำลังรัน', detail=f'ตัด N/% · {label}')
        names = ', '.join(p.name for p in cut_percent(path, job.cut_percent_mode))
        emit('log', text=f'ตัด N/% {path.name} ({label}) → {names}')
        notes.append(f'ตัด N/% {label}: {names}')
    return ' | '.join(notes)


CUT_LABELS = {'both': 'N + %', 'count': 'N Only', 'percent': '% Only'}


def post_main():
    """`AutoLychee_OneFile.py --post` (exe: `… --post`): the post-processing child. Reads one
    task per stdin line, runs post_process() on it and reports the row's final status itself.
    A separate process so the CPU-heavy file work never competes with the Lyche driver (and its
    window hider) for the GIL, while Lyche already runs the next queue row."""
    global POST_CHILD
    POST_CHILD = True
    low_priority()
    for line in sys.stdin:
        if not line.strip():
            continue
        task = json.loads(line)
        job, path = Job(**task['job']), Path(task['path'])
        try:
            started = time.process_time()
            detail = task['detail'] + ' | ' + post_process(job, path)
            if os.environ.get('LYCHE_CPU_DEBUG'):
                emit('log', text=f'cpu post-processing {path.name}: {time.process_time() - started:.1f}s')
            emit('status', id=job.id, status='OK', detail=detail, base=task['base'])
        except Exception as exc:
            emit('status', id=job.id, status='ผิดพลาด', detail=f"{task['detail']} | หลังรันไม่สำเร็จ: {exc}",
                 base=task['base'])
            emit('log', text=f'หลังรัน {path.name} ไม่สำเร็จ: {exc}')
            try:
                log_dir = Path(task['log_dir'])
                log_dir.mkdir(parents=True, exist_ok=True)
                (log_dir / f'error-post-{datetime.now():%Y%m%d-%H%M%S}.txt').write_text(
                    traceback.format_exc(), encoding='utf-8')
            except OSError:
                pass
        emit('post_done', id=job.id)


class PostPipeline:
    """Parent side: one long-lived `--post` child, fed one task per exported Banner. Tasks run
    one at a time (1 CPU core, below-normal priority) in queue order."""
    def __init__(self, log_dir):
        import subprocess
        if getattr(sys, 'frozen', False):
            command = [sys.executable, '--post']
        else:
            command = [sys.executable, '-X', 'utf8', '-u', str(Path(__file__).resolve()), '--post']
        self.log_dir = str(log_dir)
        self.process = subprocess.Popen(
            command, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            cwd=str(Path(__file__).resolve().parent),
            creationflags=subprocess.CREATE_NO_WINDOW | subprocess.BELOW_NORMAL_PRIORITY_CLASS)
        self.job_object = self._kill_with_parent(self.process.pid)
        self.pending = {}  # row id -> (detail, base) of tasks sent but not finished
        self.dead = False
        self.condition = threading.Condition()
        threading.Thread(target=self._relay, name='post-relay', daemon=True).start()

    @staticmethod
    def _kill_with_parent(pid):
        """Put the child in a job object that is closed (and the child killed) when this worker
        exits for any reason, so no file processing outlives a stopped or killed run."""
        try:
            import win32api
            import win32con
            import win32job
            job = win32job.CreateJobObject(None, '')
            info = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
            info['BasicLimitInformation']['LimitFlags'] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, info)
            handle = win32api.OpenProcess(win32con.PROCESS_SET_QUOTA | win32con.PROCESS_TERMINATE, False, pid)
            win32job.AssignProcessToJobObject(job, handle)
            return job
        except Exception:
            return None

    def _relay(self):
        for raw in self.process.stdout:
            line = raw.decode('utf-8', errors='replace').rstrip('\r\n')
            try:
                event = json.loads(line)
            except ValueError:
                event = None
            if isinstance(event, dict) and event.get('event') == 'post_done':
                with self.condition:
                    self.pending.pop(event.get('id'), None)
                    self.condition.notify_all()
                continue
            if line:
                with _emit_lock:
                    print(line, flush=True)
        with self.condition:
            self.dead = True
            self.condition.notify_all()

    def submit(self, job, path, detail, base):
        from dataclasses import asdict
        with self.condition:
            if self.dead:
                raise RuntimeError('ตัวประมวลผลหลังรันหยุดทำงาน')
            self.pending[job.id] = (detail, base)
        task = {'job': asdict(job), 'path': str(path), 'detail': detail, 'base': base, 'log_dir': self.log_dir}
        self.process.stdin.write((json.dumps(task, ensure_ascii=False) + '\n').encode('utf-8'))
        self.process.stdin.flush()

    def wait(self, checkpoint):
        """Until every submitted row has finished post-processing (Stop / F8 still work)."""
        while True:
            with self.condition:
                if not self.pending:
                    return
                if self.dead:
                    for row_id, (detail, base) in self.pending.items():
                        emit('status', id=row_id, status='ผิดพลาด', base=base, background=True,
                             detail=f'{detail} | หลังรันไม่สำเร็จ: ตัวประมวลผลหยุดกลางคัน')
                    self.pending.clear()
                    raise RuntimeError('ตัวประมวลผลหลังรันหยุดกลางคัน')
                self.condition.wait(.5)
            checkpoint()

    def close(self, kill=False):
        try:
            if kill:
                self.process.kill()
            else:
                self.process.stdin.close()
                self.process.wait(30)
        except Exception:
            try:
                self.process.kill()
            except Exception:
                pass


def resolve_handle(config, windows, open_cross_tabulation):
    """The Cross Tabulation window is closed after each run, so a saved handle can be stale:
    find the window for the same project again, or open it (minimised) from the Lyche window."""
    import win32gui
    handle = config.get('handle')
    from driver import WINDOW_PREFIXES  # Cross Tabulation, the Matrix settings window, Tabulation Result
    if handle and win32gui.IsWindow(handle) and win32gui.GetWindowText(handle).startswith(WINDOW_PREFIXES):
        return handle
    title = config.get('window_title') or ''
    project = title[title.find('<'):title.find('>') + 1] if '<' in title else ''
    items = windows()
    opened = False
    if not items:
        emit('log', text='หน้าต่าง Cross Tabulation ปิดอยู่ — กำลังเปิดใหม่เบื้องหลัง')
        with screen_freeze(config):
            items = open_cross_tabulation(emit)
        opened = bool(items)
    emit('windows', items=items, opened=opened)
    for item in items:
        if not project or project in item['title']:
            return item['handle']
    raise ValueError('ไม่พบหน้าต่าง Cross Tabulation ของโปรเจกต์นี้ กรุณาเปิดโปรเจกต์ใน Lyche')


def cpu_debug(label):
    """With LYCHE_CPU_DEBUG=1, log this worker's CPU seconds (all threads) at phase ends."""
    if os.environ.get('LYCHE_CPU_DEBUG'):
        emit('log', text=f'cpu {label}: {time.process_time():.1f}s total, wall {time.monotonic() - STARTED:.1f}s')


STARTED = time.monotonic()


def low_priority():
    """Run the worker below normal priority so the user's own programs stay responsive
    (Del Sig / Delete Total + NA can use a full CPU core for a while)."""
    try:
        import win32api
        import win32process
        win32process.SetPriorityClass(win32api.GetCurrentProcess(), win32process.BELOW_NORMAL_PRIORITY_CLASS)
    except Exception:
        pass


def main():
    low_priority()
    config = json.loads(Path(sys.argv[1]).read_text(encoding='utf-8'))
    if config['action'] == 'banners' and config.get('source', 'Personal') == 'Personal':
        # Pure metadata read. Never enter the UI adapter for Personal Get Banner.
        try:
            from history import from_window
            # The window may have been closed after the last run; its title still names the project.
            names, path = from_window(config['handle'], config.get('source', 'Personal'), config.get('window_title'))
            emit('banners', names=names)
            emit('log', text=f'อ่าน Personal History จากไฟล์โดยไม่ใช้เมาส์: {path.name}')
            emit('done')
        except Exception as exc:
            emit('error', text=str(exc))
            sys.exit(1)
        return
    from driver import Lyche, Stopped, launchers, open_cross_tabulation, windows
    bot = None
    current = None
    post = None
    try:
        if config['action'] == 'windows':
            # Start-up / "ค้นหาหน้าต่าง": never open Cross Tabulation here (opening it can flash on
            # screen for a few ms). The project window is enough for Personal Get Banner; run, load
            # and Shared open Cross Tabulation in the background when they need it (resolve_handle).
            # One entry per open Lyche project: its Cross Tabulation window if open, else the project window.
            items, seen = [], set()
            for item in windows() + launchers():
                project = item['title'][item['title'].find('<'):item['title'].find('>') + 1]
                if project not in seen:
                    seen.add(project)
                    items.append(item)
            emit('windows', items=items, opened=False)
            return
        fresh = not windows()  # resolve_handle is about to open Cross Tabulation
        bot = Lyche(resolve_handle(config, windows, open_cross_tabulation), Path(config['control_dir']), emit,
                    config.get('timeout', 900), background=config.get('background', True))
        source = config.get('source', 'Personal')
        if config['action'] == 'banners':
            # Shared History exists only on the Lyche server: read it from the Import History list.
            with screen_freeze(config):  # Import History opens (off-screen) while it is read
                names = bot.get_banners(source)
            emit('banners', names=names, raise_app=True)
            emit('log', text='อ่าน Shared Tabulation จากหน้าต่าง Import History ของ Lyche')
        elif config['action'] == 'load':
            with bot.background_session():
                bot.load(config['banner'], source)
        elif config['action'] == 'check_items':
            # Banner Manual 'เช็คกับ Lyche': search the item list only (the Banner is not changed). UIA search
            # works with the window as it is (verified live), so nothing pops up (user request): a window on
            # screen is not moved at all; a minimised one (also one just opened) stays minimised, and the
            # parker hides it if Lyche shows it by itself — for 2 s more after a fresh open — then it is
            # minimised again.
            import win32gui
            if bot.background and win32gui.IsIconic(bot.handle):
                with bot.background_session(keep_minimised=True):
                    results = bot.check_items(config['items'])
                    if fresh:
                        time.sleep(2)
            else:
                results = bot.check_items(config['items'])
            emit('items_checked', results=results)
        elif config['action'] == 'run':
            jobs = [Job(**job) for job in config['jobs']]
            folder = Path(config['folder'])
            validate_jobs(jobs, folder)
            # Post-processing of row N runs in a child process while Lyche already works on row N+1.
            if any(job.status != 'OK' and (job.total_na or job.del_sig or job.cut_percent) for job in jobs):
                post = PostPipeline(Path(config['control_dir']).parent / 'logs')
            with bot.background_session():  # Lyche stays off-screen for the whole queue
                loaded = None
                for job in jobs:
                    if job.status == 'OK':
                        continue
                    current = job
                    emit('status', id=job.id, status='กำลังรัน', detail='กำลังโหลด Banner')
                    # Banner Manual changes the loaded Banner, so a row reuses what is loaded only when
                    # both the History and its manual items are the same as the previous row's.
                    wanted = (job.history, tuple(manual_items(job.banner_manual_items)) if job.banner_manual else ())
                    if wanted != loaded:
                        bot.load(job.history, source)
                        if wanted[1]:
                            emit('status', id=job.id, status='กำลังรัน', detail='Banner Manual')
                            bot.set_banner_items(list(wanted[1]))
                        loaded = wanted
                    base = bot.set_filter(job.filter, job.base)
                    emit('status', id=job.id, status='กำลังรัน', detail='Tabulate / Export')
                    path = folder / output_name(job.output)
                    # the row's last Step; Banner Manual rows export with Analysis Axis only (user rule:
                    # a One Sheet export of a manual Banner said Saved but left no file)
                    bot.run_export(path, one_sheet=job.export_mode == 'onesheet' and not job.banner_manual)
                    detail = f'Saved | {path}'
                    if bot.matrix and job.del_sig and job.del_sig_mode != 'MATRIX':
                        job.del_sig_mode = 'MATRIX'  # user rule: a Matrix History → Del Sig in Matrix mode
                        emit('log', text=f'{job.output}: Banner แบบ Matrix → Del Sig ใช้ประเภท Matrix')
                    base = '' if base is None else str(base)
                    cpu_debug('Lyche steps')
                    if job.total_na or job.del_sig or job.cut_percent:
                        emit('status', id=job.id, status='กำลังรัน', detail='Saved · รอทำขั้นหลังรัน', base=base)
                        post.submit(job, path, detail, base)
                    else:
                        emit('status', id=job.id, status='OK', detail=detail, base=base)
                    current = None
                try:  # user request: close the Cross Tabulation window once the queue is done
                    bot.close_window()
                except Exception as exc:
                    emit('log', text=f'ปิดหน้าต่าง Cross Tabulation ไม่ได้: {exc}')
            if post:
                if post.pending:
                    emit('log', text=f'Lyche รันครบแล้ว — รอขั้นหลังรันอีก {len(post.pending)} แถว')
                post.wait(bot.checkpoint)
                post.close()
                post = None
            emit('log', text='เสร็จครบทุกแถวที่รอรัน')
        else:
            raise ValueError('Unknown action')
        emit('done')
    except Stopped as exc:
        if post:
            post.close(kill=True)
        if current:
            emit('status', id=current.id, status='หยุด', detail=str(exc))
        emit('stopped', text=str(exc))
    except Exception as exc:
        if current:
            emit('status', id=current.id, status='ผิดพลาด', detail=str(exc))
        if post:  # rows Lyche already exported still finish their post-processing, as before
            try:
                post.wait(bot.checkpoint if bot else lambda: None)
            except Exception:
                pass
            post.close(kill=True)
        log_dir = Path(config['control_dir']).parent / 'logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime('%Y%m%d-%H%M%S')
        (log_dir / f'error-{stamp}.txt').write_text(traceback.format_exc(), encoding='utf-8')
        if bot:
            try:
                bot.main().capture_as_image().save(log_dir / f'error-{stamp}.png')
            except Exception:
                pass
        emit('error', text=str(exc))
        sys.exit(1)


# ONEFILE: `--worker request.json` and `--post` are dispatched by the loader at the top of this file.


# ====== MODULE: app ======
# from __future__ import annotations  (applied to this section by the loader)
import csv
import io
import json
import os
import re
import sys
from dataclasses import asdict
from datetime import datetime
from html import escape
from pathlib import Path
from uuid import uuid4

import onefile  # ONEFILE: the loader at the top of this file (paths, how to start the worker)

FROZEN = onefile.FROZEN  # running as an exe built with --build-exe

from PySide6.QtCore import Qt, QEvent, QProcess, QTimer, QLockFile, QLocale, QRect, QSize
from PySide6.QtGui import QColor, QCursor, QFont, QFontMetrics, QIcon, QKeySequence, QPainter, QPen, QShortcut, QTextCursor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QLineEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QMessageBox, QPlainTextEdit, QSplitter,
    QAbstractItemView, QSpinBox, QDialog, QFrame, QGraphicsDropShadowEffect, QStyledItemDelegate, QStyle,
    QButtonGroup, QCheckBox, QRadioButton, QStyleOptionViewItem,
)
from chrome import MacWindowMixin
from core import Job, manual_items, save_json, read_jobs, validate_jobs

# ONEFILE: queue/settings/logs and the icons live in onefile.DATA (%LOCALAPPDATA%\AutoLychee\OneFile).
ROOT = onefile.SINGLE_FILE.parent
DATA = onefile.DATA
QUEUE = DATA / 'last_queue.json'
QUEUE_DIR = Path.home() / 'Documents'  # default folder for Save/Open queue
ASSETS = onefile.ASSETS.as_posix()
APP_NAME = 'Auto Lychee'
LOGO = onefile.ASSETS / 'logo.ico'

# Lyche-Epoch tone: ribbon blue canvas, Banner blue + Stub purple accents, Tabulate green.
STYLE = '''
* { font-family: "Segoe UI Variable Text", "Segoe UI", "Leelawadee UI"; font-size: 13px; color: #1e2a44; }
QMainWindow { background: transparent; }
QWidget#root { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f7faff, stop:0.18 #e9f0fa, stop:1 #dfe8f5); }
QDialog { background: transparent; }
QFrame#sheet { background: #ffffff; border: 1px solid #c9d6ea; border-radius: 18px; }
QFrame#card { background: #ffffff; border: 1px solid #c9d6ea; border-radius: 12px; }
QLabel { background: transparent; }
QLabel#windowTitle { color: #5f6f8a; font-size: 12px; font-weight: 600; }
QLabel#title { font-size: 26px; font-weight: 700; color: #1b3fd0; }
QLabel#subtitle, QLabel#caption { color: #5f6f8a; }
QLabel#caption { font-size: 12px; }
QLabel#section { font-size: 15px; font-weight: 700; color: #1b3fd0; }
QLabel#status { background: #e1eefc; border: 1px solid #b9d5f6; border-radius: 13px; padding: 5px 14px; color: #1f6fd1; font-size: 12px; font-weight: 600; }

QPushButton#tl_close, QPushButton#tl_min, QPushButton#tl_max { min-width: 13px; max-width: 13px; min-height: 13px; max-height: 13px; border-radius: 6px; padding: 0; font-size: 9px; font-weight: 800; color: rgba(0, 0, 0, 0.55); }
QPushButton#tl_close { background: #ff5f57; border: 1px solid #e2463f; }
QPushButton#tl_min { background: #febc2e; border: 1px solid #e1a116; }
QPushButton#tl_max { background: #28c840; border: 1px solid #14ae2c; }
QPushButton#tl_close[inactive="true"], QPushButton#tl_min[inactive="true"], QPushButton#tl_max[inactive="true"] { background: #d3dae6; border: 1px solid #c3cbd9; }

QPushButton { border: none; border-radius: 8px; padding: 7px 15px; font-weight: 600; }
QPushButton#secondary { background: #e6edf7; color: #2d3f5f; }
QPushButton#secondary:hover { background: #d9e4f3; }
QPushButton#blue { background: #2a8de9; color: #ffffff; }
QPushButton#blue:hover { background: #3b99ef; }
QPushButton#purple { background: #ece5fb; color: #5b34b8; }
QPushButton#purple:hover { background: #e0d5f8; }
QPushButton#gold { background: #fff4cc; color: #8a6a00; }
QPushButton#gold:hover { background: #ffecaa; }
QPushButton#amber { background: #fff0d9; color: #b26a00; }
QLabel#modeLabel { color: #5f6f8a; font-size: 12px; }
QCheckBox { spacing: 10px; }
QCheckBox::indicator { width: 18px; height: 18px; border: 2px solid #9fb8dc; border-radius: 5px; background: #ffffff; }
QCheckBox::indicator:hover { border-color: #2a8de9; }
QCheckBox::indicator:checked { background: #2a8de9; border-color: #2a8de9; image: url(ASSETS/check.svg); }
QCheckBox::indicator:disabled { background: #eef2f8; border-color: #d5e0ef; }
QCheckBox:disabled { color: #a9b4c6; }
QRadioButton { spacing: 8px; }
QRadioButton::indicator { width: 16px; height: 16px; border: 2px solid #9fb8dc; border-radius: 10px; background: #ffffff; }
QRadioButton::indicator:hover { border-color: #2a8de9; }
QRadioButton::indicator:checked { border: 2px solid #2a8de9; background: qradialgradient(cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5, stop:0 #2a8de9, stop:0.5 #2a8de9, stop:0.56 #ffffff, stop:1 #ffffff); }
QRadioButton::indicator:checked:disabled { border-color: #c3d0e4; background: qradialgradient(cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5, stop:0 #c3d0e4, stop:0.5 #c3d0e4, stop:0.56 #eef2f8, stop:1 #eef2f8); }
QRadioButton::indicator:disabled { border-color: #d5e0ef; background: #eef2f8; }
QRadioButton:disabled { color: #a9b4c6; }
QPushButton#segLeft, QPushButton#segRight { background: #e6edf7; color: #5f6f8a; padding: 7px 14px; font-weight: 700; border: 1px solid #c9d6ea; }
QPushButton#segLeft { border-top-right-radius: 0; border-bottom-right-radius: 0; border-right: none; }
QPushButton#segRight { border-top-left-radius: 0; border-bottom-left-radius: 0; }
QPushButton#segLeft:checked { background: #1b3fd0; color: #ffffff; border-color: #1b3fd0; }
QPushButton#segRight:checked { background: #6c3fc8; color: #ffffff; border-color: #6c3fc8; }
QPushButton#segLeft:disabled, QPushButton#segRight:disabled { background: #eef2f8; color: #a9b4c6; }
QPushButton#segLeft:checked:disabled, QPushButton#segRight:checked:disabled { background: #9fb0d8; color: #ffffff; }
QPushButton#amber:hover { background: #ffe5bf; }
QPushButton#red { background: #fde4e4; color: #c62828; }
QPushButton#red:hover { background: #fbd3d3; }
QPushButton#plain { background: transparent; color: #1a5fd0; padding: 7px 10px; text-decoration: underline; }
QPushButton#plain:hover { background: #e8f1fc; }
QPushButton#icon { background: transparent; min-width: 30px; max-width: 30px; padding: 6px 0; }
QPushButton#icon:hover { background: #e8f1fc; }
QPushButton#primary { color: #ffffff; padding: 9px 22px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2a8de9, stop:1 #6c3fc8); }
QPushButton#primary:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #3b99ef, stop:1 #7a4fd4); }
QPushButton#primary:pressed { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1f7ad2, stop:1 #5b34b8); }
QPushButton:disabled, QPushButton#primary:disabled, QPushButton#blue:disabled, QPushButton#purple:disabled, QPushButton#gold:disabled,
QPushButton#amber:disabled, QPushButton#red:disabled, QPushButton#secondary:disabled { background: #eef2f8; color: #a9b4c6; }
QPushButton#plain:disabled, QPushButton#icon:disabled { background: transparent; color: #b3bdcc; }

QLineEdit, QComboBox, QSpinBox { background: #ffffff; border: 1px solid #c3d0e4; border-radius: 7px; padding: 6px 10px; min-height: 20px; selection-background-color: #cfe3fb; selection-color: #1e2a44; }
QLineEdit:hover, QComboBox:hover, QSpinBox:hover { border: 1px solid #9fb8dc; }
QLineEdit:focus, QComboBox:focus, QSpinBox:focus { border: 1px solid #2a8de9; }
QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled { background: #f3f6fb; color: #8b97ab; }
QComboBox::drop-down { border: none; width: 26px; }
QComboBox::down-arrow { image: url(ASSETS/chevron-down.svg); width: 12px; height: 12px; }
QComboBox QAbstractItemView { background: #ffffff; border: 1px solid #c9d6ea; border-radius: 7px; padding: 4px; outline: 0; selection-background-color: #2a8de9; selection-color: #ffffff; }
QSpinBox::up-button, QSpinBox::down-button { border: none; width: 20px; background: transparent; }
QSpinBox::up-arrow { image: url(ASSETS/chevron-up.svg); width: 10px; height: 10px; }
QSpinBox::down-arrow { image: url(ASSETS/chevron-down.svg); width: 10px; height: 10px; }

QTableWidget { background: #ffffff; border: 1px solid #d5e0ef; border-radius: 8px; outline: 0; selection-background-color: #dbe9fb; selection-color: #1e2a44; }
QTableWidget::item { border-bottom: 1px solid #edf2f9; padding-left: 8px; }
QTableWidget::item:selected { background: #dbe9fb; color: #1e2a44; }
QTableWidget QComboBox { border: none; background: transparent; padding-left: 8px; }
QTableWidget QComboBox QLineEdit { border: none; background: transparent; padding: 0; }
QTableWidget QLineEdit { border: 1px solid #2a8de9; border-radius: 5px; padding: 2px 6px; }
QHeaderView { background: #2a8de9; border: none; border-top-left-radius: 8px; border-top-right-radius: 8px; }
QHeaderView::section { background: #2a8de9; color: #ffffff; border: none; border-right: 1px solid #4ea1ee; padding: 7px 8px; font-size: 12px; font-weight: 700; }
QHeaderView::section:last { border-right: none; }
QTableCornerButton::section { background: #2a8de9; border: none; }

QPlainTextEdit#log { background: #f4f7fc; border: 1px solid #d5e0ef; border-top: 2px solid #1b3fd0; border-radius: 8px; padding: 8px 10px; color: #4a5a78; font-family: "Cascadia Mono", Consolas, "Leelawadee UI"; font-size: 12px; }
QSplitter::handle { background: transparent; }
QScrollBar:vertical { background: transparent; width: 10px; margin: 2px; }
QScrollBar::handle:vertical { background: #c3d0e4; border-radius: 3px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #9fb8dc; }
QScrollBar:horizontal { background: transparent; height: 10px; margin: 2px; }
QScrollBar::handle:horizontal { background: #c3d0e4; border-radius: 3px; min-width: 30px; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
QScrollBar::add-page, QScrollBar::sub-page { background: transparent; }
QToolTip { background: #1e2a44; color: #ffffff; border: none; border-radius: 6px; padding: 5px 8px; }
QMessageBox { background: #eef3fa; }
'''.replace('ASSETS', ASSETS)


class FreezeOverlay(QWidget):
    """A still image of one screen, always on top and transparent to input. Shown for ~1-2 s while
    Lyche opens a window in the background, so its brief appearance is never seen."""
    def __init__(self, screen):
        super().__init__(None, Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
                         | Qt.WindowDoesNotAcceptFocus | Qt.WindowTransparentForInput)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.shot = screen.grabWindow(0)
        self.setGeometry(screen.geometry())

    def paintEvent(self, event):
        QPainter(self).drawPixmap(self.rect(), self.shot)


# Per-row settings kept on the row's column-0 item: the Steps, plus Banner Manual (its own column/dialog).
MANUAL_KEYS = ('banner_manual', 'banner_manual_items')
POST_KEYS = ('total_na', 'total_na_empty_rows', 'del_sig', 'del_sig_groups', 'del_sig_mode', 'del_sig_beside',
             'cut_percent', 'cut_percent_mode', 'export_mode') + MANUAL_KEYS
CUT_MODES = (('both', 'N + %'), ('count', 'N Only'), ('percent', '% Only'))
EXPORT_MODES = (('sheets', 'แยกชีท'), ('onesheet', 'One Sheet'))


def post_settings(job):
    """The post-processing part of a Job as a plain dict (stored on the row's column-0 item)."""
    post = {key: getattr(job, key) for key in POST_KEYS}
    post['export_mode'] = export_mode(post)  # e.g. an imported queue with Banner Manual + One Sheet
    return post


# Step 1 Delete Total + NA, Step 2 Del Sig, Step 3 Cut N / %, then Export — always the LAST Step (user
# rule): a new Step goes before it (new column before the last, new card before the Export card).
STEP_COLUMNS = (6, 7, 8, 9)
STEP_NAMES = ('Delete Total + NA', 'Del Sig (ตัด Sig)', 'ตัด N / %', 'Export')
# Banner Manual is appended after the Steps (so no other column index changes) and shown next to Banner.
# A new Step would take this index: move Banner Manual to the end then.
MANUAL_COLUMN = 6 + len(STEP_COLUMNS)
GEAR_COLUMN = MANUAL_COLUMN + 1  # one gear per row (shown last): opens the row's Step settings


def export_mode(post):
    """The row's Step 4 export. Banner Manual rows always export separated sheets (Export with Analysis
    Axis): user rule 2026-10-01 — a One Sheet export of a manual Banner said Saved but left no file."""
    return 'sheets' if post.get('banner_manual') else post.get('export_mode', 'sheets')


def manual_text(post):
    """Banner Manual cell text (None when off)."""
    items = manual_items(post.get('banner_manual_items', '')) if post.get('banner_manual') else []
    return ', '.join(items) or None


def step_texts(post):
    """Text (None when off) per Step column, each showing only its own setting; Export is never off."""
    step1 = ('Total+NA' + (' · แถวว่าง' if post.get('total_na_empty_rows') else '')) if post.get('total_na') else None
    step2 = None
    if post.get('del_sig'):
        step2 = 'Sig ' + (post.get('del_sig_groups') or '?')
        if post.get('del_sig_mode') == 'MATRIX':
            step2 += ' · Matrix'
        if post.get('del_sig_beside'):
            step2 += ' · ข้าง'
    step3 = dict(CUT_MODES).get(post.get('cut_percent_mode'), 'N + %') if post.get('cut_percent') else None
    export = dict(EXPORT_MODES).get(export_mode(post), 'แยกชีท')
    return step1, step2, step3, export


def post_summary(post):
    parts = []
    if post.get('total_na'):
        parts.append('Total+NA' + (' · แถวว่าง' if post.get('total_na_empty_rows') else ''))
    if post.get('del_sig'):
        parts.append('Sig ' + (post.get('del_sig_groups') or '?') + (' · ข้าง' if post.get('del_sig_beside') else ''))
    if post.get('cut_percent'):
        parts.append(dict(CUT_MODES).get(post.get('cut_percent_mode'), 'N + %'))
    if export_mode(post) == 'onesheet':
        parts.append('Export One Sheet')
    return ' + '.join(parts) or 'Off'


def sheet_layout(dialog, width):
    """Frameless rounded sheet with a soft shadow (the settings dialogs' look); returns its layout."""
    dialog.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
    dialog.setAttribute(Qt.WA_TranslucentBackground)
    dialog.setFixedWidth(width)
    shell = QVBoxLayout(dialog)
    shell.setContentsMargins(20, 20, 20, 20)
    sheet = QFrame()
    sheet.setObjectName('sheet')
    shadow = QGraphicsDropShadowEffect(sheet)
    shadow.setBlurRadius(36)
    shadow.setOffset(0, 8)
    shadow.setColor(QColor(27, 63, 208, 60))
    sheet.setGraphicsEffect(shadow)
    shell.addWidget(sheet)
    layout = QVBoxLayout(sheet)
    layout.setContentsMargins(28, 24, 28, 22)
    layout.setSpacing(10)
    return layout


def dialog_buttons(dialog, layout):
    """'ใช้กับทุกแถว' · 'ยกเลิก' · 'ตกลง' row shared by the settings dialogs."""
    buttons = QHBoxLayout()
    buttons.setSpacing(8)
    all_rows = QPushButton('ใช้กับทุกแถว')
    all_rows.setObjectName('secondary')
    cancel = QPushButton('ยกเลิก')
    cancel.setObjectName('secondary')
    ok = QPushButton('ตกลง')
    ok.setObjectName('primary')
    ok.setDefault(True)
    for button in (all_rows, cancel, ok):
        button.setCursor(Qt.PointingHandCursor)
    all_rows.clicked.connect(dialog.accept_all)
    cancel.clicked.connect(dialog.reject)
    ok.clicked.connect(dialog.accept)
    buttons.addWidget(all_rows)
    buttons.addStretch(1)
    buttons.addWidget(cancel)
    buttons.addWidget(ok)
    layout.addLayout(buttons)


class PostProcessDialog(QDialog):
    """Settings for what runs on a Banner's Excel file after Lyche exported it:
    1) Delete Total + NA (156_DeleteTotalNA.py), 2) Del Sig (Del_Sig.py), 3) Cut N / %
    (151_CutLychee_Persence.py)."""
    def __init__(self, parent, name, settings):
        super().__init__(parent)
        self.apply_all = False
        layout = sheet_layout(self, 580)
        heading = QLabel('ตั้งค่าหลังรัน')
        heading.setStyleSheet('font-size: 19px; font-weight: 700; color: #1b3fd0;')
        layout.addWidget(heading)
        target = QLabel(escape(name) + f'  ·  Step 1 → {len(STEP_COLUMNS) - 1} ทำหลัง Lyche Save ไฟล์ · '
                                       f'Step {len(STEP_COLUMNS)} = วิธี Export จาก Lyche')
        target.setStyleSheet('color: #5f6f8a;')
        layout.addWidget(target)
        layout.addSpacing(4)
        card_style = ('QFrame { background: #f4f7fc; border: 1px solid #d5e0ef; border-radius: 12px; } '
                      'QLabel, QCheckBox, QRadioButton { border: none; background: transparent; }')

        # 1) Delete Total + NA
        card = QFrame()
        card.setStyleSheet(card_style)
        inner = QVBoxLayout(card)
        inner.setContentsMargins(16, 14, 16, 14)
        inner.setSpacing(8)
        self.enabled = QCheckBox('1  Delete Total + NA')
        self.enabled.setStyleSheet('font-size: 15px; font-weight: 700;')
        self.enabled.setChecked(bool(settings.get('total_na')))
        inner.addWidget(self.enabled)
        what = QLabel('เปิดไฟล์ด้วย Excel เบื้องหลัง · ยกเลิก Merge แถวหัวตาราง · ลบคอลัมน์ TOTAL ที่ซ้ำ\n'
                      '· ลบคอลัมน์และแถว NA · บันทึกทับไฟล์เดิม')
        what.setStyleSheet('color: #4a5a78; font-size: 12px;')
        inner.addWidget(what)
        self.empty_rows = QCheckBox('ลบแถวที่ไม่มี Label และ Frequency ทุกคอลัมน์เป็น 0 หรือว่าง')
        self.empty_rows.setChecked(bool(settings.get('total_na_empty_rows')))
        inner.addWidget(self.empty_rows)
        layout.addWidget(card)
        self.enabled.toggled.connect(self.empty_rows.setEnabled)
        self.empty_rows.setEnabled(self.enabled.isChecked())

        # 2) Del Sig — same options as Del_Sig.py
        sig_card = QFrame()
        sig_card.setStyleSheet(card_style)
        sig = QVBoxLayout(sig_card)
        sig.setContentsMargins(16, 14, 16, 14)
        sig.setSpacing(8)
        self.del_sig = QCheckBox('2  Del Sig (ตัด Sig)')
        self.del_sig.setStyleSheet('font-size: 15px; font-weight: 700;')
        self.del_sig.setChecked(bool(settings.get('del_sig')))
        sig.addWidget(self.del_sig)
        self.sig_options = QWidget()
        options = QVBoxLayout(self.sig_options)
        options.setContentsMargins(0, 0, 0, 0)
        options.setSpacing(8)
        def radio_row(label, first, second, second_on):
            row = QHBoxLayout()
            row.setSpacing(18)
            caption = QLabel(label)
            caption.setFixedWidth(110)
            caption.setStyleSheet('color: #4a5a78; font-size: 12px; font-weight: 600;')
            row.addWidget(caption)
            a, b = QRadioButton(first), QRadioButton(second)
            group = QButtonGroup(self)
            group.addButton(a)
            group.addButton(b)
            (b if second_on else a).setChecked(True)
            row.addWidget(a)
            row.addWidget(b)
            row.addStretch(1)
            options.addLayout(row)
            return a, b
        self.mode_normal, self.mode_matrix = radio_row('ประเภท Crosstab', 'Crosstab ธรรมดา', 'Matrix',
                                                       settings.get('del_sig_mode', 'NORMAL') == 'MATRIX')
        self.sig_normal, self.sig_beside = radio_row('รูปแบบ Sig', 'Sig ปกติ', 'Sig ข้าง',
                                                     bool(settings.get('del_sig_beside')))
        self.sig_beside.setToolTip('ยก Sig ขึ้นบรรทัดเดียวกับตัวเลข และลบแถว Sig เดิม')
        groups_label = QLabel('ใส่ Sig ตรงนี้')
        groups_label.setStyleSheet('color: #4a5a78; font-size: 12px; font-weight: 600;')
        options.addWidget(groups_label)
        self.sig_groups = QLineEdit(settings.get('del_sig_groups', ''))
        self.sig_groups.setPlaceholderText('เช่น ABCDEF,GHIJKL,MNOPQR,STUVWX,YZ')
        options.addWidget(self.sig_groups)
        sig.addWidget(self.sig_options)
        layout.addWidget(sig_card)
        self.del_sig.toggled.connect(self.sig_options.setEnabled)
        self.sig_options.setEnabled(self.del_sig.isChecked())

        # 3) Cut N / % — a single compact row, same choices as 151_CutLychee_Persence.py
        cut_card = QFrame()
        cut_card.setStyleSheet(card_style)
        cut = QHBoxLayout(cut_card)
        cut.setContentsMargins(16, 10, 16, 10)
        cut.setSpacing(16)
        self.cut_percent = QCheckBox('3  ตัด N / %')
        self.cut_percent.setStyleSheet('font-size: 15px; font-weight: 700;')
        self.cut_percent.setChecked(bool(settings.get('cut_percent')))
        self.cut_percent.setToolTip('สร้างไฟล์ใหม่ข้างไฟล์ Banner (ชื่อ “ชื่อไฟล์ N” / “ชื่อไฟล์ %”) ไฟล์เดิมยังอยู่')
        cut.addWidget(self.cut_percent)
        cut.addStretch(1)
        self.cut_group = QButtonGroup(self)
        current = settings.get('cut_percent_mode', 'both')
        self.cut_radios = {}
        for key, label in CUT_MODES:
            radio = QRadioButton(label)
            self.cut_group.addButton(radio)
            radio.setChecked(key == current)
            radio.setEnabled(self.cut_percent.isChecked())
            self.cut_percent.toggled.connect(radio.setEnabled)
            self.cut_radios[key] = radio
            cut.addWidget(radio)
        if self.cut_group.checkedButton() is None:
            self.cut_radios['both'].setChecked(True)
        layout.addWidget(cut_card)

        # Export — always the LAST Step (user rule): add new Steps above this card. One compact row.
        export_card = QFrame()
        export_card.setStyleSheet(card_style)
        export = QHBoxLayout(export_card)
        export.setContentsMargins(16, 10, 16, 10)
        export.setSpacing(16)
        export_label = QLabel(f'{len(STEP_COLUMNS)}  Export')
        export_label.setStyleSheet('font-size: 15px; font-weight: 700;')
        export_label.setToolTip('แยกชีท: Cross = Export with Analysis Axis · Matrix = Export all → Excel(Separated Sheets)\n'
                                'One Sheet: ทั้ง Cross และ Matrix = Export all → Excel(One Sheet)')
        export.addWidget(export_label)
        manual = bool(settings.get('banner_manual'))
        if manual:
            locked = QLabel('Banner Manual → แยกชีทเท่านั้น')
            locked.setStyleSheet('color: #8b97ab; font-size: 12px;')
            export.addWidget(locked)
        export.addStretch(1)
        self.export_group = QButtonGroup(self)
        current = export_mode(settings)
        self.export_radios = {}
        for key, label in EXPORT_MODES:
            radio = QRadioButton(label)
            self.export_group.addButton(radio)
            radio.setChecked(key == current)
            self.export_radios[key] = radio
            export.addWidget(radio)
        if self.export_group.checkedButton() is None:
            self.export_radios['sheets'].setChecked(True)
        if manual:  # Banner Manual rows: Export with Analysis Axis only (user rule)
            self.export_radios['onesheet'].setEnabled(False)
            self.export_radios['onesheet'].setToolTip('แถวที่ใช้ Banner Manual Export ได้แบบแยกชีท (Export with Analysis Axis) เท่านั้น')
        layout.addWidget(export_card)

        self.message = QLabel('')
        self.message.setStyleSheet('color: #c62828; font-size: 12px;')
        self.message.hide()
        layout.addWidget(self.message)
        note = QLabel('Delete Total + NA ต้องมี Microsoft Excel ในเครื่อง · Del Sig และตัด N / % ปิดไว้เป็นค่าเริ่มต้น')
        note.setStyleSheet('color: #8b97ab; font-size: 12px;')
        layout.addWidget(note)
        layout.addSpacing(6)
        dialog_buttons(self, layout)

    @property
    def settings(self):
        on = self.enabled.isChecked()
        sig_on = self.del_sig.isChecked()
        return {'total_na': on, 'total_na_empty_rows': on and self.empty_rows.isChecked(),
                'del_sig': sig_on, 'del_sig_groups': self.sig_groups.text().strip().upper(),
                'del_sig_mode': 'MATRIX' if self.mode_matrix.isChecked() else 'NORMAL',
                'del_sig_beside': self.sig_beside.isChecked(),
                'cut_percent': self.cut_percent.isChecked(),
                'cut_percent_mode': next(k for k, r in self.cut_radios.items() if r.isChecked()),
                'export_mode': next(k for k, r in self.export_radios.items() if r.isChecked())}

    def accept(self):
        if self.del_sig.isChecked() and not re.search(r'[A-Za-z]', self.sig_groups.text()):
            self.message.setText('เปิด Del Sig แล้ว กรุณาใส่กลุ่ม Sig เช่น ABCDEF,GHIJKL')
            self.message.show()
            self.sig_groups.setFocus()
            self.apply_all = False
            return
        super().accept()

    def accept_all(self):
        self.apply_all = True
        self.accept()


class BannerManualDialog(QDialog):
    """Banner Manual of a row: Lyche item codes that replace the History's Banner (in this order).
    The run does what a user would: Banner 'Clear all' → Yes, then per item: search → select → To Banner."""
    def __init__(self, parent, name, settings, checker=None):
        super().__init__(parent)
        self.apply_all = False
        self.checker = checker  # App.check_manual_items(items, callback): 'เช็คกับ Lyche'
        layout = sheet_layout(self, 560)
        heading = QLabel('Banner Manual')
        heading.setStyleSheet('font-size: 19px; font-weight: 700; color: #1b3fd0;')
        layout.addWidget(heading)
        target = QLabel(escape(name) + '  ·  ใส่ข้อเป็น Banner แทน Banner จาก History (Stub ใช้ของ History เหมือนเดิม)')
        target.setStyleSheet('color: #5f6f8a;')
        target.setWordWrap(True)
        layout.addWidget(target)
        layout.addSpacing(4)
        card = QFrame()
        card.setStyleSheet('QFrame { background: #f4f7fc; border: 1px solid #d5e0ef; border-radius: 12px; } '
                           'QLabel, QCheckBox { border: none; background: transparent; }')
        inner = QVBoxLayout(card)
        inner.setContentsMargins(16, 14, 16, 14)
        inner.setSpacing(8)
        self.enabled = QCheckBox('ใช้ Banner Manual')
        self.enabled.setStyleSheet('font-size: 15px; font-weight: 700;')
        self.enabled.setChecked(bool(settings.get('banner_manual')))
        inner.addWidget(self.enabled)
        caption = QLabel('ข้อที่จะใส่ เรียงตามลำดับ Banner · หลายข้อคั่นด้วย , หรือขึ้นบรรทัดใหม่')
        caption.setStyleSheet('color: #4a5a78; font-size: 12px; font-weight: 600;')
        caption.setWordWrap(True)
        inner.addWidget(caption)
        self.items = QPlainTextEdit('\n'.join(manual_items(settings.get('banner_manual_items', ''))))
        self.items.setPlaceholderText('เช่น\nQUOTA1\nQUOTA6')
        self.items.setFixedHeight(112)
        # its own style: the card's 'QFrame' rule would otherwise apply (QPlainTextEdit is a QFrame)
        self.items.setStyleSheet('QPlainTextEdit { background: #ffffff; border: 1px solid #c3d0e4; border-radius: 7px; '
                                 'padding: 4px 6px; font-size: 13px; } QPlainTextEdit:focus { border: 1px solid #2a8de9; } '
                                 'QPlainTextEdit:disabled { background: #f3f6fb; color: #8b97ab; }')
        self.items.setTabChangesFocus(True)
        inner.addWidget(self.items)
        check_row = QHBoxLayout()
        check_row.setSpacing(10)
        self.check_button = QPushButton('เช็คกับ Lyche')
        self.check_button.setObjectName('secondary')
        self.check_button.setCursor(Qt.PointingHandCursor)
        self.check_button.setToolTip('ค้นหาทุกข้อในรายการ Item ของ Lyche (เบื้องหลัง) ว่ามีจริงและชื่อตรง — ไม่แก้ Banner ใน Lyche')
        self.check_button.clicked.connect(self.check_items)
        check_row.addWidget(self.check_button, 0, Qt.AlignTop)
        self.check_result = QLabel('')
        self.check_result.setWordWrap(True)
        self.check_result.setTextFormat(Qt.RichText)
        self.check_result.setStyleSheet('font-size: 12px;')
        check_row.addWidget(self.check_result, 1)
        inner.addLayout(check_row)
        self.items.textChanged.connect(lambda: self.check_result.setText(''))  # an old result no longer applies
        how = QLabel('ตอนรัน: Clear all Banner เดิม → Yes → ค้นหาทีละข้อ → To Banner แล้วรันต่อตามปกติ · Export แยกชีทเท่านั้น\n'
                     'ใช้กับ Banner แบบ Matrix ไม่ได้')
        how.setStyleSheet('color: #4a5a78; font-size: 12px;')
        how.setWordWrap(True)
        inner.addWidget(how)
        layout.addWidget(card)
        self.enabled.toggled.connect(self.items.setEnabled)
        self.enabled.toggled.connect(self.check_button.setEnabled)
        self.items.setEnabled(self.enabled.isChecked())
        self.check_button.setEnabled(self.enabled.isChecked())
        self.message = QLabel('')
        self.message.setStyleSheet('color: #c62828; font-size: 12px;')
        self.message.hide()
        layout.addWidget(self.message)
        layout.addSpacing(6)
        dialog_buttons(self, layout)

    def check_items(self):
        items = manual_items(self.items.toPlainText())
        if not items:
            self.check_result.setText('<span style="color:#c62828">ใส่ข้ออย่างน้อย 1 ข้อก่อนเช็ค</span>')
            return
        if self.checker is None:
            return
        self.check_button.setEnabled(False)
        self.check_result.setText('<span style="color:#1f6fd1">กำลังเช็คกับ Lyche (เบื้องหลัง)…</span>')
        self.checker(items, self.show_check)

    def show_check(self, results, error=''):
        """Result of 'เช็คกับ Lyche': ✓ item → Lyche's name, ✗ item not found (or the error)."""
        try:
            if not self.isVisible():
                return
        except RuntimeError:  # the dialog was closed and deleted meanwhile
            return
        self.check_button.setEnabled(self.enabled.isChecked())
        if error or results is None:
            self.check_result.setText(f'<span style="color:#c62828">เช็คไม่สำเร็จ: {escape(error or "ไม่ทราบสาเหตุ")}</span>')
            return
        lines = []
        for result in results:
            if result['found']:
                label = result['label'] if len(result['label']) <= 48 else result['label'][:47] + '…'
                lines.append(f'<span style="color:#1f9a3e">✓ {escape(result["item"])}</span>'
                             f' <span style="color:#5f6f8a">→ {escape(label)}</span>')
            else:
                lines.append(f'<span style="color:#c62828">✗ {escape(result["item"])} — ไม่พบใน Lyche</span>')
        missing = sum(not r['found'] for r in results)
        lines.append('<b style="color:#1f9a3e">ตรงกับ Lyche ครบทุกข้อ</b>' if not missing
                     else f'<b style="color:#c62828">ไม่พบ {missing} ข้อ — แก้ชื่อข้อก่อนรัน</b>')
        self.check_result.setText('<br>'.join(lines))

    @property
    def settings(self):
        on = self.enabled.isChecked()
        settings = {'banner_manual': on, 'banner_manual_items': ', '.join(manual_items(self.items.toPlainText()))}
        if on:
            settings['export_mode'] = 'sheets'  # Banner Manual exports with Analysis Axis only (user rule)
        return settings

    def accept(self):
        if self.enabled.isChecked() and not manual_items(self.items.toPlainText()):
            self.message.setText('เปิด Banner Manual แล้ว กรุณาใส่ข้ออย่างน้อย 1 ข้อ เช่น QUOTA1')
            self.message.show()
            self.items.setFocus()
            self.apply_all = False
            return
        super().accept()

    def accept_all(self):
        self.apply_all = True
        self.accept()


class HiddenTextDelegate(QStyledItemDelegate):
    """Column 0 keeps the Banner name as item text (queue data) but the row's dropdown shows it;
    paint background/selection only so the two never overlap."""
    def paint(self, painter, option, index):
        opt = option.__class__(option)
        self.initStyleOption(opt, index)
        opt.text = ''
        widget = opt.widget
        (widget.style() if widget else QApplication.style()).drawControl(QStyle.CE_ItemViewItem, opt, painter, widget)


class StepDelegate(QStyledItemDelegate):
    """Step / Banner Manual cells: a compact chip with the setting when on, a faint dash when off (a
    dashed '+' while hovered). One gear per row (GearDelegate) is the obvious way into the settings."""
    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        text, opt.text = opt.text, ''
        widget = opt.widget
        (widget.style() if widget else QApplication.style()).drawControl(QStyle.CE_ItemViewItem, opt, painter, widget)
        on = text not in ('', 'Off')
        hover = bool(option.state & QStyle.State_MouseOver)
        rect = option.rect.adjusted(4, 9, -4, -9)
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)
        font = QFont(option.font)
        painter.setFont(font)
        metrics = QFontMetrics(font)
        if on:
            label = metrics.elidedText(text, Qt.ElideRight, rect.width() - 14)
            chip = QRect(0, rect.y(), min(rect.width(), metrics.horizontalAdvance(label) + 18), rect.height())
            chip.moveCenter(rect.center())
            painter.setPen(QPen(QColor('#c9b6f3'), 1))
            painter.setBrush(QColor('#e0d4fb' if hover else '#ece5fb'))
            painter.drawRoundedRect(chip, 8, 8)
            painter.setPen(QColor('#5b34b8'))
            painter.drawText(chip, Qt.AlignCenter, label)
        elif hover:
            chip = QRect(0, rect.y(), min(rect.width(), 44), rect.height())
            chip.moveCenter(rect.center())
            painter.setPen(QPen(QColor('#8fb3e8'), 1, Qt.DashLine))
            painter.setBrush(QColor('#eef4fd'))
            painter.drawRoundedRect(chip, 8, 8)
            painter.setPen(QColor('#1a5fd0'))
            painter.drawText(chip, Qt.AlignCenter, '＋')
        else:
            painter.setPen(QColor('#b4bfd1'))
            painter.drawText(rect, Qt.AlignCenter, '–')
        painter.restore()


class GearDelegate(QStyledItemDelegate):
    """The row's single gear: opens its Step settings."""
    gear = None
    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        opt.text = ''
        widget = opt.widget
        (widget.style() if widget else QApplication.style()).drawControl(QStyle.CE_ItemViewItem, opt, painter, widget)
        if GearDelegate.gear is None:
            GearDelegate.gear = QIcon(f'{ASSETS}/gear.svg')
        hover = bool(option.state & QStyle.State_MouseOver)
        box = QRect(0, 0, 30, 30)
        box.moveCenter(option.rect.center())
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(QPen(QColor('#c9b6f3' if hover else '#d9cdf6'), 1))
        painter.setBrush(QColor('#e0d4fb' if hover else '#f3eefc'))
        painter.drawRoundedRect(box, 8, 8)
        GearDelegate.gear.paint(painter, box.adjusted(7, 7, -7, -7))
        painter.restore()


class ResultDialog(QDialog):
    """Run summary shown when a queue run ends: success, error, or stop."""
    THEMES = {
        'ok': ('✓', '#1f9a3e', '#e2f5e7'),
        'warn': ('!', '#b26a00', '#fff0d9'),
        'error': ('✕', '#c62828', '#fde4e4'),
        'info': ('?', '#1f6fd1', '#e1eefc'),
    }
    PILLS = {'OK': ('#1f9a3e', '#e2f5e7'), 'ผิดพลาด': ('#c62828', '#fde4e4'), 'หยุด': ('#b26a00', '#fff0d9')}

    def __init__(self, parent, kind, title, message, results, folder):
        super().__init__(parent)
        self.folder = folder
        icon, accent, soft = self.THEMES[kind]
        self.setWindowTitle(APP_NAME)
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedWidth(500)
        shell = QVBoxLayout(self)
        shell.setContentsMargins(20, 20, 20, 20)
        sheet = QFrame()
        sheet.setObjectName('sheet')
        shadow = QGraphicsDropShadowEffect(sheet)
        shadow.setBlurRadius(36)
        shadow.setOffset(0, 8)
        shadow.setColor(QColor(27, 63, 208, 60))
        sheet.setGraphicsEffect(shadow)
        shell.addWidget(sheet)
        layout = QVBoxLayout(sheet)
        layout.setContentsMargins(28, 28, 28, 22)
        layout.setSpacing(10)

        badge = QLabel(icon)
        badge.setFixedSize(60, 60)
        badge.setAlignment(Qt.AlignCenter)
        badge.setStyleSheet(f'background: {soft}; color: {accent}; border-radius: 30px; font-size: 28px; font-weight: 700;')
        layout.addWidget(badge, 0, Qt.AlignHCenter)
        layout.addSpacing(4)
        heading = QLabel(title)
        heading.setAlignment(Qt.AlignCenter)
        heading.setStyleSheet('font-size: 19px; font-weight: 700;')
        layout.addWidget(heading)
        detail = QLabel(message)
        detail.setAlignment(Qt.AlignCenter)
        detail.setWordWrap(True)
        detail.setStyleSheet('color: #6e6e73;')
        layout.addWidget(detail)
        layout.addSpacing(8)

        if results is None:  # plain notice (e.g. Lyche not found): no run statistics
            results = []
            show_stats = False
        else:
            show_stats = True
        counts = {status: sum(r[1] == status for r in results) for status in self.PILLS}
        stats = QHBoxLayout()
        stats.setSpacing(8)
        for status, (color, bg) in self.PILLS.items():
            tile = QLabel(f'<div style="font-size:20px;font-weight:700;color:{color if counts[status] else "#c7c7cc"}">{counts[status]}</div>'
                          f'<div style="color:#86868b;font-size:12px">{status}</div>')
            tile.setAlignment(Qt.AlignCenter)
            tile.setStyleSheet(f'background: {bg if counts[status] else "#ffffff"}; border: 1px solid {bg if counts[status] else "#e5e5ea"}; border-radius: 12px; padding: 8px;')
            stats.addWidget(tile, 1)
        if show_stats:
            layout.addLayout(stats)
        else:
            for index in range(stats.count()):
                stats.itemAt(index).widget().hide()

        if results:
            card = QFrame()
            card.setStyleSheet('QFrame { background: #ffffff; border: 1px solid #e5e5ea; border-radius: 12px; } QLabel { border: none; background: transparent; }')
            rows = QVBoxLayout(card)
            rows.setContentsMargins(14, 10, 14, 10)
            rows.setSpacing(8)
            shown = results[:8]
            for name, status, note in shown:
                color, _ = self.PILLS.get(status, ('#86868b', ''))
                line = QHBoxLayout()
                line.setSpacing(10)
                dot = QLabel('●')
                dot.setStyleSheet(f'color: {color}; font-size: 10px;')
                line.addWidget(dot, 0, Qt.AlignTop)
                label = QLabel(f'<span style="font-weight:600">{escape(name or "-")}</span>'
                               + (f'<br><span style="color:#86868b;font-size:12px">{escape(note)}</span>' if note and status != 'OK' else ''))
                label.setWordWrap(True)
                line.addWidget(label, 1)
                state = QLabel(status)
                state.setStyleSheet(f'color: {color}; font-size: 12px; font-weight: 600;')
                line.addWidget(state, 0, Qt.AlignTop)
                rows.addLayout(line)
            if len(results) > len(shown):
                more = QLabel(f'และอีก {len(results) - len(shown)} รายการ — ดูในตาราง')
                more.setStyleSheet('color: #86868b; font-size: 12px;')
                rows.addWidget(more)
            layout.addWidget(card)

        layout.addSpacing(8)
        buttons = QHBoxLayout()
        buttons.setSpacing(8)
        if folder and Path(folder).is_dir():
            open_folder = QPushButton('เปิดโฟลเดอร์')
            open_folder.setObjectName('secondary')
            open_folder.setCursor(Qt.PointingHandCursor)
            open_folder.clicked.connect(self.open_folder)
            buttons.addWidget(open_folder, 1)
        close = QPushButton('ตกลง')
        close.setObjectName('primary')
        close.setCursor(Qt.PointingHandCursor)
        close.setDefault(True)
        close.clicked.connect(self.accept)
        buttons.addWidget(close, 1)
        layout.addLayout(buttons)

    def open_folder(self):
        import os
        os.startfile(self.folder)


class App(MacWindowMixin, QMainWindow):
    def __init__(self):
        super().__init__()
        self.restore_sounds()  # in case the last run ended while System Sounds were muted
        self.setWindowTitle(f'{APP_NAME} — Table Runner')
        self.setWindowIcon(QIcon(str(LOGO)))
        self.resize(1210, 820)
        self.setMinimumSize(980, 660)
        self.banner_names = []
        self.busy = False
        self.process = None
        self.buffer = ''
        self.control_dir = None
        self.action = None
        self.locked_widgets = []
        self.overlays = []
        self.freeze_timer = QTimer(self)
        self.freeze_timer.setSingleShot(True)
        self.freeze_timer.timeout.connect(self.unfreeze_screen)
        self.focus_guard = QTimer(self)
        self.focus_guard.setInterval(100)
        self.focus_guard.timeout.connect(self.keep_focus)
        self.build_ui()
        self.restore_settings()
        if not self.table.rowCount():
            self.add_job(Job())
        QTimer.singleShot(250, self.scan_windows)

    def center_on_screen(self):
        """Open centred on the monitor under the mouse, shrinking to fit small screens."""
        screen = QApplication.screenAt(QCursor.pos()) or QApplication.primaryScreen()
        area = screen.availableGeometry()
        width = min(self.width(), area.width() - 40)
        height = min(self.height(), area.height() - 40)
        self.setGeometry(area.x() + (area.width() - width) // 2, area.y() + (area.height() - height) // 2, width, height)

    def button(self, label, callback, row, primary=False, lock=True, kind=None, tip=None):
        button = QPushButton(label)
        button.clicked.connect(callback)
        button.setCursor(Qt.PointingHandCursor)
        button.setObjectName('primary' if primary else kind or 'secondary')
        if tip:
            button.setToolTip(tip)
        row.addWidget(button)
        if lock:
            self.locked_widgets.append(button)
        return button

    @staticmethod
    def card(parent_layout, stretch=0):
        frame = QFrame()
        frame.setObjectName('card')
        parent_layout.addWidget(frame, stretch)
        inner = QVBoxLayout(frame)
        inner.setContentsMargins(18, 16, 18, 16)
        inner.setSpacing(12)
        return inner

    @staticmethod
    def caption(text):
        label = QLabel(text)
        label.setObjectName('caption')
        return label

    def build_ui(self):
        central = QWidget()
        central.setObjectName('root')
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        outer.addWidget(self.setup_chrome(f'{APP_NAME} — Table Runner'))
        layout = QVBoxLayout()
        layout.setContentsMargins(28, 4, 28, 22)
        layout.setSpacing(14)
        outer.addLayout(layout, 1)

        # Header: title + live status pill (drag the window from here too)
        header = QHBoxLayout()
        header.setSpacing(14)
        logo = QLabel()
        logo.setFixedSize(54, 54)
        logo.setScaledContents(True)
        logo.setPixmap(QIcon(str(onefile.ASSETS / 'logo.png')).pixmap(108, 108))
        header.addWidget(logo, 0, Qt.AlignVCenter)
        titles = QVBoxLayout()
        titles.setSpacing(2)
        title = QLabel(APP_NAME)
        title.setObjectName('title')
        titles.addWidget(title)
        subtitle = QLabel('Import History  ·  Filter queue  ·  Export with Analysis Axis')
        subtitle.setObjectName('subtitle')
        titles.addWidget(subtitle)
        header.addLayout(titles, 1)
        self.add_drag_widgets(central, title, subtitle, logo)
        self.state = QLabel('พร้อม')
        self.state.setObjectName('status')
        header.addWidget(self.state, 0, Qt.AlignVCenter)
        layout.addLayout(header)

        # Setup card
        setup = self.card(layout)
        row = QHBoxLayout()
        row.setSpacing(10)
        label = self.caption('หน้าต่าง Lyche')
        label.setFixedWidth(104)
        row.addWidget(label)
        self.window_combo = QComboBox()
        self.window_combo.setPlaceholderText('เลือกหน้าต่าง Lyche')
        self.window_combo.setMinimumWidth(420)
        row.addWidget(self.window_combo, 1)
        self.button('ค้นหาหน้าต่าง', self.scan_windows, row)
        setup.addLayout(row)
        row = QHBoxLayout()
        row.setSpacing(10)
        label = self.caption('Import History')
        label.setFixedWidth(104)
        row.addWidget(label)
        self.source = QComboBox()
        self.source.addItems(['Personal', 'Shared'])
        self.source.setFixedWidth(118)
        row.addWidget(self.source)
        self.banner = QComboBox()
        self.banner.setEditable(True)
        self.banner.setMinimumWidth(210)
        self.banner.setPlaceholderText('เลือก Banner')
        row.addWidget(self.banner, 1)
        self.button('Get Banner', self.get_banners, row, kind='blue')
        self.button('โหลด Banner', self.load_banner, row, kind='purple')
        setup.addLayout(row)

        # Queue card: toolbar + table
        queue = self.card(layout, 1)
        row = QHBoxLayout()
        row.setSpacing(6)
        section = QLabel('รายการรัน')
        section.setObjectName('section')
        row.addWidget(section)
        row.addSpacing(8)
        row.addWidget(self.caption('Filter:  -  ไม่กรอง   ·   QUOTA6 = 1  Include   ·   QUOTA6 ^= 1  Exclude'))
        row.addStretch()
        settings_button = self.button('ตั้งค่า Step', self.edit_selected_post, row, kind='purple',
                    tip='ตั้งค่าของแถวที่เลือก: Step 1 Delete Total + NA · Step 2 Del Sig · Step 3 ตัด N / % · Step 4 Export\n'
                        '(หรือคลิกที่ช่อง Step ในตาราง)')
        settings_button.setIcon(QIcon(f'{ASSETS}/gear.svg'))
        settings_button.setIconSize(QSize(15, 15))
        row.addSpacing(6)
        for icon, callback, tip in [
                ('plus', self.add_row, 'เพิ่มแถว'), ('duplicate', self.duplicate, 'คัดลอกแถว'), ('minus', self.remove, 'ลบแถว'),
                ('arrow-up', lambda: self.move(-1), 'เลื่อนขึ้น'), ('arrow-down', lambda: self.move(1), 'เลื่อนลง')]:
            button = self.button('', callback, row, kind='icon', tip=tip)
            button.setIcon(QIcon(f'{ASSETS}/{icon}.svg'))
            button.setIconSize(QSize(16, 16))
        row.addSpacing(6)
        self.button('วางจาก Excel', self.paste, row, kind='plain', tip='วาง 2 คอลัมน์: ชื่อผลลัพธ์ / Filter')
        self.button('นำเข้า Excel', self.import_excel, row, kind='plain')
        self.button('ส่งออก Excel', self.export_excel, row, kind='plain', tip='บันทึกคิวเป็นไฟล์ Excel (นำเข้ากลับได้)')
        queue.addLayout(row)
        self.table = QTableWidget(0, GEAR_COLUMN + 1)
        self.table.setItemDelegateForColumn(0, HiddenTextDelegate(self.table))
        self.step_delegate = StepDelegate(self.table)
        for column in (*STEP_COLUMNS, MANUAL_COLUMN):
            self.table.setItemDelegateForColumn(column, self.step_delegate)
        self.table.setItemDelegateForColumn(GEAR_COLUMN, GearDelegate(self.table))
        self.table.setHorizontalHeaderLabels(['Banner', 'ชื่อไฟล์ผลลัพธ์', 'Filter', 'Base', 'สถานะ', 'รายละเอียด',
                                              *[f'Step {n}' for n in range(1, len(STEP_COLUMNS) + 1)], 'Banner Manual', ''])
        # Columns 6+ (the Steps, Export last) are appended so every other column index stays the
        # same; they are only *shown* between Base and สถานะ. Column 5 (detail) keeps its data for
        # the run summary / queue export but is hidden: progress is shown in the log below.
        header = self.table.horizontalHeader()
        for position, column in enumerate(STEP_COLUMNS, 4):
            header.moveSection(header.visualIndex(column), position)
        header.moveSection(header.visualIndex(MANUAL_COLUMN), 1)  # right after Banner
        self.table.horizontalHeaderItem(MANUAL_COLUMN).setToolTip(
            'Banner Manual · ใส่ข้อเป็น Banner แทน Banner จาก History\nคลิกที่ช่องในแถวเพื่อตั้งค่า')
        self.table.horizontalHeaderItem(GEAR_COLUMN).setToolTip('ตั้งค่า Step 1–4 ของแถว')
        self.table.setColumnHidden(5, True)
        for number, (column, name) in enumerate(zip(STEP_COLUMNS, STEP_NAMES), 1):
            self.table.horizontalHeaderItem(column).setToolTip(f'Step {number} · {name}\nคลิกที่ช่องในแถวเพื่อตั้งค่า')
        self.table.cellDoubleClicked.connect(self.cell_double_clicked)
        self.table.cellClicked.connect(self.cell_clicked)  # Step chips open the settings with one click
        self.table.setMouseTracking(True)
        self.table.viewport().installEventFilter(self)  # hand cursor over the Step chips
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(False)
        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(44)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.table.horizontalHeader().setHighlightSections(False)
        self.table.setWordWrap(False)  # one line per cell; long Sig groups end with … (full text in the tooltip)
        # Compact widths so every column fits the window (user request); the file name takes what is left.
        for index, width in enumerate([150, 160, 104, 52, 76, 300, 88, 112, 80, 88, 136, 46]):
            self.table.setColumnWidth(index, width)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(GEAR_COLUMN, QHeaderView.Fixed)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.itemChanged.connect(self.edited)
        self.log = QPlainTextEdit()
        self.log.setObjectName('log')
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(600)
        self.log.setPlaceholderText('ความคืบหน้าจะแสดงที่นี่')
        self.progress_lines = {}  # row id -> (cursor on its progress line, line text, step)
        split = QSplitter(Qt.Vertical)
        split.setHandleWidth(10)
        split.addWidget(self.table)
        split.addWidget(self.log)
        split.setSizes([300, 190])  # the log is where every step shows its progress
        queue.addWidget(split, 1)

        # Output card
        output = self.card(layout)
        row = QHBoxLayout()
        row.setSpacing(10)
        label = self.caption('โฟลเดอร์ผลลัพธ์')
        label.setFixedWidth(104)
        row.addWidget(label)
        self.folder = QLineEdit()
        self.folder.setPlaceholderText('เลือกโฟลเดอร์บันทึกไฟล์ Excel')
        row.addWidget(self.folder, 1)
        self.button('เลือกโฟลเดอร์', self.choose_folder, row, kind='gold')
        row.addSpacing(12)
        row.addWidget(self.caption('รอสูงสุด/ขั้นตอน'))
        self.timeout = QSpinBox()
        self.timeout.setRange(1, 180)
        self.timeout.setValue(15)
        self.timeout.setSuffix('  นาที')
        self.timeout.setFixedWidth(104)
        row.addWidget(self.timeout)
        output.addLayout(row)

        # Action bar
        row = QHBoxLayout()
        row.setSpacing(8)
        for label, callback in [('บันทึกคิว', self.save_as), ('เปิดคิว', self.open_queue), ('ตรวจคิว', self.validate), ('รีเซ็ตสถานะแถวที่เลือก', self.reset_status)]:
            self.button(label, callback, row, kind='plain')
        row.addStretch()
        # Run mode toggle: BG (default) = Lyche kept out of sight; Preview = visible, mouse/keyboard.
        mode = QLabel('โหมดรัน')
        mode.setObjectName('modeLabel')
        row.addWidget(mode)
        segment = QHBoxLayout()
        segment.setSpacing(0)
        self.bg_button = QPushButton('BG')
        self.bg_button.setObjectName('segLeft')
        self.bg_button.setToolTip('รันเบื้องหลัง — ไม่เห็นหน้าต่าง Lyche (ค่าเริ่มต้น)')
        self.preview_button = QPushButton('Preview')
        self.preview_button.setObjectName('segRight')
        self.preview_button.setToolTip('รันแบบเห็นหน้าจอ Lyche — ใช้เมาส์/คีย์บอร์ดจริง ห้ามใช้เครื่องระหว่างรัน')
        self.mode_group = QButtonGroup(self)
        for button in (self.bg_button, self.preview_button):
            button.setCheckable(True)
            button.setCursor(Qt.PointingHandCursor)
            self.mode_group.addButton(button)
            segment.addWidget(button)
            self.locked_widgets.append(button)
        self.bg_button.setChecked(True)
        self.mode_group.buttonClicked.connect(lambda *_: self.autosave())
        row.addLayout(segment)
        row.addSpacing(12)
        self.pause_button = self.button('พัก', self.pause, row, lock=False, kind='amber')
        self.stop_button = self.button('หยุด  F8', self.stop, row, lock=False, kind='red')
        self.run_button = self.button('▶  รันแถวที่ยังไม่ OK', self.start_run, row, True)
        self.pause_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        layout.addLayout(row)

        self.locked_widgets += [self.table, self.window_combo, self.source, self.banner, self.folder, self.timeout]
        self.source.currentTextChanged.connect(self.clear_banners)
        self.window_combo.currentIndexChanged.connect(self.clear_banners)
        self.paste_shortcut = QShortcut(QKeySequence('Ctrl+V'), self.table)
        self.paste_shortcut.setContext(Qt.WidgetShortcut)
        self.paste_shortcut.activated.connect(self.paste)
        QShortcut(QKeySequence('F8'), self, activated=self.stop)
        self.setStyleSheet(STYLE)

    def clear_banners(self, *_):
        self.update_banner_choices([])
        if not self.busy and self.window_combo.currentData():  # Personal = file read, Shared = Lyche list
            QTimer.singleShot(0, self.get_banners)

    def update_banner_choices(self, names):
        self.banner_names = list(names)
        current = self.banner.currentText()
        self.banner.blockSignals(True)
        self.banner.clear()
        self.banner.addItems(self.banner_names)
        self.banner.setCurrentIndex(0 if names else -1)
        if current in names:
            self.banner.setCurrentText(current)
        self.banner.blockSignals(False)
        for row in range(self.table.rowCount()):
            combo = self.table.cellWidget(row, 0)
            if combo is None:
                continue
            value = self.table.item(row, 0).text()
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(self.banner_names)
            combo.setCurrentText(value)
            combo.blockSignals(False)

    def banner_changed(self, job_id, value):
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item.data(Qt.UserRole) == job_id:
                item.setText(value.strip())
                return

    def jobs(self):
        result = []
        for row in range(self.table.rowCount()):
            values = [self.table.item(row, col).text().strip() for col in range(6)]
            post = self.table.item(row, 0).data(Qt.UserRole + 1) or {}
            defaults = post_settings(Job())
            result.append(Job(*values, id=self.table.item(row, 0).data(Qt.UserRole),
                              **{key: post.get(key, defaults[key]) for key in POST_KEYS}))
        return result

    def add_job(self, job):
        self.table.blockSignals(True)
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col, value in enumerate((job.history, job.output, job.filter, job.base, job.status, job.detail)):
            item = QTableWidgetItem(value)
            if col >= 4:
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            if col == 0:
                item.setData(Qt.UserRole, job.id)
                item.setData(Qt.UserRole + 1, post_settings(job))
            self.table.setItem(row, col, item)
        for column in (*STEP_COLUMNS, MANUAL_COLUMN, GEAR_COLUMN):
            step_item = QTableWidgetItem()
            step_item.setFlags(step_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, column, step_item)
        self.show_post(row)
        self.table.blockSignals(False)
        combo = QComboBox(self.table)
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.addItems(self.banner_names)
        combo.setCurrentText(job.history)
        combo.lineEdit().setPlaceholderText('เลือก Banner ▾')
        combo.lineEdit().installEventFilter(self)  # double-click the Banner → post-processing settings
        combo.lineEdit().setProperty('job_id', job.id)
        combo.currentTextChanged.connect(lambda text, job_id=job.id: self.banner_changed(job_id, text))
        self.table.setCellWidget(row, 0, combo)
        self.color_status(row)

    def show_post(self, row):
        post = self.table.item(row, 0).data(Qt.UserRole + 1) or {}
        tip = post_summary(post) + '\nคลิกเพื่อตั้งค่าหลังรัน (Step 1 Delete Total + NA, Step 2 Del Sig, Step 3 ตัด N / %, Step 4 Export)'
        for column, text in zip(STEP_COLUMNS, step_texts(post)):
            item = self.table.item(row, column)
            item.setText(text or 'Off')  # StepDelegate draws the chip
            item.setTextAlignment(Qt.AlignCenter)
            item.setToolTip(tip)
        manual = manual_text(post)
        item = self.table.item(row, MANUAL_COLUMN)
        item.setText(manual or 'Off')
        item.setTextAlignment(Qt.AlignCenter)
        self.table.item(row, GEAR_COLUMN).setToolTip('ตั้งค่า Step 1–4 ของแถวนี้\n' + post_summary(post))
        item.setToolTip(('Banner Manual: ' + manual if manual else 'Banner Manual ปิดอยู่ (ใช้ Banner จาก History)')
                        + '\nคลิกเพื่อตั้งค่า Banner Manual')

    def eventFilter(self, obj, event):
        if hasattr(self, 'table') and obj is self.table.viewport():
            if event.type() == QEvent.MouseMove:
                index = self.table.indexAt(event.position().toPoint())
                on_chip = index.isValid() and index.column() in (*STEP_COLUMNS, MANUAL_COLUMN, GEAR_COLUMN) and not self.busy
                obj.setCursor(Qt.PointingHandCursor if on_chip else Qt.ArrowCursor)
            elif event.type() == QEvent.Leave:
                obj.unsetCursor()
            return super().eventFilter(obj, event)
        if event.type() == QEvent.MouseButtonDblClick and obj.property('job_id') and not self.busy:
            for row in range(self.table.rowCount()):
                if self.table.item(row, 0).data(Qt.UserRole) == obj.property('job_id'):
                    QTimer.singleShot(0, lambda row=row: self.edit_post(row))
                    return True
        return super().eventFilter(obj, event)

    def cell_double_clicked(self, row, column):
        if column == MANUAL_COLUMN and not self.busy:
            self.edit_post(row, manual=True)
        elif column in (0, 4, 5, *STEP_COLUMNS) and not self.busy:  # columns 1-3 keep double-click-to-edit
            self.edit_post(row)

    def cell_clicked(self, row, column):
        if column in (*STEP_COLUMNS, MANUAL_COLUMN, GEAR_COLUMN) and not self.busy and not QApplication.keyboardModifiers():
            self.edit_post(row, manual=column == MANUAL_COLUMN)

    def edit_selected_post(self):
        """Toolbar '⚙ ตั้งค่า Step': the selected rows (all get the same settings), else the first row."""
        if not self.table.rowCount():
            self.error('ยังไม่มีแถวในคิว กด + เพื่อเพิ่มแถวก่อน')
            return
        rows = self.selected_rows() or [0]
        self.edit_post(rows[0], rows)

    def edit_post(self, row, rows=None, manual=False):
        """Post-processing settings (or Banner Manual) of a row, or of `rows` (a multi-row selection)."""
        if getattr(self, '_editing_post', False):  # a click and a double-click on the same chip
            return
        self._editing_post = True
        try:
            self._edit_post(row, rows, manual)
        finally:
            self._editing_post = False

    def _edit_post(self, row, rows=None, manual=False):
        item = self.table.item(row, 0)
        current = item.data(Qt.UserRole + 1) or {}
        name = self.table.item(row, 1).text() or item.text() or f'แถว {row + 1}'
        if rows and len(rows) > 1:
            name = f'{len(rows)} แถวที่เลือก'
        dialog = BannerManualDialog(self, name, current, self.check_manual_items) if manual \
            else PostProcessDialog(self, name, current)
        if dialog.exec() != QDialog.Accepted:
            return
        rows = range(self.table.rowCount()) if dialog.apply_all else (rows or [row])
        for target in rows:
            before = self.table.item(target, 0).data(Qt.UserRole + 1) or {}
            after = {**before, **dialog.settings}  # each dialog changes only its own keys
            after['export_mode'] = export_mode(after)  # e.g. 'apply to all' onto a Banner Manual row
            if before == after:
                continue
            self.table.blockSignals(True)
            self.table.item(target, 0).setData(Qt.UserRole + 1, after)
            self.table.item(target, 4).setText('รอรัน')  # the output would change: run the row again
            self.table.item(target, 5).setText('')
            self.table.blockSignals(False)
            self.show_post(target)
            self.color_status(target)
        self.autosave()

    def add_row(self):
        self.add_job(Job(history=self.banner.currentText().strip()))
        self.table.selectRow(self.table.rowCount() - 1)
        self.autosave()

    def selected_rows(self):
        return sorted({index.row() for index in self.table.selectionModel().selectedRows()})

    def duplicate(self):
        jobs = self.jobs()
        for row in self.selected_rows():
            job = jobs[row]
            self.add_job(Job(history=job.history, output=job.output + '_copy', filter=job.filter, base=job.base,
                             **post_settings(job)))
        self.autosave()

    def remove(self):
        for row in reversed(self.selected_rows()):
            self.table.removeRow(row)
        self.autosave()

    def move(self, direction):
        selected = self.selected_rows()
        if len(selected) != 1:
            return
        source = selected[0]
        target = source + direction
        if not 0 <= target < self.table.rowCount():
            return
        jobs = self.jobs()
        jobs[source], jobs[target] = jobs[target], jobs[source]
        self.replace_jobs(jobs)
        self.table.selectRow(target)
        self.autosave()

    def replace_jobs(self, jobs):
        self.table.setRowCount(0)
        for job in jobs:
            self.add_job(job)

    def paste(self):
        if self.busy:
            return
        rows = csv.reader(io.StringIO(QApplication.clipboard().text()), delimiter='\t')
        for job in read_jobs(rows, self.banner.currentText().strip()):
            self.add_job(job)
        self.autosave()

    def import_excel(self):
        filename, _ = QFileDialog.getOpenFileName(self, 'เลือกไฟล์คิว (ไฟล์ที่ส่งออกจากโปรแกรม หรือชีทแรก: Banner, ชื่อไฟล์, Filter, Base)', '', 'Excel (*.xlsx);;CSV (*.csv)')
        if not filename:
            return
        try:
            if filename.lower().endswith('.csv'):
                with open(filename, encoding='utf-8-sig', newline='') as stream:
                    jobs = read_jobs(csv.reader(stream), self.banner.currentText())
            else:
                from openpyxl import load_workbook
                workbook = load_workbook(filename, read_only=True, data_only=True)
                try:
                    jobs = read_jobs(workbook.worksheets[0].iter_rows(values_only=True), self.banner.currentText())
                finally:
                    workbook.close()
            for job in jobs:
                self.add_job(job)
            self.autosave()
            self.write_log(f'นำเข้าคิว {len(jobs)} แถว (พร้อมค่าตั้ง Step 1-3 ถ้ามีในไฟล์) ← {Path(filename).name}')
        except Exception as exc:
            self.error(str(exc))

    def export_excel(self):
        """Save the queue as .xlsx with every Step 1-3 setting (core.QUEUE_COLUMNS), so it can be
        edited in Excel (dropdowns) and imported back complete; status/detail are for reference."""
        jobs = [job for job in self.jobs() if job.history or job.output or job.filter not in ('', '-')]
        if not jobs:
            self.error('ยังไม่มีรายการในคิว')
            return
        default = QUEUE_DIR / f'Auto Lychee queue {datetime.now():%Y%m%d-%H%M}.xlsx'
        filename, _ = QFileDialog.getSaveFileName(self, 'ส่งออกคิวเป็น Excel', str(default), 'Excel (*.xlsx)')
        if not filename:
            return
        if not filename.lower().endswith('.xlsx'):
            filename += '.xlsx'
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Font, PatternFill
            from openpyxl.utils import get_column_letter
            from openpyxl.worksheet.datavalidation import DataValidation
            from core import QUEUE_COLUMNS, setting_cell
            book = Workbook()
            sheet = book.active
            sheet.title = 'Queue'
            sheet.append([header for header, _ in QUEUE_COLUMNS])
            for job in jobs:
                row = []
                for header, field in QUEUE_COLUMNS:
                    if field is None:
                        row.append(job.status if header == 'สถานะ' else job.detail)
                    elif field == 'base':
                        row.append(int(job.base) if job.base.isdigit() else job.base)
                    elif field in ('history', 'output', 'filter'):
                        row.append(getattr(job, field))
                    else:
                        row.append(setting_cell(job, field))
                sheet.append(row)
            fills = {'Step 1': '6C3FC8', 'Step 2': '7B52D1', 'Step 3': '8A65DA'}
            for cell in sheet[1]:
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor=fills.get(str(cell.value)[:6], '2A8DE9'))
                cell.alignment = Alignment(vertical='center', wrap_text=True)
            sheet.row_dimensions[1].height = 32
            columns = {field or header: get_column_letter(index) for index, (header, field) in enumerate(QUEUE_COLUMNS, 1)}
            status = columns['สถานะ']
            colors = {'OK': ('1F9A3E', 'E2F5E7'), 'ผิดพลาด': ('C62828', 'FDE4E4'), 'หยุด': ('B26A00', 'FFF0D9')}
            for index in range(2, len(jobs) + 2):
                cell = sheet[f'{status}{index}']
                if cell.value in colors:
                    fg, bg = colors[cell.value]
                    cell.font = Font(bold=True, color=fg)
                    cell.fill = PatternFill('solid', fgColor=bg)
            # dropdowns so the settings can be edited in Excel and still import correctly
            last = max(len(jobs) + 1, 200)
            for fields, choices in ((('banner_manual', 'total_na', 'total_na_empty_rows', 'del_sig', 'cut_percent'), 'เปิด,ปิด'),
                                    (('del_sig_mode',), 'Crosstab ธรรมดา,Matrix'),
                                    (('del_sig_beside',), 'Sig ปกติ,Sig ข้าง'),
                                    (('cut_percent_mode',), 'N + %,N Only,% Only'),
                                    (('export_mode',), 'แยกชีท,One Sheet')):
                rule = DataValidation(type='list', formula1=f'"{choices}"', allow_blank=True)
                sheet.add_data_validation(rule)
                for field in fields:
                    rule.add(f'{columns[field]}2:{columns[field]}{last}')
            widths = {'history': 16, 'banner_manual_items': 22, 'output': 26, 'filter': 18, 'base': 8, 'del_sig_groups': 26,
                      'del_sig_mode': 17, 'del_sig_beside': 13, 'cut_percent_mode': 13, 'export_mode': 13,
                      'สถานะ': 11, 'รายละเอียด': 60}
            for key, letter in columns.items():
                sheet.column_dimensions[letter].width = widths.get(key, 13)
            sheet.freeze_panes = 'E2'  # Banner, Banner Manual (on/off, items), output name
            sheet.auto_filter.ref = sheet.dimensions
            book.save(filename)
        except PermissionError:
            self.error('บันทึกไม่ได้ — ไฟล์นี้อาจเปิดอยู่ใน Excel กรุณาปิดไฟล์แล้วลองใหม่')
            return
        except Exception as exc:
            self.error(f'ส่งออก Excel ไม่สำเร็จ: {exc}')
            return
        self.write_log(f'ส่งออกคิว {len(jobs)} แถว → {filename}')
        ResultDialog(self, 'ok', 'ส่งออกคิวสำเร็จ',
                     f'บันทึก {len(jobs)} แถว พร้อมค่าตั้ง Step 1-3 เป็นไฟล์ Excel แล้ว\n{Path(filename).name}\nกด “นำเข้า Excel” เพื่อโหลดกลับมาได้ครบทุกค่า',
                     None, str(Path(filename).parent)).exec()

    def edited(self, item):
        if item.column() < 4:
            self.table.blockSignals(True)
            self.table.item(item.row(), 4).setText('รอรัน')
            self.table.item(item.row(), 5).setText('')
            self.table.blockSignals(False)
            self.color_status(item.row())
        self.autosave()

    def color_status(self, row):
        item = self.table.item(row, 4)
        colors = {'OK': ('#1f9a3e', '#e2f5e7'), 'ผิดพลาด': ('#c62828', '#fde4e4'),
                  'กำลังรัน': ('#1f6fd1', '#e1eefc'), 'หยุด': ('#b26a00', '#fff0d9')}
        fg, bg = colors.get(item.text(), ('#5f6f8a', None))
        item.setForeground(QColor(fg))
        item.setBackground(QColor(bg) if bg else QColor(0, 0, 0, 0))
        font = item.font()
        font.setWeight(QFont.DemiBold if item.text() in colors else QFont.Normal)
        item.setFont(font)
        item.setTextAlignment(Qt.AlignCenter)

    def reset_status(self):
        for row in self.selected_rows():
            self.table.item(row, 4).setText('รอรัน')
            self.table.item(row, 5).setText('')
            self.color_status(row)
        self.autosave()

    def config(self):
        return {'version': 1, 'source': self.source.currentText(), 'folder': self.folder.text().strip(),
                'timeout': self.timeout.value() * 60, 'background': self.bg_button.isChecked(),
                'jobs': [asdict(job) for job in self.jobs()]}

    def autosave(self):
        try:
            save_json(QUEUE, self.config())
        except OSError as exc:
            self.state.setText(f'บันทึกคิวไม่ได้: {exc}')

    def restore_settings(self):
        """Start clean: keep only source/timeout preferences, never the old queue or output folder."""
        try:
            data = json.loads(QUEUE.read_text(encoding='utf-8'))
            self.source.setCurrentText(data.get('source', 'Personal'))
            self.timeout.setValue(data.get('timeout', 900) // 60)
            (self.bg_button if data.get('background', True) else self.preview_button).setChecked(True)
        except (OSError, ValueError, TypeError):
            pass

    def restore(self, path=QUEUE):
        if not path.exists():
            return
        try:
            data = json.loads(path.read_text(encoding='utf-8'))
            self.source.setCurrentText(data.get('source', 'Personal'))
            self.folder.setText(data.get('folder', ''))
            self.timeout.setValue(data.get('timeout', 900) // 60)
            jobs = [Job(**value) for value in data['jobs']]
            for job in jobs:
                if job.status == 'กำลังรัน':
                    job.status = 'หยุด'
                    job.detail = 'งานก่อนหน้าถูกขัดจังหวะ ตรวจ Lyche และไฟล์ก่อนรันต่อ'
            self.replace_jobs(jobs)
        except Exception as exc:
            self.error(f'เปิดคิวไม่ได้: {exc}')

    def save_as(self):
        filename, _ = QFileDialog.getSaveFileName(self, 'บันทึกคิว', str(QUEUE_DIR / 'my_queue.json'), 'Queue (*.json)')
        if filename:
            save_json(Path(filename), self.config())

    def open_queue(self):
        filename, _ = QFileDialog.getOpenFileName(self, 'เปิดคิว', str(QUEUE_DIR), 'Queue (*.json)')
        if filename:
            self.restore(Path(filename))
            self.autosave()

    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'โฟลเดอร์ผลลัพธ์', self.folder.text())
        if folder:
            self.folder.setText(folder)
            self.autosave()

    def validate(self, notify=True):
        try:
            if not self.folder.text().strip():
                raise ValueError('ยังไม่ได้เลือกโฟลเดอร์ผลลัพธ์')
            validate_jobs(self.jobs(), Path(self.folder.text().strip()))
            if notify:
                self.write_log('ตรวจคิวผ่าน: ชื่อไฟล์ / Filter / Base / ไฟล์ซ้ำ')
            return True
        except ValueError as exc:
            self.error(str(exc))
            return False

    def scan_windows(self):
        self.launch('windows')

    def get_banners(self):
        self.launch('banners')

    def load_banner(self):
        name = self.banner.currentText().strip()
        if not name:
            self.error('เลือกหรือพิมพ์ชื่อ Banner ก่อน')
            return
        self.launch('load', banner=name)

    def check_manual_items(self, items, callback):
        """Banner Manual 'เช็คกับ Lyche': a background worker searches each item in Lyche's item list;
        `callback(results, error)` gets [{'item', 'found', 'label'}] or an error text."""
        if self.busy:
            callback(None, 'โปรแกรมกำลังทำงานอยู่ รอให้เสร็จก่อน')
            return
        if not self.window_combo.currentData():
            callback(None, 'ค้นหาและเลือกหน้าต่าง Lyche ก่อน')
            return
        self.check_callback, self.check_results, self.check_error = callback, None, ''
        self.launch('check_items', items=items)

    def start_run(self):
        self.table.clearFocus()
        if self.validate(False):
            self.autosave()
            self.launch('run')

    def launch(self, action, **extra):
        if self.busy:
            return
        handle = self.window_combo.currentData()
        if action != 'windows' and not handle:
            self.error('ค้นหาและเลือกหน้าต่าง Lyche ก่อน')
            return
        self.control_dir = DATA / ('session-' + uuid4().hex)
        self.action = action
        self.control_dir.mkdir(parents=True, exist_ok=True)
        config = {**self.config(), 'action': action, 'handle': handle, 'window_title': self.window_combo.currentText(),
                  'control_dir': str(self.control_dir), **extra}
        request = self.control_dir / 'request.json'
        save_json(request, config)
        self.buffer = ''
        self.run_results = {}
        self.run_message = ''
        self.run_event = ''
        self.set_busy(True)
        self.state.setText({'windows': 'กำลังค้นหาหน้าต่าง…', 'banners': 'กำลังดึง Banner…', 'load': 'กำลังโหลด Banner…',
                            'check_items': 'กำลังเช็คข้อกับ Lyche…', 'run': 'กำลังรันคิว • F8 หยุดได้ทุกเมื่อ'}[action])
        self.process = QProcess(self)
        self.process.setWorkingDirectory(str(ROOT))
        self.process.readyReadStandardOutput.connect(self.read_output)
        self.process.readyReadStandardError.connect(self.read_error)
        self.process.finished.connect(self.finished)
        self.process.errorOccurred.connect(self.process_error)
        program, arguments = onefile.worker_command(str(request))  # ONEFILE: this file with --worker
        self.process.start(program, arguments)
        if action in ('run', 'load', 'check_items'):
            import ctypes
            self.user_window = ctypes.windll.user32.GetForegroundWindow()  # where the user is now
            self.focus_guard.start()

    def read_output(self):
        self.buffer += bytes(self.process.readAllStandardOutput()).decode('utf-8', errors='replace')
        while '\n' in self.buffer:
            line, self.buffer = self.buffer.split('\n', 1)
            try:
                event = json.loads(line)
            except ValueError:
                self.write_log(line)
                continue
            kind = event.get('event')
            if kind == 'windows':
                # Mid-action refresh (the worker reopened a closed window): don't trigger Get Banner.
                previous = self.window_combo.currentText()
                project = previous[previous.find('<'):previous.find('>') + 1] if '<' in previous else ''
                self.window_combo.blockSignals(True)
                self.window_combo.clear()
                for item in event['items']:
                    self.window_combo.addItem(item['title'], item['handle'])
                if self.action == 'windows' and len(event['items']) > 1:
                    self.window_combo.setCurrentIndex(-1)  # several projects: the user must choose
                elif event['items']:
                    # with a placeholder set, Qt does not pick the first item by itself
                    self.window_combo.setCurrentIndex(0)
                    for index in range(self.window_combo.count()):  # mid-action refresh: same project
                        if project and project in self.window_combo.itemText(index):
                            self.window_combo.setCurrentIndex(index)
                            break
                self.window_combo.blockSignals(False)
                self.write_log(f'พบ {len(event["items"])} หน้าต่าง Lyche')
                if event.get('opened'):  # Lyche grabbed focus when it opened
                    if self.action in ('run', 'load', 'check_items') and getattr(self, 'user_window', None):
                        QTimer.singleShot(0, lambda: self.activate_window(self.user_window))  # back to the user's work
                    else:
                        QTimer.singleShot(0, self.bring_to_front)
            elif kind == 'freeze':
                self.freeze_screen()
            elif kind == 'unfreeze':
                self.unfreeze_screen()
            elif kind == 'items_checked':
                self.check_results = event['results']
            elif kind == 'banners':
                self.update_banner_choices(event['names'])
                if event.get('raise_app'):
                    QTimer.singleShot(0, self.bring_to_front)
                if not self.table.rowCount():
                    self.add_job(Job())
                self.write_log(f'ดึงได้ {len(self.banner_names)} รายการ: ' + ', '.join(self.banner_names))
            elif kind == 'status':
                for row in range(self.table.rowCount()):
                    if self.table.item(row, 0).data(Qt.UserRole) == event['id']:
                        self.table.item(row, 4).setText(event['status'])
                        self.table.item(row, 5).setText(event.get('detail', ''))
                        if event.get('base'):  # actual respondents after the filter
                            self.table.blockSignals(True)  # editing Base would otherwise reset the status
                            self.table.item(row, 3).setText(event['base'])
                            self.table.blockSignals(False)
                        self.log_status(row, event['id'], event['status'], event.get('detail', ''), event.get('base', ''))
                        if event['status'] in ResultDialog.PILLS:
                            self.run_results[event['id']] = (self.table.item(row, 1).text(), event['status'], event.get('detail', ''))
                        self.color_status(row)
                        if not event.get('background'):  # post-processing of an earlier row: keep the selection on Lyche's row
                            self.table.selectRow(row)
                            self.table.scrollToItem(self.table.item(row, 0))
                        self.autosave()
                        break
            elif kind in ('log', 'stopped', 'error'):
                self.write_log(event['text'])
                if kind != 'log' and self.action == 'run':
                    self.run_message, self.run_event = event['text'], kind  # shown in the run summary, not a second popup
                elif kind == 'error' and self.action == 'check_items':
                    self.check_error = event['text']  # shown in the Banner Manual dialog, not a second popup
                elif kind == 'error':
                    self.error(event['text'])

    def process_error(self, error):
        if error == QProcess.FailedToStart:
            self.set_busy(False)
            self.error(f'เปิด Worker ไม่สำเร็จ: {self.process.errorString()}')

    def read_error(self):
        text = bytes(self.process.readAllStandardError()).decode('utf-8', errors='replace').strip()
        if text:
            self.write_log(text)

    @staticmethod
    def restore_sounds():
        """A worker that ended while Windows' System Sounds were muted (sounds.py) left a marker: unmute."""
        if (DATA / 'sounds-muted').exists():
            try:
                from sounds import restore_if_left_muted
                restore_if_left_muted(DATA / 'sounds-muted')
            except Exception:
                pass

    def finished(self, code, *_):
        self.focus_guard.stop()
        self.unfreeze_screen()
        self.restore_sounds()
        self.read_output()
        for row in range(self.table.rowCount()):
            if self.table.item(row, 4).text() == 'กำลังรัน':
                self.table.item(row, 4).setText('หยุด')
                self.table.item(row, 5).setText('Worker สิ้นสุดก่อนยืนยันผล ตรวจไฟล์ก่อนรันใหม่')
                self.color_status(row)
                self.log_status(row, self.table.item(row, 0).data(Qt.UserRole), 'หยุด', self.table.item(row, 5).text())
                self.run_results[self.table.item(row, 0).data(Qt.UserRole)] = (
                    self.table.item(row, 1).text(), 'หยุด', self.table.item(row, 5).text())
        self.autosave()
        self.set_busy(False)
        self.state.setText('พร้อม' if code == 0 else 'หยุดแล้ว — ดูรายละเอียดในบันทึก')
        if self.action == 'run':
            QTimer.singleShot(0, lambda: self.show_run_summary(code))
        if code == 0 and self.action == 'windows':
            QTimer.singleShot(0, self.after_scan)
        if self.action == 'check_items' and getattr(self, 'check_callback', None):
            callback, self.check_callback = self.check_callback, None
            error = self.check_error or ('' if self.check_results is not None else 'Worker สิ้นสุดก่อนได้ผล ดูรายละเอียดในบันทึก')
            QTimer.singleShot(0, lambda: callback(self.check_results, error))

    def after_scan(self):
        """After looking for Lyche: none → ask to open it; several → ask which one; one → Get Banner."""
        count = self.window_combo.count()
        if count == 0:
            self.state.setText('ไม่พบ Lyche')
            ResultDialog(self, 'warn', 'ไม่พบ Lyche',
                         'กรุณาเปิด Lyche และเปิดโปรเจกต์ก่อน แล้วกด “ค้นหาหน้าต่าง”', None, None).exec()
        elif self.window_combo.currentIndex() < 0:
            self.state.setText('เลือกหน้าต่าง Lyche')
            ResultDialog(self, 'info', f'พบ Lyche {count} หน้าต่าง',
                         'กรุณาเลือก Lyche ที่จะใช้ในช่อง “หน้าต่าง Lyche”', None, None).exec()
            self.window_combo.showPopup()
        elif self.window_combo.currentData():
            self.get_banners()

    def show_run_summary(self, code):
        results = list(self.run_results.values())
        failed = sum(r[1] == 'ผิดพลาด' for r in results)
        stopped = sum(r[1] == 'หยุด' for r in results)
        done = sum(r[1] == 'OK' for r in results)
        if code == 0 and not failed and not stopped and not self.run_event:
            kind, title = 'ok', 'รันเสร็จครบแล้ว'
            message = f'บันทึกไฟล์ Excel สำเร็จ {done} ไฟล์' if done else 'ไม่มีแถวที่ต้องรัน'
        elif (stopped or self.run_event == 'stopped') and not failed and self.run_event != 'error':
            kind, title = 'warn', 'หยุดการรันแล้ว'
            message = self.run_message or 'รันแถวที่เหลือต่อได้ด้วยปุ่ม “รันแถวที่ยังไม่ OK”'
        else:
            kind, title = 'error', 'รันไม่สำเร็จ'
            message = self.run_message or 'ดูรายละเอียดในบันทึกด้านล่าง'
        self.bring_to_front()
        if sys.platform == 'win32':
            import winsound
            winsound.MessageBeep(winsound.MB_ICONASTERISK if kind == 'ok' else winsound.MB_ICONEXCLAMATION)
        ResultDialog(self, kind, title, message, results, self.folder.text().strip()).exec()

    def freeze_screen(self):
        self.unfreeze_screen()
        self.overlays = [FreezeOverlay(screen) for screen in QApplication.screens()]
        for overlay in self.overlays:
            overlay.show()
        QApplication.processEvents()
        control = self.control_dir
        # give DWM a frame to put the overlay on screen before the worker lets Lyche open anything
        QTimer.singleShot(80, lambda: control and control.exists() and (control / 'frozen').touch())
        self.freeze_timer.start(15000)  # never leave the screen frozen

    def unfreeze_screen(self):
        self.freeze_timer.stop()
        for overlay in self.overlays:
            overlay.close()
            overlay.deleteLater()
        self.overlays = []

    def keep_focus(self):
        """While the bot drives Lyche off-screen, Lyche dialogs can still activate themselves and take
        the keyboard. Remember the window the user is working in (any non-Lyche window, this app
        included) and hand focus straight back to it. Skipped while the worker holds `hold-focus`
        (the two short steps that need Lyche active; the worker gives focus back itself)."""
        if not self.busy or sys.platform != 'win32':
            return
        import win32gui
        import win32process
        handle = self.window_combo.currentData()
        foreground = win32gui.GetForegroundWindow()
        if not handle or not foreground or not win32gui.IsWindow(handle):
            return
        lyche = win32process.GetWindowThreadProcessId(handle)[1]
        if win32process.GetWindowThreadProcessId(foreground)[1] != lyche:
            self.user_window = foreground
            return
        if win32gui.GetWindowRect(foreground)[0] > -15000:
            return  # a visible Lyche window: the user is working in Lyche
        if self.control_dir and (self.control_dir / 'hold-focus').exists():
            return
        target = getattr(self, 'user_window', None)
        if target and win32gui.IsWindow(target) and target != int(self.winId()):
            self.activate_window(target)
        else:
            self.bring_to_front(alert=False)

    @staticmethod
    def activate_window(hwnd):
        """Foreground another window from this process (borrow the foreground thread's input queue)."""
        import ctypes
        user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
        current = kernel32.GetCurrentThreadId()
        threads = {user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), None),
                   user32.GetWindowThreadProcessId(hwnd, None)} - {current, 0}
        attached = [t for t in threads if user32.AttachThreadInput(current, t, True)]
        try:
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
        finally:
            for t in attached:
                user32.AttachThreadInput(current, t, False)

    def bring_to_front(self, alert=True):
        """Raise the app over Lyche; Windows blocks plain SetForegroundWindow from a background app."""
        if self.isMinimized():
            self.showNormal()
        self.show()
        self.raise_()
        self.activateWindow()
        if sys.platform != 'win32':
            return
        import ctypes
        user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
        hwnd = int(self.winId())
        foreground = user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), None)
        current = kernel32.GetCurrentThreadId()
        attached = foreground != current and user32.AttachThreadInput(foreground, current, True)
        try:
            flags = 0x0001 | 0x0002 | 0x0040  # NOSIZE | NOMOVE | SHOWWINDOW
            user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, flags)  # topmost, then back to normal
            user32.SetWindowPos(hwnd, -2, 0, 0, 0, 0, flags)
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
        finally:
            if attached:
                user32.AttachThreadInput(foreground, current, False)
        if alert:
            QApplication.alert(self)

    def set_busy(self, busy):
        self.busy = busy
        for widget in self.locked_widgets:
            widget.setEnabled(not busy)
        self.pause_button.setEnabled(busy)
        self.stop_button.setEnabled(busy)
        self.pause_button.setText('พัก')

    def pause(self):
        if not self.busy:
            return
        flag = self.control_dir / 'pause'
        if flag.exists():
            flag.unlink()
            self.pause_button.setText('พัก')
            self.state.setText('ทำต่อแล้ว')
        else:
            flag.touch()
            self.pause_button.setText('ทำต่อ')
            self.state.setText('ขอพักก่อนคำสั่งถัดไป • Lyche อาจยังคำนวณอยู่')

    def stop(self):
        if not self.busy:
            return
        (self.control_dir / 'stop').touch()
        self.state.setText('กำลังหยุดบอต • ไม่ปิด Lyche และไม่ลบไฟล์ผลลัพธ์')
        process = self.process
        QTimer.singleShot(3000, lambda: process.kill() if process.state() != QProcess.NotRunning else None)

    def log_status(self, row, row_id, status, detail, base=''):
        """Every row status change goes to the log (the detail column is hidden). Progress of the
        same step of the same row (e.g. 'ชีต 51/204') rewrites its own line instead of adding one
        per second; it is written to the log file only when the step starts."""
        name = self.table.item(row, 1).text().strip() or self.table.item(row, 0).text().strip() or f'แถว {row + 1}'
        if status == 'กำลังรัน':
            text = f'{name} · {detail}' if detail else f'{name} · กำลังรัน'
            step = detail.split(' · ')[0]
            cursor, last, previous_step = self.progress_lines.get(row_id, (None, None, None))
            if cursor is not None and step == previous_step and cursor.block().text() == last:
                line = f'{datetime.now():%H:%M:%S}  {text}'
                cursor.movePosition(QTextCursor.StartOfBlock)
                cursor.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
                cursor.insertText(line)
                cursor.movePosition(QTextCursor.StartOfBlock)  # at the end it would follow text appended below
                self.progress_lines[row_id] = (cursor, line, step)
                return
            line = self.write_log(text)
            cursor = QTextCursor(self.log.document().lastBlock())
            self.progress_lines[row_id] = (cursor, line, step)
            return
        self.progress_lines.pop(row_id, None)
        parts = [part for part in detail.split(' | ') if part]
        if status == 'OK':
            path = parts[1] if len(parts) > 1 and parts[0] == 'Saved' else ''
            text = f'✓ {name} · OK' + (f' · Base {base}' if base else '') + (f' · {path}' if path else '')
        else:
            text = f'{"✗" if status == "ผิดพลาด" else "■"} {name} · {status}' + (f' · {detail}' if detail else '')
        self.write_log(text)

    def write_log(self, text):
        line = f'{datetime.now():%H:%M:%S}  {text}'
        self.log.appendPlainText(line)
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())
        log_dir = DATA / 'logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        with (log_dir / f'{datetime.now():%Y%m%d}.log').open('a', encoding='utf-8') as stream:
            stream.write(line + '\n')
        return line

    def error(self, text):
        self.write_log(text)
        QMessageBox.warning(self, APP_NAME, text)

    def closeEvent(self, event):
        if self.busy:
            self.stop()
            self.state.setText('กำลังหยุด กรุณาปิดอีกครั้งเมื่อบอตหยุดแล้ว')
            event.ignore()
        else:
            self.autosave()
            event.accept()


# ONEFILE: --worker / --post are handled by the loader at the top of this file.

def main():  # ONEFILE: called by the loader (was the `if __name__ == '__main__':` block)
    if sys.platform == 'win32':  # own taskbar identity, so Windows shows our logo instead of Python's
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('Intage.AutoLychee')
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setWindowIcon(QIcon(str(LOGO)))
    QLocale.setDefault(QLocale(QLocale.English, QLocale.UnitedStates))
    DATA.mkdir(parents=True, exist_ok=True)
    lock = QLockFile(str(DATA / 'app.lock'))
    if not lock.tryLock(100):
        QMessageBox.information(None, APP_NAME, 'โปรแกรมเปิดอยู่แล้ว กรุณาใช้หน้าต่างเดิม')
        sys.exit(0)
    app.setStyle('Fusion')  # native Windows 11 style ignores parts of the stylesheet
    app.setFont(QFont('Segoe UI', 10))
    window = App()
    window.center_on_screen()
    window.show()
    sys.exit(app.exec())


# ====== MODULE: onefile_assets ======
"""Icons for the app: written to onefile.ASSETS by the loader. SVG as text (editable), others base64."""
TEXT = {
    'arrow-down.svg': '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16"><path d="M8 3.5v9M4.5 9 8 12.5 11.5 9" fill="none" stroke="#2d3f5f" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    'arrow-up.svg': '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16"><path d="M8 12.5v-9M4.5 7 8 3.5 11.5 7" fill="none" stroke="#2d3f5f" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    'check.svg': '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 14 14"><path d="M3 7.2 5.8 10 11 4.2" fill="none" stroke="#ffffff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    'chevron-down.svg': '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 12 12"><path d="M3 4.5 6 7.5 9 4.5" fill="none" stroke="#5f6f8a" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    'chevron-up.svg': '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 12 12"><path d="M3 7.5 6 4.5 9 7.5" fill="none" stroke="#5f6f8a" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    'duplicate.svg': '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16"><rect x="5.5" y="5.5" width="7" height="7" rx="1.6" fill="none" stroke="#2d3f5f" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M10.5 3.5h-5a2 2 0 0 0-2 2v5" fill="none" stroke="#2d3f5f" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    'gear.svg': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="#5b34b8" d="M19.14 12.94c.04-.3.06-.61.06-.94 0-.32-.02-.64-.07-.94l2.03-1.58a.49.49 0 0 0 .12-.61l-1.92-3.32a.49.49 0 0 0-.59-.22l-2.39.96c-.5-.38-1.03-.7-1.62-.94l-.36-2.54a.48.48 0 0 0-.48-.41h-3.84c-.24 0-.43.17-.47.41l-.36 2.54c-.59.24-1.13.57-1.62.94l-2.39-.96a.49.49 0 0 0-.59.22L2.74 8.87c-.12.21-.08.47.12.61l2.03 1.58c-.05.3-.09.63-.09.94s.02.64.07.94l-2.03 1.58a.49.49 0 0 0-.12.61l1.92 3.32c.12.22.37.29.59.22l2.39-.96c.5.38 1.03.7 1.62.94l.36 2.54c.05.24.24.41.48.41h3.84c.24 0 .44-.17.47-.41l.36-2.54c.59-.24 1.13-.56 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32c.12-.22.07-.47-.12-.61l-2.01-1.58zM12 15.6c-1.98 0-3.6-1.62-3.6-3.6s1.62-3.6 3.6-3.6 3.6 1.62 3.6 3.6-1.62 3.6-3.6 3.6z"/></svg>',
    'logo.svg': '<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" viewBox="0 0 256 256">\n  <defs>\n    <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">\n      <stop offset="0" stop-color="#2a8de9"/>\n      <stop offset="1" stop-color="#6c3fc8"/>\n    </linearGradient>\n    <linearGradient id="badge" x1="0" y1="0" x2="1" y2="1">\n      <stop offset="0" stop-color="#4ade80"/>\n      <stop offset="1" stop-color="#16a34a"/>\n    </linearGradient>\n  </defs>\n\n  <!-- tile -->\n  <rect x="8" y="8" width="240" height="240" rx="56" fill="url(#bg)"/>\n  <path d="M 64 8 H 192 A 56 56 0 0 1 248 64 V 92 C 180 114 90 100 8 128 V 64 A 56 56 0 0 1 64 8 Z" fill="#ffffff" opacity="0.08"/>\n\n  <!-- queue: tables stacked behind -->\n  <rect x="68" y="34" width="148" height="124" rx="16" fill="#ffffff" opacity="0.2"/>\n  <rect x="54" y="48" width="148" height="124" rx="16" fill="#ffffff" opacity="0.38"/>\n\n  <!-- front table card -->\n  <rect x="40" y="62" width="148" height="124" rx="16" fill="#ffffff"/>\n  <!-- header band with rounded top corners -->\n  <path d="M 56 62 H 172 A 16 16 0 0 1 188 78 V 92 H 40 V 78 A 16 16 0 0 1 56 62 Z" fill="#dbe9fb"/>\n  <g stroke-linecap="round" stroke-width="7">\n    <line x1="54" y1="77" x2="74" y2="77" stroke="#1b3fd0"/>\n    <line x1="104" y1="77" x2="122" y2="77" stroke="#1b3fd0"/>\n    <line x1="152" y1="77" x2="168" y2="77" stroke="#1b3fd0"/>\n  </g>\n  <!-- grid -->\n  <g stroke="#d5e0ef" stroke-width="3">\n    <line x1="40" y1="123" x2="188" y2="123"/>\n    <line x1="40" y1="154" x2="188" y2="154"/>\n    <line x1="90" y1="92" x2="90" y2="186"/>\n    <line x1="139" y1="92" x2="139" y2="186"/>\n  </g>\n  <!-- data -->\n  <g stroke-linecap="round" stroke-width="7">\n    <line x1="54" y1="108" x2="76" y2="108" stroke="#2a8de9"/>\n    <line x1="103" y1="108" x2="124" y2="108" stroke="#9fb8dc"/>\n    <line x1="54" y1="139" x2="72" y2="139" stroke="#2a8de9"/>\n    <line x1="103" y1="139" x2="126" y2="139" stroke="#9fb8dc"/>\n    <line x1="54" y1="170" x2="74" y2="170" stroke="#6c3fc8"/>\n    <line x1="103" y1="170" x2="118" y2="170" stroke="#9fb8dc"/>\n  </g>\n  <!-- done ticks -->\n  <path d="M 152 108 l 6 6 l 13 -13" fill="none" stroke="#1f9a3e" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"/>\n  <path d="M 152 139 l 6 6 l 13 -13" fill="none" stroke="#1f9a3e" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"/>\n\n  <!-- automation badge: lightning bolt -->\n  <circle cx="190" cy="192" r="40" fill="url(#badge)" stroke="#ffffff" stroke-width="8"/>\n  <path d="M 197 165 L 175 197 H 189 L 182 220 L 206 186 H 192 Z" fill="#ffffff" stroke="#ffffff" stroke-width="3" stroke-linejoin="round"/>\n</svg>',
    'minus.svg': '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16"><path d="M3.5 8h9" fill="none" stroke="#2d3f5f" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    'plus.svg': '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16"><path d="M8 3.5v9M3.5 8h9" fill="none" stroke="#2d3f5f" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
}
BINARY = {
    'logo.ico': (
        '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'
        'aMbto54KN06J1osvqhIUcV2mAphjRynjQ+asUdPM0PxNdgMe6Xle5Qimpydiwbg5Cvzvhnbj33bcjyY7q7mGCEfJFDq/mNJf'
        'KQwe9YiLpDp5IOXgob0NdMICnI7Db+ue5zKJvZRvV27CcVa36uTIHw28Tj2Rnd63V80LcxumIc0SSgESKLr+6cFn8Plt/42/'
        '3XIHfnVsm+qkGJRx82trUPTLSkzIDNICUbsMT+JstZMEHNxdgm2RDVHOwL7ubfPcLnRxvhS3aG0m5Uuc25JpvnlHweuASyCZ'
        '4NjzZhk9PS5mZNowNz1dOTx6N0sgzRNoYElMTrXi4gkLlKL75q5uvNS3HVkrBT8IrQUd5nILaG86HbMap+LIoTL2vV5CIkmL'
        'axGHv6jmtrSDq29sNCj+iPDLpLq0v/YOgo+ekYNEHtvTT/cpcNdNvljJs06xaWVIlmFaciJm56bhmd4t+N7vH0KTk1WmL6pH'
        'MTW1+7uZV6p+nvvlEQwd84jDeBC4NPQjNOy2iR2S609LJDvVPe0Zv1zYbtlpqhC8k+Aj00Oym04xPPnoUQwMelja2o5r2t6P'
        'w6VjKgCi+mTzO8afo1j/pi2rVVCkZD6MZun+YLEPfzPjclw0+d0YHHCx6cFepNI8sFiK+V5hO/rwLGHu7maBYv+lXZusNWuY'
        'Z4PdbllJUhLynQQf1SUgtJNECY11PzigKt0042p8pG0ZDpeOwgt8xeor2hbja9vX4a38ISXnhN0CUyJwsNCLa09ZjlXtK1X7'
        '9av3ofdQGYkEkxYSzIJ9+5qX3+11Ld2kzE3kzbOuLrBNmzbx9Jy2zU4it9ArDQYWI4P6zoA365LzSOcBr71+ioroqHz39w9j'
        '9Vs/Q19pAPObZ2JL/xvKTab0lxu4KPsexieaFNt/ds6HVZtH7juA+7+zD9ksCxzeYLne0As7vMIFHR0dYtUqPSs7ssREhKee'
        'WuYvn/3adcIrv2DzREYErqDz9u8U+Iq/IZGlWH7tfhTzAT5ybRs+OWMFLjtpER7Y9ww29ryCk5MTUQxKKghqS7Viccs8dJ7S'
        'ganZiar9A2v34+c/PIBMhgkOhwdBKc+EvI4wdnTI+FAtq/fJzCU3bFth2+mfQQZcBj45SPaogP7I4M1YRMn8kI/2c3K48qOT'
        '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'
    ),
    'logo.png': (
        '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'
        'RK5CYII='
    ),
}

# ====== END OF MODULES ======
