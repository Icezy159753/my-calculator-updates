# hook-savReaderWriter.py (เวอร์ชันที่ง่ายที่สุด)
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules('savReaderWriter')