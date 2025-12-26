# hook-savReaderWriter.py
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# ส่วนที่ 1: บอกให้ PyInstaller รู้จักโมดูลย่อย (submodules) ที่ซ่อนอยู่ทั้งหมด
# แก้ปัญหา "ModuleNotFoundError"
hiddenimports = collect_submodules('savReaderWriter')

# ส่วนที่ 2: บอกให้ PyInstaller ก็อปปี้โฟลเดอร์ spssio ที่มีไฟล์ .dll ไปด้วย
datas = collect_data_files('savReaderWriter.spssio')