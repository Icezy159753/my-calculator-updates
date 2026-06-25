# =============================================================================
# === โค้ดสำหรับแก้ไข Path ของ savReaderWriter เมื่อเป็นไฟล์ .exe (PyInstaller) ===
# ส่วนนี้ต้องอยู่บนสุด ก่อนการ import อื่นๆ ที่เกี่ยวข้อง
import sys
import os

# Compatibility shim for Python 3.10+ where collections.* moved to collections.abc
try:
    import collections
    from collections.abc import Iterable, Mapping, MutableMapping, Sequence
    if not hasattr(collections, "Iterable"):
        collections.Iterable = Iterable
    if not hasattr(collections, "Mapping"):
        collections.Mapping = Mapping
    if not hasattr(collections, "MutableMapping"):
        collections.MutableMapping = MutableMapping
    if not hasattr(collections, "Sequence"):
        collections.Sequence = Sequence
except Exception:
    pass

# ตรวจสอบว่าโปรแกรมกำลังรันในรูปแบบ "frozen" (ไฟล์ .exe) หรือไม่
# รองรับทั้งการรันจาก Main Process และ Subprocess
_spss_found = False
base_dirs = []

# 1. ลอง _MEIPASS ก่อน (subprocess อาจไม่มี)
if hasattr(sys, "_MEIPASS"):
    base_dirs.append(sys._MEIPASS)

# 2. ลองโฟลเดอร์ที่ .exe อยู่
if getattr(sys, 'frozen', False):
    base_dirs.append(os.path.dirname(sys.executable))

# 3. ลองใช้ environment variable ที่ Main Process ส่งมา
if 'MAIN_PROGRAM_DIR' in os.environ:
    base_dirs.append(os.environ['MAIN_PROGRAM_DIR'])
    base_dirs.append(os.path.join(os.environ['MAIN_PROGRAM_DIR'], '_internal'))

# 4. ลองหาจาก path ของไฟล์นี้เอง
base_dirs.append(os.path.dirname(os.path.abspath(__file__)))
base_dirs.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

print(f"SPSS_FIXER_DEBUG: Looking for savReaderWriter/spssio in: {base_dirs}")

for base_dir in base_dirs:
    spss_home_path = os.path.join(base_dir, 'savReaderWriter', 'spssio')
    print(f"SPSS_FIXER_DEBUG: Checking: {spss_home_path} -> exists: {os.path.isdir(spss_home_path)}")
    if os.path.isdir(spss_home_path):
        os.environ['SPSS_HOME'] = spss_home_path
        print(f"SPSS_FIXER_DEBUG: SPSS_HOME set to: {spss_home_path}")
        # เพิ่มโฟลเดอร์ที่มี package savReaderWriter เข้า sys.path
        # เพื่อให้ import ได้แม้รันไฟล์นี้ตรงๆ จาก All_Programs
        if base_dir not in sys.path:
            sys.path.insert(0, base_dir)
            print(f"SPSS_FIXER_DEBUG: Added to sys.path: {base_dir}")
        _spss_found = True
        break

if not _spss_found:
    print("SPSS_FIXER_WARNING: Could not find savReaderWriter/spssio folder!")
# =============================================================================


# --- ส่วน import หลักของโปรแกรม ---
import tkinter as tk
from tkinter import filedialog, messagebox
# --- ส่วนที่แก้ไข: import SavReader และ SavWriter โดยตรง ---
from savReaderWriter.savReader import SavReader
from savReaderWriter.savWriter import SavWriter


# --- ฟังก์ชันช่วยเหลือ ---
def _decode_recursive(obj, encoding):
    """ถอดรหัส bytes ทุกตัว (ทั้ง key และ value ใน dict/list/tuple) ด้วย encoding ที่กำหนด"""
    if isinstance(obj, bytes):
        return obj.decode(encoding, errors="replace")
    if isinstance(obj, dict):
        return {_decode_recursive(k, encoding): _decode_recursive(v, encoding)
                for k, v in obj.items()}
    if isinstance(obj, list):
        return [_decode_recursive(x, encoding) for x in obj]
    if isinstance(obj, tuple):
        return tuple(_decode_recursive(x, encoding) for x in obj)
    return obj


def center_window(window, width, height):
    """ฟังก์ชันสำหรับจัดหน้าต่างให้อยู่กึ่งกลางจอ"""
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = int((screen_width / 2) - (width / 2))
    y = int((screen_height / 2) - (height / 2))
    window.geometry(f'{width}x{height}+{x}+{y}')

# --- ฟังก์ชันการทำงานของโปรแกรม ---
def select_file():
    """ฟังก์ชันสำหรับเปิดหน้าต่างเลือกไฟล์ SPSS (.sav)"""
    file_path = filedialog.askopenfilename(
        title="เลือกไฟล์ SPSS ต้นฉบับ",
        filetypes=(("SPSS Files", "*.sav"), ("All files", "*.*"))
    )
    if file_path:
        # entry_filepath ถูกสร้างใน run_this_app และถูกเข้าถึงผ่าน global
        entry_filepath.delete(0, tk.END)
        entry_filepath.insert(0, file_path)

def clone_and_fix_spss_file():
    """
    ฟังก์ชันสำหรับโคลนไฟล์ SPSS ทั้งข้อมูลและ Metadata
    แล้วบันทึกเป็นไฟล์ใหม่ที่เข้ารหัสเป็น UTF-8
    """
    spss_file_path = entry_filepath.get()

    if not spss_file_path:
        messagebox.showerror("ข้อผิดพลาด", "กรุณาเลือกไฟล์ SPSS ก่อน")
        return

    if not os.path.exists(spss_file_path):
        messagebox.showerror("ข้อผิดพลาด", "ไม่พบไฟล์ที่ระบุ")
        return

    # encoding ของไฟล์ SPSS ภาษาไทยรุ่นเก่า (TIS-620 / Windows-874)
    SOURCE_ENCODING = "cp874"

    import tempfile
    import shutil
    import uuid

    # SPSS C DLL เปิด/สร้างไฟล์ที่ path มีภาษาไทย (non-ASCII) ไม่ได้ -> SPSS_FILE_OERROR
    # วิธีแก้: ทำงานผ่านโฟลเดอร์ temp ที่เป็น ASCII ล้วน แล้วค่อย move ไฟล์กลับด้วย Python
    temp_dir = tempfile.mkdtemp(prefix="spss_fix_")
    temp_in = os.path.join(temp_dir, "input.sav")
    temp_out = os.path.join(temp_dir, "output.sav")

    try:
        # คัดลอกไฟล์ต้นฉบับไปยัง temp path (ASCII) เพื่อให้ DLL อ่านได้ชัวร์
        shutil.copy2(spss_file_path, temp_in)

        # --- ขั้นตอนที่ 1: อ่านทุกอย่างจากไฟล์ต้นฉบับ ---
        # สำคัญ: อ่านเป็น bytes ดิบ (ioUtf8=False) แล้วถอดรหัสด้วย cp874 เอง
        # ไม่บังคับ UTF-8 ตอนอ่าน เพราะไฟล์ต้นฉบับยังไม่ใช่ UTF-8
        print(f"กำลังอ่านไฟล์ต้นฉบับและ Metadata: {spss_file_path}")
        records = []
        metadata = {}

        with SavReader(temp_in, ioUtf8=False) as reader:
            metadata['varNames'] = _decode_recursive(reader.header, SOURCE_ENCODING)
            metadata['varTypes'] = _decode_recursive(getattr(reader, 'varTypes', {}), SOURCE_ENCODING)
            metadata['varLabels'] = _decode_recursive(getattr(reader, 'varLabels', {}), SOURCE_ENCODING)
            metadata['valueLabels'] = _decode_recursive(getattr(reader, 'valueLabels', {}), SOURCE_ENCODING)
            metadata['missingValues'] = _decode_recursive(getattr(reader, 'missingValues', {}), SOURCE_ENCODING)
            metadata['measureLevels'] = _decode_recursive(getattr(reader, 'measureLevels', {}), SOURCE_ENCODING)
            metadata['columnWidths'] = _decode_recursive(getattr(reader, 'columnWidths', {}), SOURCE_ENCODING)
            metadata['formats'] = _decode_recursive(getattr(reader, 'formats', {}), SOURCE_ENCODING)

            for record in reader:
                records.append(_decode_recursive(record, SOURCE_ENCODING))
        print("อ่านข้อมูลและ Metadata ทั้งหมดสำเร็จ")

        # --- ขยายความกว้างของตัวแปร String ---
        # ภาษาไทย 1 ตัวอักษร: cp874 ใช้ 1 byte แต่ UTF-8 ใช้ 3 bytes
        # ถ้าไม่ขยายความกว้าง ข้อความไทยจะถูกตัดหายตอนเขียนเป็น UTF-8
        # สำคัญ: ต้องแก้ทั้ง varTypes (ความกว้าง) และ formats (เช่น A255) ให้ตรงกัน
        #        มิฉะนั้น SPSS จะ error ว่า format misspecified
        MAX_STR_WIDTH = 32767  # ความยาว String สูงสุดของ SPSS
        var_types = metadata.get('varTypes', {})
        formats = metadata.get('formats', {})
        for var_name, type_len in list(var_types.items()):
            if isinstance(type_len, int) and type_len > 0:  # > 0 = ตัวแปร String
                new_width = min(type_len * 3, MAX_STR_WIDTH)
                var_types[var_name] = new_width
                # format ของ String ต้องเป็น A<width> ให้ตรงกับความกว้างใหม่
                formats[var_name] = f"A{new_width}"

        # --- ขั้นตอนที่ 2: เขียนทุกอย่างลงในไฟล์ใหม่ ---
        base_name = os.path.basename(spss_file_path)
        file_name_without_ext = os.path.splitext(base_name)[0]
        output_dir = os.path.dirname(spss_file_path)
        output_sav_path = os.path.join(output_dir, f"{file_name_without_ext}_preserved_utf8.sav")

        print(f"กำลังเขียนไฟล์ใหม่ที่สมบูรณ์: {output_sav_path}")
        # เขียนลง temp path (ASCII) ก่อน เพราะ DLL เขียน path ภาษาไทยไม่ได้
        with SavWriter(temp_out, ioUtf8=True, **metadata) as writer:
            for record in records:
                writer.writerow(record)

        # ย้ายไฟล์ผลลัพธ์จาก temp ไปยังปลายทางจริง (Python รองรับ unicode path)
        # ถ้ามีไฟล์เดิมอยู่แล้วให้เขียนทับ
        if os.path.exists(output_sav_path):
            os.remove(output_sav_path)
        shutil.move(temp_out, output_sav_path)

        print("เขียนไฟล์ใหม่สำเร็จ!")
        messagebox.showinfo("สำเร็จสมบูรณ์!", f"ไฟล์ SPSS ใหม่ถูกสร้างขึ้นแล้ว\n\nบันทึกที่: {output_sav_path}")

    except Exception as e:
        # พิมพ์ traceback เพื่อให้เราเห็นรายละเอียดข้อผิดพลาดใน Console
        import traceback
        traceback.print_exc()
        
        detailed_error = (
            "เกิดข้อผิดพลาดในระหว่างการโคลนไฟล์\n\n"
            f"ประเภทข้อผิดพลาด: {type(e).__name__}\n"
            f"ข้อความ: {e}"
        )
        messagebox.showerror("เกิดข้อผิดพลาด", detailed_error)

    finally:
        # ลบโฟลเดอร์ temp ทิ้งเสมอ ไม่ว่าจะสำเร็จหรือไม่
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass

# --- Entry Point หลักของโปรแกรมนี้ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None):
    """
    ฟังก์ชันหลักสำหรับสร้างและรันหน้าจอโปรแกรม
    """
    print("--- SPSS Fixer: Starting via run_this_app() ---")
    try:
        # สร้าง GUI ทั้งหมดภายในฟังก์ชันนี้
        root = tk.Tk()
        root.title("โปรแกรมซ่อมไฟล์ SPSS V1")
        root.resizable(False, False)

        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(expand=True, fill=tk.BOTH)

        label_filepath = tk.Label(main_frame, text="เลือกไฟล์ SPSS ที่ต้องการซ่อม:")
        label_filepath.pack(pady=(0, 5), anchor='w')

        # จัด Entry และ Button ให้อยู่ใน Frame เดียวกันเพื่อความสวยงาม
        entry_frame = tk.Frame(main_frame)
        entry_frame.pack(fill=tk.X, expand=True, pady=5)
        
        # ทำให้ entry_filepath เป็นตัวแปรที่ฟังก์ชันอื่นเข้าถึงได้
        global entry_filepath
        entry_filepath = tk.Entry(entry_frame)
        entry_filepath.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_browse = tk.Button(entry_frame, text="เลือกไฟล์...", command=select_file)
        btn_browse.pack(side=tk.LEFT, padx=(5, 0))

        btn_convert = tk.Button(main_frame, text="สร้างไฟล์ SAV ใหม่ (คงค่าทั้งหมด & แก้ไข UTF-8)", command=clone_and_fix_spss_file, bg="#28A745", fg="white", font=('Helvetica', 10, 'bold'))
        btn_convert.pack(pady=10, fill=tk.X, ipady=8)

        # จัดหน้าต่างให้อยู่กลางจอหลังจากสร้าง Widget ทั้งหมด
        root.update_idletasks()
        center_window(root, 550, 160) # ปรับความสูงเล็กน้อยให้พอดี

        root.mainloop()
        print("--- SPSS Fixer: Mainloop finished. ---")

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"SPSS_FIXER_ERROR: An error occurred during execution: {e}")
        # สร้าง Popup ชั่วคราวถ้ามีปัญหา
        root_temp = tk.Tk()
        root_temp.withdraw()
        messagebox.showerror("Application Error", f"An unexpected error occurred:\n{e}", parent=root_temp)
        root_temp.destroy()

# --- ส่วนที่ใช้สำหรับรันไฟล์นี้โดยตรงเพื่อทดสอบ (ไม่เกี่ยวกับ Launcher) ---
if __name__ == "__main__":
    print("--- Running SPSS Fixer module directly for testing ---")
    run_this_app()
    print("--- Finished direct execution ---")
