# updater.py (เวอร์ชันมี GUI และไอคอน)
import sys
import os
import time
import requests
import psutil
import threading
import tkinter as tk
from tkinter import ttk
from tkinter.font import Font
import zipfile
import shutil
import tempfile
import subprocess

# --- เพิ่มเข้ามา: ฟังก์ชันสำหรับหา Path ของไฟล์ที่แนบมากับ .exe ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)
# -------------------------------------------------------------

class UpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("กำลังอัปเดต...")
        self._relaunch_in_progress = False
        self.root.withdraw()

        # --- เพิ่มเข้ามา: ตั้งค่าไอคอนของหน้าต่าง ---
        try:
            icon_path = resource_path("setting.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Could not load icon: {e}")
        # ----------------------------------------

        self.root.geometry("400x120")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.disable_close)

        # ... (ส่วนที่เหลือของโค้ดเหมือนเดิมทุกประการ) ...

        # รับ arguments จากโปรแกรมหลัก
        self.running_from_temp = False
        self.ready_file_path = None
        if "--run-from-temp" in sys.argv:
            self.running_from_temp = True
        if "--ready-file" in sys.argv:
            try:
                idx = sys.argv.index("--ready-file")
                self.ready_file_path = sys.argv[idx + 1]
            except Exception:
                self.ready_file_path = None

        clean_args = []
        skip_next = False
        for arg in sys.argv[1:]:
            if skip_next:
                skip_next = False
                continue
            if arg == "--run-from-temp":
                continue
            if arg == "--ready-file":
                skip_next = True
                continue
            clean_args.append(arg)

        if len(clean_args) not in (3, 4):
            raise RuntimeError("Invalid updater arguments.")

        self.parent_pid = int(clean_args[0])
        self.app_dir = None
        self.exe_name = None
        self.old_exe_path = None
        self.update_url = None
        if len(clean_args) == 3:
            self.old_exe_path = clean_args[1]
            self.update_url = clean_args[2]
            self.app_dir = os.path.dirname(self.old_exe_path)
            self.exe_name = os.path.basename(self.old_exe_path)
        else:
            self.app_dir = clean_args[1]
            self.exe_name = clean_args[2]
            self.update_url = clean_args[3]

        self._maybe_relaunch_from_temp()
        if self._relaunch_in_progress:
            return

        # สร้าง Widgets
        self.status_label = tk.Label(root, text="กำลังเตรียมอัปเดต...", font=Font(size=10))
        self.status_label.pack(pady=10)

        self.progress = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
        self.progress.pack(pady=5)
        
        self.percent_label = tk.Label(root, text="0%", font=Font(size=9))
        self.percent_label.pack()
        self.root.deiconify()
        
        # เริ่มกระบวนการอัปเดตใน Thread ใหม่ เพื่อไม่ให้ GUI ค้าง
        self.update_thread = threading.Thread(target=self.run_update_process, daemon=True)
        if self.running_from_temp and self.ready_file_path:
            try:
                with open(self.ready_file_path, "w", encoding="utf-8") as f:
                    f.write("ready")
            except Exception:
                pass
        self.update_thread.start()

    def disable_close(self):
        pass

    def _maybe_relaunch_from_temp(self):
        if "--run-from-temp" in sys.argv:
            self.running_from_temp = True
            return
        if not self.app_dir:
            return
        exe_path = sys.executable
        if not exe_path.lower().endswith(".exe"):
            return
        try:
            app_dir = os.path.abspath(self.app_dir)
            exe_path_abs = os.path.abspath(exe_path)
            if os.path.commonpath([exe_path_abs, app_dir]) != app_dir:
                return
        except Exception:
            return

        temp_dir = tempfile.mkdtemp(prefix="Updater_", dir=tempfile.gettempdir())
        temp_exe = os.path.join(temp_dir, os.path.basename(exe_path))
        try:
            shutil.copy2(exe_path, temp_exe)
        except Exception:
            return

        ready_file = os.path.join(temp_dir, "updater_ready.flag")
        args = [temp_exe] + sys.argv[1:] + ["--run-from-temp", "--ready-file", ready_file]
        try:
            proc = subprocess.Popen(args, close_fds=True)
            alive = False
            for _ in range(10):
                time.sleep(0.3)
                if os.path.exists(ready_file):
                    alive = True
                    break
            if alive:
                self._relaunch_in_progress = True
                try:
                    self.root.withdraw()
                except Exception:
                    pass
                self.root.after(200, self.root.destroy)
            else:
                self.status_label.config(
                    text="ไม่สามารถเริ่มตัวอัปเดตชั่วคราว กำลังดำเนินการต่อ...",
                    fg="red"
                )
        except Exception:
            return

    def _wait_for_process_exit(self, timeout=30):
        start_time = time.time()
        while psutil.pid_exists(self.parent_pid):
            if time.time() - start_time > timeout:
                return False
            time.sleep(0.5)
        return True

    def _kill_process_by_name(self, exe_name):
        killed = False
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and proc.info['name'].lower() == exe_name.lower():
                    proc.terminate()
                    try:
                        proc.wait(timeout=5)
                    except Exception:
                        proc.kill()
                    killed = True
            except Exception:
                continue
        return killed

    def _kill_processes_in_app_dir(self, app_dir):
        killed = 0
        try:
            app_dir_norm = os.path.normcase(os.path.abspath(app_dir)) + os.sep
        except Exception:
            return killed
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            try:
                if proc.pid == os.getpid():
                    continue
                exe_path = proc.info.get('exe')
                if not exe_path:
                    continue
                exe_norm = os.path.normcase(os.path.abspath(exe_path))
                if exe_norm.startswith(app_dir_norm):
                    proc.terminate()
                    try:
                        proc.wait(timeout=5)
                    except Exception:
                        proc.kill()
                    killed += 1
            except Exception:
                continue
        return killed

    def _safe_rename(self, src, dst, retries=10, delay=0.5):
        for _ in range(retries):
            try:
                shutil.move(src, dst)
                return True
            except Exception:
                time.sleep(delay)
        return False

    def run_update_process(self):
        zip_path = None
        temp_dir = None
        try:
            # 1. รอให้โปรแกรมหลักปิดตัว
            self.status_label.config(text="กำลังรอให้โปรแกรมหลักปิดตัว...")
            self.root.update_idletasks()
            
            if not self._wait_for_process_exit(timeout=30):
                self.status_label.config(text="ยังมีโปรแกรมค้างอยู่ กำลังพยายามปิด...", fg="red")
                self.root.update_idletasks()
                if self.exe_name:
                    self._kill_process_by_name(self.exe_name)
                self._wait_for_process_exit(timeout=10)

            # 2. ดาวน์โหลดไฟล์เวอร์ชันใหม่
            self.status_label.config(text="กำลังดาวน์โหลดเวอร์ชันใหม่...")
            work_dir = tempfile.gettempdir()
            zip_path = os.path.join(work_dir, "Main_Program_update.zip")
            new_exe_path = None
            temp_dir = tempfile.mkdtemp(prefix="Main_Program_update_tmp_", dir=work_dir)
            if os.path.exists(zip_path):
                try:
                    os.remove(zip_path)
                except Exception:
                    pass

            with requests.get(self.update_url, stream=True) as r:
                r.raise_for_status()
                total_size = int(r.headers.get('content-length', 0))
                downloaded_size = 0

                is_zip = self.update_url.lower().endswith(".zip")
                target_path = zip_path if is_zip else (self.old_exe_path + ".new")
                new_exe_path = target_path if not is_zip else None

                with open(target_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
                        downloaded_size += len(chunk)

                        progress_percent = (downloaded_size / total_size) * 100 if total_size > 0 else 0
                        self.progress['value'] = progress_percent
                        self.percent_label.config(text=f"{progress_percent:.0f}%")
                        self.root.update_idletasks()
            
            # 3. ติดตั้งอัปเดต
            self.status_label.config(text="กำลังติดตั้งอัปเดต...")
            self.percent_label.config(text="")
            self.root.update_idletasks()
            time.sleep(1)
            
            if self.exe_name:
                self._kill_process_by_name(self.exe_name)
            if self.app_dir:
                self._kill_processes_in_app_dir(self.app_dir)
                time.sleep(0.5)

            if self.update_url.lower().endswith(".zip"):
                with zipfile.ZipFile(zip_path, 'r') as zf:
                    zf.extractall(temp_dir)

                entries = [d for d in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, d))]
                if len(entries) == 1:
                    new_app_dir = os.path.join(temp_dir, entries[0])
                else:
                    new_app_dir = temp_dir

                backup_dir = self.app_dir + ".old"
                if os.path.exists(backup_dir):
                    shutil.rmtree(backup_dir, ignore_errors=True)
                if not self._safe_rename(self.app_dir, backup_dir):
                    self._kill_processes_in_app_dir(self.app_dir)
                    time.sleep(1.0)
                    if not self._safe_rename(self.app_dir, backup_dir):
                        raise RuntimeError("ไม่สามารถย้ายโฟลเดอร์เดิมได้ (ยังถูกใช้งานอยู่)")
                if not self._safe_rename(new_app_dir, self.app_dir):
                    raise RuntimeError("ไม่สามารถย้ายโฟลเดอร์ใหม่ได้")
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir, ignore_errors=True)
                if os.path.exists(zip_path):
                    os.remove(zip_path)
            else:
                if self.old_exe_path is None or new_exe_path is None:
                    raise RuntimeError("Invalid updater arguments for EXE update.")
                if self.app_dir:
                    self._kill_processes_in_app_dir(self.app_dir)
                os.remove(self.old_exe_path)
                os.rename(new_exe_path, self.old_exe_path)
            
            # 4. เปิดโปรแกรมเวอร์ชันใหม่ขึ้นมา
            new_exe_path = os.path.join(self.app_dir, self.exe_name)
            launched = False
            for _ in range(5):
                if os.path.exists(new_exe_path):
                    try:
                        os.startfile(new_exe_path)
                        launched = True
                        break
                    except Exception:
                        time.sleep(0.5)
                else:
                    time.sleep(0.5)

            if not launched:
                self.status_label.config(text="อัปเดตเสร็จแล้ว กรุณาเปิดโปรแกรมใหม่อีกครั้ง", fg="red")
                self.root.after(8000, self.root.quit)
                return

            # 5. ปิดตัวเอง
            self.root.quit()

        except Exception as e:
            self.status_label.config(text=f"เกิดข้อผิดพลาด: {e}", fg="red")
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                except Exception:
                    pass
            if zip_path and os.path.exists(zip_path):
                try:
                    os.remove(zip_path)
                except Exception:
                    pass
            self.root.after(10000, self.root.quit)

if __name__ == "__main__":
    # --- เพิ่มเข้ามา: ตรวจสอบว่ามี arguments ส่งมาหรือไม่ก่อนรัน ---
    # ป้องกัน Error เวลาเผลอดับเบิ้ลคลิกไฟล์ .py โดยตรง
    raw_args = sys.argv[1:]
    clean_args = []
    skip_next = False
    for arg in raw_args:
        if skip_next:
            skip_next = False
            continue
        if arg == "--run-from-temp":
            continue
        if arg == "--ready-file":
            skip_next = True
            continue
        clean_args.append(arg)

    if len(clean_args) not in (3, 4):
        print("This script is intended to be run by the main application.")
        print("Usage: updater.py <pid> <exe_path> <download_url>")
        print("   or: updater.py <pid> <app_dir> <exe_name> <download_url>")
        sys.exit(1)
    # -----------------------------------------------------------
    
    root = tk.Tk()
    app = UpdaterApp(root)
    root.mainloop()
