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
import atexit
import re
import json
import getpass
import socket
from update_cache import find_cached_package, store_cached_package
from release_files import (apply_file_update, recover_pending, RollbackFailed, protected,
                           extract_full_package, install_full_update)

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

def _log_update_event(message):
    try:
        base_dir = os.environ.get("UPDATER_LOG_DIR")
        if not base_dir:
            base_dir = os.path.dirname(sys.executable) if sys.executable else os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(base_dir, "updater_debug.log")
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception:
        pass

def _get_bsdiff4():
    try:
        import bsdiff4  # type: ignore
        return bsdiff4
    except Exception as e:
        _log_update_event(f"bsdiff4 import failed: {e}")
        return None

TELEGRAM_BOT_TOKEN = "8207273310:AAEwpcDWP8yRP5Q74R3ic5jpZ_BOPQwJ_PQ"
TELEGRAM_CHAT_ID = "8556512706"

class UpdaterApp:
    def __init__(self, root):
        _log_update_event(f"Updater start argv={sys.argv}")
        self.root = root
        self.root.title("กำลังอัปเดต...")
        self._relaunch_in_progress = False
        self._lock_path = None
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

        update_kind = "full"
        current_version = None
        new_version = None
        patch_manifest_path = None
        release_url = None
        clean_args = []
        skip_next = False
        args = sys.argv[1:]
        idx = 0
        while idx < len(args):
            arg = args[idx]
            if skip_next:
                skip_next = False
                idx += 1
                continue
            if arg == "--run-from-temp":
                idx += 1
                continue
            if arg == "--ready-file":
                skip_next = True
                idx += 1
                continue
            if arg == "--update-kind" and idx + 1 < len(args):
                update_kind = args[idx + 1]
                idx += 2
                continue
            if arg == "--current-version" and idx + 1 < len(args):
                current_version = args[idx + 1]
                idx += 2
                continue
            if arg == "--new-version" and idx + 1 < len(args):
                new_version = args[idx + 1]
                idx += 2
                continue
            if arg == "--patch-manifest" and idx + 1 < len(args):
                patch_manifest_path = args[idx + 1]
                idx += 2
                continue
            if arg == "--release-url" and idx + 1 < len(args):
                release_url = args[idx + 1]
                idx += 2
                continue
            clean_args.append(arg)
            idx += 1

        if len(clean_args) not in (3, 4):
            _log_update_event(f"Invalid updater arguments: {clean_args}")
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

        self.update_kind = update_kind
        self.current_version = current_version
        self.new_version = new_version
        self.patch_manifest_path = patch_manifest_path
        self.release_url = release_url
        if self.app_dir:
            os.environ["UPDATER_LOG_DIR"] = self.app_dir
        _log_update_event(
            f"Parsed args: kind={self.update_kind}, current={self.current_version}, "
            f"new={self.new_version}, update_url={self.update_url}, patch_manifest={self.patch_manifest_path}, "
            f"release_url={self.release_url}"
        )

        self._maybe_relaunch_from_temp()
        if self._relaunch_in_progress:
            _log_update_event("Relaunching from temp; exiting original updater.")
            return
        if not self._acquire_lock():
            _log_update_event("Updater lock already held; exiting.")
            try:
                self.root.destroy()
            except Exception:
                pass
            return

        # สร้าง Widgets
        self.status_label = tk.Label(root, text="กำลังเตรียมอัปเดต...", font=Font(size=10))
        self.status_label.pack(pady=10)

        self.progress = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
        self.progress.pack(pady=5)
        
        self.percent_label = tk.Label(root, text="0%", font=Font(size=9))
        self.percent_label.pack()
        self._center_window(400, 120)
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

    def _center_window(self, width, height):
        try:
            self.root.update_idletasks()
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            x = int((screen_width - width) / 2)
            y = int((screen_height - height) / 2)
            self.root.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
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
            _log_update_event("Failed to copy updater to temp.")
            return

        ready_file = os.path.join(temp_dir, "updater_ready.flag")
        args = [temp_exe] + sys.argv[1:] + ["--run-from-temp", "--ready-file", ready_file]
        try:
            env = os.environ.copy()
            proc = subprocess.Popen(args, close_fds=True, env=env)
            alive = False
            for _ in range(10):
                time.sleep(0.3)
                if os.path.exists(ready_file):
                    alive = True
                    break
            if alive:
                _log_update_event("Temp updater started successfully.")
                self._relaunch_in_progress = True
                try:
                    self.root.withdraw()
                except Exception:
                    pass
                self.root.after(200, self.root.destroy)
            else:
                _log_update_event("Temp updater did not signal readiness.")
                self.status_label.config(
                    text="ไม่สามารถเริ่มตัวอัปเดตชั่วคราว กำลังดำเนินการต่อ...",
                    fg="red"
                )
        except Exception:
            _log_update_event("Failed to launch temp updater.")
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
        installation = os.path.normcase(os.path.realpath(self.app_dir)) + os.sep
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            try:
                executable = os.path.normcase(os.path.realpath(proc.info.get('exe') or ''))
                if (proc.pid != os.getpid() and executable.startswith(installation)
                        and proc.info['name'] and proc.info['name'].lower() == exe_name.lower()):
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
    
    def _acquire_lock(self):
        if not self.app_dir:
            return True
        lock_dir = os.path.dirname(os.path.abspath(self.app_dir))
        self._lock_path = os.path.join(lock_dir, "updater.lock")
        try:
            if os.path.exists(self._lock_path):
                try:
                    with open(self._lock_path, "r", encoding="utf-8") as f:
                        existing_pid = int(f.read().strip() or "0")
                except Exception:
                    existing_pid = 0
                if existing_pid and psutil.pid_exists(existing_pid):
                    _log_update_event(f"Updater lock held by pid {existing_pid}.")
                    return False
            with open(self._lock_path, "w", encoding="utf-8") as f:
                f.write(str(os.getpid()))
            atexit.register(self._release_lock)
            return True
        except Exception:
            return True
    
    def _release_lock(self):
        try:
            if self._lock_path and os.path.exists(self._lock_path):
                os.remove(self._lock_path)
        except Exception:
            pass
    
    def _remove_empty_dir(self, path):
        try:
            if os.path.isdir(path) and not os.listdir(path):
                os.rmdir(path)
                return True
        except Exception:
            pass
        return False
    
    def _safe_remove_path(self, path, retries=3, delay=0.3):
        for _ in range(retries):
            try:
                if os.path.isdir(path):
                    shutil.rmtree(path, ignore_errors=True)
                elif os.path.exists(path):
                    os.remove(path)
                return True
            except Exception:
                time.sleep(delay)
        return False

    def _clean_install_root(self, keep_names):
        if not self.app_dir:
            return
        try:
            for name in os.listdir(self.app_dir):
                if name in keep_names:
                    continue
                self._safe_remove_path(os.path.join(self.app_dir, name))
        except Exception:
            pass

    def _copy_tree_overwrite(self, src, dst, preserve_files=None, preserve_dirs=None):
        preserve_files = set(preserve_files or [])
        preserve_dirs = set(preserve_dirs or [])
        for root_dir, dirnames, filenames in os.walk(src):
            rel_path = os.path.relpath(root_dir, src)
            target_dir = dst if rel_path == "." else os.path.join(dst, rel_path)
            if rel_path == "." and preserve_dirs:
                dirnames[:] = [
                    d for d in dirnames
                    if not (d in preserve_dirs and os.path.exists(os.path.join(target_dir, d)))
                ]
            os.makedirs(target_dir, exist_ok=True)
            for filename in filenames:
                src_path = os.path.join(root_dir, filename)
                dst_path = os.path.join(target_dir, filename)
                relative_name = os.path.relpath(src_path, src).replace(os.sep, '/')
                if protected(relative_name) and os.path.exists(dst_path):
                    continue
                if rel_path == "." and filename in preserve_files and os.path.exists(dst_path):
                    continue
                try:
                    if os.path.exists(dst_path):
                        os.remove(dst_path)
                except Exception:
                    pass
                shutil.copy2(src_path, dst_path)
                try:
                    os.utime(dst_path, None)
                except Exception:
                    pass

    def _ensure_seed_assets(self):
        if not self.app_dir:
            return
        internal_dir = os.path.join(self.app_dir, "_internal")
        if not os.path.isdir(internal_dir):
            return
        seed_file = "Itemdef - Format.xlsx"
        seed_dir = "savReaderWriter"
        seed_key_file = "Test3.json"
        dest_file = os.path.join(self.app_dir, seed_file)
        src_file = os.path.join(internal_dir, seed_file)
        if not os.path.exists(dest_file) and os.path.exists(src_file):
            try:
                shutil.copy2(src_file, dest_file)
            except Exception:
                pass
        dest_key_file = os.path.join(self.app_dir, seed_key_file)
        src_key_file = os.path.join(internal_dir, seed_key_file)
        if not os.path.exists(dest_key_file) and os.path.exists(src_key_file):
            try:
                shutil.copy2(src_key_file, dest_key_file)
            except Exception:
                pass
        dest_dir = os.path.join(self.app_dir, seed_dir)
        src_dir = os.path.join(internal_dir, seed_dir)
        if not os.path.exists(dest_dir) and os.path.isdir(src_dir):
            try:
                shutil.copytree(src_dir, dest_dir)
            except Exception:
                pass

    def _format_bytes(self, num_bytes):
        try:
            num_bytes = float(num_bytes)
        except Exception:
            return "0 MB"
        return f"{num_bytes / (1024 * 1024):.1f} MB"

    def _get_user_machine_info(self):
        try:
            username = getpass.getuser()
        except Exception:
            username = "Unknown"
        try:
            hostname = socket.gethostname()
        except Exception:
            hostname = "Unknown"
        try:
            ip_address = socket.gethostbyname(hostname)
        except Exception:
            ip_address = "IP N/A"
        return username, hostname, ip_address

    def _send_telegram_update_notice(self):
        if "TELEGRAM_BOT_TOKEN" in TELEGRAM_BOT_TOKEN or "TELEGRAM_CHAT_ID" in TELEGRAM_CHAT_ID:
            print("TELEGRAM_SKIP: Missing token or chat id.")
            return
        old_version = self.current_version or "N/A"
        new_version = self.new_version or "N/A"
        username, hostname, ip_address = self._get_user_machine_info()
        release_url = self.release_url or ""
        message_text = (
            "✅ <b>อัปเดตสำเร็จ</b>\n\n"
            f"🖥️ <b>เครื่อง:</b> {hostname}\n"
            f"👤 <b>ผู้ใช้:</b> {username}\n"
            f"🌐 <b>IP:</b> {ip_address}\n"
            f"🧩 <b>เวอร์ชัน:</b> {old_version} → {new_version}\n\n"
            f"🔗 <b>Release:</b>\n{release_url}"
        )
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        payload = {
            "chat_id": TELEGRAM_CHAT_ID,
            "text": message_text,
            "parse_mode": "HTML",
            "disable_web_page_preview": False
        }
        try:
            response = requests.post(url, json=payload, timeout=8)
            print(f"TELEGRAM_STATUS: {response.status_code}")
            print(f"TELEGRAM_BODY: {response.text}")
        except Exception as e:
            print(f"TELEGRAM_ERROR: {e}")

    def _get_updates_dir(self):
        if not self.app_dir:
            return None
        updates_dir = os.path.join(self.app_dir, "_internal", "updates")
        os.makedirs(updates_dir, exist_ok=True)
        return updates_dir

    def _extract_version_from_filename(self, path_or_url):
        if not path_or_url:
            return None
        filename = os.path.basename(path_or_url)
        match = re.search(r"(\d+\.\d+\.\d+)", filename)
        if match:
            return match.group(1)
        return None

    def _get_cached_zip_path(self, version):
        return find_cached_package(self.app_dir, version)

    def _try_file_update(self, work_dir):
        # Old installed launchers already pass both versions and download the newest updater.
        # Discover the small package directly without an API call or a cached base ZIP.
        if not all(isinstance(v, str) and re.fullmatch(r'\d+\.\d+\.\d+', v)
                   for v in (self.current_version, self.new_version)):
            return False
        base = f'https://github.com/Icezy159753/my-calculator-updates/releases/download/v{self.new_version}'
        url = f'{base}/Main_Program_files_{self.current_version}_to_{self.new_version}.zip'
        package = os.path.join(work_dir, 'file-update.zip')
        try:
            self.status_label.config(text='กำลังดาวน์โหลดเฉพาะไฟล์ที่เปลี่ยน...')
            self.root.update_idletasks()
            with requests.get(url, stream=True, timeout=(15, 90)) as response:
                if response.status_code == 404:
                    return False
                response.raise_for_status()
                total = int(response.headers.get('content-length', 0))
                downloaded = 0
                with open(package, 'wb') as output:
                    for chunk in response.iter_content(1024 * 1024):
                        output.write(chunk)
                        downloaded += len(chunk)
                        self.progress['value'] = downloaded * 100 / total if total else 0
                        self.percent_label.config(text=self._format_bytes(downloaded))
                        self.root.update_idletasks()
            self._kill_processes_in_app_dir(self.app_dir)
            self.status_label.config(text='กำลังตรวจไฟล์และติดตั้งอัปเดต...')
            self.root.update_idletasks()
            apply_file_update(self.app_dir, package, self.current_version, self.new_version)
            _log_update_event(f'File update installed: {self.current_version} -> {self.new_version}; bytes={downloaded}')
            return True
        except RollbackFailed:
            raise  # Keep backups and stop; never install over an incomplete rollback.
        except Exception as error:
            _log_update_event(f'File update rejected; full package fallback: {error}')
            self.status_label.config(text='ตรวจไฟล์เดิมไม่ผ่านหรือดาวน์โหลดไม่สำเร็จ กำลังใช้ชุดโปรแกรมเต็ม...')
            self.update_kind = 'full'
            self.update_url = f'{base}/Main_Program_full_{self.new_version}.zip'
            return False

    def _load_patch_manifest(self):
        if not self.patch_manifest_path:
            return None
        if not os.path.exists(self.patch_manifest_path):
            return None
        try:
            with open(self.patch_manifest_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    def run_update_process(self):
        zip_path = None
        temp_dir = None
        work_dir = None
        try:
            _log_update_event("Update process started.")
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
            self.status_label.config(text="กำลังเตรียมดาวน์โหลดเวอร์ชันใหม่...")
            work_dir = tempfile.mkdtemp(prefix='Main_Program_download_')
            zip_path = os.path.join(work_dir, "Main_Program_update.zip")
            new_exe_path = None
            temp_dir = tempfile.mkdtemp(prefix="Main_Program_update_tmp_", dir=work_dir)
            self._kill_processes_in_app_dir(self.app_dir)
            recover_pending(self.app_dir)
            if self._try_file_update(temp_dir):
                new_exe = os.path.join(self.app_dir, self.exe_name)
                try:
                    os.startfile(new_exe)
                except OSError:
                    self.status_label.config(text='อัปเดตสำเร็จ กรุณาเปิดโปรแกรมใหม่', fg='red')
                    self.root.after(8000, self.root.quit)
                    return
                shutil.rmtree(temp_dir, ignore_errors=True)
                self._send_telegram_update_notice()
                self.root.quit()
                return
            if self.update_kind == 'files':
                self.update_kind = 'full'
                self.update_url = f'https://github.com/Icezy159753/my-calculator-updates/releases/download/v{self.new_version}/Main_Program_full_{self.new_version}.zip'
            if os.path.exists(zip_path):
                try:
                    os.remove(zip_path)
                except Exception:
                    pass

            kind = (self.update_kind or "").lower()
            _log_update_event(f"Update kind={kind}, url={self.update_url}")
            is_patch = kind == "patch" or self.update_url.lower().endswith(".bsdiff")
            is_patch_chain = kind == "patch-chain"
            is_zip = self.update_url.lower().endswith(".zip") or is_patch
            patch_chain_failed = False
            if is_patch_chain:
                try:
                    self.status_label.config(text="กำลังดาวน์โหลด Patch แบบต่อเนื่อง...")
                    self.root.update_idletasks()
                    manifest = self._load_patch_manifest()
                    if not manifest:
                        raise RuntimeError("Missing patch manifest.")
                    chain = manifest.get("patches", [])
                    if not chain:
                        raise RuntimeError("Patch chain empty.")
                    cached_zip_path = self._get_cached_zip_path(self.current_version)
                    if not cached_zip_path or not os.path.exists(cached_zip_path):
                        raise RuntimeError("Missing cached package for patch chain.")

                    current_zip_path = cached_zip_path
                    for idx, entry in enumerate(chain, start=1):
                        patch_url = entry.get("url")
                        to_version = entry.get("to")
                        if not patch_url or not to_version:
                            raise RuntimeError("Invalid patch entry.")
                        patch_path = os.path.join(work_dir, f"Main_Program_patch_{idx}.bsdiff")
                        if os.path.exists(patch_path):
                            try:
                                os.remove(patch_path)
                            except Exception:
                                pass
                        self.status_label.config(text=f"ดาวน์โหลด Patch {idx}/{len(chain)}...")
                        self.root.update_idletasks()
                        with requests.get(patch_url, stream=True) as r:
                            r.raise_for_status()
                            total_size = int(r.headers.get('content-length', 0))
                            downloaded_size = 0
                            with open(patch_path, 'wb') as f:
                                for chunk in r.iter_content(chunk_size=8192):
                                    f.write(chunk)
                                    downloaded_size += len(chunk)
                                    progress_percent = (downloaded_size / total_size) * 100 if total_size > 0 else 0
                                    self.progress['value'] = progress_percent
                                    size_text = f"{self._format_bytes(downloaded_size)} / {self._format_bytes(total_size)}"
                                    self.percent_label.config(text=f"{progress_percent:.0f}% ({size_text})")
                                    self.root.update_idletasks()
                        self.status_label.config(text=f"กำลังสร้างไฟล์อัปเดต {idx}/{len(chain)}...")
                        self.percent_label.config(text="")
                        self.root.update_idletasks()
                        next_zip_path = os.path.join(work_dir, f"Main_Program_update_{to_version}.zip")
                        if os.path.exists(next_zip_path):
                            try:
                                os.remove(next_zip_path)
                            except Exception:
                                pass
                        bsdiff4 = _get_bsdiff4()
                        if not bsdiff4:
                            raise RuntimeError("bsdiff4 not available.")
                        bsdiff4.file_patch(current_zip_path, next_zip_path, patch_path)
                        current_zip_path = next_zip_path
                    zip_path = current_zip_path
                    is_zip = True
                    self.new_version = manifest.get("target_version", self.new_version)
                except Exception as e:
                    patch_chain_failed = True
                    _log_update_event(f"Full package fallback: patch chain failed: {e}")
                    self.status_label.config(text=f"Patch ล้มเหลว กำลังดาวน์โหลดไฟล์เต็ม... ({e})", fg="red")
                    self.root.update_idletasks()
                    is_patch_chain = False
                    is_patch = False
                    kind = "full"

            if (not is_patch_chain) and is_patch:
                self.status_label.config(text="กำลังดาวน์โหลด Patch (ขนาดเล็ก)...")
                self.root.update_idletasks()
                if not self.current_version or not self.new_version:
                    raise RuntimeError("Missing version info for patch update.")
                cached_zip_path = self._get_cached_zip_path(self.current_version)
                if not cached_zip_path or not os.path.exists(cached_zip_path):
                    raise RuntimeError("Missing cached package for patch update.")

                patch_path = os.path.join(work_dir, "Main_Program_update.bsdiff")
                if os.path.exists(patch_path):
                    try:
                        os.remove(patch_path)
                    except Exception:
                        pass

                with requests.get(self.update_url, stream=True) as r:
                    r.raise_for_status()
                    total_size = int(r.headers.get('content-length', 0))
                    downloaded_size = 0
                    with open(patch_path, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=8192):
                            f.write(chunk)
                            downloaded_size += len(chunk)
                            progress_percent = (downloaded_size / total_size) * 100 if total_size > 0 else 0
                            self.progress['value'] = progress_percent
                            size_text = f"{self._format_bytes(downloaded_size)} / {self._format_bytes(total_size)}"
                            self.percent_label.config(text=f"{progress_percent:.0f}% ({size_text})")
                            self.root.update_idletasks()

                self.status_label.config(text="กำลังสร้างไฟล์อัปเดตจาก Patch...")
                self.percent_label.config(text="")
                self.root.update_idletasks()
                zip_path = os.path.join(work_dir, f"Main_Program_update_{self.new_version}.zip")
                if os.path.exists(zip_path):
                    try:
                        os.remove(zip_path)
                    except Exception:
                        pass
                try:
                    bsdiff4 = _get_bsdiff4()
                    if not bsdiff4:
                        raise RuntimeError("bsdiff4 not available.")
                    bsdiff4.file_patch(cached_zip_path, zip_path, patch_path)
                except Exception:
                    bsdiff4 = _get_bsdiff4()
                    if not bsdiff4:
                        raise RuntimeError("bsdiff4 not available.")
                    with open(cached_zip_path, "rb") as old_f:
                        old_data = old_f.read()
                    with open(patch_path, "rb") as patch_f:
                        patch_data = patch_f.read()
                    new_data = bsdiff4.patch(old_data, patch_data)
                    with open(zip_path, "wb") as new_f:
                        new_f.write(new_data)
            else:
                self.status_label.config(text="กำลังดาวน์โหลดไฟล์เต็ม (Full package)...")
                self.root.update_idletasks()
                with requests.get(self.update_url, stream=True) as r:
                    r.raise_for_status()
                    total_size = int(r.headers.get('content-length', 0))
                    downloaded_size = 0

                    target_path = zip_path if is_zip else (self.old_exe_path + ".new")
                    new_exe_path = target_path if not is_zip else None

                    with open(target_path, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=8192):
                            f.write(chunk)
                            downloaded_size += len(chunk)

                            progress_percent = (downloaded_size / total_size) * 100 if total_size > 0 else 0
                            self.progress['value'] = progress_percent
                            size_text = f"{self._format_bytes(downloaded_size)} / {self._format_bytes(total_size)}"
                            self.percent_label.config(text=f"{progress_percent:.0f}% ({size_text})")
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

            if is_zip:
                extract_full_package(zip_path, temp_dir)

                entries = [d for d in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, d))]
                if self.exe_name and os.path.exists(os.path.join(temp_dir, self.exe_name)):
                    new_app_dir = temp_dir
                elif len(entries) == 1:
                    new_app_dir = os.path.join(temp_dir, entries[0])
                else:
                    new_app_dir = temp_dir

                if not self.app_dir:
                    raise RuntimeError("ไม่พบโฟลเดอร์ติดตั้งของโปรแกรม")
                os.makedirs(self.app_dir, exist_ok=True)
                install_full_update(
                    self.app_dir, new_app_dir,
                    self.current_version or '0.0.0',
                    self.new_version or self._extract_version_from_filename(self.update_url) or '0.0.0',
                    work_dir,
                )
                self._ensure_seed_assets()
                cache_warning = None
                try:
                    cache_version = self.new_version or self._extract_version_from_filename(self.update_url)
                    cached_new_zip = store_cached_package(self.app_dir, cache_version, zip_path)
                    _log_update_event(f"Package cache saved: {cached_new_zip}")
                except (OSError, ValueError) as error:
                    cache_warning = 'อัปเดตสำเร็จ แต่เก็บไฟล์สำหรับอัปเดตครั้งถัดไปไม่ได้ ครั้งถัดไปอาจต้องดาวน์โหลดไฟล์เต็ม'
                    _log_update_event(f"Package cache save failed: {error}")
                if cache_warning:
                    from tkinter import messagebox
                    messagebox.showwarning('ไฟล์สำหรับอัปเดตครั้งถัดไป', cache_warning, parent=self.root)
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir, ignore_errors=True)
                if zip_path and os.path.exists(zip_path):
                    try:
                        os.remove(zip_path)
                    except Exception:
                        pass
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

            self._send_telegram_update_notice()
            # 5. ปิดตัวเอง
            self.root.quit()

        except Exception as e:
            _log_update_event(f"Update process error: {e}")
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
        finally:
            self._release_lock()
            if work_dir and os.path.realpath(work_dir).startswith(os.path.realpath(tempfile.gettempdir()) + os.sep):
                shutil.rmtree(work_dir, ignore_errors=True)

if __name__ == "__main__":
    if sys.argv[1:] == ['--self-test']:
        # CI exercises the packaged updater without opening windows or touching an installation.
        from pathlib import Path
        from release_files import inventory, make_file_package
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            old, new, installed = (base / name for name in ('old', 'new', 'installed'))
            for folder in (old, new, installed):
                folder.mkdir()
            (old / 'Main_Program.exe').write_bytes(b'old')
            (new / 'Main_Program.exe').write_bytes(b'new')
            shutil.copytree(old, installed, dirs_exist_ok=True)
            package = base / 'update.zip'
            make_file_package(inventory(old), new, '1.1.96', '1.1.97', package)
            apply_file_update(installed, package, '1.1.96', '1.1.97')
            if inventory(installed) != inventory(new):
                raise RuntimeError('Frozen updater transaction failed')
        if sys.stdout is None:
            import ctypes
            import msvcrt
            handle = ctypes.windll.kernel32.GetStdHandle(-11)
            sys.stdout = open(msvcrt.open_osfhandle(handle, os.O_WRONLY), 'w', encoding='utf-8')
        print('Updater transaction self-test OK', flush=True)
        raise SystemExit(0)
    # --- เพิ่มเข้ามา: ตรวจสอบว่ามี arguments ส่งมาหรือไม่ก่อนรัน ---
    # ป้องกัน Error เวลาเผลอดับเบิ้ลคลิกไฟล์ .py โดยตรง
    raw_args = sys.argv[1:]
    clean_args = []
    skip_next = False
    idx = 0
    while idx < len(raw_args):
        arg = raw_args[idx]
        if skip_next:
            skip_next = False
            idx += 1
            continue
        if arg == "--run-from-temp":
            idx += 1
            continue
        if arg == "--ready-file":
            skip_next = True
            idx += 1
            continue
        if arg in ("--update-kind", "--current-version", "--new-version", "--patch-manifest", "--release-url"):
            idx += 2
            continue
        clean_args.append(arg)
        idx += 1

    if len(clean_args) not in (3, 4):
        print("This script is intended to be run by the main application.")
        print("Usage: updater.py <pid> <exe_path> <download_url>")
        print("   or: updater.py <pid> <app_dir> <exe_name> <download_url>")
        sys.exit(1)
    # -----------------------------------------------------------
    
    root = tk.Tk()
    app = UpdaterApp(root)
    root.mainloop()
