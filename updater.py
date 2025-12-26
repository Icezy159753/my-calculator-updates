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
        self.parent_pid = int(sys.argv[1])
        self.old_exe_path = sys.argv[2]
        self.new_exe_url = sys.argv[3]

        # สร้าง Widgets
        self.status_label = tk.Label(root, text="กำลังเตรียมอัปเดต...", font=Font(size=10))
        self.status_label.pack(pady=10)

        self.progress = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
        self.progress.pack(pady=5)
        
        self.percent_label = tk.Label(root, text="0%", font=Font(size=9))
        self.percent_label.pack()
        
        # เริ่มกระบวนการอัปเดตใน Thread ใหม่ เพื่อไม่ให้ GUI ค้าง
        self.update_thread = threading.Thread(target=self.run_update_process, daemon=True)
        self.update_thread.start()

    def disable_close(self):
        pass

    def run_update_process(self):
        try:
            # 1. รอให้โปรแกรมหลักปิดตัว
            self.status_label.config(text="กำลังรอให้โปรแกรมหลักปิดตัว...")
            self.root.update_idletasks()
            
            while psutil.pid_exists(self.parent_pid):
                time.sleep(0.5)

            # 2. ดาวน์โหลดไฟล์เวอร์ชันใหม่
            self.status_label.config(text=f"กำลังดาวน์โหลดเวอร์ชันใหม่...")
            new_exe_path = self.old_exe_path + ".new"
            
            with requests.get(self.new_exe_url, stream=True) as r:
                r.raise_for_status()
                total_size = int(r.headers.get('content-length', 0))
                downloaded_size = 0
                
                with open(new_exe_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
                        downloaded_size += len(chunk)
                        
                        progress_percent = (downloaded_size / total_size) * 100 if total_size > 0 else 0
                        self.progress['value'] = progress_percent
                        self.percent_label.config(text=f"{progress_percent:.0f}%")
                        self.root.update_idletasks()
            
            # 3. ลบไฟล์เก่าแล้วแทนที่ด้วยไฟล์ใหม่
            self.status_label.config(text="กำลังติดตั้งอัปเดต...")
            self.percent_label.config(text="")
            self.root.update_idletasks()
            time.sleep(1)

            os.remove(self.old_exe_path)
            os.rename(new_exe_path, self.old_exe_path)
            
            # 4. เปิดโปรแกรมเวอร์ชันใหม่ขึ้นมา
            os.startfile(self.old_exe_path)

            # 5. ปิดตัวเอง
            self.root.quit()

        except Exception as e:
            self.status_label.config(text=f"เกิดข้อผิดพลาด: {e}", fg="red")
            self.root.after(10000, self.root.quit)

if __name__ == "__main__":
    # --- เพิ่มเข้ามา: ตรวจสอบว่ามี arguments ส่งมาหรือไม่ก่อนรัน ---
    # ป้องกัน Error เวลาเผลอดับเบิ้ลคลิกไฟล์ .py โดยตรง
    if len(sys.argv) < 4:
        print("This script is intended to be run by the main application.")
        print("Usage: updater.py <pid> <exe_path> <download_url>")
        sys.exit(1)
    # -----------------------------------------------------------
    
    root = tk.Tk()
    app = UpdaterApp(root)
    root.mainloop()