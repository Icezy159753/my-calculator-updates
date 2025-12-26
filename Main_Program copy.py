import customtkinter as ctk
import os
import sys
from tkinter import messagebox
from PIL import Image
import importlib
from multiprocessing import Process, freeze_support
import subprocess
import requests
from packaging.version import parse as parse_version
# เพิ่มการ import ที่จำเป็นเหล่านี้ไว้ด้านบนสุดของไฟล์ Main_Program.py
import requests
import getpass
import threading
from datetime import datetime
import socket # เพิ่มเข้ามาเพื่อดึง IP Address (ถ้าต้องการ)

# --- เพิ่มค่าคงที่นี้ไว้ใกล้ๆกับค่าคงที่อื่นๆ ด้านบน ---
# !!! สำคัญ: ให้เปลี่ยน URL นี้เป็น URL ของ Web App ที่ได้จากการ Deploy บน Google Apps Script ของคุณ
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzLISY7ormRaB05x3qBD41apZ8zVMx2_-nNrlSz1RP26DCXXQgfpfESxS6ppgxkyOSm/exec" # <--- ใส่ URL ของคุณที่นี่


# --- ค่าคงที่สำหรับชื่อโฟลเดอร์ ---
PROGRAM_SUBFOLDER = "All_Programs"
ICON_FOLDER = "Icon"
# --- ข้อมูลโปรแกรมและ GitHub (สำคัญมาก: ต้องเปลี่ยนเป็นของคุณ) ---
CURRENT_VERSION = "1.0.53"
REPO_OWNER = "Icezy159753"  # << เปลี่ยนเป็นชื่อ Username ของคุณ
REPO_NAME = "my-calculator-updates"    # << เปลี่ยนเป็นชื่อ Repository ของคุณ


def get_executable_path():
    """หาตำแหน่งไฟล์ .exe ที่กำลังรันอยู่"""
    if getattr(sys, 'frozen', False):
        return sys.executable
    else:
        return os.path.abspath(__file__)

def check_for_updates(app_window):
    """ตรวจสอบอัปเดตและเรียกใช้ updater"""
    print("Checking for updates...")
    try:
        api_url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
        response = requests.get(api_url, timeout=5)
        response.raise_for_status()

        latest_release = response.json()
        latest_version = latest_release["tag_name"]

        if parse_version(latest_version) > parse_version(CURRENT_VERSION):
            print(f"New version found: {latest_version}")

            # ถามผู้ใช้ก่อนอัปเดต
            if messagebox.askyesno("Update Available", f"มีเวอร์ชันใหม่ ({latest_version})!\nต้องการอัปเดตตอนนี้เลยหรือไม่?"):
                # --- เพิ่มส่วนนี้เข้าไป ---
                # ดึงรายละเอียดการอัปเดตจาก "body" ของ release
                changelog_text = latest_release.get("body", "ไม่มีรายละเอียดการอัปเดต")

                # กำหนด path ของไฟล์ชั่วคราว
                changelog_path = os.path.join(os.path.dirname(get_executable_path()), 'changelog.tmp')

                # เขียน changelog ลงไฟล์ชั่วคราว
                try:
                    with open(changelog_path, "w", encoding="utf-8") as f:
                        f.write(changelog_text)
                    print(f"Changelog saved to {changelog_path}")
                except Exception as e:
                    print(f"Could not save changelog file: {e}")
                # -------------------------
                # หา URL ของไฟล์ exe ทั้งสองตัวจาก release ล่าสุด
                updater_url = None
                app_url = None
                for asset in latest_release['assets']:
                    if asset['name'] == 'updater.exe':
                        updater_url = asset['browser_download_url']
                    if asset['name'] == 'Main_Program.exe':
                        app_url = asset['browser_download_url']

                if not updater_url or not app_url:
                    messagebox.showerror("Error", "ไม่พบไฟล์สำหรับอัปเดตใน Release ล่าสุด")
                    return

                # ดาวน์โหลด updater.exe
                updater_path = os.path.join(os.path.dirname(get_executable_path()), 'updater.exe')
                print(f"Downloading updater from {updater_url} to {updater_path}")
                with requests.get(updater_url, stream=True) as r:
                    r.raise_for_status()
                    with open(updater_path, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=8192):
                            f.write(chunk)

                # เรียกใช้ updater.exe และส่ง argument ที่จำเป็นไปให้
                # แล้วปิดโปรแกรมหลักทันที
                subprocess.Popen([updater_path, str(os.getpid()), get_executable_path(), app_url])
                app_window.destroy() # หรือ sys.exit()

        else:
            print("You have the latest version.")

    except Exception as e:
        print(f"Could not check for updates: {e}")

# --- เพิ่มฟังก์ชันนี้เข้าไปใหม่ทั้งหมด ---
def create_custom_changelog_window(changelog_content):
    """
    สร้างและแสดงหน้าต่าง Changelog แบบกำหนดเองด้วย CustomTkinter
    """
    try:
        # สร้างหน้าต่างใหม่แบบ Toplevel (หน้าต่างย่อย)
        changelog_win = ctk.CTkToplevel()
        changelog_win.title(f"อัปเดตสำเร็จเป็นเวอร์ชัน {CURRENT_VERSION}!")

        # --- ตั้งค่าไอคอนของหน้าต่าง Popup ---
        try:
            icon_path = resource_path(os.path.join(ICON_FOLDER, "I_Main.ico")) # หรือไอคอนที่คุณต้องการ
            if os.path.exists(icon_path):
                changelog_win.iconbitmap(icon_path)
        except Exception as e:
            print(f"Could not set changelog window icon: {e}")

        # ทำให้หน้าต่างนี้อยู่ด้านหน้าเสมอ และผู้ใช้ต้องปิดหน้าต่างนี้ก่อน
        changelog_win.transient()
        changelog_win.grab_set()

        # สร้าง Widgets ภายในหน้าต่าง
        changelog_win.grid_columnconfigure(0, weight=1)
        changelog_win.grid_rowconfigure(1, weight=1)

        main_label = ctk.CTkLabel(changelog_win, text="มีอะไรใหม่ในเวอร์ชันนี้:", font=ctk.CTkFont(size=16, weight="bold"))
        main_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # ใช้ CTkTextbox เพื่อให้สามารถ scroll ข้อความยาวๆ ได้
        textbox = ctk.CTkTextbox(changelog_win, corner_radius=10, font=ctk.CTkFont(size=12))
        textbox.grid(row=1, column=0, sticky="nsew", padx=20, pady=5)
        textbox.insert("1.0", changelog_content)
        textbox.configure(state="disabled") # ทำให้แก้ไขข้อความไม่ได้

        ok_button = ctk.CTkButton(changelog_win, text="OK", width=100, command=changelog_win.destroy)
        ok_button.grid(row=2, column=0, padx=20, pady=(10, 20))

        # จัดหน้าต่างให้อยู่กลางจอ
        changelog_win.update_idletasks()
        win_width = 500
        win_height = 350
        x = (changelog_win.winfo_screenwidth() // 2) - (win_width // 2)
        y = (changelog_win.winfo_screenheight() // 2) - (win_height // 2)
        changelog_win.geometry(f'{win_width}x{win_height}+{x}+{y}')

        # รอจนกว่าหน้าต่างนี้จะถูกปิด
        changelog_win.wait_window()

    except Exception as e:
        print(f"Failed to create custom changelog window: {e}")
        # ถ้าสร้างหน้าต่าง custom ไม่ได้ ให้กลับไปใช้ messagebox แบบเดิม
        messagebox.showinfo(
            f"อัปเดตสำเร็จเป็นเวอร์ชัน {CURRENT_VERSION}!",
            f"มีอะไรใหม่ในเวอร์ชันนี้:\n\n{changelog_content}"
        )


def show_changelog_if_exists():
    """
    ตรวจสอบหาไฟล์ changelog ชั่วคราว ถ้าเจอให้แสดงเนื้อหาแล้วลบทิ้ง
    """
    try:
        # ใช้ get_executable_path() เพื่อให้ path ถูกต้องเสมอ
        base_dir = os.path.dirname(get_executable_path())
        changelog_path = os.path.join(base_dir, 'changelog.tmp')

        if os.path.exists(changelog_path):
            with open(changelog_path, "r", encoding="utf-8") as f:
                changelog_content = f.read()

            # ลบไฟล์ทิ้งทันทีหลังจากอ่านเสร็จ (สำคัญมาก!)
            os.remove(changelog_path)

            # แสดงผลใน Messagebox (หรือจะสร้างหน้าต่างใหม่สวยๆ ก็ได้)
            # ในฟังก์ชัน show_changelog_if_exists
            if changelog_content.strip(): # เช็กว่ามีเนื้อหาจริงๆ
                # --- เรียกใช้ฟังก์ชันสร้างหน้าต่างใหม่ของเราแทน ---
                create_custom_changelog_window(changelog_content)
    except Exception as e:
        print(f"Could not process changelog file: {e}")



# --- ฟังก์ชัน Resource Path ---
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


#ส่วนแก้ไข Theme สีต่างๆ
# กำหนด Path ไปยัง Theme ใหม่ของคุณ
custom_theme_file = resource_path("themes/blue.json") # หรือ Path แบบเต็มถ้าไม่ได้ใช้ resource_path
# --- กำหนดค่าเริ่มต้นของ CustomTkinter ---

if os.path.exists(custom_theme_file):
    ctk.set_default_color_theme(custom_theme_file)
    print(f"LAUNCHER_INFO: Loaded custom theme from: {custom_theme_file}")
else:
    ctk.set_default_color_theme("blue") # Fallback ไปยัง Theme ที่มีอยู่แล้วถ้าหาไฟล์ไม่เจอ
    print(f"LAUNCHER_WARNING: Custom theme file not found at '{custom_theme_file}'. Using default 'blue' theme.")

ctk.set_appearance_mode("Dark")  # "System", "Light", "Dark"
#ctk.set_default_color_theme("Red")  # "blue", "green", "dark-blue",Red"

# --- กำหนดรายการโปรแกรม (เพิ่ม "category" และปรับ "module_path") ---
PROGRAMS = [
    {
        "id": "spss_log",
        "name": "สร้าง Itemdef จากSPSS V3",
        "description": "โปรแกรมแปลง SPSS เป็น Excel",
        "type": "local_py_module",
        "module_path": "Program_ItemdefSPSS_Log", # ชื่อไฟล์ .py (ไม่รวม .py)
        "entry_point": "run_this_app",
        "icon": "Peth.ico",
        "category": "Lychee", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "t2b_itemdef",
        "name": "ทำ TB/T2B จากไฟล์ Itemdef V3",
        "description": "โปรแกรมทำ TB/T2B จาก Itemdef",
        "type": "local_py_module",
        "module_path": "Program_T2B_Itermdef", # ชื่อไฟล์ .py
        "entry_point": "run_this_app",
        "icon": "T2B.ico",
        "category": "Lychee", # <--- เพิ่ม category
        "enabled": True
    },
    #{
        #"id": "Program_Get Value",
        #"name": "Program_Get Value V6",
        #"description": "โปรแกรมดึง Var+Value จาก SPSS",
        #"type": "local_py_module",
        #"module_path": "Program_Var_Value_SPSS_V6",
        #"entry_point": "run_this_app",
        #"icon": "Iconcat.ico",
        #"category": "SPSS", # <--- เพิ่ม category
        #"enabled": True
    #},

    {
        "id": "โปรแกรมสร้าง Promt แปะ Eng v1",
        "name": "โปรแกรม GetValue+Promt แปะ Eng",
        "description": "เอาไว้ GetValue+Copy Promt แปะ Eng",
        "type": "local_py_module",
        "module_path": "117_Newen_Promt", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "promt.ico",
        "category": "Lychee", # <--- เพิ่ม category
        "enabled": True
    },

    {
        "id": "Cleandata SPSS",
        "name": "CleanData+Frequenzy SPSS V1",
        "description": "โปรแกรม Clean Data SPSS + Frequncy",
        "type": "local_py_module",
        "module_path": "99_CleanSPSS_Germini", # ชื่อไฟล์ .py
        "entry_point": "run_this_app",
        "icon": "Clean.ico",
        "category": "SPSS", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรมตัดชุด",
        "name": "โปรแกรมตัดชุด ตามใบบรีฟ V1",
        "description": "โปรแกรมตัดชุด ตามใบบรีฟจากไฟล์ Excel",
        "type": "local_py_module",
        "module_path": "99_Excel", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "Cut.ico",
        "category": "Excel", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรม RenameSheet",
        "name": "โปรแกรม RenameSheet V1",
        "description": "โปรแกรมช่วย Rename Sheet ใน Excel",
        "type": "local_py_module",
        "module_path": "Rename Sheet", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "Rename.ico",
        "category": "Excel", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรม เก็บ Norm",
        "name": "เก็บ Norm V1",
        "description": "โปรแกรมช่วย เก็บ Norm",
        "type": "local_py_module",
        "module_path": "Norm_2025", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "Norm.ico",
        "category": "Key Norm", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรม Logic_Generator Itemdef",
        "name": "โปรแกรม Logic_Generator Itemdef V8",
        "description": "โปรแกรม GenSyntaxClean Lychee Itemdef",
        "type": "local_py_module",
        "module_path": "Logic_Generator", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "Logic.ico",
        "category": "Lychee", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรม Get SPSS",
        "name": "โปรแกรม Get SPSS V2",
        "description": "โปรแกรม GenSyntax Get SPSS",
        "type": "local_py_module",
        "module_path": "105_GetSPSS", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "Get.ico",
        "category": "SPSS", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรม แปลงCE Otherจาก Edit V2",
        "name": "โปรแกรม แปลงCE Otherจาก Edit V2",
        "description": "แปลงไฟล์ CE Other จาก Edit เป็นไฟล์ Excel",
        "type": "local_py_module",
        "module_path": "106_Map_spss_Excel", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "CE_Edit.ico",
        "category": "Lychee", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรม Move Sheet Excel V1",
        "name": "โปรแกรม Move Sheet Excel V1",
        "description": "เอาไว้ Move Sheet Excel",
        "type": "local_py_module",
        "module_path": "107_Movesheet", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "movesheet.ico",
        "category": "Excel", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรมCheck Rotation Diary V1",
        "name": "โปรแกรมCheck Rotation Diary V1",
        "description": "เอาไว้เช็ค Rotation Diary + จำนวน RD",
        "type": "local_py_module",
        "module_path": "109_Diary", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "diary.ico",
        "category": "Diary", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรมซ่อมไฟล์SPSS V1",
        "name": "โปรแกรมซ่อมไฟล์SPSS V1",
        "description": "เอาไว้แปลงไฟล์ SPSS ที่มีปัญหาเป็น UTF-8",
        "type": "local_py_module",
        "module_path": "convert_SPSS_UTF8", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "convert.ico",
        "category": "SPSS", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรมแปลงไฟล์ SPSS To Excel V1",
        "name": "โปรแกรมแปลงไฟล์ SPSS To Excel V1",
        "description": "เอาไว้แปลงไฟล์ SPSS เป็น Excelแบบเลือกได้",
        "type": "local_py_module",
        "module_path": "ConvertSPSS_Excel", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "SPSS_Excel.ico",
        "category": "SPSS", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรมลบ Sig จาก TableLychee V1",
        "name": "โปรแกรมลบ Sig จาก TableLychee V1",
        "description": "เอาไว้ลบ Sig จาก Table Lychee ตามที่ระบุ",
        "type": "local_py_module",
        "module_path": "Del_Sig", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "Del_Sig.ico",
        "category": "Lychee", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรม Check Codes_Other V1",
        "name": "โปรแกรม Check Codes_Other V1",
        "description": "เอาไว้ Check Codes_Other ว่าเกิดที่ข้อไหนบ้างและเช็คกับไฟล์ Edit ",
        "type": "local_py_module",
        "module_path": "CheckOther", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "Check Other.ico",
        "category": "Lychee", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรมแตก CodeNA จากไฟล์ SPSS&Excel V1",
        "name": "โปรแกรมแตก CodeNA จากไฟล์ SPSS&Excel V1",
        "description": "เอาไว้ แตก CodeNA จากไฟล์ SPSS และ Excel เพื่อเอาเข้า Lychee",
        "type": "local_py_module",
        "module_path": "113_ProgramCodeNA", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "NA Code.ico",
        "category": "Lychee", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรมรัน Correlation จาก SPSS Data V1",
        "name": "โปรแกรมรัน Correlation จาก SPSS Data V1",
        "description": "โปรแกรม รัน Correlationและต่อ Data",
        "type": "local_py_module",
        "module_path": "104_Correlation", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "Cor.ico",
        "category": "Statistic", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรมลบN=0 OEในLychee V1",
        "name": "โปรแกรมลบN=0 OEในLychee V1",
        "description": "เอาไว้ลบช่องว่างใน TableOE ของ Lychee กรณีรันCE+OE",
        "type": "local_py_module",
        "module_path": "114_DelblankLychee", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "delete_table.ico",
        "category": "Lychee", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "โปรแกรมCreate Format_Kao By_DP V3",
        "name": "โปรแกรมCreate Format_Kao By_DP V3",
        "description": "เอาไว้สร้าง Format_Kao จากLychee แบบไม่ต้องสร้างข้อดูดติดใน Lychee",
        "type": "local_py_module",
        "module_path": "119_Create_Format_Kao", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "Kao2.ico",
        "category": "Lychee", # <--- เพิ่ม category
        "enabled": True
    },  
    {
        "id": "โปรแกรมรัน BPI Brand Power Index V1",
        "name": "โปรแกรมรัน BPI Brand Power Index V1",
        "description": "เอาไว้รัน BPI Brand Power Index จากไฟล์ SPSS",
        "type": "local_py_module",
        "module_path": "120_bpi", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "BPI.ico",
        "category": "Statistic", # <--- เพิ่ม category
        "enabled": True
    },  
    {
        "id": "โปรแกรม Multidimensional Scaling (MDS) V12",
        "name": "Multidimensional Scaling (MDS) V12",
        "description": "เอาไว้รัน Multidimensional Scaling (MDS) จากไฟล์ Excel",
        "type": "local_py_module",
        "module_path": "MDS", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "MDS.ico",
        "category": "Statistic", # <--- เพิ่ม category
        "enabled": True
    },     
    {
        "id": "โปรแกรม ดูดติด MA _O จาก Togo V1",
        "name": "ปรแกรม ดูดติด MA _O จาก Togo V1",
        "description": "เอาไว้ดูดติด _O ทุกข้อให้อยู่ช่องเดียวกัน(เลือกตัวคั่นได้)",
        "type": "local_py_module",
        "module_path": "121_Merge_MA_V2", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "Merge.ico",
        "category": "Statistic", # <--- เพิ่ม category
        "enabled": True
    },      
    {
        "id": "MRSET Auto-Generator v5.0",
        "name": "MRSET Auto-Generator v5.0",
        "description": "เอาไว้สร้างโค้ด MRSET อัตโนมัติจากไฟล์ SPSS",
        "type": "local_py_module",
        "module_path": "121_SPSS_MRSET", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "MRSET.ico",
        "category": "SPSS", # <--- เพิ่ม category
        "enabled": True
    },
    {
        "id": "PSM Pricezen V1",
        "name": "PSM Pricezen V1",
        "description": "เอาไว้รัน PSM Pricezen จากไฟล์ Excel",
        "type": "local_py_module",
        "module_path": "122_PSM", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "PSM2.ico",
        "category": "Statistic", # <--- เพิ่ม category
        "enabled": True
    },      
    {
        "id": "Program BrandSence V1",
        "name": "Program BrandSence V1",
        "description": "เอาไว้รัน Program BrandSence",
        "type": "local_py_module",
        "module_path": "123_Program_Run_Brandsence_Add_C All", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "BrandS.ico",
        "category": "Statistic", # <--- เพิ่ม category
        "enabled": True
    },   
    {
        "id": "Program Gen Table Reporter V1",
        "name": "Program Gen Table Reporter V1",
        "description": "เอาไว้ช่วยสร้าง Table Reporter",
        "type": "local_py_module",
        "module_path": "124_Table_Reporter", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "Reporter.ico",
        "category": "อื่นๆ", # <--- เพิ่ม category
        "enabled": True
    },   
    {
        "id": "Program Convert JPG to SVG V1",
        "name": "Program Convert JPG to SVG V1",
        "description": "เอาไว้แปลง Convert JPG to SVG",
        "type": "local_py_module",
        "module_path": "125_Convert_Icon_Svg2", # <--- ปรับชื่อ module_path
        "entry_point": "run_this_app",
        "icon": "SVG.ico",
        "category": "อื่นๆ", # <--- เพิ่ม category
        "enabled": True
    },   
]

# --- ขนาดไอคอน ---
ICON_SIZE = (60, 60) # ปรับขนาดไอคอนในการ์ด
CARD_DESCRIPTION_WRAPLENGTH = 150 # ปรับความกว้างของคำอธิบายในการ์ด

# --- ฟังก์ชันสำหรับรันโมดูลย่อยในโปรเซสใหม่ (เหมือนเดิม) ---
def run_module_entrypoint(module_name_in_subfolder, entry_point_func_name="main", args=(), script_kwargs={}):
    try:
        # Ensure the subfolder is part of the module name if not already
        if not module_name_in_subfolder.startswith(PROGRAM_SUBFOLDER + "."):
            full_module_name = f"{PROGRAM_SUBFOLDER}.{module_name_in_subfolder}"
        else:
            full_module_name = module_name_in_subfolder

        print(f"LAUNCHER_INFO: Importing module: {full_module_name}")
        module = importlib.import_module(full_module_name)
        print(f"LAUNCHER_INFO: Module {full_module_name} imported successfully.")

        if hasattr(module, entry_point_func_name):
            entry_point_function = getattr(module, entry_point_func_name)
            print(f"LAUNCHER_INFO: Found entry point: {entry_point_func_name}. Running with script_kwargs: {script_kwargs}")
            entry_point_function(*args, **script_kwargs)
            print(f"LAUNCHER_INFO: Finished running {entry_point_func_name} in {full_module_name}")
        else:
            print(f"LAUNCHER_ERROR: Entry point function '{entry_point_func_name}' not found in module '{full_module_name}'.")
            messagebox.showerror("Launch Error", f"ไม่พบฟังก์ชันหลัก '{entry_point_func_name}'\nในโมดูล '{module_name_in_subfolder}'.")

    except ImportError as e:
        print(f"LAUNCHER_ERROR: Error importing module {full_module_name}: {e}")
        messagebox.showerror("Launch Error", f"ไม่สามารถโหลดโมดูล '{module_name_in_subfolder}' ได้:\n{e}\n\nตรวจสอบว่าไฟล์ .py อยู่ในโฟลเดอร์ '{PROGRAM_SUBFOLDER}' และ sys.path ถูกต้อง")
    except Exception as e:
        print(f"LAUNCHER_ERROR: Error running module {full_module_name}: {e}")
        messagebox.showerror("Runtime Error", f"เกิดข้อผิดพลาดขณะรัน '{module_name_in_subfolder}':\n{e}")


class AppLauncher(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Program All DP v{CURRENT_VERSION}")
        # --- ส่วนที่ปรับแก้เพื่อจัดหน้าต่างให้อยู่กลางจอ ---
        window_width = 1000
        window_height = 700

        # ดึงขนาดของหน้าจอคอมพิวเตอร์
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # คำนวณหาตำแหน่งกึ่งกลาง (x, y)
        center_x = int((screen_width / 2) - (window_width / 2))
        center_y = int((screen_height / 2) - (window_height / 2))

        # ตั้งค่าขนาดและตำแหน่งของหน้าต่าง
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        # ----------------------------------------------------

        self.launcher_base_dir = resource_path('.')
        self.icon_dir = os.path.join(self.launcher_base_dir, ICON_FOLDER)
        self.program_dir = os.path.join(self.launcher_base_dir, PROGRAM_SUBFOLDER)

        if self.launcher_base_dir not in sys.path:
            sys.path.insert(0, self.launcher_base_dir)
            print(f"LAUNCHER_INFO: Added to sys.path: {self.launcher_base_dir}")
        # print(f"LAUNCHER_DEBUG: Current sys.path: {sys.path}")

        self.icon_cache = {}
        self.category_buttons = {} # เก็บปุ่ม category เพื่อเปลี่ยนสไตล์

        # --- สร้างโครงสร้าง UI หลัก ---
        self.grid_columnconfigure(1, weight=1) # ให้ content_frame ขยาย
        self.grid_rowconfigure(0, weight=1)    # ให้ content_frame ขยาย

        # Sidebar Frame
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsw", rowspan=2)
        # Configure row 0 to take up extra space if needed, or last button row for "Exit"
        # self.sidebar_frame.grid_rowconfigure(5, weight=1) # Example: if 5th row is last before a spacer

        # Content Frame (Scrollable) - เพิ่ม scrollbar_button_color และปรับความเร็วการเลื่อน
        self.content_frame = ctk.CTkScrollableFrame(
            self,
            label_text="All Programs",
            label_font=ctk.CTkFont(size=16, weight="bold"),
            scrollbar_button_color=("gray70", "gray30"),
            scrollbar_button_hover_color=("gray60", "gray40")
        )
        self.content_frame.grid(row=0, column=1, padx=15, pady=15, sticky="nsew")

        # ปรับความเร็วการ scroll ให้ลื่นขึ้น
        self.content_frame._parent_canvas.configure(yscrollincrement=5)

        # เพิ่มการ bind mousewheel เพื่อให้ scroll ลื่นขึ้น
        self.bind_smooth_scroll()

        self.create_sidebar_menu()
        self.show_category_programs("All") # แสดงโปรแกรมทั้งหมดเมื่อเริ่มต้น

    def bind_smooth_scroll(self):
        """เพิ่มการ scroll ที่ลื่นไหลกว่าเดิม"""
        def on_mousewheel(event):
            # ปรับความเร็วการ scroll ให้ลื่นขึ้นด้วยการลดขนาด delta
            canvas = self.content_frame._parent_canvas
            canvas.yview_scroll(int(-1 * (event.delta / 60)), "units")
            return "break"  # ป้องกันไม่ให้ scroll ซ้ำ

        # Bind สำหรับ Windows/Mac
        self.content_frame._parent_canvas.bind_all("<MouseWheel>", on_mousewheel)
        # Bind สำหรับ Linux
        self.content_frame._parent_canvas.bind_all("<Button-4>", lambda e: self.content_frame._parent_canvas.yview_scroll(-1, "units"))
        self.content_frame._parent_canvas.bind_all("<Button-5>", lambda e: self.content_frame._parent_canvas.yview_scroll(1, "units"))

    # เพิ่ม 2 ฟังก์ชันนี้เข้าไปในคลาส AppLauncher ของคุณ
    def fetch_update_history_text(self):
        """
        ดึงข้อมูลทุก Release จาก GitHub API แล้วจัดรูปแบบเป็นข้อความ
        """
        history_log = []
        try:
            # เปลี่ยนจาก /latest เป็น /releases เพื่อดึงทุกเวอร์ชัน
            api_url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases"
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()

            releases = response.json()
            if not releases:
                return "ไม่พบประวัติการอัปเดต"

            for release in releases:
                version = release.get("tag_name", "N/A")
                date = release.get("published_at", "").split("T")[0] # เอาเฉพาะวันที่
                body = release.get("body", "ไม่มีรายละเอียด")

                # จัดรูปแบบของแต่ละเวอร์ชัน
                log_entry = (
                    f"--- เวอร์ชัน {version} ({date}) ---\n"
                    f"{body}\n"
                    "----------------------------------\n"
                )
                history_log.append(log_entry)

            return "\n".join(history_log)

        except requests.exceptions.RequestException as e:
            print(f"Error fetching update history: {e}")
            return f"ไม่สามารถดึงข้อมูลได้: {e}"
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            return f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}"

    def show_update_log_window(self):
        """
        สร้างและแสดงหน้าต่างประวัติการอัปเดต
        """
        log_win = ctk.CTkToplevel(self)
        log_win.title("ประวัติการอัปเดต")
        log_win.geometry("600x500")

        # ตั้งค่าไอคอน (ถ้าต้องการ)
        try:
            icon_path = resource_path(os.path.join(ICON_FOLDER, "I_Main.ico"))
            if os.path.exists(icon_path):
                log_win.iconbitmap(icon_path)
        except Exception as e:
            print(f"Could not set log window icon: {e}")

        # ทำให้หน้าต่างนี้อยู่ด้านหน้าเสมอ
        log_win.transient(self)
        log_win.grab_set()

        # สร้าง Textbox สำหรับแสดงผล
        textbox = ctk.CTkTextbox(log_win, font=ctk.CTkFont(family="tahoma", size=12))
        textbox.pack(expand=True, fill="both", padx=15, pady=15)

        # แสดงข้อความว่ากำลังโหลด
        textbox.insert("1.0", "กำลังดึงข้อมูลประวัติการอัปเดต กรุณารอสักครู่...")
        textbox.configure(state="disabled")
        log_win.update()

        # ดึงข้อมูลประวัติแล้วแสดงผล
        history_text = self.fetch_update_history_text()
        textbox.configure(state="normal") # เปิดให้แก้ไขได้ชั่วคราว
        textbox.delete("1.0", "end")
        textbox.insert("1.0", history_text)
        textbox.configure(state="disabled") # ปิดการแก้ไข



    # 1. เพิ่มฟังก์ชันใหม่นี้เข้าไปในคลาส AppLauncher
    def open_web_script_link(self):
        """
        เปิดลิงก์ไปยังเว็บสคริปต์ในเบราว์เซอร์เริ่มต้นของผู้ใช้
        """
        import webbrowser
        from tkinter import messagebox

        url = "https://script.google.com/macros/s/AKfycbxS0tIe8TnDUf-QNn6Y0NdlT-MbQY-FmyM2_uR03muhIcg05z_0F9mFHt9ReNGpns2H/exec"
        try:
            print(f"LAUNCHER_INFO: Opening URL: {url}")
            webbrowser.open_new_tab(url)
        except Exception as e:
            print(f"LAUNCHER_ERROR: Could not open URL: {e}")
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดลิงก์ได้:\n{e}")

    def _on_press_yt_button(self, event):
        """เมื่อกดปุ่มค้างไว้ ให้ปุ่มยุบลงเล็กน้อย"""
        self.web_script_button.grid_configure(pady=(8, 2))

    def _on_release_yt_button(self, event):
        """เมื่อปล่อยปุ่ม ให้ปุ่มกลับที่เดิมและเรียกใช้คำสั่ง"""
        self.web_script_button.grid_configure(pady=(5, 5))
        self.open_web_script_link()


    # 2. แก้ไขฟังก์ชัน create_sidebar_menu ทั้งหมดด้วยโค้ดนี้
    def create_sidebar_menu(self):
        logo_label = ctk.CTkLabel(self.sidebar_frame, text="หมวดหมู่", font=ctk.CTkFont(size=20, weight="bold"))
        logo_label.grid(row=0, column=0, padx=20, pady=(20, 15))

        # ดึง categories ที่ไม่ซ้ำกันจาก PROGRAMS และเพิ่ม "All"
        raw_categories = set(p["category"] for p in PROGRAMS if p.get("category") and p.get("enabled", True))
        categories = ["All"] + sorted(list(raw_categories))

        for i, category_name in enumerate(categories):
            button = ctk.CTkButton(self.sidebar_frame, text=category_name,
                                     command=lambda cat=category_name: self.show_category_programs(cat),
                                     anchor="w", font=ctk.CTkFont(size=13))
            button.grid(row=i + 1, column=0, padx=15, pady=4, sticky="ew")
            self.category_buttons[category_name] = button

        # Appearance Mode Toggler
        appearance_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w", font=ctk.CTkFont(size=12))
        appearance_label.grid(row=len(categories) + 1, column=0, padx=15, pady=(15,0), sticky="ew")
        self.appearance_mode_menu = ctk.CTkOptionMenu(self.sidebar_frame,
                                                       values=["Light", "Dark", "System"],
                                                       command=self.change_appearance_mode_event,
                                                       font=ctk.CTkFont(size=12))
        self.appearance_mode_menu.grid(row=len(categories) + 2, column=0, padx=15, pady=(0,10), sticky="ew")
        self.appearance_mode_menu.set(ctk.get_appearance_mode())


        # กำหนดให้แถวว่างนี้ยืดได้ เพื่อดันทุกอย่างลงไปด้านล่าง
        self.sidebar_frame.grid_rowconfigure(len(categories) + 3, weight=1)

        # --- ส่วนที่ปรับเพิ่มและแก้ไข ---
        web_script_icon = self.load_icon("yt.ico")
        self.web_script_button = ctk.CTkButton(self.sidebar_frame,
                                               text="",
                                               image=web_script_icon,
                                               fg_color="transparent",
                                               hover_color="#4A4A4A",
                                               width=40,
                                               command=None)
        # ปรับ pady ลดช่องว่างด้านล่างของปุ่ม
        self.web_script_button.grid(row=len(categories) + 4, column=0, padx=15, pady=(5, 0), sticky="s")

        self.web_script_button.bind("<Button-1>", self._on_press_yt_button)
        self.web_script_button.bind("<ButtonRelease-1>", self._on_release_yt_button)

        # ปรับข้อความและ pady ของป้ายข้อความ
        yt_label = ctk.CTkLabel(self.sidebar_frame, text="คลิกเพื่อดู VDO การใช้งาน",
                                font=ctk.CTkFont(size=12), text_color="gray60")
        yt_label.grid(row=len(categories) + 5, column=0, padx=15, pady=(2, 10), sticky="n")
        # --- จบส่วนที่ปรับเพิ่มและแก้ไข ---

        # ปุ่มสำหรับดูประวัติการอัปเดต
        log_button = ctk.CTkButton(self.sidebar_frame, text="ประวัติการอัปเดต",
                                 command=self.show_update_log_window,
                                 font=ctk.CTkFont(size=13))
        log_button.grid(row=len(categories) + 6, column=0, padx=15, pady=(5, 5), sticky="sew")

        # Label แสดงเวอร์ชัน
        version_text = f"เวอร์ชัน {CURRENT_VERSION}"
        version_label = ctk.CTkLabel(self.sidebar_frame, text=version_text,
                                     font=ctk.CTkFont(size=11),
                                     text_color="gray60")
        version_label.grid(row=len(categories) + 7, column=0, padx=15, pady=(5, 5), sticky="sew")

        # ปุ่ม Exit
        exit_button = ctk.CTkButton(self.sidebar_frame, text="Exit Launcher",
                                  command=self.quit_app,
                                  fg_color="transparent", border_width=1,
                                  text_color=("gray10", "#DCE4EE"),
                                  font=ctk.CTkFont(size=13))
        exit_button.grid(row=len(categories) + 8, column=0, padx=15, pady=10, sticky="sew")

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

    def quit_app(self):
        self.quit()
        self.destroy()


    def show_category_programs(self, category_name):
        self.content_frame.configure(label_text=f"โปรแกรม: {category_name}")

        for widget in self.content_frame.winfo_children():
            widget.destroy()

        if category_name == "All":
            programs_to_display = [p for p in PROGRAMS if p.get("enabled", True)]
        else:
            programs_to_display = [p for p in PROGRAMS if p.get("enabled", True) and p.get("category") == category_name]

        if not programs_to_display:
            no_program_label = ctk.CTkLabel(self.content_frame, text=f"ไม่พบโปรแกรมในหมวดหมู่ '{category_name}'",
                                            font=ctk.CTkFont(size=16))
            no_program_label.pack(pady=50, padx=20, anchor="center", expand=True)
            return

        self.create_program_widgets(programs_to_display)

        # Highlight active category button
        active_button_color = ctk.ThemeManager.theme["CTkButton"]["hover_color"]
        default_button_color = ctk.ThemeManager.theme["CTkButton"]["fg_color"]

        for cat, btn in self.category_buttons.items():
            if cat == category_name:
                btn.configure(fg_color=active_button_color)
            else:
                # Ensure fg_color is reset correctly based on theme
                # For buttons not matching the "Exit" style, use default fg_color
                if btn.cget("text") != "Exit Launcher": # Example check if you have special buttons
                    btn.configure(fg_color=default_button_color)


    def get_icon_path(self, icon_name):
        if not icon_name: return None
        icon_path = os.path.join(self.icon_dir, icon_name)
        if os.path.exists(icon_path):
            return icon_path
        else:
            print(f"LAUNCHER_WARNING: Icon file not found: {icon_path}")
            return None

    def load_icon(self, icon_name):
        if not icon_name: return None
        if icon_name in self.icon_cache: return self.icon_cache[icon_name]
        icon_path = self.get_icon_path(icon_name)
        if icon_path:
            try:
                image = Image.open(icon_path)
                if image.mode != 'RGBA':
                    image = image.convert("RGBA")
                ctk_image = ctk.CTkImage(light_image=image, dark_image=image, size=ICON_SIZE)
                self.icon_cache[icon_name] = ctk_image
                return ctk_image
            except Exception as e:
                print(f"LAUNCHER_ERROR: Loading icon '{icon_name}' from path '{icon_path}': {e}")
        return None

    def create_program_widgets(self, programs_list):
        row_num = 0; col_num = 0
        max_cols = 3 # จำนวนการ์ดสูงสุดต่อแถว (ปรับได้ตามต้องการ)

        #ปรับสีการด ตรงส่วนที่โชวโปรแกรม

        card_fg_color_light = "#DCDCDC" # สีพื้นหลังการ์ดที่สว่างกว่า content_frame เล็กน้อย
        card_fg_color_dark = "#2b2b2b"  # สีพื้นหลังการ์ดสำหรับ Dark mode
        card_border_color_light = "#DCDCDC"
        card_border_color_dark = "#2b2b2b"
        name_text_color_light = "#202020"
        name_text_color_dark = "#FFFFFF"
        desc_text_color_light = "#F30101"
        desc_text_color_dark = "#FFFFFF"
        # เราจะใช้ tuple (light_color, dark_color) ให้ CTk จัดการ
        card_fg_color = (card_fg_color_light, card_fg_color_dark)
        card_border_color = (card_border_color_light, card_border_color_dark)
        name_text_color = (name_text_color_light, name_text_color_dark)
        desc_text_color = (desc_text_color_light, desc_text_color_dark)

        for i, program in enumerate(programs_list):
            card = ctk.CTkFrame(self.content_frame,
                                corner_radius=10, # เพิ่มมุมมนให้สวยงามขึ้น
                                border_width=1,   # เพิ่มเส้นขอบ
                                fg_color=card_fg_color, # <--- กำหนดสีพื้นหลังการ์ด
                                border_color=card_border_color) # <--- กำหนดสีเส้นขอบ
            card.grid(row=row_num, column=col_num, padx=10, pady=10, sticky="nsew") # เพิ่ม padx/pady ให้ห่างกันเล็กน้อย

            card.grid_rowconfigure(0, weight=0) # Icon
            card.grid_rowconfigure(1, weight=0) # Name
            card.grid_rowconfigure(2, weight=1) # Description (ให้ยืดหยุ่น)
            card.grid_rowconfigure(3, weight=0) # Button
            card.grid_columnconfigure(0, weight=1)

            icon_image = self.load_icon(program.get("icon"))
            icon_label = ctk.CTkLabel(card, text="", image=icon_image, fg_color="transparent") # ให้พื้นหลัง icon โปร่งใส
            icon_label.grid(row=0, column=0, pady=(15, 5)) # เพิ่ม pady ด้านบนของ icon

            name_label = ctk.CTkLabel(card, text=program["name"],
                                      font=ctk.CTkFont(size=13, weight="bold"), # เพิ่มขนาดตัวอักษรชื่อ
                                      text_color=name_text_color, # <--- กำหนดสีตัวอักษรชื่อ
                                      fg_color="transparent")
            name_label.grid(row=1, column=0, pady=(0,5), padx=10, sticky="ew")

            desc_label = ctk.CTkLabel(card, text=program["description"],
                                      wraplength=CARD_DESCRIPTION_WRAPLENGTH,
                                      font=ctk.CTkFont(size=10),
                                      justify="left", anchor="nw",
                                      height=50, # ลด height ลงเล็กน้อยถ้า ICON_SIZE ใหญ่ขึ้น
                                      text_color=desc_text_color, # <--- กำหนดสีตัวอักษรคำอธิบาย
                                      fg_color="transparent")
            desc_label.grid(row=2, column=0, pady=(0,10), padx=10, sticky="new") # เพิ่ม pady ด้านล่างคำอธิบาย

            launch_button = ctk.CTkButton(card, text="เปิดโปรแกรม",
                                          font=ctk.CTkFont(size=12, weight="bold"), # ปรับ font ปุ่ม
                                          # fg_color=button_fg_color,      # <--- กำหนดสีปุ่ม (ถ้าต้องการ custom)
                                          # hover_color=button_hover_color, # <---
                                          # text_color=button_text_color,   # <---
                                          command=lambda p=program: self.launch_program(p))
            launch_button.grid(row=3, column=0, pady=(5, 15), padx=10, sticky="ew") # เพิ่ม pady ปุ่ม

            col_num += 1
            if col_num >= max_cols:
                col_num = 0
                row_num += 1

        for i in range(max_cols):
            self.content_frame.grid_columnconfigure(i, weight=1)
        if programs_list:
            self.content_frame.grid_rowconfigure(row_num + 1 , weight=1)

    # --- เพิ่มฟังก์ชันใหม่นี้เข้าไปในคลาส AppLauncher ---
    def log_session_to_sheet(self, program_name, user_info, start_time, end_time, duration_formatted):
        """
        ส่งข้อมูลเซสชันการใช้งาน (เวลาเริ่ม-จบ, ระยะเวลา) ไปยัง Google Sheet
        (เวอร์ชันนี้ส่งระยะเวลาเป็นรูปแบบ HH:MM:SS)
        """
        if "YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE" in GOOGLE_SCRIPT_URL:
            print("LOGGING_WARNING: กรุณาเปลี่ยน GOOGLE_SCRIPT_URL เป็น URL ของ Web App จริง")
            return

        try:
            # เตรียมข้อมูล (payload) ที่จะส่ง
            payload = {
                'startDate': start_time.strftime("%Y-%m-%d"),
                'startTime': start_time.strftime("%H:%M:%S"),
                'endTime': end_time.strftime("%H:%M:%S"),
                'duration': duration_formatted, # <--- ส่งค่าที่จัดรูปแบบแล้ว
                'programName': program_name,
                'user': user_info
            }

            print(f"LOGGING: กำลังส่งข้อมูลเซสชัน: {payload}")
            response = requests.post(GOOGLE_SCRIPT_URL, data=payload, timeout=15)
            response.raise_for_status()
            print(f"LOGGING: บันทึกข้อมูลเซสชันสำเร็จ. Response: {response.text}")

        except requests.exceptions.RequestException as e:
            print(f"LOGGING_ERROR: ไม่สามารถเชื่อมต่อเพื่อบันทึกข้อมูลเซสชันได้: {e}")
        except Exception as e:
            print(f"LOGGING_ERROR: เกิดข้อผิดพลาดที่ไม่คาดคิดระหว่างการบันทึกเซสชัน: {e}")

    def _wait_and_log_session(self, process_to_watch, program_info):
        """
        ฟังก์ชันนี้จะทำงานใน Thread แยก, รอจนกว่า process จะจบ,
        แล้วคำนวณเวลาก่อนเรียกฟังก์ชันเพื่อส่งข้อมูล
        """
        program_name = program_info.get("name", "Unknown Program")
        print(f"MONITOR: เริ่มเฝ้าดูโปรแกรม '{program_name}' (PID: {process_to_watch.pid})")
        
        start_time = datetime.now()
        
        try:
            username = getpass.getuser()
            hostname = socket.gethostname()
            ip_address = socket.gethostbyname(hostname)
            user_info = f"{username} ({ip_address})"
        except Exception:
            user_info = f"{getpass.getuser()} (IP N/A)"

        process_to_watch.join()

        end_time = datetime.now()
        print(f"MONITOR: โปรแกรม '{program_name}' (PID: {process_to_watch.pid}) ถูกปิดแล้ว")

        # --- ส่วนที่แก้ไข: คำนวณระยะเวลาเป็นรูปแบบ HH:MM:SS ---
        duration_seconds = int((end_time - start_time).total_seconds())
        hours, remainder = divmod(duration_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        duration_formatted = f"{hours:02}:{minutes:02}:{seconds:02}"
        # ----------------------------------------------------

        # เรียกฟังก์ชันเพื่อส่งข้อมูลทั้งหมดในครั้งเดียว
        self.log_session_to_sheet(program_name, user_info, start_time, end_time, duration_formatted)


    # --- แก้ไขฟังก์ชัน launch_program ทั้งหมด ให้เป็นไปตามนี้ ---
    def launch_program(self, program_info):
        """
        ฟังก์ชันนี้ถูกแก้ไขเพื่อสร้าง Thread แยกสำหรับเฝ้าดูแต่ละโปรแกรมที่เปิด
        """
        program_name = program_info.get("name", "Unknown Program")
        program_type = program_info.get("type", "unknown")
        process = None # ประกาศตัวแปร process ไว้ก่อน

        print(f"LAUNCHER_INFO: กำลังเตรียมเปิด '{program_name}' (Type: {program_type})")

        if program_type == "local_py_module":
            module_path = program_info.get("module_path")
            entry_point = program_info.get("entry_point", "main")
            if not module_path:
                messagebox.showerror("Config Error", f"ไม่พบ 'module_path' สำหรับ '{program_name}'")
                return
            
            try:
                kwargs = {'working_dir': self.program_dir}
                process = Process(target=run_module_entrypoint, args=(module_path, entry_point), kwargs={'script_kwargs': kwargs})
                process.start()
            except Exception as e:
                messagebox.showerror("Process Error", f"ไม่สามารถเริ่มโปรเซสสำหรับ '{program_name}' ได้:\n{e}")
                print(f"LAUNCHER_ERROR: Process creation failed for '{program_name}': {e}")
                return

        elif program_type == "external_exe":
            messagebox.showinfo("แจ้งเพื่อทราบ", "การคำนวณระยะเวลาใช้งานยังไม่รองรับโปรแกรมประเภท External .exe")
            command = program_info.get("command")
            if not command:
                messagebox.showerror("Config Error", f"ไม่พบ 'command' สำหรับ '{program_name}'")
                return
            try:
                subprocess.Popen(command, shell=True, cwd=self.launcher_base_dir)
            except Exception as e:
                messagebox.showerror("Error", f"ไม่สามารถเปิดโปรแกรม '{program_name}' ได้:\n{e}")
            return

        else:
            messagebox.showwarning("ไม่รองรับ", f"ไม่รู้จักประเภทโปรแกรม '{program_type}'")
            return

        if process and process.is_alive():
            monitor_thread = threading.Thread(
                target=self._wait_and_log_session,
                args=(process, program_info)
            )
            monitor_thread.daemon = True
            monitor_thread.start()
        else:
            print(f"LAUNCHER_WARNING: ไม่สามารถเริ่มเฝ้าดู '{program_name}' ได้เนื่องจาก process ไม่ทำงาน")




if __name__ == "__main__":
    freeze_support()

    launcher_dir_init = os.path.dirname(os.path.abspath(__file__))
    icon_dir_path = os.path.join(launcher_dir_init, ICON_FOLDER)
    program_dir_path = os.path.join(launcher_dir_init, PROGRAM_SUBFOLDER)

    # Create necessary folders
    for folder_path, folder_name in [(icon_dir_path, ICON_FOLDER), (program_dir_path, PROGRAM_SUBFOLDER)]:
        if not os.path.exists(folder_path):
            print(f"LAUNCHER_WARNING: ไม่พบโฟลเดอร์ '{folder_name}'.")
            try:
                os.makedirs(folder_path)
                print(f"LAUNCHER_INFO: สร้างโฟลเดอร์ '{folder_name}' ให้แล้ว")
            except Exception as e:
                print(f"LAUNCHER_ERROR: ไม่สามารถสร้างโฟลเดอร์ '{folder_name}': {e}")

    # สร้าง dummy Python files ใน All_Programs ถ้ายังไม่มี (สำหรับทดสอบ)
    # คุณควรลบส่วนนี้ออกเมื่อมีโปรแกรมจริง
    dummy_programs_to_create = {
        "Program_ItemdefSPSS_Log.py": "def run_this_app(working_dir=None):\n    print(f'Hello from ItemdefSPSS_Log! Working dir: {working_dir}')\n    import tkinter as tk\n    app = tk.Tk()\n    app.title('Itemdef SPSS Log')\n    tk.Label(app, text='This is Convert SPSS To Itemdef').pack()\n    app.mainloop()",
        "Program_T2B_Itermdef.py": "def run_this_app(working_dir=None):\n    print(f'Hello from T2B_Itermdef! Working dir: {working_dir}')\n    import tkinter as tk\n    app = tk.Tk()\n    app.title('T2B Itemdef')\n    tk.Label(app, text='This is Making T2B For Itemdef').pack()\n    app.mainloop()",
        "Program_Var_Value_SPSS_V6.py": "def run_this_app(working_dir=None):\n    print(f'Hello from Var_Value_SPSS_V6! Working dir: {working_dir}')\n    import tkinter as tk\n    app = tk.Tk()\n    app.title('Get Value SPSS')\n    tk.Label(app, text='This is Program_Get Value V6').pack()\n    app.mainloop()",
        "99_CleanSPSS_Germini.py": "def run_this_app(working_dir=None):\n    print(f'Hello from CleanSPSS_Germini! Working dir: {working_dir}')\n    import tkinter as tk\n    app = tk.Tk()\n    app.title('Clean SPSS')\n    tk.Label(app, text='This is CleanSPSS+Frequenzy').pack()\n    app.mainloop()",
        "Program_CutBrief_Excel.py": "def run_this_app(working_dir=None):\n    print(f'Hello from CutBrief_Excel! Working dir: {working_dir}')\n    import tkinter as tk\n    app = tk.Tk()\n    app.title('Cut Brief')\n    tk.Label(app, text='This is โปรแกรมตัดชุด').pack()\n    app.mainloop()",
        "Program_RenameSheet.py": "def run_this_app(working_dir=None):\n    print(f'Hello from RenameSheet! Working dir: {working_dir}')\n    import tkinter as tk\n    app = tk.Tk()\n    app.title('Rename Sheet')\n    tk.Label(app, text='This is โปรแกรม RenameSheet').pack()\n    app.mainloop()",
        "Dummy_ReporterToolAlpha.py": "def run_this_app(working_dir=None):\n    print(f'Hello from Dummy_ReporterToolAlpha! Working dir: {working_dir}')\n    import tkinter as tk\n    app = tk.Tk()\n    app.title('Reporter Tool Alpha')\n    tk.Label(app, text='This is Reporter Tool Alpha (Dummy)').pack()\n    app.mainloop()",
        "Dummy_UtilityX.py": "def run_this_app(working_dir=None):\n    print(f'Hello from Dummy_UtilityX! Working dir: {working_dir}')\n    import tkinter as tk\n    app = tk.Tk()\n    app.title('Utility X')\n    tk.Label(app, text='This is Utility X (Dummy)').pack()\n    app.mainloop()",
    }
    if not os.path.exists(program_dir_path): os.makedirs(program_dir_path)
    # Create __init__.py in All_Programs to make it a package
    init_py_path = os.path.join(program_dir_path, "__init__.py")
    if not os.path.exists(init_py_path):
        with open(init_py_path, "w") as f:
            f.write("# This file makes All_Programs a package\n")
            print(f"LAUNCHER_INFO: Created '{init_py_path}'")

    for filename, content in dummy_programs_to_create.items():
        filepath = os.path.join(program_dir_path, filename)
        if not os.path.exists(filepath):
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(content)
            print(f"LAUNCHER_INFO: Created dummy file '{filepath}' for testing.")


    app = AppLauncher()
    # --- เรียกใช้ฟังก์ชันแสดง Changelog ที่นี่! ---
    show_changelog_if_exists()
    # ---------------------------------------------
    app.after(1000, lambda: check_for_updates(app))
    try:
        main_icon_relative_path = os.path.join(ICON_FOLDER, "I_Main.ico")
        main_icon_actual_path = resource_path(main_icon_relative_path)

        if os.path.exists(main_icon_actual_path):
            app.iconbitmap(main_icon_actual_path)
            print(f"LAUNCHER_INFO: Main application icon set from: {main_icon_actual_path}")
        else:
            print(f"LAUNCHER_WARNING: Main application icon ('I_Main.ico') not found at: {main_icon_actual_path}")
    except Exception as e:
        print(f"LAUNCHER_WARNING: ไม่สามารถโหลดไอคอนหลักของโปรแกรมได้: {e}")

    app.mainloop()
