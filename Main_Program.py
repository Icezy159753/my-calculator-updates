import os
import sys
from PyQt6 import QtCore, QtGui, QtWidgets
import importlib
from multiprocessing import Process, freeze_support
import subprocess
import requests
from packaging.version import parse as parse_version
import getpass
import threading
from datetime import datetime
import socket # เพิ่มเข้ามาเพื่อดึง IP Address (ถ้าต้องการ)
import ctypes


class Spinner(QtWidgets.QWidget):
    def __init__(self, parent=None, radius=10, line_width=3, speed=140):
        super().__init__(parent)
        self._radius = radius
        self._line_width = line_width
        self._angle = 0
        self._timer = QtCore.QTimer(self)
        self._timer.timeout.connect(self._on_timeout)
        self._timer.start(speed)
        size = (radius * 2) + line_width * 2
        self.setFixedSize(size, size)

    def _on_timeout(self):
        self._angle = (self._angle + 30) % 360
        self.update()

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        painter.translate(self.width() / 2, self.height() / 2)
        painter.rotate(self._angle)
        pen = QtGui.QPen(QtGui.QColor(255, 255, 255, 0))
        pen.setWidth(self._line_width)
        painter.setPen(pen)

        # Draw 12 segments with fading alpha
        for i in range(12):
            color = QtGui.QColor(self.palette().color(QtGui.QPalette.ColorRole.Text))
            color.setAlphaF((i + 1) / 12)
            pen.setColor(color)
            painter.setPen(pen)
            painter.drawLine(0, -self._radius, 0, -self._radius + (self._radius // 2))
            painter.rotate(30)

# --- เพิ่มค่าคงที่นี้ไว้ใกล้ๆกับค่าคงที่อื่นๆ ด้านบน ---
# !!! สำคัญ: ให้เปลี่ยน URL นี้เป็น URL ของ Web App ที่ได้จากการ Deploy บน Google Apps Script ของคุณ
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzLISY7ormRaB05x3qBD41apZ8zVMx2_-nNrlSz1RP26DCXXQgfpfESxS6ppgxkyOSm/exec" # <--- ใส่ URL ของคุณที่นี่


# --- ค่าคงที่สำหรับชื่อโฟลเดอร์ ---
PROGRAM_SUBFOLDER = "All_Programs"
ICON_FOLDER = "Icon"
# --- ข้อมูลโปรแกรมและ GitHub (สำคัญมาก: ต้องเปลี่ยนเป็นของคุณ) ---
CURRENT_VERSION = "1.0.86"
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
            if ask_yes_no(app_window, "Update Available", f"มีเวอร์ชันใหม่ ({latest_version})!\nต้องการอัปเดตตอนนี้เลยหรือไม่?"):
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
                # หา URL ของไฟล์ updater.exe และไฟล์ zip จาก release ล่าสุด
                updater_url = None
                app_url = None
                for asset in latest_release['assets']:
                    if asset['name'] == 'updater.exe':
                        updater_url = asset['browser_download_url']
                    if asset['name'] == 'Main_Program.zip':
                        app_url = asset['browser_download_url']

                if not updater_url or not app_url:
                    show_message(app_window, "Error", "ไม่พบไฟล์สำหรับอัปเดตใน Release ล่าสุด", QtWidgets.QMessageBox.Icon.Critical)
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
                app_exe_path = get_executable_path()
                app_dir = os.path.dirname(app_exe_path)
                app_exe_name = os.path.basename(app_exe_path)
                subprocess.Popen([updater_path, str(os.getpid()), app_dir, app_exe_name, app_url])
                app_window.close() # หรือ sys.exit()

        else:
            print("You have the latest version.")

    except Exception as e:
        print(f"Could not check for updates: {e}")

# --- เพิ่มฟังก์ชันนี้เข้าไปใหม่ทั้งหมด ---
def create_custom_changelog_window(parent, changelog_content):
    """
    สร้างและแสดงหน้าต่าง Changelog แบบกำหนดเองด้วย PyQt6
    """
    try:
        dialog = QtWidgets.QDialog(parent)
        dialog.setWindowTitle(f"อัปเดตสำเร็จเป็นเวอร์ชัน {CURRENT_VERSION}!")
        dialog.setModal(True)

        icon_path = resource_path(os.path.join(ICON_FOLDER, "I_Main.ico"))
        if os.path.exists(icon_path):
            dialog.setWindowIcon(QtGui.QIcon(icon_path))

        dialog.resize(520, 380)

        layout = QtWidgets.QVBoxLayout(dialog)
        title_label = QtWidgets.QLabel("มีอะไรใหม่ในเวอร์ชันนี้:")
        title_label.setStyleSheet("font-weight: 600; font-size: 14px;")
        layout.addWidget(title_label)

        textbox = QtWidgets.QTextEdit()
        textbox.setReadOnly(True)
        textbox.setText(changelog_content)
        layout.addWidget(textbox)

        ok_button = QtWidgets.QPushButton("OK")
        ok_button.clicked.connect(dialog.accept)
        ok_button.setFixedWidth(100)
        ok_row = QtWidgets.QHBoxLayout()
        ok_row.addStretch(1)
        ok_row.addWidget(ok_button)
        ok_row.addStretch(1)
        layout.addLayout(ok_row)

        dialog.exec()

    except Exception as e:
        print(f"Failed to create custom changelog window: {e}")
        show_message(
            parent,
            f"อัปเดตสำเร็จเป็นเวอร์ชัน {CURRENT_VERSION}!",
            f"มีอะไรใหม่ในเวอร์ชันนี้:\n\n{changelog_content}",
            QtWidgets.QMessageBox.Icon.Information
        )


def show_changelog_if_exists(parent):
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
                create_custom_changelog_window(parent, changelog_content)
    except Exception as e:
        print(f"Could not process changelog file: {e}")



# --- ฟังก์ชัน Resource Path ---
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# --- UI theme constants (PyQt6) ---
THEME_LIGHT = {
    "app_bg": "#F2F4F7",
    "sidebar_bg": "#FFFFFF",
    "content_bg": "#F6F7FA",
    "card_bg": "#FFFFFF",
    "card_border": "#E6E9EF",
    "text_primary": "#1C2430",
    "text_muted": "#6B7785",
    "accent": "#2E7D6B",
    "accent_hover": "#276B5C",
    "search_bg": "#FFFFFF",
    "search_border": "#D8DDE6"
}

THEME_DARK = {
    "app_bg": "#171A1F",
    "sidebar_bg": "#12161B",
    "content_bg": "#1A1F25",
    "card_bg": "#232A32",
    "card_border": "#2F3742",
    "text_primary": "#F4EDE7",
    "text_muted": "#FFFFFF",
    "accent": "#3BA88F",
    "accent_hover": "#32947E",
    "search_bg": "#1F262E",
    "search_border": "#303845"
}

DEFAULT_APPEARANCE_MODE = "Light"


def show_message(parent, title, text, icon):
    box = QtWidgets.QMessageBox(parent)
    box.setWindowTitle(title)
    box.setText(text)
    box.setIcon(icon)
    box.exec()


def ask_yes_no(parent, title, text):
    result = QtWidgets.QMessageBox.question(
        parent,
        title,
        text,
        QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No
    )
    return result == QtWidgets.QMessageBox.StandardButton.Yes


def show_error_dialog(title, text):
    if QtWidgets.QApplication.instance() is None:
        print(f"{title}: {text}")
        return
    show_message(None, title, text, QtWidgets.QMessageBox.Icon.Critical)

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
MAX_COLUMNS = 3
CARD_MIN_WIDTH = 240
CARD_MAX_WIDTH = 2000
CARD_HEIGHT = 320

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
            show_error_dialog("Launch Error", f"ไม่พบฟังก์ชันหลัก '{entry_point_func_name}'\nในโมดูล '{module_name_in_subfolder}'.")

    except ImportError as e:
        print(f"LAUNCHER_ERROR: Error importing module {full_module_name}: {e}")
        show_error_dialog("Launch Error", f"ไม่สามารถโหลดโมดูล '{module_name_in_subfolder}' ได้:\n{e}\n\nตรวจสอบว่าไฟล์ .py อยู่ในโฟลเดอร์ '{PROGRAM_SUBFOLDER}' และ sys.path ถูกต้อง")
    except Exception as e:
        print(f"LAUNCHER_ERROR: Error running module {full_module_name}: {e}")
        show_error_dialog("Runtime Error", f"เกิดข้อผิดพลาดขณะรัน '{module_name_in_subfolder}':\n{e}")



class AppLauncher(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Program All DP v{CURRENT_VERSION}")

        self._sized_once = False
        window_width = 1120
        window_height = 760
        self.setFixedSize(window_width, window_height)
        screen = QtWidgets.QApplication.primaryScreen()
        if screen:
            screen_rect = screen.availableGeometry()
            self.move(
                (screen_rect.width() - window_width) // 2,
                (screen_rect.height() - window_height) // 2
            )

        self.font_title = QtGui.QFont("Tahoma", 20, QtGui.QFont.Weight.Bold)
        self.font_section = QtGui.QFont("Tahoma", 14, QtGui.QFont.Weight.Bold)
        self.font_body = QtGui.QFont("Tahoma", 12)
        self.font_small = QtGui.QFont("Tahoma", 11)
        self.font_card_title = QtGui.QFont("Tahoma", 13, QtGui.QFont.Weight.Bold)
        self.font_button = QtGui.QFont("Tahoma", 12, QtGui.QFont.Weight.Bold)

        self.launcher_base_dir = resource_path(".")
        self.icon_dir = os.path.join(self.launcher_base_dir, ICON_FOLDER)
        self.program_dir = os.path.join(self.launcher_base_dir, PROGRAM_SUBFOLDER)

        if self.launcher_base_dir not in sys.path:
            sys.path.insert(0, self.launcher_base_dir)
            print(f"LAUNCHER_INFO: Added to sys.path: {self.launcher_base_dir}")

        self.icon_cache = {}
        self.current_category = "All"
        self.current_programs = []
        self.last_columns = 0
        self.last_card_width = 0
        self.launch_dialog = None
        self.launch_handle = None
        self.launch_wait_started = None

        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        main_layout = QtWidgets.QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.sidebar_frame = QtWidgets.QFrame()
        self.sidebar_frame.setObjectName("Sidebar")
        self.sidebar_frame.setFixedWidth(220)
        main_layout.addWidget(self.sidebar_frame)

        self.content_frame = QtWidgets.QFrame()
        self.content_frame.setObjectName("Content")
        main_layout.addWidget(self.content_frame)

        self.build_sidebar()
        self.build_content()

        self.apply_theme(DEFAULT_APPEARANCE_MODE)
        self.show_category_programs("All")
        QtCore.QTimer.singleShot(0, self.update_program_grid)

    def build_sidebar(self):
        layout = QtWidgets.QVBoxLayout(self.sidebar_frame)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(6)

        brand_row = QtWidgets.QHBoxLayout()
        brand_row.setContentsMargins(0, 0, 0, 0)
        brand_row.setSpacing(8)

        brand_icon = QtWidgets.QLabel()
        brand_pixmap = self.load_icon_pixmap("I_Main.ico", (28, 28))
        if brand_pixmap:
            brand_icon.setPixmap(brand_pixmap)
        brand_icon.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter)
        brand_row.addWidget(brand_icon)

        brand_text = QtWidgets.QVBoxLayout()
        brand_title = QtWidgets.QLabel("Program All DP")
        brand_title.setFont(self.font_section)
        brand_title.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter)
        brand_text.addWidget(brand_title)

        brand_row.addLayout(brand_text)
        layout.addLayout(brand_row)

        logo_label = QtWidgets.QLabel("หมวดหมู่")
        logo_label.setFont(self.font_title)
        layout.addWidget(logo_label)

        raw_categories = set(p["category"] for p in PROGRAMS if p.get("category") and p.get("enabled", True))
        categories = ["All"] + sorted(list(raw_categories))

        self.category_group = QtWidgets.QButtonGroup(self)
        self.category_group.setExclusive(True)
        self.category_buttons = {}

        for category_name in categories:
            button = QtWidgets.QPushButton(category_name)
            button.setCheckable(True)
            button.setFont(self.font_body)
            button.setObjectName("CategoryButton")
            button.clicked.connect(lambda checked, cat=category_name: self.show_category_programs(cat))
            self.category_group.addButton(button)
            layout.addWidget(button)
            self.category_buttons[category_name] = button

        appearance_label = QtWidgets.QLabel("Appearance Mode:")
        appearance_label.setFont(self.font_small)
        appearance_label.setObjectName("Muted")
        layout.addSpacing(8)
        layout.addWidget(appearance_label)

        self.appearance_mode_menu = QtWidgets.QComboBox()
        self.appearance_mode_menu.addItems(["Light", "Dark", "System"])
        self.appearance_mode_menu.setCurrentText(DEFAULT_APPEARANCE_MODE)
        self.appearance_mode_menu.currentTextChanged.connect(self.change_appearance_mode_event)
        layout.addWidget(self.appearance_mode_menu)

        layout.addStretch(1)

        web_script_icon = self.load_icon("yt.ico")
        self.web_script_button = QtWidgets.QPushButton()
        if web_script_icon:
            self.web_script_button.setIcon(web_script_icon)
            self.web_script_button.setIconSize(QtCore.QSize(26, 26))
        self.web_script_button.setFixedSize(52, 52)
        self.web_script_button.clicked.connect(self.open_web_script_link)
        self.web_script_button.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        layout.addWidget(self.web_script_button, alignment=QtCore.Qt.AlignmentFlag.AlignHCenter)

        yt_label = QtWidgets.QLabel("คลิกเพื่อดู VDO การใช้งาน")
        yt_label.setFont(self.font_small)
        yt_label.setObjectName("Muted")
        yt_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(yt_label)

        log_button = QtWidgets.QPushButton("ประวัติการอัปเดต")
        log_button.setObjectName("PrimaryButton")
        log_button.setFont(self.font_body)
        log_button.clicked.connect(self.show_update_log_window)
        layout.addWidget(log_button)

        version_label = QtWidgets.QLabel(f"เวอร์ชัน {CURRENT_VERSION}")
        version_label.setFont(self.font_body)
        version_label.setStyleSheet("color: #1E63B5; font-weight: 600;")
        version_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(version_label)

        exit_button = QtWidgets.QPushButton("ปิดProgram")
        exit_button.setObjectName("ExitButton")
        exit_button.setFont(QtGui.QFont("Tahoma", 12, QtGui.QFont.Weight.Bold))
        exit_button.clicked.connect(self.quit_app)
        layout.addWidget(exit_button)

    def build_content(self):
        layout = QtWidgets.QVBoxLayout(self.content_frame)
        layout.setContentsMargins(4, 14, 4, 14)
        layout.setSpacing(10)

        header_row = QtWidgets.QHBoxLayout()
        header_row.setSpacing(12)

        self.header_title = QtWidgets.QLabel("All Tools")
        self.header_title.setFont(self.font_title)
        header_row.addWidget(self.header_title)

        self.header_count = QtWidgets.QLabel("")
        self.header_count.setFont(self.font_small)
        self.header_count.setObjectName("Muted")
        header_row.addWidget(self.header_count)

        header_row.addStretch(1)

        self.search_entry = QtWidgets.QLineEdit()
        self.search_entry.setPlaceholderText("Search tools")
        self.search_entry.setFixedHeight(34)
        self.search_entry.setFixedWidth(220)
        self.search_entry.setFont(self.font_body)
        self.search_entry.textChanged.connect(self.update_program_grid)
        self.search_entry.setClearButtonEnabled(True)
        header_row.addWidget(self.search_entry)

        layout.addLayout(header_row)

        self.scroll_area = QtWidgets.QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        self.scroll_area.verticalScrollBar().setSingleStep(16)
        self.scroll_area.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self.cards_container = QtWidgets.QWidget()
        self.cards_container.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Expanding
        )
        self.cards_grid = QtWidgets.QGridLayout(self.cards_container)
        self.cards_grid.setContentsMargins(0, 0, 0, 0)
        self.cards_grid.setHorizontalSpacing(10)
        self.cards_grid.setVerticalSpacing(12)
        self.cards_grid.setAlignment(QtCore.Qt.AlignmentFlag.AlignTop | QtCore.Qt.AlignmentFlag.AlignLeft)
        self.scroll_area.setWidget(self.cards_container)
        layout.addWidget(self.scroll_area)

    def apply_theme(self, mode):
        if mode == "Dark":
            theme = THEME_DARK
        elif mode == "Light":
            theme = THEME_LIGHT
        else:
            theme = THEME_LIGHT
        self.theme = theme

        self.setStyleSheet(f"""
            QMainWindow {{ background: {theme['app_bg']}; }}
            QFrame#Sidebar {{ background: {theme['sidebar_bg']}; }}
            QFrame#Content {{ background: {theme['content_bg']}; }}
            QScrollArea {{
                background: {theme['content_bg']};
            }}
            QScrollArea > QWidget {{
                background: {theme['content_bg']};
            }}
            QScrollArea > QWidget > QWidget {{
                background: transparent;
            }}
            QLabel {{ color: {theme['text_primary']}; background: transparent; }}
            QLabel#Muted {{ color: {theme['text_muted']}; }}
            QLineEdit {{
                background: {theme['search_bg']};
                border: 1px solid {theme['search_border']};
                border-radius: 14px;
                padding: 6px 10px;
                color: {theme['text_primary']};
            }}
            QPushButton#CategoryButton {{
                background: transparent;
                border-radius: 10px;
                padding: 6px 10px;
                text-align: left;
                color: {theme['text_primary']};
            }}
            QPushButton#CategoryButton:hover {{
                background: {theme['card_border']};
            }}
            QPushButton#CategoryButton:checked {{
                background: {theme['accent']};
                color: white;
            }}
            QPushButton#PrimaryButton {{
                background: {theme['accent']};
                color: white;
                border-radius: 10px;
                padding: 6px 12px;
            }}
            QPushButton#PrimaryButton:hover {{
                background: {theme['accent_hover']};
            }}
            QPushButton#ExitButton {{
                background: #8B1E1E;
                border: 1px solid #701919;
                border-radius: 10px;
                padding: 6px 12px;
                color: white;
            }}
            QPushButton#ExitButton:hover {{
                background: #701919;
            }}
            QFrame#Card {{
                background: {theme['card_bg']};
                border: 1px solid {theme['card_border']};
                border-radius: 14px;
            }}
            QTableView::item:selected {{
                background: transparent;
                color: {theme['text_primary']};
            }}
            QTableView::item:selected:active {{
                background: transparent;
                color: {theme['text_primary']};
            }}
            QTextEdit {{
                background: {theme['card_bg']};
                border: 1px solid {theme['card_border']};
                color: {theme['text_primary']};
            }}
            QComboBox {{
                background: {theme['search_bg']};
                border: 1px solid {theme['search_border']};
                border-radius: 8px;
                padding: 4px 8px;
                color: {theme['text_primary']};
            }}
        """)

    def change_appearance_mode_event(self, new_appearance_mode: str):
        self.apply_theme(new_appearance_mode)

    def quit_app(self):
        self.close()

    def fetch_update_history_text(self):
        """
        ดึงข้อมูลทุก Release จาก GitHub API แล้วจัดรูปแบบเป็นข้อความ
        """
        history_log = []
        try:
            api_url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases"
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()

            releases = response.json()
            if not releases:
                return "ไม่พบประวัติการอัปเดต"

            for release in releases:
                version = release.get("tag_name", "N/A")
                date = release.get("published_at", "").split("T")[0]
                body = release.get("body", "ไม่มีรายละเอียด")
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
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("ประวัติการอัปเดต")
        dialog.resize(600, 500)
        icon_path = resource_path(os.path.join(ICON_FOLDER, "I_Main.ico"))
        if os.path.exists(icon_path):
            dialog.setWindowIcon(QtGui.QIcon(icon_path))

        layout = QtWidgets.QVBoxLayout(dialog)
        textbox = QtWidgets.QTextEdit()
        textbox.setReadOnly(True)
        textbox.setFont(self.font_body)
        textbox.setText("กำลังดึงข้อมูลประวัติการอัปเดต กรุณารอสักครู่...")
        layout.addWidget(textbox)
        QtWidgets.QApplication.processEvents()

        history_text = self.fetch_update_history_text()
        textbox.setText(history_text)
        dialog.exec()

    def open_web_script_link(self):
        """
        เปิดลิงก์ไปยังเว็บสคริปต์ในเบราว์เซอร์เริ่มต้นของผู้ใช้
        """
        import webbrowser

        url = "https://script.google.com/macros/s/AKfycbxS0tIe8TnDUf-QNn6Y0NdlT-MbQY-FmyM2_uR03muhIcg05z_0F9mFHt9ReNGpns2H/exec"
        try:
            print(f"LAUNCHER_INFO: Opening URL: {url}")
            webbrowser.open_new_tab(url)
        except Exception as e:
            print(f"LAUNCHER_ERROR: Could not open URL: {e}")
            show_message(self, "เกิดข้อผิดพลาด", f"ไม่สามารถเปิดลิงก์ได้:\n{e}", QtWidgets.QMessageBox.Icon.Critical)

    def show_category_programs(self, category_name):
        self.current_category = category_name
        button = self.category_buttons.get(category_name)
        if button:
            button.setChecked(True)
        self.update_program_grid()

    def update_program_grid(self):
        search_query = self.search_entry.text().strip().lower()

        if self.current_category == "All":
            programs_to_display = [p for p in PROGRAMS if p.get("enabled", True)]
            header_text = "All Tools"
        else:
            programs_to_display = [
                p for p in PROGRAMS
                if p.get("enabled", True) and p.get("category") == self.current_category
            ]
            header_text = f"หมวดหมู่: {self.current_category}"

        if search_query:
            programs_to_display = [
                p for p in programs_to_display
                if search_query in p.get("name", "").lower()
                or search_query in p.get("description", "").lower()
            ]
            header_text = f"{header_text} • ค้นหา: {self.search_entry.text().strip()}"

        self.header_title.setText(header_text)
        total_programs = len([p for p in PROGRAMS if p.get("enabled", True)])
        visible_programs = len(programs_to_display)
        if total_programs == visible_programs:
            self.header_count.setText(f"ทั้งหมด {total_programs} โปรแกรม")
        else:
            self.header_count.setText(f"แสดง {visible_programs} / {total_programs} โปรแกรม")
        self.current_programs = programs_to_display
        self.render_program_cards()


    def render_program_cards(self):
        for i in reversed(range(self.cards_grid.count())):
            item = self.cards_grid.takeAt(i)
            widget = item.widget()
            if widget:
                widget.setParent(None)

        if not self.current_programs:
            empty_label = QtWidgets.QLabel("ไม่พบโปรแกรมที่ตรงกับคำค้นหา")
            empty_label.setFont(self.font_section)
            empty_label.setObjectName("Muted")
            empty_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            self.cards_grid.addWidget(empty_label, 0, 0)
            return

        max_cols = MAX_COLUMNS
        card_width = self.calculate_card_width()
        if card_width != self.last_card_width:
            self.last_card_width = card_width
        row_num = 0
        col_num = 0

        for program in self.current_programs:
            card = self.create_card_widget(program, card_width)
            self.cards_grid.addWidget(card, row_num, col_num)
            col_num += 1
            if col_num >= max_cols:
                col_num = 0
                row_num += 1

        total_spacing = self.cards_grid.horizontalSpacing() * (max_cols - 1)
        grid_width = (card_width * max_cols) + total_spacing
        viewport_width = max(1, self.scroll_area.viewport().width())
        target_width = max(grid_width, viewport_width)
        self.cards_container.setMinimumWidth(target_width)
        self.cards_container.setMaximumWidth(target_width)
        self.last_columns = max_cols

    def calculate_card_width(self):
        available_width = max(1, self.scroll_area.viewport().width() - 4)
        total_spacing = self.cards_grid.horizontalSpacing() * (MAX_COLUMNS - 1)
        raw_width = max(1, (available_width - total_spacing) / MAX_COLUMNS)
        card_width = int(raw_width)
        if card_width < CARD_MIN_WIDTH:
            return max(170, card_width)
        return min(CARD_MAX_WIDTH, card_width)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.current_programs:
            new_width = self.calculate_card_width()
            if new_width != self.last_card_width:
                self.render_program_cards()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._sized_once:
            screen = self.screen()
            if screen:
                screen_rect = screen.availableGeometry()
                target_width = min(1120, screen_rect.width())
                target_height = min(760, screen_rect.height())
                self.setFixedSize(target_width, target_height)
                self.move(
                    (screen_rect.width() - target_width) // 2,
                    (screen_rect.height() - target_height) // 2
                )
            QtCore.QTimer.singleShot(0, self.update_program_grid)
            QtCore.QTimer.singleShot(50, self.update_program_grid)
            self._sized_once = True

    def get_icon_path(self, icon_name):
        if not icon_name:
            return None
        icon_path = os.path.join(self.icon_dir, icon_name)
        if os.path.exists(icon_path):
            return icon_path
        print(f"LAUNCHER_WARNING: Icon file not found: {icon_path}")
        return None

    def load_icon(self, icon_name):
        if not icon_name:
            return None
        if icon_name in self.icon_cache:
            return self.icon_cache[icon_name]
        icon_path = self.get_icon_path(icon_name)
        if not icon_path:
            return None
        icon = QtGui.QIcon(icon_path)
        self.icon_cache[icon_name] = icon
        return icon

    def load_icon_pixmap(self, icon_name, size):
        icon_path = self.get_icon_path(icon_name)
        if not icon_path:
            return None
        pixmap = QtGui.QPixmap(icon_path)
        return pixmap.scaled(
            size[0],
            size[1],
            QtCore.Qt.AspectRatioMode.KeepAspectRatio,
            QtCore.Qt.TransformationMode.SmoothTransformation
        )

    def create_card_widget(self, program, card_width):
        card = QtWidgets.QFrame()
        card.setObjectName("Card")
        card.setFixedWidth(card_width)
        card.setFixedHeight(CARD_HEIGHT)
        card_layout = QtWidgets.QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(6)

        icon_label = QtWidgets.QLabel()
        pixmap = self.load_icon_pixmap(program.get("icon"), ICON_SIZE)
        if pixmap:
            icon_label.setPixmap(pixmap)
        icon_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(icon_label)

        name_label = QtWidgets.QLabel(program["name"])
        name_label.setFont(self.font_card_title)
        name_label.setWordWrap(True)
        name_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        title_metrics = QtGui.QFontMetrics(self.font_card_title)
        name_label.setFixedHeight((title_metrics.lineSpacing() * 5) + 6)
        card_layout.addWidget(name_label)

        desc_label = QtWidgets.QLabel(program["description"])
        desc_label.setFont(self.font_small)
        desc_label.setObjectName("Muted")
        desc_label.setWordWrap(True)
        desc_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignTop)
        desc_label.setFixedHeight(72)
        card_layout.addWidget(desc_label)

        launch_button = QtWidgets.QPushButton("เปิดโปรแกรม")
        launch_button.setFont(self.font_button)
        launch_button.setObjectName("PrimaryButton")
        launch_button.clicked.connect(lambda checked=False, p=program: self.launch_program(p))
        card_layout.addWidget(launch_button)

        return card

    def show_launching_dialog(self, program_name):
        if self.launch_dialog:
            return
        dialog = QtWidgets.QDialog(self)
        dialog.setModal(False)
        dialog.setWindowFlags(
            QtCore.Qt.WindowType.FramelessWindowHint
            | QtCore.Qt.WindowType.WindowStaysOnTopHint
            | QtCore.Qt.WindowType.Tool
        )
        dialog.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground, True)
        dialog.setFixedSize(360, 160)

        outer_layout = QtWidgets.QVBoxLayout(dialog)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        card = QtWidgets.QFrame()
        card.setObjectName("LaunchOverlay")
        card_layout = QtWidgets.QVBoxLayout(card)
        card_layout.setContentsMargins(18, 16, 18, 16)
        card_layout.setSpacing(10)

        title = QtWidgets.QLabel("กำลังเปิดโปรแกรม")
        title.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        title.setFont(self.font_section)
        card_layout.addWidget(title)

        label = QtWidgets.QLabel(f"{program_name}\nกรุณารอสักครู่")
        label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        label.setFont(self.font_body)
        card_layout.addWidget(label)

        spinner = Spinner(card, radius=11, line_width=3, speed=120)
        spinner_row = QtWidgets.QHBoxLayout()
        spinner_row.addStretch(1)
        spinner_row.addWidget(spinner)
        spinner_row.addStretch(1)
        card_layout.addLayout(spinner_row)

        shadow = QtWidgets.QGraphicsDropShadowEffect()
        shadow.setBlurRadius(24)
        shadow.setOffset(0, 6)
        shadow.setColor(QtGui.QColor(0, 0, 0, 80))
        card.setGraphicsEffect(shadow)

        card.setStyleSheet(f"""
            QFrame#LaunchOverlay {{
                background: {self.theme['card_bg']};
                border: 1px solid {self.theme['card_border']};
                border-radius: 16px;
            }}
            QLabel {{
                color: {self.theme['text_primary']};
            }}
        """)

        outer_layout.addWidget(card)

        self.launch_status_label = label
        self.launch_dots = 0
        self.launch_timer = QtCore.QTimer(dialog)
        self.launch_timer.setInterval(300)
        self.launch_timer.timeout.connect(self._tick_launching_animation)
        self.launch_timer.start()
        self.launch_wait_started = QtCore.QElapsedTimer()
        self.launch_wait_started.start()

        dialog.show()
        dialog.raise_()
        dialog.activateWindow()
        QtWidgets.QApplication.processEvents()
        self.launch_dialog = dialog

    def _tick_launching_animation(self):
        if not self.launch_status_label:
            return
        dots = "." * (self.launch_dots % 4)
        base = self.launch_status_label.text().split("\n")[0]
        self.launch_status_label.setText(f"{base}\nกรุณารอสักครู่{dots}")
        self.launch_dots += 1

    def close_launching_dialog(self):
        if self.launch_dialog:
            if hasattr(self, "launch_timer") and self.launch_timer:
                self.launch_timer.stop()
            self.launch_status_label = None
            self.launch_dots = 0
            self.launch_handle = None
            self.launch_wait_started = None
            self.launch_dialog.close()
            self.launch_dialog = None

    def _wait_for_launch_ready(self, handle):
        if not self.launch_dialog:
            return
        if self.launch_wait_started and self.launch_wait_started.elapsed() > 30000:
            self.close_launching_dialog()
            return

        is_ready = False
        if handle is None:
            is_ready = True
        else:
            pid = None
            if isinstance(handle, Process):
                pid = handle.pid
            else:
                try:
                    pid = handle.pid
                except Exception:
                    pid = None

            if pid:
                try:
                    is_ready = self._has_visible_window(pid)
                except Exception:
                    is_ready = False
            else:
                is_ready = False

        if is_ready:
            self.close_launching_dialog()
        else:
            QtCore.QTimer.singleShot(200, lambda: self._wait_for_launch_ready(handle))

    def _has_visible_window(self, pid):
        user32 = ctypes.windll.user32
        visible = False

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def enum_proc(hwnd, _):
            nonlocal visible
            if not user32.IsWindowVisible(hwnd):
                return True
            length = user32.GetWindowTextLengthW(hwnd)
            if length == 0:
                return True
            lpdw_process_id = ctypes.c_uint()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(lpdw_process_id))
            if lpdw_process_id.value == pid:
                visible = True
                return False
            return True

        user32.EnumWindows(enum_proc, 0)
        return visible

    def log_session_to_sheet(self, program_name, user_info, start_time, end_time, duration_formatted):
        """
        ส่งข้อมูลเซสชันการใช้งาน (เวลาเริ่ม-จบ, ระยะเวลา) ไปยัง Google Sheet
        (เวอร์ชันนี้ส่งระยะเวลาเป็นรูปแบบ HH:MM:SS)
        """
        if "YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE" in GOOGLE_SCRIPT_URL:
            print("LOGGING_WARNING: กรุณาเปลี่ยน GOOGLE_SCRIPT_URL เป็น URL ของ Web App จริง")
            return

        try:
            payload = {
                'startDate': start_time.strftime("%Y-%m-%d"),
                'startTime': start_time.strftime("%H:%M:%S"),
                'endTime': end_time.strftime("%H:%M:%S"),
                'duration': duration_formatted,
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

        duration_seconds = int((end_time - start_time).total_seconds())
        hours, remainder = divmod(duration_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        duration_formatted = f"{hours:02}:{minutes:02}:{seconds:02}"

        self.log_session_to_sheet(program_name, user_info, start_time, end_time, duration_formatted)


    def launch_program(self, program_info):
        """
        ฟังก์ชันนี้ถูกแก้ไขเพื่อสร้าง Thread แยกสำหรับเฝ้าดูแต่ละโปรแกรมที่เปิด
        """
        program_name = program_info.get("name", "Unknown Program")
        program_type = program_info.get("type", "unknown")
        process = None

        print(f"LAUNCHER_INFO: กำลังเตรียมเปิด '{program_name}' (Type: {program_type})")

        if program_type == "local_py_module":
            module_path = program_info.get("module_path")
            entry_point = program_info.get("entry_point", "main")
            if not module_path:
                show_message(self, "Config Error", f"ไม่พบ 'module_path' สำหรับ '{program_name}'", QtWidgets.QMessageBox.Icon.Critical)
                return
            
            try:
                self.show_launching_dialog(program_name)
                kwargs = {'working_dir': self.program_dir}
                process = Process(target=run_module_entrypoint, args=(module_path, entry_point), kwargs={'script_kwargs': kwargs})
                process.start()
                self.launch_handle = process
                QtCore.QTimer.singleShot(200, lambda: self._wait_for_launch_ready(process))
            except Exception as e:
                self.close_launching_dialog()
                show_message(self, "Process Error", f"ไม่สามารถเริ่มโปรเซสสำหรับ '{program_name}' ได้:\n{e}", QtWidgets.QMessageBox.Icon.Critical)
                print(f"LAUNCHER_ERROR: Process creation failed for '{program_name}': {e}")
                return

        elif program_type == "external_exe":
            command = program_info.get("command")
            if not command:
                show_message(self, "Config Error", f"ไม่พบ 'command' สำหรับ '{program_name}'", QtWidgets.QMessageBox.Icon.Critical)
                return
            try:
                self.show_launching_dialog(program_name)
                popen_proc = subprocess.Popen(command, shell=True, cwd=self.launcher_base_dir)
                self.launch_handle = popen_proc
                QtCore.QTimer.singleShot(200, lambda: self._wait_for_launch_ready(popen_proc))
            except Exception as e:
                self.close_launching_dialog()
                show_message(self, "Error", f"ไม่สามารถเปิดโปรแกรม '{program_name}' ได้:\n{e}", QtWidgets.QMessageBox.Icon.Critical)
            return

        else:
            show_message(self, "ไม่รองรับ", f"ไม่รู้จักประเภทโปรแกรม '{program_type}'", QtWidgets.QMessageBox.Icon.Warning)
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


    if hasattr(QtCore.Qt.ApplicationAttribute, "AA_EnableHighDpiScaling"):
        QtCore.QCoreApplication.setAttribute(
            QtCore.Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True
        )
    if hasattr(QtCore.Qt.ApplicationAttribute, "AA_UseHighDpiPixmaps"):
        QtCore.QCoreApplication.setAttribute(
            QtCore.Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True
        )
    if hasattr(QtCore.Qt, "HighDpiScaleFactorRoundingPolicy"):
        QtGui.QGuiApplication.setHighDpiScaleFactorRoundingPolicy(
            QtCore.Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )

    qt_app = QtWidgets.QApplication(sys.argv)
    qt_app.setFont(QtGui.QFont("Tahoma", 10))
    window = AppLauncher()
    # --- เรียกใช้ฟังก์ชันแสดง Changelog ที่นี่! ---
    show_changelog_if_exists(window)
    # ---------------------------------------------
    QtCore.QTimer.singleShot(1000, lambda: check_for_updates(window))
    try:
        main_icon_relative_path = os.path.join(ICON_FOLDER, "I_Main.ico")
        main_icon_actual_path = resource_path(main_icon_relative_path)

        if os.path.exists(main_icon_actual_path):
            window.setWindowIcon(QtGui.QIcon(main_icon_actual_path))
            print(f"LAUNCHER_INFO: Main application icon set from: {main_icon_actual_path}")
        else:
            print(f"LAUNCHER_WARNING: Main application icon ('I_Main.ico') not found at: {main_icon_actual_path}")
    except Exception as e:
        print(f"LAUNCHER_WARNING: ไม่สามารถโหลดไอคอนหลักของโปรแกรมได้: {e}")

    window.show()
    sys.exit(qt_app.exec())
