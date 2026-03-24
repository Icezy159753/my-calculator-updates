# Context.md - Main_Program.py

วันที่จัดทำ: 2026-03-16  
ไฟล์ต้นทาง: `Main_Program.py`  
เวอร์ชันปัจจุบันในโค้ด: `1.1.59`

## 1) วัตถุประสงค์ของไฟล์
`Main_Program.py` คือโปรแกรม Launcher หลัก (PyQt6) สำหรับรวมเครื่องมือภายใน โดยมีหน้าที่หลักดังนี้
- แสดงการ์ดโปรแกรม แยกตามหมวดหมู่
- เปิดโปรแกรมย่อยจากโฟลเดอร์ `All_Programs`
- ตรวจสอบอัปเดตจาก GitHub Releases
- บันทึกการใช้งานไปยัง Google Apps Script
- ส่งแจ้งเตือนไป Telegram เมื่อจบการใช้งาน

## 2) ลำดับการทำงานภาพรวม
1. ขั้นเตรียมระบบตอนเริ่มไฟล์
- ถ้ารันแบบไฟล์ `.exe` จะพยายามตั้งค่า `SPSS_HOME` สำหรับ `savReaderWriter`

2. โหมด Fast-path (`--run-module`)
- `_fast_launch_submodule()` จะทำงานก่อน import หนัก
- ถ้ามี `--run-module` จะ import เฉพาะโมดูลเป้าหมายและรัน entry point ทันที
- ช่วยให้เปิดโปรแกรมย่อยได้เร็วโดยไม่ต้องโหลด UI ทั้งหมด

3. โหมด Launcher ปกติ
- โหลด PyQt6 และสร้างหน้าต่าง `AppLauncher`
- แสดง sidebar + ค้นหา + การ์ดโปรแกรม
- แสดง changelog ถ้ามีไฟล์ `changelog.tmp`
- ตั้งเวลาเช็กอัปเดตหลังเปิดโปรแกรม 15 วินาที (โหมดแจ้งเตือนอย่างเดียว ไม่บังคับอัปเดตทันที)

## 3) ค่าคงที่สำคัญ
ค่าตั้งค่าหลักด้านบนไฟล์
- `CURRENT_VERSION`
- `REPO_OWNER`, `REPO_NAME`
- `PROGRAM_SUBFOLDER` (ปกติคือ `All_Programs`)
- `ICON_FOLDER` (ปกติคือ `Icon`)
- `GOOGLE_SCRIPT_URL`
- `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID`, `TELEGRAM_DASHBOARD_URL`
- `UPDATE_HISTORY_URL`

ค่าด้าน UI/Layout
- `THEME_LIGHT`, `THEME_DARK`, `DEFAULT_APPEARANCE_MODE`
- `ICON_SIZE`, `MAX_COLUMNS`, `CARD_MIN_WIDTH`, `CARD_MAX_WIDTH`, `CARD_HEIGHT`

## 4) โครงสร้างรายการโปรแกรม (`PROGRAMS`)
`PROGRAMS` คือแหล่งข้อมูลหลักของการ์ดทั้งหมดใน launcher
ฟิลด์ที่ใช้บ่อยในแต่ละรายการ
- `id`, `name`, `description`
- `type` (`local_py_module` หรือ `external_exe`)
- `module_path` (ชื่อโมดูลใน `All_Programs`)
- `entry_point` (มักใช้ `run_this_app`)
- `icon`, `category`, `enabled`

ถ้าการ์ดไม่ขึ้น ให้เช็กตามลำดับ
- `enabled` เป็น `True` หรือไม่
- หมวดหมู่ตรงกับที่เลือกหรือไม่
- มีผลจากคำค้นหาในช่อง Search หรือไม่
- ไฟล์ไอคอนมีอยู่จริงหรือไม่
- `module_path` และ `entry_point` ถูกต้องหรือไม่

## 5) ส่วนประกอบหลักและหน้าที่
### ฟังก์ชัน helper ระดับบน
- `resource_path(relative_path)` หา path ที่ถูกต้องทั้งตอนรันจาก source และตอนรันแบบ exe
- `show_message`, `ask_yes_no`, `show_error_dialog` สำหรับ dialog มาตรฐาน
- `parse_module_launch_args(argv)` แปลงค่า `--run-module`, `--entry-point`, `--working-dir`
- `run_module_entrypoint(...)` import และเรียก entry point พร้อม error handling

### ฟังก์ชันเกี่ยวกับอัปเดต
- `check_for_updates(app_window, notify_only=False)`
  - เรียก GitHub API หา release ล่าสุด
  - เทียบเวอร์ชันด้วย `packaging.version.parse`
  - ถ้า `notify_only=True`: อัปเดตเฉพาะสถานะที่แถบล่างซ้ายของ UI
  - ถ้า `notify_only=False`: หา asset ที่ต้องใช้ เช่น `updater.exe`, patch/full package แล้วเริ่มกระบวนการอัปเดต
- `_build_patch_chain(...)` ใช้สร้างลำดับ patch แบบ incremental
- `create_custom_changelog_window(...)` และ `show_changelog_if_exists(...)` สำหรับแสดงบันทึกการอัปเดต

### คลาสหลักด้าน UI
- `Spinner` แสดงแอนิเมชันโหลดตอนกำลังเปิดโปรแกรม
- `AppLauncher`
  - สร้าง sidebar/content
  - กรองตามหมวดหมู่/คำค้นหา
  - แสดงการ์ดโปรแกรม
  - มีแถบสถานะล่างซ้ายสำหรับอัปเดต (`สถานะอัปเดต`, ปุ่ม `อัปเดตตอนนี้`, `ภายหลัง`)
  - เปิดโปรแกรมย่อย
  - เฝ้าดูสถานะ process
  - บันทึก session + ส่ง Telegram เมื่อโปรแกรมย่อยปิด

## 6) ลำดับการเปิดโปรแกรมย่อย
กรณี `local_py_module`
1. แสดงหน้าต่างกำลังเปิดโปรแกรม
2. สร้าง subprocess ด้วยคำสั่งประมาณ
- `python Main_Program.py --run-module <module> --entry-point <entry>` (ตอน dev)
- `<exe> --run-module ...` (ตอน build แล้ว)
3. ส่ง `--working-dir` หากมี
4. ปิด overlay เมื่อถือว่าโปรแกรมพร้อม
5. เริ่ม thread เฝ้ารอจน process จบ
6. เมื่อจบ: คำนวณเวลา -> log ไป Google Script -> ส่ง Telegram

กรณี `external_exe`
- ใช้ `subprocess.Popen(command, shell=True, cwd=launcher_base_dir)`

## 7) ระบบภายนอกที่เชื่อมต่อ
1. GitHub Releases API
- ล่าสุด: `/repos/{owner}/{repo}/releases/latest`
- ประวัติทั้งหมด: `/repos/{owner}/{repo}/releases`

2. Google Apps Script
- รับข้อมูลการใช้งาน: วันที่, เวลาเริ่ม/จบ, ระยะเวลา, ชื่อโปรแกรม, ผู้ใช้

3. Telegram Bot API
- ส่งข้อความแจ้งเตือนแบบ HTML
- มีการคุมความถี่ด้วย `TELEGRAM_MIN_INTERVAL_SECONDS`
- มี retry สำหรับกรณีโดน rate limit (429)

## 8) ผลข้างเคียงตอนเริ่มโปรแกรมที่ควรรู้
ใน `if __name__ == "__main__":` โหมดปกติ จะมีการ
- ตรวจ/สร้างโฟลเดอร์ `Icon/` และ `All_Programs/`
- สร้าง `All_Programs/__init__.py` ถ้ายังไม่มี
- อาจสร้าง dummy modules บางไฟล์ถ้ายังไม่พบ
- ตั้งค่า High DPI
- สร้างและแสดงหน้าต่างหลักของ Qt

หมายเหตุ: การสร้าง dummy file มีประโยชน์ตอนทดสอบ แต่ควรระวังในการใช้งานจริง

## 9) จุดเสี่ยงและเช็กลิสต์บำรุงรักษา
1. การจัดการความลับ
- token/chat id ถูกเขียนคงที่ใน source
- ควรพิจารณาย้ายไปใช้ environment variables

2. ความปลอดภัยกระบวนการอัปเดต
- ยังไม่มีการตรวจ hash/signature ของไฟล์ที่ดาวน์โหลด
- ควรเพิ่มการตรวจสอบความถูกต้องของไฟล์

3. การใช้ `shell=True`
- ใช้ได้ถ้าคำสั่งเป็นค่าคงที่ที่เชื่อถือได้
- มีความเสี่ยงถ้ามีข้อมูล input ที่ไม่ปลอดภัย

4. ตรรกะรอโปรแกรมพร้อม
- ตอนนี้ใช้เงื่อนไขเวลาร่วมกับสถานะ process
- ควรนึกถึงจุดนี้เมื่อ debug overlay ค้าง/ปิดเร็วเกิน

5. Encoding ของไฟล์
- คอมเมนต์มีทั้งไทยและอังกฤษ
- ควรรักษา encoding ให้เป็น UTF-8 สม่ำเสมอ

## 10) วิธีเพิ่มโปรแกรมใหม่อย่างปลอดภัย
1. เพิ่มไฟล์โมดูลใน `All_Programs/`
2. สร้างฟังก์ชัน entry เช่น `run_this_app(working_dir=None)`
3. เพิ่มไฟล์ไอคอนใน `Icon/`
4. เพิ่มรายการใน `PROGRAMS` ให้ครบ (`module_path`, `entry_point`, `category`, `enabled=True`)
5. ทดสอบ
- การ์ดขึ้นถูกต้อง
- ไอคอนโหลดได้
- โปรแกรมเปิดได้
- ไม่มี import/entry-point error
- ระบบ logging ทำงาน

## 11) Dependencies ที่ไฟล์นี้ใช้งาน
- `PyQt6`
- `requests` (import แบบ lazy ในฟังก์ชันที่ใช้เน็ต)
- `packaging` (เทียบเวอร์ชัน)
- standard library: `os`, `sys`, `argparse`, `importlib`, `subprocess`, `threading`, `multiprocessing`, `datetime`, `socket`, `time`, `getpass`

## 12) โครงสร้างไฟล์/โฟลเดอร์ที่ launcher คาดหวัง
- `Main_Program.py`
- `Icon/` (รวม `I_Main.ico` และไอคอนย่อย)
- `All_Programs/` (รวมโมดูลย่อยและ `__init__.py`)
- ไฟล์ที่อาจถูกสร้างระหว่างอัปเดต: `updater.exe`, `changelog.tmp`, `update_debug.log`

## 13) หมายเหตุสภาพแวดล้อมพัฒนา (อัปเดตล่าสุด)
- ใช้ `venv` เป็น environment หลักของโปรเจกต์
- `.vscode/settings.json` ตั้ง `python.defaultInterpreterPath` ไปที่ `${workspaceFolder}\venv\Scripts\python.exe`
- มี `.venv` แบบ junction ชี้ไป `venv` เพื่อรองรับคำสั่งเดิมที่ยังอ้าง `.venv`
