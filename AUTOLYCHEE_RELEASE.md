# Auto Lychee — เช็คลิสต์ก่อน Push ขึ้น GitHub

> สำหรับ AI: อ่านไฟล์นี้ทุกครั้งที่มีการแก้ `All_Programs/158_AutoLychee_OneFile.py`
> หรือก่อน push / bump version ที่มีการเปลี่ยนไฟล์ Auto Lychee
> ทำตามทุกข้อ ถ้าข้อไหนไม่ผ่าน **ห้าม push** ให้แจ้งผู้ใช้ก่อน (ตอบผู้ใช้เป็นภาษาไทย)

## 1. ไฟล์มาจากไหน

- ต้นทางอยู่ที่ `C:\Users\songklod\Desktop\All Python\6_RunLychee\` (ไม่ได้อยู่ใน git)
- `make_onefile.py` ในโฟลเดอร์นั้นสร้าง `AutoLychee_OneFile.py`
  ส่วน loader ด้านบนของไฟล์ copy มาจากตัวแปร `LOADER` ใน `make_onefile.py`
- ผู้ใช้ copy ไฟล์นั้นมาวางทับเป็น `All_Programs/158_AutoLychee_OneFile.py`
- ถ้าต้องเพิ่มของใน loader ให้แก้ที่ `LOADER` ใน `make_onefile.py` ด้วย ไม่อย่างนั้นการวางทับครั้งหน้าจะลบของที่เพิ่มไปทิ้ง

## 2. สิ่งที่ต้องมีในไฟล์ 158 (ห้ามหาย)

ใน `_main()` ของ loader ต้องมีครบทุกข้อ:

| คำสั่ง | ต้องทำอะไร | ใครเรียก |
|---|---|---|
| `--smoke-test` | `_smoke_test()`: โหลดทุก section แล้วสร้าง GUI (`App()`) แล้วปิดทันที พร้อมพิมพ์ `Auto Lychee GUI smoke test OK` (`if FROZEN: _bind_stdio()` ก่อน) | CI ขั้น "Verify packaged Auto Lychee EXE" |
| `--check` | `if FROZEN: _bind_stdio()` แล้ว `_check()` | CI ผ่าน `Main_Program.exe --run-module ... --check` |
| `--post` | `_bind_stdio(stdin=True)` แล้ว `worker.post_main()` อ่าน stdin จนจบแล้ว exit 0 | worker / CI |
| `--worker <json>` | `_bind_stdio()` แล้ว `worker.main()` | GUI |

**บทเรียนจาก V1.1.92:** ไฟล์ใหม่ที่วางทับไม่มี `--smoke-test` ทำให้ exe เปิด GUI เต็มและไม่ปิดเอง
CI ค้าง 6 ชั่วโมงจนถูกยกเลิก แก้ไว้ใน V1.1.93

**ตอน V1.1.94 ไฟล์ที่วางใหม่ก็ไม่มีอีก** (`make_onefile.py` ต้นทางยังไม่ได้แก้)
จนกว่าต้นทางจะแก้ ให้ถือว่าทุกครั้งที่ผู้ใช้วางไฟล์ใหม่ `_smoke_test` จะหายไป
ให้เช็คด้วย `git diff All_Programs/158_AutoLychee_OneFile.py` แล้วใส่โค้ดด้านล่างกลับเข้าไปเองทุกครั้ง

โค้ดที่ต้องมี (วางก่อน `def _main()`):

```python
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
```

และใน `_main()`:

```python
    elif args[:1] == ['--check']:
        if FROZEN:
            _bind_stdio()
        _check()
    elif args == ['--smoke-test']:  # used by the Main Program release build (CI)
        if FROZEN:
            _bind_stdio()
        _smoke_test()
```

ถ้า section `app` เปลี่ยนชื่อคลาสหน้าต่างหลัก (`App`) หรือเลิกใช้ `QApplication` ให้แก้ `_smoke_test` ให้ตรง

## 3. สิ่งที่ห้ามเปลี่ยน (Main_Program เรียก Auto Lychee แบบนี้)

- `Main_Program.py` ใน `PROGRAMS`: `module_path` คือ `158_AutoLychee_OneFile`, `entry_point` คือ `"__main__"`
  และ `frozen_executable` คือ `"AutoLychee/AutoLychee.exe"`
- ตอนเป็น exe, `_fast_launch_submodule()` และ `_start_local_module_process()` เปิด
  `_internal/AutoLychee/AutoLychee.exe` เป็นโปรเซสแยก ห้าม import ไฟล์ 158 ตรงๆ
  เพราะไฟล์ 158 ใช้ PySide6 แต่ Main_Program ใช้ PyQt6 และไฟล์จะ `raise ImportError` เมื่อถูก import
- `AutoLychee.spec` build แบบ onedir แยก (PySide6) เพื่อลดเวลาที่เสียไปกับการแตกไฟล์ทุกครั้งที่เปิด
  และ `Main_Program.spec` ต้องมี `dist/AutoLychee/AutoLychee.exe` ก่อน พร้อมคัดลอกทั้งโฟลเดอร์รวม `_internal`
  ลำดับการ build: `AutoLychee.spec` ก่อน แล้วจึง `Main_Program.spec`

## 4. ถ้าไฟล์ใหม่มี section หรือ package ใหม่

- เช็ค section: `grep -n "^# ====== MODULE:" All_Programs/158_AutoLychee_OneFile.py`
- ถ้ามี section ใหม่ ให้เพิ่มชื่อใน `excludes` ของ `AutoLychee.spec`
  ตอนนี้มี: core, fast_styles, history, chrome, driver, sounds, post_total_na, post_del_sig,
  post_cut_percent, worker, app, onefile_assets
- ถ้ามี package ภายนอกใหม่ ให้เพิ่มใน `hiddenimports` ของ `AutoLychee.spec`
  และในขั้น "Install dependencies" ของ `.github/workflows/windows-release.yml`

## 5. ทดสอบก่อน push (ต้องผ่านทุกข้อ)

รันจากโฟลเดอร์ `Main_Program` (Git Bash):

```bash
python -m unittest test_autolychee_packaging test_update_cache test_release_files test_release_build -v
export QT_QPA_PLATFORM=offscreen AUTOLYCHEE_DATA="$TEMP/al-smoke"
timeout 120 python -X utf8 All_Programs/158_AutoLychee_OneFile.py --smoke-test   # ต้องเห็น "Auto Lychee GUI smoke test OK"
timeout 60  python -X utf8 All_Programs/158_AutoLychee_OneFile.py --post </dev/null  # exit 0
timeout 120 python -X utf8 Main_Program.py --run-module 158_AutoLychee_OneFile --entry-point run_this_app --check  # exit 0
```

ใส่ `timeout` ทุกครั้ง ถ้าคำสั่งไหนค้างจนหมดเวลา แปลว่าเสียแล้ว (มักเป็นเพราะ `--smoke-test` หาย)

ถ้ามีเวลา ให้ build exe ในเครื่องแล้วลองแบบเดียวกับ CI:

```bash
pyinstaller --noconfirm AutoLychee.spec
pyinstaller --noconfirm Main_Program.spec
QT_QPA_PLATFORM=offscreen timeout 120 dist/Main_Program/_internal/AutoLychee/AutoLychee.exe --smoke-test
```

## 6. ขั้นตอนขึ้น GitHub

**กฎเลขเวอร์ชัน (ผู้ใช้สั่งไว้):** ทุกครั้งที่ผู้ใช้สั่ง "เอาขึ้น GitHub" ให้เช็ค `CURRENT_VERSION` ใน `Main_Program.py` (ประมาณบรรทัด 202)
แล้วเทียบกับ tag ล่าสุด (`git tag -l "v1.1.*" --sort=-v:refname | head -1`)
ถ้าเลขยังเหมือนเดิม หรือมี tag เลขนั้นอยู่แล้ว ให้เพิ่มเลขท้ายขึ้น 1 เอง (เช่น `1.1.94` เป็น `1.1.95`) ไม่ต้องถามผู้ใช้
และแก้เลขในไฟล์ `คำสั่ง Build exe Update ขึ้น Github.txt` ให้ตรงด้วย
ถ้าไม่ขึ้นเลขใหม่ GitHub จะไม่ build exe (auto-tag ทำงานเฉพาะเมื่อ `CURRENT_VERSION` เปลี่ยน)

1. แก้ `CURRENT_VERSION` ใน `Main_Program.py` เป็นเลขใหม่ตามกฎด้านบน (ห้ามใช้เลขที่มี tag อยู่แล้ว)
2. Commit แล้ว push:
   ```bash
   git add <ไฟล์ที่เปลี่ยน>
   git commit -m "Bump version to 1.1.XX"
   git push origin main
   git tag v1.1.XX
   git push origin v1.1.XX
   ```
3. ห้ามย้ายหรือลบ tag เดิม ถ้า build ของเลขไหนพัง ให้ขึ้นเลขถัดไป
4. ดูผล build: `gh run list -L 3 --workflow windows-release.yml` และ `gh run watch <run-id> --exit-status`
   build ปกติใช้เวลาประมาณ 12–15 นาที ถ้าเกิน 30 นาทีแปลว่ามีปัญหา
   (CI ตั้งให้ขั้น verify หมดเวลาที่ 10 นาที และทั้งงานหมดเวลาที่ 60 นาที)
5. ถ้า build fail: `gh run view <run-id> --log-failed` แล้วสรุปสาเหตุให้ผู้ใช้เป็นภาษาไทย

## 7. ระบบอัปเดตไฟล์เฉพาะที่เปลี่ยน (เริ่ม v1.1.99)

- `release_build.py prepare` ดาวน์โหลดชุดเต็มล่าสุดและตรวจ checksum กับ `release_manifest.json`
  ใช้ fingerprint ของ source/spec/dependencies เพื่อตัดสินใจ build Main, AutoLychee และ updater แยกกัน
  หากรีลีสก่อนยังไม่มี manifest จะ build ใหม่ทั้งหมด; ถ้า checksum ไม่ตรงจะหยุดการเผยแพร่
- Main อ่านเวอร์ชันจาก `_internal/release_version.json` เมื่อเป็น EXE จึง reuse EXE ได้เมื่อเปลี่ยนเลขเวอร์ชันอย่างเดียว
  ห้ามเอาไฟล์นี้ออกจากแพ็กเกจ แม้ source `CURRENT_VERSION` ถูก bump แล้วก็ตาม
- แก้โปรแกรมใน `All_Programs` โดย imports เดิม: reuse runtime และเปลี่ยนไฟล์ source ที่แพ็กไว้
  เพิ่ม imports/เปลี่ยน requirements/spec: rebuild ส่วนที่เกี่ยวข้อง ต้องเพิ่ม hiddenimports/dependencies ให้ครบตามเดิม
- AutoLychee เปลี่ยน: build AutoLychee แล้วแทนโฟลเดอร์ `_internal/AutoLychee` ทั้งชุด; Main ไม่จำเป็นต้อง rebuild
- `Main_Program_files_<from>_to_<to>.zip` มี manifest และไฟล์เปลี่ยนพร้อม SHA-256
  updater ใหม่รองรับ arguments เดิม และค้นแพ็กเกจนี้ได้เองแม้ Main บนเครื่องผู้ใช้ยังเป็นรุ่นเก่า
  ไม่ต้องมี cache ZIP ฐาน; เก็บแพ็กเกจตรงจากฐานรุ่นแรกและสองฐานล่าสุดสำหรับเครื่องที่ข้ามเวอร์ชัน
  ถ้าไม่มีแพ็กเกจสำหรับฐานที่ติดตั้งหรือไฟล์ฐานไม่ตรงจะใช้ Full Package
- ก่อนแทนไฟล์ ต้องตรวจ checksum/เส้นทางทั้งหมดและสำรองไฟล์ก่อนเปลี่ยน
  journal อยู่ที่ `_internal/update-transactions`; ถ้า rollback ไม่สำเร็จห้ามลบ backup
  updater ครั้งถัดไปต้องกู้คืน transaction ก่อนเริ่มอัปเดตใหม่
  Main รุ่นใหม่ตรวจ journal ก่อนโหลด GUI และเรียก `updater.exe --recover-only` เพื่อคืนไฟล์
  ถ้า Main เปิดไม่ได้ ให้ดับเบิลคลิก updater.exe ในโฟลเดอร์ติดตั้งเพื่อกู้คืนจาก backup
- ไม่ลบไฟล์งานนอกแพ็กเกจ และรักษา Test3.json, openrouter.json, Itemdef - Format.xlsx ที่มีอยู่
- CI ตรวจ source ของโปรแกรมทั้ง 91 ไฟล์, smoke test EXE, updater EXE self-test และทดลองอัปจาก full release ก่อนหน้า
  ต้องได้ checksum ของ managed files เหมือน full release ใหม่ก่อนอัปโหลด
- ไม่มี bsdiff ของ ZIP ทั้งชุดในรีลีสใหม่; Full ZIP และ updater.exe ยังต้องมีเพื่อรองรับเครื่องเดิม
- `workflow_dispatch` ใช้สำหรับ preview เท่านั้น ไม่เผยแพร่ release หรือส่ง Telegram
  ตัวอย่าง: `gh workflow run windows-release.yml --ref main -f preview_version=1.1.101`
  ให้ใช้เลขมากกว่ารีลีสล่าสุด แล้วตรวจขั้น Select build components และเวลางานเพื่อยืนยันความเร็วจริง

- ตั้งแต่ v1.1.100 Main reuse updater.exe เมื่อ SHA-256 ตรงกับ asset digest ของ GitHub
  ดาวน์โหลดลงไฟล์ชั่วคราวและตรวจ EXE/checksum ก่อนแทนไฟล์เดิม
