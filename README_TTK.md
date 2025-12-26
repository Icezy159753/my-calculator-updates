# Dashboard Program - ttkbootstrap Version

## การติดตั้ง

ติดตั้ง library ที่จำเป็น:

```bash
pip install ttkbootstrap pillow requests packaging
```

## การใช้งาน

รันโปรแกรมด้วยคำสั่ง:

```bash
python Main_Program_TTK.py
```

## ฟีเจอร์ที่ปรับปรุง

### 1. **UI Framework เปลี่ยนเป็น ttkbootstrap**
   - ใช้ ttkbootstrap แทน CustomTkinter
   - หน้าตาสวยงาม ทันสมัย
   - มี Theme มากกว่า 8 แบบให้เลือก

### 2. **Theme ที่รองรับ**
   - **cyborg** (Dark theme - Default)
   - **darkly** (Dark theme)
   - **superhero** (Dark theme)
   - **solar** (Dark theme)
   - **cosmo** (Light theme)
   - **flatly** (Light theme)
   - **journal** (Light theme)
   - **litera** (Light theme)

### 3. **ฟีเจอร์ทั้งหมดทำงานเหมือนเดิม**
   ✅ เปิดโปรแกรมย่อยได้
   ✅ ค้นหาโปรแกรม (Search with Debounce)
   ✅ กรองตามหมวดหมู่
   ✅ ตรวจสอบอัปเดต (Auto Update)
   ✅ แสดง Changelog
   ✅ บันทึกการใช้งานไปยัง Google Sheets
   ✅ เปลี่ยน Theme ได้
   ✅ ประวัติการอัปเดต
   ✅ เปิด VDO การใช้งาน

### 4. **การปรับปรุง Performance**
   - Search Debounce (300ms)
   - Icon Caching
   - Optimized Widget Creation
   - ScrolledFrame สำหรับ Sidebar

### 5. **Layout ใหม่**
   - ใช้ PanedWindow สำหรับ Sidebar + Content
   - ScrolledFrame สำหรับแสดงโปรแกรม
   - Grid Layout 4 คอลัมน์
   - Responsive Design

## ความแตกต่างจาก CustomTkinter

| ฟีเจอร์ | CustomTkinter | ttkbootstrap |
|---------|---------------|--------------|
| Base Library | Custom Widget | tkinter.ttk |
| Theme | 3 แบบ | 8+ แบบ |
| Performance | ดี | ดีกว่า |
| File Size | ใหญ่กว่า | เล็กกว่า |
| Styling | Limited | Bootstrap-like |
| Dependencies | มาก | น้อย |

## โครงสร้างไฟล์

```
Main_Program/
├── Main_Program_TTK.py     # โปรแกรมหลัก (ttkbootstrap)
├── Main_Program.py         # โปรแกรมเดิม (CustomTkinter)
├── Icon/                   # โฟลเดอร์ไอคอน
│   ├── I_Main.ico
│   ├── Peth.ico
│   ├── T2B.ico
│   └── ...
└── All_Programs/           # โฟลเดอร์โปรแกรมย่อย
    ├── __init__.py
    ├── Program_ItemdefSPSS_Log.py
    ├── Program_T2B_Itermdef.py
    └── ...
```

## การเปลี่ยน Theme

1. คลิกที่ Dropdown "⚙️ ธีม:" ใน Sidebar
2. เลือก Theme ที่ต้องการ
3. Theme จะเปลี่ยนทันที

## Tips & Tricks

### เพิ่มโปรแกรมใหม่
แก้ไข list `PROGRAMS` ในไฟล์ Main_Program_TTK.py:

```python
{
    "id": "program_id",
    "name": "ชื่อโปรแกรม",
    "description": "คำอธิบาย",
    "type": "local_py_module",
    "module_path": "ชื่อไฟล์_โดยไม่มี.py",
    "entry_point": "run_this_app",
    "icon": "icon.ico",
    "category": "หมวดหมู่",
    "enabled": True
}
```

### เปลี่ยน Default Theme
แก้ไขบรรทัด:
```python
super().__init__(themename="cyborg")  # เปลี่ยนเป็น theme ที่ต้องการ
```

### ปรับขนาดไอคอน
แก้ไขค่า `ICON_SIZE`:
```python
ICON_SIZE = (60, 60)  # (width, height)
```

### ปรับจำนวนคอลัมน์
แก้ไขในฟังก์ชัน `create_program_widgets`:
```python
max_cols = 4  # เปลี่ยนเป็น 3, 5, 6, etc.
```

## Troubleshooting

### ปัญหา: โปรแกรมไม่เปิด
**แก้ไข:** ตรวจสอบว่าติดตั้ง ttkbootstrap แล้ว
```bash
pip install ttkbootstrap
```

### ปัญหา: ไอคอนไม่แสดง
**แก้ไข:** ตรวจสอบว่าไฟล์ .ico อยู่ในโฟลเดอร์ `Icon/`

### ปัญหา: Font ไม่สวย
**แก้ไข:** ติดตั้ง Font "TH Sarabun New" (สำหรับภาษาไทย)

### ปัญหา: Theme ไม่เปลี่ยน
**แก้ไข:** ปิดแล้วเปิดโปรแกรมใหม่

## Known Issues

1. บางครั้ง Sidebar อาจมี scrollbar ไม่แสดง (แก้โดยรีไซส์หน้าต่าง)
2. Icon จาก .ico file บางไฟล์อาจมี background สีขาว
3. Font Thai ต้องติดตั้ง "TH Sarabun New" ก่อน

## Credits

- **Original Author:** DP Team
- **ttkbootstrap Version:** Converted from CustomTkinter
- **Framework:** ttkbootstrap by Israel Dryer

## License

MIT License - ใช้ได้ฟรี สำหรับทุกโปรเจค
