# New Tools — แผนย้ายโปรแกรม Desktop (PyQt6/Tkinter) ขึ้นเว็บ React + FastAPI

> เอกสารวางแผน **ก่อนลงมือ** สำหรับโปรเจกต์ใหม่ที่จะย้ายโปรแกรมย่อย ~35–40 ตัว
> จาก Desktop App (`Main_Program.py` + `All_Programs/`) ขึ้นเป็น **Web App**
> เป้าหมาย: เลิกแจก `.exe` / เลิก auto-update, ใช้ร่วมกันหลายคน, UI สวยทันสมัย
>
> **กลยุทธ์หลัก: ย้ายทีละตัว แล้วทดสอบว่าทำงานได้ "เหมือนเดิม" ก่อนไปตัวถัดไป**

---

## 1. เป้าหมายและขอบเขต

| หัวข้อ | รายละเอียด |
|--------|------------|
| เป้าหมาย | เว็บแอปรวมเครื่องมือ DP/วิจัย ใช้ผ่าน browser ไม่ต้องลงโปรแกรม |
| ผู้ใช้ | หลายคนในทีม (ต้องมี login) |
| Deploy | Frontend → Vercel, Backend → AWS/Render/Railway, ไฟล์ → S3 |
| แนวทาง | React + FastAPI เต็มรูปแบบ |
| วิธีทำ | ย้ายทีละโปรแกรม → ทดสอบเทียบผลกับตัวเดิม → ค่อยทำตัวถัดไป |

### ข่าวดีจากการสแกนโค้ดเดิม
- โปรแกรม SPSS **เกือบทั้งหมดใช้ `pyreadstat` อยู่แล้ว** → ขึ้น Linux/Cloud ได้ทันที
- มีแค่ **`convert_SPSS_UTF8.py`** ที่ยังใช้ `savReaderWriter` (ผูก DLL Windows) → ต้องแก้เป็น `pyreadstat`
- ทุกโปรแกรมมี entry point มาตรฐานชื่อ `run_this_app` → ใช้เป็นจุดอ้างอิงตอนแยก logic

### สิ่งที่ "ต้องเขียนใหม่" vs "ใช้ซ้ำได้"
| ส่วน | สถานะ |
|------|-------|
| Logic ประมวลผล (pandas/openpyxl/pyreadstat/คำนวณสถิติ) | ✅ ใช้ซ้ำได้ ~80–90% |
| UI (tkinter/customtkinter/PyQt6, filedialog, messagebox) | ❌ เขียนใหม่เป็น React ทั้งหมด |
| การเลือกไฟล์ในเครื่อง | ❌ เปลี่ยนเป็น upload/download ผ่าน HTTP |

---

## 2. สถาปัตยกรรมเป้าหมาย

```
┌──────────────────┐      HTTPS/REST      ┌─────────────────────┐
│  React (Vercel)  │  ◄───────────────►   │  FastAPI (server)   │
│  - หน้า list tool │                      │  - 1 router / tool  │
│  - ฟอร์ม upload   │                      │  - เรียก service     │
│  - progress bar  │                      │  - คืน job/ไฟล์ผล    │
│  - ปุ่ม download   │                      └─────────┬───────────┘
└──────────────────┘                                │
        ▲                                           ▼
        │ presigned URL                  ┌─────────────────────┐
        └────────────────────────────────│  S3 (ไฟล์ชั่วคราว)   │
                                          │  upload + auto-expire│
                                          └─────────────────────┘
        Auth: Auth0 / Clerk / Cognito (JWT)
        งานหนัก/นาน: Celery + Redis (หรือ BackgroundTasks ช่วงแรก)
```

**กฎเหล็ก:** ไฟล์วิจัยเป็นความลับ → เก็บบน S3 แบบ **มีวันหมดอายุ + ลบหลังประมวลผลเสร็จ** + แยกตาม user + ทุก endpoint ต้องผ่าน auth

---

## 3. Tech Stack

| ชั้น | เครื่องมือ | หมายเหตุ |
|------|-----------|----------|
| Frontend | React + Vite + TypeScript + Tailwind + shadcn/ui | UI ทันสมัย, component สำเร็จรูป |
| State/API | TanStack Query (React Query) + axios | จัดการ async/upload/polling |
| Backend | FastAPI + Uvicorn + Pydantic | Python ล้วน เอา logic เดิมมาต่อ |
| งานหนัก | Celery + Redis (เริ่มด้วย BackgroundTasks ก่อนได้) | กัน request timeout |
| ไฟล์ | AWS S3 + boto3 (presigned URL) | ไม่เก็บไฟล์ค้างบน server |
| SPSS | pyreadstat | แทน savReaderWriter |
| Excel | openpyxl / pandas | เหมือนเดิม |
| Auth | Auth0 / Clerk / Cognito | ไม่ต้องเขียน auth เอง |
| Deploy FE | Vercel | |
| Deploy BE | Render / Railway / AWS ECS (Docker) | Vercel รัน backend หนักไม่ได้ |
| Container | Docker + docker-compose (dev) | reproducible |

---

## 4. โครงสร้างโปรเจกต์ (Monorepo)

```
new-tools/
├── New Tools.md                  # เอกสารนี้
├── docker-compose.yml            # dev: api + redis + worker + minio(S3 จำลอง)
├── README.md
│
├── backend/
│   ├── app/
│   │   ├── main.py               # สร้าง FastAPI app, รวม routers
│   │   ├── config.py             # settings (env: S3, redis, auth)
│   │   ├── deps.py               # auth dependency, get_current_user
│   │   │
│   │   ├── core/                 # โครงกลางที่ใช้ซ้ำทุก tool
│   │   │   ├── storage.py        # upload/download/presigned S3 + auto-expire
│   │   │   ├── jobs.py           # ระบบ job: create/status/result
│   │   │   ├── spss.py           # helper อ่าน/เขียน .sav ด้วย pyreadstat
│   │   │   └── excel.py          # helper openpyxl ที่ใช้ร่วมกัน
│   │   │
│   │   ├── routers/              # 1 ไฟล์ / 1 โปรแกรม (thin layer)
│   │   │   ├── rename_sheet.py
│   │   │   ├── move_sheet.py
│   │   │   ├── bpi.py
│   │   │   └── ...
│   │   │
│   │   ├── services/             # ★ Logic บริสุทธิ์ (พอร์ตจาก All_Programs)
│   │   │   ├── rename_sheet.py   # def run(input_path, params) -> output_path
│   │   │   ├── move_sheet.py
│   │   │   ├── bpi.py
│   │   │   └── ...
│   │   │
│   │   ├── schemas/              # Pydantic models (request/response ต่อ tool)
│   │   └── registry.py           # ทะเบียนรวม tool ทั้งหมด (id/ชื่อ/หมวด/route)
│   │
│   ├── legacy/                   # ★ โค้ด Desktop เดิม เก็บไว้อ้างอิง (ห้ามแก้)
│   │   ├── All_Programs/         # copy ทั้งโฟลเดอร์มาไว้ดูเทียบ
│   │   └── README.md             # ที่มา + วิธีรันตัวเดิมไว้เทียบผล
│   │
│   ├── tests/
│   │   ├── fixtures/             # ไฟล์ตัวอย่าง .sav/.xlsx + ผลลัพธ์ที่ถูกต้อง
│   │   └── test_<tool>.py        # เทียบ output ใหม่ == output เดิม
│   │
│   ├── requirements.txt
│   └── Dockerfile
│
├── frontend/
│   ├── src/
│   │   ├── main.tsx
│   │   ├── App.tsx
│   │   ├── lib/
│   │   │   ├── api.ts            # axios client + auth header
│   │   │   └── useJob.ts         # hook: submit → poll progress → ได้ผล
│   │   ├── components/
│   │   │   ├── FileUpload.tsx    # drag-drop upload (ใช้ซ้ำทุก tool)
│   │   │   ├── JobProgress.tsx   # progress bar / สถานะ
│   │   │   ├── ResultDownload.tsx
│   │   │   └── ToolCard.tsx
│   │   ├── pages/
│   │   │   ├── ToolList.tsx      # หน้าแรก รวม tool ตามหมวด (เหมือน launcher เดิม)
│   │   │   └── tools/            # 1 หน้า / 1 โปรแกรม
│   │   │       ├── RenameSheet.tsx
│   │   │       ├── MoveSheet.tsx
│   │   │       └── ...
│   │   └── registry.ts          # ทะเบียน tool ฝั่ง FE (id/ชื่อ/หมวด/ไอคอน)
│   ├── package.json
│   └── vite.config.ts
│
└── infra/
    ├── README.md                # วิธี deploy
    └── ...                      # IaC/สคริปต์ deploy (ทำทีหลัง)
```

> **หมายเหตุเรื่อง "เอา Python 40 ตัวมาไว้ในโปรเจกต์":**
> ก๊อปทั้งหมดไปไว้ใน `backend/legacy/All_Programs/` ก่อน — **ไม่แก้ของเดิม** ใช้เป็น "ตัวเทียบผล"
> เวลาพอร์ตแต่ละตัว ให้สกัด logic ออกมาเขียนใหม่ใน `backend/app/services/` แทน
> (ตัวเดิมยังเปิดบน Windows ได้ ไว้รันเทียบว่าผลตรงกัน)

---

## 5. Pattern การพอร์ต 1 โปรแกรม (ทำซ้ำทุกตัว)

แต่ละโปรแกรมทำตาม 6 ขั้นนี้ ถือว่า "เสร็จ" เมื่อผลลัพธ์ตรงกับตัวเดิม:

1. **อ่านโค้ดเดิม** ใน `legacy/All_Programs/<file>.py` — หา 3 อย่าง:
   - input: รับไฟล์/พารามิเตอร์อะไรบ้าง (จาก `filedialog`, ช่องกรอก, checkbox)
   - process: ฟังก์ชันคำนวณหลักอยู่ตรงไหน
   - output: เขียนไฟล์อะไรออก / แสดงผลยังไง

2. **สกัด Logic** → `backend/app/services/<tool>.py`
   ```python
   def run(input_paths: list[str], params: dict) -> str:
       # ยกโค้ดคำนวณเดิมมา ตัดส่วน GUI ออก
       # return path ไฟล์ผลลัพธ์
   ```
   - ลบทุกอย่างที่เป็น `tk`, `ctk`, `QtWidgets`, `messagebox`, `filedialog`
   - แทน `filedialog.askopenfilename()` ด้วย argument `input_paths`
   - แทน `messagebox.showinfo/showerror` ด้วย `return` / `raise ValueError`

3. **เขียน Schema** → `backend/app/schemas/<tool>.py` (Pydantic: พารามิเตอร์ที่ผู้ใช้ตั้งได้)

4. **เขียน Router** → `backend/app/routers/<tool>.py`
   - `POST /tools/<tool>` → รับไฟล์ + params → สร้าง job → เรียก service → คืน job_id
   - ใช้ helper จาก `core/jobs.py`, `core/storage.py` (ไม่เขียน upload/download ซ้ำ)

5. **เขียน UI** → `frontend/src/pages/tools/<Tool>.tsx`
   - ใช้ `<FileUpload>`, `<JobProgress>`, `<ResultDownload>` ที่มีอยู่แล้ว
   - ทำแค่ฟอร์มพารามิเตอร์เฉพาะของ tool นั้น

6. **ทดสอบเทียบผล (สำคัญสุด)**
   - เอาไฟล์ตัวอย่างจริงรันผ่าน "ตัวเดิม" (Desktop) → เก็บผลไว้เป็น expected
   - รันผ่าน "ตัวใหม่" (web) → เทียบว่าไฟล์ผลลัพธ์ตรงกัน (`tests/test_<tool>.py`)
   - ✅ ตรง → ติ๊ก done ในตาราง §7 แล้วไปตัวถัดไป

---

## 6. แผนงานเป็นเฟส (Roadmap)

### เฟส 0 — เตรียมโปรเจกต์ (ทำครั้งเดียว)
- [ ] สร้าง repo `new-tools/` + โครงโฟลเดอร์ตาม §4
- [ ] ก๊อป `All_Programs/` → `backend/legacy/All_Programs/`
- [ ] ตั้ง docker-compose (api + redis + minio สำหรับ S3 จำลองตอน dev)
- [ ] FastAPI skeleton + health check + CORS
- [ ] React skeleton (Vite + Tailwind + shadcn) + หน้า ToolList ว่างๆ

### เฟส 1 — โครงกลางที่ใช้ซ้ำ (ทำครั้งเดียว ลงแรงมากสุด)
- [ ] `core/storage.py` — upload/presigned/download S3 + auto-expire
- [ ] `core/jobs.py` — create/status/result + progress (ใช้ BackgroundTasks ก่อน)
- [ ] `core/spss.py` — wrapper pyreadstat (อ่าน/เขียน .sav + metadata)
- [ ] `core/excel.py` — helper openpyxl ที่หลายโปรแกรมใช้ร่วม
- [ ] Auth (Auth0/Clerk) + `deps.get_current_user` + ป้องกันทุก endpoint
- [ ] FE: `FileUpload`, `JobProgress`, `ResultDownload`, `useJob` hook
- [ ] FE: หน้า ToolList อ่านจาก `registry.ts` แสดงเป็นการ์ดตามหมวด (เหมือน launcher เดิม)

### เฟส 2 — นำร่อง 2 ตัว (พิสูจน์ pattern)
- [ ] **Rename Sheet** (Excel ล้วน ง่ายสุด) — ครบ loop upload→run→download
- [ ] **Get SPSS** หรือ **ConvertSPSS_Excel** (ตัวแตะ SPSS) — พิสูจน์ pyreadstat บน server
- [ ] เขียน test เทียบผลทั้ง 2 ตัว → ยืนยัน pattern ใช้ได้จริง

### เฟส 3 — ทยอยพอร์ตที่เหลือทีละกลุ่ม (งานซ้ำ เร็วขึ้น)
ไล่ตามหมวดเดิม โดยเรียงจากง่าย→ยาก (ดูตาราง §7):
- [ ] กลุ่ม Excel (ส่วนใหญ่ง่าย)
- [ ] กลุ่ม Lychee
- [ ] กลุ่ม SPSS
- [ ] กลุ่ม Statistic (ซับซ้อนสุด — มีกราฟ/โมเดล เก็บไว้ท้าย)
- [ ] กลุ่ม Diary / Key Norm / อื่นๆ

### เฟส 4 — เก็บงาน + ขึ้น production
- [ ] ย้าย job หนักไป Celery + Redis (ถ้า BackgroundTasks เริ่มไม่พอ)
- [ ] แก้ `convert_SPSS_UTF8` จาก savReaderWriter → pyreadstat
- [ ] ระบบลบไฟล์อัตโนมัติ + กำหนด quota/limit ขนาดไฟล์
- [ ] Deploy: FE→Vercel, BE→Render/AWS, ตั้ง env/secret, โดเมน, HTTPS
- [ ] เก็บ log/monitoring + หน้า admin (ถ้าต้องการ)

---

## 7. ทะเบียนโปรแกรม + ช่องติดตามความคืบหน้า

> ความซับซ้อน: 🟢 ง่าย (Excel/แปลงไฟล์ตรงๆ) · 🟡 กลาง (อ่าน SPSS + transform) · 🔴 ยาก (สถิติ/กราฟ/หลายขั้นตอน)
> สถานะ: ⬜ ยังไม่ทำ · 🔧 กำลังพอร์ต · ✅ เสร็จ+เทียบผลแล้ว

| # | โปรแกรม (ตัวเดิม) | module เดิม | หมวด | SPSS | ระดับ | สถานะ |
|---|------------------|-------------|------|:----:|:----:|:----:|
| 1 | RenameSheet V1 | `Rename Sheet` | Excel | – | 🟢 | ⬜ |
| 2 | Move Sheet Excel V1 | `107_Movesheet` | Excel | – | 🟢 | ⬜ |
| 3 | ตัดQuota Pro V1 | `99_Excel` | Excel | ✓ | 🟡 | ⬜ |
| 4 | เช็ค Data Excel 2 ไฟล์ V1 | `146_Mapdata` | Excel | – | 🟢 | ⬜ |
| 5 | ลบTotal+NA Table Lychee V1 | `150_Delete_NA_Lychee` | Excel | – | 🟢 | ⬜ |
| 6 | full Itemdef+Genpromt Beta V1 | `108_GenPromt_NewBeta` | Lychee | ✓ | 🟡 | ⬜ |
| 7 | สร้าง Itemdef จาก SPSS V3 | `Program_ItemdefSPSS_Log` | Lychee | ✓ | 🟡 | ⬜ |
| 8 | ทำ TB/T2B จาก Itemdef V3 | `Program_T2B_Itermdef` | Lychee | – | 🟡 | ⬜ |
| 9 | GetValue+Promt แปะ Eng | `117_Newen_Promt` | Lychee | ✓ | 🟡 | ⬜ |
| 10 | Logic_Generator Itemdef V8 | `Logic_Generator` | Lychee | – | 🟡 | ⬜ |
| 11 | แปลง CE Other จาก Edit V2 | `106_Map_spss_Excel` | Lychee | ✓ | 🟡 | ⬜ |
| 12 | ลบ Sig จาก TableLychee V1 | `Del_Sig` | Lychee | – | 🟢 | ⬜ |
| 13 | Check Codes_Other V1 | `CheckOther` | Lychee | ✓ | 🟡 | ⬜ |
| 14 | แตก CodeNA V1 | `113_ProgramCodeNA` | Lychee | ✓ | 🟡 | ⬜ |
| 15 | ลบ N=0 OE ใน Lychee V1 | `114_DelblankLychee` | Lychee | – | 🟢 | ⬜ |
| 16 | Create Format_Kao By_DP V3 | `119_Create_Format_Kao` | Lychee | – | 🟡 | ⬜ |
| 17 | N% To % Lychee V1 | `151_CutLychee_Persence` | Lychee | – | 🟢 | ⬜ |
| 18 | SPSS For Lychee V1 | `152_spss_converter_gui` | Lychee | ✓ | 🟡 | ⬜ |
| 19 | CleanData+Frequenzy SPSS V1 | `99_CleanSPSS_Germini` | SPSS | ✓ | 🔴 | ⬜ |
| 20 | Get SPSS V2 | `105_GetSPSS` | SPSS | ✓ | 🟡 | ⬜ |
| 21 | ซ่อมไฟล์ SPSS V1 ⚠️ | `convert_SPSS_UTF8` | SPSS | ✓* | 🟡 | ⬜ |
| 22 | แปลงไฟล์ SPSS To Excel V1 | `ConvertSPSS_Excel` | SPSS | ✓ | 🟡 | ⬜ |
| 23 | MRSET Auto-Generator v5.0 | `121_SPSS_MRSET` | SPSS | ✓ | 🟡 | ⬜ |
| 24 | รัน Correlation V1 | `104_Correlation` | Statistic | ✓ | 🔴 | ⬜ |
| 25 | BPI Brand Power Index V1 | `120_bpi` | Statistic | ✓ | 🔴 | ⬜ |
| 26 | BrandSpace Well-being Index V1 | `148_BrandSpace` | Statistic | ✓ | 🔴 | ⬜ |
| 27 | Penality Analysis V1 | `147_Penalty` | Statistic | ✓ | 🔴 | ⬜ |
| 28 | Multidimensional Scaling (MDS) V12 | `MDS` | Statistic | – | 🔴 | ⬜ |
| 29 | ดูดติด MA _O จาก Togo V1 | `121_Merge_MA_V2` | Statistic | ✓ | 🟡 | ⬜ |
| 30 | PSM Pricezen V1 | `122_PSM` | Statistic | – | 🔴 | ⬜ |
| 31 | BrandSence V1 | `123_Program_Run_Brandsence2026` | Statistic | ✓ | 🔴 | ⬜ |
| 32 | Check Rotation Diary V1 | `109_Diary` | Diary | – | 🟡 | ⬜ |
| 33 | เก็บ Norm V1 | `Norm_2025` | Key Norm | ✓ | 🟡 | ⬜ |
| 34 | Gen Table Reporter V1 | `124_Table_Reporter` | อื่นๆ | ✓ | 🔴 | ⬜ |
| 35 | Convert JPG to SVG V1 | `125_Convert_Icon_Svg2` | อื่นๆ | – | 🟢 | ⬜ |

> ⚠️ #21 `convert_SPSS_UTF8` = ตัวเดียวที่ใช้ `savReaderWriter` → ต้องแก้เป็น `pyreadstat` ตอนพอร์ต
> ✓* = SPSS แต่ใช้ไลบรารีคนละตัว

**ลำดับแนะนำให้เริ่ม:** #1 → #2 → #20 (หรือ #22) แล้วค่อยกวาด 🟢 ที่เหลือ → 🟡 → 🔴

---

## 8. ประเด็นเทคนิคที่ต้องตัดสินใจ/ระวัง

1. **ไฟล์ใหญ่ + งานนาน** — โปรแกรมสถิติ (BrandSpace/MDS/PSM/Correlation) อาจรันหลายสิบวินาที
   → ต้องเป็น async job + progress ตั้งแต่ออกแบบ ห้ามทำเป็น request ตรงๆ ที่รอคำตอบ

2. **กราฟ/ภาพ** — โปรแกรมที่ render กราฟ (matplotlib ฯลฯ) บน server ต้องตั้ง backend แบบ headless (`Agg`) แล้วส่งภาพ/ไฟล์กลับ ไม่ใช่เปิดหน้าต่าง

3. **ความลับข้อมูล (สำคัญ)** — ไฟล์ .sav/.xlsx เป็นข้อมูลวิจัย
   - เก็บ S3 แบบ private + presigned URL หมดอายุสั้น
   - ลบไฟล์ทันทีหลังดาวน์โหลด/หมดเวลา (เช่น 1–24 ชม.)
   - แยก path ตาม user id
   - เลือก region ใกล้ (สิงคโปร์/ไทย) + เข้ารหัส at-rest

4. **การตั้งค่าที่เคยเป็นไฟล์ Excel** — บางโปรแกรม (เช่น Move Sheet) save/load settings เป็นไฟล์
   → บนเว็บเปลี่ยนเป็นบันทึก preset ใน DB ต่อ user หรือ export/import JSON

5. **เริ่มเรียบง่ายก่อน** — เฟสแรกใช้ FastAPI `BackgroundTasks` + เก็บ job ใน memory/Redis ธรรมดา
   ค่อยอัปเป็น Celery ตอนงานหนักจริง อย่า over-engineer ตั้งแต่วันแรก

6. **DB เริ่มจำเป็นเมื่อไหร่** — ถ้าต้องเก็บประวัติงาน/preset/ผู้ใช้ ค่อยเพิ่ม PostgreSQL
   ช่วงนำร่องยังไม่ต้องมีก็ได้

---

## 9. Definition of Done (ต่อ 1 โปรแกรม)
- [ ] Logic อยู่ใน `services/` เป็น function บริสุทธิ์ ไม่มีโค้ด GUI
- [ ] มี router + schema + หน้า React ครบ
- [ ] อัปโหลดไฟล์จริง → ได้ผลลัพธ์ดาวน์โหลดได้
- [ ] มี test เทียบ output ใหม่ == output จากตัวเดิม (ผ่าน)
- [ ] จัดการ error/ไฟล์ผิดรูปแบบได้ (ไม่ค้าง ไม่ 500 เปล่าๆ)
- [ ] ลบไฟล์ชั่วคราวหลังเสร็จ
- [ ] ติ๊ก ✅ ในตาราง §7

---

_อัปเดตล่าสุด: 2026-06-07 — แก้ตารางสถานะระหว่างพอร์ตแต่ละตัว_
