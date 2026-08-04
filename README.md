# THE PY-LIB — ระบบห้องสมุดอัจฉริยะ

**THE PY-LIB** คือโปรแกรมจัดการห้องสมุดที่พัฒนาด้วย Python (CustomTkinter + SQLite) เน้นความปลอดภัย ความสะดวก และการใช้งานจริงในโรงเรียนหรือสถาบันการศึกษา

---

## ฟีเจอร์เด่น

- จัดการข้อมูลสมาชิกและหนังสือผ่าน GUI แบบโมดูลาร์ (แยกชั้น UI / Service / Repository)
- สแกน QR Code ด้วยกล้อง Webcam เพื่อยืม-คืนและบันทึกเข้าออก
- ข้อมูล QR เข้ารหัสด้วย AES-256-CBC, รหัสผ่าน admin แฮชด้วย bcrypt
- แสดงสถานะคืนช้า พร้อมสีเตือน (เขียว / เหลือง / แดง)
- ส่งออกบัตรสมาชิก PDF, ประวัติยืม-คืน, ประวัติเข้าออก พร้อมฟอนต์ไทย (Sarabun)
- นำเข้าหนังสือจาก Excel (pandas)
- ระบบล็อกอินผู้ดูแลระบบ
- บันทึกประวัติการยืม/คืน และการเข้าใช้งานห้องสมุด
- UI ทันสมัย, รองรับธีมสว่าง/มืด, Sidebar พร้อม Grouped Navigation

---

## การติดตั้งและใช้งาน

### ติดตั้ง Python

ต้องการ Python >= 3.8

### สร้าง Virtual Environment และติดตั้ง Dependencies

```bash
python -m venv venv
venv\Scripts\activate        # Windows
pip install -r requirements.txt
```

### รันโปรแกรม

```bash
python main.py
```

ค่าเริ่มต้น: ชื่อผู้ใช้ `admin` / รหัสผ่าน `admin123` (จะถูกแฮชด้วย bcrypt อัตโนมัติครั้งแรก)

---

## โครงสร้างโปรเจกต์

```
THE-PY-LIB/
├── main.py                      # Entry point
├── app.py                       # LibraryApp shell + navigation controller
├── config/
│   ├── constants.py             # DB_PATH, paths, colors
│   └── __init__.py
├── core/
│   ├── security.py              # bcrypt hash/verify, AES-256-CBC encrypt/decrypt (PBKDF2)
│   ├── logger.py                # logging config
│   └── __init__.py
├── db/
│   ├── schema.py                # CREATE TABLE + admin seeding/migration
│   ├── repositories/
│   │   ├── base_repo.py         # Thread-safe SQLite (per-query connection + Lock)
│   │   ├── member_repo.py       # CRUD สมาชิก
│   │   ├── book_repo.py         # CRUD หนังสือ
│   │   ├── borrow_repo.py       # ยืม-คืน
│   │   └── access_repo.py       # Access log
│   └── __init__.py
├── services/
│   ├── qr_service.py            # สร้าง/ถอดรหัส QR (AES-256-CBC, fallback base64)
│   ├── pdf_service.py           # บัตรสมาชิก PDF + รายงาน
│   ├── camera_service.py        # กล้อง + ถอดรหัส QR
│   ├── excel_service.py         # Import/Export Excel
│   └── __init__.py
├── ui/
│   ├── theme.py                 # Design tokens, font registration, widget helpers
│   ├── views/
│   │   ├── login_view.py        # Split-panel login (brand + form)
│   │   ├── dashboard_view.py    # Stat cards + action tiles
│   │   ├── member_view.py       # Two-column: form + searchable list
│   │   ├── book_view.py         # Two-column: form + searchable list
│   │   ├── borrow_view.py       # 3-step flow (scan → book → confirm)
│   │   ├── return_view.py       # Scan + borrowed list + return
│   │   ├── history_view.py      # Table search + PDF export
│   │   ├── access_view.py       # Scanner + history
│   │   └── settings_view.py     # Theme switcher + about
│   └── widgets/
│       ├── confirm_dialog.py    # Modal confirm dialog
│       ├── toast.py             # Floating toast notifications
│       └── scan_window.py       # Shared camera QR scan window
├── assets/
│   ├── fonts/                   # Sarabun-Regular.ttf, Sarabun-Bold.ttf
│   ├── logos/                   # school_logo.png
│   ├── qrcodes/                 # Generated QR images
│   └── cards/                   # Generated member card PDFs
├── db/
│   └── library.db               # SQLite database
├── tests/
│   ├── test_security.py         # Password hashing + QR encryption
│   ├── test_member_repo.py      # Member CRUD
│   └── test_book_repo.py        # Book CRUD
├── legacy_main.py               # Original monolith (preserved for rollback)
├── .env.example                 # SECRET_PASSPHRASE, SECRET_SALT template
├── .env                         # Secrets (gitignored)
├── .gitignore
├── requirements.txt
└── agents-progress.md           # Refactoring progress tracker
```

---

## ระบบความปลอดภัย

| คุณสมบัติ | วิธีใช้ |
|---|---|
| รหัสผ่าน admin | `bcrypt` hash/verify, auto-migrated from plain-text on boot |
| เข้ารหัส QR | `AES-256-CBC` + PBKDF2 key derivation (ไม่ใช่ base64 อีกต่อไป) |
| Legacy QR | Fallback base64 decoding สำหรับ QR ที่สร้างก่อน refactor |
| SQL Injection | Parameterized queries ทุกที่ |
| Thread safety | `BaseRepository` ใช้ connection-per-query + `threading.Lock` |
| Secrets | `.env` file (gitignored) สำหรับ SECRET_PASSPHRASE / SECRET_SALT |

---

## รัน Tests

```bash
# Windows
venv\Scripts\python -m unittest discover -s tests

# Linux/Mac
venv/bin/python -m unittest discover -s tests
```

---

## ฐานข้อมูล (SQLite3)

| Table | คำอธิบาย |
|---|---|
| `admin_users` | ผู้ดูแลระบบ (bcrypt hashed password) |
| `members` | ข้อมูลสมาชิกห้องสมุด (ชื่อ, ชั้น, เลขที่, QR, วันหมดอายุ) |
| `books` | หนังสือในระบบ (รหัส, ชื่อ, สถานะ) |
| `borrow_log` | บันทึกการยืม-คืน (สมาชิก, หนังสือ, วันยืม, กำหนดคืน, สถานะคืน) |
| `access_log` | บันทึกการเข้า-ออกห้องสมุด |

---

## ไลบรารีที่ใช้

| แพ็กเกจ | วัตถุประสงค์ |
|---|---|
| `customtkinter` | GUI framework (Modern Tkinter) |
| `opencv-python` | กล้อง Webcam + QR scan |
| `pyzbar` | ถอดรหัส QR Code |
| `Pillow` | จัดการรูปภาพ |
| `cryptography` | AES-256-CBC encryption |
| `bcrypt` | Password hashing |
| `reportlab` | สร้าง PDF |
| `numpy` | Image processing (opencv) |
| `qrcode` | สร้าง QR Code |
| `pandas` | Excel import/export |
| `python-dotenv` | โหลด .env |
| `packaging` | CustomTkinter dependency |

---

## ผู้พัฒนา

- เด็กหญิง ขวัญชนก อุ่นศิริ เลขที่ 4 (UX/UI)
- เด็กชาย ณัฐชนน รอดน้อย เลขที่ 30 (Developer)
- **โรงเรียน:** DSNPRU

---

## ขอบคุณ

- ผู้ทดสอบ beta
- อาจารย์ที่ปรึกษา

---

**MIT License © 2025**
