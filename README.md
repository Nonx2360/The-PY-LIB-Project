# THE PY-LIB — ระบบห้องสมุดอัจฉริยะ

**THE PY-LIB** คือโปรแกรมจัดการห้องสมุดที่พัฒนาด้วย Python (CustomTkinter + Flask + SQLite) เน้นความปลอดภัย ความสะดวก และการใช้งานจริงในโรงเรียนหรือสถาบันการศึกษา รองรับทั้งแอปเดสท็อปและเว็บไซต์

---

## ฟีเจอร์เด่น

### แอปเดสท็อป
- GUI แบบโมดูลาร์ (แยกชั้น config / core / db / services / ui)
- สแกน QR Code ด้วยกล้อง Webcam เพื่อยืม-คืนและบันทึกเข้าออก
- ยืม-คืนแบบ atomic transaction — ป้องกันการยืมซ้ำ/ข้อมูลค้างจากความผิดพลาด
- แสดงสถานะคืนช้า พร้อมสีเตือน (เขียว / เหลือง / แดง)
- หยุดให้บริการ (retire) / คืนสถานะหนังสือได้ โดยประวัติการยืมยังถูกเก็บไว้
- ส่งออกบัตรสมาชิก PDF, ประวัติยืม-คืน, ประวัติเข้าออก พร้อมฟอนต์ไทย (Sarabun)
- นำเข้า/ส่งออกหนังสือจากไฟล์ Excel (pandas + openpyxl) — นำเข้ารอบละพันเล่มขึ้นไปได้ (จัดการใน transaction เดียว)
- UI ทันสมัย, รองรับธีมสว่าง/มืด, Sidebar พร้อม Grouped Navigation
- Toast notifications, Confirm dialogs, Shared scan window
- Built-in PDF viewer (PyMuPDF) สำหรับดูบัตรสมาชิก
- Backups (SQLite backup + checksum) ใช้ก่อนการเปลี่ยนแปลงสำคัญ

### เว็บไซต์ (Flask Web UI)
- เปิดใช้งานจาก sidebar เมนู "Web UI" ในแอปเดสท็อป หรือรันแยกต่างหาก
- ฟีเจอร์เทียบเท่าเดสท็อป: สมาชิก, หนังสือ, ยืม-คืน, ประวัติ, ตั้งค่า
- บันทึกเข้า-ออกผ่านเว็บ (สลับเข้า/ออกอัตโนมัติ)
- สแกนบัตรสมาชิกจากกล้องเว็บเบราว์เซอร์ (jsQR vendored) — ถอดรหัสเฉพาะฝั่งเซิร์ฟเวอร์
- หยุดให้บริการ/กลับมา และลบถาวรแบบ 2 ขั้นตอน (retire → purge) พร้อมเก็บอย่างปลอดภัย
- นำเข้าหนังสือจากไฟล์ Excel ออนเว็บ (สูงสุด 5,000 แถว/ไฟล์) พร้อมดาวน์โหลดแม่แบบ — รหัสซ้ำ/แถวไม่สมบูรณ์ถูกข้ามและรายงานผล
- ตารางหนังสือ/สมาชิก/ประวัติแบ่งหน้า server-side (50 แถว/หน้า) — แคตตาล็อก 1,000+ เล่มเปิดเร็วตลอด
- SVG icons (Lucide), School logo, Search/filter ทุกหน้า
- Responsive design, รองรับ Light/Dark mode
- Frontend stack vendored ในเครื่อง (ทำงานได้แม้ไม่มีอินเทอร์เน็ต) — ดู `web/static/vendor/README.md`:
  - **SweetAlert2** — กล่องยืนยัน/แจ้งเตือนธีมเดียวกับระบบ, flash messages เป็น toast
  - **HTMX** — ค้นหา/กรองหนังสือ (รหัส/ชื่อ/สถานะ), สมาชิก, ประวัติ แบบ server-side โดยไม่โหลดหน้าใหม่
  - **Alpine.js** — ล็อกสถานะหน้ายืม (ค้างส่ง/ครบจำนวน) ประเมินทันทีในเบราว์เซอร์
  - **Chart.js** — แดชบอร์ด: การยืม 14 วันล่าสุด, เข้าออก 7 วันล่าสุด, สัดส่วนสถานะหนังสือ (`/api/analytics`)
  - **Flatpickr (th)** — ปฏิทินเลือกกำหนดคืน พร้อมปุ่มด่วน 7/14/21/30 วัน
  - ลำดับโหลด: sweetalert2 → htmx → ui.js → alpine (สำคัญ อย่าสลับ)

### ความปลอดภัย
- ข้อมูล QR เข้ารหัสแบบ AEAD **AES-256-GCM (`v2:`)** และยังถอดรหัสการ์ดเก่าแบบ AES-256-CBC อย่างเข้มงวด (strict legacy fallback)
- PBKDF2 key derivation (100k iterations) สำหรับ QR encryption key
- รหัสผ่าน admin แฮชด้วย bcrypt — ไม่มีรหัสผ่านดีฟอลต์ และตรวจแบบ timing-equalized
- จำกัดการล็อกอินล้มเหลว (5 ครั้ง / 15 นาที ต่อผู้ใช้) ทั้งเว็บและเดสท็อป
- CSRF token ทุกฟอร์ม POST และ session cookie (HttpOnly/SameSite=Lax)
- Parameterized queries ป้องกัน SQL Injection ทุกที่
- Thread-safe database access (connection-per-query + Lock) และ FK enforcement เปิดทุก connection
- Secret migration ด้วย `PRAGMA user_version` ตาม schema version
- Secrets เก็บใน `.env` (gitignored) — สร้างแบบสุ่มให้อัตโนมัติครั้งแรก

---

## การติดตั้งและใช้งาน

### ติดตั้ง Python

ต้องการ Python >= 3.10

### สร้าง Virtual Environment และติดตั้ง Dependencies

```bash
python -m venv venv
venv\Scripts\activate        # Windows
pip install -r requirements.txt
```

รองรับการใช้ `req.bat` สำหรับติดตั้ง dependencies อัตโนมัติ

### รันแอปเดสท็อป

```bash
python main.py              # โหมดปกติ
python main.py --debug      # โหมด debug (log ลง debug.log)
```

ค่าเริ่มต้น: ไม่มีรหัสผ่านดีฟอลต์อีกต่อไป — การบูตครั้งแรกจะ**สร้างรหัสผ่านแบบสุ่ม** แสดงบนหน้าจอคอนโซลเพียงครั้งเดียว และบันทึกไว้ที่ `db/INITIAL_ADMIN_PASSWORD.txt` (ไม่ถูกติดตามใน Git) ถ้าตรวจพบรหัสผ่านดีฟอลต์เดิม (`admin123`) ในฐานข้อมูลเดิม ระบบจะเปลี่ยนให้เป็นรหัสใหม่อัตโนมัติทุกครั้งที่บูต

> การจำกัดการล็อกอิน (login throttling): ล็อกอินผิดเกิน 5 ครั้งใน 15 นาที (ต่อชื่อผู้ใช้) จะต้องรอครบกำหนด ทั้งหน้าเว็บและเดสท็อป

### รันเว็บไซต์

เปิดแอปเดสท็อป → เมนู "เพิ่มเติม" → "Web UI" → กด "เริ่ม Web UI"

หรือรันแยกต่างหาก (ผูกกับเฉพาะเครื่องนี้ 127.0.0.1):

```bash
venv\Scripts\python -c "from web.app import create_app; create_app().run(port=5000)"
```

เปิดเบราว์เซอร์ไปที่ `http://127.0.0.1:5000`

> ความปลอดภัย: เว็บยืนยันตัวตนด้วย session cookie (HttpOnly/SameSite=Lax), CSRF token
> ในทุกฟอร์ม POST, จำกัดการล็อกอินล้มเหลว (5 ครั้ง/15 นาที) และไม่บันทึกข้อมูลส่วนบุคคล
> ลง log เว็บถูกผูกกับ `127.0.0.1` เป็นค่าเริ่มต้น — ถ้าต้องการเปิดบน LAN ดูขั้นตอน HTTPS
> ที่ `docs/SECURITY.md` และเปลี่ยนค่า `webui_host` ในหน้าตั้งค่า

> คีย์ QR: ไฟล์ `.env` เก็บ `SECRET_PASSPHRASE` / `SECRET_SALT` / `FLASK_SECRET_KEY`
> (gitignored; สร้างให้อัตโนมัติครั้งแรก ค้างเก็บให้ปลอดภัย) รูปแบบข้อมูล QR ใหม่ใช้
> AES-256-GCM (`v2:`) การหมุนเวียนคีย์/ย้ายเครื่อง/กู้คืน — อ่าน `docs/SECURITY.md`

> สแกนบัตรสมาชิกบนเว็บ: กล้องเว็บเบราว์เซอร์ต้องต้องเปิดใน HTTPS (ยกเว้น localhost)
> ดูรายละเอียดการเปิด HTTPS ที่ `docs/SECURITY.md` § deployment

---

## โครงสร้างโปรเจกต์

```
THE-PY-LIB/
├── main.py                      # Entry point (--debug flag)
├── app.py                       # LibraryApp shell + navigation controller
├── legacy_main.py               # Original monolith (preserved for rollback)
├── config/
│   ├── constants.py             # DB_PATH, font paths, book status constants
│   ├── settings.py              # User settings (config/settings.json)
│   ├── env_store.py             # .env secrets: auto-generate, guard placeholders
│   └── settings.json            # User settings (auto-created)
├── core/
│   ├── security.py              # bcrypt hash/verify, AES-256-GCM `v2:` + legacy AES-CBC (PBKDF2)
│   ├── logger.py                # logging config (debug.log support)
│   └── throttle.py              # Sliding-window login lockout (desktop + web)
├── db/                          # บางไฟล์ถูกสร้างตอน runtime: library.db, INITIAL_ADMIN_PASSWORD.txt (gitignored)
│   ├── schema.py                # CREATE TABLE + user_version migrations + admin bootstrap
│   └── repositories/
│       ├── base_repo.py         # Thread-safe SQLite (per-query connection + Lock)
│       ├── member_repo.py       # CRUD สมาชิก (+ badge uniqueness check)
│       ├── book_repo.py         # CRUD หนังสือ (uses book status constants)
│       ├── borrow_repo.py       # try_borrow / try_return — atomic invariants
│       ├── access_repo.py       # Access log
│       └── admin_repo.py        # admin_users + timing-equalized verify_credentials
├── services/
│   ├── qr_service.py            # สร้าง/ถอดรหัส QR (AES)
│   ├── pdf_service.py           # บัตรสมาชิก PDF + รายงาน
│   ├── camera_service.py        # กล้อง + ถอดรหัส QR
│   ├── excel_service.py         # Import/Export Excel
│   ├── backup_service.py        # SQLite backups + checksum verification
│   └── __init__.py
├── ui/
│   ├── theme.py                 # Design tokens, font registration, widget helpers
│   ├── views/
│   │   ├── login_view.py        # Split-panel login + throttling
│   │   ├── dashboard_view.py    # Stat cards + action tiles
│   │   ├── member_view.py       # Two-column: form + searchable list
│   │   ├── book_view.py         # Two-column: form + searchable list, retire/restore
│   │   ├── borrow_view.py       # 3-step flow (scan → book → confirm)
│   │   ├── return_view.py       # Scan + borrowed list + return
│   │   ├── history_view.py      # Table search + PDF export
│   │   ├── access_view.py       # Scanner + history (AccessScannerView/AccessHistoryView)
│   │   ├── settings_view.py     # Theme, loan settings, school info, DB backup
│   │   └── webui_view.py        # Web UI launcher control panel
│   └── widgets/
│       ├── confirm_dialog.py    # Modal confirm dialog
│       ├── toast.py             # Floating toast notifications
│       ├── scan_window.py       # Shared camera QR scan window
│       └── pdf_viewer.py        # Built-in PDF viewer (PyMuPDF)
├── web/
│   ├── app.py                   # Flask web application (routes, guards, pagination, /api/analytics, Excel import)
│   ├── csrf.py                  # Per-session CSRF tokens (form field หรือ X-CSRF-Token header)
│   ├── static/
│   │   ├── style.css            # Web UI design system
│   │   ├── vendor/              # Vendored JS libs, pinned versions (see README.md inside)
│   │   │   ├── alpine.min.js    # Alpine.js 3.17.5
│   │   │   ├── htmx.min.js      # HTMX 2.0.11
│   │   │   ├── sweetalert2.min.js/.css  # SweetAlert2 11.26.25
│   │   │   ├── chart.umd.min.js # Chart.js 4.6.1
│   │   │   └── flatpickr.min.*
│   │   └── js/
│   │       ├── jsQR.js          # Vendored jsQR (in-browser QR decode)
│   │       ├── scan.js          # Camera capture + /scan/resolve calls
│   │       └── ui.js            # PyLib helpers: dialogs, toasts, data-confirm, flatpickr init
│   └── templates/
│       ├── icons.svg            # Lucide SVG icon sprite
│       ├── base.html            # Layout + sidebar + vendor stack wiring
│       ├── login.html           # Login page
│       ├── index.html           # Dashboard + Chart.js widgets
│       ├── members.html         # Member management
│       ├── books.html           # Book management (retire/restore/purge)
│       ├── borrow.html          # Borrow (Alpine lock state + Flatpickr due date)
│       ├── return_book.html     # Return with overdue badges
│       ├── history.html         # Borrow history
│       ├── access.html          # Access recording + history
│       ├── settings.html        # Settings + about
│       └── partials/            # HTMX table partials (_books/_members/_history)
├── tools/
│   ├── rotate_qr_key.py         # Key rotation: backup → decrypt → re-encrypt
│   └── regenerate_member_cards.py
├── tests/
│   ├── _harness.py              # Temp-DB/env isolation for tests
│   ├── test_security.py         # Password hashing + QR crypto
│   ├── test_qr_security.py      # QR payload formats & rejection paths
│   ├── test_migrations.py       # Schema v1→v2 + book-code dedupe
│   ├── test_member_repo.py      # Member CRUD
│   ├── test_book_repo.py        # Book CRUD
│   ├── test_borrow_atomic.py    # Atomic borrow/return invariants
│   ├── test_admin_bootstrap.py  # No-default-password bootstrap
│   ├── test_backup_restore.py   # Backup + verify
│   ├── test_web_security.py     # CSRF/throttle/session
│   ├── test_web_scan.py         # Badge scanner endpoint
│   ├── test_web_delete_guards.py# History-bearing delete refusals
│   ├── test_web_book_retire.py  # Retire/restore/purge flow
│   ├── test_web_bulk.py         # 1000-row Excel import + list pagination
│   ├── test_web_partials.py     # HTMX partials + header CSRF + analytics
│   └── test_web_pages_smoke.py  # Every authenticated page renders 200
├── docs/
│   ├── SECURITY.md              # Security runbook (key rotation, HTTPS/LAN)
│   └── DESIGN.md                # Visual design system notes
├── assets/
│   ├── fonts/                   # Sarabun-Regular.ttf, Sarabun-Bold.ttf
│   ├── logos/                   # school_logo.png
│   ├── qrcodes/                 # Generated QR images
│   └── cards/                   # Generated member card PDFs
├── DB/                          # (runtime, gitignored) INITIAL_ADMIN_PASSWORD.txt lives here when created
├── .env.example                 # SECRET_PASSPHRASE, SECRET_SALT, FLASK_SECRET_KEY template
├── .env                         # Secrets (gitignored, auto-generated)
├── requirements.txt
└── README.md
```

---

## การตั้งค่า

เปิดแอป → เมนู "ตั้งค่า" หรือเว็บ → เมนู "ตั้งค่า":

| ตัวเลือก | คำอธิบาย | ค่าเริ่มต้น |
|---|---|---|
| ธีม | System / Light / Dark | System |
| ระยะเวลาการยืม | จำนวนวันเริ่มต้นเมื่อยืมหนังสือ | 7 วัน |
| จำนวนหนังสือสูงสุดต่อคน | จำกัดจำนวนหนังสือที่ยืมได้พร้อมกัน | 3 เล่ม |
| อายุบัตรสมาชิก | จำนวนวันก่อนบัตรหมดอายุ | 365 วัน |
| ชื่อโรงเรียน | แสดงในบัตรสมาชิกและหน้าเว็บ | DSNPRU |
| โลโก้โรงเรียน | ไฟล์ PNG สำหรับบัตรสมาชิกและ sidebar | — |
| Web UI host | 127.0.0.1 (loopback) หรือ 0.0.0.0 (LAN + HTTPS) | 127.0.0.1 |

---

## ระบบความปลอดภัย

| คุณสมบัติ | วิธีใช้ |
|---|---|
| รหัสผ่าน admin | `bcrypt` hash/verify; ไม่มีดีฟอลต์ — บูตแรกสร้างรหัสสุ่มและปริ้นครั้งเดียว; 'admin123' เก่าถูกปลดอัตโนมัติทุกบูต |
| เข้ารหัส QR | `AES-256-GCM` (`v2:` authenticated envelope) ใหม่; legacy `AES-256-CBC` ถอดรหัสแบบ strict จนกว่าจะย้ายด้วย `tools/rotate_qr_key.py` |
| Key derivation | PBKDF2-HMAC-SHA256 (100k iters) จาก secrets ใน `.env` ที่ถูก generate อัตโนมัติและกัน placeholder |
| Login throttling | เว็บและเดสท็อป: 5 ครั้งล้มเหลวต่อ 15 นาที → ล็อกชั่วคราว (เว็บต่อ IP+username) |
| Web session | HttpOnly/SameSite=Lax cookie, CSRF token ทุก POST (constant-time compare), login timing-equalized กัน enumeration |
| SQL Injection | Parameterized queries ทุกที่, status เปรียบเทียบผ่านพารามิเตอร์ |
| Data integrity | FK enforcement, UNIQUE index บน books.code, borrow_log CHECK (returned IN (0,1)), borrow/return เป็น single IMMEDIATE transaction |
| Thread safety | `BaseRepository` ใช้ connection-per-query + `threading.Lock`, `transact()` สำหรับ multi-statement |
| Schema migrations | `PRAGMA user_version` (v1 legacy → v2 constraints → v3 bulk-import indexes) รันอัตโนมัติและ idempotent ทุกบูต |
| Secrets | `.env` (gitignored), ป้องกันค่า placeholder, การเปลี่ยนมี runbook ใน `docs/SECURITY.md` |

---

## เครื่องมือเสริม

```bash
# หมุนเวียนคีย์เข้ารหัส QR (สำคัญ: ก่อนเปิดแอปครั้งถัดไป เมื่อ .env เปลี่ยน)
venv\Scripts\python tools\rotate_qr_key.py [--no-cards]

# สร้างบัตรสมาชิก PDF ซ้ำจากฐานข้อมูลปัจจุบัน
venv\Scripts\python tools\regenerate_member_cards.py
```

---

## รัน Tests

```bash
# Windows
venv\Scripts\python -m unittest discover -s tests

# Linux/Mac
venv/bin/python -m unittest discover -s tests
```

ปัจจุบันมี 111 test cases (15 test files) ครอบคลุม security, repositories, migrations, atomic borrow/return, backups, bulk imports, HTMX partials/analytics และ guards ของเว็บ

---

## ฐานข้อมูล (SQLite3)

| Table | คำอธิบาย |
|---|---|
| `admin_users` | ผู้ดูแลระบบ (bcrypt hashed password, ไม่มีดีฟอลต์) |
| `members` | ข้อมูลสมาชิกห้องสมุด (ชื่อ, ชั้น, เลขที่, QR, วันหมดอายุ) — badge key (ชั้น, เลขที่) ต้องไม่ซ้ำกันในระบบใหม่ |
| `books` | หนังสือในระบบ (รหัส [UNIQUE], ชื่อ, สถานะ: `ว่าง` / `ยืมแล้ว` / `หยุดให้บริการ`) |
| `borrow_log` | บันทึกการยืม-คืน (returned CHECK 0/1, FK ไป members/books) |
| `access_log` | บันทึกการเข้า-ออกห้องสมุด |

---

## ไลบรารีที่ใช้

| แพ็กเกจ | วัตถุประสงค์ |
|---|---|
| `customtkinter` | GUI framework (Modern Tkinter) |
| `flask` | Web UI framework |
| `opencv-python` | กล้อง Webcam + QR scan |
| `pyzbar` | ถอดรหัส QR Code (เดสท็อป) |
| `Pillow` | จัดการรูปภาพ |
| `cryptography` | AES-256-GCM / AES-256-CBC encryption, PBKDF2 |
| `bcrypt` | Password hashing |
| `reportlab` | สร้าง PDF (บัตรสมาชิก, รายงาน) |
| `PyMuPDF` (`pymupdf`) | PDF viewer ในแอป |
| `numpy` | Image processing (opencv) |
| `qrcode` | สร้าง QR Code |
| `pandas` | Excel import/export |
| `openpyxl` | Excel engine ของ pandas (.xlsx) |
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
