# แผน Refactor: THE PY-LIB (ระบบห้องสมุดอัจฉริยะ)

> เป้าหมาย: เปลี่ยนจาก `main.py` ไฟล์เดียว 2,526 บรรทัด ให้เป็นโครงสร้างแบบแยกชั้น (layered), แก้ช่องโหว่ความปลอดภัยจริง, และทำให้ codebase ดูแลง่ายขึ้นโดยไม่พังของเดิม

---

## 0. สถานะปัจจุบัน (จากการอ่านโค้ดจริงใน repo)

- ทุกอย่างอยู่ใน `main.py` คลาสเดียว `LibraryApp` — GUI (CustomTkinter), DB (sqlite3 ตรงๆ), QR generation/scan, PDF export (reportlab), Excel import (pandas) ปนกันหมด
- **รหัสผ่าน admin เก็บเป็น plaintext** ในตาราง `admin_users` และ login เทียบ string ตรงๆ — ทั้งที่ README อ้างว่าใช้ bcrypt/AES-256/PBKDF2/HMAC ซึ่ง**ไม่ตรงกับโค้ดจริง**
- ข้อมูล QR เข้ารหัสด้วย `base64` เฉยๆ (ไม่ใช่ encryption จริง) — ใครสแกนแล้ว decode ก็เห็นชื่อ-ชั้น-เลขที่ตรงๆ
- DB connection (`self.conn`/`self.cursor`) ตัวเดียวถูกใช้ทั้ง main thread และ QR-scan thread พร้อมกัน → เสี่ยง race condition
- ทุกหน้าจอ rebuild ด้วย `winfo_children()` destroy แล้ว pack ใหม่ทั้งหมด — ใช้งานได้แต่ไม่ scale และไม่มีทางเทสต์ UI logic แยกจาก DB logic ได้เลย

---

## 1. โครงสร้างไฟล์เป้าหมาย

```
py-lib/
├── main.py                     # entry point เท่านั้น (สร้าง App แล้ว run)
├── app.py                      # LibraryApp shell + navigation controller
├── config/
│   ├── settings.py              # โหลด/เซฟ settings.json (path, theme, font)
│   └── constants.py             # ค่าคงที่ เช่น due-date default, ตาราง DB
├── core/
│   ├── security.py              # bcrypt hash/verify, AES-256-CBC encrypt/decrypt, key derivation (PBKDF2)
│   └── logger.py                # logging config (แทน print())
├── db/
│   ├── connection.py             # เปิด sqlite connection แบบ per-thread (check_same_thread=False + lock หรือ thread-local)
│   ├── schema.py                  # CREATE TABLE ทั้งหมด (migration-friendly)
│   └── repositories/
│       ├── member_repo.py         # CRUD สมาชิก
│       ├── book_repo.py           # CRUD หนังสือ
│       ├── borrow_repo.py         # ยืม-คืน
│       └── access_repo.py         # access log
├── services/
│   ├── qr_service.py              # gen QR (เข้ารหัสจริงด้วย AES แทน base64), decode
│   ├── pdf_service.py             # สร้างบัตรสมาชิก/รายงาน PDF (reportlab)
│   ├── excel_service.py           # import/export หนังสือผ่าน pandas
│   └── camera_service.py          # เปิดกล้อง + สแกน QR (แยก thread logic ออกจาก UI)
├── ui/
│   ├── views/
│   │   ├── login_view.py
│   │   ├── dashboard_view.py
│   │   ├── member_view.py
│   │   ├── book_view.py
│   │   ├── borrow_view.py
│   │   ├── return_view.py
│   │   ├── history_view.py
│   │   ├── access_view.py
│   │   └── settings_view.py
│   └── widgets/
│       ├── date_picker.py         # ของเดิมมีอยู่แล้วแค่ย้ายมา
│       └── confirm_dialog.py
├── assets/                        # เหมือนเดิม
├── database/
│   └── library.db
├── tests/
│   ├── test_security.py
│   ├── test_member_repo.py
│   └── test_book_repo.py
├── requirements.txt
└── .env.example                   # เก็บ AES key/secret แยกจากโค้ด (ห้าม commit .env จริง)
```

---

## 2. Phase 1 — ความปลอดภัย (ทำก่อนสุด, กระทบข้อมูลผู้ใช้จริง)

1. **Hash รหัสผ่าน admin ด้วย bcrypt**
   - เขียน migration script: อ่าน `admin_users` เดิม → hash รหัสผ่านที่มีอยู่ → เขียนกลับ
   - แก้ `login()` ให้ใช้ `bcrypt.checkpw()` แทนเทียบ string ตรงๆ
   - บังคับให้ admin คนแรกต้องตั้งรหัสใหม่ตอน migrate (อย่าปล่อย `admin123` เป็นรหัสจริงต่อ)
2. **เข้ารหัสข้อมูล QR ด้วย AES-256-CBC จริง** (ไม่ใช่ base64)
   - ใช้ `cryptography` หรือ `pycryptodome` (มีอยู่ใน requirements.txt แล้วแต่ยังไม่ได้ใช้จริงกับ QR)
   - key derive ด้วย PBKDF2-HMAC-SHA256 จาก secret ใน `.env` (ห้าม hardcode key ในซอร์ส)
   - เพิ่ม HMAC ตรวจสอบความถูกต้องของข้อมูลก่อน decode ตอนสแกน
3. ย้าย secret/config (AES key, default admin password) ออกจากโค้ดไปที่ `.env` + `.env.example` เป็นตัวอย่าง
4. เพิ่ม rate-limit/lockout ตอน login ผิดเกิน N ครั้ง (README อ้างว่ามีแล้วแต่ในโค้ดที่เห็นยังไม่มี)

---

## 3. Phase 2 — แยกชั้น Data (DB → Repository pattern)

- ย้าย `CREATE TABLE` ทั้งหมดไป `db/schema.py`, เรียกตอน init ครั้งเดียว
- สร้าง repository class ต่อ entity (`MemberRepository`, `BookRepository`, ...) ที่รับ connection แล้วมีเมธอด `add()`, `get_all()`, `delete()`, `find_by_id()` ฯลฯ — UI จะไม่ยิง SQL ตรงๆ อีกต่อไป
- แก้ปัญหา thread-safety: ใช้ `sqlite3.connect(..., check_same_thread=False)` ร่วมกับ `threading.Lock()` รอบทุก query หรือเปิด connection ใหม่ต่อ thread (แนะนำอย่างหลังถ้าจะสแกน QR บ่อย)

---

## 4. Phase 3 — แยกชั้น Service (business logic)

- `qr_service.py`: รวม logic generate/encrypt/decrypt QR ที่กระจายอยู่หลายจุดใน main.py (add_member, view_member_qr, process_qr_data ฯลฯ) ให้เหลือจุดเดียว
- `pdf_service.py`: ย้าย `generate_member_card_pdf` และฟังก์ชันสร้างรายงานออกมา
- `camera_service.py`: ย้าย logic scan thread (`scan_qr`, queue, cv2 loop) ออกจาก UI class — UI แค่ subscribe ผลลัพธ์ผ่าน callback/queue

---

## 5. Phase 4 — แยกชั้น UI (Views)

- แต่ละหน้าจอ (login, dashboard, member, book, borrow, return, history, access, settings) เป็นคลาส View แยกไฟล์ รับ `parent` + `repositories`/`services` ที่ต้องใช้ผ่าน constructor (dependency injection แบบง่ายๆ)
- `app.py` ทำหน้าที่แค่ navigation: สร้าง view ใหม่/ทำลาย view เก่า แทนที่ logic กระจายอยู่ใน `LibraryApp` ทุกเมธอด
- เก็บพฤติกรรมเดิมไว้ (responsive font บน resize, quick-date buttons ฯลฯ) แค่ย้ายที่อยู่

---

## 6. Phase 5 — Logging, error handling, tests

- แทน `print()` ทั้งหมดด้วย `logging` module (log ไปไฟล์ + console, level แยก DEBUG/ERROR)
- เปลี่ยน bare `except Exception as e: self.show_error(...)` ให้ log stack trace ด้วย ไม่ใช่แค่โชว์ข้อความสั้นๆ ให้ user
- เขียน unit test สำหรับ `core/security.py` และ repository layer (ใช้ in-memory SQLite เหมือนที่เคยทำกับ DSNPRU_REG ได้เลย)

---

## 7. ลำดับการทำจริง (แนะนำ)

1. Phase 1 (ความปลอดภัย) — ทำก่อนเพราะกระทบข้อมูลจริงของนักเรียนที่ใช้อยู่
2. Phase 2 (repository) — เป็นฐานให้ phase อื่นแยกง่ายขึ้น
3. Phase 3 (services)
4. Phase 4 (views) — ใหญ่สุด ทำทีละหน้าจอ ทดสอบให้พฤติกรรมเดิมไม่เปลี่ยนก่อนไปหน้าถัดไป
5. Phase 5 (logging/tests) — ทำคู่ขนานได้ตลอด ไม่ต้องรอ phase อื่นเสร็จ

แต่ละ phase commit แยก เพื่อ rollback ได้ถ้าอะไรพัง

---

## 8. หมายเหตุสำหรับ coding agent ที่จะรับแผนนี้ไปทำต่อ

- ห้ามลบฟีเจอร์เดิม (import excel, generate card PDF, quick due-date buttons, resize font ฯลฯ) — refactor คือย้ายที่ ไม่ใช่ตัดทิ้ง
- ต้อง migrate ข้อมูลเดิมใน `db/library.db` ได้ ไม่ใช่สร้างฐานใหม่ทิ้งของเก่า
- ทดสอบทีละ phase ก่อนไป phase ถัดไป
