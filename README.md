<div align="center">

# THE PY-LIB

**ระบบบริหารห้องสมุดโรงเรียนครบวงจร — แอปเดสก์ท็อป, เว็บสำหรับผู้ดูแล และหน้าสมาชิกส่วนบุคคล**
QR ประจำตัวเข้ารหัส AES-256-GCM (รวมตั๋วหมุนเวียนแบบมีอายุ) • ยืม-คืน atomic • 193 การทดสอบ + CI

[![CI](https://github.com/Nonx2360/The-PY-LIB-Project/actions/workflows/ci.yml/badge.svg)](https://github.com/Nonx2360/The-PY-LIB-Project/actions/workflows/ci.yml)
![Python](https://img.shields.io/badge/Python-3.10%2B-306998?logo=python&logoColor=FFD43B)
![Platform](https://img.shields.io/badge/Desktop-Windows%2010%2B-0078D6?logo=windows&logoColor=white)
![Tests](https://img.shields.io/badge/tests-193%20passing-2EA44F)
![License](https://img.shields.io/badge/license-MIT-8A6BE8)

</div>

---

THE PY-LIB คือระบบจัดการห้องสมุดที่ออกแบบมาเพื่อห้องสมุดโรงเรียนจริง: จัดการสมาชิก/หนังสือ/การยืม-คืนผ่านแอปเดสก์ท็อปและเว็บสำหรับผู้ดูแล พร้อม **หน้าสมาชิก** ที่นักเรียนเข้าดูสถานะการยืมและแสดง **QR หมุนเวียนที่มีอายุตามเวลา** แทนการพึ่งบัตรพลาสติกใบเดียวตลอดกาล ทุกเลเยอร์ — การเข้ารหัส, สิทธิ์เข้าถึง, การทรานแซกชัน — ถูกออกแบบร่วมกันและครอบด้วยชุดทดสอบความปลอดภัยที่รันต่อเนื่องใน CI

## สารบัญ

- [สถาปัตยกรรมและสิทธิ์การเข้าถึง](#สถาปัตยกรรมและสิทธิ์การเข้าถึง)
- [คุณสมบัติ](#คุณสมบัติ)
- [เริ่มใช้งานอย่างรวดเร็ว](#เริ่มใช้งานอย่างรวดเร็ว)
- [วงจรชีวิตบัญชีสมาชิก](#วงจรชีวิตบัญชีสมาชิก)
- [การตั้งค่า](#การตั้งค่า)
- [เส้นทาง HTTP](#เส้นทาง-http)
- [การปฏิบัติการและบำรุงรักษา](#การปฏิบัติการและบำรุงรักษา)
- [รูปแบบความปลอดภัย](#รูปแบบความปลอดภัย)
- [ฐานข้อมูล](#ฐานข้อมูล)
- [การทดสอบและ CI](#การทดสอบและ-ci)
- [โครงสร้างโปรเจกต์](#โครงสร้างโปรเจกต์)
- [แพ็กเกจที่ใช้](#แพ็กเกจที่ใช้)
- [แผนการพัฒนา](#แผนการพัฒนา)
- [การมีส่วนร่วม](#การมีส่วนร่วม)
- [ทีมและใบอนุญาต](#ทีมและใบอนุญาต)

---

## สถาปัตยกรรมและสิทธิ์การเข้าถึง

```
                         ┌─────────────────────────────────────┐
                         │  แอปเดสก์ท็อป (Tk)                   │
                         │  สมาชิก • หนังสือ • ยืม-คืน             │
                         │  เข้า-ออก • ประวัติ • ตั้งค่า             │
                         └──────────────┬──────────────────────┘
                                        │ ใช้ db/ และกลุ่ม service เดียวกัน
┌────────────────────────┐    ┌─────────▼──────────────────────┐    ┌────────────────────────┐
│  เว็บผู้ดูแล (Flask)       │    │  คอร์กลางของระบบ               │    │  หน้าสมาชิก (เว็บ)         │
│  ตาราง server-side     ├───►│  db/repositories (atomic)      │◄───│  ล็อกอิน PIN (bcrypt)    │
│  jsQR → เซิร์ฟเวอร์        │    │  services/qr + member_identity │    │  เห็นเฉพาะของตัวเอง       │
│  bulk + analytics      │    │  core/security (AES-GCM)       │    │  ตั๋วหมุนเวียน v3          │
└────────────────────────┘    │  config/settings + .env        │    └────────────────────────┘
                              └─────────┬──────────────────────┘
                                        │
                              ┌─────────▼─────────┐
                              │   SQLite (db/)    │
                              └───────────────────┘  user_version v1→v4 อัตโนมัติ
```

**เมทริกซ์สิทธิ์** — ทุกหน้าถูกกั้นชัดเจนว่าใครเห็นอะไร:

| ความสามารถ | ผู้ดูแล (เดสก์ท็อป) | ผู้ดูแล (เว็บ) | สมาชิก (หน้าเว็บ) |
|---|:---:|:---:|:---:|
| เพิ่ม/แก้/ลบ สมาชิก + หนังสือ, จัดการแบบกลุ่ม | ✓ | ✓ | ✗ |
| ยืม • คืน • บันทึกเข้า-ออกห้องสมุด | ✓ | ✓ | ✗ |
| ออกรหัสเปิดใช้งาน / ล้าง PIN ของสมาชิก | ✓ | ✓ | ✗ |
| ปรับตั้งค่าระบบ (รวมเปิด/ปิดหน้าสมาชิก) | ✓ | ✓ | ✗ |
| ดูข้อมูลของตนเอง + QR หมุนเวียนของตนเอง | — | — | ✓ |

Session ของผู้ดูแลและสมาชิกเป็น key คนละชุดในคุกกี้เดียวกัน — การล็อกอินฝั่งหนึ่งไม่ให้สิทธิ์อีกฝั่งหนึ่ง และถูกตรวจสอบด้วยชุดทดสอบเฉพาะ

## คุณสมบัติ

### แอปเดสก์ท็อป (สำหรับผู้ดูแล)

- **ตัวตนสมาชิกด้วย QR ที่เข้ารหัส** — สแกนเว็บแคมเพื่อยืม คืน และบันทึกเข้า-ออก; ถอดรหัสฝั่งเครื่องด้วยคีย์จาก `.env` เท่านั้น
- **ยืม-คืนแบบ atomic** — `BEGIN IMMEDIATE` หนึ่งทรานแซกชันต่อการยืม/คืน: หนังสือเปลี่ยนสถานะด้วย conditional UPDATE (หนึ่งเล่มมีผู้ชนะได้เพียงรายเดียวแม้สแกนพร้อมกันจากสองเครื่อง) และความผิดพลาดกลางทางไม่ทิ้งข้อมูลค้าง
- **จัดการจำนวนมากแบบกลุ่ม** — เลือกหนังสือ/สมาชิกหลายรายการแล้วหยุดให้บริการ / กลับมา / ลบพร้อมกัน รายการที่มีประวัติถูกข้ามอัตโนมัติ พร้อมสรุป "สำเร็จ X / ข้าม Y"
- **นำเข้าหนังสือจาก Excel** — พันเล่มต่อรอบขึ้นไป เป็น transaction เดียวรันเบื้องหลัง หน้าต่างไม่ค้าง
- **ตารางหน้าแบ่งชุด (จำกัด resource)** — เปิดแคตตาล็อก 1,000+ เล่มได้ทันที กัน crash ของ Tk
- **เอกสารและบัตร** — บัตรสมาชิก PDF, รายงานยืม-คืน/เข้า-ออก ฟอนต์ไทย Sarabun, ดู PDF ในแอปได้
- **สำรองฐานข้อมูล** — SQLite online backup + ตรวจ integrity/foreign key ก่อนใช้ ใช้ก่อนการเปลี่ยนแปลงสำคัญ
- UI ธีมสว่าง/มืด, toast notifications, confirm dialogs, หน้าต่างสแกนใช้ร่วมกันระหว่างหน้า

### เว็บสำหรับผู้ดูแล (Flask)

- เต็มเทียบเดสก์ท็อป: สมาชิก หนังสือ ยืม-คืน ประวัติ เข้า-ออก ตั้งค่า — รันจากแอป หรือแยกเอง (ผูก `127.0.0.1` ตามค่าเริ่มต้น)
- **ตาราง server-side** (jQuery DataTables — เวอร์ชัน vendored) ค้น/เรียง/แบ่งหน้าที่เซิร์ฟเวอร์ ~50 แถว/หน้า สำหรับแคตตาล็อกหลักพันเล่ม
- **หยุดให้บริการ → ลบถาวรแบบ 2 ขั้น** (retire → purge) ประวัติถูกเก็บอย่างถูกต้องเสมอ
- Import Excel ออนเว็บสูงสุด 5,000 แถว/ไฟล์ พร้อมแม่แบบดาวน์โหลด รายงานแถวซ้ำ/ไม่สมบูรณ์รายตัว
- หน้าแดชบอร์ด Chart.js (ยืม 14 วัน, เข้า-ออก 7 วัน, สัดส่วนสถานะหนังสือ)
- **ทำงานออฟไลน์สมบูรณ์** — ทุกไลบรารี frontend (SweetAlert2, jQuery + DataTables, Alpine.js, Chart.js, Flatpickr) และฟอนต์ Sarabun/Prompt ถูก pin ไว้ใน repo ไม่มี CDN เลย

### หน้าสมาชิก (Member Portal)

- **เข้าสู่ระบบเอง** ด้วย ชั้น + เลขที่ + PIN ที่ตั้งเอง (เก็บเฉพาะ bcrypt hash)
- **สถานะแบบอ่านอย่างเดียว แยกหนึ่งหน้าต่อโมดูล**: หน้าเต็มจอที่มีแถบเมนูข้างด้านซ้าย (ปุ่มพับ/กางได้; บนจอแคบจะเปิดเป็น drawer พร้อมปุ่มปิด) เชื่อมไปยัง `หนังสือที่กำลังยืม` (ป้ายเกินกำหนด/ครบวันนี้) · `รายการคืนล่าสุด` · `ประวัติการเข้าใช้` โดยหน้าหลักยังมีการ์ดสรุปจำนวนพร้อมป้ายเกินกำหนด — ไม่มีปุ่มยืม/คืน การยืม-คืนยังเป็นหน้าที่ของเคาน์เตอร์; ดีเทล UX ก็ครบ: ปุ่มข้ามไปยังเนื้อหาหลัก, โฟกัสคีย์บอร์ดชัดเจน, ตารางเลื่อนแนวนอนในการ์ดบนจอแคบ, สถานะว่างที่มีคำอธิบาย และรองรับ prefers-reduced-motion
- **QR หมุนเวียน (ตั๋ว v3)** — หน้าของสมาชิกแสดง QR ประจำตัวที่เปลี่ยนเองตามรอบเวลา (ผู้ดูแลตั้ง 1–72 ชั่วโมง, เผื่อรอบข้างเคียง 0–2 รอบ) ภาพถ่าย QR เก่า ๆ ใช้ไม่ได้เมื่อรอบเปลี่ยน
- **เปิดใช้งานครั้งแรกแบบ 2 ขั้น** — สแกนบัตรของตนเอง หรือใช้รหัสครั้งเดียวจากผู้ดูแล แล้วตั้ง PIN ภายในหน้าต่าง 10 นาที
- **ควบคุมส่วนกลางได้** — ปิดหน้าสมาชิกทั้งระบบ (`/member*` → 404) และเลือกได้ว่าจะยังรับ QR บนบัตรพลาสติกแบบคงที่หรือไม่
- จำกัดล็อกอินผิด 5 ครั้ง/15 นาที (ต่อ IP + ชั้น + เลขที่) และข้อความผิดพลาดไม่บอกว่า "ไม่มีสมาชิกนี้" กันการเดา

## เริ่มใช้งานอย่างรวดเร็ว

**ข้อกำหนด:** Python **3.10+** (ชุด dependencies ตรวจบน Windows / Python 3.10; CI รันบน 3.11) — เดสก์ท็อปมีเป้าหมายเป็น **Windows 10+** (สแกนผ่าน OpenCV + pyzbar)

```bash
# 1) สร้างสภาพแวดล้อมและติดตั้ง (หรือรัน req.bat แทนได้)
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# 2) เริ่มแอปเดสก์ท็อป
python main.py            # เพิ่ม --debug เพื่อเขียน debug.log
```

**การบูตครั้งแรกเสริมความปลอดภัยให้เอง:** ไม่มีรหัสผ่านดีฟอลต์ ระบบสร้างรหัสผ่านแบบสุ่ม แสดงบนคอนโซล **ครั้งเดียว** และบันทึกไว้ที่ `db/INITIAL_ADMIN_PASSWORD.txt` (gitignored — ควรถือเป็นของลับ) ทุกครั้งที่บูต หากยังพบรหัสเดิม `admin123` ในฐานข้อมูล ระบบจะแทนที่ให้อัตโนมัติ

**เปิดเว็บ:** จากแอป → เมนู "เพิ่มเติม" → "Web UI" → "เริ่ม Web UI" หรือรันแยก:

```bash
venv\Scripts\python -c "from web.app import create_app; create_app().run(port=5000)"
```

เปิด `http://127.0.0.1:5000` — การเปิดบน LAN ต้องใส่ HTTPS จึงปลอดภัย ดู [การปฏิบัติการและบำรุงรักษา](#การปฏิบัติการและบำรุงรักษา)

**ไฟล์ของลับ:** `.env` (gitignored) ถูกสร้างแบบสุ่มอัตโนมัติครั้งแรก ต้องถือเป็นคู่ของฐานข้อมูล — ย้ายเครื่องด้วยกันทั้งชุด ดู [รูปแบบความปลอดภัย](#รูปแบบความปลอดภัย)

> กล้องบนเบราว์เซอร์ใช้ได้ผ่าน `localhost` หรือ HTTPS เท่านั้น (ข้อจำกัดด้านความปลอดภัยของ getUserMedia)

## วงจรชีวิตบัญชีสมาชิก

| ขั้น | ผู้ทำ | รายละเอียดและข้อจำกัด |
|---|---|---|
| 1 · สมัครสมาชิก | ผู้ดูแล | เพิ่มสมาชิก → ระบบออก QR บัตร (v2, เข้ารหัส) และบัตร PDF สำหรับพิมพ์ |
| 2 · เปิดใช้งานครั้งแรก | สมาชิกเอง | ที่ `/member/activate`: สแกน **บัตรของตัวเอง** หรือใส่รหัสครั้งเดียวจากผู้ดูแล (แฮช SHA-256, ใช้ได้ 15 นาที, แสดงบนหน้าจอครั้งเดียว) · แล้วตั้ง PIN 6–64 ตัวอักษรภายใน 10 นาที · บัญชีที่มี PIN แล้วปฏิเสธการเปิดใช้งานซ้ำทุกช่องทาง |
| 3 · ใช้งานปกติ | สมาชิก | ล็อกอินด้วย ชั้น + เลขที่ + PIN → ดูสถานะของตัวเอง + แสดง QR หมุนเวียน (จำกัด 30 ภาพ/นาที) |
| 4 · ที่เคาน์เตอร์ | เจ้าหน้าที่ | สแกน QR ของรอบปัจจุบัน (หรือบัตรคงที่ หากยังเปิดใช้) → ยืม / คืน / เข้า-ออก สแกนเว็บและเดสก์ท็อปใช้แนวปฏิบัติเดียวกัน |
| 5 · ยกเลิก/รีเซ็ต | ผู้ดูแล | ปุ่ม **ล้าง PIN** (ยกเลิกโค้ดค้างหมดในทรานแซกชันเดียว) → สมาชิกต้องยืนยันตัวตนใหม่ด้วยบัตรจริง |

จุดออกแบบสำคัญ: ระบบ **ไม่มี** ทางเปิดใช้งานแทนสมาชิกให้เสร็จได้ — สิทธิ์เข้าหน้าสมาชิกเกิดขึ้นได้เมื่อสมาชิกเป็นคนกระทำตามขั้นที่ 2 ด้วยตัวเองเท่านั้น และการใช้โค้ดของบุคคลอื่นจะไม่ได้ผล (สิทธิ์ผูกกับสมาชิกเจ้าของโค้ด)

## การตั้งค่า

ทั้งหมดแก้ได้ที่หน้า "ตั้งค่า" (เดสก์ท็อปและเว็บใช้ชุดเดียวกัน — `config/settings.json`):

| ตัวเลือก | คำอธิบาย | ค่าเริ่มต้น |
|---|---|---|
| ธีม | System / Light / Dark | System |
| ระยะเวลาการยืมเริ่มต้น | จำนวนวันเมื่อยืม (clamp 1–90 วัน) | 7 วัน |
| จำนวนหนังสือสูงสุดต่อคน | จำกัดยืมพร้อมกัน (clamp 1–20 เล่ม) | 3 เล่ม |
| อายุบัตรสมาชิก | จำนวนวันก่อนบัตรหมดอายุ (clamp 1–3,650 วัน) | 365 วัน |
| ชื่อโรงเรียน | แสดงบนบัตรและหน้าเว็บ | DSNPRU |
| โลโก้โรงเรียน | ไฟล์ PNG สำหรับบัตรและ sidebar | `assets/logos/school_logo.png` |
| Web UI host | `127.0.0.1` (loopback) หรือ `0.0.0.0` (LAN + HTTPS) | `127.0.0.1` |
| เปิดหน้าสมาชิก | ปิดแล้วทุกเส้นทาง `/member*` ตอบ 404 ทันที | เปิด |
| รอบเปลี่ยน QR สมาชิก | ช่วงเวลาต่อรอบของตั๋ว v3 (clamp 1–72 ชม.) | 12 ชั่วโมง |
| เผื่อรอบข้างเคียง | จำนวนรอบเผื่อสำหรับนาฬิกาคลาดเคลื่อน (clamp 0–2) | 1 รอบ |
| รับ QR บัตรเดิม | ยังรับบัตรพลาสติก/กระดาษแบบไม่มีอายุ — แนะนำปิดเมื่อสมาชิกเปิดใช้งานครบ | เปิด |

ค่าตัวเลขถูก clamp ที่เซิร์ฟเวอร์เสมอ (`min`/`max` ของฟอร์มเป็น UX เท่านั้น) และหน้าเว็บแยกฟอร์มบันทึกเป็นกลุ่ม general / member_portal เพื่อไม่ให้กล่องติ๊กที่ถูกละเว้นลบค่าอีกกลุ่ม

## เส้นทาง HTTP

| โซน | เส้นทาง | วิธี | บทบาท |
|---|---|---|---|
| ยืนยันตัวตน | `/login`, `/logout` | GET/POST, GET | ผู้ดูแล |
| หน้าผู้ดูแล | `/`, `/members`, `/books`, `/borrow`, `/return`, `/history`, `/access`, `/settings` | GET | ผู้ดูแล |
| API ตาราง | `/api/members`, `/api/books`, `/api/history`, `/api/stats`, `/api/analytics` | GET | ผู้ดูแล |
| เขียนผู้ดูแล | `/members/add`, `/members/<id>/delete`, `/members/<id>/activation-code`, `/members/<id>/reset-pin`, `/members/bulk`, `/books/add`, `/books/import`, `/books/template`, `/books/<id>/delete`, `/books/<id>/retire`, `/books/<id>/restore`, `/books/bulk` | POST | ผู้ดูแล |
| สแกน | `/scan/resolve` | POST | ผู้ดูแล (ถอดรหัส QR ทุกชนิดฝั่งเซิร์ฟเวอร์) |
| เข้า-ออก | `/access/record` | POST | ผู้ดูแล |
| หน้าสมาชิก | `/member/login`, `/member/logout`, `/member`, `/member/loans`, `/member/returns`, `/member/access`, `/member/qr.png`, `/member/activate`, `/member/activate/scan`, `/member/activate/code`, `/member/activate/set-pin` | GET/POST | สมาชิก (session แยกจากผู้ดูแล) |
| ไฟล์ | `/assets/<path>` | GET | — |

ทุก POST ต้องมี CSRF token (ในฟอร์มหรือ header `X-CSRF-Token`) และเส้นทางที่กีดกันบทบาทจะคืน 302 ไปหน้าล็อกอินที่ต้องการหากไม่มีสิทธิ์

## การปฏิบัติการและบำรุงรักษา

**สำรอง / กู้คืน** — โดยปกติทำผ่านตั้งค่าเดสก์ท็อป หรือ CLI:

```bash
# สำรอง (SQLite online backup ปลอดภัยแม้เซิร์ฟเวอร์กำลังทำงาน) + ตรวจ
venv\Scripts\python -c "from services.backup_service import create_backup, verify_backup; p = create_backup(); print(verify_backup(p), p)"

# กู้คืน (ปฏิเสธไฟล์ที่ตรวจไม่ผ่าน; แทนที่ไฟล์เดสติเนชันแบบ atomic)
venv\Scripts\python -c "from services.backup_service import restore_backup; restore_backup(r'backups\library_backup_YYYYMMDD_HHMMSS.db')"
```

**หมุนเวียนคีย์ QR** — เมื่อสงสัยว่าคีย์รั่วหรือกำลังจะย้ายเครื่อง ให้รันเครื่องมือที่แถมมากับโครงการนี้ ซึ่งหมุนคีย์โดยบัตรที่ออกไปแล้วไม่หาย:

```bash
venv\Scripts\python tools\rotate_qr_key.py [--no-cards]   # backup → ถอดรหัส → เข้ารหัสใหม่ทั้งหมดแบบ v2
venv\Scripts\python tools\regenerate_member_cards.py      # สร้างบัตร PDF ซ้ำจากฐานปัจจุบัน
```

หลังหมุนเวียน: กระจาย `.env` ใหม่ให้ทุกเครื่องทันที **ก่อน**บูตครั้งถัดไป และย้ายโดยคุมชุด (`.env` + `db/` + `assets/`) พร้อมกันเสมอ

**เปิดบน LAN / HTTPS** — ใช้ reverse proxy ยืนหน้าแล้ว terminate TLS แทน จากนั้นตั้ง `webui_host = 0.0.0.0` และเปิด `SESSION_COOKIE_SECURE` — รายละเอียดที่ `docs/SECURITY.md` § 9

**ปัญหาที่พบบ่อยและจุดตรวจ:**

| อาการ | จุดตรวจ |
|---|---|
| กล้องเว็บใช้ไม่ได้ | ต้องเป็น `localhost` หรือ HTTPS; อนุญาตกล้องในเบราว์เซอร์ |
| ล็อกอินสมาชิกคืน 429 | ล้มเหลว 5 ครั้ง/15 นาทีต่อชุดตัวตน รอครบกำหนด หรือให้ผู้ดูแลล้าง PIN |
| บัตรเก่าถอดรหัสไม่ได้ทั้งระบบ | `.env` ไม่ตรงกันระหว่างเครื่อง — จัดชุด (`.env` + `db/` + `assets/`) ให้ตรงกัน หรือรัน `tools/rotate_qr_key.py` |
| หน้าสมาชิกทั้งระบบคืน 404 | มีคนปิด "เปิดหน้าสมาชิก" ในการตั้งค่า |
| ตารางเว็บช้าหรือว่าง | ตรวจว่า migration รันครบ (`PRAGMA user_version` = 4) และ `/api/*` คืนข้อมูล |
| ลืมรหัสผ่าน admin | เปิดไฟล์ `db/INITIAL_ADMIN_PASSWORD.txt`; เคสอื่นดู `docs/SECURITY.md` |

## รูปแบบความปลอดภัย

| ด้าน | แนวทาง |
|---|---|
| ตัวตน QR | **v2** — บัตรคงที่ `AES-256-GCM` (nonce 12B + ป้องกันแก้ไขด้วย tag) · **v3** — ตั๋วหมุนเวียน `AES-256-GCM` + AAD ผูกโปรโตคอล (`pylib-member-ticket-v3`) เนื้อในเป็น JSON ที่ตรวจเข้มทั้งฟิลด์และชนิด (`{"m": id, "e": epoch}`) · **v1 legacy** — AES-CBC อ่านได้เท่านั้นและถอดรหัสแบบเข้มงวด (PKCS7 + UTF-8 เต็มรูปแบบ) — **ไม่มี plaintext fallback** โค้ดผิดรูปแบบถูกปฏิเสธทั้งหมด |
| รอบเวลาของตั๋ว | `epoch = floor(unix_time / รอบ)` — เซิร์ฟเวอร์พิจารณารอบปัจจุบันและเผื่อรอบข้างเคียงเท่านั้น และ **stateless** (ไม่ต้องเก็บสถานะที่ฝั่งใด ๆ) |
| การสร้างคีย์ | PBKDF2-HMAC-SHA256 100,000 รอบ จาก `SECRET_PASSPHRASE` + `SECRET_SALT` (สร้างอัตโนมัติ พร้อมปรับค่า placeholder ให้เป็นค่าสุ่มเสมอ) |
| รหัสผ่านผู้ดูแล | bcrypt, ไม่มีดีฟอลต์ — บูตแรกสุ่มและแสดงครั้งเดียว; ตรวจแบบ timing-equalized; `admin123` ถูกปลดอัตโนมัติทุกบูต |
| รหัส PIN สมาชิก | bcrypt ใน `members.pin_hash` (NULL จนกว่าสมาชิกเปิดใช้งานเอง); การเปิดใช้งานเป็นทรานแซกชันเดียว (อ้างสิทธิ์โค้ด + ตั้ง PIN + เงื่อนไข `pin_hash IS NULL` กันการแทนที่) |
| รหัสเปิดใช้งาน | เก็บเฉพาะแฮช SHA-256, อายุ 15 นาที, ใช้ได้ครั้งเดียว (claim แบบ atomic), แสดงบนหน้าจอครั้งเดียว — ตัวรหัสเปล่าไม่ถูกบันทึกที่ใดเลย |
| การจำกัดการล็อกอิน | 5 ครั้ง / 15 นาที แยกขอบเขตอิสระ: ผู้ดูแล (IP + username) · สมาชิก (IP + ชั้น + เลขที่) |
| เซสชันเว็บ | คุกกี้ HttpOnly + SameSite=Lax อายุ 4 ชั่วโมง, CSRF token ทุก POST (เทียบแบบ constant-time), ข้อความล้มเหลวกัน user enumeration |
| ขอบเขตสิทธิ์ | หน้าสมาชิกอ่านเฉพาะข้อมูลของตนเอง ไม่มีเส้นทางแก้ไขข้อมูลใด ๆ; ปิดทั้งโดเมนได้ (kill switch) |
| ความถูกต้องข้อมูล | FK เปิดทุก connection, `books.code UNIQUE`, `borrow_log.returned CHECK (0,1)`, ยืม/คืนเป็นทรานแซกชัน `BEGIN IMMEDIATE` เดียว, parameterized SQL ทุกจุด |
| Migrations | `PRAGMA user_version` v1→v4 (constraints → ดัชนี bulk-import → หน้าสมาชิก: pin_hash + ตารางรหัสเปิดใช้งาน) — รันทุกบูตและ idempotent |
| การบันทึก log | ไม่เขียนข้อมูลส่วนบุคคลลง log (เฉพาะ id และจำนวน); เคส payload ที่ล้มเหลวคืนเฉพาะผลลัพธ์ไม่สะท้อนเนื้อหาจริง |

เอกสารเชิงปฏิบัติฉบับเต็ม (รูปแบบ payload ทุกเวอร์ชัน, runbook หมุนเวียนคีย์, deployment HTTPS, เมทริกซ์การทดสอบความปลอดภัย): [`docs/SECURITY.md`](docs/SECURITY.md)

## ฐานข้อมูล

SQLite3 ไฟล์เดียว (`db/library.db`) — โครงสร้างถูกจัดการด้วย migration รายเวอร์ชัน:

| ตาราง | เนื้อหา |
|---|---|
| `admin_users` | ผู้ดูแลระบบ (bcrypt, ไม่มีดีฟอลต์) |
| `members` | ชื่อ, ชั้น, เลขที่, QR, วันหมดอายุ, **`pin_hash`** (v4 — NULL จนสมาชิกเปิดใช้งาน); badge ชั้น+เลขที่ ต้องไม่ซ้ำ |
| `books` | รหัส (**UNIQUE**), ชื่อ, สถานะ `ว่าง` / `ยืมแล้ว` / `หยุดให้บริการ` |
| `borrow_log` | ทรานแซกชันยืม-คืน — `returned CHECK (0,1)`, FK ไป members/books, ดัชนีสำหรับ bulk (v3) |
| `access_log` | บันทึกเข้า-ออกห้องสมุด |
| `member_activation_codes` | (v4) แฮช SHA-256 ของรหัสเปิดใช้งาน, อายุ 15 นาที, `code_hash UNIQUE`, `used_at` บันทึกเวลาที่ถูกใช้ (ใช้ได้ครั้งเดียว) |

## การทดสอบและ CI

```bash
venv\Scripts\python -m unittest discover -s tests        # Windows
venv/bin/python -m unittest discover -s tests            # Linux/Mac
```

**193 กรณีทดสอบ รวม 18 ไฟล์** — ครอบคลุม crypto (v1/v2/v3), ตัวตรวจตัวตน (member identity), หน้าสมาชิกครบวงจร, repository และการยืม-คืนแบบ atomic, migration v1→v4, สำรอง/กู้คืน, bulk + import Excel, API ตาราง server-side, guard ของเว็บ (CSRF/throttle/session/สิทธิ์) และ offline-safety

CI (GitHub Actions) รัน 3 งานแยกกันทุก push/PR:

| Job | สิ่งที่ทำ |
|---|---|
| `tests + lint` | ชุดทดสอบทั้งหมด (unittest) + `ruff check` (กำหนดเองใน `pyproject.toml`, รวม rule bandit) |
| `pip-audit` | สแกน CVE ของ dependency ทุกตัวที่ติดตั้ง |
| `gitleaks` | ค้นหา secret ที่ commit เข้าโค้ด |

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
│   ├── security.py              # bcrypt hash/verify, AES-GCM `v2:` badges + `v3:` rotating tickets + legacy AES-CBC (PBKDF2)
│   ├── logger.py                # logging config (debug.log support)
│   └── throttle.py              # Sliding-window login lockout (desktop + web)
├── db/                          # (runtime artifacts: library.db, INITIAL_ADMIN_PASSWORD.txt — gitignored)
│   ├── schema.py                # CREATE TABLE + user_version migrations (v1→v4) + admin bootstrap
│   └── repositories/
│       ├── base_repo.py         # Thread-safe SQLite (per-query connection + Lock)
│       ├── member_repo.py       # Member CRUD (+ PIN bcrypt, activation codes, atomic activation)
│       ├── book_repo.py         # Book CRUD (uses book status constants)
│       ├── borrow_repo.py       # try_borrow / try_return — atomic invariants
│       ├── access_repo.py       # Access log
│       └── admin_repo.py        # admin_users + timing-equalized verify_credentials
├── services/
│   ├── qr_service.py            # สร้าง/ถอดรหัส QR (AES)
│   ├── member_identity.py       # Resolver เดียวสำหรับ QR ทุกชนิด (v3 ตั๋ว/v2 บัตร/legacy) + นโยบายหน้าต่างเวลา
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
│   │   ├── member_view.py       # Two-column: form + searchable list, PIN tools
│   │   ├── book_view.py         # Two-column: form + searchable list, retire/restore
│   │   ├── borrow_view.py       # 3-step flow (scan → book → confirm)
│   │   ├── return_view.py       # Scan + borrowed list + return
│   │   ├── history_view.py      # Table search + PDF export
│   │   ├── access_view.py       # Scanner + history (AccessScannerView/AccessHistoryView)
│   │   ├── settings_view.py     # Theme, loan settings, member portal, school info, DB backup
│   │   └── webui_view.py        # Web UI launcher control panel
│   └── widgets/
│       ├── confirm_dialog.py    # Modal confirm dialog
│       ├── toast.py             # Floating toast notifications
│       ├── scan_window.py       # Shared camera QR scan window
│       └── pdf_viewer.py        # Built-in PDF viewer (PyMuPDF)
├── web/
│   ├── app.py                   # Flask app: routes, guards, DataTables /api feeds, bulk, analytics, member portal, Excel import
│   ├── csrf.py                  # Per-session CSRF tokens (form field หรือ X-CSRF-Token header)
│   ├── static/
│   │   ├── style.css            # Web UI design system
│   │   ├── vendor/              # Vendored JS libs, pinned versions (see README.md inside)
│   │   │   ├── alpine.min.js    # Alpine.js 3.17.5
│   │   │   ├── jquery.min.js    # jQuery 3.7.1
│   │   │   ├── datatables.min.js # DataTables 1.13.8 (server-side tables)
│   │   │   ├── sweetalert2.min.js/.css  # SweetAlert2 11.26.25
│   │   │   ├── chart.umd.min.js # Chart.js 4.6.1
│   │   │   ├── flatpickr.min.*  # Flatpickr 4.6.13 + Thai locale (th)
│   │   │   └── htmx.min.js      # htmx 2.0.11 (retained, currently unused)
│   │   ├── fonts/               # Sarabun + Prompt TTFs (OFL) — @font-face in style.css, offline
│   │   ├── lang/
│   │   │   └── datatables-th.json # hand-made Thai i18n for DataTables
│   │   └── js/
│   │       ├── jsQR.js          # Vendored jsQR (in-browser QR decode)
│   │       ├── scan.js          # Camera capture + /scan/resolve calls
│   │       ├── member_scan.js   # กล้องสแกนบัตร บนหน้าเปิดใช้งานสมาชิก
│   │       ├── member_qr.js     # นับเวลาถอยหลัง + สลับ QR เมื่อรอบเปลี่ยน
│   │       ├── member_layout.js # ปุ่มพับ/กาง sidebar ของหน้าสมาชิก (จำค่าไว้, drawer บนมือถือ)
│   │       ├── tables.js        # DataTables configs (books/members/history) + bulk bar
│   │       └── ui.js            # PyLib helpers: dialogs, toasts, data-confirm, flatpickr init
│   └── templates/
│       ├── icons.svg            # Lucide SVG icon sprite
│       ├── base.html            # Layout + sidebar + vendor stack wiring
│       ├── login.html           # Login page
│       ├── member_base.html     # หน้าสมาชิก (เต็มจอ + sidebar เมนูข้าง + ปุ่มพับ/กาง)
│       ├── member_login.html    # สมาชิกเข้าสู่ระบบ (ชั้น+เลขที่+PIN)
│       ├── member_activate.html # เปิดใช้งานครั้งแรก (สแกนบัตร หรือรหัสแอดมิน) + ตั้ง PIN
│       ├── member_home.html     # หน้าหลักสมาชิก: สถานะ + QR หมุนเวียนพร้อมนับเวลา + ชิปสรุป
│       ├── member_loans.html    # หนังสือที่กำลังยืม (หน้าโมดูลแยก)
│       ├── member_returns.html  # รายการคืนล่าสุด (หน้าโมดูลแยก)
│       ├── member_access.html   # ประวัติการเข้าใช้ห้องสมุด (หน้าโมดูลแยก)
│       ├── index.html           # Dashboard + Chart.js widgets
│       ├── members.html         # Member management
│       ├── books.html           # Book management (retire/restore/purge)
│       ├── borrow.html          # Borrow (Alpine lock state + Flatpickr due date)
│       ├── return_book.html     # Return with overdue badges
│       ├── history.html         # Borrow history
│       ├── access.html          # Access recording + history
│       └── settings.html        # Settings + about
├── tools/
│   ├── rotate_qr_key.py         # Key rotation: backup → decrypt → re-encrypt
│   └── regenerate_member_cards.py
├── tests/                       # _harness.py + 18 test files (ดูหัวข้อ การทดสอบและ CI)
├── docs/
│   ├── SECURITY.md              # Security runbook (key rotation, HTTPS/LAN, payload formats)
│   └── DESIGN.md                # Visual design system notes
├── assets/
│   ├── fonts/                   # Sarabun + Prompt TTFs (desktop PDFs + copied to web/static/fonts)
│   ├── logos/                   # school_logo.png
│   ├── qrcodes/                 # Generated QR images (runtime, gitignored)
│   └── cards/                   # Generated member card PDFs (runtime, gitignored)
├── .env.example                 # SECRET_PASSPHRASE, SECRET_SALT, FLASK_SECRET_KEY template
├── .env                         # (runtime, gitignored, auto-generated)
├── pyproject.toml               # ruff lint/style configuration
├── requirements.txt             # pinned versions
└── README.md
```

## แพ็กเกจที่ใช้

| แพ็กเกจ | บทบาท |
|---|---|
| `customtkinter` | GUI เดสก์ท็อป (Modern Tkinter) |
| `flask` | เว็บ UI |
| `opencv-python` + `numpy` | กล้องเว็บแคม + กระบวนการภาพ (เดสก์ท็อป) |
| `pyzbar` | ถอดรหัส QR จากภาพกล้อง (เดสก์ท็อป) |
| `qrcode` | สร้างรูป QR |
| `cryptography` | AES-256-GCM / AES-256-CBC, PBKDF2 |
| `bcrypt` | แฮชรหัสผ่าน admin / PIN สมาชิก |
| `reportlab` | บัตรสมาชิกและรายงาน PDF |
| `pymupdf` | ดู PDF ในแอป |
| `pillow` | จัดการรูปภาพ |
| `pandas` + `openpyxl` | นำเข้า/ส่งออก Excel |
| `python-dotenv` | โหลด `.env` |
| `packaging` | dependency ของ customtkinter |

เวอร์ชันถูก pin ทั้งหมดใน `requirements.txt` (ไฟล์ระบุไว้เองว่าตรวจบน Windows / Python 3.10)

## แผนการพัฒนา

สถานะ: ทุกงานหลักใช้งานได้จริง — รายการต่อไปอยู่บนแผน:

- [ ] สร้างบัตร/รายงาน PDF จากเว็บ (ปัจจุบันทำได้บนเดสก์ท็อปเท่านั้น)
- [ ] สมาชิกเปลี่ยน PIN ของตนเองในหน้าของตัว (ปัจจุบันรีเซ็ตโดยผู้ดูแล)
- [ ] สำรองตามเวลาอัตโนมัติ + นโยบายเก็บสำรองย้อนหลัง
- [ ] จุดเช็กสถานะแบบไม่ต้องล็อกอิน (kiosk)

พบบั๊กหรือมีข้อเสนอ เปิด [issue](https://github.com/Nonx2360/The-PY-LIB-Project/issues) ได้โดยตรง

## การมีส่วนร่วม

1. Fork แล้วตั้ง branch (`feature/…` หรือ `fix/…`)
2. `ruff check .` ต้องผ่านทั้งหมด — กำหนดเองใน `pyproject.toml`
3. ทุกการเปลี่ยนแปลงต้องไม่เสียชุดทดสอบ: `python -m unittest discover -s tests`
4. การแตะด้านความปลอดภัย (การเข้ารหัส / สิทธิ์ / session) ควรเพิ่มกรณีทดสอบ regression ให้ตรงเมทริกซ์ใน [`docs/SECURITY.md`](docs/SECURITY.md) § 11
5. เปิด PR พร้อมเหตุผลและภาพหน้าจอ (หากแตะ UI)

ยังเปิดรับข้อเสนอที่ไม่ใช่โค้ด: เปิด issue พร้อมเล่าสถานการณ์การใช้งานจริงในห้องสมุดของคุณ

## ทีมและใบอนุญาต

- เด็กหญิง ขวัญชนก อุ่นศิริ เลขที่ 4 (UX/UI)
- เด็กชาย ณัฐชนน รอดน้อย เลขที่ 30 (Developer)
- **โรงเรียน:** DSNPRU

ขอบคุณ — ผู้ทดสอบ beta และอาจารย์ที่ปรึกษา

MIT License © 2025
