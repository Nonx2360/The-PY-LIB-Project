<div align="center">

# THE PY-LIB

**ระบบจัดการห้องสมุดโรงเรียนแบบครบวงจร — แอปเดสก์ท็อป, เว็บสำหรับผู้ดูแล, หน้าสมาชิก และตู้บริการตนเอง (Kiosk)**
QR ประจำตัวเข้ารหัสด้วย AES-256-GCM (รองรับตั๋วหมุนเวียนแบบมีอายุ) • ยืม-คืนแบบ atomic • จองหนังสือ (ระบบคิว/ข้อเสนอ) • ย้ายปีการศึกษา • ตรวจนับสต็อก • 214 เทสเคส + CI

[![CI](https://github.com/Nonx2360/The-PY-LIB-Project/actions/workflows/ci.yml/badge.svg)](https://github.com/Nonx2360/The-PY-LIB-Project/actions/workflows/ci.yml)
![Python](https://img.shields.io/badge/Python-3.10%2B-306998?logo=python&logoColor=FFD43B)
![Platform](https://img.shields.io/badge/Desktop-Windows%2010%2B-0078D6?logo=windows&logoColor=white)
![Tests](https://img.shields.io/badge/tests-214%20passing-2EA44F)
![License](https://img.shields.io/badge/license-MIT-8A6BE8)

</div>

---

THE PY-LIB คือระบบจัดการห้องสมุดที่สร้างขึ้นมาเพื่อใช้จริงในโรงเรียน ครอบคลุมตั้งแต่การจัดการสมาชิก หนังสือ และการยืม-คืน ผ่านทั้งแอปเดสก์ท็อปและเว็บสำหรับผู้ดูแล รวมถึง **หน้าสมาชิก** ที่ให้นักเรียนเข้าดูสถานะการยืมและแสดง **QR หมุนเวียนที่มีอายุตามรอบเวลา** โดยไม่ต้องพึ่งพาบัตรพลาสติกใบเดิมตลอดไป ทุกชั้น ทั้งการเข้ารหัส, สิทธิ์การเข้าถึง และการทำทรานแซกชัน ถูกออกแบบมาให้ทำงานร่วมกันและครอบคลุมด้วยชุดทดสอบความปลอดภัยที่รันอัตโนมัติใน CI

## สารบัญ

- [สถาปัตยกรรมและสิทธิ์การเข้าถึง](#สถาปัตยกรรมและสิทธิ์การเข้าถึง)
- [คุณสมบัติ](#คุณสมบัติ)
- [เริ่มใช้งานอย่างรวดเร็ว](#เริ่มใช้งานอย่างรวดเร็ว)
- [วงจรชีวิตบัญชีสมาชิก](#วงจรชีวิตบัญชีสมาชิก)
- [การตั้งค่า](#การตั้งค่า)
- [เส้นทาง HTTP](#เส้นทาง-http)
- [การดูแลและบำรุงรักษา](#การดูแลและบำรุงรักษา)
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
│  jsQR → เซิร์ฟเวอร์        │    │  services/qr + member_identity │    │  เห็นเฉพาะข้อมูลตัวเอง    │
│  bulk + analytics      │    │  core/security (AES-GCM)       │    │  ตั๋วหมุนเวียน v3          │
└────────────────────────┘    │  config/settings + .env        │    └────────────────────────┘
                              └─────────┬──────────────────────┘
                                        │
                              ┌─────────▼─────────┐
                              │   SQLite (db/)    │
                              └───────────────────┘  user_version v1→v7 อัตโนมัติ (idempotent)
```

**เมทริกซ์สิทธิ์** — แต่ละหน้าถูกกำหนดชัดเจนว่าใครเข้าถึงได้:

| ความสามารถ | ผู้ดูแล (เดสก์ท็อป) | ผู้ดูแล (เว็บ) | สมาชิก (หน้าเว็บ) |
|---|:---:|:---:|:---:|
| เพิ่ม/แก้ไข/ลบ สมาชิก + หนังสือ, จัดการแบบกลุ่ม | ✓ | ✓ | ✗ |
| ยืม • คืน • บันทึกเข้า-ออกห้องสมุด | ✓ | ✓ | ✗ |
| ออกรหัสเปิดใช้งาน / ล้าง PIN ของสมาชิก | ✓ | ✓ | ✗ |
| ปรับตั้งค่าระบบ (รวมถึงเปิด/ปิดหน้าสมาชิก) | ✓ | ✓ | ✗ |
| ดูข้อมูลของตัวเอง + QR หมุนเวียนของตัวเอง | — | — | ✓ |

Session ของผู้ดูแลและสมาชิกใช้คีย์คนละชุดในคุกกี้เดียวกัน ดังนั้นการล็อกอินฝั่งหนึ่งจะไม่ให้สิทธิ์อีกฝั่ง และมีชุดทดสอบเฉพาะตรวจสอบเรื่องนี้ไว้

## คุณสมบัติ

### แอปเดสก์ท็อป (สำหรับผู้ดูแล)

- **ระบุตัวตนสมาชิกด้วย QR เข้ารหัส** — สแกนผ่านเว็บแคมเพื่อยืม คืน และบันทึกเข้า-ออก โดยถอดรหัสบนเครื่องด้วยคีย์จาก `.env` เท่านั้น
- **ยืม-คืนแบบ atomic** — หนึ่งทรานแซกชัน `BEGIN IMMEDIATE` ต่อการยืม/คืน: หนังสือเปลี่ยนสถานะด้วย conditional UPDATE ที่ทำให้มีผู้ชนะได้เพียงคนเดียวแม้สแกนพร้อมกันจากสองเครื่อง และจะไม่มีข้อมูลค้างแม้เกิดข้อผิดพลาดกลางทาง
- **จัดการแบบกลุ่ม** — เลือกหนังสือหรือสมาชิกหลายรายการแล้วหยุดให้บริการ / คืนสถานะ / ลบพร้อมกัน รายการที่มีประวัติจะถูกข้ามอัตโนมัติพร้อมสรุป "สำเร็จ X / ข้าม Y"
- **นำเข้าหนังสือจาก Excel** — นำเข้าได้หลายพันเล่มต่อครั้งในทรานแซกชันเดียว รันเบื้องหลังโดยไม่ทำให้หน้าต่างค้าง
- **ตารางแบบแบ่งหน้า (ประหยัด resource)** — เปิดแคตตาล็อกกว่า 1,000 เล่มได้ทันที ป้องกัน Tk crash
- **เอกสารและบัตร** — บัตรสมาชิก PDF, รายงานยืม-คืน/เข้า-ออก ใช้ฟอนต์ไทย Sarabun, ดู PDF ได้ในแอปเลย
- **สำรองฐานข้อมูล** — SQLite online backup พร้อมตรวจสอบ integrity และ foreign key ก่อนใช้ ควรทำก่อนการเปลี่ยนแปลงสำคัญทุกครั้ง
- UI ธีมสว่าง/มืด, toast notifications, confirm dialogs และหน้าต่างสแกนที่ใช้ร่วมกันได้ระหว่างหน้า

### เว็บสำหรับผู้ดูแล (Flask)

- ครบเทียบเท่าเดสก์ท็อป: สมาชิก หนังสือ ยืม-คืน ประวัติ เข้า-ออก ตั้งค่า — รันจากแอปหรือแยกเองก็ได้ (ผูกกับ `127.0.0.1` เป็นค่าเริ่มต้น)
- **ตาราง server-side** (jQuery DataTables เวอร์ชัน vendored) ค้นหา/เรียง/แบ่งหน้าที่เซิร์ฟเวอร์ ประมาณ 50 แถว/หน้า รองรับแคตตาล็อกหลักพันเล่ม
- **หยุดให้บริการ → ลบถาวรแบบ 2 ขั้น** (retire → purge) ประวัติถูกรักษาไว้ถูกต้องเสมอ
- นำเข้า Excel บนเว็บได้สูงสุด 5,000 แถว/ไฟล์ พร้อมแม่แบบให้ดาวน์โหลด และรายงานแถวซ้ำ/ไม่สมบูรณ์รายตัว
- แดชบอร์ด Chart.js (ยืม 14 วัน, เข้า-ออก 7 วัน, สัดส่วนสถานะหนังสือ)
- **ย้ายปีการศึกษา / เลื่อนชั้น (Academic Year Migration)** — มี dry-run แบบอ่านอย่างเดียวให้ตรวจก่อนเสมอ, ตรวจทานรายบุคคล (เลื่อนชั้น/คงเดิม/จบการศึกษา/ออก/ตรวจสอบเอง), ส่งออก CSV, ยืนยันด้วยการพิมพ์ข้อความ, สำรองและตรวจสอบก่อนบันทึก, ทำงานในทรานแซกชันเดียวพร้อมบันทึกประวัติ (ไม่เขียนทับประวัติยืม/เข้าออกเดิม รหัสสมาชิกยังคงเดิม และสามารถย้อนกลับได้โดยกู้คืนจากไฟล์สำรอง)
- **ตรวจนับหนังสือ (Inventory Audit)** — เปิด/พัก/ทำต่อ/ปิดรอบตรวจนับ, สแกนหรือกรอกรหัส (resolve ฝั่งเซิร์ฟเวอร์), จับของซ้ำ/รหัสไม่รู้จัก, แยกแยะ "กำลังถูกยืม" กับ "ขาดหายจากชั้น", บันทึกการจัดการความคลาดเคลื่อนและส่งออกรายงาน — เป็นการกระทบยอดเท่านั้น ไม่เปลี่ยนสถานะหรือลบข้อมูลโดยอัตโนมัติ
- **การจองหนังสือ (Reservations)** — จองแบบรายเล่ม คิว FIFO ชัดเจน เมื่อคืนหรือว่างจะสร้างข้อเสนอ (offer) พร้อมวันหมดเวลาแบบปลอดภัยทางธุรกรรม สมาชิก/เจ้าหน้าที่ต้อง "รับ/ยืม" อย่างชัดเจน (ไม่มีการยืมอัตโนมัติ) และจัดการคิว/ยกเลิกได้ทั้งจากผู้ดูแลและหน้าสมาชิก
- **อัปเดตแบบเรียลไทม์ (WebSocket)** — ตารางผู้ดูแล/แดชบอร์ด/หน้าสมาชิกรีเฟรชข้อมูลเองโดยไม่ต้องกด F5 การเปลี่ยนแปลงจากทั้งเว็บและแอปเดสก์ท็อปจะถูกตรวจจับผ่าน WebSocket ภายในประมาณ 2 วินาที
- **ทำงานออฟไลน์ได้สมบูรณ์** — ไลบรารี frontend ทุกตัว (SweetAlert2, jQuery + DataTables, Alpine.js, Chart.js, Flatpickr) และฟอนต์ Sarabun/Prompt ถูก pin ไว้ใน repo ไม่พึ่ง CDN เลย

### หน้าสมาชิก (Member Portal)

- **เข้าสู่ระบบเอง** ด้วยชั้น + เลขที่ + PIN ที่ตั้งเอง (เก็บเฉพาะ bcrypt hash)
- **สถานะแบบอ่านอย่างเดียว แยกเป็นโมดูล**: QR ประจำตัว (หมุนเวียน v3), หนังสือที่กำลังยืม, รายการคืนล่าสุด, ประวัติการเข้าใช้ห้องสมุด และ **การจองหนังสือ** (ดูคิว/ข้อเสนอ, จองตามรหัส, รับหนังสือเมื่อถึงคิว, ยกเลิก)
- **QR หมุนเวียน (ตั๋ว v3)** — เปลี่ยนอัตโนมัติตามรอบเวลา ภาพถ่าย QR เก่าจะใช้ไม่ได้เมื่อรอบหมุนไปแล้ว
- เปิดใช้งานครั้งแรกแบบ 2 ขั้น, จำกัดจำนวนครั้งล็อกอินผิด และควบคุมจากส่วนกลางได้

### ตู้บริการตนเอง (Self-service Kiosk)

- เส้นทางแยก `/kiosk` (session แยกจากผู้ดูแล/หน้าสมาชิก) สำหรับยืม-คืนด้วยตนเอง ระบุตัวตนด้วยการสแกนบัตร สแกนหรือกรอกรหัสหนังสือ และใช้บริการยืม-คืนแบบ atomic เดียวกับระบบหลัก
- รีเซ็ตอัตโนมัติเมื่อไม่มีการใช้งาน, จำกัดจำนวนครั้งสแกน และแสดงเฉพาะข้อมูลที่จำเป็น (ไม่เปิดหน้าหรือสิทธิ์ผู้ดูแล)

## เริ่มใช้งานอย่างรวดเร็ว

**ข้อกำหนด:** Python **3.10+** (ตรวจสอบ dependencies บน Windows / Python 3.10 แล้ว; CI รันบน 3.11) — เดสก์ท็อปรองรับ **Windows 10+** (สแกนผ่าน OpenCV + pyzbar)

```bash
# 1) สร้าง virtual environment และติดตั้ง (หรือรัน req.bat แทนก็ได้)
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# 2) เปิดแอปเดสก์ท็อป
python main.py            # เพิ่ม --debug เพื่อบันทึก debug.log
```

**บูตครั้งแรกจัดการความปลอดภัยให้เองอัตโนมัติ:** ไม่มีรหัสผ่านเริ่มต้น ระบบจะสุ่มสร้างรหัสผ่านขึ้นมา แสดงบนคอนโซล **ครั้งเดียว** และบันทึกไว้ที่ `db/INITIAL_ADMIN_PASSWORD.txt` (gitignored — ควรเก็บเป็นความลับ) นอกจากนี้ทุกครั้งที่บูต หากระบบพบรหัสผ่าน `admin123` ในฐานข้อมูลจะแทนที่ให้อัตโนมัติ

**เปิดเว็บ UI:** ไปที่เมนู "เพิ่มเติม" → "Web UI" → "เริ่ม Web UI" ในแอป หรือรันแยกเองด้วย:

```bash
venv\Scripts\python -c "from web.app import create_app; create_app().run(port=5000)"
```

เปิดที่ `http://127.0.0.1:5000` — หากต้องการเปิดบน LAN ต้องใช้ HTTPS จึงจะปลอดภัย ดูรายละเอียดที่ [การดูแลและบำรุงรักษา](#การดูแลและบำรุงรักษา)

**ไฟล์ความลับ:** `.env` (gitignored) จะถูกสร้างแบบสุ่มอัตโนมัติในครั้งแรก ต้องถือว่าเป็นคู่กับฐานข้อมูล — ย้ายเครื่องต้องนำทั้งชุดไปด้วย ดูที่ [รูปแบบความปลอดภัย](#รูปแบบความปลอดภัย)

> กล้องในเบราว์เซอร์ใช้ได้เฉพาะผ่าน `localhost` หรือ HTTPS เท่านั้น (ข้อจำกัดด้านความปลอดภัยของ getUserMedia)

## วงจรชีวิตบัญชีสมาชิก

| ขั้น | ผู้ดำเนินการ | รายละเอียดและข้อจำกัด |
|---|---|---|
| 1 · สมัครสมาชิก | ผู้ดูแล | เพิ่มสมาชิก → ระบบออก QR บัตร (v2, เข้ารหัส) และบัตร PDF สำหรับพิมพ์ |
| 2 · เปิดใช้งานครั้งแรก | สมาชิกเอง | ที่ `/member/activate`: สแกน **บัตรของตัวเอง** หรือใส่รหัสครั้งเดียวจากผู้ดูแล (แฮช SHA-256, ใช้ได้ภายใน 15 นาที, แสดงบนหน้าจอเพียงครั้งเดียว) · จากนั้นตั้ง PIN 6–64 ตัวอักษรภายใน 10 นาที · บัญชีที่ตั้ง PIN แล้วจะปฏิเสธการเปิดใช้งานซ้ำทุกช่องทาง |
| 3 · ใช้งานปกติ | สมาชิก | ล็อกอินด้วยชั้น + เลขที่ + PIN → ดูสถานะของตัวเอง + แสดง QR หมุนเวียน (จำกัด 30 ภาพ/นาที) |
| 4 · ที่เคาน์เตอร์ | เจ้าหน้าที่ | สแกน QR ของรอบปัจจุบัน (หรือบัตรคงที่ หากยังเปิดใช้อยู่) → ยืม / คืน / เข้า-ออก ทั้งเว็บและเดสก์ท็อปใช้ขั้นตอนเดียวกัน |
| 5 · ยกเลิก/รีเซ็ต | ผู้ดูแล | กดปุ่ม **ล้าง PIN** (ยกเลิกโค้ดค้างทั้งหมดในทรานแซกชันเดียว) → สมาชิกต้องยืนยันตัวตนใหม่ด้วยบัตรจริง |

สิ่งที่ออกแบบไว้ชัดเจน: ระบบ **ไม่มีทาง** ให้ผู้ดูแลเปิดใช้งานแทนสมาชิกได้จนเสร็จ — สิทธิ์เข้าหน้าสมาชิกจะเกิดขึ้นได้ก็ต่อเมื่อสมาชิกทำขั้นที่ 2 ด้วยตัวเองเท่านั้น และการใช้รหัสของคนอื่นจะไม่มีผล เพราะสิทธิ์ผูกกับเจ้าของรหัสนั้น

## การตั้งค่า

ทั้งหมดปรับได้ที่หน้า "ตั้งค่า" (เดสก์ท็อปและเว็บใช้ชุดเดียวกัน — `config/settings.json`):

| ตัวเลือก | คำอธิบาย | ค่าเริ่มต้น |
|---|---|---|
| ธีม | System / Light / Dark | System |
| ระยะเวลาการยืมเริ่มต้น | จำนวนวันเมื่อยืม (clamp 1–90 วัน) | 7 วัน |
| จำนวนหนังสือสูงสุดต่อคน | จำกัดยืมพร้อมกัน (clamp 1–20 เล่ม) | 3 เล่ม |
| อายุบัตรสมาชิก | จำนวนวันก่อนบัตรหมดอายุ (clamp 1–3,650 วัน) | 365 วัน |
| ชื่อโรงเรียน | แสดงบนบัตรและหน้าเว็บ | DSNPRU |
| โลโก้โรงเรียน | ไฟล์ PNG สำหรับบัตรและ sidebar | `assets/logos/school_logo.png` |
| Web UI host | `127.0.0.1` (loopback) หรือ `0.0.0.0` (LAN + HTTPS) | `127.0.0.1` |
| เปิดหน้าสมาชิก | ถ้าปิด เส้นทาง `/member*` ทั้งหมดจะตอบ 404 ทันที | เปิด |
| รอบเปลี่ยน QR สมาชิก | ช่วงเวลาต่อรอบของตั๋ว v3 (clamp 1–72 ชม.) | 12 ชั่วโมง |
| เผื่อรอบข้างเคียง | จำนวนรอบเผื่อสำหรับนาฬิกาที่คลาดเคลื่อน (clamp 0–2) | 1 รอบ |
| รับ QR บัตรเดิม | ยังรับบัตรพลาสติก/กระดาษแบบไม่มีอายุ — แนะนำให้ปิดเมื่อสมาชิกเปิดใช้งานครบแล้ว | เปิด |
| เปิดตู้บริการตนเอง (Kiosk) | เปิด/ปิดเส้นทาง `/kiosk` | เปิด |
| เวลาหมดอายุ Kiosk | รีเซ็ต session อัตโนมัติเมื่อไม่ใช้งาน (วินาที) | 90 |
| อายุข้อเสนอการจอง (ชั่วโมง) | เมื่อหนังสือว่างและถึงคิว ระบบจะรอให้รับภายในเวลานี้ (ต้องรับ/ยืมจริง) | 48 |

ค่าตัวเลขถูก clamp ที่เซิร์ฟเวอร์เสมอ (`min`/`max` ของฟอร์มเป็นแค่ UX) และหน้าเว็บแยกฟอร์มบันทึกเป็นกลุ่ม general / member_portal เพื่อป้องกันไม่ให้ checkbox ที่ไม่ได้ติ๊กไปลบค่าของกลุ่มอื่น

## เส้นทาง HTTP

| โซน | เส้นทาง | วิธี | บทบาท |
|---|---|---|---|
| ยืนยันตัวตน | `/login`, `/logout` | GET/POST, GET | ผู้ดูแล |
| หน้าผู้ดูแล | `/`, `/members`, `/books`, `/borrow`, `/return`, `/history`, `/access`, `/academic-year-migration`, `/inventory-audit`, `/reservations`, `/settings` | GET | ผู้ดูแล |
| API ตาราง/สถิติ | `/api/members`, `/api/books`, `/api/history`, `/api/stats`, `/api/analytics` | GET | ผู้ดูแล |
| เขียนผู้ดูแล | `/members/add`, `/members/<id>/delete`, `/members/<id>/activation-code`, `/members/<id>/reset-pin`, `/members/bulk`, `/books/add`, `/books/import`, `/books/template`, `/books/<id>/delete`, `/books/<id>/retire`, `/books/<id>/restore`, `/books/bulk` | POST | ผู้ดูแล |
| ย้ายปีการศึกษา | `/academic-year-migration`, `/academic-year-migration/export`, `/academic-year-migration/apply` | GET/POST | ผู้ดูแล (preview + สำรองก่อนบันทึก) |
| ตรวจนับหนังสือ | `/inventory-audit`, `/inventory-audit/create`, `/inventory-audit/<id>/status`, `/inventory-audit/<id>/scan`, `/inventory-audit/<id>/resolve`, `/inventory-audit/<id>/export` | GET/POST | ผู้ดูแล (กระทบยอดเท่านั้น) |
| การจอง (ผู้ดูแล) | `/reservations` | GET/POST | ผู้ดูแล |
| สแกน | `/scan/resolve` | POST | ผู้ดูแล (ถอดรหัส QR ทุกชนิดฝั่งเซิร์ฟเวอร์) |
| ช่องเรียลไทม์ | `/admin/ws`, `/member/ws` | GET (WebSocket) | ผู้ดูแล / สมาชิก ตาม session (same-origin เท่านั้น) |
| เข้า-ออก | `/access/record` | POST | ผู้ดูแล |
| หน้าสมาชิก | `/member/login`, `/member/logout`, `/member`, `/member/loans`, `/member/returns`, `/member/access`, `/member/reservations`, `/member/qr.png`, `/member/activate`, `/member/activate/scan`, `/member/activate/code`, `/member/activate/set-pin` | GET/POST | สมาชิก (session แยกจากผู้ดูแล) |
| ตู้บริการตนเอง (Kiosk) | `/kiosk`, `/kiosk/resolve`, `/kiosk/borrow`, `/kiosk/return`, `/kiosk/reset` | GET/POST | สาธารณะ/จุดบริการ (session แยก, ไม่ใช่ผู้ดูแล) |
| ไฟล์ | `/assets/<path>` | GET | — |

ทุก POST ต้องมี CSRF token (ในฟอร์มหรือ header `X-CSRF-Token`) และเส้นทางที่กำหนดบทบาทไว้จะ redirect 302 ไปหน้าล็อกอินที่เหมาะสมหากไม่มีสิทธิ์

## การดูแลและบำรุงรักษา

**สำรอง / กู้คืน** — ทำได้ผ่านหน้าตั้งค่าในเดสก์ท็อป หรือสั่งผ่าน CLI:

```bash
# สำรอง (SQLite online backup ปลอดภัยแม้เซิร์ฟเวอร์กำลังทำงานอยู่) + ตรวจสอบ
venv\Scripts\python -c "from services.backup_service import create_backup, verify_backup; p = create_backup(); print(verify_backup(p), p)"

# กู้คืน (ปฏิเสธไฟล์ที่ตรวจไม่ผ่าน; แทนที่ไฟล์ปลายทางแบบ atomic)
venv\Scripts\python -c "from services.backup_service import restore_backup; restore_backup(r'backups\library_backup_YYYYMMDD_HHMMSS.db')"
```

**หมุนเวียนคีย์ QR** — เมื่อสงสัยว่าคีย์รั่ว หรือกำลังจะย้ายเครื่อง ให้รันเครื่องมือที่มาพร้อมโปรเจกต์นี้ ซึ่งจะหมุนคีย์โดยที่บัตรที่ออกไปแล้วไม่หาย:

```bash
venv\Scripts\python tools\rotate_qr_key.py [--no-cards]   # backup → ถอดรหัส → เข้ารหัสใหม่ทั้งหมดแบบ v2
venv\Scripts\python tools\regenerate_member_cards.py      # สร้างบัตร PDF ซ้ำจากฐานข้อมูลปัจจุบัน
```

หลังหมุนเวียนคีย์: แจกจ่าย `.env` ใหม่ให้ทุกเครื่อง **ก่อน**บูตครั้งถัดไป และย้ายพร้อมกันทั้งชุดเสมอ (`.env` + `db/` + `assets/`)

**เปิดบน LAN / HTTPS** — ใช้ reverse proxy มายืนหน้าและ terminate TLS จากนั้นตั้ง `webui_host = 0.0.0.0` และเปิด `SESSION_COOKIE_SECURE` — รายละเอียดอยู่ที่ `docs/SECURITY.md` § 9

**ปัญหาที่พบบ่อยและวิธีตรวจสอบ:**

| อาการ | วิธีตรวจสอบ |
|---|---|
| กล้องเว็บใช้ไม่ได้ | ต้องใช้ผ่าน `localhost` หรือ HTTPS; ตรวจสอบว่าอนุญาตกล้องในเบราว์เซอร์แล้ว |
| ล็อกอินสมาชิกได้รับ 429 | ล้มเหลว 5 ครั้ง/15 นาทีต่อชุดตัวตน รอให้ครบกำหนด หรือให้ผู้ดูแลล้าง PIN |
| บัตรเก่าถอดรหัสไม่ได้ทั้งระบบ | `.env` ไม่ตรงกันระหว่างเครื่อง — จัดชุด (`.env` + `db/` + `assets/`) ให้ตรงกัน หรือรัน `tools/rotate_qr_key.py` |
| หน้าสมาชิกทั้งระบบคืน 404 | มีใครปิด "เปิดหน้าสมาชิก" ในหน้าตั้งค่า |
| ตารางเว็บช้าหรือแสดงว่าง | ตรวจว่า migration รันครบ (`PRAGMA user_version` = 7) และ `/api/*` ส่งข้อมูลกลับมา |
| ลืมรหัสผ่าน admin | เปิดไฟล์ `db/INITIAL_ADMIN_PASSWORD.txt`; กรณีอื่นดูที่ `docs/SECURITY.md` |

## รูปแบบความปลอดภัย

| ด้าน | แนวทาง |
|---|---|
| ตัวตน QR | **v2** — บัตรคงที่ `AES-256-GCM` (nonce 12B + ป้องกันการแก้ไขด้วย tag) · **v3** — ตั๋วหมุนเวียน `AES-256-GCM` + AAD ผูกโปรโตคอล (`pylib-member-ticket-v3`) เนื้อในเป็น JSON ที่ตรวจสอบทั้งฟิลด์และชนิดข้อมูลอย่างเข้มงวด (`{"m": id, "e": epoch}`) · **v1 legacy** — AES-CBC อ่านได้เท่านั้น ถอดรหัสแบบเข้มงวด (PKCS7 + UTF-8 เต็มรูปแบบ) — **ไม่มี plaintext fallback** รหัสที่ผิดรูปแบบจะถูกปฏิเสธทั้งหมด |
| รอบเวลาของตั๋ว | `epoch = floor(unix_time / รอบ)` — เซิร์ฟเวอร์พิจารณาเฉพาะรอบปัจจุบันและรอบข้างเคียงที่เผื่อไว้ และเป็น **stateless** (ไม่ต้องเก็บสถานะฝั่งใดเลย) |
| การสร้างคีย์ | PBKDF2-HMAC-SHA256 100,000 รอบ จาก `SECRET_PASSPHRASE` + `SECRET_SALT` (สร้างอัตโนมัติ พร้อมแทนที่ค่า placeholder ให้เป็นค่าสุ่มเสมอ) |
| รหัสผ่านผู้ดูแล | bcrypt, ไม่มีค่าเริ่มต้น — บูตแรกสุ่มและแสดงครั้งเดียว; ตรวจสอบแบบ timing-equalized; `admin123` ถูกแทนที่อัตโนมัติทุกบูต |
| PIN สมาชิก | bcrypt ใน `members.pin_hash` (NULL จนกว่าสมาชิกจะเปิดใช้งานเอง); การเปิดใช้งานเป็นทรานแซกชันเดียว (อ้างสิทธิ์โค้ด + ตั้ง PIN + เงื่อนไข `pin_hash IS NULL` ป้องกันการแทนที่) |
| รหัสเปิดใช้งาน | เก็บเฉพาะแฮช SHA-256, อายุ 15 นาที, ใช้ได้ครั้งเดียว (claim แบบ atomic), แสดงบนหน้าจอเพียงครั้งเดียว — ตัวรหัสจริงไม่ถูกบันทึกไว้ที่ไหนเลย |
| จำกัดการล็อกอิน | 5 ครั้ง / 15 นาที แยกขอบเขตอิสระ: ผู้ดูแล (IP + username) · สมาชิก (IP + ชั้น + เลขที่) |
| Session เว็บ | คุกกี้ HttpOnly + SameSite=Lax อายุ 4 ชั่วโมง, CSRF token ทุก POST (เปรียบเทียบแบบ constant-time), ข้อความแสดงข้อผิดพลาดออกแบบให้ไม่เปิดเผยว่ามีหรือไม่มีสมาชิกนั้น |
| ขอบเขตสิทธิ์ | หน้าสมาชิกอ่านได้เฉพาะข้อมูลของตัวเอง ไม่มีเส้นทางแก้ไขข้อมูลใดๆ; ปิดทั้งโดเมนได้ด้วย kill switch |
| ความถูกต้องของข้อมูล | FK เปิดทุก connection, `books.code UNIQUE`, `borrow_log.returned CHECK (0,1)`, ยืม/คืนเป็นทรานแซกชัน `BEGIN IMMEDIATE` เดียว, parameterized SQL ทุกจุด |
| Migrations | `PRAGMA user_version` v1→v7 (constraints → ดัชนี bulk-import → หน้าสมาชิก → placements/ย้ายปีการศึกษา → ตรวจนับหนังสือ → การจองหนังสือ) — รันทุกบูตและ idempotent |
| การบันทึก log | ไม่เขียนข้อมูลส่วนบุคคลลง log (เฉพาะ id และจำนวน); payload ที่ล้มเหลวจะคืนเฉพาะผลลัพธ์โดยไม่สะท้อนเนื้อหากลับมา |

เอกสารเชิงปฏิบัติฉบับเต็ม (รูปแบบ payload ทุกเวอร์ชัน, runbook หมุนเวียนคีย์, deployment HTTPS, เมทริกซ์ทดสอบความปลอดภัย): [`docs/SECURITY.md`](docs/SECURITY.md)

## ฐานข้อมูล

SQLite3 ไฟล์เดียว (`db/library.db`) — โครงสร้างถูกจัดการด้วย migration แยกตามเวอร์ชัน:

| ตาราง | เนื้อหา |
|---|---|
| `admin_users` | ผู้ดูแลระบบ (bcrypt, ไม่มีค่าเริ่มต้น) |
| `members` | ชื่อ, ชั้น, เลขที่, QR, วันหมดอายุ, **`pin_hash`** (v4), **`status`** (v5: active/graduated/left); badge ชั้น+เลขที่ คือข้อมูลปัจจุบัน |
| `books` | รหัส (**UNIQUE**), ชื่อ, สถานะ `ว่าง` / `ยืมแล้ว` / `หยุดให้บริการ` |
| `borrow_log` | ทรานแซกชันยืม-คืน — `returned CHECK (0,1)`, FK ไป members/books (ประวัติไม่ถูกเขียนทับโดยการย้ายปีการศึกษา) |
| `access_log` | บันทึกเข้า-ออกห้องสมุด |
| `member_activation_codes` | (v4) แฮช SHA-256 ของรหัสเปิดใช้งาน, อายุ 15 นาที, ใช้ได้ครั้งเดียว |
| `academic_years`, `member_placements` | (v5) ปีการศึกษาและ placement ปัจจุบัน/ประวัติของสมาชิก (หนึ่ง current ต่อสมาชิก) |
| `academic_year_migrations`, `academic_year_migration_items` | (v5) บันทึกการย้ายปีการศึกษาและผลรายบุคคล (audit trail) |
| `inventory_audits`, `inventory_audit_scans`, `inventory_audit_resolutions` | (v6) รอบตรวจนับ, การสแกน/ผลกระทบยอด, และการจัดการความคลาดเคลื่อน |
| `book_reservations` | (v7) คิวจองรายเล่ม (waiting/offered/fulfilled/cancelled/expired) พร้อมเวลาเสนอ/หมดเวลา |

## การทดสอบและ CI

```bash
venv\Scripts\python -m unittest discover -s tests        # Windows
venv/bin/python -m unittest discover -s tests            # Linux/Mac
```

**214 เทสเคส** — ครอบคลุม crypto (v1/v2/v3), ระบบระบุตัวตนสมาชิก, หน้าสมาชิก, ช่อง WebSocket เรียลไทม์, repository และการยืม-คืนแบบ atomic, migration v1→v7, สำรอง/กู้คืน, bulk + import Excel, API ตาราง server-side, guard ของเว็บ (CSRF/throttle/session/สิทธิ์), ย้ายปีการศึกษา, ตรวจนับหนังสือ, การจองหนังสือ และ offline-safety

CI (GitHub Actions) รัน 3 job แยกกันทุก push/PR:

| Job | สิ่งที่ทำ |
|---|---|
| `tests + lint` | ชุดทดสอบทั้งหมด (unittest) + `ruff check` (กำหนดเองใน `pyproject.toml` รวม rule bandit) |
| `pip-audit` | สแกน CVE ของ dependency ทุกตัวที่ติดตั้ง |
| `gitleaks` | ค้นหา secret ที่ commit เข้ามาใน codebase |

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
│   ├── live.py                  # WebSocket hub + DB-marker ticker (realtime push)
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
│   │       ├── member_live.js   # หน้าสมาชิก: อัปเดตจำนวน/ตารางจาก /member/ws (เรียลไทม์)
│   │       ├── admin_live.js    # หน้าผู้ดูแล: สั่งโหลดตาราง/แดชบอร์ดใหม่เมื่อข้อมูลเปลี่ยน
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
| `opencv-python` + `numpy` | เว็บแคม + ประมวลผลภาพ (เดสก์ท็อป) |
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

เวอร์ชันทั้งหมด pin ไว้ใน `requirements.txt` (ไฟล์ระบุไว้ว่าตรวจสอบบน Windows / Python 3.10)

## แผนการพัฒนา

ตอนนี้ฟีเจอร์หลักทุกอย่างใช้งานได้จริงแล้ว — สิ่งที่อยู่ในแผนต่อไป:

- [ ] สร้างบัตร/รายงาน PDF จากเว็บ (ตอนนี้ทำได้บนเดสก์ท็อปเท่านั้น)
- [ ] ให้สมาชิกเปลี่ยน PIN เองในหน้าของตัวเอง (ตอนนี้รีเซ็ตโดยผู้ดูแลเท่านั้น)
- [ ] สำรองข้อมูลอัตโนมัติตามเวลา + นโยบายเก็บสำรองย้อนหลัง
- [x] ตู้บริการตนเอง (self-service kiosk) — ยืม/คืนด้วยตนเองแบบแยก session, รีเซ็ตอัตโนมัติ, ใช้บริการ atomic เดียวกับระบบหลัก
- [x] ย้ายปีการศึกษา / เลื่อนชั้น (Academic Year Migration) — preview/ยืนยัน/สำรองก่อนบันทึก, เก็บประวัติการย้าย
- [x] ตรวจนับหนังสือ (Inventory Audit) — กระทบยอด, แยกกำลังยืมจากขาดหาย, รายงาน/จัดการความคลาดเคลื่อน
- [x] การจองหนังสือ (Book Reservation) — คิว FIFO, ข้อเสนอพร้อมหมดเวลา, รับ/ยืมจริงอย่างชัดเจน (ไม่มีการยืมอัตโนมัติ)

พบบั๊กหรือมีข้อเสนอแนะอะไร เปิด [issue](https://github.com/Nonx2360/The-PY-LIB-Project/issues) ได้เลย

## การมีส่วนร่วม

1. Fork แล้วตั้ง branch (`feature/…` หรือ `fix/…`)
2. `ruff check .` ต้องผ่านทั้งหมด — กำหนดเองใน `pyproject.toml`
3. ทุกการเปลี่ยนแปลงต้องไม่ทำให้ชุดทดสอบพัง: `python -m unittest discover -s tests`
4. การแตะด้านความปลอดภัย (การเข้ารหัส / สิทธิ์ / session) ควรเพิ่ม regression test ให้ตรงกับเมทริกซ์ใน [`docs/SECURITY.md`](docs/SECURITY.md) § 11
5. เปิด PR พร้อมอธิบายเหตุผลและแนบภาพหน้าจอ (ถ้าแตะ UI)

ยินดีรับข้อเสนอที่ไม่ใช่โค้ดด้วยนะ — เปิด issue แล้วเล่าสถานการณ์การใช้งานจริงในห้องสมุดของคุณให้ฟัง

## ทีมและใบอนุญาต

- เด็กหญิง ขวัญชนก อุ่นศิริ เลขที่ 4 (UX/UI)
- เด็กชาย ณัฐชนน รอดน้อย เลขที่ 30 (Developer)
- **โรงเรียน:** DSNPRU

ขอบคุณผู้ทดสอบ beta และอาจารย์ที่ปรึกษาทุกท่าน

MIT License © 2025
