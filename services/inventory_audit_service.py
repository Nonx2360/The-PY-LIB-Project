# services/inventory_audit_service.py
import csv
import io

from core import security
from db.repositories.inventory_audit_repo import InventoryAuditRepository


class InventoryAuditService:
    def __init__(self, repo=None):
        self.repo = repo or InventoryAuditRepository()

    def create_session(self, name, location, actor):
        active = self.repo.has_active_session()
        if active:
            raise ValueError(f"มีรอบตรวจนับที่ยังไม่เสร็จสิ้นอยู่ (รหัส {active}) — กรุณาพัก/ปิดรอบเดิมก่อน")
        audit_id = self.repo.create(name, location, actor)
        return audit_id

    def set_status(self, audit_id, status, actor):
        audit = self.repo.get(audit_id)
        if not audit:
            raise ValueError("ไม่พบรอบตรวจนับ")
        cur = audit[4]
        if status == cur:
            return
        # allow completing/cancelling from in_progress or paused
        if status == "completed" and cur in ("in_progress", "paused"):
            pass
        elif status == "cancelled" and cur in ("in_progress", "paused"):
            pass
        elif status == "in_progress" and cur == "paused":
            pass
        elif status == "paused" and cur == "in_progress":
            pass
        else:
            raise ValueError(f"ไม่สามารถเปลี่ยนสถานะจาก {cur} เป็น {status} ได้")
        self.repo.set_status(audit_id, status, actor)

    def _decode_payload(self, raw_value):
        """Best-effort: return candidate book code strings from scan payload."""
        raw = (raw_value or "").strip()
        if not raw:
            return []
        candidates = [raw]
        # try decrypting as QR payload (member/book cards may be encrypted)
        for val in (raw,):
            try:
                dec = security.decrypt_qr_data(val)
                if dec and dec not in candidates:
                    candidates.append(dec.strip())
            except Exception:
                pass
        # encrypted payloads are often just the code; also try splitting if
        # someone scans a member-like badge by mistake we won't match a book
        return [c for c in candidates if c]

    def scan(self, audit_id, raw_value, actor, manual=False):
        audit = self.repo.get(audit_id)
        if not audit:
            raise ValueError("ไม่พบรอบตรวจนับ")
        if audit[4] != "in_progress":
            raise ValueError("รอบตรวจนับนี้ไม่ได้อยู่ในสถานะกำลังดำเนินการ")
        candidates = self._decode_payload(raw_value)
        book = None
        matched_value = (raw_value or "").strip()
        for cand in candidates:
            book = self.repo.find_book_by_code(cand)
            if book:
                matched_value = cand
                break
        if not book:
            self.repo.record_scan(audit_id, None, raw_value, "unknown", actor)
            return {"ok": True, "outcome": "unknown",
                    "message": "ไม่พบหนังสือจากรหัส/QR นี้ในฐานข้อมูล (บันทึกเป็นรายการไม่รู้จัก)",
                    "raw": raw_value}

        bid, code, title, status = book[0], book[1], book[2], book[3]
        if status == "หยุดให้บริการ":
            self.repo.record_scan(audit_id, bid, matched_value, "retired", actor)
            return {"ok": True, "outcome": "retired", "book_id": bid,
                    "code": code, "title": title, "status": status,
                    "message": f"{code} — {title}: หนังสือถูกหยุดให้บริการ (ไม่นับในรายการที่คาดบนชั้น)"}

        if self.repo.already_found(audit_id, bid):
            self.repo.record_scan(audit_id, bid, matched_value, "duplicate", actor)
            loan = self.repo.book_loan_state(bid)
            return {"ok": True, "outcome": "duplicate", "book_id": bid,
                    "code": code, "title": title, "status": status,
                    "on_loan": bool(loan),
                    "message": f"{code} — {title}: สแกนซ้ำในรอบนี้แล้ว (ไม่นับเพิ่ม)"}

        # try to count as found once (unique constraint guards races)
        counted = self.repo.mark_found_once(audit_id, bid, matched_value, actor)
        loan = self.repo.book_loan_state(bid)
        if not counted:
            # race: someone else recorded it; treat as duplicate event
            self.repo.record_scan(audit_id, bid, matched_value, "duplicate", actor)
            return {"ok": True, "outcome": "duplicate", "book_id": bid,
                    "code": code, "title": title, "status": status,
                    "on_loan": bool(loan),
                    "message": f"{code} — {title}: สแกนซ้ำในรอบนี้แล้ว"}
        msg = f"พบ: {code} — {title}"
        if loan:
            msg += " (ระบบระบุว่ากำลังถูกยืม — บันทึกว่าพบแล้ว แต่ควรตรวจสอบสถานะการยืม)"
        return {"ok": True, "outcome": "found", "book_id": bid,
                "code": code, "title": title, "status": status,
                "on_loan": bool(loan), "message": msg}

    def report(self, audit_id):
        prog = self.repo.progress(audit_id)
        if not prog:
            raise ValueError("ไม่พบรอบตรวจนับ")
        prog["shelf_missing_books"] = self.repo.books_by_ids(prog["shelf_missing_ids"][:200])
        prog["on_loan_books"] = self.repo.books_by_ids(prog["on_loan_not_scanned_ids"][:200])
        prog["recent_scans"] = self.repo.recent_scans(audit_id, limit=50)
        prog["resolutions"] = self.repo.resolutions(audit_id)
        return prog

    def resolve(self, audit_id, book_id, scan_id, resolution, note, actor):
        audit = self.repo.get(audit_id)
        if not audit:
            raise ValueError("ไม่พบรอบตรวจนับ")
        allowed = {
            "confirm_found", "not_found_recheck", "flag_investigation",
            "location_corrected", "note",
        }
        if resolution not in allowed:
            raise ValueError("ประเภทการจัดการไม่ถูกต้อง")
        self.repo.add_resolution(audit_id, book_id or None, scan_id or None,
                                 resolution, note, actor)
        return True

    def export_csv(self, audit_id):
        prog = self.report(audit_id)
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(["inventory_audit_report"])
        w.writerow(["audit_id", prog["audit"]["id"]])
        w.writerow(["name", prog["audit"]["name"]])
        w.writerow(["status", prog["audit"]["status"]])
        w.writerow([])
        w.writerow(["metric", "value"])
        for k, v in prog["counts"].items():
            w.writerow([k, v])
        w.writerow([])
        w.writerow(["shelf_missing_books"])
        w.writerow(["id", "code", "title", "status"])
        for b in prog["shelf_missing_books"]:
            w.writerow([b[0], b[1], b[2], b[3]])
        w.writerow([])
        w.writerow(["on_loan_not_scanned_books"])
        w.writerow(["id", "code", "title", "status"])
        for b in prog["on_loan_books"]:
            w.writerow([b[0], b[1], b[2], b[3]])
        w.writerow([])
        w.writerow(["resolutions"])
        w.writerow(["id", "book_id", "code", "title", "resolution", "note", "actor", "resolved_at"])
        for r in prog["resolutions"]:
            w.writerow(list(r))
        return "\ufeff" + buf.getvalue()
