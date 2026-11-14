# db/repositories/inventory_audit_repo.py
from datetime import datetime

from .base_repo import BaseRepository


class InventoryAuditRepository(BaseRepository):
    def now(self):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def create(self, name, location, actor, scope="all"):
        now = self.now()
        # expected = shelf-eligible books in scope (all non-retired by default
        # when scope=='all'; callers can refine later). Snapshot count for
        # a stable denominator at creation time.
        expected = self.fetchone(
            "SELECT COUNT(*) FROM books WHERE status != 'หยุดให้บริการ'")[0]
        return self.execute(
            "INSERT INTO inventory_audits "
            "(name, location, scope, status, created_by, notes, expected_count, created_at, updated_at) "
            "VALUES (?,?,?,?,?,?,?, ?, ?)",
            ((name or "การตรวจนับ").strip() or "การตรวจนับ",
             (location or "").strip() or None,
             scope or "all", "in_progress", actor, None,
             int(expected), now, now))

    def get(self, audit_id):
        return self.fetchone(
            "SELECT id, name, location, scope, status, created_by, notes, "
            "expected_count, created_at, updated_at, completed_at "
            "FROM inventory_audits WHERE id = ?", (audit_id,))

    def list_recent(self, limit=30):
        return self.fetchall(
            "SELECT id, name, location, scope, status, created_by, expected_count, "
            "created_at, updated_at, completed_at FROM inventory_audits "
            "ORDER BY id DESC LIMIT ?", (limit,))

    def set_status(self, audit_id, status, actor=None):
        now = self.now()
        if status == "completed":
            self.execute(
                "UPDATE inventory_audits SET status=?, updated_at=?, completed_at=? WHERE id=?",
                (status, now, now, audit_id))
        else:
            self.execute(
                "UPDATE inventory_audits SET status=?, updated_at=? WHERE id=?",
                (status, now, audit_id))

    def has_active_session(self, exclude_id=None):
        if exclude_id:
            row = self.fetchone(
                "SELECT id FROM inventory_audits "
                "WHERE status IN ('in_progress','paused') AND id != ? LIMIT 1",
                (exclude_id,))
        else:
            row = self.fetchone(
                "SELECT id FROM inventory_audits "
                "WHERE status IN ('in_progress','paused') LIMIT 1")
        return row[0] if row else None

    # ---------------- scans & reconciliation ----------------
    def find_book_by_code(self, code):
        code = (code or "").strip()
        if not code:
            return None
        # exact, then case-insensitive fallback
        row = self.fetchone(
            "SELECT id, code, title, status FROM books WHERE code = ?", (code,))
        if row:
            return row
        return self.fetchone(
            "SELECT id, code, title, status FROM books WHERE lower(code) = lower(?)",
            (code,))

    def book_loan_state(self, book_id):
        row = self.fetchone(
            "SELECT id, member_id, borrow_date, return_due FROM borrow_log "
            "WHERE book_id = ? AND returned = 0 ORDER BY id DESC LIMIT 1",
            (book_id,))
        if not row:
            return None
        return {"record_id": row[0], "member_id": row[1],
                "borrow_date": row[2], "return_due": row[3]}

    def already_found(self, audit_id, book_id):
        return self.fetchone(
            "SELECT id FROM inventory_audit_scans "
            "WHERE audit_id = ? AND book_id = ? AND outcome = 'found' LIMIT 1",
            (audit_id, book_id))

    def record_scan(self, audit_id, book_id, raw_value, outcome, actor):
        return self.execute(
            "INSERT INTO inventory_audit_scans "
            "(audit_id, book_id, raw_value, outcome, scanned_by, scanned_at) "
            "VALUES (?,?,?,?,?,?)",
            (audit_id, book_id, (raw_value or "")[:500], outcome, actor, self.now()))

    def mark_found_once(self, audit_id, book_id, raw_value, actor):
        """Insert a unique 'found' scan; return True if newly counted."""

        def _ins(conn):
            try:
                conn.execute(
                    "INSERT INTO inventory_audit_scans "
                    "(audit_id, book_id, raw_value, outcome, scanned_by, scanned_at) "
                    "VALUES (?,?,?,?,?,?)",
                    (audit_id, book_id, (raw_value or "")[:500], 'found', actor, self.now()))
                return True
            except Exception:
                return False
        return self.transact(_ins)

    def progress(self, audit_id):
        audit = self.get(audit_id)
        if not audit:
            return None
        # unique found (real books)
        found = self.fetchone(
            "SELECT COUNT(DISTINCT book_id) FROM inventory_audit_scans "
            "WHERE audit_id = ? AND outcome = 'found' AND book_id IS NOT NULL",
            (audit_id,))[0]
        unknown = self.fetchone(
            "SELECT COUNT(*) FROM inventory_audit_scans "
            "WHERE audit_id = ? AND outcome = 'unknown'", (audit_id,))[0]
        duplicates = self.fetchone(
            "SELECT COUNT(*) FROM inventory_audit_scans "
            "WHERE audit_id = ? AND outcome = 'duplicate'", (audit_id,))[0]
        retired = self.fetchone(
            "SELECT COUNT(*) FROM inventory_audit_scans "
            "WHERE audit_id = ? AND outcome = 'retired'", (audit_id,))[0]
        # on-loan books (currently borrowed) - separate from shelf-missing
        on_loan_total = self.fetchone(
            "SELECT COUNT(DISTINCT b.id) FROM books b "
            "JOIN borrow_log bl ON bl.book_id = b.id AND bl.returned = 0 "
            "WHERE b.status = 'ยืมแล้ว'")[0]
        # expected shelf set for this scope (all non-retired, not on-loan treated separately in report)
        expected_ids = [r[0] for r in self.fetchall(
            "SELECT id FROM books WHERE status != 'หยุดให้บริการ'")]
        found_ids = {r[0] for r in self.fetchall(
            "SELECT DISTINCT book_id FROM inventory_audit_scans "
            "WHERE audit_id = ? AND outcome = 'found' AND book_id IS NOT NULL",
            (audit_id,))}
        not_scanned_ids = [bid for bid in expected_ids if bid not in found_ids]
        # split not-scanned into on-loan vs shelf-missing
        on_loan_ids = {r[0] for r in self.fetchall(
            "SELECT DISTINCT b.id FROM books b "
            "JOIN borrow_log bl ON bl.book_id = b.id AND bl.returned = 0")}
        shelf_missing = [bid for bid in not_scanned_ids if bid not in on_loan_ids]
        on_loan_not_scanned = [bid for bid in not_scanned_ids if bid in on_loan_ids]
        expected = len(expected_ids)
        pct = round((found / expected) * 100, 1) if expected else 100.0
        return {
            "audit": {
                "id": audit[0], "name": audit[1], "location": audit[2],
                "scope": audit[3], "status": audit[4], "created_by": audit[5],
                "notes": audit[6], "expected_count": audit[7],
                "created_at": audit[8], "updated_at": audit[9],
                "completed_at": audit[10],
            },
            "counts": {
                "expected": expected,
                "found_unique": found,
                "not_scanned": len(not_scanned_ids),
                "shelf_missing": len(shelf_missing),
                "on_loan_total": on_loan_total,
                "on_loan_not_scanned": len(on_loan_not_scanned),
                "unknown_scans": unknown,
                "duplicate_scans": duplicates,
                "retired_encountered": retired,
                "progress_pct": pct,
            },
            "shelf_missing_ids": shelf_missing,
            "on_loan_not_scanned_ids": on_loan_not_scanned,
        }

    def books_by_ids(self, ids):
        if not ids:
            return []
        marks = ",".join("?" * len(ids))
        return self.fetchall(
            f"SELECT id, code, title, status FROM books WHERE id IN ({marks}) ORDER BY code",  # noqa: S608 - {marks} is purely '?' placeholders; values are bound
            tuple(ids))

    def recent_scans(self, audit_id, limit=50):
        return self.fetchall(
            "SELECT s.id, s.book_id, b.code, b.title, s.raw_value, s.outcome, "
            "s.scanned_by, s.scanned_at "
            "FROM inventory_audit_scans s "
            "LEFT JOIN books b ON b.id = s.book_id "
            "WHERE s.audit_id = ? ORDER BY s.id DESC LIMIT ?",
            (audit_id, limit))

    def add_resolution(self, audit_id, book_id, scan_id, resolution, note, actor):
        self.execute(
            "INSERT INTO inventory_audit_resolutions "
            "(audit_id, book_id, scan_id, resolution, note, actor, resolved_at) "
            "VALUES (?,?,?,?,?,?,?)",
            (audit_id, book_id, scan_id, resolution, (note or "")[:500], actor, self.now()))

    def resolutions(self, audit_id):
        return self.fetchall(
            "SELECT r.id, r.book_id, b.code, b.title, r.scan_id, r.resolution, r.note, "
            "r.actor, r.resolved_at FROM inventory_audit_resolutions r "
            "LEFT JOIN books b ON b.id = r.book_id "
            "WHERE r.audit_id = ? ORDER BY r.id DESC",
            (audit_id,))
