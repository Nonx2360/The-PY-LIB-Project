# tests/test_inventory_audit.py — inventory audit reconciliation (no auto mutations)
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

_env = harness.TempEnv()

import db.schema as schema_mod  # noqa: E402
from db.repositories.base_repo import BaseRepository  # noqa: E402
from db.repositories.member_repo import MemberRepository  # noqa: E402
from db.repositories.book_repo import BookRepository  # noqa: E402
from db.repositories.borrow_repo import BorrowRepository  # noqa: E402
from services.inventory_audit_service import InventoryAuditService  # noqa: E402
from config import settings as settings_mod  # noqa: E402

ORIG_SETTINGS_FILE = settings_mod.SETTINGS_FILE


class TestInventoryAudit(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        _env.apply()
        _env.patch()
        schema_mod.init_db()
        settings_mod.SETTINGS_FILE = os.path.join(_env.dir, "settings.json")
        settings_mod.load()
        cls.base = BaseRepository()
        cls.members = MemberRepository()
        cls.books = BookRepository()
        cls.borrow = BorrowRepository()
        cls.svc = InventoryAuditService()

    def setUp(self):
        # isolate audit sessions between tests
        self.base.execute("DELETE FROM inventory_audit_resolutions")
        self.base.execute("DELETE FROM inventory_audit_scans")
        self.base.execute("DELETE FROM inventory_audits")

    @classmethod
    def tearDownClass(cls):
        settings_mod.SETTINGS_FILE = ORIG_SETTINGS_FILE
        settings_mod.load()
        _env.restore()
        harness.cleanup(_env.dir)

    def test_scan_found_duplicate_unknown_and_no_mutation(self):
        bid = self.books.add("INV-1", "เล่มหนึ่ง", "ว่าง")
        before = self.base.fetchone("SELECT status FROM books WHERE id = ?", (bid,))[0]
        aid = self.svc.create_session("รอบทดสอบ", None, "tester")
        r1 = self.svc.scan(aid, "INV-1", "tester")
        self.assertEqual(r1["outcome"], "found")
        r2 = self.svc.scan(aid, "INV-1", "tester")
        self.assertEqual(r2["outcome"], "duplicate")
        r3 = self.svc.scan(aid, "NO-SUCH", "tester")
        self.assertEqual(r3["outcome"], "unknown")
        after = self.base.fetchone("SELECT status FROM books WHERE id = ?", (bid,))[0]
        self.assertEqual(before, after, "audit must not mutate book status")

    def test_on_loan_separate_from_shelf_missing(self):
        # setup: one available scanned, one on-loan (not shelf-missing), one truly missing
        self.books.add("INV-A", "พบ", "ว่าง")
        b_loan = self.books.add("INV-B", "ยืม", "ว่าง")
        b_miss = self.books.add("INV-C", "หายจากชั้น", "ว่าง")
        m = self.members.add("ผู้ยืม", "ม.1", "1", "2026-01-01", "2027-01-01", "qr", None)
        # create an active loan record and mark borrowed
        self.borrow.try_borrow(m, b_loan, "2026-11-01", "2026-11-30")
        self.assertEqual(self.books.fetchone("SELECT status FROM books WHERE id=?", (b_loan,))[0], "ยืมแล้ว")
        aid = self.svc.create_session("รอบแยกยืม", None, "tester")
        self.svc.scan(aid, "INV-A", "tester")
        # do not scan loaned or missing
        rep = self.svc.report(aid)
        self.assertIn(b_loan, rep["on_loan_not_scanned_ids"])
        self.assertNotIn(b_loan, rep["shelf_missing_ids"])
        self.assertIn(b_miss, rep["shelf_missing_ids"])
        self.assertNotIn(b_miss, rep["on_loan_not_scanned_ids"])
        # ensure books table unchanged (still borrowed/available)
        self.assertEqual(self.books.fetchone("SELECT status FROM books WHERE id=?", (b_loan,))[0], "ยืมแล้ว")

    def test_status_transitions_and_resolution_recorded(self):
        aid = self.svc.create_session("รอบสถานะ", None, "tester")
        self.svc.set_status(aid, "paused", "tester")
        self.assertEqual(self.svc.repo.get(aid)[4], "paused")
        self.svc.set_status(aid, "in_progress", "tester")
        self.assertEqual(self.svc.repo.get(aid)[4], "in_progress")
        bid = self.books.add("INV-R", "จัดการ", "ว่าง")
        self.svc.resolve(aid, bid, None, "flag_investigation", "ตรวจสอบ", "tester")
        res = self.svc.repo.resolutions(aid)
        self.assertTrue(any(r[1] == bid and r[5] == "flag_investigation" for r in res))
        self.svc.set_status(aid, "completed", "tester")
        self.assertEqual(self.svc.repo.get(aid)[4], "completed")
        with self.assertRaises(ValueError):
            self.svc.set_status(aid, "in_progress", "tester")

    def test_active_session_guard(self):
        aid = self.svc.create_session("รอบแรก", None, "tester")
        with self.assertRaises(ValueError):
            self.svc.create_session("รอบซ้ำ", None, "tester")
        self.svc.set_status(aid, "cancelled", "tester")
        aid2 = self.svc.create_session("รอบใหม่", None, "tester")
        self.assertGreater(aid2, 0)
        self.svc.set_status(aid2, "completed", "tester")


if __name__ == "__main__":
    unittest.main()
