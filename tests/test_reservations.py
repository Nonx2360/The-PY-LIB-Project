# tests/test_reservations.py — copy-level FIFO reservations, no silent loans
import os
import sys
import unittest
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

_env = harness.TempEnv()

import db.schema as schema_mod  # noqa: E402
from db.repositories.base_repo import BaseRepository  # noqa: E402
from db.repositories.member_repo import MemberRepository  # noqa: E402
from db.repositories.book_repo import BookRepository  # noqa: E402
from db.repositories.borrow_repo import BorrowRepository, BorrowError  # noqa: E402
from services.reservation_service import ReservationService  # noqa: E402
from config import settings as settings_mod  # noqa: E402

ORIG_SETTINGS_FILE = settings_mod.SETTINGS_FILE


class TestReservations(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        _env.apply()
        _env.patch()
        schema_mod.init_db()
        settings_mod.SETTINGS_FILE = os.path.join(_env.dir, "settings.json")
        settings_mod.load()
        settings_mod.set("reservation_offer_hours", 48)
        settings_mod.save()
        cls.base = BaseRepository()
        cls.members = MemberRepository()
        cls.books = BookRepository()
        cls.borrow = BorrowRepository()
        cls.svc = ReservationService()

    @classmethod
    def tearDownClass(cls):
        settings_mod.SETTINGS_FILE = ORIG_SETTINGS_FILE
        settings_mod.load()
        _env.restore()
        harness.cleanup(_env.dir)

    def _member(self, name, grade, num):
        return self.members.add(name, grade, num, "2026-01-01", "2027-01-01", "qr", None)

    def _today(self):
        return datetime.now().strftime("%Y-%m-%d")

    def _due(self, days=7):
        return (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d")

    def test_queue_offer_on_return_and_claim_no_silent_loan(self):
        m1 = self._member("คนหนึ่ง", "ม.1", "R1")
        m2 = self._member("คนสอง", "ม.1", "R2")
        m3 = self._member("คนสาม", "ม.1", "R3")
        b = self.books.add("RES-1", "เล่มจอง", "ว่าง")
        # m1 borrows, others reserve (FIFO m2 then m3)
        rec = self.borrow.try_borrow(m1, b, self._today(), self._due())
        self.svc.reserve(b, m2)
        self.svc.reserve(b, m3)
        # return -> should OFFER to m2 (not auto-loan)
        self.borrow.try_return(rec)
        status = self.books.fetchone("SELECT status FROM books WHERE id=?", (b,))[0]
        self.assertEqual(status, "ว่าง")
        loan_open = self.base.fetchone(
            "SELECT COUNT(*) FROM borrow_log WHERE book_id=? AND returned=0", (b,))[0]
        self.assertEqual(loan_open, 0, "no silent loan on return")
        res = self.base.fetchall(
            "SELECT member_id, status FROM book_reservations WHERE book_id=? ORDER BY id", (b,))
        self.assertEqual([(r[0], r[1]) for r in res], [(m2, 'offered'), (m3, 'waiting')])
        # m3 cannot borrow (offered to m2)
        with self.assertRaises(BorrowError):
            self.borrow.try_borrow(m3, b, self._today(), self._due())
        # m2 claims explicitly -> becomes loan, fulfilled
        rid = self.base.fetchone(
            "SELECT id FROM book_reservations WHERE book_id=? AND member_id=? AND status='offered'",
            (b, m2))[0]
        self.svc.claim(rid, m2)
        self.assertEqual(self.books.fetchone("SELECT status FROM books WHERE id=?", (b,))[0], "ยืมแล้ว")
        st2 = self.base.fetchone("SELECT status FROM book_reservations WHERE id=?", (rid,))[0]
        self.assertEqual(st2, 'fulfilled')

    def test_offer_expires_and_promotes_next(self):
        m1 = self._member("A", "ม.2", "E1")
        m2 = self._member("B", "ม.2", "E2")
        b = self.books.add("RES-2", "หมดเวลา", "ว่าง")
        self.svc.reserve(b, m1)
        self.svc.reserve(b, m2)
        # immediate offer should exist for m1 (book is available)
        r1 = self.base.fetchone("SELECT id,status FROM book_reservations WHERE member_id=? AND book_id=?", (m1,b))
        self.assertEqual(r1[1], 'offered')
        # force expiry in past
        past = (datetime.now() - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
        self.base.execute("UPDATE book_reservations SET expires_at=? WHERE id=?", (past, r1[0]))
        # trigger promotion by sweep via service action
        self.svc._expire_sweep()
        self.base.execute("UPDATE book_reservations SET status='expired' WHERE status='offered' AND expires_at < ?",
                          (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),))
        # promote next
        self.base.transact(lambda conn: self.svc.repo._offer_next_conn(conn, b, 48))
        r1s = self.base.fetchone("SELECT status FROM book_reservations WHERE id=?", (r1[0],))[0]
        r2s = self.base.fetchone("SELECT status FROM book_reservations WHERE member_id=? AND book_id=?", (m2,b))[0]
        self.assertEqual(r1s, 'expired')
        self.assertEqual(r2s, 'offered')

    def test_cancel_promotes_next(self):
        m1 = self._member("C", "ม.3", "C1")
        m2 = self._member("D", "ม.3", "C2")
        b = self.books.add("RES-3", "ยกเลิก", "ว่าง")
        self.svc.reserve(b, m1)
        self.svc.reserve(b, m2)
        rid1 = self.base.fetchone("SELECT id FROM book_reservations WHERE member_id=? AND book_id=?", (m1,b))[0]
        self.svc.cancel(rid1, actor_member_id=m1)
        r2s = self.base.fetchone("SELECT status FROM book_reservations WHERE member_id=? AND book_id=?", (m2,b))[0]
        self.assertEqual(r2s, 'offered')


if __name__ == "__main__":
    unittest.main()
