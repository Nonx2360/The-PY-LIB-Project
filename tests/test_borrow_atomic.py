# tests/test_borrow_atomic.py — transaction guarantees for borrow/return
import os
import sys
import threading
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402


class TestBorrowAtomicity(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = harness.TempEnv()
        cls.env.apply()
        cls.env.patch()
        import db.schema as schema_mod
        schema_mod.init_db(log=lambda *a, **k: None)
        cls.schema = schema_mod

        from db.repositories.member_repo import MemberRepository
        from db.repositories.book_repo import BookRepository
        from db.repositories.borrow_repo import BorrowRepository
        cls.member_repo = MemberRepository()
        cls.book_repo = BookRepository()
        cls.borrow_repo = BorrowRepository()

    @classmethod
    def tearDownClass(cls):
        cls.env.restore()
        harness.cleanup(cls.env.dir)

    def _reset(self):
        conn = self.borrow_repo.get_connection()
        try:
            conn.execute("PRAGMA foreign_keys=OFF")
            conn.execute("DELETE FROM borrow_log")
            conn.execute("PRAGMA foreign_keys=ON")
            conn.execute("DELETE FROM members")
            conn.execute("DELETE FROM books")
        finally:
            conn.close()
        self.member = self.member_repo.add("สมชาย ใจดี", "ม.4/1", "30",
                                           "2026-10-30", "2027-10-30", "x", "y")
        self.book = self.book_repo.add("T-001", "หนังสือ A", "ว่าง")
        self.book2 = self.book_repo.add("T-002", "หนังสือ B", "ว่าง")

    def _status(self, book_id):
        return self.book_repo.fetchone("SELECT status FROM books WHERE id=?",
                                       (book_id,))[0]

    # ---------------- happy path ----------------

    def test_borrow_then_return_updates_both_sides(self):
        self._reset()
        rec = self.borrow_repo.try_borrow(self.member, self.book,
                                          "2026-10-30", "2026-11-06")
        self.assertIsNotNone(rec)
        self.assertEqual(self._status(self.book), "ยืมแล้ว")
        returned_book = self.borrow_repo.try_return(rec)
        self.assertEqual(returned_book, self.book)
        self.assertEqual(self._status(self.book), "ว่าง")
        self.assertEqual(len(self.borrow_repo.get_active_borrows()), 0)

    # audit-test: two simultaneous attempts to borrow one book -> at most one loan
    def test_concurrent_double_borrow_lends_once(self):
        self._reset()
        results = {"ok": 0, "fail": 0, "lock": threading.Lock()}

        def attempt():
            from db.repositories.borrow_repo import BorrowError
            try:
                self.borrow_repo.try_borrow(self.member, self.book,
                                            "2026-10-30", "2026-11-06")
                with results["lock"]:
                    results["ok"] += 1
            except BorrowError:
                with results["lock"]:
                    results["fail"] += 1

        threads = [threading.Thread(target=attempt) for _ in range(2)]
        for t in threads:
            t.start()
        for t in threads:
            t.join(30)

        self.assertEqual(results["ok"], 1, "only one borrower may win")
        self.assertEqual(results["fail"], 1)
        active = self.borrow_repo.fetchall(
            "SELECT COUNT(*) FROM borrow_log WHERE returned=0")[0][0]
        self.assertEqual(active, 1)

    # audit-test: database error midway through borrowing -> no partial transaction
    def test_fk_failure_midway_rolls_back_book_status(self):
        self._reset()
        from db.repositories.borrow_repo import BorrowError
        with self.assertRaises(BorrowError):
            # nonexistent member: the book status is flipped first (as it would
            # be in the real flow), THEN the member/INSERT step fails -> the
            # transaction must roll back everything.
            self.borrow_repo.try_borrow(9999, self.book,
                                        "2026-10-30", "2026-11-06")
        self.assertEqual(self._status(self.book), "ว่าง",
                         "book status must not stay flipped after a failed borrow")
        count = self.borrow_repo.fetchone("SELECT COUNT(*) FROM borrow_log")[0]
        self.assertEqual(count, 0, "no orphan borrow row may remain")

    def test_locked_out_after_overdue(self):
        self._reset()
        from db.repositories.borrow_repo import BorrowError
        conn = self.borrow_repo.get_connection()
        try:
            conn.execute(
                "INSERT INTO borrow_log (member_id, book_id, borrow_date, return_due, returned) "
                "VALUES (?, ?, '2026-01-01', '2026-01-08', 0)", (self.member, self.book))
        finally:
            conn.close()
        self.book_repo.update_status(self.book2, "ว่าง")
        with self.assertRaises(BorrowError):
            self.borrow_repo.try_borrow(self.member, self.book2,
                                        "2026-10-30", "2026-11-06")
        self.assertEqual(self._status(self.book2), "ว่าง")

    def test_max_books_limit_enforced(self):
        self._reset()
        from db.repositories.borrow_repo import BorrowError
        from config.settings import DEFAULTS
        max_books = DEFAULTS["max_books_per_member"]
        for i in range(max_books):
            extra = self.book_repo.add(f"T-9{i:02d}", f"หนังสือ {i}", "ว่าง")
            self.borrow_repo.try_borrow(self.member, extra, "2026-10-30", "2026-11-06")
        with self.assertRaises(BorrowError):
            self.borrow_repo.try_borrow(self.member, self.book,
                                        "2026-10-30", "2026-11-06")
        self.assertEqual(self._status(self.book), "ว่าง")

    # audit-test: return nonexistent/mismatched record -> rejected w/o side effects
    def test_return_nonexistent_record_rejected(self):
        self._reset()
        from db.repositories.borrow_repo import BorrowError
        with self.assertRaises(BorrowError):
            self.borrow_repo.try_return(99999)
        self.assertEqual(self._status(self.book), "ว่าง")

    def test_return_cannot_flip_another_book(self):
        self._reset()
        # record 1 borrows book1; an attacker-supplied book_id for book2 is IGNORED
        self.borrow_repo.try_borrow(self.member, self.book,
                                    "2026-10-30", "2026-11-06")
        rec2 = self.borrow_repo.try_borrow(self.member, self.book2,
                                           "2026-10-30", "2026-11-06")
        # returning record2 must only free book2
        freed = self.borrow_repo.try_return(rec2)
        self.assertEqual(freed, self.book2)
        self.assertEqual(self._status(self.book2), "ว่าง")
        self.assertEqual(self._status(self.book), "ยืมแล้ว")

    # audit-test: SQL injection strings in search fields -> ordinary input
    def test_sqli_string_is_ordinary_data(self):
        self._reset()
        payload = "'; DROP TABLE books;--"
        row = self.book_repo.fetchone("SELECT * FROM books WHERE code = ?", (payload,))
        self.assertIsNone(row)
        # the table still exists and is intact
        titles = self.book_repo.fetchall("SELECT title FROM books ORDER BY id")
        self.assertEqual(len(titles), 2)
        # full payload stored as a literal code (parameterized insert)
        malicious = self.book_repo.add(payload, "หนังสือ SQL Injection")
        found = self.book_repo.get_by_code(payload)
        self.assertIsNotNone(found)
        self.assertEqual(found[0], malicious)


if __name__ == "__main__":
    unittest.main()
