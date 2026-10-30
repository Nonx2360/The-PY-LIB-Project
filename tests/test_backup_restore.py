# tests/test_backup_restore.py — verified backup / restore / integrity checks
import os
import sqlite3
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402


class TestBackupRestore(unittest.TestCase):
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
        from services.backup_service import create_backup, verify_backup
        cls.member_repo = MemberRepository()
        cls.book_repo = BookRepository()
        cls.borrow_repo = BorrowRepository()
        cls.create_backup = create_backup
        cls.verify_backup = verify_backup

        # seed initial data, then snapshot it
        cls.member = cls.member_repo.add("สมใจ เรียนดี", "ม.2/1", "7",
                                         "2026-10-30", "2027-10-30", "x", "y")
        cls.book = cls.book_repo.add("T-100", "หนังสือต้ั้งพื้นฐาน", "ว่าง")
        cls.record = cls.borrow_repo.try_borrow(cls.member, cls.book,
                                                "2026-10-30", "2026-11-06")
        cls.backup_path = cls.create_backup(cls.env.db_path, backup_dir=cls.env.dir)
        cls.report = cls.verify_backup(cls.backup_path)

    @classmethod
    def tearDownClass(cls):
        cls.env.restore()
        harness.cleanup(cls.env.dir)

    # audit-test: Backup and restore -> records and relationships preserved

    def test_backup_passes_verification(self):
        self.assertTrue(self.report["ok"], f"backup verification failed: {self.report}")

    def test_restore_preserves_records_and_relationships(self):
        # mutate live data after the backup
        self.book_repo.add("T-999", "หนังสือหลังแบ็กอัพ", "ว่าง")
        from services.backup_service import restore_backup
        result = restore_backup(self.backup_path, self.env.db_path)
        self.assertTrue(result["verification"]["ok"])
        conn = sqlite3.connect(self.env.db_path)
        try:
            books_after = conn.execute(
                "SELECT id, code, title FROM books ORDER BY id").fetchall()
            borrowed = conn.execute(
                """SELECT bl.id, bl.member_id, bl.book_id
                   FROM borrow_log bl
                   JOIN members m ON bl.member_id = m.id
                   JOIN books b ON bl.book_id = b.id""").fetchall()
            fk = conn.execute("PRAGMA foreign_key_check").fetchall()
        finally:
            conn.close()
        self.assertEqual([(b[1], b[2]) for b in books_after], [("T-100", "หนังสือต้ั้งพื้นฐาน")])
        self.assertEqual(len(borrowed), 1)
        self.assertEqual(borrowed[0][2], books_after[0][0],
                         "borrow_log must still reference the restored book")
        self.assertEqual(borrowed[0][1], self.member)
        self.assertEqual(fk, [], "restored db must pass foreign_key_check")

    def test_corrupt_backup_is_refused(self):
        from services.backup_service import restore_backup
        corrupt = os.path.join(self.env.dir, "corrupt.db")
        with open(corrupt, "wb") as f:
            f.write(b"this is not a sqlite database " * 64)
        with self.assertRaises(ValueError):
            restore_backup(corrupt, self.env.db_path)
        # live db untouched and still readable
        conn = sqlite3.connect(self.env.db_path)
        try:
            ok = conn.execute("PRAGMA integrity_check").fetchone()[0]
        finally:
            # explicit close: `with` on sqlite3 does NOT close, and an open
            # handle would make the later os.replace() restore fail on Windows
            conn.close()
        self.assertEqual(ok, "ok")

    def test_backup_restoration_of_records(self):
        # verify the snapshot contains all rows present at snapshot time
        conn = sqlite3.connect(self.backup_path)
        try:
            self.assertIsNotNone(conn.execute(
                "SELECT id FROM members WHERE id=?", (self.member,)).fetchone())
            self.assertIsNotNone(conn.execute(
                "SELECT id FROM books WHERE id=?", (self.book,)).fetchone())
            self.assertIsNotNone(conn.execute(
                "SELECT id FROM borrow_log WHERE id=?", (self.record,)).fetchone())
        finally:
            conn.close()


if __name__ == "__main__":
    unittest.main()
