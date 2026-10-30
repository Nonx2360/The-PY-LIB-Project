# tests/test_member_repo.py
import os
import tempfile
import unittest

import db.repositories.base_repo as base_repo_mod
import db.schema
from db.repositories.member_repo import MemberRepository


class TestMemberRepository(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp_dir = tempfile.mkdtemp()
        cls.db_path = os.path.join(cls._tmp_dir, "test.db")
        # Point the repository module at the temp DB
        cls._orig_base_path = base_repo_mod.DB_PATH
        base_repo_mod.DB_PATH = cls.db_path
        cls._orig_schema_path = db.schema.DB_PATH
        db.schema.DB_PATH = cls.db_path
        db.schema.init_db()

    @classmethod
    def tearDownClass(cls):
        # restore EVERY global this class patched — a leak here would redirect
        # later test classes (e.g. the web suite) to this temp path
        db.schema.DB_PATH = cls._orig_schema_path
        base_repo_mod.DB_PATH = cls._orig_base_path
        os.remove(cls.db_path)

    def setUp(self):
        self.repo = MemberRepository()
        conn = self.repo.get_connection()
        try:
            conn.execute("DELETE FROM members")
        finally:
            conn.close()

    def test_add_and_get_all(self):
        mid = self.repo.add("สมชาย ใจดี", "ม.1/2", "15", "2026-01-01", "2027-01-01",
                            "qr-data", "assets/qrcodes/x.png")
        self.assertIsNotNone(mid)
        rows = self.repo.get_all()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][1], "สมชาย ใจดี")

    def test_get_by_id(self):
        mid = self.repo.add("สมหญิง", "ม.2/1", "3", "2026-01-01", "2027-01-01",
                            "abc", "assets/qrcodes/y.png")
        row = self.repo.get_by_id(mid)
        self.assertEqual(row[1], "สมหญิง")

    def test_update(self):
        mid = self.repo.add("เดิม", "ม.1/1", "1", "2026-01-01", "2027-01-01",
                            "x", "assets/qrcodes/z.png")
        self.repo.update(mid, name="ใหม่", grade="ม.3/1", number="2",
                         exp_date="2028-01-01", qr_data="y", qr_path="assets/qrcodes/w.png")
        row = self.repo.get_by_id(mid)
        self.assertEqual(row[1], "ใหม่")
        self.assertEqual(row[2], "ม.3/1")

    def test_delete(self):
        mid = self.repo.add("ลบทิ้ง", "ม.1/1", "1", "2026-01-01", "2027-01-01",
                            "x", "assets/qrcodes/v.png")
        self.repo.delete(mid)
        self.assertIsNone(self.repo.get_by_id(mid))


if __name__ == "__main__":
    unittest.main()
