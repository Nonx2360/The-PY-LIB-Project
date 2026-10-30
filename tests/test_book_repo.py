# tests/test_book_repo.py
import os
import tempfile
import unittest

import db.repositories.base_repo as base_repo_mod
import db.schema
from db.repositories.book_repo import BookRepository


class TestBookRepository(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp_dir = tempfile.mkdtemp()
        cls.db_path = os.path.join(cls._tmp_dir, "test.db")
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
        self.repo = BookRepository()
        conn = self.repo.get_connection()
        try:
            conn.execute("DELETE FROM books")
        finally:
            conn.close()

    def test_add_get_all(self):
        bid = self.repo.add("B-001", "Python เบื้องต้น", "available")
        self.assertIsNotNone(bid)
        rows = self.repo.get_all()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][1], "B-001")

    def test_get_by_code(self):
        self.repo.add("B-002", "คณิตศาสตร์", "available")
        row = self.repo.get_by_code("B-002")
        self.assertEqual(row[2], "คณิตศาสตร์")
        self.assertIsNone(self.repo.get_by_code("MISSING"))

    def test_update_status(self):
        bid = self.repo.add("B-003", "ฟิสิกส์", "available")
        self.repo.update_status(bid, "borrowed")
        row = self.repo.get_by_code("B-003")
        self.assertEqual(row[3], "borrowed")

    def test_delete(self):
        bid = self.repo.add("B-004", "เคมี", "available")
        self.repo.delete(bid)
        self.assertIsNone(self.repo.get_by_code("B-004"))


if __name__ == "__main__":
    unittest.main()
