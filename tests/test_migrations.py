# tests/test_migrations.py — user_version migrations + integrity constraint enforcement
import os
import shutil
import sqlite3
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness  # noqa: E402

from db.repositories.base_repo import BaseRepository  # noqa: E402


LEGACY_SCHEMA = """
CREATE TABLE admin_users (username TEXT PRIMARY KEY, password TEXT);
CREATE TABLE members (id INTEGER PRIMARY KEY, name TEXT, grade TEXT, number TEXT,
    register_date TEXT, expire_date TEXT, qrcode_data TEXT, qrcode_path TEXT);
CREATE TABLE books (id INTEGER PRIMARY KEY, code TEXT, title TEXT, status TEXT);
CREATE TABLE borrow_log (id INTEGER PRIMARY KEY, member_id INTEGER, book_id INTEGER,
    borrow_date TEXT, return_due TEXT, returned INTEGER,
    FOREIGN KEY (member_id) REFERENCES members (id),
    FOREIGN KEY (book_id) REFERENCES books (id));
CREATE TABLE access_log (id INTEGER PRIMARY KEY, member_id INTEGER,
    access_time DATETIME, action TEXT, FOREIGN KEY (member_id) REFERENCES members (id));
"""

BOOKS = [
    (1, "T-001", "A"),
    (2, "T-002", "B"),
    (3, "T-002", "B-copy"),
    (4, "T-001", "A-copy"),
    (5, "T-003", "C"),
]

BORROWS = [
    (1, 1, 1, "2026-01-02", "2026-01-09", 0),      # valid open loan
    (2, 1, 3, "2026-02-02", "2026-02-09", 1),      # valid returned
    (3, 1, 4, "2026-03-02", "2026-03-09", 5),      # legacy invalid value
    (4, 1, 5, "2026-04-02", "2026-04-09", None),   # legacy NULL
]


class TestMigrations(unittest.TestCase):
    dirs_to_remove = []

    @classmethod
    def _make_legacy_db(cls, path):
        conn = sqlite3.connect(path)
        try:
            conn.executescript(LEGACY_SCHEMA)
            conn.execute("INSERT INTO admin_users VALUES ('admin', 'not-a-real-hash')")
            conn.execute("INSERT INTO members VALUES (1, 'สมใจ', 'ม.1/1', '12', "
                         "'2026-01-01', '2027-01-01', 'x', 'y')")
            conn.executemany("INSERT INTO books VALUES (?,?,?, 'ว่าง')", [b[0:2] + (b[2],) for b in BOOKS])
            conn.executemany("INSERT INTO borrow_log VALUES (?,?,?,?,?,?)", BORROWS)
            conn.commit()
        finally:
            conn.close()

    @classmethod
    def setUpClass(cls):
        import db.schema as schema_mod
        import db.repositories.base_repo as base_mod
        cls.schema_mod = schema_mod
        cls._orig_schema = schema_mod.DB_PATH
        cls._orig_base = base_mod.DB_PATH

        cls.dir = os.path.join(os.path.dirname(os.path.realpath(__file__)), "_migrations_tmp")
        shutil.rmtree(cls.dir, ignore_errors=True)  # wipe leftovers from any earlier run
        os.makedirs(cls.dir, exist_ok=True)
        cls.db_path = os.path.join(cls.dir, "legacy.db")
        cls._make_legacy_db(cls.db_path)

        schema_mod.DB_PATH = cls.db_path
        base_mod.DB_PATH = cls.db_path
        schema_mod.init_db(log=lambda *a, **k: None)
        cls.dirs_to_remove.append(cls.dir)

    @classmethod
    def tearDownClass(cls):
        cls.schema_mod.DB_PATH = cls._orig_schema
        import db.repositories.base_repo as base_mod
        base_mod.DB_PATH = cls._orig_base
        for d in cls.dirs_to_remove:
            shutil.rmtree(d, ignore_errors=True)

    # ---------------- upgrade results ----------------

    def test_user_version_is_2(self):
        version = self._user_version()
        self.assertEqual(version, 2)

    def _user_version(self):
        conn = sqlite3.connect(self.db_path)
        try:
            return conn.execute("PRAGMA user_version").fetchone()[0]
        finally:
            conn.close()

    def test_duplicate_book_codes_renumbered_and_unique(self):
        conn = sqlite3.connect(self.db_path)
        try:
            codes = [r[0] for r in conn.execute("SELECT code FROM books").fetchall()]
        finally:
            conn.close()
        self.assertEqual(len(codes), len(set(codes)), f"codes not unique: {codes}")
        self.assertEqual(len(codes), 5)
        self.assertIn("T-001", codes)
        self.assertIn("T-003", codes)

    def test_borrow_rows_preserved_and_states_valid(self):
        conn = sqlite3.connect(self.db_path)
        try:
            rows = conn.execute("SELECT id, returned FROM borrow_log ORDER BY id").fetchall()
        finally:
            conn.close()
        by_id = dict(rows)
        self.assertEqual(len(rows), 4)
        self.assertEqual(by_id[1], 0, "valid open loan preserved")
        self.assertEqual(by_id[2], 1, "valid returned loan preserved")
        self.assertEqual(by_id[3], 1, "legacy invalid value coerced to returned")
        self.assertEqual(by_id[4], 1, "legacy NULL coerced to returned")

    # ---------------- DB-enforced constraints (audit tests) ----------------

    def test_unique_code_enforced(self):
        repo = BaseRepository()
        with self.assertRaises(sqlite3.IntegrityError):
            repo.execute("INSERT INTO books (code, title, status) VALUES (?, ?, ?)",
                         ("T-001", "dup", "ว่าง"))

    def test_borrow_state_check_enforced(self):
        repo = BaseRepository()
        with self.assertRaises(sqlite3.IntegrityError):
            repo.execute(
                "INSERT INTO borrow_log (member_id, book_id, borrow_date, return_due, returned) "
                "VALUES (?, ?, ?, ?, ?)", (1, 1, "2026-10-30", "2026-11-06", 2))

    def test_borrow_notnull_member_enforced(self):
        repo = BaseRepository()
        with self.assertRaises(sqlite3.IntegrityError):
            repo.execute(
                "INSERT INTO borrow_log (member_id, book_id, borrow_date, return_due, returned) "
                "VALUES (NULL, 1, '2026-10-30', '2026-11-06', 0)")

    def test_foreign_keys_enforced_on_repository_connections(self):
        # FK enforcement must be ON for every repository connection
        repo = BaseRepository()
        conn = repo.get_connection()
        try:
            fk_state = conn.execute("PRAGMA foreign_keys").fetchone()[0]
        finally:
            conn.close()
        self.assertEqual(fk_state, 1)
        with self.assertRaises(sqlite3.IntegrityError):
            repo.execute(
                "INSERT INTO borrow_log (member_id, book_id, borrow_date, return_due, returned) "
                "VALUES (9999, 1, '2026-10-30', '2026-11-06', 0)")

    def test_migrations_are_idempotent(self):
        self.schema_mod.init_db(log=lambda *a, **k: None)
        self.assertEqual(self._user_version(), 2)


class TestFreshInstallLayout(unittest.TestCase):
    """Fresh installs are created directly in the constrained (v2) layout."""

    @classmethod
    def setUpClass(cls):
        import db.schema as schema_mod
        import db.repositories.base_repo as base_mod
        cls.schema_mod = schema_mod
        cls._orig_schema = schema_mod.DB_PATH
        cls._orig_base = base_mod.DB_PATH

        cls.dir = os.path.join(os.path.dirname(os.path.realpath(__file__)), "_fresh_tmp")
        shutil.rmtree(cls.dir, ignore_errors=True)  # wipe leftovers from any earlier run
        os.makedirs(cls.dir, exist_ok=True)
        cls.db_path = os.path.join(cls.dir, "fresh.db")
        schema_mod.DB_PATH = cls.db_path
        base_mod.DB_PATH = cls.db_path
        cls.dirs_to_remove = [cls.dir]
        schema_mod.init_db(log=lambda *a, **k: None)

    @classmethod
    def tearDownClass(cls):
        cls.schema_mod.DB_PATH = cls._orig_schema
        import db.repositories.base_repo as base_mod
        base_mod.DB_PATH = cls._orig_base
        for d in cls.dirs_to_remove:
            shutil.rmtree(d, ignore_errors=True)

    def test_no_default_admin_password(self):
        stored = _harness.query(
            self.db_path,
            "SELECT password FROM admin_users WHERE username='admin'")[0]
        self.assertIsNotNone(stored)
        from core.security import verify_password
        self.assertFalse(verify_password("admin123", stored[0]))

    def test_constraints_apply_immediately(self):
        repo = BaseRepository()
        repo.execute("INSERT INTO books (code, title, status) VALUES (?, ?, ?)",
                     ("F-001", "fresh", "ว่าง"))
        with self.assertRaises(sqlite3.IntegrityError):
            repo.execute("INSERT INTO books (code, title, status) VALUES (?, ?, ?)",
                         ("F-001", "dup", "ว่าง"))


if __name__ == "__main__":
    unittest.main()
