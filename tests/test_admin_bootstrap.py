# tests/test_admin_bootstrap.py — no shared default admin password remains
import os
import sqlite3
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness  # noqa: E402


def _bootstrap(env: "_harness.TempEnv"):
    env.apply()
    env.patch()
    import db.schema as schema_mod
    import core.security  # noqa: F401 — loads secrets from the temp .env
    return schema_mod


class TestAdminBootstrap(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.env = _harness.TempEnv()
        cls.schema = _bootstrap(cls.env)
        cls.schema.init_db(log=lambda *a, **k: None)

    @classmethod
    def tearDownClass(cls):
        cls.env.restore()
        _harness.cleanup(cls.env.dir)

    def _admin_hash(self):
        return _harness.query(
            self.env.db_path,
            "SELECT password FROM admin_users WHERE username='admin'")[0][0]

    def _first_boot_password(self):
        """The generated one-time password recorded by the bootstrap file."""
        path = os.path.join(os.path.dirname(self.env.db_path), "INITIAL_ADMIN_PASSWORD.txt")
        self.assertTrue(os.path.exists(path), "one-time credentials file was not written")
        with open(path, encoding="utf-8") as f:
            lines = [ln for ln in f.read().splitlines() if "admin:" in ln]
        self.assertTrue(lines)
        return lines[-1].split("admin:", 1)[1].split(" (", 1)[0].strip()

    # ---- audit-test: First launch -> No shared default password remains ----
    def test_fresh_boot_has_no_default_password(self):
        from core.security import verify_password
        self.schema.init_db()
        stored = self._admin_hash()
        self.assertFalse(verify_password("admin123", stored),
                         "default password must not authenticate")
        boot_pw = self._first_boot_password()
        self.assertTrue(verify_password(boot_pw, stored),
                        "generated one-time password must work")

    def test_boot_is_idempotent(self):
        from core.security import verify_password
        pw = self._first_boot_password()
        self.schema.init_db()
        self.schema.init_db()
        stored = self._admin_hash()
        self.assertTrue(verify_password(pw, stored))
        self.assertFalse(verify_password("admin123", stored))

    # ---- audit-test: historic default credential is retired on upgrade ----
    def test_legacy_db_with_default_admin123_is_retired(self):
        import db.schema as schema_mod
        db2 = os.path.join(self.env.dir, "legacy.db")
        conn = sqlite3.connect(db2)
        try:
            conn.executescript("""
                CREATE TABLE admin_users (username TEXT PRIMARY KEY, password TEXT);
                CREATE TABLE members (id INTEGER PRIMARY KEY, name TEXT, grade TEXT,
                    number TEXT, register_date TEXT, expire_date TEXT,
                    qrcode_data TEXT, qrcode_path TEXT);
                CREATE TABLE books (id INTEGER PRIMARY KEY, code TEXT, title TEXT, status TEXT);
                CREATE TABLE borrow_log (id INTEGER PRIMARY KEY, member_id INTEGER,
                    book_id INTEGER, borrow_date TEXT, return_due TEXT, returned INTEGER);
                CREATE TABLE access_log (id INTEGER PRIMARY KEY, member_id INTEGER,
                    access_time DATETIME, action TEXT);
                INSERT INTO admin_users VALUES ('admin', 'admin123');
            """)
            conn.commit()
        finally:
            conn.close()

        orig = schema_mod.DB_PATH
        try:
            schema_mod.DB_PATH = db2
            schema_mod.init_db()
            conn = sqlite3.connect(db2)
            try:
                stored = conn.execute(
                    "SELECT password FROM admin_users WHERE username='admin'").fetchone()[0]
            finally:
                conn.close()
            from core.security import verify_password
            self.assertNotEqual(stored, "admin123")
            self.assertFalse(verify_password("admin123", stored))
            self.assertTrue(stored.startswith("$2"), "retired password must be bcrypt-hashed")
        finally:
            schema_mod.DB_PATH = orig

    def test_failed_login_rejects_wrong_password(self):
        # audit-test: Wrong administrator password -> Login rejected
        from core.security import verify_password
        stored = self._admin_hash()
        self.assertFalse(verify_password("wrong-password", stored))


if __name__ == "__main__":
    unittest.main()
