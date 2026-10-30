# tests/test_web_security.py — route auth, CSRF, throttling, validation
import os
import re
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

_env = harness.TempEnv()
_env.apply()
_env.patch()

import web.app as web_app_mod  # noqa: E402
from config.settings import DEFAULTS  # noqa: E402
web_app_mod.DB_PATH = _env.db_path  # login uses a module-level binding

CSRF_RE = re.compile(r'name="csrf_token" value="([^"]+)"')

PROTECTED_ROUTES = ["/", "/members", "/books", "/borrow", "/return",
                    "/history", "/access", "/settings", "/api/stats"]


class TestWebSecurity(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from config import settings as settings_mod
        cls._orig_settings_file = settings_mod.SETTINGS_FILE
        settings_mod.SETTINGS_FILE = os.path.join(_env.dir, "settings.json")
        cls.app = web_app_mod.create_app()
        cls.client = cls.app.test_client()

    @classmethod
    def tearDownClass(cls):
        from config import settings as settings_mod
        settings_mod.SETTINGS_FILE = cls._orig_settings_file
        harness.cleanup(_env.dir)

    # ---------------- helpers ----------------

    def _csrf_token(self, path="/login"):
        resp = self.client.get(path, follow_redirects=True)
        match = CSRF_RE.search(resp.get_data(as_text=True))
        self.assertIsNotNone(match, f"no csrf token rendered on {path}")
        return match.group(1)

    def _first_boot_password(self):
        path = os.path.join(os.path.dirname(_env.db_path), "INITIAL_ADMIN_PASSWORD.txt")
        with open(path, encoding="utf-8") as f:
            for line in f.read().splitlines():
                if "admin:" in line:
                    return line.split("admin:", 1)[1].split(" (", 1)[0].strip()
        self.fail("no generated admin password file")

    def _login(self, username="admin", password=None):
        if password is None:
            password = self._first_boot_password()
        token = self._csrf_token()
        return self.client.post("/login", data={
            "username": username, "password": password, "csrf_token": token,
        })

    def _add_book(self, code="W-001", title="หนังสือเว็บ"):
        token = self._csrf_token("/books")
        return self.client.post("/books/add", data={"code": code, "title": title,
                                                    "csrf_token": token})

    def _count(self, table):
        return harness.query(_env.db_path, f"SELECT COUNT(*) FROM {table}")[0][0]

    # ---------------- auth (audit-test: wrong password rejected) ----------------

    def test_login_wrong_password_rejected(self):
        resp = self._login(password="not-the-password")
        body = resp.get_data(as_text=True)
        self.assertIn("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง", body)

    def test_login_success_with_generated_password(self):
        resp = self._login()
        self.assertEqual(resp.status_code, 302)
        resp = self.client.get("/settings")
        self.assertEqual(resp.status_code, 200)
        self.assertIn("การตั้งค่าระบบ", resp.get_data(as_text=True))

    # ---------------- unauthorized access (audit-test) ----------------

    def test_unauthenticated_requests_redirect(self):
        self.client.get("/logout")
        for route in PROTECTED_ROUTES:
            with self.subTest(route=route):
                resp = self.client.get(route)
                self.assertEqual(resp.status_code, 302, route)
                self.assertIn("/login", resp.headers["Location"])

    def test_unauthenticated_post_rejected(self):
        self.client.get("/logout")
        before = self._count("books")
        resp = self.client.post("/books/add", data={"code": "X", "title": "Y"})
        self.assertEqual(resp.status_code, 302)
        self.assertIn("/login", resp.headers["Location"])
        self.assertEqual(self._count("books"), before)

    # ---------------- CSRF (audit-test: POST without token -> rejected) ----------------

    def test_post_without_csrf_rejected_and_state_unchanged(self):
        self._login()
        before = self._count("books")
        resp = self.client.post("/books/add", data={"code": "NOC-1", "title": "no token"})
        self.assertEqual(resp.status_code, 302)  # rejected via redirect
        self.assertEqual(self._count("books"), before, "state must not change")

    def test_post_with_csrf_succeeds(self):
        before = self._count("books")
        resp = self._add_book("W-001")
        status = resp.status_code
        flash_page = self.client.get("/books").get_data(as_text=True)
        detail = (f"POST status={status}; page-contains-W-001={'W-001' in flash_page}; "
                  f"alert={flash_page[flash_page.find('alert'):][:200] if 'alert' in flash_page else 'none'}")
        self.assertEqual(self._count("books"), before + 1, detail)

    def test_csrf_token_is_per_session(self):
        # token issued to session A must not authorize session B
        self._login()
        token = self._csrf_token("/books")
        self.client.get("/logout")
        resp = self.client.post("/books/add", data={
            "code": "ST-1", "title": "stolen token", "csrf_token": token})
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(self._count("books"), 0, "foreign token must not authorize")

    # ---------------- throttling (audit-test: login throttling) ----------------

    def test_login_throttling_after_failures(self):
        username = "throttle-probe"  # isolated key so earlier tests don't interfere
        for i in range(5):
            self.client.post("/login", data={
                "username": username, "password": f"wrong-{i}",
                "csrf_token": self._csrf_token()})
        resp = self.client.post("/login", data={
            "username": username, "password": "whatever", "csrf_token": self._csrf_token()})
        self.assertEqual(resp.status_code, 429, "6th attempt must be throttled")
        self.assertIn("นาที", resp.get_data(as_text=True))

    def test_throttle_cleared_on_success(self):
        self._login()
        # another immediate failed attempt is allowed (count reset by success)
        resp = self._login(password="bad-bad")
        self.assertEqual(resp.status_code, 200)
        self.assertIn("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง", resp.get_data(as_text=True))
        self.assertNotIn("429", str(resp.status_code))

    # ---------------- validation ----------------

    def test_settings_post_clamps_out_of_bounds_values(self):
        self._login()
        token = self._csrf_token("/settings")
        self.client.post("/settings", data={
            "school_name": "DSNPRU", "default_loan_days": "99999",
            "max_books_per_member": "-5", "member_expiry_days": "not-an-int",
            "school_logo_path": "assets/logos/school_logo.png",
            "csrf_token": token})
        from config import settings
        self.assertTrue(1 <= settings.get("default_loan_days") <= 90)
        self.assertTrue(1 <= settings.get("max_books_per_member") <= 20)
        self.assertEqual(settings.get("member_expiry_days"), DEFAULTS["member_expiry_days"])

    # ---------------- SQL injection as ordinary data (audit-test) ----------------

    def test_sqli_in_web_fields_treated_as_data(self):
        import db.repositories.book_repo as br_mod
        repo = br_mod.BookRepository()
        payload = "X'; DROP TABLE books;--"
        repo.add(payload, "หนังสือ SQLi")
        row = repo.get_by_code(payload)
        self.assertIsNotNone(row, "payload stored as literal code")
        tables = harness.query(_env.db_path,
                               "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='books'"
                               )[0][0]
        self.assertEqual(tables, 1, "books table must still exist")
        repo.delete(row[0])


if __name__ == "__main__":
    unittest.main()
