# tests/test_web_delete_guards.py — delete routes refuse history-bearing rows
#
# FK policy (schema v2): borrow_log -> members/books and access_log ->
# members must stay intact, so a delete of a referenced row must be refused
# with a Thai message — never an HTTP 500 or a raw DB constraint error.
import os
import re
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

_env = harness.TempEnv()

# NOTE: env/binding mutations happen in setUpClass, NOT at import (see
# test_web_scan.py for the discovery-collision explanation).
import web.app as web_app_mod  # noqa: E402

CSRF_RE = re.compile(r'name="csrf_token" value="([^"]+)"')


class TestWebDeleteGuards(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from config import settings as settings_mod
        _env.apply()
        _env.patch()
        web_app_mod.DB_PATH = _env.db_path
        cls._orig_settings_file = settings_mod.SETTINGS_FILE
        settings_mod.SETTINGS_FILE = os.path.join(_env.dir, "settings.json")
        # member/book file artifacts (QR PNGs, card PDFs) are written relative
        # to the CWD by design — chdir into the temp env so tests can never
        # touch or overwrite files in the real assets/ folders again.
        cls._orig_cwd = os.getcwd()
        os.chdir(_env.dir)
        harness.stage_static_assets(_env.dir)
        cls.app = web_app_mod.create_app()
        cls.client = cls.app.test_client()
        cls._login()

    @classmethod
    def tearDownClass(cls):
        from config import settings as settings_mod
        settings_mod.SETTINGS_FILE = cls._orig_settings_file
        os.chdir(cls._orig_cwd)
        _env.restore()
        harness.cleanup(_env.dir)

    # ---------------- helpers ----------------

    @classmethod
    def _csrf(cls, path):
        resp = cls.client.get(path)
        match = CSRF_RE.search(resp.get_data(as_text=True))
        if not match:
            resp = cls.client.get("/borrow")
            match = CSRF_RE.search(resp.get_data(as_text=True))
        return match.group(1)

    @classmethod
    def _login(cls):
        # first-boot admin password file written by init_db()
        pwd_path = os.path.join(os.path.dirname(_env.db_path),
                                "INITIAL_ADMIN_PASSWORD.txt")
        with open(pwd_path, encoding="utf-8") as fh:
            content = fh.read()
        token = cls._csrf("/login")
        for line in content.splitlines():
            if "admin:" in line:
                password = line.split("admin:", 1)[1].split(" (", 1)[0].strip()
                resp = cls.client.post("/login", data={
                    "username": "admin", "password": password,
                    "csrf_token": token,
                })
                if resp.status_code != 302:  # classmethods can't use assertEqual
                    raise AssertionError(f"login failed: HTTP {resp.status_code}")
                break

    def setUp(self):
        self._login()

    def _add_member(self, number):
        resp = self.client.post("/members/add", data={
            "name": f"นักเรียนทดสอบ {number}", "grade": "ม.4/1", "number": number,
            "csrf_token": self._csrf("/members")})
        self.assertEqual(resp.status_code, 302)
        return harness.query(_env.db_path,
                             "SELECT id FROM members WHERE number=?",
                             (number,))[0][0]

    def _add_book(self, code):
        resp = self.client.post("/books/add", data={
            "code": code, "title": "หนังสือทดสอบ",
            "csrf_token": self._csrf("/books")})
        self.assertEqual(resp.status_code, 302)
        return harness.query(_env.db_path,
                             "SELECT id FROM books WHERE code=?", (code,))[0][0]

    def _borrow(self, member_id, book_id):
        resp = self.client.post("/borrow", data={
            "member_id": member_id, "book_id": book_id, "due_date": "2027-12-31",
            "csrf_token": self._csrf("/borrow")})
        self.assertEqual(resp.status_code, 302)

    def _delete_book(self, bid):
        return self.client.post(f"/books/{bid}/delete",
                                data={"csrf_token": self._csrf("/books")},
                                follow_redirects=True)

    def _delete_member(self, mid):
        return self.client.post(f"/members/{mid}/delete",
                                data={"csrf_token": self._csrf("/members")},
                                follow_redirects=True)

    # ---------------- tests ----------------

    def test_delete_book_with_returned_history_refused(self):
        # borrow AND return, then try to delete: history rows block it
        mid = self._add_member("30001")
        bid = self._add_book("G-001")
        self._borrow(mid, bid)
        record = harness.query(_env.db_path,
                               "SELECT id FROM borrow_log ORDER BY id DESC LIMIT 1")[0][0]
        resp = self.client.post("/return", data={
            "record_id": record, "csrf_token": self._csrf("/return")},
            follow_redirects=True)
        self.assertEqual(resp.status_code, 200)

        resp = self._delete_book(bid)
        body = resp.get_data(as_text=True)
        self.assertIn("มีประวัติการยืม", body)
        self.assertNotIn("ลบหนังสือสำเร็จ", body)
        remaining = harness.query(_env.db_path,
                                  "SELECT COUNT(*) FROM books WHERE id=?", (bid,))
        self.assertEqual(remaining, [(1,)])

    def test_delete_book_with_active_borrow_refused(self):
        mid = self._add_member("30002")
        bid = self._add_book("G-002")
        self._borrow(mid, bid)
        resp = self._delete_book(bid)
        body = resp.get_data(as_text=True)
        self.assertIn("มีการยืมหนังสือเล่มนี้อยู่", body)
        remaining = harness.query(_env.db_path,
                                  "SELECT COUNT(*) FROM books WHERE id=?", (bid,))
        self.assertEqual(remaining, [(1,)])

    def test_delete_book_without_history_succeeds(self):
        bid = self._add_book("G-003")
        resp = self._delete_book(bid)
        self.assertIn("ลบหนังสือสำเร็จ", resp.get_data(as_text=True))
        remaining = harness.query(_env.db_path,
                                  "SELECT COUNT(*) FROM books WHERE id=?", (bid,))
        self.assertEqual(remaining, [(0,)])

    def test_delete_member_with_borrow_history_refused(self):
        mid = self._add_member("30003")
        bid = self._add_book("G-004")
        self._borrow(mid, bid)
        resp = self._delete_member(mid)
        body = resp.get_data(as_text=True)
        self.assertIn("ประวัติการยืมหนังสือ", body)
        self.assertNotIn("ลบสมาชิกสำเร็จ", body)
        remaining = harness.query(_env.db_path,
                                  "SELECT COUNT(*) FROM members WHERE id=?", (mid,))
        self.assertEqual(remaining, [(1,)])

    def test_delete_member_with_access_history_refused(self):
        mid = self._add_member("30004")
        resp = self.client.post("/access/record", data={
            "member_id": mid, "csrf_token": self._csrf("/access")},
            follow_redirects=True)
        self.assertEqual(resp.status_code, 200)
        resp = self._delete_member(mid)
        body = resp.get_data(as_text=True)
        self.assertIn("ประวัติการเข้าออก", body)
        remaining = harness.query(_env.db_path,
                                  "SELECT COUNT(*) FROM members WHERE id=?", (mid,))
        self.assertEqual(remaining, [(1,)])

    def test_delete_member_without_history_succeeds(self):
        mid = self._add_member("30005")
        qr_path = harness.query(
            _env.db_path, "SELECT qrcode_path FROM members WHERE id=?",
            (mid,))[0][0]
        self.assertTrue(qr_path)
        resp = self._delete_member(mid)
        self.assertIn("ลบสมาชิกสำเร็จ", resp.get_data(as_text=True))
        remaining = harness.query(_env.db_path,
                                  "SELECT COUNT(*) FROM members WHERE id=?", (mid,))
        self.assertEqual(remaining, [(0,)])
        # the QR image is removed only after the row really disappears
        self.assertFalse(os.path.exists(qr_path))
