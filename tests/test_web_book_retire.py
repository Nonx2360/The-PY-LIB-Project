# tests/test_web_book_retire.py — book retire/restore lifecycle
#
# Retire (สถานะ หยุดให้บริการ) is the sanctioned way to take a history-
# bearing book out of circulation: deletion stays FK-blocked forever, but
# the book must leave every borrow flow until restored.
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


class TestWebBookRetire(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from config import settings as settings_mod
        _env.apply()
        _env.patch()
        web_app_mod.DB_PATH = _env.db_path
        cls._orig_settings_file = settings_mod.SETTINGS_FILE
        settings_mod.SETTINGS_FILE = os.path.join(_env.dir, "settings.json")
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

    def _add_book(self, code):
        resp = self.client.post("/books/add", data={
            "code": code, "title": "หนังสือทดสอบ " + code,
            "csrf_token": self._csrf("/books")})
        self.assertEqual(resp.status_code, 302)
        return harness.query(_env.db_path,
                             "SELECT id, status FROM books WHERE code=?",
                             (code,))[0]

    def _add_member(self, number):
        resp = self.client.post("/members/add", data={
            "name": f"นักเรียนทดสอบ {number}", "grade": "ม.4/1", "number": number,
            "csrf_token": self._csrf("/members")})
        self.assertEqual(resp.status_code, 302)
        return harness.query(_env.db_path,
                             "SELECT id FROM members WHERE number=?",
                             (number,))[0][0]

    def _retire(self, bid):
        return self.client.post(f"/books/{bid}/retire",
                                data={"csrf_token": self._csrf("/books")},
                                follow_redirects=True)

    def _restore(self, bid):
        return self.client.post(f"/books/{bid}/restore",
                                data={"csrf_token": self._csrf("/books")},
                                follow_redirects=True)

    def _status(self, bid):
        return harness.query(_env.db_path,
                             "SELECT status FROM books WHERE id=?", (bid,))[0][0]

    # ---------------- tests ----------------

    def test_retire_removes_from_borrow_flow_then_restore(self):
        bid, _ = self._add_book("R-001")
        comma = self._add_member("40001")

        # available book is listed in the borrow page dropdown
        page = self.client.get("/borrow").get_data(as_text=True)
        self.assertIn("R-001", page)

        resp = self._retire(bid)
        self.assertIn("หยุดให้บริการแล้ว", resp.get_data(as_text=True))
        self.assertEqual(self._status(bid), "หยุดให้บริการ")

        # retired book is excluded from the borrow dropdown...
        page = self.client.get("/borrow").get_data(as_text=True)
        self.assertNotIn("R-001", page)
        # ...and a direct POST slip past the dropdown is refused
        resp = self.client.post("/borrow", data={
            "member_id": comma, "book_id": bid, "due_date": "2027-12-31",
            "csrf_token": self._csrf("/borrow")}, follow_redirects=True)
        self.assertIn("หยุดให้บริการ", resp.get_data(as_text=True))
        self.assertNotIn("ยืมหนังสือสำเร็จ", resp.get_data(as_text=True))
        self.assertEqual(harness.query(
            _env.db_path, "SELECT COUNT(*) FROM borrow_log WHERE book_id=?",
            (bid,)), [(0,)])

        resp = self._restore(bid)
        self.assertIn("กลับมาให้บริการ", resp.get_data(as_text=True))
        self.assertEqual(self._status(bid), "ว่าง")
        page = self.client.get("/borrow").get_data(as_text=True)
        self.assertIn("R-001", page)

    def test_retire_borrowed_book_refused_until_returned(self):
        bid, _ = self._add_book("R-002")
        mid = self._add_member("40002")
        resp = self.client.post("/borrow", data={
            "member_id": mid, "book_id": bid, "due_date": "2027-12-31",
            "csrf_token": self._csrf("/borrow")})
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(self._status(bid), "ยืมแล้ว")

        resp = self._retire(bid)
        self.assertIn("กรุณารอคืนก่อน", resp.get_data(as_text=True))
        self.assertEqual(self._status(bid), "ยืมแล้ว")   # unchanged

    def test_restore_non_retired_book_refused(self):
        bid, _ = self._add_book("R-003")   # still ว่าง, never retired
        resp = self._restore(bid)
        self.assertIn("กำลังให้บริการอยู่แล้ว", resp.get_data(as_text=True))
        self.assertEqual(self._status(bid), "ว่าง")

    def test_double_retire_refused(self):
        bid, _ = self._add_book("R-004")
        self._retire(bid)
        resp = self._retire(bid)
        self.assertIn("หยุดให้บริการอยู่แล้ว", resp.get_data(as_text=True))
        self.assertEqual(self._status(bid), "หยุดให้บริการ")

    def test_retired_book_keeps_deletion_guard(self):
        # give the book borrow history, retire+restore it, then delete:
        # the history guard must still refuse (retire never hides history)
        bid, _ = self._add_book("R-005")
        mid = self._add_member("40003")
        self.client.post("/borrow", data={
            "member_id": mid, "book_id": bid, "due_date": "2027-12-31",
            "csrf_token": self._csrf("/borrow")})
        record = harness.query(_env.db_path,
                               "SELECT id FROM borrow_log ORDER BY id DESC LIMIT 1")[0][0]
        self.client.post("/return", data={
            "record_id": record, "csrf_token": self._csrf("/return")})
        self._retire(bid)
        self._restore(bid)
        resp = self.client.post(f"/books/{bid}/delete",
                                data={"csrf_token": self._csrf("/books")},
                                follow_redirects=True)
        self.assertIn("มีประวัติการยืม", resp.get_data(as_text=True))
        remaining = harness.query(_env.db_path,
                                  "SELECT COUNT(*) FROM books WHERE id=?", (bid,))
        self.assertEqual(remaining, [(1,)])
