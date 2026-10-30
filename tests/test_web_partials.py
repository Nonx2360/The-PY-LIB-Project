# tests/test_web_partials.py — HTMX partials + header CSRF + analytics API
#
# The vendored frontend stack (HTMX/Alpine/SweetAlert2) leans on three
# server-side behaviors:
#   * /books /members /history honor ?q= (and /books ?status=) and render
#     bare table partials for HX-Request:"true" — no page chrome;
#   * POSTs may carry the session CSRF token in the X-CSRF-Token header
#     (htmx.body hx-headers wiring) instead of a hidden form field;
#   * /api/analytics serves zero-filled dashboard series.
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
HX = {"HX-Request": "true"}
_HX = {**HX, "X-CSRF-Token": ""}


class TestWebPartials(unittest.TestCase):
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
        _HX["X-CSRF-Token"] = cls._csrf("/books")

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

    def _add_book(self, code, title):
        resp = self.client.post("/books/add", data={
            "code": code, "title": title,
            "csrf_token": self._csrf("/books")})
        self.assertEqual(resp.status_code, 302)
        return harness.query(_env.db_path,
                             "SELECT id, status FROM books WHERE code=?",
                             (code,))[0]

    def _add_member(self, number, name_suffix=""):
        resp = self.client.post("/members/add", data={
            "name": f"นักเรียนทดสอบ {number}{name_suffix}",
            "grade": "ม.4/1", "number": number,
            "csrf_token": self._csrf("/members")})
        self.assertEqual(resp.status_code, 302)
        return harness.query(_env.db_path,
                             "SELECT id FROM members WHERE number=?",
                             (number,))[0][0]

    def _browse(self, path, qs=None):
        return self.client.get(path, headers=HX, query_string=qs or {})

    # ---------------- partials ----------------

    def test_books_partial_is_fragment_not_page(self):
        self._add_book("PW-100", "บทนำ Python")
        resp = self._browse("/books", {"q": "Python"})
        body = resp.get_data(as_text=True)
        self.assertNotEqual(resp.status_code, 400)
        self.assertNotIn("<html", body, "HX request must render the bare partial")
        self.assertIn('id="booksTableWrap"', body)
        self.assertIn("บทนำ Python", body)

    def test_books_q_filters_rows(self):
        self._add_book("PW-200", "บทนำ Python")
        self._add_book("PW-201", "อัญมณีในแม่น้ำ")
        resp = self._browse("/books", {"q": "Python"})
        self.assertEqual(resp.status_code, 200)
        body = resp.get_data(as_text=True)
        self.assertNotIn("<html", body)
        self.assertIn("บทนำ Python", body)
        self.assertNotIn("อัญมณีในแม่น้ำ", body)

    def test_books_status_filter(self):
        self._add_book("PW-300", "บทนำ Python")
        book2 = self._add_book("PW-301", "อัญมณีในแม่น้ำ")
        # retire one through the sanctioned web flow, then filter by status
        resp = self.client.post(f"/books/{book2[0]}/retire",
                                data={"csrf_token": self._csrf("/books")})
        self.assertEqual(resp.status_code, 302)
        body = self._browse("/books", {"status": "ว่าง"}).get_data(as_text=True)
        self.assertIn("PW-300", body)
        self.assertNotIn("PW-301", body)
        body = self._browse("/books", {"status": "หยุดให้บริการ"}).get_data(as_text=True)
        self.assertIn("PW-301", body)
        self.assertNotIn("PW-300", body)

    def test_books_full_page_has_vendor_wiring(self):
        resp = self.client.get("/books")
        self.assertEqual(resp.status_code, 200)
        body = resp.get_data(as_text=True)
        self.assertIn("<html", body)
        self.assertIn("vendor/sweetalert2.min.js", body)
        self.assertIn("vendor/htmx.min.js", body)
        self.assertIn("vendor/alpine.min.js", body)
        self.assertIn("hx-target", body)
        self.assertIn("X-CSRF-Token", body)

    def test_members_q_filter(self):
        self._add_member("91001")
        self._add_member("91002", "นาย")
        body = self._browse("/members", {"q": "91002"}).get_data(as_text=True)
        self.assertNotIn("<html", body)
        self.assertIn("91002", body)
        self.assertNotIn("91001", body)

    def test_history_q_filter(self):
        member_id = self._add_member("91003")
        book = self._add_book("PW-400", "บทนำ Python")
        due = "2099-01-01"
        resp = self.client.post("/borrow", data={
            "member_id": member_id, "book_id": book[0], "due_date": due,
            "csrf_token": self._csrf("/borrow")})
        self.assertEqual(resp.status_code, 302)
        body = self._browse("/history", {"q": "บทนำ Python"}).get_data(as_text=True)
        self.assertNotIn("<html", body)
        self.assertIn("บทนำ Python", body)
        self.assertIn("กำลังยืม", body)
        body = self._browse("/history", {"q": "หนังสือที่ไม่มีตรงกันเลย"}).get_data(as_text=True)
        self.assertIn("ไม่พบข้อมูลที่ตรงกับ", body)

    # ---------------- header CSRF ----------------

    def test_post_with_csrf_header_accepted(self):
        token = self._csrf("/books")
        resp = self.client.post("/books/add",
                                data={"code": "PW-HDR", "title": "บทนำ Python"},
                                headers={"X-CSRF-Token": token})
        self.assertEqual(resp.status_code, 302)
        row = harness.query(_env.db_path,
                            "SELECT id FROM books WHERE code='PW-HDR'")
        self.assertEqual(len(row), 1)

    def test_post_with_bad_csrf_header_rejected(self):
        token = self._csrf("/books")
        resp = self.client.post("/books/add",
                                data={"code": "PW-BAD", "title": "x"},
                                headers={"X-CSRF-Token": token + "tampered"})
        self.assertEqual(resp.status_code, 302)   # classic POST -> redirect
        row = harness.query(_env.db_path,
                            "SELECT id FROM books WHERE code='PW-BAD'")
        self.assertEqual(len(row), 0)

    def test_post_with_bad_csrf_header_hx_gets_403(self):
        token = self._csrf("/books")
        resp = self.client.post("/books/add",
                                data={"code": "PW-BAD2", "title": "x"},
                                headers={**HX, "X-CSRF-Token": token + "tampered"})
        self.assertEqual(resp.status_code, 403)

    # ---------------- analytics ----------------

    def test_analytics_shape(self):
        resp = self.client.get("/api/analytics")
        self.assertEqual(resp.status_code, 200)
        payload = resp.get_json()
        self.assertEqual(len(payload["borrow_last_14_days"]), 14)
        self.assertEqual(len(payload["access_last_7_days"]), 7)
        self.assertIn("book_status", payload)
        self.assertEqual(
            {"available", "borrowed", "retired", "other"},
            set(payload["book_status"].keys()))

    def test_analytics_counts_today_borrow(self):
        member_id = self._add_member("91004")
        book = self._add_book("PW-500", "บทนำ Python")
        self.client.post("/borrow", data={
            "member_id": member_id, "book_id": book[0],
            "due_date": "2099-01-01", "csrf_token": self._csrf("/borrow")})
        payload = self.client.get("/api/analytics").get_json()
        self.assertGreaterEqual(payload["borrow_last_14_days"][-1]["count"], 1)
        # the just-borrowed book left ว่าง, so the donut sees it as borrowed
        self.assertGreaterEqual(payload["book_status"]["borrowed"], 1)


if __name__ == "__main__":
    unittest.main()
