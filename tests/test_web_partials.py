# tests/test_web_partials.py — list-table API + header CSRF + analytics API
#
# The vendored frontend stack (jQuery/DataTables/Alpine/SweetAlert2) leans
# on three server-side behaviors:
#   * /books /members /history render a full-page skeleton table; all rows
#     come from the /api/{books,members,history} DataTables feeds;
#   * POSTs may carry the session CSRF token in the X-CSRF-Token header
#     (fetch-based table rows/bulk bar) instead of a hidden form field;
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


class TestWebTableApi(unittest.TestCase):
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

    def _dt(self, path, q="", order_col=0, order_dir="desc", **extra):
        resp = self.client.get(path, query_string={
            "draw": 1, "start": 0, "length": 50,
            "search[value]": q,
            "order[0][column]": order_col, "order[0][dir]": order_dir,
            **extra})
        self.assertEqual(resp.status_code, 200)
        return resp.get_json()

    def _post(self, path, data):
        data = dict(data)
        data["csrf_token"] = self._csrf(path)
        resp = self.client.post(path, data=data)
        return resp

    def _add_book(self, code, title):
        resp = self._post("/books/add", {"code": code, "title": title})
        self.assertEqual(resp.status_code, 302)
        return harness.query(_env.db_path,
                             "SELECT id, status FROM books WHERE code=?",
                             (code,))[0]

    def _add_member(self, number, name_suffix=""):
        resp = self._post("/members/add", {
            "name": f"นักเรียนทดสอบ {number}{name_suffix}",
            "grade": "ม.4/1", "number": number})
        self.assertEqual(resp.status_code, 302)
        return harness.query(_env.db_path,
                             "SELECT id FROM members WHERE number=?",
                             (number,))[0][0]

    # ---------------- page skeletons (no rows, no page chrome) ----------------

    def test_books_full_page_has_vendor_wiring(self):
        resp = self.client.get("/books")
        self.assertEqual(resp.status_code, 200)
        body = resp.get_data(as_text=True)
        self.assertIn("<html", body)
        self.assertIn("vendor/sweetalert2.min.js", body)
        self.assertIn("vendor/jquery.min.js", body)
        self.assertIn("vendor/datatables.min.js", body)
        self.assertIn('id="booksTable"', body)
        self.assertIn("vendor/alpine.min.js", body)
        self.assertIn('name="csrf_token"', body)
        self.assertNotIn("htmx", body)

    def test_page_counts_render_in_header(self):
        self._add_book("PW-000", "หัวกระดาษหน้า")
        body = self.client.get("/books").get_data(as_text=True)
        self.assertIn("หนังสือทั้งหมด (", body)
        body = self.client.get("/members").get_data(as_text=True)
        self.assertIn("รายชื่อสมาชิก (", body)
        body = self.client.get("/history").get_data(as_text=True)
        self.assertIn("ประวัติการยืม-คืน (", body)

    # ---------------- /api books ----------------

    def test_api_books_q_filters_rows(self):
        self._add_book("PW-200", "บทนำส่วนนำ Python ฉบับเข้ม")
        self._add_book("PW-201", "อัญมณีในแม่น้ำ")
        payload = self._dt("/api/books", q="ฉบับเข้ม")
        titles = [row["title"] for row in payload["data"]]
        self.assertIn("บทนำส่วนนำ Python ฉบับเข้ม", titles)
        self.assertNotIn("อัญมณีในแม่น้ำ", titles)
        self.assertEqual(payload["recordsFiltered"], 1)

    def test_api_books_status_filter(self):
        book1 = self._add_book("PW-300", "กรองสถานะ ว่าง")
        book2 = self._add_book("PW-301", "กรองสถานะ หยุด")
        resp = self._post(f"/books/{book2[0]}/retire", {})
        self.assertEqual(resp.status_code, 302)
        codes = lambda p: [r["code"] for r in p["data"]]  # noqa: E731
        ok = self._dt("/api/books", q="กรองสถานะ", status="ว่าง")
        self.assertEqual(codes(ok), ["PW-300"])
        gone = self._dt("/api/books", q="กรองสถานะ", status="หยุดให้บริการ")
        self.assertEqual(codes(gone), ["PW-301"])

    # ---------------- /api members + history ----------------

    def test_api_members_q_filter(self):
        self._add_member("91001")
        self._add_member("91002", "นาย")
        payload = self._dt("/api/members", q="91002")
        numbers = [row["number"] for row in payload["data"]]
        self.assertIn("91002", numbers)
        self.assertNotIn("91001", numbers)
        self.assertEqual(payload["recordsFiltered"], 1)

    def test_api_history_borrow_overdue_shape(self):
        member_id = self._add_member("91003")
        book = self._add_book("PW-400", "บทนำส่วนนำ Python เล่มนำโชค")
        resp = self._post("/borrow", {
            "member_id": member_id, "book_id": book[0],
            "due_date": "2099-01-01"})
        self.assertEqual(resp.status_code, 302)
        payload = self._dt("/api/history", q="เล่มนำโชค")
        rows = [r for r in payload["data"] if r["code"] == "PW-400"]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["returned"], 0)   # กำลังยืม
        self.assertEqual(rows[0]["name"], "นักเรียนทดสอบ 91003")
        miss = self._dt("/api/history", q="ไม่มีอะไรตรงเลย")
        self.assertEqual(miss["recordsFiltered"], 0)
        self.assertEqual(miss["data"], [])

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

    def test_post_with_bad_csrf_header_ajax_gets_403(self):
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
        self._post("/borrow", {
            "member_id": member_id, "book_id": book[0],
            "due_date": "2099-01-01"})
        payload = self.client.get("/api/analytics").get_json()
        self.assertGreaterEqual(payload["borrow_last_14_days"][-1]["count"], 1)
        # the just-borrowed book left ว่าง, so the donut sees it as borrowed
        self.assertGreaterEqual(payload["book_status"]["borrowed"], 1)


if __name__ == "__main__":
    unittest.main()
