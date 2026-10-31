# tests/test_web_bulk.py — bulk Excel import + large-catalog pagination
#
# The library loads ~1000+ books at once. Two things must hold:
#   * an Excel import of 1000 rows lands in ONE transaction (no hang, no
#     partial catalog) and reports what it skipped instead of erroring;
#   * every list page paginates server-side (PAGE_SIZE rows + pager), so
#     the render stays cheap no matter how big the tables grow.
import io
import os
import re
import sys
import unittest

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

_env = harness.TempEnv()

# NOTE: env/binding mutations happen in setUpClass, NOT at import (see
# test_web_scan.py for the discovery-collision explanation).
import web.app as web_app_mod  # noqa: E402

CSRF_RE = re.compile(r'name="csrf_token" value="([^"]+)"')
HX = {"HX-Request": "true"}


def _xlsx(rows, columns=("รหัสหนังสือ", "ชื่อเรื่อง")) -> io.BytesIO:
    df = pd.DataFrame(rows, columns=list(columns))
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf


class TestWebBulk(unittest.TestCase):
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
                if resp.status_code != 302:
                    raise AssertionError(f"login failed: HTTP {resp.status_code}")
                break

    def _upload(self, buf, filename="books.xlsx", follow=True):
        resp = self.client.post(
            "/books/import",
            data={"file": (buf, filename),
                  "csrf_token": self._csrf("/books")},
            content_type="multipart/form-data")
        if follow and resp.status_code == 302:
            resp = self.client.get(resp.headers["Location"])
        return resp

    def _hx(self, path, qs=None):
        return self.client.get(path, headers=HX, query_string=qs or {})

    def _book_count(self):
        return len(harness.query(_env.db_path, "SELECT 1 FROM books"))

    # ---------------- repo bulk insert ----------------

    def test_add_many_single_transaction_skip_semantics(self):
        from db.repositories.book_repo import BookRepository
        repo = BookRepository()
        repo.add("BQ-000", "มีอยู่ก่อนแล้ว")
        inserted, skipped = repo.add_many([
            ("BQ-000", "รหัสซ้ำกับในตาราง"),     # skipped: UNIQUE code
            ("BQ-001", "หนังสือเล่มที่ 1"),
            ("BQ-001", "รหัสซ้ำในไฟล์เดียวกัน"),  # skipped: intra-file dup
            ("BQ-002", "หนังสือเล่มที่ 2"),
            ("BQ-003", ""),                        # skipped: blank title
            ("", "ไม่มีรหัส"),                     # skipped: blank code
        ])
        self.assertEqual(inserted, 2)
        self.assertEqual(skipped, 4)
        rows = harness.query(_env.db_path,
                             "SELECT code, title, status FROM books "
                             "WHERE code LIKE 'BQ-%' ORDER BY code")
        self.assertEqual([r[0] for r in rows], ["BQ-000", "BQ-001", "BQ-002"])
        self.assertEqual(rows[1][1], "หนังสือเล่มที่ 1")  # first occurrence wins
        from config.constants import BOOK_STATUS_AVAILABLE
        self.assertEqual(rows[1][2], BOOK_STATUS_AVAILABLE)

    # ---------------- Excel import flow ----------------

    def test_import_1000_books_reports_each_skip(self):
        from db.repositories.book_repo import BookRepository
        repo = BookRepository()
        repo.add_many([("IMP-000", "มีอยู่ก่อนแล้ว")])
        rows = [("IMP-000", "รหัสนี้ซ้ำ")]           # already stored
        rows += [(f"IMP-{i:04d}", f"หนังสือนำเข้า เล่มที่ {i}")
                 for i in range(1, 1001)]            # 1000 unique rows
        rows += [("IMP-0001", "รหัสซ้ำในไฟล์เดียวกัน")]  # intra-file dup

        resp = self._upload(_xlsx(rows))
        self.assertEqual(
            len(harness.query(_env.db_path,
                              "SELECT 1 FROM books WHERE code LIKE 'IMP-%'")),
            1001)                                    # 1000 new + 1 preexisting
        body = resp.get_data(as_text=True) if resp.status_code == 200 else "no-flash"
        self.assertIn("นำเข้าหนังสือสำเร็จ 1000 เล่ม", body)
        self.assertIn("รหัสซ้ำข้าม 2 รายการ", body)   # table dup + in-file dup

        # the books page must still render with 1000 new rows present
        resp = self.client.get("/books")
        self.assertEqual(resp.status_code, 200)
        self.assertIn("หนังสือทั้งหมด (", resp.get_data(as_text=True))

    def test_import_rejects_wrong_extension(self):
        before = self._book_count()
        resp = self._upload(io.BytesIO(b"not an excel"), filename="books.xls")
        self.assertIn("รองรับเฉพาะไฟล์ .xlsx", resp.get_data(as_text=True))
        self.assertEqual(self._book_count(), before)  # nothing touched

    def test_import_rejects_bad_column_layout(self):
        buf = _xlsx([("001", "หนังสือ")], columns=("code", "ชื่อเรื่อง"))
        resp = self._upload(buf)
        self.assertIn("รูปแบบไฟล์ Excel ไม่ถูกต้อง", resp.get_data(as_text=True))

    def test_import_over_5000_rows_rejected(self):
        rows = [(f"OVER-{i}", f"เล่มที่ {i}") for i in range(5001)]
        resp = self._upload(_xlsx(rows))
        self.assertIn("5,000 แถว", resp.get_data(as_text=True))

    def test_413_guard_tells_user_to_split_file(self):
        original = self.app.config["MAX_CONTENT_LENGTH"]
        self.app.config["MAX_CONTENT_LENGTH"] = 1024
        try:
            buf = io.BytesIO(b"x" * 4096)
            resp = self.client.post(
                "/books/import",
                data={"file": (buf, "t.xlsx"),
                      "csrf_token": self._csrf("/books")},
                content_type="multipart/form-data")
            self.assertEqual(resp.status_code, 302)
            body = self.client.get(resp.headers["Location"]).get_data(as_text=True)
            self.assertIn("แบ่งการนำเข้าเป็นหลายไฟล์", body)
        finally:
            self.app.config["MAX_CONTENT_LENGTH"] = original

    def test_template_download(self):
        resp = self.client.get("/books/template")
        self.assertEqual(resp.status_code, 200)
        self.assertIn("spreadsheetml", resp.headers["Content-Type"])
        self.assertGreater(len(resp.get_data()), 3000)  # a real .xlsx grows here

    # ---------------- pagination ----------------

    def test_books_pagination_slices_and_clamps(self):
        from db.repositories.book_repo import BookRepository
        from web.app import PAGE_SIZE
        BookRepository().add_many(
            [(f"PAG-{i:04d}", f"หนังสือแบ่งหน้า {i}") for i in range(1, 121)])
        total = self._book_count()                   # includes sibling seeds
        pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        # page 1 holds the 50 newest rows, which are this block: PAG-0120..0071
        body = self._hx("/books").get_data(as_text=True)
        self.assertIn(f"รวม {total} เล่ม", body)
        self.assertIn("PAG-0120", body)
        self.assertIn("PAG-0071", body)
        self.assertNotIn("PAG-0070", body)
        # page 3 exists and shading marks it active
        body = self._hx("/books", {"page": 3}).get_data(as_text=True)
        self.assertIn(f"หน้า 3 จาก {pages}", body)
        self.assertIn("page-active", body)
        self.assertNotIn("PAG-0069", body)           # page-2 territory
        # out-of-range request clamps to the last page (no crash, no empty lie)
        body = self._hx("/books", {"page": 999}).get_data(as_text=True)
        self.assertIn(f"หน้า {pages} จาก {pages}", body)
        self.assertNotIn("PAG-0120", body)           # newest row is on page 1

    def test_books_pagination_keeps_search_context(self):
        # seed its own rows (OR IGNORE keeps re-seeding idempotent; unittest
        # discovery runs methods alphabetically, so ordering may vary)
        from db.repositories.book_repo import BookRepository
        BookRepository().add_many(
            [(f"PAGS-{i:04d}", f"ค้นหาแบ่งหน้า {i}") for i in range(1, 121)])
        # q matches exactly the 120 seeded rows -> multi-page with q carried on
        body = self._hx("/books", {"q": "ค้นหาแบ่งหน้า"}).get_data(as_text=True)
        self.assertIn("รวม 120 เล่ม", body)
        self.assertIn("PAGS-0120", body)
        self.assertNotIn("PAGS-0069", body)             # page 2 territory
        self.assertIn("q=%", body)                      # q survives in pager links

    def test_members_pagination(self):
        from db.repositories.member_repo import MemberRepository
        repo = MemberRepository()
        for i in range(1, 61):
            repo.add(f"สมาชิกแบ่งหน้า {i:02d}", f"ม.5/{(i % 9) + 1}",
                     f"95{i:03d}", "2026-01-01", "2027-01-01", "qr", None)
        total = len(harness.query(_env.db_path, "SELECT 1 FROM members"))
        seeded = len(harness.query(
            _env.db_path, "SELECT 1 FROM members WHERE name LIKE 'สมาชิกแบ่งหน้า %'"))
        self.assertEqual(seeded, 60)
        # page 1 = the 50 newest ids, page 2 = the remainder
        body = self._hx("/members").get_data(as_text=True)
        self.assertIn(f"รวม {total} คน", body)
        self.assertIn("สมาชิกแบ่งหน้า 60", body)
        body = self._hx("/members", {"page": 2}).get_data(as_text=True)
        self.assertIn("หน้า 2 จาก 2", body)
        self.assertIn("สมาชิกแบ่งหน้า 10 ", body)
        self.assertIn("สมาชิกแบ่งหน้า 01 ", body)
        self.assertNotIn("สมาชิกแบ่งหน้า 60", body)

    def test_history_pagination(self):
        from db.repositories.book_repo import BookRepository
        from db.repositories.member_repo import MemberRepository
        from db.repositories.base_repo import BaseRepository
        member_id = MemberRepository().add(
            "สมาชิกประวัติ", "ม.6/1", "96001", "2026-01-01", "2027-01-01", "qr", None)
        BookRepository().add_many(
            [(f"HST-{i:04d}", f"หนังสือประวัติ {i}") for i in range(1, 11)])
        book_ids = [r[0] for r in harness.query(
            _env.db_path, "SELECT id FROM books WHERE code LIKE 'HST-%'")]

        def _borrows(conn):
            conn.executemany(
                "INSERT INTO borrow_log (member_id, book_id, borrow_date, "
                "return_due, returned) VALUES (?, ?, '2026-01-05', "
                "'2026-01-20', 1)",
                [(member_id, book_ids[i % 10], ) for i in range(120)])
        BaseRepository().transact(_borrows)

        body = self._hx("/history").get_data(as_text=True)
        self.assertIn("รวม 120 รายการ", body)
        self.assertIn("หน้า 1 จาก 3", body)
        body = self._hx("/history", {"page": 3}).get_data(as_text=True)
        self.assertIn("หน้า 3 จาก 3", body)
        self.assertIn("สมาชิกประวัติ", body)


if __name__ == "__main__":
    unittest.main()
