# tests/test_web_bulk.py — bulk Excel import + server-side DataTables feeds
#
# The library loads ~1000+ books at once. Three things must hold:
#   * an Excel import of 1000 rows lands in ONE transaction (no hang, no
#     partial catalog) and reports what it skipped instead of erroring;
#   * the list pages are fed by /api/{books,members,history} in the
#     DataTables server-side protocol, bounded per request so a 10,000-row
#     catalog costs the same as a 10-row one (only one page is ever SELECTed);
#   * the bulk endpoints (/books/bulk, /members/bulk) apply many actions in
#     one atomic shot and skip rows that must not be touched instead of
#     failing the batch.
import io
import json
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

    def _dt(self, path, draw=1, start=0, length=50, q="", status=None,
            order_col=0, order_dir="desc"):
        """Minimal replay of what the vendored DataTables sends."""
        qs = {"draw": draw, "start": start, "length": length,
              "search[value]": q,
              "order[0][column]": order_col,
              "order[0][dir]": order_dir}
        if status:
            qs["status"] = status
        resp = self.client.get(path, query_string=qs)
        self.assertEqual(resp.status_code, 200)
        self.assertTrue(resp.content_type.startswith("application/json"))
        payload = resp.get_json()
        self.assertIn("recordsTotal", payload)
        self.assertEqual(payload["draw"], draw)
        return payload

    def _bulk(self, path, payload, token=None):
        return self.client.post(
            path,
            data=json.dumps(payload),
            content_type="application/json",
            headers={"X-CSRF-Token": token or self._csrf("/books")})

    def _book_count(self):
        return len(harness.query(_env.db_path, "SELECT 1 FROM books"))

    def _codes(self, payload, key="code"):
        return [row[key] for row in payload["data"]]

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

    # ---------------- DataTables server-side API ----------------

    def test_books_api_page1_slices_newest_first(self):
        from db.repositories.book_repo import BookRepository
        BookRepository().add_many(
            [(f"PAG-{i:04d}", f"หนังสือแบ่งหน้า {i}") for i in range(1, 121)])
        total = self._book_count()                   # includes sibling seeds
        payload = self._dt("/api/books")
        self.assertEqual(payload["recordsTotal"], total)
        self.assertEqual(payload["recordsFiltered"], total)
        self.assertEqual(len(payload["data"]), 50)   # page 1 of the newest
        codes = self._codes(payload)
        self.assertEqual(codes[0], "PAG-0120")       # newest first
        self.assertEqual(codes[49], "PAG-0071")
        self.assertNotIn("PAG-0070", codes)          # page-2 territory
        self.assertIn("status", payload["data"][0])  # row shape for the UI

    def test_books_api_page2_offset_and_clamped_start(self):
        # start=100 walks past page 1 into the 51st row; overshooting start
        # (a bulk delete can empty the browser's page) clamps back onto the
        # last full page instead of answering with an empty sheet
        payload = self._dt("/api/books", start=100)
        codes = self._codes(payload)
        self.assertEqual(len(codes), 50)             # full 50-row page
        self.assertEqual(codes[0], "PAG-0020")       # 100 rows deep
        payload = self._dt("/api/books", start=99999)
        total = payload["recordsTotal"]
        self.assertEqual(len(self._codes(payload)), (total - 1) % 50 + 1)
        first_of_last = self._codes(
            self._dt("/api/books", start=((total - 1) // 50) * 50))[0]
        self.assertEqual(self._codes(payload)[0], first_of_last)

    def test_books_api_search_filters(self):
        from db.repositories.book_repo import BookRepository
        BookRepository().add_many(
            [(f"PAGS-{i:04d}", f"ค้นหาแบ่งหน้า {i}") for i in range(1, 121)])
        payload = self._dt("/api/books", q="ค้นหาแบ่งหน้า")
        self.assertEqual(payload["recordsFiltered"], 120)
        codes = self._codes(payload)
        self.assertEqual(codes[0], "PAGS-0120")
        self.assertTrue(all(c.startswith("PAGS-") for c in codes))

    def test_books_api_order_asc_and_desc(self):
        from db.repositories.book_repo import BookRepository
        BookRepository().add_many(
            [(f"ORD-{i:04d}", f"เรียงหน้าที่ {i}") for i in range(1, 13)])
        up = self._dt("/api/books", q="เรียงหน้าที่", order_col=1,
                      order_dir="asc")
        self.assertEqual(self._codes(up)[0], "ORD-0001")
        down = self._dt("/api/books", q="เรียงหน้าที่", order_col=1,
                        order_dir="desc")
        self.assertEqual(self._codes(down)[0], "ORD-0012")

    def test_books_api_length_clamped_to_whitelist(self):
        # only 10/25/50/100 are offered by the UI; anything else -> 50
        from db.repositories.book_repo import BookRepository
        BookRepository().add_many(
            [(f"CLP-{i:04d}", f"จำกัดหน้าละ {i}") for i in range(1, 61)])
        payload = self._dt("/api/books", q="จำกัดหน้าละ", length=999)
        self.assertEqual(len(payload["data"]), 50)

    def test_books_api_order_column_out_of_whitelist_defaults_newest(self):
        payload = self._dt("/api/books", q="จำกัดหน้าละ", order_col=77)
        self.assertEqual(self._codes(payload)[0], "CLP-0060")

    def test_books_api_status_filter(self):
        from db.repositories.book_repo import BookRepository
        BookRepository().add("SRS-001", "กรองสถานะ ว่าง")
        srs2 = BookRepository().add("SRS-002", "กรองสถานะ หยุด")
        resp = self.client.post(f"/books/{srs2}/retire",
                                data={"csrf_token": self._csrf("/books")})
        self.assertEqual(resp.status_code, 302)
        ok = self._dt("/api/books", q="กรองสถานะ", status="ว่าง")
        self.assertEqual(self._codes(ok), ["SRS-001"])
        gone = self._dt("/api/books", q="กรองสถานะ", status="หยุดให้บริการ")
        self.assertEqual(self._codes(gone), ["SRS-002"])

    def test_members_api_pagination(self):
        from db.repositories.member_repo import MemberRepository
        repo = MemberRepository()
        for i in range(1, 61):
            repo.add(f"สมาชิกแบ่งหน้า {i:02d}", f"ม.5/{(i % 9) + 1}",
                     f"95{i:03d}", "2026-01-01", "2027-01-01", "qr", None)
        payload = self._dt("/api/members")
        total_members = len(harness.query(_env.db_path, "SELECT 1 FROM members"))
        self.assertEqual(payload["recordsTotal"], total_members)
        self.assertEqual(payload["data"][0]["name"], "สมาชิกแบ่งหน้า 60")
        page2 = self._dt("/api/members", start=50)
        names = [row["name"] for row in page2["data"]]
        self.assertIn("สมาชิกแบ่งหน้า 10", names)
        self.assertIn("สมาชิกแบ่งหน้า 01", names)
        self.assertNotIn("สมาชิกแบ่งหน้า 60", names)

    def test_history_api_pagination(self):
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

        total = len(harness.query(_env.db_path, "SELECT 1 FROM borrow_log"))
        payload = self._dt("/api/history")
        self.assertEqual(payload["recordsTotal"], total)
        self.assertEqual(payload["recordsFiltered"], total)
        self.assertEqual(len(payload["data"]), 50)
        self.assertEqual(payload["data"][0]["title"], "หนังสือประวัติ 10")
        page3 = self._dt("/api/history", start=100)
        self.assertEqual(page3["data"][0]["title"], "หนังสือประวัติ 10")
        self.assertEqual(page3["data"][0]["returned"], 1)
        self.assertEqual(len(page3["data"]), total - 100)

    def test_history_api_search(self):
        payload = self._dt("/api/history", q="สมาชิกประวัติ")
        self.assertEqual(payload["recordsFiltered"], 120)
        miss = self._dt("/api/history", q="ไม่มีอะไรตรงเลย")
        self.assertEqual(miss["recordsFiltered"], 0)
        self.assertEqual(miss["data"], [])

    # ---------------- bulk endpoints ----------------

    def test_books_bulk_csrf_header(self):
        token = self._csrf("/books")
        ok = self._bulk("/books/bulk",
                        {"action": "retire", "ids": []}, token=token)
        self.assertEqual(ok.status_code, 200)     # the fetch header path
        bad = self.client.post(
            "/books/bulk", data=json.dumps({"action": "delete", "ids": [1]}),
            content_type="application/json",
            headers={**HX, "X-CSRF-Token": token + "tampered"})
        self.assertEqual(bad.status_code, 403)    # nothing deleted

    def test_books_bulk_delete_semantics(self):
        from db.repositories.book_repo import BookRepository
        from db.repositories.member_repo import MemberRepository
        from db.repositories.base_repo import BaseRepository
        plain = BookRepository().add("BKD-PLAIN", "ลบกลุ่ม เล่มธรรมดา")
        active = BookRepository().add("BKD-ACTIVE", "ลบกลุ่ม ยืมอยู่")
        purged = BookRepository().add("BKD-PURGED", "ลบกลุ่ม หยุดให้บริการ")
        member_id = MemberRepository().add(
            "สมาชิกลบกลุ่ม", "ม.4/3", "97002",
            "2026-01-01", "2027-01-01", "qr", None)
        resp = self.client.post("/borrow", data={
            "member_id": member_id, "book_id": active,
            "due_date": "2099-01-01", "csrf_token": self._csrf("/borrow")})
        self.assertEqual(resp.status_code, 302)

        def _past_history(conn):
            conn.execute(
                "INSERT INTO borrow_log (member_id, book_id, borrow_date, "
                "return_due, returned) VALUES (?, ?, '2025-01-01', "
                "'2025-01-15', 1)", (member_id, purged))
        BaseRepository().transact(_past_history)
        resp = self.client.post(f"/books/{purged}/retire",
                                data={"csrf_token": self._csrf("/books")})
        self.assertEqual(resp.status_code, 302)

        resp = self._bulk("/books/bulk", {"action": "delete",
                                          "ids": [plain, active, purged]})
        self.assertEqual(resp.status_code, 200)
        out = resp.get_json()
        self.assertEqual(out["status"], "ok")
        self.assertEqual(out["deleted"], 2)       # plain + retired-purge
        self.assertEqual(out["skipped"], 1)       # the actively-borrowed copy
        self.assertEqual(
            len(harness.query(_env.db_path,
                              "SELECT 1 FROM books WHERE code='BKD-PLAIN'")), 0)
        self.assertEqual(
            len(harness.query(_env.db_path,
                              "SELECT 1 FROM books WHERE code='BKD-ACTIVE'")), 1)
        # the skipped copy keeps its open borrow_log row untouched
        self.assertEqual(
            len(harness.query(
                _env.db_path,
                "SELECT 1 FROM borrow_log WHERE book_id=? AND returned=0",
                (active,))), 1)
        # the retired copy is purged WITH its borrow_log rows
        self.assertEqual(
            len(harness.query(
                _env.db_path,
                "SELECT 1 FROM books WHERE code='BKD-PURGED'")), 0)
        self.assertEqual(
            len(harness.query(
                _env.db_path,
                "SELECT 1 FROM borrow_log WHERE book_id=?", (purged,))), 0)

    def test_books_bulk_id_cap_is_400(self):
        resp = self._bulk("/books/bulk",
                          {"action": "retire",
                           "ids": [str(n) for n in range(1, 202)]})
        self.assertEqual(resp.status_code, 400)
        self.assertIn("เลือกได้ครั้งละสูงสุด 200 รายการ",
                      resp.get_json()["message"])

    def test_books_bulk_retire_and_restore(self):
        from db.repositories.book_repo import BookRepository
        from db.repositories.member_repo import MemberRepository
        # BR-001 ว่าง | BR-002 ยืมอยู่ | BR-003 หยุดให้บริการแล้ว
        br1 = BookRepository().add("BR-001", "หนังสือหยุดกลุ่ม A")
        br2 = BookRepository().add("BR-002", "หนังสือหยุดกลุ่ม B")
        br3 = BookRepository().add("BR-003", "หนังสือหยุดกลุ่ม C")
        member_id = MemberRepository().add(
            "สมาชิกหยุดกลุ่ม", "ม.4/2", "97001",
            "2026-01-01", "2027-01-01", "qr", None)
        resp = self.client.post("/borrow", data={
            "member_id": member_id, "book_id": br2,
            "due_date": "2099-01-01", "csrf_token": self._csrf("/borrow")})
        self.assertEqual(resp.status_code, 302)
        resp = self.client.post(f"/books/{br3}/retire",
                                data={"csrf_token": self._csrf("/books")})
        self.assertEqual(resp.status_code, 302)

        resp = self._bulk("/books/bulk",
                          {"action": "retire", "ids": [br1, br2, br3]})
        self.assertEqual(resp.status_code, 200)
        out = resp.get_json()
        self.assertEqual(out["status"], "ok")
        self.assertEqual(out["changed"], 1)       # only BR-001 was ว่าง
        self.assertEqual(out["skipped"], 2)
        rows = {r[0]: r[1] for r in harness.query(
            _env.db_path,
            "SELECT id, status FROM books WHERE id IN (?, ?, ?)",
            (br1, br2, br3))}
        self.assertEqual(rows[br1], "หยุดให้บริการ")
        self.assertEqual(rows[br2], "ยืมแล้ว")     # untouched
        self.assertEqual(rows[br3], "หยุดให้บริการ")

        resp = self._bulk("/books/bulk",
                          {"action": "restore", "ids": [br1, br2, br3]})
        self.assertEqual(resp.status_code, 200)
        out = resp.get_json()
        self.assertEqual(out["changed"], 2)       # BR-001 + BR-003 were retired
        self.assertEqual(out["skipped"], 1)       # BR-002 is borrowed, untouched
        rows = {r[0]: r[1] for r in harness.query(
            _env.db_path,
            "SELECT id, status FROM books WHERE id IN (?, ?, ?)",
            (br1, br2, br3))}
        self.assertEqual(rows[br1], "ว่าง")
        self.assertEqual(rows[br3], "ว่าง")

    def test_books_bulk_unknown_action_is_400(self):
        resp = self._bulk("/books/bulk", {"action": "แก้สถานะลอยๆ", "ids": [1]})
        self.assertEqual(resp.status_code, 400)
        self.assertIn("retire/restore/delete",
                      resp.get_json()["message"])
        resp = self._bulk("/members/bulk", {"action": "retire", "ids": [1]})
        self.assertEqual(resp.status_code, 400)
        self.assertIn("'delete'", resp.get_json()["message"])

    def test_members_bulk_delete(self):
        from db.repositories.member_repo import MemberRepository
        from db.repositories.book_repo import BookRepository
        keep = MemberRepository().add(
            "สมาชิกลบกลุ่ม (มีประวัติ)", "ม.4/4", "97003",
            "2026-01-01", "2027-01-01", "qr", None)
        drop = MemberRepository().add(
            "สมาชิกลบกลุ่ม (ใหม่)", "ม.4/4", "97004",
            "2026-01-01", "2027-01-01", "qr", None)
        book_id = BookRepository().add("BMD-001", "หนังสือลบกลุ่ม")
        resp = self.client.post("/borrow", data={
            "member_id": keep, "book_id": book_id,
            "due_date": "2099-01-01", "csrf_token": self._csrf("/borrow")})
        self.assertEqual(resp.status_code, 302)
        resp = self._bulk("/members/bulk",
                          {"action": "delete", "ids": [keep, drop]})
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.get_json(),
                         {"status": "ok", "deleted": 1, "skipped": 1})
        self.assertEqual(len(harness.query(
            _env.db_path, "SELECT 1 FROM members WHERE id=?", (keep,))), 1)
        self.assertEqual(len(harness.query(
            _env.db_path, "SELECT 1 FROM members WHERE id=?", (drop,))), 0)


if __name__ == "__main__":
    unittest.main()
