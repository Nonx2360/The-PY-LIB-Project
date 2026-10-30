# tests/test_web_scan.py — /scan/resolve endpoint (member badge scanning)
import os
import re
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

_env = harness.TempEnv()

# NOTE: environment/binding mutations happen in setUpClass, NOT at import:
# unittest discovery imports every module BEFORE running any test class, so
# module-level TempEnv patching would let the last-imported module silently
# repoint shared globals (DB_PATH, PYLIB_ENV_FILE) for every other web class.
import web.app as web_app_mod  # noqa: E402
from core import security  # noqa: E402

CSRF_RE = re.compile(r'name="csrf_token" value="([^"]+)"')

MEMBER_PAYLOAD = "อรทัย ใจดี|ม.4/1|12345|2026-01-01|2027-01-01"


def _SECRETS_SALT():
    from config.env_store import ensure_secrets
    return ensure_secrets()["SECRET_SALT"]


class TestScanResolve(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from config import settings as settings_mod
        _env.apply()
        _env.patch()
        web_app_mod.DB_PATH = _env.db_path  # login uses a module-level binding
        cls._orig_settings_file = settings_mod.SETTINGS_FILE
        settings_mod.SETTINGS_FILE = os.path.join(_env.dir, "settings.json")
        # QR/card artifacts are written relative to CWD — chdir into the temp
        # env so member-add tests never touch the real assets/ folders
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
    def _csrf(cls, path="/login"):
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
        token = cls._csrf()
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

    def _scan(self, payload, token=None):
        token = token if token is not None else self._csrf("/borrow")
        return self.client.post("/scan/resolve",
                                data={"payload": payload, "csrf_token": token})

    # ---------------- tests ----------------

    def setUp(self):
        # any test that called /logout leaves the session cleared; ensure a
        # fresh login for every test so ordering can never leak session state
        self._login()

    def test_valid_member_badge_resolves(self):
        resp = self.client.post("/members/add", data={
            "name": "อรทัย ใจดี", "grade": "ม.4/1", "number": "12345",
            "csrf_token": self._csrf("/members"),
        })
        self.assertEqual(resp.status_code, 302)
        payload = security.encrypt_qr_data(MEMBER_PAYLOAD)  # what the camera sees
        resp = self._scan(payload)
        self.assertEqual(resp.status_code, 200)
        body = resp.get_json()
        self.assertTrue(body["ok"])
        self.assertEqual(body["member"]["number"], "12345")
        self.assertEqual(body["member"]["grade"], "ม.4/1")
        self.assertFalse(body["member"]["expired"])

    def test_tampered_payload_rejected(self):
        payload = security.encrypt_qr_data(MEMBER_PAYLOAD)
        bad = payload[: -4] + ("AAAA" if payload[-4:] != "AAAA" else "BBBB")
        resp = self._scan(bad)
        self.assertEqual(resp.status_code, 400)          # test matrix: no plaintext fallback
        self.assertEqual(resp.get_json()["error"], "invalid")

    def test_wrong_key_token_rejected(self):
        # a badge sealed under a DIFFERENT passphrase than the current .env
        key = security.derive_qr_key("definitely-not-the-real-passphrase",
                                     _SECRETS_SALT())
        foreign = security.encrypt_with_key(key, MEMBER_PAYLOAD)
        resp = self._scan(foreign)
        self.assertEqual(resp.status_code, 400)
        self.assertEqual(resp.get_json()["error"], "invalid")

    def test_not_a_member_badge_rejected(self):
        payload = security.encrypt_qr_data("some other library system output")
        resp = self._scan(payload)
        self.assertEqual(resp.status_code, 400)
        self.assertEqual(resp.get_json()["error"], "invalid")

    def test_unknown_member_number(self):
        payload = security.encrypt_qr_data("คนอื่น|ม.6/1|99999|2026-01-01|2027-01-01")
        resp = self._scan(payload)
        self.assertEqual(resp.status_code, 404)
        self.assertEqual(resp.get_json()["error"], "not_found")

    def test_unauthenticated_scan_rejected(self):
        self.client.get("/logout")
        resp = self.client.post("/scan/resolve", data={"payload": "zzz"})
        self.assertEqual(resp.status_code, 302)          # redirected to login
        self.assertTrue(resp.headers["Location"].endswith("/login"))

    def test_scan_without_csrf_rejected(self):
        payload = security.encrypt_qr_data(MEMBER_PAYLOAD)
        resp = self.client.post("/scan/resolve", data={"payload": payload})  # no csrf_token
        # CSRF before_request flashes + redirects (does not resolve anything)
        self.assertNotIn(b'"ok":true', resp.get_data())
        self.assertNotIn(b'"ok": true', resp.get_data())

    def test_return_filter_by_member(self):
        # add member + book, borrow, then /return?member_id=N lists only them
        token = self._csrf("/members")
        self.client.post("/members/add", data={
            "name": "อรทัย ใจดี", "grade": "ม.4/1", "number": "12345",
            "csrf_token": token})
        self.client.post("/members/add", data={
            "name": "ประเสิรฐ ดีงาม", "grade": "ม.5/2", "number": "54321",
            "csrf_token": token})
        self.client.post("/books/add", data={"code": "W-001", "title": "เล่มหนึ่ง",
                                             "csrf_token": self._csrf("/books")})
        self.client.post("/books/add", data={"code": "W-002", "title": "เล่มสอง",
                                             "csrf_token": self._csrf("/books")})
        member_rows = harness.query(_env.db_path, "SELECT id, number FROM members")
        by_number = {num: mid for mid, num in member_rows}
        book_rows = [r[0] for r in harness.query(_env.db_path, "SELECT id FROM books")]
        for mid, bid in ((by_number["12345"], book_rows[0]),
                         (by_number["54321"], book_rows[1])):
            resp = self.client.post("/borrow", data={
                "member_id": mid, "book_id": bid, "due_date": "2026-12-31",
                "csrf_token": self._csrf("/borrow")})
            self.assertEqual(resp.status_code, 302)

        row_a = harness.query(
            _env.db_path, "SELECT id FROM borrow_log ORDER BY id LIMIT 1")[0][0]
        member_a = harness.query(
            _env.db_path,
            "SELECT bl.member_id FROM borrow_log bl WHERE bl.id=?",
            (row_a,))[0][0]

        resp = self.client.get(f"/return?member_id={member_a}")
        body = resp.get_data(as_text=True)
        self.assertIn("ล้างตัวกรอง", body)
        self.assertEqual(body.count('class="btn btn-primary btn-sm"'), 1)

    def test_return_without_filter_lists_all(self):
        resp = self.client.get("/return")
        body = resp.get_data(as_text=True)
        self.assertNotIn("ล้างตัวกรอง", body)
