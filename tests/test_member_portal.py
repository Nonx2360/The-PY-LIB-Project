# tests/test_member_portal.py — member login, first activation, status page,
# live QR, admin pin management, and the admin/member permission wall.
#
# Follows the test_web_scan.py harness pattern (TempEnv + two clients:
# an admin session and a session the portal uses exclusively).
import os
import random
import re
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

_env = harness.TempEnv()

import web.app as web_app_mod  # noqa: E402
from core import security  # noqa: E402
from db.repositories.base_repo import BaseRepository  # noqa: E402
from db.repositories.member_repo import MemberRepository  # noqa: E402
from config import settings as settings_mod  # noqa: E402

CSRF_RE = re.compile(r'name="csrf_token" value="([^"]+)"')
ORIG_SETTINGS_FILE = settings_mod.SETTINGS_FILE


class TestMemberPortal(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        _env.apply()
        _env.patch()
        web_app_mod.DB_PATH = _env.db_path
        settings_mod.SETTINGS_FILE = os.path.join(_env.dir, "settings.json")
        cls._orig_cwd = os.getcwd()
        os.chdir(_env.dir)
        harness.stage_static_assets(_env.dir)
        cls.base = BaseRepository()
        cls.repos = MemberRepository()
        cls.app = web_app_mod.create_app()
        cls.admin_client = cls.app.test_client()
        cls._admin_login(cls.admin_client)

    @classmethod
    def tearDownClass(cls):
        settings_mod.SETTINGS_FILE = ORIG_SETTINGS_FILE
        os.chdir(cls._orig_cwd)
        _env.restore()
        harness.cleanup(_env.dir)

    # ---------------- helpers ----------------

    @classmethod
    def _admin_login(cls, client):
        pwd_path = os.path.join(_env.dir, "INITIAL_ADMIN_PASSWORD.txt")
        with open(pwd_path, encoding="utf-8") as fh:
            content = fh.read()
        resp = client.get("/login")
        token = CSRF_RE.search(resp.get_data(as_text=True)).group(1)
        for line in content.splitlines():
            if "admin:" in line:
                password = line.split("admin:", 1)[1].split(" (", 1)[0].strip()
                resp = client.post("/login", data={
                    "username": "admin", "password": password,
                    "csrf_token": token,
                })
                if resp.status_code != 302:
                    raise AssertionError(f"admin login failed: {resp.status_code}")
                break

    @staticmethod
    def _csrf(client, path):
        resp = client.get(path)
        match = CSRF_RE.search(resp.get_data(as_text=True))
        if not match:                       # maybe redirected to the portal
            resp = client.get("/member/login")
            match = CSRF_RE.search(resp.get_data(as_text=True))
        return match.group(1)

    @classmethod
    def _make_member(cls, prefix="สมชาย ทดสอบ ", grade="ม.4/1", pin=None):
        """A fresh member per call; optional pre-set portal PIN."""
        suffix = f"{random.randint(10000, 99999)}"
        mid = cls.repos.add(prefix + suffix, grade, suffix,
                            "2026-01-01", "2027-01-01", "qr", None)
        if pin is not None:
            cls.repos.set_pin(mid, pin)
        name, number = prefix + suffix, suffix
        return mid, name, grade, number

    @classmethod
    def _activate_and_login(cls, client, grade, number, pin="654321",
                            via_card=True, payload=None, code=None):
        """Full first-activation flow in one go; raises on any refusal."""
        def _error_flash(txt):
            matched = re.findall(
                r'alert alert-(?:error|warning)[^>]*>\s*([^<]{5,200})', txt)
            return ("activation step 1 refused: " + matched[0]
                    if matched else "activation step 1 failed (no flash)")
        token = cls._csrf(client, "/member/login")
        if via_card:
            plaintext = payload or f"สมชาย ทดสอบ|{grade}|{number}|2026-01-01|2027-01-01"
            resp = client.post(
                "/member/activate/scan",
                data={"payload": security.encrypt_qr_data(plaintext),
                      "csrf_token": token}, follow_redirects=True)
            if "ขั้นตอนที่ 2" not in resp.get_data(as_text=True):
                raise AssertionError(_error_flash(resp.get_data(as_text=True)))
        else:
            resp = client.post(
                "/member/activate/code",
                data={"code": code, "csrf_token": token}, follow_redirects=True)
            if "ขั้นตอนที่ 2" not in resp.get_data(as_text=True):
                raise AssertionError(_error_flash(resp.get_data(as_text=True)))
        resp = client.post(
            "/member/activate/set-pin",
            data={"pin1": pin, "pin2": pin, "csrf_token": token})
        if resp.status_code != 302:
            raise AssertionError(f"set-pin failed: {resp.status_code}")
        return resp

    # ---------------- activation flows ----------------

    def test_activation_via_card_scan_then_pin(self):
        client = self.app.test_client()
        mid, name, grade, number = self._make_member()
        self._activate_and_login(client, grade, number)
        resp = client.get("/member")
        body = resp.get_data(as_text=True)
        self.assertEqual(resp.status_code, 200)
        self.assertIn(name, body)
        self.assertIn("QR ประจำตัว", body)
        # takeover guard state: the row now has a pin (stated, never printed)
        self.assertEqual(self.repos.pin_state(mid), "set")

    def test_activation_refused_when_pin_already_set(self):
        client = self.app.test_client()
        mid, _name, grade, number = self._make_member(pin="654321")
        token = self._csrf(client, "/member/login")
        resp = client.post(
            "/member/activate/scan",
            data={"payload": security.encrypt_qr_data(
                f"สมชาย ทดสอบ|{grade}|{number}|2026-01-01|2027-01-01"),
                  "csrf_token": token},
            follow_redirects=True)
        body = resp.get_data(as_text=True)
        self.assertIn("เปิดใช้งานหน้าของตนเองอยู่แล้ว", body)
        self.assertNotIn("ขั้นตอนที่ 2", body)
        self.assertEqual(self.repos.pin_state(mid), "set")

    def test_activation_rejects_ticket_payload(self):
        client = self.app.test_client()
        mid, _name, _grade, _number = self._make_member()
        token = self._csrf(client, "/member/login")
        resp = client.post(
            "/member/activate/scan",
            data={"payload": security.encrypt_member_ticket(mid, 12),
                  "csrf_token": token},
            follow_redirects=True)
        self.assertIn("ต้องใช้ QR จากบัตร", resp.get_data(as_text=True))

    def test_activation_via_admin_code(self):
        client = self.app.test_client()
        mid, _name, _grade, _number = self._make_member()
        # admin issues the one-time code (JSON row-action shape)
        resp = self.admin_client.post(
            f"/members/{mid}/activation-code",
            headers={"X-CSRF-Token": self._csrf(self.admin_client, "/members")})
        self.assertEqual(resp.status_code, 200)
        self.assertTrue(resp.get_json()["ok"])
        code = resp.get_json()["code"]
        self._activate_and_login(client, "ม.4/1", "0", via_card=False, code=code)
        resp = client.get("/member")
        self.assertEqual(resp.status_code, 200)
        # the code is burnt: log out, then a reuse attempt uses a fresh
        # session (members with a session get sent to the portal instead)
        client.post("/member/logout",
                    data={"csrf_token": self._csrf(client, "/member")})
        token = self._csrf(client, "/member/login")
        resp = client.post(
            "/member/activate/code",
            data={"code": code, "csrf_token": token}, follow_redirects=True)
        self.assertIn("ไม่ถูกต้องหรือหมดอายุ", resp.get_data(as_text=True))

    def test_activation_code_for_activated_member_refused(self):
        mid, _name, _grade, _number = self._make_member()
        members, token = self.repos, self._csrf(self.admin_client, "/members")
        members.set_pin(mid, "654321")
        resp = self.admin_client.post(
            f"/members/{mid}/activation-code",
            headers={"X-CSRF-Token": token})
        self.assertEqual(resp.get_json()["ok"], False)
        self.assertIn("อยู่แล้ว", resp.get_json()["message"])

    def test_activation_window_expires(self):
        client = self.app.test_client()
        _mid, _name, grade, number = self._make_member()
        token = self._csrf(client, "/member/login")
        client.post(
            "/member/activate/scan",
            data={"payload": security.encrypt_qr_data(
                f"สมชาย ทดสอบ|{grade}|{number}|2026-01-01|2027-01-01"),
                  "csrf_token": token}, follow_redirects=True)
        # force the 10-minute window into the past
        with client.session_transaction() as sess:
            sess["activation_at"] = 1.0
        resp = client.post(
            "/member/activate/set-pin",
            data={"pin1": "123456", "pin2": "123456", "csrf_token": token},
            follow_redirects=True)
        self.assertIn("หมดเวลาเปิดใช้งาน", resp.get_data(as_text=True))

    def test_pin_length_enforced(self):
        client = self.app.test_client()
        _mid, _name, grade, number = self._make_member()
        token = self._csrf(client, "/member/login")
        client.post(
            "/member/activate/scan",
            data={"payload": security.encrypt_qr_data(
                f"สมชาย ทดสอบ|{grade}|{number}|2026-01-01|2027-01-01"),
                  "csrf_token": token}, follow_redirects=True)
        resp = client.post(
            "/member/activate/set-pin",
            data={"pin1": "1234", "pin2": "1234", "csrf_token": token},
            follow_redirects=True)
        self.assertIn("6–64", resp.get_data(as_text=True))

    # ---------------- login / logout ----------------

    def test_login_logout_roundtrip(self):
        client = self.app.test_client()
        mid, _name, grade, number = self._make_member(pin="654321")
        token = self._csrf(client, "/member/login")
        resp = client.post("/member/login", data={
            "grade": grade, "number": number, "pin": "654321",
            "csrf_token": token})
        self.assertEqual(resp.status_code, 302)
        self.assertTrue(resp.headers["Location"].endswith("/member"))
        self.assertEqual(client.get("/member").status_code, 200)
        resp = client.post("/member/logout",
                           data={"csrf_token": self._csrf(client, "/member")})
        self.assertEqual(resp.status_code, 302)
        resp = client.get("/member", follow_redirects=False)
        self.assertEqual(resp.status_code, 302)
        self.assertTrue(resp.headers["Location"].endswith("/member/login"))

    def test_login_errors_are_generic(self):
        client = self.app.test_client()
        _mid, _name, grade, number = self._make_member(pin="654321")
        token = self._csrf(client, "/member/login")
        for wrong in ({"grade": grade, "number": number, "pin": "000000"},
                      {"grade": "ม.9/9", "number": "99999", "pin": "654321"}):
            resp = client.post("/member/login",
                               data={**wrong, "csrf_token": token})
            body = resp.get_data(as_text=True)
            self.assertIn("ชั้น เลขที่ หรือรหัส PIN ไม่ถูกต้อง", body)
            self.assertNotIn("ไม่พบสมาชิก", body)   # no user enumeration

    def test_login_throttling(self):
        client = self.app.test_client()
        _mid, _name, grade, number = self._make_member(pin="654321")
        token = self._csrf(client, "/member/login")
        bad = {"grade": grade, "number": number, "pin": "000000",
               "csrf_token": token}
        for _ in range(5):
            client.post("/member/login", data=bad)
        resp = client.post("/member/login", data=bad)
        self.assertEqual(resp.status_code, 429)

    # ---------------- permission wall ----------------

    def test_member_session_cannot_open_admin_pages(self):
        client = self.app.test_client()
        mid, _name, grade, number = self._make_member(pin="654321")
        client.post("/member/login", data={
            "grade": grade, "number": number, "pin": "654321",
            "csrf_token": self._csrf(client, "/member/login")})
        for path in ("/", "/members", "/books", "/settings", "/history",
                     "/api/members"):
            resp = client.get(path)
            self.assertEqual(resp.status_code, 302, path)
            self.assertTrue(resp.headers["Location"].endswith("/login"), path)
        # and cannot mutate: posting a member-add is bounced the same way
        resp = client.post("/members/add", data={
            "name": "x", "grade": "y", "number": "z",
            "csrf_token": self._csrf(client, "/member")})
        self.assertEqual(resp.status_code, 302)
        self.assertTrue(resp.headers["Location"].endswith("/login"))

    def test_admin_session_does_not_open_portal_urls(self):
        # the admin client has no member session yet -> member_required bounces
        resp = self.admin_client.get("/member")
        self.assertEqual(resp.status_code, 302)
        self.assertTrue(resp.headers["Location"].endswith("/member/login"))

    # ---------------- portal page + QR ----------------

    def test_portal_shows_only_own_data(self):
        client = self.app.test_client()
        mid_a, name_a, grade_a, number_a = self._make_member(pin="654321")
        mid_b, _n, grade_b, number_b = self._make_member(pin="654321",
                                                         prefix="อรุณ แสงดี ")
        bid_a = self.base.execute(
            "INSERT INTO books (code, title, status) VALUES (?, ?, 'ว่าง')",
            (f"PORTAL-{random.randint(100000,999999)}", "หนังสือของเอ"))
        bid_b = self.base.execute(
            "INSERT INTO books (code, title, status) VALUES (?, ?, 'ว่าง')",
            (f"PORTAL-{random.randint(100000,999999)}", "หนังสือของบี"))
        self.base.execute(
            "INSERT INTO borrow_log (member_id, book_id, borrow_date, "
            "return_due, returned) VALUES (?, ?, ?, ?, 0)",
            (mid_a, bid_a, "2026-11-01", "2026-12-01"))
        self.base.execute(
            "INSERT INTO borrow_log (member_id, book_id, borrow_date, "
            "return_due, returned) VALUES (?, ?, ?, ?, 0)",
            (mid_b, bid_b, "2026-10-01", "2026-10-02"))    # overdue for b
        client.post("/member/login", data={
            "grade": grade_a, "number": number_a, "pin": "654321",
            "csrf_token": self._csrf(client, "/member/login")})
        body = client.get("/member").get_data(as_text=True)
        self.assertIn(name_a, body)
        self.assertIn("หนังสือของเอ", body)
        self.assertNotIn("หนังสือของบี", body)
        self.assertNotIn("เกินกำหนด", body)      # a's loan is not overdue
        client.post("/member/logout",
                    data={"csrf_token": self._csrf(client, "/member")})
        client.post("/member/login", data={
            "grade": grade_b, "number": number_b, "pin": "654321",
            "csrf_token": self._csrf(client, "/member/login")})
        body = client.get("/member").get_data(as_text=True)
        self.assertIn("หนังสือของบี", body)
        self.assertIn("เกินกำหนด", body)
        self.assertNotIn("หนังสือของเอ", body)

    def test_qr_png_bound_and_rate_limited(self):
        client = self.app.test_client()
        _mid, _name, grade, number = self._make_member(pin="654321")
        client.post("/member/login", data={
            "grade": grade, "number": number, "pin": "654321",
            "csrf_token": self._csrf(client, "/member/login")})
        resp = client.get("/member/qr.png")
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.mimetype, "image/png")
        self.assertIn("no-store", resp.headers.get("Cache-Control", ""))
        self.assertTrue(resp.get_data()[:4] == b"\x89PNG")
        resp_next = client.get("/member/qr.png?next=1")
        self.assertEqual(resp_next.status_code, 200)
        # picture rate limit: 30/window per member session — push past it
        statuses = [client.get("/member/qr.png").status_code
                    for _ in range(40)]
        self.assertIn(429, statuses)
        code, count = None, 0
        for code in statuses:
            if code == 429:
                break
            count += 1
        self.assertGreaterEqual(count, 27)   # 2 earlier + ~28 fresh before hit

    # ---------------- admin pin management ----------------

    def test_admin_reset_pin_locks_member_out(self):
        mid, _name, grade, number = self._make_member(prefix="รีเซ็ต ทดสอบ",
                                                      pin="654321")
        token = self._csrf(self.admin_client, "/members")
        resp = self.admin_client.post(
            f"/members/{mid}/reset-pin", data={"csrf_token": token})
        self.assertEqual(resp.status_code, 302)
        # outstanding codes (if any) died with the reset
        self.assertEqual(
            self.base.fetchone(
                "SELECT COUNT(*) FROM member_activation_codes WHERE member_id=?",
                (mid,))[0], 0)
        client = self.app.test_client()
        resp = client.post("/member/login", data={
            "grade": grade, "number": number, "pin": "654321",
            "csrf_token": self._csrf(client, "/member/login")})
        # the old PIN is dead: the login re-renders (200) with the generic
        # error — never a redirect into the portal
        self.assertEqual(resp.status_code, 200)
        self.assertIn("ชั้น เลขที่ หรือรหัส PIN ไม่ถูกต้อง",
                      resp.get_data(as_text=True))
        portal = client.get("/member", follow_redirects=False)
        self.assertTrue(portal.headers["Location"].endswith("/member/login"))

    # ---------------- settings group ----------------

    def test_settings_member_group_only(self):
        try:
            token = self._csrf(self.admin_client, "/settings")
            resp = self.admin_client.post("/settings", data={
                "settings_group": "member_portal",
                "member_portal_enabled": "on",
                "member_qr_rotation_hours": "7",
                "member_qr_grace_steps": "0",
                "accept_static_card_qr": "on",
                "csrf_token": token})
            self.assertEqual(resp.status_code, 302)
            self.assertEqual(settings_mod.get("member_qr_rotation_hours"), 7)
            self.assertEqual(settings_mod.get("member_qr_grace_steps"), 0)
            # a general-group save must NOT wipe the member/portal group
            resp = self.admin_client.post("/settings", data={
                "settings_group": "general",
                "school_name": "TESTSCHOOL",
                "default_loan_days": "9",
                "max_books_per_member": "4",
                "member_expiry_days": "400",
                "school_logo_path": "assets/logos/school_logo.png",
                "csrf_token": token})
            self.assertEqual(resp.status_code, 302)
            self.assertEqual(settings_mod.get("member_qr_rotation_hours"), 7)
            self.assertEqual(settings_mod.get("school_name"), "TESTSCHOOL")
            # out-of-range values fall back to defaults
            resp = self.admin_client.post("/settings", data={
                "settings_group": "member_portal",
                "member_portal_enabled": "on",
                "member_qr_rotation_hours": "500",
                "member_qr_grace_steps": "9",
                "csrf_token": token})
            self.assertEqual(settings_mod.get("member_qr_rotation_hours"), 12)
            self.assertEqual(settings_mod.get("member_qr_grace_steps"), 1)
        finally:
            # restore the values other tests read from settings
            settings_mod.set("member_qr_rotation_hours", 12)
            settings_mod.set("member_qr_grace_steps", 1)
            settings_mod.set("accept_static_card_qr", True)
            settings_mod.set("member_portal_enabled", True)

    def test_portal_switch_off_404s_everything(self):
        try:
            settings_mod.set("member_portal_enabled", False)
            for path in ("/member/login", "/member", "/member/activate"):
                resp = self.admin_client.get(path)
                self.assertEqual(resp.status_code, 404, path)
        finally:
            settings_mod.set("member_portal_enabled", True)

    def test_repo_activate_with_pin_atomic(self):
        mid, _name, _grade, _number = self._make_member()
        self.repos.set_pin(mid, "654321")
        # second activation (any path) must refuse, not overwrite
        self.assertFalse(self.repos.activate_with_pin(mid, "999999"))
        self.assertTrue(self.repos.verify_member_pin(mid, "654321"))

    def test_repo_activate_with_pin_claims_code(self):
        mid, _name, _grade, _number = self._make_member()
        code = self.repos.issue_activation_code(mid)
        # a foreign member cannot win the code, and ends up with no pin
        other_mid, *_ = self._make_member()
        self.assertFalse(
            self.repos.activate_with_pin(other_mid, "999999", code=code))
        self.assertEqual(self.repos.pin_state(other_mid), "none")
        # the rightful member wins code + pin in one move
        self.assertTrue(self.repos.activate_with_pin(mid, "111111", code=code))
        self.assertEqual(self.repos.pin_state(mid), "set")
        self.assertIsNone(self.repos.find_activation_member(code))


if __name__ == "__main__":
    unittest.main()
