# tests/test_web_pages_smoke.py — every authenticated page renders HTTP 200
#
# Cheap render gauntlet: any Jinja/route error (new blocks, template renames,
# vendor wiring) trips a non-200 here before real users see it.
import os
import re
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

_env = harness.TempEnv()

import web.app as web_app_mod  # noqa: E402

CSRF_RE = re.compile(r'name="csrf_token" value="([^"]+)"')


class TestPageSmoke(unittest.TestCase):
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

    def test_dashboard_renders_charts_section(self):
        resp = self.client.get("/")
        self.assertEqual(resp.status_code, 200)
        body = resp.get_data(as_text=True)
        self.assertIn("chartBorrow", body)
        self.assertIn("chartAccess", body)
        self.assertIn("chartStatus", body)
        self.assertIn("vendor/chart.umd.min.js", body)

    def test_all_pages_render_200(self):
        for path in ("/", "/books", "/members", "/borrow", "/return",
                     "/history", "/access", "/settings"):
            with self.subTest(path=path):
                resp = self.client.get(path)
                self.assertEqual(resp.status_code, 200)

    def test_borrow_page_wires_alpine_and_flatpickr(self):
        resp = self.client.get("/borrow")
        self.assertEqual(resp.status_code, 200)
        body = resp.get_data(as_text=True)
        self.assertIn("x-data=", body)
        self.assertIn("data-flatpickr", body)
        self.assertIn("vendor/flatpickr.min.js", body)
        self.assertIn("vendor/flatpickr-th.js", body)

    def test_unauthenticated_redirects_to_login(self):
        self.client.get("/logout")
        for path in ("/", "/books", "/api/analytics"):
            with self.subTest(path=path):
                resp = self.client.get(path)
                self.assertEqual(resp.status_code, 302)
                self.assertIn("/login", resp.headers["Location"])


if __name__ == "__main__":
    unittest.main()
