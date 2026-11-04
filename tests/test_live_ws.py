# tests/test_live_ws.py — the realtime WebSocket channel: auth walls,
# cross-process tick detection, and member snapshot push (integration
# over a real Werkzeug server + simple_websocket.Client).
#
# Follows the test_member_portal.py harness pattern (TempEnv + admin
# login). Member sessions are fabricated with session_transaction()
# because /member/ws only trusts the session key it shares with pages.
import json
import os
import random
import re
import sys
import threading
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

_env = harness.TempEnv()

import web.app as web_app_mod  # noqa: E402
from db.repositories.base_repo import BaseRepository  # noqa: E402
from db.repositories.member_repo import MemberRepository  # noqa: E402
from config import settings as settings_mod  # noqa: E402
from web import live as live_mod  # noqa: E402
import simple_websocket  # noqa: E402
from simple_websocket.errors import ConnectionClosed  # noqa: E402
from werkzeug.serving import make_server  # noqa: E402

CSRF_RE = re.compile(r'name="csrf_token" value="([^"]+)"')
ORIG_SETTINGS_FILE = settings_mod.SETTINGS_FILE


class TestLiveWebSocket(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        _env.apply()
        _env.patch()
        web_app_mod.DB_PATH = _env.db_path
        settings_mod.SETTINGS_FILE = os.path.join(_env.dir, "settings.json")
        settings_mod.load()
        cls._orig_cwd = os.getcwd()
        os.chdir(_env.dir)
        harness.stage_static_assets(_env.dir)
        cls.base = BaseRepository()
        cls.repos = MemberRepository()
        cls.app = web_app_mod.create_app()

        # real server so the socket upgrade actually happens
        cls.server = make_server("127.0.0.1", 0, cls.app, threaded=True)
        cls.port = cls.server.server_address[1]
        cls.server_thread = threading.Thread(
            target=cls.server.serve_forever, daemon=True)
        cls.server_thread.start()

        cls.admin_client = cls.app.test_client()
        cls._admin_login(cls.admin_client)

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        settings_mod.SETTINGS_FILE = ORIG_SETTINGS_FILE
        settings_mod.load()
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
    def _cookie_header(client):
        """The signed session cookie of a test client, as an
        (name, value) header pair; None when the client has none."""
        cookie = client.get_cookie("pylib_session")
        if cookie is None:
            return None
        return ("Cookie", "pylib_session=" + cookie.value)

    @classmethod
    def _connect(cls, path, client):
        """Open a real socket to the test server; returns the Client."""
        url = f"ws://127.0.0.1:{cls.port}{path}"
        headers = [("Origin", f"http://127.0.0.1:{cls.port}")]
        if cls._cookie_header(client):
            headers.append(cls._cookie_header(client))
        return simple_websocket.Client(url, headers=headers)

    @staticmethod
    def _safe_close(sock):
        """close() on an already-closed client raises ConnectionClosed."""
        try:
            sock.close()
        except Exception:
            pass

    # ---------------- tests ----------------

    def test_ticker_diffs_markers(self):
        """tick_once() reports a topic once per change, then goes quiet."""
        marker = {"value": (1, 0)}

        def read_a():
            return marker["value"]

        def read_b():
            return (9,)

        ticker = live_mod.LiveTicker(live_mod.LiveHub(),
                                     topics={"a": read_a, "b": read_b})
        # the constructor seeds the baseline: first beat is quiet
        self.assertEqual(ticker.tick_once(), [])
        marker["value"] = (2, 0)
        self.assertEqual(ticker.tick_once(), ["a"])
        self.assertEqual(ticker.tick_once(), [])
        self.assertEqual(ticker.tick_once(), [])

    def test_admin_ws_rejects_anonymous_and_member_sessions(self):
        client = self.app.test_client()
        self.repos.add("ประชาชน ไลฟ์", "ม.4/1",
                       str(random.randint(10000, 99999)),
                       "2026-01-01", "2027-01-01", "qr", None)
        # anonymous: the handler drops the socket before any frame
        sock = self._connect("/admin/ws", client)
        with self.assertRaises(ConnectionClosed):
            sock.receive(timeout=3)
        self._safe_close(sock)

    def test_member_session_cannot_use_admin_socket(self):
        client = self.app.test_client()
        mid = self.repos.add("สมาชิก ไลฟ์", "ม.5/2",
                             str(random.randint(10000, 99999)),
                             "2026-01-01", "2027-01-01", "qr", None)
        with client.session_transaction() as sess:
            sess["member_id"] = mid
        sock = self._connect("/admin/ws", client)
        with self.assertRaises(ConnectionClosed):
            sock.receive(timeout=3)
        self._safe_close(sock)

    def test_admin_ws_pushes_changed_topics(self):
        client = self.app.test_client()
        self._admin_login(client)
        sock = self._connect("/admin/ws", client)
        try:
            self.repos.add("เพิ่งสมัคร ไลฟ์", "ม.6/1",
                           str(random.randint(10000, 99999)),
                           "2026-01-01", "2027-01-01", "qr", None)
            msg = json.loads(sock.receive(timeout=8))
            self.assertEqual(msg["type"], "topics")
            self.assertIn("members", msg["topics"])
        finally:
            self._safe_close(sock)

    def test_member_ws_snapshot_and_live_update(self):
        client = self.app.test_client()
        mid = self.repos.add("ผู้รับ ไลฟ์", "ม.4/2",
                             str(random.randint(10000, 99999)),
                             "2026-01-01", "2027-01-01", "qr", None)
        with client.session_transaction() as sess:
            sess["member_id"] = mid

        # portal off → the socket must die before any snapshot flies
        settings_mod.set("member_portal_enabled", False)
        try:
            sock = self._connect("/member/ws", client)
            with self.assertRaises(ConnectionClosed):
                sock.receive(timeout=3)
            self._safe_close(sock)
        finally:
            settings_mod.set("member_portal_enabled", True)

        # portal on → first paint snapshot, then a live push after a desk
        # (desktop-app style, direct DB write) add of one loan
        sock = self._connect("/member/ws", client)
        snap = json.loads(sock.receive(timeout=8))
        self.assertEqual(snap["type"], "snapshot")
        self.assertEqual(snap["counts"]["loans"], 0)
        self.assertEqual(snap["counts"]["overdue"], 0)
        self.assertEqual(snap["loans"], [])

        bid = self.base.execute(
            "INSERT INTO books (code, title, status) VALUES (?, ?, 'ว่าง')",
            (f"LIVE-{random.randint(100000, 999999)}", "หนังสือไลฟ์"))
        self.base.execute(
            "INSERT INTO borrow_log (member_id, book_id, borrow_date, "
            "return_due, returned) VALUES (?, ?, ?, ?, 0)",
            (mid, bid, "2026-11-01", "2026-12-01"))
        snap = json.loads(sock.receive(timeout=8))
        self.assertEqual(snap["type"], "snapshot")
        self.assertEqual(snap["counts"]["loans"], 1)
        self.assertEqual(snap["loans"][0][0], "หนังสือไลฟ์")
        self._safe_close(sock)


if __name__ == "__main__":
    unittest.main()
