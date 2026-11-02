# tests/test_rotating_tickets.py — v3 member ticket crypto + time-window math
#
# Every identity payload the system touches: sealed v3 tickets (rotating),
# static v2 card envelopes and legacy v1 badges — plus the window policy
# (epoch math + grace) and services.member_identity's dispatch rules.
import os
import sys
import unittest
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

_env = harness.TempEnv()

from cryptography.hazmat.primitives.ciphers.aead import AESGCM  # noqa: E402
from core import security  # noqa: E402
from services import member_identity  # noqa: E402

from config import settings as settings_mod  # noqa: E402
from db.repositories.member_repo import MemberRepository  # noqa: E402


class _TempDBCase(unittest.TestCase):
    """Shared: temp env + DB + one member row (created in setUpClass)."""

    member_id = None

    @classmethod
    def setUpClass(cls):
        _env.apply()
        _env.patch()
        settings_mod.SETTINGS_FILE = os.path.join(_env.dir, "settings.json")
        import db.schema as schema_mod
        schema_mod.init_db()
        repo = MemberRepository()
        cls.member_id = repo.add("สมชาย ใจดี", "ม.4/1", "15",
                                 "2026-01-01", "2027-01-01", "qr-x", None)

    @classmethod
    def tearDownClass(cls):
        settings_mod.SETTINGS_FILE = _orig_settings_file[0]
        _env.restore()
        harness.cleanup(_env.dir)


_orig_settings_file = [None]


def _patch_settings_file():
    _orig_settings_file[0] = settings_mod.SETTINGS_FILE


_patch_settings_file()


class TestTicketCrypto(_TempDBCase):
    """encrypt/decrypt roundtrips, tamper rejection, strict JSON."""

    def test_roundtrip_same_member_and_epoch(self):
        epoch = security.ticket_epoch(12)
        tok = security.encrypt_member_ticket(self.member_id, 12)
        self.assertTrue(tok.startswith("v3:"))
        mid, ep = security.decrypt_member_ticket(tok)
        self.assertEqual(mid, self.member_id)
        self.assertEqual(ep, epoch)

    def test_explicit_epoch_roundtrip(self):
        tok = security.encrypt_member_ticket(self.member_id, 12, epoch=100)
        self.assertEqual(security.decrypt_member_ticket(tok), (self.member_id, 100))

    def test_payload_is_short_enough_for_scanning(self):
        # desk cameras need version-1-sized QRs to stay reliable
        tok = security.encrypt_member_ticket(self.member_id, 12, epoch=41516)
        self.assertLessEqual(len(tok), 90)

    def test_wrong_key_rejected(self):
        from config.env_store import ensure_secrets
        key = security.derive_qr_key("another-passphrase",
                                     ensure_secrets()["SECRET_SALT"])
        tok = security.encrypt_member_ticket(self.member_id, 12)
        with self.assertRaises(ValueError):
            security.decrypt_member_ticket(tok, key=key)

    def test_tamper_rejected(self):
        tok = security.encrypt_member_ticket(self.member_id, 12)
        raw_half = len(tok) // 2
        bad = tok[:raw_half] + ("A" if tok[raw_half] != "A" else "B") + tok[raw_half + 1:]
        with self.assertRaises(ValueError):
            security.decrypt_member_ticket(bad)

    def test_foreign_aad_rejected(self):
        # sealed with the right key but a different AAD (protocol confusion)
        import json as _json
        plain = _json.dumps({"m": self.member_id, "e": 1}).encode()
        nonce = os.urandom(security.GCM_NONCE_LEN)
        sealed = AESGCM(security._AES_KEY).encrypt(nonce, plain, b"other-aad")
        fake = security.TICKET_PREFIX + security.base64.urlsafe_b64encode(
            nonce + sealed).decode("ascii")
        with self.assertRaises(ValueError):
            security.decrypt_member_ticket(fake)

    def test_ticks_touching_base64_rejected(self):
        with self.assertRaises(ValueError):
            security.decrypt_member_ticket("v3:!!!not-base64!!!")
        with self.assertRaises(ValueError):
            security.decrypt_member_ticket("v3:AAAA")     # truncated
        with self.assertRaises(ValueError):
            security.decrypt_member_ticket("v2:AAAA")     # wrong version


class TestTicketJsonStrictness(_TempDBCase):
    """parse_ticket_json: exact {m, e}, ints, sane ranges."""

    def test_valid_body(self):
        self.assertEqual(security.parse_ticket_json('{"e": 5, "m": 7}'), (7, 5))

    def test_extra_key_rejected(self):
        with self.assertRaises(ValueError):
            security.parse_ticket_json('{"m": 1, "e": 2, "x": 3}')

    def test_string_value_rejected(self):
        with self.assertRaises(ValueError):
            security.parse_ticket_json('{"m": "1", "e": 2}')

    def test_bool_value_rejected(self):
        with self.assertRaises(ValueError):
            security.parse_ticket_json('{"m": true, "e": 2}')

    def test_zero_member_rejected(self):
        with self.assertRaises(ValueError):
            security.parse_ticket_json('{"m": 0, "e": 2}')

    def test_negative_epoch_rejected(self):
        with self.assertRaises(ValueError):
            security.parse_ticket_json('{"m": 1, "e": -1}')

    def test_non_dict_rejected(self):
        with self.assertRaises(ValueError):
            security.parse_ticket_json('[1, 2]')
        with self.assertRaises(ValueError):
            security.parse_ticket_json('not json at all')


class TestEpochWindow(_TempDBCase):
    """Ticket epoch math and the accepted-window policy."""

    def test_epoch_matches_floor_of_time_divided_by_window(self):
        dt = datetime(2026, 5, 17, 14, 30, 0)
        expected = int(dt.timestamp() // (12 * 3600))
        self.assertEqual(security.ticket_epoch(12, dt), expected)

    def test_seconds_left_runs_down(self):
        # epoch boundaries are Unix-epoch multiples, NOT local midnight —
        # build a boundary clock relative to the window's own math
        base = datetime(2026, 5, 17, 0, 0, 0)
        rem = int(base.timestamp()) % 43200
        flip = base + timedelta(seconds=43200 - rem)   # exact next boundary
        self.assertEqual(security.ticket_epoch_seconds_left(12, flip), 43200)
        self.assertEqual(
            security.ticket_epoch_seconds_left(12, flip + timedelta(seconds=30)),
            43200 - 30)

    def test_epoch_flips_at_the_computed_boundary(self):
        base = datetime(2026, 5, 17, 0, 0, 0)
        rem = int(base.timestamp()) % 43200
        flip = base + timedelta(seconds=43200 - rem)
        just_before = flip + timedelta(seconds=-1)
        self.assertEqual(
            security.ticket_epoch(12, flip),
            security.ticket_epoch(12, just_before) + 1)

    def test_hours_are_clamped_to_positive(self):
        # rotation_hours is admin-managed; garbage must not divide by zero
        self.assertEqual(security.ticket_epoch(0), security.ticket_epoch(1))
        self.assertEqual(security.ticket_epoch(-5), security.ticket_epoch(1))

    def test_grace_zero_only_accepts_current(self):
        cur = security.ticket_epoch(12)
        self.assertTrue(security.ticket_window_ok(cur, 12, grace=0))
        self.assertFalse(security.ticket_window_ok(cur - 1, 12, grace=0))
        self.assertFalse(security.ticket_window_ok(cur + 1, 12, grace=0))

    def test_grace_one_accepts_neighbours(self):
        cur = security.ticket_epoch(12)
        self.assertTrue(security.ticket_window_ok(cur - 1, 12, grace=1))
        self.assertTrue(security.ticket_window_ok(cur + 1, 12, grace=1))
        self.assertFalse(security.ticket_window_ok(cur - 2, 12, grace=1))
        self.assertFalse(security.ticket_window_ok(cur + 2, 12, grace=1))

    def test_grace_two_range(self):
        cur = security.ticket_epoch(12)
        self.assertTrue(security.ticket_window_ok(cur - 2, 12, grace=2))
        self.assertTrue(security.ticket_window_ok(cur + 2, 12, grace=2))
        self.assertFalse(security.ticket_window_ok(cur - 3, 12, grace=2))


class TestIdentityFlows(_TempDBCase):
    """services.member_identity dispatch + policy, end to end."""

    def _hours(self):
        return member_identity.rotation_hours()

    def test_settings_clamped(self):
        try:
            settings_mod.set("member_qr_rotation_hours", 500)
            self.assertEqual(member_identity.rotation_hours(), 72)
            settings_mod.set("member_qr_rotation_hours", 0)
            self.assertEqual(member_identity.rotation_hours(), 1)
        finally:
            settings_mod.set("member_qr_rotation_hours", 12)

    def test_grace_clamped(self):
        try:
            settings_mod.set("member_qr_grace_steps", 9)
            self.assertEqual(member_identity.grace_steps(), 2)
            settings_mod.set("member_qr_grace_steps", -1)
            self.assertEqual(member_identity.grace_steps(), 0)
        finally:
            settings_mod.set("member_qr_grace_steps", 1)

    def test_ticket_envelope_resolves(self):
        tok = security.encrypt_member_ticket(self.member_id, self._hours())
        ctx = member_identity.resolve_ticket_envelope(tok)
        self.assertEqual(ctx["member_id"], self.member_id)
        self.assertEqual(ctx["kind"], "ticket")
        self.assertEqual(ctx["name"], "สมชาย ใจดี")

    def test_stale_ticket_rejected_with_thai_message(self):
        hours = self._hours()
        cur = security.ticket_epoch(hours)
        stale = security.encrypt_member_ticket(
            self.member_id, hours, epoch=cur - 5)
        with self.assertRaises(member_identity.MemberIdentityError) as ctx:
            member_identity.resolve_ticket_envelope(stale)
        self.assertIn("QR หมดเวลา", str(ctx.exception))

    def test_future_beyond_grace_rejected(self):
        hours = self._hours()
        cur = security.ticket_epoch(hours)
        far = security.encrypt_member_ticket(
            self.member_id, hours, epoch=cur + 3)
        with self.assertRaises(member_identity.MemberIdentityError) as ctx:
            member_identity.resolve_ticket_envelope(far)
        self.assertIn("QR หมดเวลา", str(ctx.exception))

    def test_decode_any_returns_canonical_json(self):
        tok = security.encrypt_member_ticket(self.member_id, self._hours())
        text = member_identity.decode_any(tok)
        mid, ep = security.parse_ticket_json(text)
        self.assertEqual(mid, self.member_id)

    def test_resolve_decoded_accepts_ticket_body(self):
        tok = security.encrypt_member_ticket(self.member_id, self._hours())
        text = member_identity.decode_any(tok)
        ctx = member_identity.resolve_decoded(text)
        self.assertEqual(ctx["member_id"], self.member_id)

    def test_static_card_accepted_and_resolved(self):
        payload = security.encrypt_qr_data("สมชาย ใจดี|ม.4/1|15|2026-01-01|2027-01-01")
        ctx = member_identity.resolve_scanned(payload)
        self.assertEqual(ctx["kind"], "static")
        self.assertEqual(ctx["member_id"], self.member_id)
        self.assertFalse(ctx["expired"])

    def test_static_cards_can_be_disabled(self):
        payload = security.encrypt_qr_data("สมชาย ใจดี|ม.4/1|15|2026-01-01|2027-01-01")
        try:
            settings_mod.set("accept_static_card_qr", False)
            with self.assertRaises(member_identity.MemberIdentityError) as ctx:
                member_identity.resolve_scanned(payload)
            self.assertIn("บัตร QR แบบเดิม", str(ctx.exception))
        finally:
            settings_mod.set("accept_static_card_qr", True)

    def test_scanned_dispatches_ticket_envelope(self):
        tok = security.encrypt_member_ticket(self.member_id, self._hours())
        ctx = member_identity.resolve_scanned(tok)
        self.assertEqual(ctx["kind"], "ticket")

    def test_scanned_rejects_foreign_payload(self):
        with self.assertRaises(member_identity.MemberIdentityError):
            member_identity.resolve_scanned("v2:AAAAAAAAAAAAAAAA")
        with self.assertRaises(member_identity.MemberIdentityError):
            member_identity.resolve_scanned("random text")

    def test_legacy_v1_badge_still_accepted(self):
        # strict-decode legacy card (the migration path for old members)
        legacy = security.encrypt_legacy_with_key(
            security._AES_KEY, "สมชาย ใจดี|ม.4/1|15|2026-01-01|2027-01-01")
        ctx = member_identity.resolve_scanned(legacy)
        self.assertEqual(ctx["kind"], "static")
        self.assertEqual(ctx["member_id"], self.member_id)

    def test_unknown_ticket_member_rejected(self):
        tok = security.encrypt_member_ticket(999999, self._hours())
        with self.assertRaises(member_identity.MemberIdentityError) as ctx:
            member_identity.resolve_ticket_envelope(tok)
        self.assertIn("ไม่พบสมาชิก", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
