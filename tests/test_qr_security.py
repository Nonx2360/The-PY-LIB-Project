# tests/test_qr_security.py — tampering, key mismatch, no plaintext fallback
import base64
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

# Isolate secrets before crypto modules load
_env = harness.TempEnv()
_env.apply()

from core import security  # noqa: E402

MEMBER_PAYLOAD = "สมชาย ใจดี|ม.4/1|30|2026-10-30|2027-10-30"


class TestQRv2Roundtrip(unittest.TestCase):
    def test_roundtrip_member_payload(self):
        encrypted = security.encrypt_qr_data(MEMBER_PAYLOAD)
        self.assertTrue(encrypted.startswith(security.GCM_PREFIX))
        self.assertEqual(security.decrypt_qr_data(encrypted), MEMBER_PAYLOAD)

    def test_ciphertext_is_not_plaintext(self):
        encrypted = security.encrypt_qr_data(MEMBER_PAYLOAD)
        self.assertNotIn(MEMBER_PAYLOAD.encode("utf-8"),
                         base64.urlsafe_b64decode(encrypted[len(security.GCM_PREFIX):]))
        self.assertNotIn(MEMBER_PAYLOAD, encrypted)


class TestTamperRejection(unittest.TestCase):
    # audit-test: Tampered QR ciphertext -> Rejected; no plaintext fallback
    def test_flipped_byte_rejected(self):
        encrypted = security.encrypt_qr_data(MEMBER_PAYLOAD)  # v2
        self.assertTrue(encrypted.startswith(security.GCM_PREFIX))
        raw = bytearray(base64.urlsafe_b64decode(encrypted[len(security.GCM_PREFIX):]))
        raw[-1] ^= 0x01  # the last byte is part of the GCM tag
        tampered = security.GCM_PREFIX + base64.urlsafe_b64encode(bytes(raw)).decode("ascii")
        with self.assertRaises(ValueError):
            security.decrypt_qr_data(tampered)

    def test_wrong_key_v2_rejected(self):
        other_key = security.derive_qr_key("another-passphrase", "another-salt")
        payload = security.encrypt_with_key(other_key, MEMBER_PAYLOAD)
        with self.assertRaises(ValueError):
            security.decrypt_qr_data(payload)

    def test_wrong_key_legacy_cbc_rejected(self):
        # audit-test: QR encrypted with an old key -> rejected
        other_key = security.derive_qr_key("old-exposed-passphrase", "oldsalt")
        # 30 different payloads: chance of a wrong-key decryption passing strict
        # PKCS7+UTF-8 validation is ~1/255 per payload -> flake probability ~0.
        for i in range(30):
            legacy = security.encrypt_legacy_with_key(other_key, f"member{i}|grade|num|2026|2027")
            with self.assertRaises(ValueError):
                security.decrypt_qr_data(legacy)


class TestNoPlaintextFallback(unittest.TestCase):
    def test_legacy_plain_base64_no_longer_accepted(self):
        # audit-test: pre-refactor plaintext cards are rejected (fail closed)
        legacy = base64.b64encode("oldmember|ม.1/1|15|2025-01-01|2026-01-01".encode()).decode("utf-8")
        with self.assertRaises(ValueError):
            security.decrypt_qr_data(legacy)

    def test_truncated_payload_rejected(self):
        short = security.encrypt_qr_data("x")[:-3]
        with self.assertRaises(ValueError):
            security.decrypt_qr_data(short)

    def test_garbage_rejected(self):
        for bad in ("", "v2:", "v2:%", "not-base64!!!", "v2:AAAA"):
            with self.assertRaises(ValueError):
                security.decrypt_qr_data(bad)


class TestLegacyCBCStillSupported(unittest.TestCase):
    def test_current_key_legacy_payload_decodes(self):
        # Existing member payloads (v1 CBC under the CURRENT key) keep working
        legacy = security.encrypt_legacy_with_key(security._AES_KEY, MEMBER_PAYLOAD)
        self.assertEqual(security.decrypt_qr_data(legacy), MEMBER_PAYLOAD)


class TestKeyRotation(unittest.TestCase):
    def test_rotate_and_reencrypt_roundtrip(self):
        # equivalent of tools/rotate_qr_key.py core operation
        old_key = security.derive_qr_key("exposed-pass", "exposed-salt")
        stored = security.encrypt_with_key(old_key, MEMBER_PAYLOAD)
        new_pass, new_salt = "new-pass", "new-salt"
        new_key = security.derive_qr_key(new_pass, new_salt)
        # old key reads it
        self.assertEqual(security.decrypt_with_key(old_key, stored), MEMBER_PAYLOAD)
        # new key rejects the old payload...
        with self.assertRaises(ValueError):
            security.decrypt_with_key(new_key, stored)
        # ...and after re-encryption the new key reads it, old rejects it
        rotated = security.encrypt_with_key(new_key, MEMBER_PAYLOAD)
        self.assertEqual(security.decrypt_with_key(new_key, rotated), MEMBER_PAYLOAD)
        with self.assertRaises(ValueError):
            security.decrypt_with_key(old_key, rotated)


if __name__ == "__main__":
    unittest.main()
