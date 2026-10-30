# tests/test_security.py
import base64
import unittest

import core.security as sec


class TestPasswordHashing(unittest.TestCase):
    def test_hash_and_verify(self):
        hashed = sec.hash_password("secret123")
        self.assertTrue(hashed.startswith("$2"))
        self.assertTrue(sec.verify_password("secret123", hashed))

    def test_wrong_password_rejected(self):
        hashed = sec.hash_password("secret123")
        self.assertFalse(sec.verify_password("wrong", hashed))

    def test_hashes_are_salted(self):
        self.assertNotEqual(sec.hash_password("same"), sec.hash_password("same"))


class TestQRDataEncryption(unittest.TestCase):
    def test_roundtrip(self):
        payload = "ม.1/2|สมชาย|เลขที่ 15|2026-01-01|2027-01-01"
        encrypted = sec.encrypt_qr_data(payload)
        self.assertEqual(sec.decrypt_qr_data(encrypted), payload)

    def test_new_payloads_use_authenticated_encryption(self):
        encrypted = sec.encrypt_qr_data("secret-member-data")
        self.assertTrue(encrypted.startswith(sec.GCM_PREFIX),
                        "new payloads must use the v2 (AEAD) envelope")

    def test_ciphertext_not_plaintext(self):
        payload = "secret-member-data"
        encrypted = sec.encrypt_qr_data(payload)
        self.assertNotIn(payload, encrypted)
        # plain base64 decode of the payload must NOT yield the data
        raw = base64.urlsafe_b64decode(encrypted[len(sec.GCM_PREFIX):])
        self.assertNotIn(payload.encode("utf-8"), raw)

    def test_legacy_plain_base64_rejected(self):
        # audit-test: plaintext/base64 fallback was removed — fail closed
        legacy = base64.b64encode(b"legacy|data|123").decode("utf-8")
        with self.assertRaises(ValueError):
            sec.decrypt_qr_data(legacy)


if __name__ == "__main__":
    unittest.main()
