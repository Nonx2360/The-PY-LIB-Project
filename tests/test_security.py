# tests/test_security.py
import base64
import unittest

from core.security import hash_password, verify_password, encrypt_qr_data, decrypt_qr_data


class TestPasswordHashing(unittest.TestCase):
    def test_hash_and_verify(self):
        hashed = hash_password("secret123")
        self.assertTrue(hashed.startswith("$2"))
        self.assertTrue(verify_password("secret123", hashed))

    def test_wrong_password_rejected(self):
        hashed = hash_password("secret123")
        self.assertFalse(verify_password("wrong", hashed))

    def test_hashes_are_salted(self):
        self.assertNotEqual(hash_password("same"), hash_password("same"))


class TestQRDataEncryption(unittest.TestCase):
    def test_roundtrip(self):
        payload = "ม.1/2|สมชาย|เลขที่ 15|2026-01-01|2027-01-01"
        encrypted = encrypt_qr_data(payload)
        self.assertEqual(decrypt_qr_data(encrypted), payload)

    def test_ciphertext_not_plaintext(self):
        payload = "secret-member-data"
        encrypted = encrypt_qr_data(payload)
        self.assertNotIn(payload, encrypted)
        # plain base64 decode of the payload must NOT yield the data
        raw = base64.urlsafe_b64decode(encrypted)
        self.assertNotIn(payload.encode("utf-8"), raw)

    def test_legacy_plain_base64_fallback(self):
        legacy = base64.b64encode("legacy|data|123".encode("utf-8")).decode("utf-8")
        self.assertEqual(decrypt_qr_data(legacy), "legacy|data|123")


if __name__ == "__main__":
    unittest.main()