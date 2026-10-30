# core/security.py — password hashing + QR payload encryption
#
# QR payload formats:
#   v2 (current): "v2:" + base64url( nonce[12] || AES-256-GCM ciphertext || tag[16] )
#                 Authenticated: any tampering or wrong key is rejected.
#   v1 (legacy):  base64url( IV[16] || AES-256-CBC ciphertext )  — no integrity.
#                 Decode is STRICT (full PKCS7 + UTF-8 validation) and exists
#                 only so existing member cards keep working until they are
#                 migrated with tools/rotate_qr_key.py. It can be removed once
#                 every member has been re-issued a card.
# Removed: base64/plaintext fallback and exception swallowing. Anything that
# fails validation raises ValueError.
#
# Key derivation: PBKDF2-HMAC-SHA256 (100k iterations) over SECRET_PASSPHRASE
# with SECRET_SALT, both loaded from .env via config.env_store. Never
# hardcoded — a fresh install generates strong random values on first boot.
import base64
import os

import bcrypt
from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from dotenv import load_dotenv

from config.env_store import ensure_secrets

GCM_PREFIX = "v2:"
GCM_NONCE_LEN = 12
GCM_TAG_LEN = 16
CBC_IV_LEN = 16
KDF_ITERATIONS = 100_000

load_dotenv()  # pre-refactor convenience: pick up .env if present


def derive_qr_key(passphrase: str, salt: str) -> bytes:
    """PBKDF2-HMAC-SHA256 -> 32-byte AES-256 key."""
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt.encode("utf-8"),
        iterations=KDF_ITERATIONS,
    )
    return kdf.derive(passphrase.encode("utf-8"))


# Module-level working key (created once at import; the rotation tool re-derives
# keys explicitly with derive_qr_key and encrypt_with_key/decrypt_with_key).
_SECRETS = ensure_secrets()
_AES_KEY = derive_qr_key(_SECRETS["SECRET_PASSPHRASE"], _SECRETS["SECRET_SALT"])


def hash_password(password: str) -> str:
    """Hash password using bcrypt."""
    salt = bcrypt.gensalt()
    hashed = bcrypt.hashpw(password.encode('utf-8'), salt)
    return hashed.decode('utf-8')


def verify_password(password: str, hashed: str) -> bool:
    """Verify password using bcrypt."""
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))
    except Exception:
        return False


# ------------------------- AEAD (v2) helpers -------------------------

def encrypt_with_key(key: bytes, plain_text: str) -> str:
    """v2 envelope encrypted with an explicit key (used by the rotation tool)."""
    nonce = os.urandom(GCM_NONCE_LEN)
    sealed = AESGCM(key).encrypt(nonce, plain_text.encode("utf-8"), None)
    return GCM_PREFIX + base64.urlsafe_b64encode(nonce + sealed).decode("ascii")


def decrypt_with_key(key: bytes, payload: str) -> str:
    """Decrypt a v2 envelope with an explicit key. Raises ValueError on any failure."""
    if not payload.startswith(GCM_PREFIX):
        raise ValueError("not a v2 payload")
    try:
        raw = base64.urlsafe_b64decode(payload[len(GCM_PREFIX):].encode("ascii"))
        nonce, sealed = raw[:GCM_NONCE_LEN], raw[GCM_NONCE_LEN:]
        if len(nonce) != GCM_NONCE_LEN or len(sealed) < GCM_TAG_LEN:
            raise ValueError("payload truncated")
        plain = AESGCM(key).decrypt(nonce, sealed, None)
        return plain.decode("utf-8")
    except (InvalidTag, UnicodeDecodeError, ValueError, IndexError) as exc:
        raise ValueError(
            "QR payload rejected: wrong key, tampered data, or unsupported format"
        ) from exc


# ------------------------- legacy (v1 CBC) -------------------------

def encrypt_legacy_with_key(key: bytes, plain_text: str) -> str:
    """v1 envelope (AES-256-CBC, IV-prefixed) — only used by the rotation tool
    and tests to create representative legacy payloads."""
    iv = os.urandom(CBC_IV_LEN)
    data = plain_text.encode("utf-8")
    pad_len = 16 - (len(data) % 16)
    padded = data + bytes([pad_len]) * pad_len
    cipher = Cipher(algorithms.AES(key), modes.CBC(iv))
    enc = cipher.encryptor()
    ct = enc.update(padded) + enc.finalize()
    return base64.urlsafe_b64encode(iv + ct).decode("ascii")
    # NOTE: intentionally no trailing zero-padding trim needed; PKCS7 is clean.


def _decrypt_legacy_cbc_strict(payload: str) -> str:
    """Strict v1 decode: full PKCS7 check, UTF-8 required, no fallback."""
    try:
        combined = base64.urlsafe_b64decode(payload.encode("utf-8"))
        iv, ct = combined[:CBC_IV_LEN], combined[CBC_IV_LEN:]
        if len(iv) != CBC_IV_LEN or (len(ct) == 0 or len(ct) % 16 != 0):
            raise ValueError("payload truncated")
        cipher = Cipher(algorithms.AES(_AES_KEY), modes.CBC(iv))
        dec = cipher.decryptor()
        padded = dec.update(ct) + dec.finalize()
        pad_len = padded[-1]
        if not (1 <= pad_len <= 16):
            raise ValueError("bad padding")
        if padded[-pad_len:] != bytes([pad_len]) * pad_len:
            raise ValueError("bad padding")
        return padded[:-pad_len].decode("utf-8")
    except (UnicodeDecodeError, ValueError) as exc:
        raise ValueError(
            "QR payload rejected: wrong key, tampered data, or unsupported format"
        ) from exc


# ------------------------- public API -------------------------

def encrypt_qr_data(plain_text: str) -> str:
    """Encrypt member data for a QR code (v2 AEAD envelope, working key)."""
    return encrypt_with_key(_AES_KEY, plain_text)


def decrypt_qr_data(payload: str) -> str:
    """Decrypt a QR payload under the working key.

    Accepted formats: v2 (AES-GCM, authenticated) and v1 (strict CBC).
    Plain base64 payloads and anything that fails validation are REJECTED.
    """
    if payload.startswith(GCM_PREFIX):
        return decrypt_with_key(_AES_KEY, payload)
    return _decrypt_legacy_cbc_strict(payload)
