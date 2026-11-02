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
import binascii
import json
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


# ------------------------- rotating ticket (v3) -------------------------
# A member's identity credential that expires on its own: the payload seals
# only {"m": member_id, "e": epoch}, where epoch = floor(unix_time /
# rotation_seconds). The scanner accepts the current epoch plus `grace`
# neighbours, so a photographed/s stolen QR is worthless after that window —
# unlike the static v2 card QR, which never changes.
# Stateless by design: nothing is written to the database when a ticket is
# issued or verified — the "rotation" is pure time math shared by the web
# app (which renders the member portal QR) and the desktop scanner.

TICKET_PREFIX = "v3:"
TICKET_AAD = b"pylib-member-ticket-v3"
_TICKET_JSON_KEYS = {"m", "e"}


def ticket_epoch(rotation_hours: int, now=None) -> int:
    """Current window index for the given rotation period.

    `rotation_hours` is the admin-configured member_qr_rotation_hours
    (clamped by callers); epochs from different periods are not comparable,
    so changing the period simply invalidates QRs issued before the change
    fell out of the ±grace window.
    """
    from datetime import datetime
    if now is None:
        now = datetime.now()
    seconds = max(1, int(round(rotation_hours))) * 3600
    return int(now.timestamp() // seconds)


def ticket_epoch_seconds_left(rotation_hours: int, now=None) -> int:
    """Whole seconds until the current window flips to the next one."""
    from datetime import datetime
    if now is None:
        now = datetime.now()
    seconds = max(1, int(round(rotation_hours))) * 3600
    return seconds - (int(now.timestamp()) % seconds)


def ticket_window_ok(epoch: int, rotation_hours: int, grace: int = 1,
                     now=None) -> bool:
    """True when an epoch belongs to the accepted range around now.

    grace covers clock drift between the machine that issued the QR
    (member page, desktop login page) and the machine that scans it. The
    ± range means a member can also present the "next" QR slightly early —
    harmless because it is still anchored to their own id.
    """
    return abs(int(epoch) - ticket_epoch(rotation_hours, now)) <= max(0, int(grace))


def encrypt_member_ticket(member_id: int, rotation_hours: int, epoch=None,
                          *, key=None) -> str:
    """Seal a v3 ticket for one member + one window (v3: prefix, GCM+AAD)."""
    if key is None:
        key = _AES_KEY
    if int(member_id) < 1:
        raise ValueError("ticket needs a real member id")
    if epoch is None:
        epoch = ticket_epoch(rotation_hours)
    plain = json.dumps({"m": int(member_id), "e": int(epoch)},
                       sort_keys=True, separators=(",", ":"))
    nonce = os.urandom(GCM_NONCE_LEN)
    sealed = AESGCM(key).encrypt(nonce, plain.encode("utf-8"), TICKET_AAD)
    return TICKET_PREFIX + base64.urlsafe_b64encode(
        nonce + sealed).decode("ascii")


def parse_ticket_json(text):
    """Validate an ALREADY-DECRYPTED ticket body: returns (member_id, epoch).

    Strict: exact keyset {m, e}, integers only (bools rejected), sane
    ranges. This exists so the desktop camera flow — which performs the
    AES-GCM decrypt with the .env key inside its frame loop — can re-run
    the same field validation without accepting unsealed JSON anywhere an
    attacker-supplied payload goes.
    """
    try:
        data = json.loads(text)
        if (not isinstance(data, dict) or set(data.keys()) != _TICKET_JSON_KEYS
                or isinstance(data["m"], bool) or isinstance(data["e"], bool)
                or not isinstance(data["m"], int)
                or not isinstance(data["e"], int)):
            raise ValueError("ticket body is not {m, e} integers")
        if data["m"] < 1 or data["e"] < 0:
            raise ValueError("ticket fields out of range")
        return data["m"], data["e"]
    except (json.JSONDecodeError, ValueError) as exc:
        raise ValueError(
            "member ticket rejected: wrong key, tampered data, or "
            "unsupported format") from exc


def decrypt_member_ticket(payload: str, *, key=None):
    """Open a v3 ticket; returns (member_id, epoch). Raises ValueError."""
    if key is None:
        key = _AES_KEY
    if not payload.startswith(TICKET_PREFIX):
        raise ValueError("not a v3 ticket")
    try:
        raw = base64.urlsafe_b64decode(
            payload[len(TICKET_PREFIX):].encode("ascii"))
        nonce, sealed = raw[:GCM_NONCE_LEN], raw[GCM_NONCE_LEN:]
        if len(nonce) != GCM_NONCE_LEN or len(sealed) < GCM_TAG_LEN:
            raise ValueError("ticket truncated")
        plain = AESGCM(key).decrypt(nonce, sealed, TICKET_AAD)
        return parse_ticket_json(plain.decode("utf-8"))
    except (InvalidTag, UnicodeDecodeError, ValueError, IndexError,
            binascii.Error) as exc:
        raise ValueError(
            "member ticket rejected: wrong key, tampered data, or "
            "unsupported format") from exc


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
