# tools/rotate_qr_key.py — rotate the QR encryption key and re-encrypt stored payloads
#
# Run this ONCE on each machine that has an existing library database, BEFORE
# the next app launch, whenever SECRET_PASSPHRASE/SECRET_SALT must change
# (e.g. the old values were ever exposed).
#
# What it does (in order, aborting on any failure):
#   1. Takes a verified SQLite backup of the database into backups/.
#   2. Decrypts every member's stored qrcode_data under the CURRENT key loaded
#      from the existing .env. Rows in legacy plain-base64 format are also
#      handled (they are re-encrypted/upgraded in place).
#   3. Generates a NEW SECRET_PASSPHRASE + SECRET_SALT.
#   4. Re-encrypts every payload with the new key (AES-GCM "v2:"), rewrites the
#      member QR PNG images and member card PDFs.
#   5. Writes the new secrets to .env (preserving every other key).
#
# Usage (from the project root):
#   venv\Scripts\python tools\rotate_qr_key.py [--no-cards]
#
# See docs/SECURITY.md for the full runbook.
import argparse
import base64
import os
import secrets
import sqlite3
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

from config.constants import DB_PATH  # noqa: E402
from config.env_store import update_secrets  # noqa: E402
from core import security  # noqa: E402  (loads CURRENT .env -> old working key)
from services.backup_service import create_backup, verify_backup  # noqa: E402


def _decode_current(payload):
    """Decode a stored payload under the current key; None if unreadable."""
    if not payload:
        return None
    try:
        return security.decrypt_qr_data(payload)
    except Exception:
        pass
    # legacy pre-refactor rows stored plain base64 text
    try:
        raw = base64.b64decode(payload).decode("utf-8")
        if "|" in raw:
            return raw  # plaintext member data; upgraded to AEAD below
    except Exception:
        pass
    return None


def main():
    parser = argparse.ArgumentParser(description="rotate QR encryption keys")
    parser.add_argument("--no-cards", action="store_true",
                        help="skip regenerating member card PDFs (DB + QR images always updated)")
    args = parser.parse_args()

    if not os.path.exists(DB_PATH):
        print(f"No database found at {DB_PATH} — nothing to rotate.")
        return 0

    # 1) backup first
    backup = create_backup(DB_PATH)
    report = verify_backup(backup)
    if not report["ok"]:
        raise SystemExit(f"ABORT: fresh backup failed verification: {report}")
    print(f"[1/5] Backup OK: {backup}")

    # 2) decrypt everything under the current key before touching anything
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("PRAGMA foreign_keys=ON")
        rows = conn.execute(
            "SELECT id, name, grade, number, register_date, expire_date, qrcode_data, qrcode_path "
            "FROM members ORDER BY id").fetchall()
    finally:
        conn.close()

    decoded = {}
    failed = []
    for mid, name, grade, number, reg, exp, qrcode_data, qrcode_path in rows:
        plain = _decode_current(qrcode_data)
        if plain is None:
            failed.append(mid)
        else:
            decoded[mid] = (name, grade, number, reg, exp, plain, qrcode_path)

    if failed:
        raise SystemExit(
            f"ABORT: {len(failed)} member payload(s) could not be decoded under the "
            f"CURRENT key (ids: {failed}). The old .env secret may already have been "
            f"replaced, or the data is corrupt. Nothing was modified."
        )
    print(f"[2/5] Decoded {len(decoded)}/{len(rows)} member payloads under the current key.")

    # 3) new key material
    new_passphrase = secrets.token_urlsafe(32)
    new_salt = secrets.token_urlsafe(12)
    new_key = security.derive_qr_key(new_passphrase, new_salt)
    print("[3/5] New passphrase + salt generated.")

    # 4) re-encrypt payloads and rewrite QR images
    import qrcode
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("PRAGMA foreign_keys=ON")
        cur = conn.cursor()
        updated = 0
        for mid, meta in decoded.items():
            qrcode_path = meta[6]
            new_payload = security.encrypt_with_key(new_key, meta[5])
            cur.execute("UPDATE members SET qrcode_data = ? WHERE id = ?",
                        (new_payload, mid))
            if qrcode_path:
                os.makedirs(os.path.dirname(qrcode_path) or ".", exist_ok=True)
                qr = qrcode.QRCode(version=1, box_size=10, border=4)
                qr.add_data(new_payload)
                qr.make(fit=True)
                qr.make_image(fill_color="black", back_color="white").save(qrcode_path)
            updated += 1
        conn.commit()
    finally:
        conn.close()

    cards_regenerated = 0
    if not args.no_cards:
        from services.pdf_service import PDFService
        conn = sqlite3.connect(DB_PATH)
        try:
            conn.execute("PRAGMA foreign_keys=ON")
            for (mid, payload) in conn.execute(
                    "SELECT id, qrcode_data FROM members ORDER BY id").fetchall():
                if mid not in decoded:
                    continue
                name, grade, number, reg, exp, plain, _ = decoded[mid]
                card_path = f"assets/cards/member_{mid}.pdf"
                os.makedirs("assets/cards", exist_ok=True)
                PDFService.generate_member_card_pdf(name, number,
                                                    "assets/logos/school_logo.png",
                                                    card_path, payload)
                cards_regenerated += 1
        finally:
            conn.close()

    if args.no_cards:
        print(f"[4/5] Re-encrypted {updated} payloads; QR images rewritten (cards skipped).")
    else:
        print(f"[4/5] Re-encrypted {updated} payloads; QR images rewritten; "
              f"{cards_regenerated} card PDFs regenerated.")

    # 5) persist new secrets
    update_secrets({
        "SECRET_PASSPHRASE": new_passphrase,
        "SECRET_SALT": new_salt,
    })
    security._AES_KEY = new_key  # keep this process consistent for programmatic use
    print("[5/5] .env updated with the new secrets.")
    print("DONE. Follow-ups:")
    print("  * Copy the updated .env to any other machine running this app BEFORE its next boot.")
    print("  * Physical member cards still carry old-format payloads — re-issue them.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
