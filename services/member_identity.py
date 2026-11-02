# services/member_identity.py — one resolver for every member QR payload
#
# The web scanner (/scan/resolve) and the desktop camera flow used to each
# decode + badge-match separately. With rotating tickets (v3) introduced,
# both flows MUST accept exactly the same payload family and apply the same
# expiry-window policy, so the logic lives here once.
#
# Accepted payload families (raw === encrypted as it comes off the scanner):
#   v3:"..."  member ticket (rotating identity) — member_id + epoch inside
#   v2:"..."  static card QR (printed card)     — name|grade|number|reg|exp
#   (no prefix) legacy card QR (v1 CBC, strict) — same badge shape
#
# Every invalid path raises MemberIdentityError whose message is the Thai
# text shown to the desk — no payload content ever echoes back, and logs
# carry outcome counters only (no PII), matching the bulk endpoints.
import json

from core.security import (decrypt_member_ticket, decrypt_qr_data,
                           parse_ticket_json, ticket_window_ok)
from db.repositories.member_repo import MemberRepository


class MemberIdentityError(ValueError):
    """Raised with a Thai message safe to show at the desk."""


def rotation_hours():
    """Admin-set window length, clamped to what the UI offers (1..72 h)."""
    from config import settings
    try:
        value = int(settings.get("member_qr_rotation_hours"))
    except (TypeError, ValueError):
        value = 12
    return min(72, max(1, value))


def grace_steps():
    from config import settings
    try:
        value = int(settings.get("member_qr_grace_steps"))
    except (TypeError, ValueError):
        value = 1
    return min(2, max(0, value))


def static_cards_accepted():
    """Whether the printed card's static QR still identifies members."""
    from config import settings
    return bool(settings.get("accept_static_card_qr"))


def _today():
    from datetime import datetime
    return datetime.now().strftime("%Y-%m-%d")


def _member_row(repo, sql, params):
    return (repo or MemberRepository()).fetchone(sql, params)


def resolve_badge(plaintext, repo=None):
    """Legacy identity path: 'name|grade|number|reg|exp' -> member."""
    if not static_cards_accepted():
        raise MemberIdentityError(
            "ระบบปิดการรับบัตร QR แบบเดิมแล้ว — กรุณาใช้ QR จากหน้าของนักเรียน")
    parts = [p.strip() for p in (plaintext or "").split("|")]
    if len(parts) != 5 or not all(parts[:3]):
        raise MemberIdentityError("รูปแบบ QR ไม่ถูกต้อง")
    name, grade, number, _reg, _exp = parts
    rows = (repo or MemberRepository()).fetchall(
        "SELECT id, name, grade, number, register_date, expire_date "
        "FROM members WHERE number = ?", (number,))
    if not rows:
        raise MemberIdentityError("ไม่พบสมาชิกนี้ในระบบ")
    exact = [r for r in rows if r[1] == name and r[2] == grade]
    row = (exact or rows)[0]
    return {"kind": "static", "member_id": row[0], "name": row[1],
            "grade": row[2], "number": row[3],
            "expire_date": row[5], "expired": bool(row[5] and row[5] < _today())}


def resolve_ticket_envelope(raw_payload, repo=None):
    """Rotating identity path for a SEALED 'v3:...' payload (web flow)."""
    try:
        member_id, epoch = decrypt_member_ticket(raw_payload)
    except ValueError:
        # the lower-layer message may reference payload content — suppress it
        raise MemberIdentityError(
            "QR ไม่ถูกต้อง (ระบบเปิดไม่ได้หรือคีย์ไม่ตรง)") from None
    return _member_context(member_id, epoch, repo)


def _member_context(member_id, epoch, repo):
    """Window check + row lookup shared by both ticket entry points."""
    if not ticket_window_ok(epoch, rotation_hours(), grace_steps()):
        raise MemberIdentityError(
            "QR หมดเวลา — ให้สมาชิกเปิดหน้าของตัวเองแล้วแสดง QR ใหม่")
    row = _member_row(
        repo, "SELECT id, name, grade, number, register_date, expire_date "
        "FROM members WHERE id = ?", (member_id,))
    if row is None:
        raise MemberIdentityError("ไม่พบสมาชิกนี้ในระบบ")
    return {"kind": "ticket", "member_id": row[0], "name": row[1],
            "grade": row[2], "number": row[3],
            "expire_date": row[5], "expired": bool(row[5] and row[5] < _today())}


def resolve_scanned(raw_payload, repo=None):
    """Full resolver for an ENCRYPTED scan payload (the web flow).

    Dispatches on the envelope prefix; decodes with the .env key; runs the
    window policy for tickets and the static-card policy for cards.
    """
    raw = raw_payload or ""
    if raw.startswith("v3:"):
        return resolve_ticket_envelope(raw, repo)
    try:
        plaintext = decrypt_qr_data(raw)
    except ValueError:
        # the lower-layer message may reference payload content — suppress it
        raise MemberIdentityError(
            "QR ไม่ถูกต้อง (ระบบเปิดไม่ได้หรือคีย์ไม่ตรง)") from None
    return resolve_badge(plaintext, repo)


def resolve_decoded(decoded_text, repo=None):
    """Full resolver for an ALREADY-DECRYPTED scan output (desktop flow).

    CameraService hands back decrypted text: badge shape for static cards,
    raw ticket JSON ('{"e":...,\"m\":...}') for rotating ones. The ticket
    branch re-validates the fields and window but does NOT treat this as
    unauthenticated input — the camera loop already performed the AES-GCM
    decrypt with the same .env key the issuing page used.
    """
    text = (decoded_text or "").strip()
    if text.startswith("{") and text.endswith("}"):
        try:
            member_id, epoch = parse_ticket_json(text)
        except ValueError:
            # the lower-layer message may reference payload content — suppress it
            raise MemberIdentityError(
                "QR ไม่ถูกต้อง (ข้อมูลโค้ดเสียหาย)") from None
        return _member_context(member_id, epoch, repo)
    return resolve_badge(text, repo)


def member_ticket_plaintext(member_id, epoch):
    """JSON body a v3 ticket decrypts to (for rendering + tests)."""
    return json.dumps({"m": int(member_id), "e": int(epoch)},
                      sort_keys=True, separators=(",", ":"))


def decode_any(raw_payload):
    """Decrypt ANY accepted envelope ('v3:...', 'v2:...', legacy) to text.

    Used by the desktop camera loop, which hands the decoded text to the
    views. Ticket output is the canonical {'e':..,'m':..} JSON body.
    """
    raw = raw_payload or ""
    if raw.startswith("v3:"):
        member_id, epoch = decrypt_member_ticket(raw)
        return member_ticket_plaintext(member_id, epoch)
    return decrypt_qr_data(raw)
