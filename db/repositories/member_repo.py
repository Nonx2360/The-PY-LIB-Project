# db/repositories/member_repo.py
import hashlib
import secrets
from datetime import datetime, timedelta

from .base_repo import BaseRepository
from core.security import hash_password, verify_password

class MemberRepository(BaseRepository):
    def get_all(self):
        return self.fetchall("SELECT * FROM members")

    def get_by_id(self, member_id):
        return self.fetchone("SELECT * FROM members WHERE id = ?", (member_id,))

    def find_by_badge(self, grade, number):
        """Row registered with the same class/number badge, else None.

        เลขที่ (class attendance number) is unique within one grade but can
        repeat across grades, so a member's badge key is (grade, number).
        QR badge matching relies on that key being unique — see
        web/app.py scan_resolve().
        """
        return self.fetchone(
            "SELECT id FROM members WHERE grade = ? AND number = ? LIMIT 1",
            (grade, number))

    def add(self, name, grade, number, reg_date, exp_date, qr_data, qr_path):
        query = '''
            INSERT INTO members (name, grade, number, register_date, expire_date, qrcode_data, qrcode_path)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        '''
        return self.execute(query, (name, grade, number, reg_date, exp_date, qr_data, qr_path))

    def update(self, member_id, name, grade, number, exp_date, qr_data, qr_path):
        query = '''
            UPDATE members
            SET name=?, grade=?, number=?, expire_date=?, qrcode_data=?, qrcode_path=?
            WHERE id=?
        '''
        self.execute(query, (name, grade, number, exp_date, qr_data, qr_path, member_id))

    def delete(self, member_id):
        self.execute("DELETE FROM members WHERE id = ?", (member_id,))

    def delete_many(self, ids):
        """Delete several members in ONE transaction, skipping the unsafe.

        A member referenced by borrow_log or access_log (foreign keys) is
        skipped instead of crashing the batch; missing ids count as skips
        too. Returns (deleted, skipped).
        """
        ids = [i for i in ids if i]
        if not ids:
            return 0, 0

        def _del(conn):
            deleted = skipped = 0
            for mid in ids:
                if not conn.execute(
                        "SELECT 1 FROM members WHERE id=?", (mid,)).fetchone():
                    skipped += 1
                    continue
                if conn.execute(
                        "SELECT 1 FROM borrow_log WHERE member_id=? LIMIT 1",
                        (mid,)).fetchone():
                    skipped += 1
                    continue
                if conn.execute(
                        "SELECT 1 FROM access_log WHERE member_id=? LIMIT 1",
                        (mid,)).fetchone():
                    skipped += 1
                    continue
                conn.execute("DELETE FROM members WHERE id=?", (mid,))
                deleted += 1
            return deleted, skipped

        return self.transact(_del)

    # ================= member portal (rotating QR identity) =================

    ACTIVATION_CODE_TTL_MINUTES = 15

    @staticmethod
    def _hash_code(code: str) -> str:
        """SHA-256 of a one-time activation code (plaintext never at rest)."""
        return hashlib.sha256(code.encode("utf-8")).hexdigest()

    def get_by_badge_for_login(self, grade, number):
        """Login lookup: id + pin_hash only (never logs out the row)."""
        return self.fetchone(
            "SELECT id, pin_hash FROM members "
            "WHERE grade = ? AND number = ? ORDER BY id LIMIT 1",
            (grade, number))

    def has_pin(self, member_id):
        """True once the member activated their portal login."""
        row = self.fetchone("SELECT pin_hash FROM members WHERE id = ?",
                            (member_id,))
        return bool(row and row[0])

    def verify_member_pin(self, member_id, pin):
        row = self.fetchone("SELECT pin_hash FROM members WHERE id = ?",
                            (member_id,))
        if not row or not row[0]:
            return False
        return verify_password(pin, row[0])

    def set_pin(self, member_id, pin):
        """Store a bcrypt-hashed portal PIN (plaintext never persisted)."""
        self.execute("UPDATE members SET pin_hash = ? WHERE id = ?",
                     (hash_password(pin), member_id))

    def clear_pin(self, member_id):
        """Admin reset: member must re-activate (card or new one-time code)."""

        def _reset(conn):
            conn.execute("UPDATE members SET pin_hash = NULL WHERE id = ?",
                         (member_id,))
            # outstanding codes for this member die with the reset, so an
            # old code can't be used to re-seize the account afterwards
            conn.execute(
                "DELETE FROM member_activation_codes WHERE member_id = ?",
                (member_id,))
        self.transact(_reset)

    def pin_state(self, member_id):
        """'set' | 'none' — whether the member activated their login."""
        row = self.fetchone("SELECT pin_hash FROM members WHERE id = ?",
                            (member_id,))
        return "set" if (row and row[0]) else "none"

    def issue_activation_code(self, member_id, ttl_minutes=None):
        """Mint a one-time activation code; returns the PLAINTEXT once.

        Only the SHA-256 hash is stored (see _hash_code). The code never
        appears in logs. TTL defaults to 15 minutes.
        """
        if ttl_minutes is None:
            ttl_minutes = self.ACTIVATION_CODE_TTL_MINUTES
        code = secrets.token_urlsafe(9)          # 12 chars, url-safe
        now = datetime.now()
        self.execute(
            "INSERT INTO member_activation_codes "
            "(member_id, code_hash, created_at, expires_at) VALUES (?, ?, ?, ?)",
            (member_id, self._hash_code(code),
             now.strftime("%Y-%m-%d %H:%M:%S"),
             (now + timedelta(minutes=ttl_minutes)).strftime(
                 "%Y-%m-%d %H:%M:%S")))
        return code

    def find_activation_member(self, code, today=None):
        """Valid (unused, unexpired) code -> member_id, or None.

        Marks nothing yet — the caller claims the code in one transaction
        when the member finishing setting their PIN.
        """
        if not code:
            return None
        row = self.fetchone(
            "SELECT member_id, expires_at, used_at FROM member_activation_codes "
            "WHERE code_hash = ?", (self._hash_code(code.strip()),))
        if not row or row[2]:
            return None
        if today is None:
            today = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if row[1] < today:
            return None
        member = self.fetchone("SELECT id FROM members WHERE id = ?", (row[0],))
        return row[0] if member else None

    def claim_activation_code(self, code, member_id):
        """Mark a code used atomically; True only if WE claimed it first."""
        def _claim(conn):
            cur = conn.execute(
                "UPDATE member_activation_codes "
                "SET used_at = COALESCE(used_at, ?) "
                "WHERE code_hash = ? AND member_id = ? AND used_at IS NULL",
                (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 self._hash_code(code.strip()), member_id))
            return cur.rowcount == 1
        return self.transact(_claim)

    def activate_with_pin(self, member_id, pin, code=None):
        """Finish activation in ONE transaction: claim code (optional) + set PIN.

        Both writes succeed or neither does. The PIN write carries a
        `pin_hash IS NULL` guard — activation must never overwrite an
        existing login (take-over protection), regardless of which
        activation path (card scan or admin code) brought the member here.

        Returns True when activation happened, False when refused:
          * the code is missing/used/expired/for another member, or
          * the member already has a PIN.
        The caller then shows its generic Thai message either way.
        """
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        def _activate(conn):
            if code is not None:
                cur = conn.execute(
                    "UPDATE member_activation_codes "
                    "SET used_at = COALESCE(used_at, ?) "
                    "WHERE code_hash = ? AND member_id = ? AND used_at IS NULL",
                    (now, self._hash_code(code.strip()), member_id))
                if cur.rowcount != 1:
                    return False
            cur = conn.execute(
                "UPDATE members SET pin_hash = ? "
                "WHERE id = ? AND pin_hash IS NULL",
                (hash_password(pin), member_id))
            return cur.rowcount == 1

        return self.transact(_activate)

    def drop_expired_activation_codes(self):
        """Housekeeping: codes older than their TTL can be removed."""
        self.execute(
            "DELETE FROM member_activation_codes WHERE expires_at < ?",
            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),))
