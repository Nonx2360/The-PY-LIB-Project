# db/repositories/reservation_repo.py
from datetime import datetime, timedelta

from .base_repo import BaseRepository


class ReservationRepository(BaseRepository):
    def now(self):
        return datetime.now()

    def now_str(self):
        return self.now().strftime("%Y-%m-%d %H:%M:%S")

    def _expire_offers_conn(self, conn, book_id=None):
        """Mark expired offers as expired (does not auto-promote; caller may)."""
        now = self.now_str()
        if book_id:
            conn.execute(
                "UPDATE book_reservations SET status='expired' "
                "WHERE status='offered' AND expires_at IS NOT NULL AND expires_at < ? AND book_id = ?",
                (now, book_id))
        else:
            conn.execute(
                "UPDATE book_reservations SET status='expired' "
                "WHERE status='offered' AND expires_at IS NOT NULL AND expires_at < ?",
                (now,))

    def _offer_next_conn(self, conn, book_id, offer_hours):
        """If book available and no active offer/waiting-offer conflict, offer next waiting (FIFO)."""
        # only when book is available
        row = conn.execute("SELECT status FROM books WHERE id = ?", (book_id,)).fetchone()
        if not row or row[0] != 'ว่าง':
            return None
        self._expire_offers_conn(conn, book_id)
        active_offer = conn.execute(
            "SELECT id, member_id FROM book_reservations "
            "WHERE book_id = ? AND status = 'offered' AND (expires_at IS NULL OR expires_at >= ?) "
            "ORDER BY id LIMIT 1", (book_id, self.now_str())).fetchone()
        if active_offer:
            return active_offer[0]
        nxt = conn.execute(
            "SELECT id, member_id FROM book_reservations "
            "WHERE book_id = ? AND status = 'waiting' ORDER BY created_at, id LIMIT 1",
            (book_id,)).fetchone()
        if not nxt:
            return None
        now = self.now()
        offered_at = now.strftime("%Y-%m-%d %H:%M:%S")
        expires_at = (now + timedelta(hours=int(offer_hours or 48))).strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            "UPDATE book_reservations SET status='offered', offered_at=?, expires_at=? WHERE id=?",
            (offered_at, expires_at, nxt[0]))
        return nxt[0]

    def active_offer_for_book_conn(self, conn, book_id):
        self._expire_offers_conn(conn, book_id)
        return conn.execute(
            "SELECT id, member_id, expires_at FROM book_reservations "
            "WHERE book_id = ? AND status = 'offered' AND (expires_at IS NULL OR expires_at >= ?) "
            "ORDER BY id LIMIT 1", (book_id, self.now_str())).fetchone()

    def queue_for_book(self, book_id):
        self.execute("UPDATE book_reservations SET status='expired' "
                     "WHERE status='offered' AND expires_at IS NOT NULL AND expires_at < ?",
                     (self.now_str(),))
        return self.fetchall(
            "SELECT r.id, r.member_id, m.name, m.grade, r.status, r.created_at, r.offered_at, r.expires_at "
            "FROM book_reservations r JOIN members m ON m.id = r.member_id "
            "WHERE r.book_id = ? AND r.status IN ('waiting','offered') "
            "ORDER BY CASE r.status WHEN 'offered' THEN 0 ELSE 1 END, r.created_at, r.id",
            (book_id,))

    def member_reservations(self, member_id):
        self.execute("UPDATE book_reservations SET status='expired' "
                     "WHERE status='offered' AND expires_at IS NOT NULL AND expires_at < ?",
                     (self.now_str(),))
        return self.fetchall(
            "SELECT r.id, r.book_id, b.code, b.title, r.status, r.created_at, r.offered_at, r.expires_at "
            "FROM book_reservations r JOIN books b ON b.id = r.book_id "
            "WHERE r.member_id = ? AND r.status IN ('waiting','offered') "
            "ORDER BY r.created_at DESC, r.id DESC", (member_id,))

    def create(self, book_id, member_id, note=None):
        now = self.now_str()
        return self.execute(
            "INSERT INTO book_reservations (book_id, member_id, status, created_at, note) "
            "VALUES (?,?,?,?,?)", (book_id, member_id, 'waiting', now, note))

    def cancel(self, reservation_id, member_id=None, by_admin=False):
        # cancel if waiting/offered (and owned or admin)
        row = self.fetchone(
            "SELECT id, book_id, member_id, status FROM book_reservations WHERE id = ?",
            (reservation_id,))
        if not row:
            raise ValueError("ไม่พบรายการจอง")
        if not by_admin and member_id is not None and row[2] != member_id:
            raise ValueError("ไม่มีสิทธิ์ยกเลิกการจองนี้")
        if row[3] not in ('waiting', 'offered'):
            raise ValueError("ไม่สามารถยกเลิกการจองในสถานะนี้")
        self.execute(
            "UPDATE book_reservations SET status='cancelled', cancelled_at=? WHERE id=?",
            (self.now_str(), reservation_id))
        return row[1]  # book_id
