# db/repositories/borrow_repo.py
import sqlite3
from datetime import datetime, timedelta

import config.settings as settings
from config.constants import BOOK_STATUS_AVAILABLE, BOOK_STATUS_BORROWED
from .base_repo import BaseRepository


class BorrowError(Exception):
    """Business-rule violation; str(e) is a user-facing Thai message."""
    pass


class BorrowRepository(BaseRepository):
    def get_active_borrows(self):
        query = '''
            SELECT bl.id, m.name, b.title, bl.borrow_date, bl.return_due
            FROM borrow_log bl
            JOIN members m ON bl.member_id = m.id
            JOIN books b ON bl.book_id = b.id
            WHERE bl.returned = 0
        '''
        return self.fetchall(query)

    def add(self, member_id, book_id, borrow_date, return_due):
        """Raw insert — callers that must uphold invariants use try_borrow()."""
        query = '''
            INSERT INTO borrow_log (member_id, book_id, borrow_date, return_due, returned)
            VALUES (?, ?, ?, ?, 0)
        '''
        return self.execute(query, (member_id, book_id, borrow_date, return_due))

    def mark_returned(self, record_id):
        query = "UPDATE borrow_log SET returned = 1 WHERE id = ?"
        self.execute(query, (record_id,))

    # ------------------- atomic invariants-enforcing operations -------------------

    def _expire_and_offer_next(self, conn, book_id):
        """Reservation-aware: expire stale offers and (if still available) offer next waiting."""
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # expire stale offers for this book
        conn.execute(
            "UPDATE book_reservations SET status='expired' "
            "WHERE book_id = ? AND status='offered' AND expires_at IS NOT NULL AND expires_at < ?",
            (book_id, now_str))
        # only if available and no active offer, promote next waiting
        brow = conn.execute("SELECT status FROM books WHERE id = ?", (book_id,)).fetchone()
        if not brow or brow[0] != BOOK_STATUS_AVAILABLE:
            return
        active = conn.execute(
            "SELECT id FROM book_reservations WHERE book_id = ? AND status='offered' "
            "AND (expires_at IS NULL OR expires_at >= ?) LIMIT 1",
            (book_id, now_str)).fetchone()
        if active:
            return
        nxt = conn.execute(
            "SELECT id FROM book_reservations WHERE book_id = ? AND status='waiting' "
            "ORDER BY created_at, id LIMIT 1", (book_id,)).fetchone()
        if not nxt:
            return
        hours = int(settings.get("reservation_offer_hours") or 48)
        offered_at = now_str
        expires_at = (datetime.now() + timedelta(hours=hours)).strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            "UPDATE book_reservations SET status='offered', offered_at=?, expires_at=? WHERE id=?",
            (offered_at, expires_at, nxt[0]))

    def try_borrow(self, member_id, book_id, borrow_date, return_due):
        """Atomically lend one book (reservation-aware, no silent loans).

        Guarantees, enforced inside a single IMMEDIATE transaction:
          * the book exists and is BOOK_STATUS_AVAILABLE ('ว่าง') — its
            status flips to BOOK_STATUS_BORROWED ('ยืมแล้ว') only via a
            conditional UPDATE (rowcount == 1), so concurrent borrowers
            cannot double-lend;
          * expired offers are cleared; an active offer for another member
            blocks borrowing, an active offer for this member is fulfilled
            as part of this explicit borrow;
          * member has no overdue loan and is below max_books_per_member;
          * on any failure the transaction rolls back — no partial state.

        Raises BorrowError with a Thai, user-facing message.
        """
        def _tx(conn):
            row = conn.execute(
                "SELECT status FROM books WHERE id = ?", (book_id,)).fetchone()
            if row is None:
                raise BorrowError("ไม่พบหนังสือในระบบ")
            if row[0] != BOOK_STATUS_AVAILABLE:
                raise BorrowError(f"หนังสือเล่มนี้ไม่พร้อมให้ยืม (สถานะ: {row[0]})")

            # reservation gate (copy-level FIFO): expire stale and enforce offer
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn.execute(
                "UPDATE book_reservations SET status='expired' "
                "WHERE book_id = ? AND status='offered' AND expires_at IS NOT NULL AND expires_at < ?",
                (book_id, now_str))
            offer = conn.execute(
                "SELECT id, member_id FROM book_reservations WHERE book_id = ? AND status='offered' "
                "AND (expires_at IS NULL OR expires_at >= ?) ORDER BY id LIMIT 1",
                (book_id, now_str)).fetchone()
            if offer and int(offer[1]) != int(member_id):
                raise BorrowError("หนังสือเล่มนี้ถูกจอง/เสนอให้สมาชิกอื่นอยู่")
            offer_id = offer[0] if offer and int(offer[1]) == int(member_id) else None

            cur = conn.execute(
                "UPDATE books SET status = ? WHERE id = ? AND status = ?",
                (BOOK_STATUS_BORROWED, book_id, BOOK_STATUS_AVAILABLE))
            if cur.rowcount != 1:
                raise BorrowError(f"หนังสือเล่มนี้ไม่พร้อมให้ยืม (สถานะ: {row[0]})")

            overdue = conn.execute(
                "SELECT COUNT(*) FROM borrow_log "
                "WHERE member_id=? AND returned=0 AND return_due < date('now')",
                (member_id,)).fetchone()[0]
            if overdue > 0:
                raise BorrowError("ไม่สามารถยืมได้ — มีหนังสือค้างส่ง")

            current = conn.execute(
                "SELECT COUNT(*) FROM borrow_log WHERE member_id=? AND returned=0",
                (member_id,)).fetchone()[0]
            max_books = settings.get("max_books_per_member")
            if current >= max_books:
                raise BorrowError(f"ยืมครบจำนวนที่กำหนดแล้ว ({max_books} เล่ม)")

            cur = conn.execute(
                "INSERT INTO borrow_log (member_id, book_id, borrow_date, return_due, returned) "
                "VALUES (?, ?, ?, ?, 0)",
                (member_id, book_id, borrow_date, return_due))
            if offer_id:
                conn.execute(
                    "UPDATE book_reservations SET status='fulfilled', fulfilled_at=? WHERE id=?",
                    (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), offer_id))
            return cur.lastrowid

        try:
            return self.transact(_tx)
        except sqlite3.IntegrityError as exc:
            raise BorrowError("ข้อมูลไม่ถูกต้อง ไม่สามารถบันทึกการยืมได้") from exc

    def try_return(self, record_id):
        """Atomically mark a borrow record returned and free its book.

        The book identifier is taken from the borrow_log row itself — client
        data is never trusted. Guarantees:
          * unknown/non-open record -> BorrowError, nothing changes;
          * borrow_log.returned=1 and books.status=BOOK_STATUS_AVAILABLE
            ('ว่าง') applied together.
          * if a waiting reservation exists, the next (FIFO) is offered in the
            same transaction (no silent loan); the book remains available but
            held for that offer until claimed/expired.
        """
        def _tx(conn):
            row = conn.execute(
                "SELECT book_id FROM borrow_log WHERE id = ? AND returned = 0",
                (record_id,)).fetchone()
            if row is None:
                raise BorrowError("ไม่พบรายการยืมที่ยังไม่ได้คืน")
            book_id = row[0]
            conn.execute(
                "UPDATE borrow_log SET returned = 1 WHERE id = ?", (record_id,))
            conn.execute(
                "UPDATE books SET status = ? WHERE id = ?",
                (BOOK_STATUS_AVAILABLE, book_id))
            # reservation-aware: promote next waiter to an explicit offer
            self._expire_and_offer_next(conn, book_id)
            return book_id

        try:
            return self.transact(_tx)
        except sqlite3.IntegrityError as exc:
            raise BorrowError("ข้อมูลไม่ถูกต้อง ไม่สามารถบันทึกการคืนได้") from exc
