# db/repositories/borrow_repo.py
import sqlite3

from config import settings
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

    def try_borrow(self, member_id, book_id, borrow_date, return_due):
        """Atomically lend one book.

        Guarantees, enforced inside a single IMMEDIATE transaction:
          * the book exists and is 'ว่าง' — its status flips to 'ยืมแล้ว'
            only via a conditional UPDATE (rowcount == 1), so concurrent
            borrowers cannot double-lend;
          * member has no overdue loan and is below max_books_per_member;
          * on any failure (rule violated or FK/integrity error) the
            transaction rolls back — no partial state (book status or log).

        Raises BorrowError with a Thai, user-facing message.
        """
        def _tx(conn):
            row = conn.execute(
                "SELECT status FROM books WHERE id = ?", (book_id,)).fetchone()
            if row is None:
                raise BorrowError("ไม่พบหนังสือในระบบ")
            if row[0] != "ว่าง":
                raise BorrowError("หนังสือเล่มนี้ไม่พร้อมให้ยืม (ถูกยืมอยู่)")

            cur = conn.execute(
                "UPDATE books SET status = 'ยืมแล้ว' WHERE id = ? AND status = 'ว่าง'",
                (book_id,))
            if cur.rowcount != 1:
                raise BorrowError("หนังสือเล่มนี้ไม่พร้อมให้ยืม (ถูกยืมอยู่)")

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
          * borrow_log.returned=1 and books.status='ว่าง' applied together.
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
                "UPDATE books SET status = 'ว่าง' WHERE id = ?", (book_id,))
            return book_id

        try:
            return self.transact(_tx)
        except sqlite3.IntegrityError as exc:
            raise BorrowError("ข้อมูลไม่ถูกต้อง ไม่สามารถบันทึกการคืนได้") from exc
