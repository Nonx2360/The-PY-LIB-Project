# services/reservation_service.py
from datetime import datetime

import config.settings as settings
from db.repositories.reservation_repo import ReservationRepository
from db.repositories.book_repo import BookRepository
from db.repositories.borrow_repo import BorrowRepository, BorrowError


class ReservationService:
    def __init__(self, repo=None):
        self.repo = repo or ReservationRepository()
        self.books = BookRepository()
        self.borrow = BorrowRepository()

    def _expire_sweep(self):
        # best-effort global expiry; book-specific promotion happens on access/return
        self.repo.execute(
            "UPDATE book_reservations SET status='expired' "
            "WHERE status='offered' AND expires_at IS NOT NULL AND expires_at < ?",
            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),))

    def reserve(self, book_id, member_id, actor_member_id=None, admin=False):
        """Place a FIFO reservation for a specific copy (book_id)."""
        self._expire_sweep()
        if not admin and actor_member_id is not None and int(actor_member_id) != int(member_id):
            raise ValueError("ไม่มีสิทธิ์จองให้ผู้อื่น")
        book = self.books.fetchone(
            "SELECT id, code, title, status FROM books WHERE id = ?", (book_id,))
        if not book:
            raise ValueError("ไม่พบหนังสือ")
        if book[3] == 'หยุดให้บริการ':
            raise ValueError("หนังสือเล่มนี้หยุดให้บริการ")
        # prevent duplicate active reservation by same member for same copy
        dup = self.repo.fetchone(
            "SELECT id FROM book_reservations WHERE book_id=? AND member_id=? "
            "AND status IN ('waiting','offered')", (book_id, member_id))
        if dup:
            raise ValueError("คุณมีการจองหนังสือเล่มนี้อยู่แล้ว")
        # if member currently borrowing this same copy, no point reserving
        borrowing = self.repo.fetchone(
            "SELECT 1 FROM borrow_log WHERE book_id=? AND member_id=? AND returned=0",
            (book_id, member_id))
        if borrowing:
            raise ValueError("คุณกำลังยืมหนังสือเล่มนี้อยู่")
        rid = self.repo.create(book_id, member_id)
        # if book is available and no one ahead, try to create immediate offer
        # do it transactionally by attempting promotion via a small transaction
        try:
            hours = int(settings.get("reservation_offer_hours") or 48)
            self.repo.transact(lambda conn: self.repo._offer_next_conn(conn, book_id, hours))
        except Exception:
            pass
        return rid

    def reserve_by_code(self, code, member_id, actor_member_id=None, admin=False):
        code = (code or "").strip()
        if not code:
            raise ValueError("กรุณาระบุรหัสหนังสือ")
        row = self.books.fetchone(
            "SELECT id FROM books WHERE code = ? OR lower(code)=lower(?)", (code, code))
        if not row:
            raise ValueError("ไม่พบหนังสือจากรหัสที่ระบุ")
        return self.reserve(row[0], member_id, actor_member_id=actor_member_id, admin=admin)

    def cancel(self, reservation_id, actor_member_id=None, admin=False):
        self._expire_sweep()
        book_id = self.repo.cancel(reservation_id, member_id=actor_member_id, by_admin=admin)
        # if an offer was cancelled, promote next
        try:
            hours = int(settings.get("reservation_offer_hours") or 48)
            self.repo.transact(lambda conn: self.repo._offer_next_conn(conn, book_id, hours))
        except Exception:
            pass
        return True

    def claim(self, reservation_id, actor_member_id):
        """Explicit claim: borrower (offered member) takes the book -> creates a real loan."""
        self._expire_sweep()
        row = self.repo.fetchone(
            "SELECT id, book_id, member_id, status, expires_at FROM book_reservations WHERE id=?",
            (reservation_id,))
        if not row:
            raise ValueError("ไม่พบรายการจอง")
        if int(row[2]) != int(actor_member_id):
            raise ValueError("ไม่มีสิทธิ์รับหนังสือรายการนี้")
        if row[3] != 'offered':
            raise ValueError("รายการนี้ยังไม่พร้อมให้รับ (ยังไม่ถึงคิวหรือหมดเวลา)")
        # check expiry
        if row[4] and row[4] < datetime.now().strftime("%Y-%m-%d %H:%M:%S"):
            # mark expired and try promote
            self.repo.execute("UPDATE book_reservations SET status='expired' WHERE id=?", (reservation_id,))
            try:
                hours = int(settings.get("reservation_offer_hours") or 48)
                self.repo.transact(lambda conn: self.repo._offer_next_conn(conn, row[1], hours))
            except Exception:
                pass
            raise ValueError("ข้อเสนอรับหนังสือหมดเวลาแล้ว")
        borrow_date = datetime.now().strftime("%Y-%m-%d")
        due = (datetime.now() + __import__('datetime').timedelta(days=int(settings.get("default_loan_days") or 7))).strftime("%Y-%m-%d")
        # try_borrow will fulfill the offered reservation atomically (same member)
        try:
            self.borrow.try_borrow(int(actor_member_id), int(row[1]), borrow_date, due)
        except BorrowError as exc:
            raise ValueError(str(exc)) from exc
        return True

    def list_for_member(self, member_id):
        self._expire_sweep()
        return self.repo.member_reservations(member_id)

    def list_for_book(self, book_id):
        self._expire_sweep()
        return self.repo.queue_for_book(book_id)
