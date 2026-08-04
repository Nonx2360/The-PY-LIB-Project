# db/repositories/borrow_repo.py
from .base_repo import BaseRepository

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
        query = '''
            INSERT INTO borrow_log (member_id, book_id, borrow_date, return_due, returned)
            VALUES (?, ?, ?, ?, 0)
        '''
        return self.execute(query, (member_id, book_id, borrow_date, return_due))
        
    def mark_returned(self, record_id):
        query = "UPDATE borrow_log SET returned = 1 WHERE id = ?"
        self.execute(query, (record_id,))
