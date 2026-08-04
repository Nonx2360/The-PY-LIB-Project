# db/repositories/book_repo.py
from .base_repo import BaseRepository

class BookRepository(BaseRepository):
    def get_all(self):
        return self.fetchall("SELECT * FROM books")
        
    def get_by_code(self, code):
        return self.fetchone("SELECT * FROM books WHERE code = ?", (code,))
        
    def add(self, code, title, status='Available'):
        query = "INSERT INTO books (code, title, status) VALUES (?, ?, ?)"
        return self.execute(query, (code, title, status))
        
    def update_status(self, book_id, status):
        query = "UPDATE books SET status = ? WHERE id = ?"
        self.execute(query, (status, book_id))
        
    def delete(self, book_id):
        self.execute("DELETE FROM books WHERE id = ?", (book_id,))
