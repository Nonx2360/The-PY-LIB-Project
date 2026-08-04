# db/repositories/base_repo.py
import sqlite3
import threading
from config.constants import DB_PATH

class BaseRepository:
    _lock = threading.Lock()
    
    def __init__(self):
        self.db_path = DB_PATH
        
    def get_connection(self):
        # Using check_same_thread=False since we use a lock
        return sqlite3.connect(self.db_path, check_same_thread=False)
        
    def execute(self, query, params=()):
        with self._lock:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                conn.commit()
                return cursor.lastrowid
                
    def fetchall(self, query, params=()):
        with self._lock:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                return cursor.fetchall()

    def fetchone(self, query, params=()):
        with self._lock:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                return cursor.fetchone()
