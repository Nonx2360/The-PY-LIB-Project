# db/repositories/base_repo.py
import sqlite3
import threading
from config.constants import DB_PATH

_BUSY_TIMEOUT_MS = 5000


class BaseRepository:
    _lock = threading.Lock()

    def __init__(self, db_path=None):
        # db_path is resolved from the module-level DB_PATH at *use* time
        # unless an explicit path is given. Tests repoint the module global
        # (tests/_harness.py, test_*_repo.py); capturing the value at
        # construction would bind repos to whichever path was active then.
        self._explicit_path = db_path

    @property
    def db_path(self):
        return self._explicit_path or DB_PATH

    def get_connection(self):
        # check_same_thread=False because access is serialized by _lock;
        # isolation_level=None -> explicit transaction control (autocommit
        # otherwise), so transact() can manage BEGIN/COMMIT/ROLLBACK itself.
        conn = sqlite3.connect(
            self.db_path, check_same_thread=False, isolation_level=None)
        # Enforce declared foreign keys (SQLite default is OFF) like borrow_log
        # -> members/books. Also tolerate transient multi-process locks.
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute(f"PRAGMA busy_timeout={_BUSY_TIMEOUT_MS}")
        return conn

    def execute(self, query, params=()):
        with self._lock:
            conn = self.get_connection()
            try:
                cursor = conn.cursor()
                cursor.execute(query, params)
                conn.commit()
                return cursor.lastrowid
            finally:
                # close promptly: leaked sqlite handles (Windows locks files)
                conn.close()

    def fetchall(self, query, params=()):
        with self._lock:
            conn = self.get_connection()
            try:
                cursor = conn.cursor()
                cursor.execute(query, params)
                return cursor.fetchall()
            finally:
                conn.close()

    def fetchone(self, query, params=()):
        with self._lock:
            conn = self.get_connection()
            try:
                cursor = conn.cursor()
                cursor.execute(query, params)
                return cursor.fetchone()
            finally:
                conn.close()

    def transact(self, fn):
        """Run fn(conn) as one atomic transaction.

        The connection is passed in — the caller MUST use `conn` directly and
        NOT call self.execute()/fetch*() (they re-acquire the class lock,
        which would deadlock). Any exception rolls the whole transaction back.
        """
        with self._lock:
            conn = self.get_connection()
            try:
                conn.execute("BEGIN IMMEDIATE")
                result = fn(conn)
                conn.commit()
                return result
            except Exception:
                try:
                    conn.rollback()
                except Exception:
                    pass
                raise
            finally:
                conn.close()
