# db/repositories/book_repo.py
from config.constants import BOOK_STATUS_AVAILABLE
from .base_repo import BaseRepository

# Loose safety caps for bulk sources (the web layer restricts further).
_CODE_MAX_LEN = 64
_TITLE_MAX_LEN = 200

class BookRepository(BaseRepository):
    def get_all(self):
        return self.fetchall("SELECT * FROM books")

    def get_by_code(self, code):
        return self.fetchone("SELECT * FROM books WHERE code = ?", (code,))

    def add(self, code, title, status=BOOK_STATUS_AVAILABLE):
        query = "INSERT INTO books (code, title, status) VALUES (?, ?, ?)"
        return self.execute(query, (code, title, status))

    def add_many(self, rows):
        """Insert many books in ONE transaction (bulk-import friendly).

        `rows` is an iterable of (code, title) string pairs, already cleaned
        by the caller. Skips are never an error:
          * repeated codes INSIDE the batch keep only the first occurrence,
          * codes already stored in books are silently skipped (the UNIQUE
            index on books.code enforces it; INSERT OR IGNORE).
        Returns (inserted, skipped):
          inserted — rows actually written;
          skipped  — intra-file duplicates + blank/partial rows + codes that
                     already existed in the table.
        Everything lands in a single BEGIN IMMEDIATE -> commit, so a 1000-row
        import is one fsync, not a thousand.
        """
        clean, seen = [], set()
        skipped = 0
        for raw_code, raw_title in rows:
            code = (raw_code or "").strip()[:_CODE_MAX_LEN]
            title = (raw_title or "").strip()[:_TITLE_MAX_LEN]
            if not code or not title:
                skipped += 1
                continue
            if code in seen:
                skipped += 1
                continue
            seen.add(code)
            clean.append((code, title))

        if not clean:
            return 0, skipped

        def _insert(conn):
            cursor = conn.cursor()
            cursor.executemany(
                "INSERT OR IGNORE INTO books (code, title, status) VALUES (?, ?, ?)",
                [(code, title, BOOK_STATUS_AVAILABLE) for code, title in clean])
            return cursor.rowcount

        inserted = self.transact(_insert)
        return inserted, skipped + (len(clean) - inserted)

    def update_status(self, book_id, status):
        query = "UPDATE books SET status = ? WHERE id = ?"
        self.execute(query, (status, book_id))

    # ---------------- bulk management ----------------

    def bulk_retire(self, ids):
        """Withdraw several available copies; already-borrowed/retired skip.

        Returns how many rows actually changed. Uses one atomic UPDATE so
        a batch application (web admin selects 50 rows) is one commit.
        """
        ids = [i for i in ids if i]
        if not ids:
            return 0
        marks = ",".join("?" * len(ids))

        def _retire(conn):
            cur = conn.execute(
                f"UPDATE books SET status = ? WHERE status = ? "
                f"AND id IN ({marks})",
                ("หยุดให้บริการ", "ว่าง", *ids))
            return cur.rowcount

        return self.transact(_retire)

    def bulk_restore(self, ids):
        """Bring retired copies back; everything else is left untouched."""
        ids = [i for i in ids if i]
        if not ids:
            return 0
        marks = ",".join("?" * len(ids))

        def _restore(conn):
            cur = conn.execute(
                f"UPDATE books SET status = ? WHERE status = ? "
                f"AND id IN ({marks})",
                ("ว่าง", "หยุดให้บริการ", *ids))
            return cur.rowcount

        return self.transact(_restore)

    def delete_many(self, ids):
        """Delete several books in ONE transaction, skipping the unsafe.

        Semantics per book:
          * หยุดให้บริการ  -> hospital purge: erased together with its
            borrow_log rows (the sanctioned two-step retire->purge flow);
          * other statuses -> hard delete ONLY when the copy has no borrow
            history (foreign key would block it anyway);
          * everything else (borrowed, has history) -> skipped.
        Returns (deleted, skipped).
        """
        ids = [i for i in ids if i]
        if not ids:
            return 0, 0

        def _del(conn):
            deleted = skipped = 0
            for bid in ids:
                row = conn.execute(
                    "SELECT status FROM books WHERE id=?", (bid,)).fetchone()
                if not row:
                    skipped += 1
                    continue
                has_log = conn.execute(
                    "SELECT 1 FROM borrow_log WHERE book_id=? LIMIT 1",
                    (bid,)).fetchone()
                if row[0] == "หยุดให้บริการ":
                    conn.execute(
                        "DELETE FROM borrow_log WHERE book_id=?", (bid,))
                    conn.execute("DELETE FROM books WHERE id=?", (bid,))
                    deleted += 1
                elif has_log:
                    skipped += 1
                else:
                    conn.execute("DELETE FROM books WHERE id=?", (bid,))
                    deleted += 1
            return deleted, skipped

        return self.transact(_del)

    def delete(self, book_id):
        self.execute("DELETE FROM books WHERE id = ?", (book_id,))
