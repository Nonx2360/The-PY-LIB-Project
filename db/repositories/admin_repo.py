# db/repositories/admin_repo.py — admin_users access + credential checks
#
# Single sanctioned path to the admin table (the desktop login view and the
# web login route both come through here; nothing opens its own
# sqlite3.connect to admin_users anymore).
import bcrypt

from core.security import verify_password
from .base_repo import BaseRepository

# Timing-equalization hash for unknown usernames: a missing user still pays
# one bcrypt compare, so response latency cannot be used to enumerate
# which usernames exist.
_DUMMY_HASH = bcrypt.hashpw(b"pylib-timing-equalizer", bcrypt.gensalt())


class AdminRepository(BaseRepository):
    def get_all(self):
        return self.fetchall("SELECT username, password FROM admin_users")

    def get_by_username(self, username):
        return self.fetchone(
            "SELECT username, password FROM admin_users WHERE username = ?",
            (username,))

    def add(self, username, password_hash):
        return self.execute(
            "INSERT INTO admin_users (username, password) VALUES (?, ?)",
            (username, password_hash))

    def update_password(self, username, password_hash):
        self.execute(
            "UPDATE admin_users SET password = ? WHERE username = ?",
            (password_hash, username))

    def delete(self, username):
        self.execute("DELETE FROM admin_users WHERE username = ?", (username,))

    def verify_credentials(self, username, password):
        """Return True only for an existing user with the matching password.

        Unknown usernames run an identical bcrypt compare against a dummy
        hash (timing equalization) and can never verify.
        """
        row = self.get_by_username(username)
        hashed = row[1] if row else _DUMMY_HASH
        ok = verify_password(password, hashed)
        return bool(row) and ok
