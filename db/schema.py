# db/schema.py — schema creation, versioned migrations, admin bootstrap.
#
# Migration policy: instead of relying only on CREATE TABLE IF NOT EXISTS,
# every connection schema state is tracked with `PRAGMA user_version`.
#   version 1 = original legacy layout (pre-refactor tables)
#   version 2 = adds data-integrity constraints (see MIGRATION notes below)
#   Fresh DBs are created directly in the v2 layout.
# Any new schema change = bump SCHEMA_VERSION + add a migration step here.
#
# Admin bootstrap policy:
#   * No default password is ever seeded. First boot generates a strong random
#     password, prints it once and records it in db/INITIAL_ADMIN_PASSWORD.txt
#     (gitignored). Passwords are bcrypt-hashed at rest.
#   * Boot is idempotent: if any admin still authenticates with the historic
#     default 'admin123', it is replaced with a fresh random password on
#     every boot until changed.
import os
import secrets
import sqlite3
import string
from datetime import datetime

from config.constants import DB_PATH

SCHEMA_VERSION = 2

_CREDS_FILE = "INITIAL_ADMIN_PASSWORD.txt"
_KNOWN_DEFAULT_PASSWORD = "admin123"

# ------------------------------- definitions -------------------------------

_BOOKS = '''
    CREATE TABLE IF NOT EXISTS books (
        id INTEGER PRIMARY KEY,
        code TEXT,
        title TEXT,
        status TEXT
    )
'''

_BORROW_LOG = '''
    CREATE TABLE IF NOT EXISTS borrow_log (
        id INTEGER PRIMARY KEY,
        member_id INTEGER NOT NULL,
        book_id INTEGER NOT NULL,
        borrow_date TEXT,
        return_due TEXT,
        returned INTEGER NOT NULL DEFAULT 0 CHECK (returned IN (0, 1)),
        FOREIGN KEY (member_id) REFERENCES members (id),
        FOREIGN KEY (book_id) REFERENCES books (id)
    )
'''


def _tables(conn):
    rows = conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    return {r[0] for r in rows}


def _create_final_schema(conn):
    """Layout used for fresh installs (v2): includes integrity constraints."""
    conn.execute('''CREATE TABLE IF NOT EXISTS admin_users (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS members (
            id INTEGER PRIMARY KEY,
            name TEXT,
            grade TEXT,
            number TEXT,
            register_date TEXT,
            expire_date TEXT,
            qrcode_data TEXT,
            qrcode_path TEXT)''')
    conn.execute(_BOOKS)
    conn.execute(_BORROW_LOG)
    conn.execute('''CREATE TABLE IF NOT EXISTS access_log (
            id INTEGER PRIMARY KEY,
            member_id INTEGER,
            access_time DATETIME,
            action TEXT,
            FOREIGN KEY (member_id) REFERENCES members (id))''')
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_books_code ON books(code)")


def _create_base_schema(conn):
    """v1 layout — creates the legacy tables only if they don't exist yet."""
    conn.execute('''CREATE TABLE IF NOT EXISTS admin_users (
            username TEXT PRIMARY KEY,
            password TEXT
        )''')
    conn.execute('''CREATE TABLE IF NOT EXISTS members (
            id INTEGER PRIMARY KEY,
            name TEXT,
            grade TEXT,
            number TEXT,
            register_date TEXT,
            expire_date TEXT,
            qrcode_data TEXT,
            qrcode_path TEXT
        )''')
    conn.execute(_BOOKS)
    conn.execute('''CREATE TABLE IF NOT EXISTS borrow_log (
            id INTEGER PRIMARY KEY,
            member_id INTEGER,
            book_id INTEGER,
            borrow_date TEXT,
            return_due TEXT,
            returned INTEGER,
            FOREIGN KEY (member_id) REFERENCES members (id),
            FOREIGN KEY (book_id) REFERENCES books (id)
        )''')
    conn.execute('''CREATE TABLE IF NOT EXISTS access_log (
            id INTEGER PRIMARY KEY,
            member_id INTEGER,
            access_time DATETIME,
            action TEXT,
            FOREIGN KEY (member_id) REFERENCES members (id)
        )''')


def _has_base_tables(conn):
    return "books" in _tables(conn) and "members" in _tables(conn)


# ------------------------------- migrations -------------------------------

def _migrate_v2(conn, log):
    """v1 -> v2: enforce unique book codes and valid loan states.

    * books: de-duplicate repeated codes (suffix `-<id>` for later rows),
      then add a UNIQUE index so data integrity is DB-enforced.
    * borrow_log: rebuild the table so `returned` is NOT NULL CHECK (0|1),
      member_id/book_id are NOT NULL and foreign keys are declared.
      Non-0/1 or NULL legacy `returned` values are coerced to 1 (returned).
      Runs inside one transaction with FK enforcement off (rebuild pattern).
    * books.status is a free display string (Thai labels) by design.
    """
    was_fk_on = conn.execute("PRAGMA foreign_keys").fetchone()[0]
    conn.isolation_level = None  # manual transaction control
    try:
        conn.execute("PRAGMA foreign_keys=OFF")
        conn.execute("BEGIN IMMEDIATE")

        # 1) unique book codes
        dup_codes = conn.execute(
            "SELECT code FROM books GROUP BY code HAVING COUNT(*) > 1").fetchall()
        for (code,) in dup_codes:
            ids = [r[0] for r in conn.execute(
                "SELECT id FROM books WHERE code = ? ORDER BY id", (code,))]
            for dup_id in ids[1:]:
                conn.execute("UPDATE books SET code = ? WHERE id = ?",
                             (f"{code}-{dup_id}", dup_id))
            if log:
                log(f"books: duplicate code '{code}' — renumbered copies to "
                    f"'{code}-{ids[1:]}' pattern (kept id {ids[0]})")
        conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_books_code ON books(code)")

        # 2) borrow_log constraints via table rebuild
        conn.execute("ALTER TABLE borrow_log RENAME TO borrow_log_legacy")
        conn.execute(_BORROW_LOG)
        conn.execute('''
            INSERT INTO borrow_log (id, member_id, book_id, borrow_date, return_due, returned)
            SELECT id, member_id, book_id, borrow_date, return_due,
                   CASE WHEN returned IN (0, 1) THEN returned ELSE 1 END
            FROM borrow_log_legacy
        ''')
        conn.execute("DROP TABLE borrow_log_legacy")

        conn.execute("COMMIT")
    except Exception:
        conn.execute("ROLLBACK")
        raise
    finally:
        conn.execute(f"PRAGMA foreign_keys={1 if was_fk_on else 0}")


def _apply_migrations(conn, log=None):
    version = conn.execute("PRAGMA user_version").fetchone()[0]
    if version >= SCHEMA_VERSION:
        return
    if version == 0:
        if _has_base_tables(conn):
            # pre-migration install: mark baseline, then upgrade
            conn.execute("PRAGMA user_version = 1")
            version = 1
        else:
            _create_final_schema(conn)
            conn.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
            if log:
                log(f"db: created fresh schema (v{SCHEMA_VERSION})")
            return
    if version == 1:
        _migrate_v2(conn, log)
        conn.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
        if log:
            log(f"db: migrated schema v1 -> v{SCHEMA_VERSION} (constraints + unique codes)")
    # future versions land here


# ------------------------------- admin bootstrap -------------------------------
# (imported lazily keeps bcrypt/crypto import cost off DB-only usage)

def _random_password(length=14):
    alphabet = string.ascii_letters + string.digits
    rng = secrets.SystemRandom()
    pw = [
        rng.choice(string.ascii_lowercase),
        rng.choice(string.ascii_uppercase),
        rng.choice(string.digits),
    ]
    pw += [rng.choice(alphabet) for _ in range(length - len(pw))]
    rng.shuffle(pw)
    return "".join(pw)


def _creds_path():
    directory = os.path.dirname(DB_PATH) or "."
    return os.path.join(directory, _CREDS_FILE)


def _notify_password(username, password, reason):
    """Print a one-time password to the console and record it on disk."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("=" * 60)
    print(f"NOTICE - {reason}")
    print(f"  admin account : {username}")
    print(f"  password      : {password}")
    print(f"  saved to      : {_creds_path()}  (gitignored - back it up)")
    print("=" * 60)
    directory = os.path.dirname(DB_PATH) or "."
    os.makedirs(directory, exist_ok=True)
    with open(_creds_path(), "a", encoding="utf-8") as f:
        f.write(f"{ts} - {username}: {password}  ({reason})\n")


def _bootstrap_admins(conn, log=None):
    from core.security import hash_password, verify_password

    rows = conn.execute("SELECT username, password FROM admin_users").fetchall()
    known_default_retired = False
    for username, pwd in rows:
        keep = True
        new_password = None
        if pwd == _KNOWN_DEFAULT_PASSWORD:          # legacy plain-text default
            keep, new_password = False, _random_password()
        else:
            try:
                if verify_password(_KNOWN_DEFAULT_PASSWORD, pwd or ""):
                    keep, new_password = False, _random_password()
            except Exception:
                pass
        if not keep:
            hashed = hash_password(new_password)
            conn.execute("UPDATE admin_users SET password = ? WHERE username = ?",
                         (hashed, username))
            _notify_password(username, new_password,
                             "shared default password removed - "
                             f"{datetime.now().strftime('%Y-%m-%d %H:%M')}")
            known_default_retired = True
            continue
        # normalize remaining legacy plain-text passwords to bcrypt
        if pwd and not (pwd.startswith('$2a$') or pwd.startswith('$2b$')
                        or pwd.startswith('$2y$')):
            conn.execute("UPDATE admin_users SET password = ? WHERE username = ?",
                         (hash_password(pwd), username))
            if log:
                log(f"admin_users: bcrypt-normalized legacy plain-text password for '{username}'")

    if not rows:
        new_password = _random_password()
        conn.execute("INSERT INTO admin_users (username, password) VALUES (?, ?)",
                     ("admin", hash_password(new_password)))
        _notify_password("admin", new_password,
                         f"first boot - initial admin password generated "
                         f"{datetime.now().strftime('%Y-%m-%d %H:%M')}")
        if log:
            log("admin_users: seeded first admin with a generated password (no default exists)")

    if known_default_retired and log:
        log("admin_users: retired historic 'admin123' default credential")


# ------------------------------- entry point -------------------------------

def init_db(log=None):
    """Create/upgrade the schema and bootstrap admin accounts.

    Safe to call on every boot; migrations and bootstrap are idempotent.
    """
    directory = os.path.dirname(DB_PATH)
    if directory and not os.path.exists(directory):
        os.makedirs(directory, exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("PRAGMA busy_timeout=5000")
        _apply_migrations(conn, log=log)
        _bootstrap_admins(conn, log=log)
        conn.commit()
    finally:
        conn.close()


def seed_admin():
    """Deprecated shim: ran legacy plain-text seeding before migrations existed.

    Retained for API compatibility — all bootstrap logic now lives in
    init_db()/_bootstrap_admins().
    """
    conn = sqlite3.connect(DB_PATH)
    try:
        _bootstrap_admins(conn)
        conn.commit()
    finally:
        conn.close()
