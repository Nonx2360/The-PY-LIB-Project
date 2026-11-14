# db/schema.py — schema creation, versioned migrations, admin bootstrap.
#
# Migration policy: instead of relying only on CREATE TABLE IF NOT EXISTS,
# every connection schema state is tracked with `PRAGMA user_version`.
#   version 1 = original legacy layout (pre-refactor tables)
#   version 2 = adds data-integrity constraints (see MIGRATION notes below)
#   version 3 = bulk-import era indexes on borrow_log (book_id/member_id) —
#               books-page history counts + lists survive 1000+ row catalogs
#   version 4 = member portal era: members.pin_hash (bcrypt, NULL until the
#               member activates via their card or an admin-issued one-time
#               code) + member_activation_codes (hashed one-time codes so
#               the plaintext is never at rest) + supporting indexes.
#   Fresh DBs are created directly in the current layout.
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

SCHEMA_VERSION = 7

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
    """Layout used for fresh installs: current schema including migrations."""
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
            qrcode_path TEXT,
            pin_hash TEXT,
            status TEXT NOT NULL DEFAULT 'active')''')
    conn.execute(_BOOKS)
    conn.execute(_BORROW_LOG)
    conn.execute('''CREATE TABLE IF NOT EXISTS access_log (
            id INTEGER PRIMARY KEY,
            member_id INTEGER,
            access_time DATETIME,
            action TEXT,
            FOREIGN KEY (member_id) REFERENCES members (id))''')
    conn.execute('''CREATE TABLE IF NOT EXISTS member_activation_codes (
            id INTEGER PRIMARY KEY,
            member_id INTEGER NOT NULL,
            code_hash TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL,
            expires_at TEXT NOT NULL,
            used_at TEXT,
            FOREIGN KEY (member_id) REFERENCES members (id))''')
    _create_academic_year_tables(conn)
    _create_inventory_audit_tables(conn)
    _create_reservation_tables(conn)
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_books_code ON books(code)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_activation_member "
                 "ON member_activation_codes(member_id)")
    _create_bulk_indexes(conn)


def _create_academic_year_tables(conn):
    """v5: explicit academic years, member placements, and migration audit."""
    conn.execute('''CREATE TABLE IF NOT EXISTS academic_years (
            id INTEGER PRIMARY KEY,
            label TEXT NOT NULL UNIQUE,
            start_date TEXT,
            end_date TEXT,
            is_active INTEGER NOT NULL DEFAULT 0 CHECK (is_active IN (0, 1)),
            created_at TEXT NOT NULL)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS member_placements (
            id INTEGER PRIMARY KEY,
            member_id INTEGER NOT NULL,
            academic_year_id INTEGER,
            grade TEXT,
            number TEXT,
            status TEXT NOT NULL DEFAULT 'current',
            source TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY (member_id) REFERENCES members (id),
            FOREIGN KEY (academic_year_id) REFERENCES academic_years (id))''')
    conn.execute('''CREATE TABLE IF NOT EXISTS academic_year_migrations (
            id INTEGER PRIMARY KEY,
            source_year_id INTEGER,
            target_year_id INTEGER,
            source_label TEXT,
            target_label TEXT,
            actor TEXT,
            status TEXT NOT NULL,
            summary_json TEXT,
            created_at TEXT NOT NULL,
            finished_at TEXT,
            notes TEXT,
            FOREIGN KEY (source_year_id) REFERENCES academic_years (id),
            FOREIGN KEY (target_year_id) REFERENCES academic_years (id))''')
    conn.execute('''CREATE TABLE IF NOT EXISTS academic_year_migration_items (
            id INTEGER PRIMARY KEY,
            migration_id INTEGER NOT NULL,
            member_id INTEGER NOT NULL,
            action TEXT NOT NULL,
            reason TEXT,
            before_json TEXT,
            after_json TEXT,
            outcome TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY (migration_id) REFERENCES academic_year_migrations (id),
            FOREIGN KEY (member_id) REFERENCES members (id))''')
    conn.execute("CREATE INDEX IF NOT EXISTS idx_member_placements_member "
                 "ON member_placements(member_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_member_placements_year "
                 "ON member_placements(academic_year_id)")
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_member_placements_current "
                 "ON member_placements(member_id) WHERE status = 'current'")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_migration_items_run "
                 "ON academic_year_migration_items(migration_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_migration_items_member "
                 "ON academic_year_migration_items(member_id)")


def _create_inventory_audit_tables(conn):
    """v6: inventory audit sessions, scans, and discrepancy resolutions."""
    conn.execute('''CREATE TABLE IF NOT EXISTS inventory_audits (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            location TEXT,
            scope TEXT NOT NULL DEFAULT 'all',
            status TEXT NOT NULL DEFAULT 'in_progress'
                CHECK (status IN ('in_progress','paused','completed','cancelled')),
            created_by TEXT,
            notes TEXT,
            expected_count INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            completed_at TEXT)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS inventory_audit_scans (
            id INTEGER PRIMARY KEY,
            audit_id INTEGER NOT NULL,
            book_id INTEGER,
            raw_value TEXT,
            outcome TEXT NOT NULL,
            scanned_by TEXT,
            scanned_at TEXT NOT NULL,
            FOREIGN KEY (audit_id) REFERENCES inventory_audits (id),
            FOREIGN KEY (book_id) REFERENCES books (id))''')
    conn.execute('''CREATE TABLE IF NOT EXISTS inventory_audit_resolutions (
            id INTEGER PRIMARY KEY,
            audit_id INTEGER NOT NULL,
            book_id INTEGER,
            scan_id INTEGER,
            resolution TEXT NOT NULL,
            note TEXT,
            actor TEXT,
            resolved_at TEXT NOT NULL,
            FOREIGN KEY (audit_id) REFERENCES inventory_audits (id),
            FOREIGN KEY (book_id) REFERENCES books (id),
            FOREIGN KEY (scan_id) REFERENCES inventory_audit_scans (id))''')
    conn.execute("CREATE INDEX IF NOT EXISTS idx_inv_scans_audit "
                 "ON inventory_audit_scans(audit_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_inv_scans_book "
                 "ON inventory_audit_scans(audit_id, book_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_inv_res_audit "
                 "ON inventory_audit_resolutions(audit_id)")
    # count each real book at most once per audit
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_inv_scan_unique_book "
                 "ON inventory_audit_scans(audit_id, book_id) "
                 "WHERE book_id IS NOT NULL AND outcome = 'found'")


def _create_reservation_tables(conn):
    """v7: copy-level reservations with deterministic FIFO offers."""
    conn.execute('''CREATE TABLE IF NOT EXISTS book_reservations (
            id INTEGER PRIMARY KEY,
            book_id INTEGER NOT NULL,
            member_id INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'waiting'
                CHECK (status IN ('waiting','offered','fulfilled','cancelled','expired')),
            created_at TEXT NOT NULL,
            offered_at TEXT,
            expires_at TEXT,
            fulfilled_at TEXT,
            cancelled_at TEXT,
            note TEXT,
            FOREIGN KEY (book_id) REFERENCES books (id),
            FOREIGN KEY (member_id) REFERENCES members (id))''')
    conn.execute("CREATE INDEX IF NOT EXISTS idx_res_book_status "
                 "ON book_reservations(book_id, status, id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_res_member_status "
                 "ON book_reservations(member_id, status, id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_res_expires "
                 "ON book_reservations(status, expires_at)")


def _migrate_v7(conn, log):
    """v5/6 -> v7: book reservations (idempotent)."""
    _create_reservation_tables(conn)
    if log:
        log("db: migrated schema v6 -> v7 (book reservations)")


def _create_bulk_indexes(conn):
    """v3 indexes: keep the books page fast when the catalog is 1000+ rows.

    * borrow_log(book_id) — the per-book history counts on /books and the
      member-delete guard are correlated subqueries; without the index each
      one is a full borrow_log scan.
    * borrow_log(member_id) — borrow/return/member-delete lookups by member.
    """
    conn.execute("CREATE INDEX IF NOT EXISTS idx_borrow_log_book "
                 "ON borrow_log(book_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_borrow_log_member "
                 "ON borrow_log(member_id)")


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


def _migrate_v3(conn, log):
    """v2 -> v3: bulk-import era indexes.

    Index-only change (no table rebuild, no data rewrite), so an existing
    catalog with 1000+ books migrates in milliseconds. Idempotent.
    """
    conn.isolation_level = None
    try:
        conn.execute("BEGIN IMMEDIATE")
        _create_bulk_indexes(conn)
        conn.execute("COMMIT")
    except Exception:
        conn.execute("ROLLBACK")
        raise


def _migrate_v4(conn, log):
    """v3 -> v4: member portal era (rotating QR identity + member login).

    * members.pin_hash — bcrypt hash of the member's PIN; NULL means the
      member has not activated their portal login yet (activation requires
      presenting the physical card or an admin-issued one-time code).
    * member_activation_codes — one-time activation codes issued by the
      admin from the members page; only the SHA-256 hash is stored so the
      plaintext code is never at rest.
    Idempotent: column/table existence is checked before altering.
    """
    cols = [r[1] for r in conn.execute("PRAGMA table_info(members)").fetchall()]
    if "pin_hash" not in cols:
        conn.execute("ALTER TABLE members ADD COLUMN pin_hash TEXT")
    conn.execute('''CREATE TABLE IF NOT EXISTS member_activation_codes (
            id INTEGER PRIMARY KEY,
            member_id INTEGER NOT NULL,
            code_hash TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL,
            expires_at TEXT NOT NULL,
            used_at TEXT,
            FOREIGN KEY (member_id) REFERENCES members (id))''')
    conn.execute("CREATE INDEX IF NOT EXISTS idx_activation_member "
                 "ON member_activation_codes(member_id)")
    if log:
        log("db: migrated schema v3 -> v4 (member portal: pin_hash + "
            "activation codes)")


def _migrate_v5(conn, log):
    """v4 -> v5: academic-year placements and migration audit trail.

    Adds members.status (default 'active') without touching historical
    borrow/access rows, creates academic year / placement / migration
    tables, and seeds a baseline "current" placement for every existing
    member so promotions can be tracked going forward. Idempotent.
    """
    cols = [r[1] for r in conn.execute("PRAGMA table_info(members)").fetchall()]
    if "status" not in cols:
        conn.execute("ALTER TABLE members ADD COLUMN status TEXT NOT NULL DEFAULT 'active'")
    _create_academic_year_tables(conn)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    baseline = conn.execute(
        "SELECT id FROM academic_years WHERE label = ?",
        ("ปัจจุบัน (ก่อนระบบปีการศึกษา)",)).fetchone()
    if baseline is None:
        cur = conn.execute(
            "INSERT INTO academic_years (label, start_date, end_date, is_active, created_at) "
            "VALUES (?, NULL, NULL, 1, ?)",
            ("ปัจจุบัน (ก่อนระบบปีการศึกษา)", now))
        baseline_id = cur.lastrowid
    else:
        baseline_id = baseline[0]
    # one current placement per member, derived from current member fields
    members = conn.execute(
        "SELECT id, grade, number FROM members").fetchall()
    for mid, grade, number in members:
        has_current = conn.execute(
            "SELECT 1 FROM member_placements WHERE member_id = ? AND status = 'current'",
            (mid,)).fetchone()
        if not has_current:
            conn.execute(
                "INSERT INTO member_placements "
                "(member_id, academic_year_id, grade, number, status, source, created_at) "
                "VALUES (?, ?, ?, ?, 'current', 'migration_v5_baseline', ?)",
                (mid, baseline_id, grade, number, now))
    if log:
        log("db: migrated schema v4 -> v5 (academic years, placements, migration audit)")


def _migrate_v6(conn, log):
    """v5 -> v6: inventory audit sessions, scans, resolutions (idempotent)."""
    _create_inventory_audit_tables(conn)
    if log:
        log("db: migrated schema v5 -> v6 (inventory audit)")


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
        version = 2
        conn.execute("PRAGMA user_version = 2")
        if log:
            log("db: migrated schema v1 -> v2 (constraints + unique codes)")
    if version == 2:
        _migrate_v3(conn, log)
        version = 3
        conn.execute("PRAGMA user_version = 3")
        if log:
            log("db: migrated schema v2 -> v3 (bulk-import indexes)")
    if version == 3:
        _migrate_v4(conn, log)
        version = 4
        conn.execute("PRAGMA user_version = 4")
        if log:
            log("db: migrated schema v3 -> v4 (member portal: pin_hash + activation codes)")
    if version == 4:
        _migrate_v5(conn, log)
        version = 5
        conn.execute("PRAGMA user_version = 5")
        if log:
            log("db: migrated schema v4 -> v5 (academic-year placements + migration audit)")
    if version == 5:
        _migrate_v6(conn, log)
        version = 6
        conn.execute("PRAGMA user_version = 6")
        if log:
            log("db: migrated schema v5 -> v6 (inventory audit)")
    if version == 6:
        _migrate_v7(conn, log)
        version = 7
        conn.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
        if log:
            log(f"db: migrated schema v6 -> v{SCHEMA_VERSION} (book reservations)")
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
