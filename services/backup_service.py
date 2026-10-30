# services/backup_service.py — SQLite backup / verification / restore
#
# Backups are created with the SQLite online-backup API (src.backup(dst)) so a
# consistent snapshot is produced even while the app/app-servers are running,
# rather than a possibly-torn file copy of a .db file.
import gc
import os
import sqlite3
import time
from datetime import datetime

from config.constants import DB_PATH

BACKUP_DIR = os.path.join("backups")


def create_backup(db_path=DB_PATH, dest_path=None, backup_dir=BACKUP_DIR):
    """Create a consistent .db backup; returns the backup file path.

    Default destination: backups/library_backup_YYYYMMDD_HHMMSS.db
    """
    if dest_path:
        dest = dest_path
    else:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        os.makedirs(backup_dir, exist_ok=True)
        dest = os.path.join(backup_dir, f"library_backup_{ts}.db")

    if os.path.abspath(dest) == os.path.abspath(db_path):
        raise ValueError("backup destination must differ from the live database")

    src = sqlite3.connect(db_path)
    try:
        dst = sqlite3.connect(dest)
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()
    return dest


def verify_backup(backup_path):
    """Check a backup file: integrity_check + foreign_key_check on a RO conn.

    Returns dict(ok=bool, integrity=str, fk_violations=list).
    Unreadable/corrupt files report ok=False instead of raising.
    """
    try:
        dst = sqlite3.connect(f"file:{backup_path}?mode=ro", uri=True)
        try:
            integrity = dst.execute("PRAGMA integrity_check").fetchone()[0]
            fk_violations = dst.execute("PRAGMA foreign_key_check").fetchall()
            return {
                "ok": integrity == "ok" and not fk_violations,
                "integrity": integrity,
                "fk_violations": fk_violations,
            }
        finally:
            dst.close()
    except sqlite3.DatabaseError as exc:
        return {"ok": False, "integrity": f"unreadable backup ({exc})", "fk_violations": [("corrupt", "", "corrupt")]}


def _atomic_replace(src, dest, attempts=5):
    """os.replace with a short backoff retry.

    On Windows, a freshly written/written-to .db can transiently stay locked
    for a moment (antivirus/indexer scanning, garbage-collected SQLite
    handles). Retrying after gc.collect() avoids spurious 'Access is denied'
    on an otherwise-valid atomic swap.
    """
    for i in range(attempts):
        try:
            os.replace(src, dest)
            return
        except PermissionError:
            if i == attempts - 1:
                raise
            gc.collect()  # drop any Collectible sqlite3/other file handles now
            time.sleep(0.05 * (2 ** i))


def restore_backup(backup_path, db_path=DB_PATH, db_dir=None):
    """Replace the live DB with a verified backup copy.

    Safety: refuses to restore over an identical path (use an explicit copy),
    refuses missing/failed-verification backups, then copies atomically
    (write temp file + os.replace). The Windows/POSIX rename is atomic on
    both platforms for same-volume files.
    """
    if os.path.abspath(backup_path) == os.path.abspath(db_path):
        raise ValueError("cannot restore a backup onto itself")

    report = verify_backup(backup_path)
    if not report["ok"]:
        raise ValueError(f"backup failed verification: {report}")

    directory = db_dir or os.path.dirname(db_path) or "."
    os.makedirs(directory, exist_ok=True)
    tmp = db_path + ".restoring"
    with open(backup_path, "rb") as src, open(tmp, "wb") as out:
        while True:
            chunk = src.read(1024 * 256)
            if not chunk:
                break
            out.write(chunk)
    _atomic_replace(tmp, db_path)  # atomic on same volume
    return {"restored_from": backup_path, "verification": report}
