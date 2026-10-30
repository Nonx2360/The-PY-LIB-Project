# tests/_harness.py — shared temp-DB isolation for the security suite
#
# Not discovered by unittest (default pattern is test*.py).
# Usage:
#     import harness; harness.use_temp_db()  (inside setUpClass)
#     ... patch module DB_PATH bindings BEFORE importing web.app / repos ...
import os
import sqlite3
import tempfile


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def query(path, sql, params=()):
    """Run a read-only statement, ALWAYS closing the connection.

    Windows note: `with sqlite3.connect(...) as conn:` commits but does NOT
    close; a lingering handle makes later os.remove/os.replace fail.
    """
    conn = sqlite3.connect(path)
    try:
        return conn.execute(sql, params).fetchall()
    finally:
        conn.close()


class TempEnv:
    """Isolated temp directory for db + secrets for one test class."""

    def __init__(self):
        self.dir = tempfile.mkdtemp(prefix="pylib_sec_")
        self.db_path = os.path.join(self.dir, "test.db")
        self.env_path = os.path.join(self.dir, ".env")

    def apply(self):
        os.environ["PYLIB_ENV_FILE"] = self.env_path

    def patch(self):
        """Point schema/base_repo modules at the temp DB."""
        import db.schema as schema_mod
        import db.repositories.base_repo as base_mod
        self._orig_schema = schema_mod.DB_PATH
        self._orig_base = base_mod.DB_PATH
        schema_mod.DB_PATH = self.db_path
        base_mod.DB_PATH = self.db_path

    def restore(self):
        import db.schema as schema_mod
        import db.repositories.base_repo as base_mod
        schema_mod.DB_PATH = self._orig_schema
        base_mod.DB_PATH = self._orig_base
        os.environ.pop("PYLIB_ENV_FILE", None)


def cleanup(dir_path):
    import shutil
    shutil.rmtree(dir_path, ignore_errors=True)
