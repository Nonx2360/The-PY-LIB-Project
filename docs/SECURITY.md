# THE PY-LIB — Security & Operations Guide

This document is the single reference for everything security-related in
THE PY-LIB: how secrets are stored, how QR payloads are encrypted, how the
database is protected and migrated, how backups are taken and restored, and
the runbooks an administrator is likely to need.

---

## 1. Secrets and the `.env` file

All project secrets live in `.env` at the project root:

| Key | Purpose | Generated when |
|---|---|---|
| `SECRET_PASSPHRASE` | Root material for the QR-encryption key (PBKDF2 input) | First boot, or replacement |
| `SECRET_SALT` | PBKDF2 salt for the QR-encryption key | First boot, or replacement |
| `FLASK_SECRET_KEY` | Flask session signing key (web UI) | First boot, or replacement |

Rules enforced by `config/env_store.py` (`ensure_secrets`):

* Missing, too-short, or known-placeholder values are **replaced by
  cryptographically random ones** and persisted to `.env` automatically.
* Historic placeholder/default values (`default_secret_py_lib_key_change_me_in_dotenv`,
  `pylibsalt123`, …) are **banned** — they are rejected and re-generated.
* Unknown keys in `.env` are preserved on every rewrite.
* `.env` must stay out of version control (it is in `.gitignore`).

> ⚠️ If `SECRET_PASSPHRASE` or `SECRET_SALT` ever change, previously-issued QR
> cards can no longer decode. See §2 (formats) and §7 (rotation runbook).

For tests and CI, the environment variable `PYLIB_ENV_FILE` redirects the
secret file to a disposable path so test runs never touch the real `.env`.

## 2. QR payload encryption

Member data stored in the QR payload: `name|grade|number|register_date|expire_date`.
Implementation: `core/security.py`. Two formats exist:

**v2 — current (AES-256-GCM, authenticated)**

```
v2: + base64url( nonce[12] || ciphertext || tag[16] )
```

* Any tampering, truncation, or wrong-key decryption is **rejected** (GCM tag).
* All *new* member registrations produce v2 payloads.

**v1 — legacy (AES-256-CBC, unauthenticated) — supported for reading only**

```
base64url( IV[16] || AES-256-CBC ciphertext )   PKCS7 padding
```

* Kept strictly so cards already issued keep working; decode is **strict**
  (full PKCS7 + UTF-8 validation). Plain-base64 legacy payloads and anything
  malformed are rejected — **there is no plaintext fallback** (this was an
  intentional removal; the old behavior silently accepted any base64 text).
* New/rotated tooling always writes v2.

**Key derivation** — `PBKDF2-HMAC-SHA256(SHA-256, 32-byte key, 100k iterations)`
over `SECRET_PASSPHRASE` with `SECRET_SALT`.

## 3. QR key rotation and the rotation tool

`tools/rotate_qr_key.py` rotates the key for an *existing* install without
losing any member card. It:

1. Takes a **verified SQLite backup** into `backups/` first.
2. Decodes every stored `qrcode_data` under the **current** key — aborts
   atomically (no writes) if any row fails.
3. Generates a new random `SECRET_PASSPHRASE` + `SECRET_SALT`.
4. Re-encrypts all payloads as **v2**, rewrites the QR PNGs and card PDFs.
5. Persists the new secrets to `.env` (preserving `FLASK_SECRET_KEY`).

```bash
venv\Scripts\python tools\rotate_qr_key.py [--no-cards]
```

After running it: distribute the new `.env` to every machine instantly, before
the next boot (§7).

**When to rotate:** the passphrase/salt were ever committed, shared between
installs outside a secure channel, or an operator with access left.

## 4. Administrator accounts & first boot

Implemented in `db/schema.py`:

* **There is no default password.** On a fresh boot, `init_db()` creates an
  `admin` account with a strong random password, prints it **once** to the
  console, and appends it to git-ignored `db/INITIAL_ADMIN_PASSWORD.txt`.
* On any boot, an admin that *still* authenticates with the historic default
  `admin123` is **retired** and replaced with a fresh random password
  (belt-and-braces so an old install cannot keep functioning on the default).
* Legacy plain-text password hashes are normalized to bcrypt on boot.
* Nothing about usernames is written to logs (`web login: failed attempt`, …).

First-boot console output looks like:

```
================================================------------
NOTICE - first boot - initial admin password generated 2026-10-30 12:00
  admin account : admin
  password      : xX3f9…
  saved to      : db\INITIAL_ADMIN_PASSWORD.txt  (gitignored - back it up)
============================================================
```

Store that password; the file is for reference only.

## 5. Database integrity & migrations

* **Versioned migrations:** schema state is tracked with `PRAGMA user_version`
  (`db/schema.py`), not just `CREATE TABLE IF NOT EXISTS`.
  * v0 (no tables) → fresh install created directly in the v2 layout.
  * v1 (pre-refactor layout) → migrated to v2.
  * v2 adds, inside one transaction:
    - **UNIQUE index** on `books(code)` (duplicate legacy codes are de-duplicated
      by renumbering later copies to `code-<id>` — nothing is deleted);
    - `borrow_log` rebuilt with `returned INTEGER NOT NULL CHECK (returned IN (0,1))`,
      `member_id`/`book_id` `NOT NULL`, and FKs declared (legacy NULL/invalid
      `returned` values are coerced to `1` — returned);
  * Migrations are idempotent and safe to run on every boot.
* **Foreign keys** are turned ON (`PRAGMA foreign_keys=ON`) for *every*
  connection created by the repositories (`db/repositories/base_repo.py`).
* **`busy_timeout=5000`** on every connection to tolerate transient locks.
* **Atomic operations:** `BaseRepository.transact()` runs a callable inside a
  single `BEGIN IMMEDIATE` transaction with rollback on any exception.
  `BorrowRepository.try_borrow()` / `try_return()` use it:
  - `try_borrow` flips `books.status` via a *conditional* UPDATE
    (`WHERE status='ว่าง'`) and checks rowcount — only one concurrent borrower
    can win; on any failure (overdue, limit, FK, integrity) everything rolls
    back — no partial borrow records or half-flipped book statuses.
  - `try_return` derives the book id **from the borrow_log row itself**;
    client-supplied book ids are never trusted (a mismatched return record is
    rejected without touching another book).

All SQL in the codebase is parameterized — injected strings are treated as
ordinary data (covered by regression tests).

## 6. Backup & recovery

`services/backup_service.py`:

* `create_backup()` — online **SQLite backup API** snapshot (safe while the
  app/web server is running); default destination `backups/library_backup_*.db`
  (gitignored) or an explicit path (used by the desktop Settings dialog).
* `verify_backup()` — read-only `PRAGMA integrity_check` +
  `PRAGMA foreign_key_check`. Returns `ok`, or `ok=False` for unreadable files.
* `restore_backup()` — refuses an unverified/failed backup, then swaps the
  live DB **atomically** (temp file + `os.replace`, with a short retry loop
  for Windows AV/indexer transient locks).

**Recovery runbook:**

1. Stop the app (optional but recommended).
2. Choose the newest backup whose `verify_backup()` passes.
3. `python -c "from services.backup_service import restore_backup; restore_backup(r'backups\library_backup_XXXXXXXX_XXXXXX.db')"`
4. Start the app; the boot migration/bootstrap remain no-ops on an intact DB.
Always keep at least two independent copies (the automatic `backups/`
directory is a machine-local convenience, not an off-site backup).

## 7. Changing machines / disaster recovery of cards

1. Copy **`.env` + `db/library.db` + `assets/qrcodes/` + `assets/cards/`**
   together — they are encrypted with each other.
2. If `.env` must differ, rotate instead: run §3 on the OLD machine, then copy
   the new `.env` + re-encrypted assets together.

## 8. Web UI hardening defaults

Implemented in `web/app.py`, `web/csrf.py`, `web/throttle.py`:

* **Binds `127.0.0.1` only** by default (`webui_host` setting; `0.0.0.0`
  requires deliberate opt-in and HTTPS — see §9).
* **Session cookies:** signed `pylib_session` cookie with `HttpOnly`,
  `SameSite=Lax`, 4-hour lifetime, 1 MiB body cap; secret from `.env`.
* **CSRF:** a per-session token is rendered in every POST form and validated
  with a constant-time comparison before any handler runs. Missing/wrong
  tokens → rejected via redirect, no state change.
* **Login throttling:** 5 failures per 15 minutes per (IP, username), cleared
  on success. Throttled and wrong-credential responses use the same generic
  message (no user enumeration).
* **Validation & clamping:** all form inputs length-limited (`LIMITS`),
  control characters stripped, numeric settings clamped to sane bounds
  server-side (client `min`/`max` are UX only).
* **PII-free logging:** student names are never logged; borrow/return/access
  actions log only ids/codes.
* SQL injection strings in any field are ordinary data (parameterized SQL).

## 9. Deploying the web UI beyond one machine

The bundled server + SQLite is designed for one site. To share on a LAN:

1. Put a reverse proxy (or the school's existing web server) in **front** and
   terminate TLS there — the app itself serves plain HTTP.
2. Set `webui_host` in Settings to `0.0.0.0` (or the LAN interface) and restart
   the Web UI.
3. Enable secure cookies once HTTPS is in front:
   set `SESSION_COOKIE_SECURE=True` in `create_app` configuration.

## 10. CI checks

`.github/workflows/ci.yml` runs on pushes and PRs:

* `ruff check .` — lint/security smells (config in `pyproject.toml`).
* the full test suite (`python -m unittest discover -s tests`) — includes the
  security regression matrix below;
* `pip-audit` — known vulnerability scanning of the dependency set;
* `gitleaks` — committed-secrets scan of the working tree.

## 11. Security regression test matrix

Answered by `tests/` (unittest):

| Test | File |
|---|---|
| Wrong admin password rejected | `test_admin_bootstrap`, `test_web_security` |
| First launch: no default password remains | `test_admin_bootstrap` |
| Tampered QR ciphertext rejected, no plaintext fallback | `test_qr_security` |
| QR from old key rejected | `test_qr_security` |
| SQL-injection as ordinary input | `test_borrow_atomic`, `test_web_security` |
| Unauthenticated POST rejected | `test_web_security` |
| POST without CSRF rejected, no state change | `test_web_security` |
| Two simultaneous borrows → at most one loan | `test_borrow_atomic` |
| DB error mid-borrow → no partial state | `test_borrow_atomic` |
| Return nonexistent/mismatched record rejected | `test_borrow_atomic` |
| Backup/restore preserves records & relationships | `test_backup_restore` |
| Migrations enforce constraints (UNIQUE/CHECK/FK) | `test_migrations` |
