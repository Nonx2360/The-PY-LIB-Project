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

Implementation: `core/security.py`. Three formats exist:

**Member identity payloads**

* Static cards (v2/v1/legacy): `name|grade|number|register_date|expire_date`.
* Rotating tickets (v3): the payload body is just
  `{"m": <member_id>, "e": <epoch>}` — no personal data inside the QR.

**v3 — rotating member ticket (current for on-screen identities)**

```
v3: + base64url( nonce[12] || AES-256-GCM( JSON {"m":id,"e":epoch} ) || tag[16] )
    AAD = b"pylib-member-ticket-v3"   (protocol binding — v2 bodies fail here and vice versa)
```

* `e` is the **time-window epoch**: `floor(unix_time / window_seconds)` with the
  admin-configured window (default 12 h, clamp 1–72 h). A ticket is accepted
  only inside its own window plus `member_qr_grace_steps` neighbours
  (default 1, clamp 0–2) to absorb small clock skew.
* The window math is **stateless** — nothing is persisted, so clocks and
  machines agree without a shared store; `e` is a Unix-epoch multiple (not
  local midnight).
* Decoding is **strict**: the JSON body must be exactly `{"m", "e"}` with
  positive-int values (extra keys, strings, booleans, floats → rejected;
  `parse_ticket_json`).
* The AAD binds the v3 envelope to the member-ticket protocol, so a v2 badge
  can never be confused for a ticket (cross-protocol replay fails the GCM tag).
* Payload is ~70–90 chars — stays comfortably inside quality-scannable QRs.

**v2 — static card (AES-256-GCM, authenticated)**

```
v2: + base64url( nonce[12] || ciphertext || tag[16] )
```

* Any tampering, truncation, or wrong-key decryption is **rejected** (GCM tag).
* All *new* member registrations produce v2 payloads.

Acceptance policy (`services/member_identity.py`, shared by the web scanner and
the desktop camera so a QR never behaves differently depending on the reader):

* `accept_static_card_qr` (admin setting, default on) — when off, printed-card
  QRs are refused in favour of rotating QRs only.
* Expired-window tickets raise a Thai "QR หมดเวลา" outcome; unknown members
  raise "ไม่พบสมาชิกนี้ในระบบ" — messages only, never payload content.

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
over `SECRET_PASSPHRASE` with `SECRET_SALT`. One key serves the badge formats
(v1/v2) and member tickets (v3) alike; rotating it (§3) invalidates every
in-flight ticket window as well as static cards.

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
  * v3 adds bulk-import indexes on `borrow_log(member_id, book_id)`;
  * v4 (member portal era) adds `members.pin_hash` (bcrypt, NULL until personal activation) and `member_activation_codes` (SHA-256 hashes, 15-min expiry, single-use);
  * v5 (academic-year migration) adds `members.status` and tables for academic years, member placements (one current per member), and migration runs/items (audit trail). Promotions keep stable member IDs and do not rewrite historical borrow/access rows; upgrades seed a baseline current placement.
  * v6 (inventory audit) adds inventory audit sessions, scans, and discrepancy resolutions (reconciliation only — no automatic destructive/status changes).
  * v7 (reservations) adds `book_reservations` for copy-level FIFO queues/offers with explicit states and offer expiry.
  * Migrations are idempotent and safe to run on every boot (fresh install == upgraded).
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

### 8.1 Member portal (`/member…`) hardening

Implemented in `web/app.py` + `db/repositories/member_repo.py`:

* **Key separation:** the member session lives in a different session key (`member_id`) than the admin one (`logged_in`). A member session cannot open any admin route, and an admin session does not silently inherit member pages.
* **Member login:** grade+number+PIN against `members.pin_hash` (bcrypt). Errors are generic; throttled 5 failures / 15 min per (IP, `member:<grade>:<number>`).
* **First activation ceremony:** scan own card or admin-issued one-time code, then a 10-minute server-side window to set a PIN (6–64 chars). Takeover guarded when `pin_hash` is already set.
* **One-time codes:** stored SHA-256-hashed, 15-minute expiry, single-use, shown once; admin PIN reset burns leftover codes atomically.
* **Portal is read-only by design:** own status/loans/returns/access/QR plus reservation actions (reserve/cancel/explicit claim). No silent loans; claiming performs an explicit borrow through shared atomic services and server-side checks.
* **Rotating QR endpoint** (`/member/qr.png`): session-bound, rate-limited, `no-store`, generated in memory.
* **Kill switch:** `member_portal_enabled=False` makes every `/member…` route answer 404.

### 8.2 Self-service kiosk (`/kiosk`) hardening

Implemented in `web/app.py` + `web/templates/kiosk.html`:

* **Isolated session/context:** kiosk uses its own session keys (identified member + idle timestamp) and does **not** grant or require admin/member-portal privileges; admin routes remain `login_required`.
* **Minimal exposure:** shows only the information needed for the current identified transaction; no admin data/PII listings.
* **Identity and actions:** member identity resolved server-side via the shared member-identity resolver; borrow/return execute through shared atomic repository services (same invariants/limits) with server-side book resolution.
* **Safety controls:** CSRF on POSTs, throttled identity resolve attempts, automatic idle reset/timeout (`kiosk_idle_timeout_seconds`, default 90s), and enable/disable via `kiosk_enabled`.
* **Reservation-aware:** borrowing respects active offers/holds (no bypass of reservation rules).

### 8.3 Reservations and circulation integrity

Implemented in `db/repositories/borrow_repo.py`, `db/repositories/reservation_repo.py`, `services/reservation_service.py`:

* **Explicit states:** copy-level reservations use `waiting` → `offered` → `fulfilled` (or `cancelled`/`expired`) with timestamps/expiry.
* **Deterministic FIFO:** queues are ordered per copy; when a copy becomes available, the next eligible waiter is promoted to an **offer** (not a loan) transaction-safely.
* **No silent loans:** returns and expiries never create `borrow_log` rows; the offered member (or staff flow) must explicitly claim/borrow.
* **Offer enforcement:** an active unexpired offer blocks borrowing by others; claiming/borrowing by the offered member fulfills the offer atomically in the borrow transaction.

### 8.4 Administrative workflows

* **Academic Year Migration:** admin-only preview (read-only), reviewed plan, explicit typed confirmation, duplicate-success guard, and **verified backup preflight** before a single-transaction apply with per-item audit records. Stable member IDs and historical loan/access records are preserved (no retroactive loan-time class rewrite). No unsafe automatic undo — recovery is restore-from-backup.
* **Inventory Audit:** admin-only reconciliation sessions (create/pause/resume/complete/cancel), server-side scan resolve, duplicate/unknown handling, and separation of on-loan copies from shelf-missing discrepancies. Audits record findings/resolutions and do **not** auto-delete or auto-mutate book/member/loan state.

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
| v4 upgrade: pin_hash NULL for legacy members, activation codes table + index, idempotent | `test_migrations` |
| Rotating ticket roundtrip; tamper/wrong-key/foreign-AAD rejected | `test_rotating_tickets` |
| Ticket JSON strictness (extra keys / wrong types / zero ids rejected) | `test_rotating_tickets` |
| Window math: epoch boundaries, grace ±1/±2, stale tickets refused | `test_rotating_tickets` |
| Shared resolver: static cards acceptable/refusable, dispatch to tickets | `test_rotating_tickets` |
| Member activation via card scan / one-time code; PIN-length rules | `test_member_portal` |
| Activation refused after PIN set (no takeover); expired activation window | `test_member_portal` |
| One-time code: single use, atomic claim, foreign member cannot win it | `test_member_portal` |
| Member session cannot reach admin routes; admin session has no portal | `test_member_portal` |
| Portal shows only the member's own loans/returns/access rows | `test_member_portal` |
| Member QR PNG: session-bound, no-store, rate-limited (429) | `test_member_portal` |
| Admin PIN reset locks the member out and burns leftover codes | `test_member_portal` |
| Settings groups save independently; portal off → /member 404 | `test_member_portal` |
| v3 ticket scan via /scan/resolve: valid 200, stale 400 expired, unknown 404 | `test_web_scan` |
| v5–v7 migrations (placements/migration audit, inventory audit, reservations) + idempotent fresh/upgrade parity | `test_migrations` |
| Academic year migration: read-only preview, atomic apply, stable IDs + history preserved, duplicate-run guard, graduate/left/manual handling | `test_academic_year_migration` |
| Inventory audit: scan found/duplicate/unknown, on-loan separate from shelf-missing, no auto destructive/status mutation, status transitions, resolutions recorded | `test_inventory_audit` |
| Reservations: FIFO offer on return (no silent loan), offer blocks others, explicit claim fulfills offer, expiry/cancel promote next in queue | `test_reservations` |
