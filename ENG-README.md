<div align="center">

# THE PY-LIB

**A complete school-library management system — desktop app, staff web UI, and a self-service member portal**
AES-256-GCM encrypted identity QR (including time-limited rotating tickets) • atomic borrowing • 193 tests + CI

[![CI](https://github.com/Nonx2360/The-PY-LIB-Project/actions/workflows/ci.yml/badge.svg)](https://github.com/Nonx2360/The-PY-LIB-Project/actions/workflows/ci.yml)
![Python](https://img.shields.io/badge/Python-3.10%2B-306998?logo=python&logoColor=FFD43B)
![Platform](https://img.shields.io/badge/Desktop-Windows%2010%2B-0078D6?logo=windows&logoColor=white)
![Tests](https://img.shields.io/badge/tests-193%20passing-2EA44F)
![License](https://img.shields.io/badge/license-MIT-8A6BE8)

</div>

---

THE PY-LIB is a library platform built for a real school media center: staff run members, books and circulations from a desktop app or a web UI, while students get a **member portal** where they check their own loans and display a **rotating, time-limited QR identity** instead of relying on one static plastic card forever. Every layer — cryptography, permissions, transactions — is designed together and covered by a security-regression suite that runs in CI.

## Contents

- [Architecture & permissions](#architecture--permissions)
- [Features](#features)
- [Quick start](#quick-start)
- [Member account lifecycle](#member-account-lifecycle)
- [Configuration](#configuration)
- [HTTP endpoints](#http-endpoints)
- [Operations](#operations)
- [Security model](#security-model)
- [Database](#database)
- [Testing & CI](#testing--ci)
- [Project structure](#project-structure)
- [Dependencies](#dependencies)
- [Roadmap](#roadmap)
- [Contributing](#contributing)
- [Team & license](#team--license)

---

## Architecture & permissions

```
                         ┌─────────────────────────────────────┐
                         │  Desktop app (Tk/customtkinter)     │
                         │  members • books • circulations     │
                         │  access log • history • settings    │
                         └──────────────┬──────────────────────┘
                                        │ shares db/ + service layers
┌────────────────────────┐    ┌─────────▼──────────────────────┐    ┌────────────────────────┐
│  Staff web UI (Flask)  │    │  shared core                   │    │  Member portal (web)   │
│  server-side tables    ├───►│  db/repositories (atomic)      │◄───│  PIN login (bcrypt)    │
│  jsQR → /scan/resolve  │    │  services/qr + member_identity │    │  own status only       │
│  bulk + analytics      │    │  core/security (AES-GCM)       │    │  rotating v3 ticket    │
└────────────────────────┘    │  config/settings + .env        │    └────────────────────────┘
                              └─────────┬──────────────────────┘
                                        │
                              ┌─────────▼─────────┐
                              │   SQLite (db/)    │
                              └───────────────────┘  user_version v1→v4, automatic
```

**Permission matrix** — every page says unmistakably who may see it:

| Capability | Admin (desktop) | Admin (web) | Member (portal) |
|---|:---:|:---:|:---:|
| Members + books CRUD, bulk operations | ✓ | ✓ | ✗ |
| Borrow • return • access check-in/out | ✓ | ✓ | ✗ |
| Issue activation codes / clear member PINs | ✓ | ✓ | ✗ |
| System settings (incl. portal on/off) | ✓ | ✓ | ✗ |
| See own status + own rotating QR | — | — | ✓ |

The admin and member sessions live in separate cookie keys — logging into one side grants nothing on the other, and both walls are covered by dedicated tests.

## Features

### Desktop app (staff)

- **Encrypted QR identities** — webcam scanning for borrow, return and check-in/out; decoding happens on-device only, keyed from `.env`
- **Atomic circulation** — one `BEGIN IMMEDIATE` transaction per borrow/return: the book status flip is a conditional UPDATE (one winner per copy even if two stations scan simultaneously) and mid-operation failures leave no half-written state
- **Bulk management** — select many books/members, then retire / restore / delete together; rows with history are skipped automatically with a "Success X / Skipped Y" summary
- **Excel import** — 1,000+ rows per batch in a single transaction running in the background; the window stays responsive
- **Chunked list windows** — a 1,000+ item catalog opens immediately without crashing Tk
- **Documents** — member card PDFs and borrow/return/access reports in Thai (Sarabun), with an in-app PDF viewer
- **Backups** — SQLite online backup + integrity/foreign-key verification before major changes
- Light/dark theme, toasts, confirm dialogs, one shared scan window

### Staff web UI (Flask)

- Desktop-parity modules: members, books, borrow/return, history, access, settings — launched from the app or standalone (bound to `127.0.0.1` by default)
- **Server-side DataTables** (vendored jQuery + DataTables): search/sort/paginate ~50 rows/page at the server — catalogs in the thousands stay fast
- **Two-stage destructive flow** — retire first, delete-purge afterward; history integrity is always preserved
- Web Excel import up to 5,000 rows/file with a downloadable template and per-row skip reporting
- Chart.js dashboard (14-day circulations, 7-day access, book status split)
- **Fully offline** — every frontend library (SweetAlert2, jQuery + DataTables, Alpine.js, Chart.js, Flatpickr) and the Sarabun/Prompt fonts are pinned in the repo; zero CDN calls

### Member portal

- **Self-service login** with class + number + self-set PIN (only the bcrypt hash is stored)
- **Read-only status page**: current loans + due dates (overdue/due-today badges), recent returns, access history — there are no borrow/return actions; circulation stays at the desk
- **Rotating QR (v3 ticket)** — the member's page shows a personal QR that re-issues itself every time window (admin sets 1–72 h, plus 0–2 grace windows); a photo of yesterday's QR is useless when the window flips
- **Two-step first activation** — verify with your own printed card or an admin-issued one-time code, then set a PIN within a 10-minute window
- **Central control** — the whole portal can be switched off (`/member*` → 404) and static plastic-card QRs can be retired independently
- Login throttled to 5 failures / 15 min per (IP + class + number); error replies never reveal whether an account exists

## Quick start

**Requirements:** Python **3.10+** (the pinned dependency set is verified on Windows / Python 3.10; CI runs on 3.11) — the desktop app targets **Windows 10+** (OpenCV + pyzbar for scanning)

```bash
# 1) environment + dependencies (or run req.bat)
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# 2) start the desktop app
python main.py            # add --debug to write debug.log
```

**The first boot hardens itself:** there is no default password — a random one is generated, printed to the console **once**, and saved to `db/INITIAL_ADMIN_PASSWORD.txt` (gitignored; treat it as a secret). On every boot, any account still holding the historic `admin123` is replaced automatically.

**Web UI:** the desktop app's "เพิ่มเติม" (More) menu → "Web UI" → "Start Web UI", or run standalone:

```bash
venv\Scripts\python -c "from web.app import create_app; create_app().run(port=5000)"
```

Open `http://127.0.0.1:5000`. Exposing it on a LAN requires HTTPS in front — see [Operations](#operations).

**Secrets:** `.env` (gitignored) is auto-generated on first boot and must travel *with* the database — always move them as one set. See [Security model](#security-model).

> Browser cameras are only available over `localhost` or HTTPS (a getUserMedia security requirement).

## Member account lifecycle

| Step | Actor | Details and limits |
|---|---|---|
| 1 · Register | Admin | Add the member → the system issues an encrypted static card QR (v2) plus a printable PDF card |
| 2 · First activation | The member | At `/member/activate`: scan **your own card** or enter an admin one-time code (SHA-256 at rest, 15-minute expiry, shown once on screen) · then set a 6–64 character PIN within 10 minutes · accounts that already have a PIN refuse re-activation on every path |
| 3 · Daily use | Member | Log in with class + number + PIN → view own status and show the rotating QR (rate-limited to 30 images/min) |
| 4 · At the desk | Librarian | Scan the current-window QR (or a static card, if still accepted) → borrow / return / check in-out; web and desktop scanners run the identical policy |
| 5 · Reset | Admin | **Clear PIN** (burns any outstanding codes in the same transaction) → the member must re-verify with the physical card |

Key design point: the system **cannot** activate a member portal account *for* the student — access only exists after the member personally completes step 2, and picking up someone else's code never grants their account (codes bind to their owner).

## Configuration

All of these live under "Settings" (desktop and web share the same store — `config/settings.json`):

| Option | Description | Default |
|---|---|---|
| Theme | System / Light / Dark | System |
| Default loan period | Borrow duration (clamped 1–90 days) | 7 days |
| Max books per member | Concurrent-loan limit (clamped 1–20) | 3 books |
| Member card validity | Days before expiry (clamped 1–3,650) | 365 days |
| School name | Shown on cards and pages | DSNPRU |
| School logo | PNG for cards and the sidebar | `assets/logos/school_logo.png` |
| Web UI host | `127.0.0.1` (loopback) or `0.0.0.0` (LAN + HTTPS) | `127.0.0.1` |
| Enable member portal | Off = every `/member*` route answers 404 | Enabled |
| Member QR rotation window | Length of one v3 ticket window (clamped 1–72 h) | 12 hours |
| Grace windows | Extra accepted windows for clock skew (clamped 0–2) | 1 window |
| Accept static card QR | Still accept no-expiry plastic/paper QRs — turn off once members are activated | Enabled |

Numeric inputs are always clamped server-side (the form `min`/`max` are UX only), and the settings page posts two independent groups (general vs. member_portal) so an unchecked box can never wipe the other group's values.

## HTTP endpoints

| Zone | Routes | Methods | Role |
|---|---|---|---|
| Auth | `/login`, `/logout` | GET/POST, GET | Admin |
| Admin pages | `/`, `/members`, `/books`, `/borrow`, `/return`, `/history`, `/access`, `/settings` | GET | Admin |
| Table APIs | `/api/members`, `/api/books`, `/api/history`, `/api/stats`, `/api/analytics` | GET | Admin |
| Admin writes | `/members/add`, `/members/<id>/delete`, `/members/<id>/activation-code`, `/members/<id>/reset-pin`, `/members/bulk`, `/books/add`, `/books/import`, `/books/template`, `/books/<id>/delete`, `/books/<id>/retire`, `/books/<id>/restore`, `/books/bulk` | POST | Admin |
| Scanning | `/scan/resolve` | POST | Admin (all QR formats decode server-side) |
| Access log | `/access/record` | POST | Admin |
| Member portal | `/member/login`, `/member/logout`, `/member`, `/member/qr.png`, `/member/activate`, `/member/activate/scan`, `/member/activate/code`, `/member/activate/set-pin` | GET/POST | Member (session separate from admin) |
| Assets | `/assets/<path>` | GET | — |

Every POST requires a CSRF token (in the form or via the `X-CSRF-Token` header), and role-gated routes answer 302 to the appropriate login when the session lacks rights.

## Operations

**Backup / restore** — normally via desktop Settings, or the CLI:

```bash
# backup (SQLite online backup — safe while the server is running) + verify
venv\Scripts\python -c "from services.backup_service import create_backup, verify_backup; p = create_backup(); print(verify_backup(p), p)"

# restore (refuses unverified files; atomically replaces the live DB)
venv\Scripts\python -c "from services.backup_service import restore_backup; restore_backup(r'backups\library_backup_YYYYMMDD_HHMMSS.db')"
```

**Rotate QR keys** — if the key may have leaked, or you are migrating machines, the bundled tool rotates without losing any issued card:

```bash
venv\Scripts\python tools\rotate_qr_key.py [--no-cards]   # backup → decrypt → re-encrypt everything as v2
venv\Scripts\python tools\regenerate_member_cards.py      # rebuild card PDFs from the current database
```

After rotating: distribute the new `.env` to every machine **before** the next boot, and always migrate machines as the set (`.env` + `db/` + `assets/`) together.

**LAN / HTTPS** — put a reverse proxy in front to terminate TLS, then set `webui_host = 0.0.0.0` and enable `SESSION_COOKIE_SECURE`. Full steps: `docs/SECURITY.md` § 9.

**Common issues quick check:**

| Symptom | Where to look |
|---|---|
| Web camera unavailable | must be `localhost` or HTTPS; grant camera permission in the browser |
| Member login returns 429 | 5 failures / 15 min per identity set — wait it out, or have an admin clear the PIN |
| Old cards fail to decode everywhere | `.env` mismatch between machines — move them as one set or run `tools/rotate_qr_key.py` |
| All member pages 404 | someone disabled "Enable member portal" in Settings |
| Web tables slow or empty | check migrations ran (`PRAGMA user_version` = 4) and that `/api/*` returns data |
| Forgot the admin password | open `db/INITIAL_ADMIN_PASSWORD.txt`; further cases in `docs/SECURITY.md` |

## Security model

| Area | Approach |
|---|---|
| QR identity | **v2** — static card, `AES-256-GCM` (12-byte nonce + forgery-proof tag) · **v3** — rotating ticket, `AES-256-GCM` with protocol-binding AAD (`pylib-member-ticket-v3`), strict-field/strict-type JSON body (`{"m": id, "e": epoch}`) · **v1 legacy** — AES-CBC, read-only and decoded strictly (full PKCS7 + UTF-8) — **no plaintext fallback**, malformed input is rejected outright |
| Ticket time window | `epoch = floor(unix_time / window)` — the server accepts only the current window plus grace windows, and the design is **stateless** (nothing persists anywhere) |
| Key derivation | PBKDF2-HMAC-SHA256, 100,000 iterations, over `SECRET_PASSPHRASE` + `SECRET_SALT` (auto-generated; placeholder values are always replaced) |
| Admin passwords | bcrypt with no default — first boot prints a random one once; timing-equalized verification; `admin123` is retired automatically on every boot |
| Member PIN | bcrypt in `members.pin_hash` (NULL until the member activates); activation is a single transaction (code claim + PIN set + a `pin_hash IS NULL` guard against takeover) |
| Activation codes | SHA-256 hash at rest, 15-minute expiry, single atomic use, shown on screen exactly once — the plaintext is never persisted |
| Login throttling | 5 failures / 15 min in independent scopes: admin (IP + username), member (IP + class + number) |
| Web sessions | HttpOnly + SameSite=Lax cookie, 4-hour lifetime, CSRF token on every POST (constant-time compare), failure messages prevent user enumeration |
| Blast radius | the portal reads only the member's own rows and has no data-mutating routes at all; the whole domain can be disabled with one switch |
| Data integrity | FKs on every connection, `books.code UNIQUE`, `borrow_log.returned CHECK (0,1)`, borrow/return as one `BEGIN IMMEDIATE` transaction, parameterized SQL everywhere |
| Migrations | `PRAGMA user_version` v1→v4 (constraints → bulk-import indexes → member portal: pin_hash + activation codes) — idempotent, run at every boot |
| Logging | no personal data is ever logged (ids and counts only); failed payloads reply with outcomes, never echoing contents |

The full operational reference (every payload format, the key-rotation runbook, HTTPS deployment, the security test matrix): [`docs/SECURITY.md`](docs/SECURITY.md)

## Database

One SQLite file (`db/library.db`) — structure is managed by versioned migrations:

| Table | Contents |
|---|---|
| `admin_users` | Admin accounts (bcrypt, no default password) |
| `members` | Name, class, number, QR, expiry, **`pin_hash`** (v4 — NULL until the member activates); class+number badge must be unique |
| `books` | Code (**UNIQUE**), title, status `ว่าง` (available) / `ยืมแล้ว` (borrowed) / `หยุดให้บริการ` (retired) |
| `borrow_log` | Circulation records — `returned CHECK (0,1)`, FKs to members/books, bulk indexes (v3) |
| `access_log` | Library entry/exit records |
| `member_activation_codes` | (v4) SHA-256 hashes of one-time activation codes, 15-minute expiry, `code_hash UNIQUE`, `used_at` burn timestamp |

## Testing & CI

```bash
venv\Scripts\python -m unittest discover -s tests        # Windows
venv/bin/python -m unittest discover -s tests            # Linux/Mac
```

**193 test cases across 18 files** — crypto (v1/v2/v3), the member-identity resolver, the full member-portal flow, repositories and atomic circulation, migrations v1→v4, backup/restore, bulk + Excel import, server-side table APIs, web guards (CSRF/throttle/sessions/permissions) and offline-safety.

CI (GitHub Actions) runs three independent jobs on every push/PR:

| Job | What it does |
|---|---|
| `tests + lint` | the entire unittest suite + `ruff check` (configured in `pyproject.toml`, including bandit rules) |
| `pip-audit` | CVE scan of the installed dependency set |
| `gitleaks` | committed-secret detection |

## Project structure

```
THE-PY-LIB/
├── main.py                      # Entry point (--debug flag)
├── app.py                       # LibraryApp shell + navigation controller
├── legacy_main.py               # Original monolith (preserved for rollback)
├── config/
│   ├── constants.py             # DB_PATH, font paths, book status constants
│   ├── settings.py              # User settings (config/settings.json)
│   ├── env_store.py             # .env secrets: auto-generate, guard placeholders
│   └── settings.json            # User settings (auto-created)
├── core/
│   ├── security.py              # bcrypt hash/verify, AES-GCM `v2:` badges + `v3:` rotating tickets + legacy AES-CBC (PBKDF2)
│   ├── logger.py                # logging config (debug.log support)
│   └── throttle.py              # Sliding-window login lockout (desktop + web)
├── db/                          # (runtime artifacts: library.db, INITIAL_ADMIN_PASSWORD.txt — gitignored)
│   ├── schema.py                # CREATE TABLE + user_version migrations (v1→v4) + admin bootstrap
│   └── repositories/
│       ├── base_repo.py         # Thread-safe SQLite (per-query connection + Lock)
│       ├── member_repo.py       # Member CRUD (+ PIN bcrypt, activation codes, atomic activation)
│       ├── book_repo.py         # Book CRUD (uses book status constants)
│       ├── borrow_repo.py       # try_borrow / try_return — atomic invariants
│       ├── access_repo.py       # Access log
│       └── admin_repo.py        # admin_users + timing-equalized verify_credentials
├── services/
│   ├── qr_service.py            # Create/decode QR (AES)
│   ├── member_identity.py       # Single resolver for all QR types (v3 ticket / v2 card / legacy) + time-window policy
│   ├── pdf_service.py           # Member card PDF + reports
│   ├── camera_service.py        # Camera + QR decode
│   ├── excel_service.py         # Import/Export Excel
│   ├── backup_service.py        # SQLite backups + checksum verification
│   └── __init__.py
├── ui/
│   ├── theme.py                 # Design tokens, font registration, widget helpers
│   ├── views/
│   │   ├── login_view.py        # Split-panel login + throttling
│   │   ├── dashboard_view.py    # Stat cards + action tiles
│   │   ├── member_view.py       # Two-column: form + searchable list, PIN tools
│   │   ├── book_view.py         # Two-column: form + searchable list, retire/restore
│   │   ├── borrow_view.py       # 3-step flow (scan → book → confirm)
│   │   ├── return_view.py       # Scan + borrowed list + return
│   │   ├── history_view.py      # Table search + PDF export
│   │   ├── access_view.py       # Scanner + history (AccessScannerView/AccessHistoryView)
│   │   ├── settings_view.py     # Theme, loan settings, member portal, school info, DB backup
│   │   └── webui_view.py        # Web UI launcher control panel
│   └── widgets/
│       ├── confirm_dialog.py    # Modal confirm dialog
│       ├── toast.py             # Floating toast notifications
│       ├── scan_window.py       # Shared camera QR scan window
│       └── pdf_viewer.py        # Built-in PDF viewer (PyMuPDF)
├── web/
│   ├── app.py                   # Flask app: routes, guards, DataTables /api feeds, bulk, analytics, member portal, Excel import
│   ├── csrf.py                  # Per-session CSRF tokens (form field or X-CSRF-Token header)
│   ├── static/
│   │   ├── style.css            # Web UI design system
│   │   ├── vendor/              # Vendored JS libs, pinned versions (see README.md inside)
│   │   │   ├── alpine.min.js    # Alpine.js 3.17.5
│   │   │   ├── jquery.min.js    # jQuery 3.7.1
│   │   │   ├── datatables.min.js # DataTables 1.13.8 (server-side tables)
│   │   │   ├── sweetalert2.min.js/.css  # SweetAlert2 11.26.25
│   │   │   ├── chart.umd.min.js # Chart.js 4.6.1
│   │   │   ├── flatpickr.min.*  # Flatpickr 4.6.13 + Thai locale (th)
│   │   │   └── htmx.min.js      # htmx 2.0.11 (retained, currently unused)
│   │   ├── fonts/               # Sarabun + Prompt TTFs (OFL) — @font-face in style.css, offline
│   │   ├── lang/
│   │   │   └── datatables-th.json # hand-made Thai i18n for DataTables
│   │   └── js/
│   │       ├── jsQR.js          # Vendored jsQR (in-browser QR decode)
│   │       ├── scan.js          # Camera capture + /scan/resolve calls
│   │       ├── member_scan.js   # Card-scan camera on the member activation page
│   │       ├── member_qr.js     # Countdown + QR swap when the window flips
│   │       ├── tables.js        # DataTables configs (books/members/history) + bulk bar
│   │       └── ui.js            # PyLib helpers: dialogs, toasts, data-confirm, flatpickr init
│   └── templates/
│       ├── icons.svg            # Lucide SVG icon sprite
│       ├── base.html            # Layout + sidebar + vendor stack wiring
│       ├── login.html           # Login page
│       ├── member_base.html     # Member pages (standalone shell, no admin nav)
│       ├── member_login.html    # Member login (class + number + PIN)
│       ├── member_activate.html # First activation (scan card or admin code) + set PIN
│       ├── member_portal.html   # Own status + rotating QR with countdown
│       ├── index.html           # Dashboard + Chart.js widgets
│       ├── members.html         # Member management
│       ├── books.html           # Book management (retire/restore/purge)
│       ├── borrow.html          # Borrow (Alpine lock state + Flatpickr due date)
│       ├── return_book.html     # Return with overdue badges
│       ├── history.html         # Borrow history
│       ├── access.html          # Access recording + history
│       └── settings.html        # Settings + about
├── tools/
│   ├── rotate_qr_key.py         # Key rotation: backup → decrypt → re-encrypt
│   └── regenerate_member_cards.py
├── tests/                       # _harness.py + 18 test files (see Testing & CI)
├── docs/
│   ├── SECURITY.md              # Security runbook (key rotation, HTTPS/LAN, payload formats)
│   └── DESIGN.md                # Visual design system notes
├── assets/
│   ├── fonts/                   # Sarabun + Prompt TTFs (desktop PDFs + copied to web/static/fonts)
│   ├── logos/                   # school_logo.png
│   ├── qrcodes/                 # Generated QR images (runtime, gitignored)
│   └── cards/                   # Generated member card PDFs (runtime, gitignored)
├── .env.example                 # SECRET_PASSPHRASE, SECRET_SALT, FLASK_SECRET_KEY template
├── .env                         # (runtime, gitignored, auto-generated)
├── pyproject.toml               # ruff lint/style configuration
├── requirements.txt             # pinned versions
└── README.md
```

## Dependencies

| Package | Role |
|---|---|
| `customtkinter` | Desktop GUI (Modern Tkinter) |
| `flask` | Web UI |
| `opencv-python` + `numpy` | Webcam capture + image processing (desktop) |
| `pyzbar` | QR decode from camera frames (desktop) |
| `qrcode` | QR image generation |
| `cryptography` | AES-256-GCM / AES-256-CBC, PBKDF2 |
| `bcrypt` | Admin password + member PIN hashing |
| `reportlab` | Member card & report PDFs |
| `pymupdf` | In-app PDF viewer |
| `pillow` | Image handling |
| `pandas` + `openpyxl` | Excel import/export |
| `python-dotenv` | `.env` loading |
| `packaging` | customtkinter dependency |

All versions are pinned in `requirements.txt` (verified on Windows / Python 3.10 per the file's header).

## Roadmap

Current status: every core capability is production-ready — next up:

- [ ] Generate member cards / report PDFs from the web (desktop-only today)
- [ ] Let members change their own PIN from the portal (currently admin-reset only)
- [ ] Scheduled automatic backups + retention policy
- [ ] No-login status kiosk page

Bugs and feature requests can be filed directly as [issues](https://github.com/Nonx2360/The-PY-LIB-Project/issues).

## Contributing

1. Fork, then branch (`feature/…` or `fix/…`)
2. `ruff check .` must pass — configured in `pyproject.toml`
3. The unittest suite must stay green: `python -m unittest discover -s tests`
4. Security-touching changes (crypto / permissions / sessions) should add regression cases matching the matrix in [`docs/SECURITY.md`](docs/SECURITY.md) § 11
5. Open a PR with the rationale and screenshots (for UI changes)

Non-code ideas are welcome too: open an issue describing a real-world library scenario.

## Team & license

- Miss Khwanchanok Unsiri, No. 4 (UX/UI)
- Mr. Natnanon Rodnoi, No. 30 (Developer)
- **School:** DSNPRU

Thanks — beta testers and advisors.

MIT License © 2025
