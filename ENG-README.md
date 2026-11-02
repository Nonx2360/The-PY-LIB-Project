# THE PY-LIB — Smart Library System

**THE PY-LIB** is a library management system developed with Python (CustomTkinter + Flask + SQLite), focusing on security, convenience, and practical use in schools or educational institutions. Supports both desktop app and website.

---

## Key Features

### Desktop App
- Modular GUI (separated config / core / db / services / ui layers)
- QR Code scanning via webcam camera for borrowing/returning and check-in/check-out logging
- Atomic transaction borrow/return — prevents double-borrowing / stale data from errors
- Shows overdue return status with warning colors (green / yellow / red)
- Retire / restore book status while keeping borrow history
- Bulk management on books and members pages: select multiple items then retire / restore / delete together (items with history or currently borrowed are automatically skipped, with "Success X / Skipped Y" summary)
- Export member cards PDF, borrow/return history, access history with Thai font (Sarabun)
- Import/export books from Excel files (pandas + openpyxl) — import 1,000+ books per batch (single transaction, works in background, window doesn't freeze)
- List tables open in chunks (first 50 items + show more button) — 1,000+ book catalog opens immediately without excessive Windows resource usage (prevents Tk crashes)
- Modern UI, supports light/dark theme, Sidebar with Grouped Navigation
- Toast notifications, Confirm dialogs, Shared scan window
- Built-in PDF viewer (PyMuPDF) for viewing member cards
- Backups (SQLite backup + checksum) used before important changes
### Website (Flask Web UI)
- Enable from sidebar menu "Web UI" in desktop app or run separately
- Desktop-equivalent features: members, books, borrow/return, history, settings
- Check-in/check-out via web (automatic toggle in/out)
- Scan member cards from web browser camera (jsQR vendored) — decoded server-side only
- Retire/restore and permanent delete in 2 steps (retire → purge) with safe storage
- Import books from Excel online on web (max 5,000 rows/file) with template download — duplicate codes/incomplete rows are skipped and results reported
- Bulk management on web equivalent to desktop: select multiple books/members then command simultaneously (max 200 items/time)
  - Books: retire / restore / delete — books currently borrowed or with history still active are automatically skipped (retired books are deleted together with borrow history)
  - Members: bulk delete — members with borrow history/access history are skipped to maintain data integrity
  - Every action through confirmation dialog + "Success X / Skipped Y" report and count-only logs (no personal data recorded)
- Books/members/history tables (web) are **server-side jQuery DataTables** (search/sort/paginate server-side, ~50 rows/page, Thai interface) — 1,000+ book catalog opens fast always
- SVG icons (Lucide), School logo, search/filter on main tables — books, members, and history (other pages use forms/field selectors according to context)
- Member cards and PDF reports are desktop-only features (web does not yet open PDF creation routes)
- Responsive design, supports Light/Dark mode
- Frontend stack vendored locally (works even without internet) — see `web/static/vendor/README.md`:
  - **SweetAlert2** — confirmation/alert boxes themed with the system, flash messages as toast
  - **jQuery + DataTables** — books/members/history tables search/filter (code/name/status)/sort/paginate server-side via `/api/{books,members,history}` with bulk bar on books/members pages and hand-made Thai i18n file
  - Web fonts **Sarabun + Prompt** (OFL license) pinned in `web/static/fonts/` via `@font-face` — no more Google Fonts calls, every page including login works fully offline
  - **Alpine.js** — locks borrow page status (pending send/complete count) evaluated instantly in browser
  - **Chart.js** — dashboard: borrowings last 14 days, check-ins last 7 days, book status distribution (`/api/analytics`)
  - **Flatpickr (th)** — calendar for due date selection with quick buttons 7/14/21/30 days
- Load order: sweetalert2 → ui.js → [jquery + datatables + tables.js] → alpine (important, do not swap)
### Member Portal
- Members log in with **class + number + self-set PIN** (bcrypt stores hash only) — sees only their own data
- Read-only status page: currently borrowed books + due dates (overdue/today badges), latest returns, library access history — no borrow/return buttons, all borrow/return still done at counter
- **Rotating QR (v3 ticket)**: Member page shows personal QR that changes automatically per time cycle (admin sets 1–72 hours + neighboring cycle allowance 0–2) — QR from card photo cannot be reused when cycle changes
- First-time identity confirmation 2 methods: scan your own card with camera or one-time activation code from admin (displayed on screen once, SHA-256 hash stored in database, expires 15 minutes)
- Admin can reset member PIN (desktop/web) — pending activation codes are cancelled together, member must re-verify identity
- Enable/disable member page system-wide from settings (disabled → /member returns 404) and choose whether to still accept QR on traditional plastic cards
- Member session separate from admin different key — two-side logins don't interfere, limit failed logins 5 times/15 minutes (per IP+class+number), error messages cannot be tampered to say "no such member"

### Security
- QR data encrypted AEAD **AES-256-GCM (`v2:`)** and still decrypts old cards via AES-256-CBC strictly (strict legacy fallback)
- **Rotating QR ticket (`v3:`)**: AES-256-GCM + AAD protocol-bound ("pylib-member-ticket-v3") — each cycle ticket time-limited, server-side decode only, small payload (easy scan) and rejects extra/incorrect-type JSON
- PBKDF2 key derivation (100k iterations) for QR encryption key
- Admin passwords hashed with bcrypt — no default password and timing-equalized verification
- Limits failed logins (5 times / 15 minutes per user) both web and desktop
- CSRF token on every POST form and session cookie (HttpOnly/SameSite=Lax)
- Parameterized queries prevent SQL Injection everywhere
- Thread-safe database access (connection-per-query + Lock) and FK enforcement enabled on every connection
- Secret migration via `PRAGMA user_version` per schema version
- Secrets stored in `.env` (gitignored) — auto-generated randomly on first run

---

## Installation and Usage

### Install Python

Requires Python >= 3.10

### Create Virtual Environment and Install Dependencies

```bash
python -m venv venv
venv\Scripts\activate        # Windows
pip install -r requirements.txt
```

Supports using `req.bat` for automatic dependencies installation

### Run Desktop App

```bash
python main.py              # normal mode
python main.py --debug      # debug mode (log to debug.log)
```

Default: no default password anymore — first boot will **create a random password** shown on console screen only once and saved at `db/INITIAL_ADMIN_PASSWORD.txt` (not tracked in Git). If old default password (`admin123`) is detected in old database, system will change to new password automatically every boot.

> Login throttling: more than 5 failed logins within 15 minutes (per username) must wait until time limit, both web page and desktop

### Run Website

Open desktop app → menu "More" → "Web UI" → press "Start Web UI"

Or run separately (bound to this machine only 127.0.0.1):

```bash
venv\Scripts\python -c "from web.app import create_app; create_app().run(port=5000)"
```

Open browser to `http://127.0.0.1:5000`

> Security: web authenticates with session cookie (HttpOnly/SameSite=Lax), CSRF token on every POST form, limits failed logins (5 times/15 minutes) and does not record personal data
> Web logging is bound to `127.0.0.1` by default — to open on LAN see HTTPS steps at `docs/SECURITY.md` and change `webui_host` value on settings page

> QR keys: `.env` file stores `SECRET_PASSPHRASE` / `SECRET_SALT` / `FLASK_SECRET_KEY` (gitignored; auto-generated on first run, keep it secure) New QR data format uses AES-256-GCM (`v2:`) key rotation/migration/recovery — read `docs/SECURITY.md`

> Member card scan on web: web browser camera must be opened via HTTPS (except localhost) See HTTPS deployment details at `docs/SECURITY.md` § deployment

---

## Project Structure

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
├── db/
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
│   ├── member_identity.py       # Single resolver for all QR types (v3 ticket/v2 card/legacy) + time window policy
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
│   │   ├── member_view.py       # Two-column: form + searchable list
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
│   ├── app.py                   # Flask web application (routes, guards, DataTables /api feeds, bulk endpoints, /api/analytics, Excel import)
│   ├── csrf.py                  # Per-session CSRF tokens (form field or X-CSRF-Token header)
│   ├── static/
│   │   ├── style.css            # Web UI design system
│   │   ├── vendor/              # Vendored JS libs, pinned versions (see README.md inside)
│   │   │   ├── alpine.min.js    # Alpine.js 3.17.5
│   │   │   ├── jquery.min.js    # jQuery 3.7.1
│   │   │   ├── datatables.min.js# DataTables 1.13.8 (server-side tables)
│   │   │   ├── sweetalert2.min.js/.css  # SweetAlert2 11.26.25
│   │   │   ├── chart.umd.min.js # Chart.js 4.6.1
│   │   │   └── flatpickr.min.*
│   │   ├── fonts/               # Sarabun + Prompt TTFs (OFL) — @font-face in style.css, offline
│   │   ├── lang/
│   │   │   └── datatables-th.json # hand-made Thai i18n for DataTables
│   │   └── js/
│   │       ├── jsQR.js          # Vendored jsQR (in-browser QR decode)
│   │       ├── scan.js          # Camera capture + /scan/resolve calls
│   │       ├── member_scan.js   # Card scan camera on member activation page
│   │       ├── member_qr.js     # Countdown timer + switch QR when cycle changes
│   │       ├── tables.js        # DataTables configs (books/members/history) + bulk bar
│   │       └── ui.js            # PyLib helpers: dialogs, toasts, data-confirm, flatpickr init
│   └── templates/
│       ├── icons.svg            # Lucide SVG icon sprite
│       ├── base.html            # Layout + sidebar + vendor stack wiring
│       ├── login.html           # Login page
│       ├── member_base.html     # Member page (separate shell, no admin menu)
│       ├── member_login.html    # Member login (class+number+PIN)
│       ├── member_activate.html # First-time activation (scan card or admin code) + set PIN
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
├── tests/
│   ├── _harness.py              # Temp-DB/env isolation for tests
│   ├── test_security.py         # Password hashing + QR crypto
│   ├── test_qr_security.py      # QR payload formats & rejection paths
│   ├── test_migrations.py       # Schema v1→v4 upgrades + book-code dedupe + fresh layout
│   ├── test_desktop_paging.py   # Capped list-window math (large catalog guard)
│   ├── test_member_repo.py      # Member CRUD + PIN/activation-code methods
│   ├── test_rotating_tickets.py # `v3:` ticket crypto, window math, identity dispatch
│   ├── test_member_portal.py    # Member portal: activation, login, portal, admin PIN reset
│   ├── test_book_repo.py        # Book CRUD
│   ├── test_borrow_atomic.py    # Atomic borrow/return invariants
│   ├── test_admin_bootstrap.py  # No-default-password bootstrap
│   ├── test_backup_restore.py   # Backup + verify
│   ├── test_web_security.py     # CSRF/throttle/session
│   ├── test_web_scan.py         # Badge scanner endpoint
│   ├── test_web_delete_guards.py# History-bearing delete refusals
│   ├── test_web_book_retire.py  # Retire/restore/purge flow
│   ├── test_web_bulk.py         # 1000-row Excel import + DataTables APIs + bulk endpoints (cap/CSRF/skip counts)
│   ├── test_web_partials.py     # List-table skeleton/page APIs + header CSRF + analytics
│   └── test_web_pages_smoke.py  # Every authenticated page renders 200
├── docs/
│   ├── SECURITY.md              # Security runbook (key rotation, HTTPS/LAN)
│   └── DESIGN.md                # Visual design system notes
├── assets/
│   ├── fonts/                   # Sarabun + Prompt TTFs (desktop PDFs + copied to web/static/fonts)
│   ├── logos/                   # school_logo.png
│   ├── qrcodes/                 # Generated QR images
│   └── cards/                   # Generated member card PDFs
├── DB/                          # (runtime, gitignored) INITIAL_ADMIN_PASSWORD.txt lives here when created
├── .env.example                 # SECRET_PASSPHRASE, SECRET_SALT, FLASK_SECRET_KEY template
├── .env                         # Secrets (gitignored, auto-generated)
├── requirements.txt
└── README.md
```

---

## Settings

Open app → menu "Settings" or web → menu "Settings":

| Option | Description | Default |
|---|---|---|
| Theme | System / Light / Dark | System |
| Loan period | Default number of days when borrowing a book | 7 days |
| Max books per person | Limit number of books borrowed simultaneously | 3 books |
| Member card validity | Number of days before card expires | 365 days |
| School name | Displayed on member card and web page | DSNPRU |
| School logo | PNG file for member card and sidebar | — |
| Web UI host | 127.0.0.1 (loopback) or 0.0.0.0 (LAN + HTTPS) | 127.0.0.1 |
| Enable member portal | Enable/disable member page system-wide (disabled → `/member*` returns 404) | Enabled |
| Member QR rotation cycle (hours) | Time interval per cycle for rotating QR on member page (1–72) | 12 hrs |
| Neighboring cycle allowance | Number of allowance cycles (0–2) for clock skew between scanner/display QR machines | 1 cycle |
| Accept legacy card QR | Still accept non-expiring QR on plastic/paper cards (disable when member activation complete) | Enabled |

---

## Security System

| Feature | Method |
|---|---|
| Admin password | `bcrypt` hash/verify; no default — first boot creates random password and prints once; old 'admin123' automatically removed every boot |
| QR encryption | `AES-256-GCM` (`v2:` fixed card; `v3:` rotating ticket bind AAD) new; legacy `AES-256-CBC` decodes strictly until migrated via `tools/rotate_qr_key.py` |
| Key derivation | PBKDF2-HMAC-SHA256 (100k iters) from `.env` secrets auto-generated and placeholder-guarded |
| Rotating member ticket | Each cycle ticket contains `member_id`+`epoch`, server-side time window check (round time + allowance), no derived state — no persist needed |
| Login throttling | Web and desktop: 5 failed attempts per 15 minutes → temporary lockout (web admin per IP+username; member page per IP+class+number) |
| Member PIN code | bcrypt stored in `members.pin_hash` (NULL = not activated yet); first activation must go through real card or admin code — old code stores SHA-256 hash + expires 15 minutes + single use (atomic claim) |
| Web session | HttpOnly/SameSite=Lax cookie, CSRF token on every POST (constant-time compare), login timing-equalized against enumeration |
| SQL Injection | Parameterized queries everywhere, status compared via parameter |
| Data integrity | FK enforcement, UNIQUE index on books.code, borrow_log CHECK (returned IN (0,1)), borrow/return as single IMMEDIATE transaction |
| Thread safety | `BaseRepository` uses connection-per-query + `threading.Lock`, `transact()` for multi-statement |
| Schema migrations | `PRAGMA user_version` (v1 legacy → v2 constraints → v3 bulk-import indexes → v4 member portal: pin_hash + activation codes) runs automatically and idempotently every boot |
| Secrets | `.env` (gitignored), guards placeholder values, changes have runbook in `docs/SECURITY.md` |

---

## Additional Tools

```bash
# Rotate QR encryption keys (important: before next app launch when .env changes)
venv\Scripts\python tools\rotate_qr_key.py [--no-cards]

# Regenerate member card PDFs from current database
venv\Scripts\python tools\regenerate_member_cards.py
```

---

## Run Tests

```bash
# Windows
venv\Scripts\python -m unittest discover -s tests

# Linux/Mac
venv/bin/python -m unittest discover -s tests
```

Currently 193 test cases (18 test files) covering security, rotating tickets, member portal, repositories, migrations, atomic borrow/return, backups, bulk imports, DataTables list APIs + bulk endpoints, offline-safety (no CDN/external fonts), desktop list-paging and web guards.

---

## Database (SQLite3)

| Table | Description |
|---|---|
| `admin_users` | System admins (bcrypt hashed password, no default) |
| `members` | Library member data (name, class, number, QR, expiry date, `pin_hash` — member portal PIN hash NULL until self-activated) — badge key (class, number) must be unique in new system |
| `books` | Books in system (code [UNIQUE], name, status: `Available` / `Borrowed` / `Retired`) |
| `borrow_log` | Borrow/return records (returned CHECK 0/1, FK to members/books) |
| `access_log` | Library check-in/check-out records |
| `member_activation_codes` | One-time activation codes (stores only SHA-256 hash, expires 15 minutes, UNIQUE) — empty codes never stored |

---

## Libraries Used

| Package | Purpose |
|---|---|
| `customtkinter` | GUI framework (Modern Tkinter) |
| `flask` | Web UI framework |
| `opencv-python` | Webcam + QR scan |
| `pyzbar` | QR Code decode (desktop) |
| `Pillow` | Image management |
| `cryptography` | AES-256-GCM / AES-256-CBC encryption, PBKDF2 |
| `bcrypt` | Password hashing |
| `reportlab` | PDF creation (member cards, reports) |
| `PyMuPDF` (`pymupdf`) | PDF viewer in app |
| `numpy` | Image processing (opencv) |
| `qrcode` | QR Code creation |
| `pandas` | Excel import/export |
| `openpyxl` | Excel engine for pandas (.xlsx) |
| `python-dotenv` | Load `.env` |
| `packaging` | CustomTkinter dependency |

---

## Developers

- Miss Khwanchanok Unsiri No. 4 (UX/UI)
- Mr. Natnanon Rodnoi No. 30 (Developer)
- **School:** DSNPRU

---

## Acknowledgments

- Beta testers
- Advisors

---

**MIT License © 2025**