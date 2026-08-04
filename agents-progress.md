# Project Refactoring & Handoff Progress Tracking

> **Status:** Phase 1-4 Complete, Phase 5 Complete (verified)  
> **Target codebase:** original `main.py` (2,500+ lines, CustomTkinter SQLite app) → now modular; preserved as `legacy_main.py`

---

## 🛠️ Completed Actions

### Phase 0 — Baseline
- [x] **Debug Mode:** `python main.py --debug` wrapper with exception hooks + method tracing (kept in `legacy_main.py`; the new entry point uses structured logging).
- [x] **Venv:** `venv/` created and used for all runs/installs/tests.

### 🔐 Phase 1: Security Implementation
- [x] **Bcrypt Admin Hashing:** `core/security.py` (hash_password/verify_password); plain-text rows auto-migrated to bcrypt; default admin seeded on first boot (`db/schema.py:seed_admin`).
- [x] **AES-256-CBC QR Cryptography:** `core/security.py` PBKDF2-derived key, IV-prepended ciphertext; `decrypt_qr_data` keeps legacy base64 fallback.
- [x] **Environment & Config:** `.env` / `.env.example` (SECRET_PASSPHRASE, SECRET_SALT) loaded via python-dotenv.
- [ ] **Login Rate-Limiter:** lockout after N failed attempts — still NOT implemented (design decision pending).

### 💾 Phase 2: Data Repositories Layer
- [x] **DB Schema:** `db/schema.py` — all CREATE TABLEs + `seed_admin()` (bcrypt default admin + legacy migration).
- [x] **Repository Classes:** `db/repositories/{base,member,book,borrow,access}_repo.py` — parameterized queries only.
- [x] **Thread-Safe Connections:** `BaseRepository` opens a fresh `check_same_thread=False` connection per query under a shared `threading.Lock`.

### ⚙️ Phase 3: Services Layer
- [x] **QR Service:** `services/qr_service.py` (generate_qr encrypts with AES, decode_qr).
- [x] **PDF Service:** `services/pdf_service.py` (member cards + borrow/access history reports).
- [x] **Camera Service:** `services/camera_service.py` (frame processing + QR decode loop; UI subscribes via queue).
- [x] **Excel Service:** `services/excel_service.py` (import books from Excel).

### 🎨 Phase 4: UI View Layer Separation
- [x] **Views:** `ui/views/` — login, dashboard, member, book, borrow, return, history, access (scanner + history), settings. Each receives `parent` + `nav` (simple DI), packs itself.
- [x] **App Shell:** `app.py` — `LibraryApp(ctk.CTk)` navigation controller: login → sidebar + content area, `_show_view()` swap (no more `winfo_children()` rebuilds), `show_*` methods for every screen.
- [x] **Entry Point:** `main.py` is now thin (init_db → LibraryApp().run()); monolith preserved as `legacy_main.py` for rollback.
- [x] **Dependencies:** `requirements.txt` updated with reportlab/pandas/numpy/dotenv; numpy pinned <2 (opencv 4.8 incompat). Installed and verified in `venv/`.

### 📊 Phase 5: Tests and Logging
- [x] **Logger:** `core/logger.py` (console + file, DEBUG flag via `setup_logger`), replaces prints.
- [x] **Unit Tests:** `tests/test_security.py`, `test_member_repo.py`, `test_book_repo.py` — 14 tests, all pass with `venv/Scripts/python -m unittest discover -s tests` (temp-DB isolation).
- [x] **GUI smoke test:** boot → login → all 9 navigation views verified.

---

## 📋 Remaining Work

- [ ] **Login Rate-Limiter** (Phase 1, unfinished): lockout after N failed attempts.
- [ ] **View refresh patterns:** some views build once in `__init__`; verify data refresh on re-entry to match old behavior.
- [ ] **Cancel stale `after()` callbacks** on view destroy (currently benign warnings).
- [ ] **Commit per phase** (nothing committed yet — repo currently has the old single commit).

## ⚠️ Notes for next agent
- Run everything from `venv/`: `& venv/Scripts/python.exe main.py`
- Existing data in `db/library.db` is preserved (schema is CREATE-IF-NOT-EXISTS + seed-only-when-empty).
- Legacy monolith lives in `legacy_main.py` — do not delete until the modular app has been user-accepted.
- `.env` is gitignored; keep `.env.example` as the template.
