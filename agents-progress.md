# Project Refactoring & Handoff Progress Tracking

> **Status:** Initiated (Phase 0 Complete)  
> **Target codebase:** [main.py](file:///C:/Users/Nonx2/Documents/The-PY-LIB-Project/main.py) (2,500+ lines, CustomTkinter SQLite app)

---

## 🛠️ Completed Actions (Phase 0)
- [x] **Debug Mode Script Argument:** Added a dynamically injected debug wrapper to `main.py`. Running `python main.py --debug`:
  - Registers custom exception hooks (`sys.excepthook` & `threading.excepthook`) to capture crashes in both the GUI thread and camera/scan threads.
  - Dynamically inspects and wraps all methods of `LibraryApp` to log entry, exit, duration, and arguments.
  - Streams logs to standard output and logs to `debug.log`.
- [x] **Venv Verification:** Advised developer to execute all runs and packaging commands from a Python Virtual Environment (`venv`) to keep packaging sizes minimal and avoid global library conflicts.

---

## 📋 Refactoring Plan & To-Do List (For Next Agent)

Follow this sequence phase-by-phase. Tick off completed items as you proceed.

### 🔐 Phase 1: Security Implementation
- [x] **Bcrypt Admin Hashing:**
  - Automatically migrates existing plain-text admin database table rows to Bcrypt on application boot.
  - Login checks verify credentials dynamically using `bcrypt.checkpw`.
- [x] **AES-256-CBC QR Cryptography:**
  - Created `core/security.py` using cryptography library.
  - Derived a secure 256-bit key from passphrase and salt configurations via PBKDF2.
  - Encrypts generated member QR code details and decodes with legacy standard base64 fallback.
- [x] **Environment & Config Configuration:**
  - Added `.env` and `.env.example` templates to manage secret variables.
- [ ] **Login Rate-Limiter:**
  - Lockout users/IPs temporarily after $N$ consecutive failed login attempts.

### 💾 Phase 2: Data Repositories Layer
- [ ] **DB Schema Definition:**
  - Extract database creation routines to [db/schema.py](file:///C:/Users/Nonx2/Documents/The-PY-LIB-Project/db/schema.py).
- [ ] **Repository Classes:**
  - Build discrete classes: `MemberRepository`, `BookRepository`, `BorrowRepository`, and `AccessRepository` to contain all queries.
- [ ] **Thread-Safe Connections:**
  - Resolve potential SQLite SQLite-busy/race conditions during simultaneous scan-thread and UI-thread operations using safe thread-local connections or mutex locks.

### ⚙️ Phase 3: Services Layer
- [ ] **QR Service (`services/qr_service.py`):**
  - Consolidate encryption/decryption and generation logic.
- [ ] **PDF Service (`services/pdf_service.py`):**
  - Relocate PDF generator commands (ReportLab canvas generation) from the UI layer.
- [ ] **Camera Service (`services/camera_service.py`):**
  - Isolate OpenCV camera routines and worker loops. Provide frame updates and barcode payloads via a queue/callbacks.

### 🎨 Phase 4: UI View Layer Separation
- [ ] **View Separation:**
  - Transition individual screens into distinct modules under `ui/views/` (e.g. `login_view.py`, `dashboard_view.py`, etc.).
- [ ] **Application Shell (`app.py`):**
  - Implement a central frame navigation manager inside `app.py` to toggle visible frames cleanly, eliminating `winfo_children()` resets.

### 📊 Phase 5: Tests and Logging
- [ ] **Formal Logger Configuration:**
  - Replace printouts with structured logging configs.
- [ ] **Unit Tests:**
  - Develop tests using python's `unittest` or `pytest` to target repositories and security services.
