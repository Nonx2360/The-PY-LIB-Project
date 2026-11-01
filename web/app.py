# web/app.py - Flask web UI for THE PY-LIB
import io
import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from flask import (Flask, render_template, request, redirect, url_for,
                   jsonify, session, flash, send_file)
from config import settings
from config.constants import (BOOK_STATUS_AVAILABLE, BOOK_STATUS_BORROWED,
                              BOOK_STATUS_RETIRED)
from config.env_store import ensure_secrets
from db.schema import init_db
from db.repositories.member_repo import MemberRepository
from db.repositories.book_repo import BookRepository
from db.repositories.borrow_repo import BorrowRepository, BorrowError
from db.repositories.access_repo import AccessRepository
from db.repositories.admin_repo import AdminRepository
from services.excel_service import ExcelService
from services.excel_service import MAX_ROWS as IMPORT_MAX_ROWS
from services.qr_service import QRService
from services.pdf_service import PDFService
from core.logger import logger
from core import throttle

from web import csrf as csrf_mod
from web.csrf import csrf_field_markup

from functools import wraps
from markupsafe import Markup
import sqlite3


def _rows_to_dicts(rows, columns):
    return [dict(zip(columns, row, strict=False)) for row in rows]


# Server-side field limits (mirror the <input maxlength> guidance)
LIMITS = {
    "member_name": 120, "member_grade": 40, "member_number": 20,
    "book_code": 40, "book_title": 200, "school_name": 80, "logo_path": 260,
}

def _clean(value, limit):
    text = "".join(c for c in (value or "") if ord(c) >= 32)
    text = text.strip()
    return text[:limit] if limit else text


def _int_in(raw, lo, hi, default):
    """Parse an int with bounds; falls back to default on any failure."""
    try:
        value = int(raw)
    except (TypeError, ValueError):
        return default
    if value < lo or value > hi:
        return default
    return value


def create_app():
    init_db()  # run schema migrations + admin bootstrap on web-only boots too

    secrets_dict = ensure_secrets()

    app = Flask(__name__, static_folder="static", template_folder="templates")
    app.secret_key = secrets_dict["FLASK_SECRET_KEY"]
    app.config.update(
        SESSION_COOKIE_NAME="pylib_session",
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        # NOTE: set True behind HTTPS (see docs/SECURITY.md § deployment)
        SESSION_COOKIE_SECURE=False,
        PERMANENT_SESSION_LIFETIME=timedelta(hours=4),
        # Excel bulk imports: a 5,000-row .xlsx is well under this — the cap
        # is a DoS guard, not a feature limit. Exceeding it hits the 413
        # handler below, which tells the user to split the file.
        MAX_CONTENT_LENGTH=8 * 1024 * 1024,
    )

    @app.route("/assets/<path:filename>")
    def serve_assets(filename):
        from flask import send_from_directory
        return send_from_directory(os.path.join(app.root_path, ".."), "assets/" + filename)

    member_repo = MemberRepository()
    book_repo = BookRepository()
    borrow_repo = BorrowRepository()
    access_repo = AccessRepository()
    admin_repo = AdminRepository()

    def login_required(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if not session.get("logged_in"):
                return redirect(url_for("login"))
            return f(*args, **kwargs)
        return decorated

    @app.context_processor
    def _inject_csrf():
        try:
            return {
                "csrf_token": csrf_mod._get_token(session),
                "csrf_field": csrf_field_markup(session),
            }
        except RuntimeError:  # outside request context
            return {"csrf_token": "", "csrf_field": Markup("")}

    @app.before_request
    def _csrf_protect():
        """Reject state-changing POSTs without a valid session CSRF token."""
        if request.method != "POST":
            return None
        if request.endpoint == "login":
            # login is rendered without login_required; session token still required
            if not csrf_mod.validate(session, request.form):
                flash("ระดับเซสชันไม่ถูกต้อง กรุณาลองอีกครั้ง", "error")
                return redirect(url_for("login"))
            return None
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        # form field first (regular pages); X-CSRF-Token header for AJAX
        # requests (the list-table engine and its bulk endpoints)
        token = (request.form.get("csrf_token")
                 or request.headers.get("X-CSRF-Token", ""))
        if not csrf_mod.validate(session, token):
            logger.warning("csrf: rejected POST (missing/invalid token) %s", request.path)
            flash("เซสชันหมดอายุหรือโทเคนไม่ถูกต้อง กรุณาลองอีกครั้ง", "error")
            # AJAX callers (fetch) get a clean status instead of a redirect
            if request.headers.get("HX-Request") == "true":
                return "", 403
            return redirect(request.referrer or url_for("dashboard"))
        return None

    @app.errorhandler(413)
    def _payload_too_large(_e):
        # MAX_CONTENT_LENGTH exceeded (e.g. an oversized Excel upload)
        flash("ไฟล์ใหญ่เกิน 8 MB — กรุณาแบ่งการนำเข้าเป็นหลายไฟล์", "error")
        return redirect(url_for("books"))

    # ── DataTables server-side helpers ────────────────────
    # List pages are driven by jQuery DataTables (vendored) in server-side
    # mode: the browser sends draw/start/length/search/order, the API
    # answers {draw, recordsTotal, recordsFiltered, data}. Everything is
    # bounded so a 10,000-row catalog costs the same as a 10-row one.

    _DT_LENGTHS = (10, 25, 50, 100)

    def _dt_args():
        """Parse + bound the DataTables server-side protocol parameters."""
        draw = _int_in(request.args.get("draw"), 1, 10**7, 1)
        start = _int_in(request.args.get("start"), 0, 10**7, 0)
        length = _int_in(request.args.get("length"), 1, 10**4, 50)
        if length not in _DT_LENGTHS:          # only values our UI offers
            length = 50
        if start > 10**7 - length:
            start = 0
        order_col = _int_in(request.args.get("order[0][column]"), 0, 20, 0)
        order_dir = (request.args.get("order[0][dir]") or "desc").lower()
        if order_dir not in ("asc", "desc"):
            order_dir = "asc"
        q = _clean(request.args.get("search[value]"), 80)
        return {"draw": draw, "start": start, "length": length,
                "order_col": order_col, "order_dir": order_dir, "q": q}

    def _dt_response(dt, repo, base, where, params, select_sql, columns):
        """Assemble the JSON payload + run the bounded page query.

        `columns` is a list of (SQL identifier, dict key) pairs — the key at
        dt['order_col'] must be orderable, anything else falls back to the
        first column (id, newest first). When `start` overshoots the
        filtered row count (e.g. a bulk delete emptied the current page) the
        page is clamped back to the last one instead of showing nothing.
        """
        total = repo.fetchone(f"SELECT COUNT(*) FROM {base}")[0]
        where_sql = f" WHERE {where}" if where else ""
        filtered = repo.fetchone(
            f"SELECT COUNT(*) FROM {base}{where_sql}", tuple(params))[0]
        order_col = min(dt["order_col"], len(columns) - 1)
        order_sql = columns[order_col][0]
        start = dt["start"]
        length = dt["length"]
        # page clamp: a bulk action can leave the current page number past
        # the end — jump back to the last full page instead of a blank one
        if filtered > 0 and start >= filtered:
            start = max(0, ((filtered - 1) // length) * length)
        rows = repo.fetchall(
            select_sql + where_sql +
            f" ORDER BY {order_sql} {dt['order_dir'].upper()} "
            f"LIMIT {length} OFFSET {start}",
            tuple(params))
        return jsonify({
            "draw": dt["draw"],
            "recordsTotal": total,
            "recordsFiltered": filtered,
            "data": _rows_to_dicts(rows, [key for _, key in columns]),
        })

    def _bulk_ids(payload):
        """Validate a bulk-action id list; returns (ids, error_message)."""
        raw = payload.get("ids") if isinstance(payload, dict) else None
        ids, invalid = [], 0
        if isinstance(raw, list):
            for v in raw:
                try:
                    i = int(v)
                except (TypeError, ValueError):
                    invalid += 1
                    continue
                if 1 <= i <= 10**9:
                    ids.append(i)
        if len(ids) > 200:                       # keep one batch lightweight
            return [], "เลือกได้ครั้งละสูงสุด 200 รายการ"
        return sorted(set(ids)), None

    # ── Auth ──────────────────────────────────────────────
    @app.route("/login", methods=["GET", "POST"])
    def login():
        if request.method == "POST":
            username = _clean(request.form.get("username"), 40)
            password = request.form.get("password", "")
            ip = request.remote_addr or "-local-"
            key = (ip, username)

            if throttle.is_locked(key):
                minutes = throttle.lock_remaining_minutes(key)
                logger.warning("web login: throttled attempt "
                               "(locked %ss remaining)",
                               throttle.lock_remaining_seconds(key))
                # one generic message for throttled AND wrong creds (no user enumeration)
                flash(f"ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง (ลองใหม่ได้ในอีก ~{minutes} นาที)", "error")
                return render_template("login.html", config=settings.all()), 429

            if admin_repo.verify_credentials(username, password):
                throttle.record_success(key)
                session["logged_in"] = True
                session["username"] = username
                logger.info("web login: successful")  # no usernames/PII in logs
                return redirect(url_for("dashboard"))

            throttle.record_failure(key)
            logger.warning("web login: failed attempt (count kept in-process)")
            flash("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง", "error")
        return render_template("login.html", config=settings.all())

    @app.route("/logout")
    def logout():
        session.clear()
        return redirect(url_for("login"))

    # ── Dashboard ─────────────────────────────────────────
    @app.route("/")
    @login_required
    def dashboard():
        members = member_repo.fetchall("SELECT COUNT(*) FROM members")[0][0]
        books = book_repo.fetchall("SELECT COUNT(*) FROM books")[0][0]
        borrowed = borrow_repo.fetchall(
            "SELECT COUNT(*) FROM borrow_log WHERE returned=0")[0][0]
        overdue = borrow_repo.fetchall(
            "SELECT COUNT(*) FROM borrow_log WHERE returned=0 AND return_due < date('now')"
        )[0][0]
        return render_template("index.html",
                               members=members, books=books,
                               borrowed=borrowed, overdue=overdue)

    # ── Members ───────────────────────────────────────────
    @app.route("/members")
    @login_required
    def members():
        total = member_repo.fetchone("SELECT COUNT(*) FROM members")[0]
        return render_template("members.html", total=total)

    @app.route("/api/members")
    @login_required
    def api_members():
        """DataTables server-side feed for the members table."""
        dt = _dt_args()
        base = "members"
        where, params = "", ()
        if dt["q"]:
            like = f"%{dt['q']}%"
            where = "name LIKE ? OR grade LIKE ? OR number LIKE ?"
            params = (like, like, like)
        return _dt_response(
            dt, member_repo, base, where, params,
            "SELECT id, name, grade, number, register_date, expire_date, "
            "qrcode_path FROM members",
            [("id", "id"), ("name", "name"), ("grade", "grade"),
             ("number", "number"), ("register_date", "register_date"),
             ("expire_date", "expire_date"), ("qrcode_path", "qrcode_path")],
        )

    @app.route("/members/bulk", methods=["POST"])
    @login_required
    def members_bulk():
        """Bulk member management: only 'delete' for now.

        Referenced members (borrow/access history) are skipped by
        delete_many, never erroring the batch.
        """
        payload = request.get_json(silent=True) or {}
        action = payload.get("action")
        ids, err = _bulk_ids(payload)
        if err:
            return jsonify({"status": "error", "message": err}), 400
        if action != "delete":
            return jsonify({"status": "error",
                            "message": "รู้จักเฉพาะคำสั่ง 'delete' เท่านั้น"}), 400
        deleted, skipped = member_repo.delete_many(ids)
        logger.info("Web: bulk member delete: requested=%d deleted=%d skipped=%d",
                    len(ids), deleted, skipped)
        return jsonify({"status": "ok", "deleted": deleted, "skipped": skipped})

    @app.route("/members/add", methods=["POST"])
    @login_required
    def member_add():
        name = _clean(request.form.get("name"), LIMITS["member_name"])
        grade = _clean(request.form.get("grade"), LIMITS["member_grade"])
        number = _clean(request.form.get("number"), LIMITS["member_number"])
        if not all([name, grade, number]):
            flash("กรุณากรอกข้อมูลให้ครบ", "error")
            return redirect(url_for("members"))
        if member_repo.find_by_badge(grade, number):
            flash(f"ชั้น {grade} มีสมาชิกเลขที่ {number} อยู่แล้ว "
                  "ไม่สามารถลงทะเบียนซ้ำได้ (เลขที่ต้องไม่ซ้ำภายในชั้นเดียวกัน)",
                  "error")
            return redirect(url_for("members"))
        reg_date = datetime.now().strftime("%Y-%m-%d")
        exp_date = (datetime.now() + timedelta(days=settings.get("member_expiry_days"))).strftime("%Y-%m-%d")
        qr_data = f"{name}|{grade}|{number}|{reg_date}|{exp_date}"
        qr_filename = f"assets/qrcodes/{os.urandom(8).hex()}.png"
        os.makedirs("assets/qrcodes", exist_ok=True)
        os.makedirs("assets/cards", exist_ok=True)
        encrypted = QRService.generate_qr(qr_data, qr_filename)
        member_id = member_repo.add(name, grade, number, reg_date, exp_date, encrypted, qr_filename)
        card_pdf = f"assets/cards/member_{member_id}.pdf"
        PDFService.generate_member_card_pdf(name, number,
                                            "assets/logos/school_logo.png",
                                            card_pdf, encrypted)
        logger.info("Web: member added: id=%s", member_id)  # student names are PII — not logged
        flash("เพิ่มสมาชิกสำเร็จ", "success")
        return redirect(url_for("members"))

    @app.route("/members/<int:mid>/delete", methods=["POST"])
    @login_required
    def member_delete(mid):
        # borrow_log/access_log carry FOREIGN KEYs to members: deleting a
        # member with history would corrupt library records AND crash as 500.
        borrowed = borrow_repo.fetchone(
            "SELECT COUNT(*) FROM borrow_log WHERE member_id=?", (mid,))[0]
        if borrowed > 0:
            flash("ไม่สามารถลบได้ — สมาชิกนี้มีประวัติการยืมหนังสือ "
                  "ระบบเก็บประวัติไว้เพื่อความถูกต้องของข้อมูล", "error")
            return redirect(url_for("members"))
        accessed = access_repo.fetchone(
            "SELECT COUNT(*) FROM access_log WHERE member_id=?", (mid,))[0]
        if accessed > 0:
            flash("ไม่สามารถลบได้ — สมาชิกนี้มีประวัติการเข้าออกห้องสมุด", "error")
            return redirect(url_for("members"))
        row = member_repo.fetchone("SELECT qrcode_path FROM members WHERE id=?", (mid,))
        try:
            member_repo.delete(mid)
        except sqlite3.IntegrityError:
            logger.info("web member delete: DB refused id=%s", mid)
            flash("ลบไม่สำเร็จ — ข้อมูลยังถูกอ้างอิงจากประวัติการใช้งาน", "error")
            return redirect(url_for("members"))
        # remove card artifacts only AFTER the row is really gone, so a failed
        # delete never loses the QR image of a member that still exists
        if row and row[0]:
            try:
                os.remove(row[0])
            except OSError:
                pass
        logger.info("Web: member deleted: id=%s", mid)
        flash("ลบสมาชิกสำเร็จ", "success")
        return redirect(url_for("members"))

    # ── Books ─────────────────────────────────────────────
    @app.route("/books")
    @login_required
    def books():
        total = book_repo.fetchone("SELECT COUNT(*) FROM books")[0]
        return render_template("books.html", total=total)

    @app.route("/api/books")
    @login_required
    def api_books():
        """DataTables server-side feed for the books table.

        `status` (optional) is an exact-match server-side filter wired to
        the page's status <select>; the free-text search goes through the
        standard search[value] parameter from DataTables.
        """
        dt = _dt_args()
        status = (request.args.get("status") or "").strip()
        base = "books bo"
        where, params = [], []
        if dt["q"]:
            like = f"%{dt['q']}%"
            where.append("(bo.code LIKE ? OR bo.title LIKE ?)")
            params += [like, like]
        if status:
            where.append("bo.status = ?")
            params.append(status)
        cond = " AND ".join(where)
        select = ("SELECT bo.id, bo.code, bo.title, bo.status, "
                  "(SELECT COUNT(*) FROM borrow_log bl WHERE bl.book_id=bo.id) "
                  "AS history FROM books bo")
        return _dt_response(
            dt, book_repo, base, cond, params, select,
            [("bo.id", "id"), ("bo.code", "code"), ("bo.title", "title"),
             ("bo.status", "status"), ("history", "history")],
        )

    @app.route("/books/bulk", methods=["POST"])
    @login_required
    def books_bulk():
        """Bulk book management: retire / restore / delete-many in JSON.

        delete is the collective version of the per-row flow: retired
        copies purge together with their borrow_log rows, plain copies
        delete only when they have no borrow history, everything else is
        skipped and reported.
        """
        payload = request.get_json(silent=True) or {}
        action = payload.get("action")
        ids, err = _bulk_ids(payload)
        if err:
            return jsonify({"status": "error", "message": err}), 400
        if action == "retire":
            changed = book_repo.bulk_retire(ids)
            out = {"status": "ok", "changed": changed, "skipped": len(ids) - changed}
        elif action == "restore":
            changed = book_repo.bulk_restore(ids)
            out = {"status": "ok", "changed": changed, "skipped": len(ids) - changed}
        elif action == "delete":
            deleted, skipped = book_repo.delete_many(ids)
            out = {"status": "ok", "deleted": deleted, "skipped": skipped}
        else:
            return jsonify({"status": "error",
                            "message": "รู้จักคำสั่ง retire/restore/delete เท่านั้น"}), 400
        logger.info("Web: bulk books: action=%s requested=%d", action, len(ids))
        return jsonify(out)

    @app.route("/books/add", methods=["POST"])
    @login_required
    def book_add():
        code = _clean(request.form.get("code"), LIMITS["book_code"])
        title = _clean(request.form.get("title"), LIMITS["book_title"])
        if not code or not title:
            flash("กรุณากรอกข้อมูลให้ครบ", "error")
            return redirect(url_for("books"))
        existing = book_repo.fetchone("SELECT id FROM books WHERE code=?", (code,))
        if existing:
            flash(f"รหัสหนังสือ '{code}' มีอยู่แล้ว", "error")
            return redirect(url_for("books"))
        book_repo.add(code, title, status=BOOK_STATUS_AVAILABLE)
        logger.info("Web: book added: code=%s", code)
        flash("เพิ่มหนังสือสำเร็จ", "success")
        return redirect(url_for("books"))

    @app.route("/books/import", methods=["POST"])
    @login_required
    def book_import():
        """Bulk import from the Thai-template .xlsx (1000+ rows is normal).

        Even huge imports are safe here because everything lands in ONE
        transaction (book_repo.add_many): one commit, duplicates skipped,
        never a partial catalog on error.
        """
        file = request.files.get("file")
        if file is None or not file.filename:
            flash("กรุณาเลือกไฟล์ Excel (.xlsx)", "error")
            return redirect(url_for("books"))
        if not file.filename.lower().endswith(".xlsx"):
            flash("รองรับเฉพาะไฟล์ .xlsx — ใช้แม่แบบที่ระบบมีให้", "error")
            return redirect(url_for("books"))
        try:
            # Werkzeug's SpooledTemporaryFile (py3.10) lacks .seekable(),
            # which openpyxl/zipfile needs — buffer the upload into BytesIO
            # first (bounded by MAX_CONTENT_LENGTH, so this stays small).
            payload = io.BytesIO(file.read())
            rows = ExcelService.read_books_from_excel(payload)
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("books"))
        except Exception:
            logger.error("Web: excel import read failed", exc_info=True)
            flash("อ่านไฟล์ไม่สำเร็จ กรุณาตรวจสอบรูปแบบไฟล์", "error")
            return redirect(url_for("books"))
        if not rows:
            flash("ไฟล์ไม่มีข้อมูลหนังสือที่ถูกต้อง", "error")
            return redirect(url_for("books"))

        # Enforce the same field limits the manual form applies
        clean, invalid = [], 0
        for row in rows:
            code = _clean(str(row.get("code", "")), LIMITS["book_code"])
            title = _clean(str(row.get("title", "")), LIMITS["book_title"])
            if not code or not title:
                invalid += 1
                continue
            clean.append((code, title))
        try:
            inserted, skipped_existing = book_repo.add_many(clean)
        except sqlite3.Error:
            logger.error("Web: excel import failed", exc_info=True)
            flash("นำเข้าไม่สำเร็จ กรุณาลองใหม่", "error")
            return redirect(url_for("books"))

        msg = f"นำเข้าหนังสือสำเร็จ {inserted} เล่ม"
        notes = []
        if skipped_existing:
            notes.append(f"รหัสซ้ำข้าม {skipped_existing} รายการ")
        if invalid:
            notes.append(f"แถวไม่สมบูรณ์ {invalid} แถว")
        if notes:
            msg += " (" + " • ".join(notes) + ")"
        logger.info("Web: excel import: inserted=%d skipped=%d invalid=%d",
                    inserted, skipped_existing, invalid)
        flash(msg, "success" if inserted else "warning")
        return redirect(url_for("books"))

    @app.route("/books/template")
    @login_required
    def book_template():
        """Download the import template (same columns as ExcelService reads)."""
        buf = io.BytesIO()
        ExcelService.generate_template_df().to_excel(buf, index=False)
        buf.seek(0)
        return send_file(
            buf,
            mimetype=("application/vnd.openxmlformats-"
                      "officedocument.spreadsheetml.sheet"),
            as_attachment=True,
            download_name="book_import_template.xlsx")

    @app.route("/books/<int:bid>/delete", methods=["POST"])
    @login_required
    def book_delete(bid):
        book = book_repo.fetchone("SELECT status FROM books WHERE id=?", (bid,))
        if book is None:
            flash("ไม่พบหนังสือในระบบ", "error")
            return redirect(url_for("books"))
        if book[0] == BOOK_STATUS_RETIRED:
            # Deliberate two-step purge: a RETIRED copy is removed together
            # with its borrow_log rows inside ONE transaction. The librarian
            # chose withdrawal first; this path makes the history loss an
            # explicit, confirmed decision — never reachable by one click
            # on an active catalog entry.
            def _purge(conn):
                conn.execute("DELETE FROM borrow_log WHERE book_id=?", (bid,))
                conn.execute("DELETE FROM books WHERE id=?", (bid,))
            try:
                borrow_repo.transact(_purge)
            except sqlite3.Error:
                logger.info("web book purge failed: id=%s", bid)
                flash("ลบไม่สำเร็จ กรุณาลองใหม่", "error")
                return redirect(url_for("books"))
            logger.info("Web: retired book purged with its history: id=%s", bid)
            flash("ลบหนังสือและประวัติการยืมของเล่มนี้แบบถาวรแล้ว", "success")
            return redirect(url_for("books"))
        active = borrow_repo.fetchone(
            "SELECT COUNT(*) FROM borrow_log WHERE book_id=? AND returned=0", (bid,))[0]
        if active > 0:
            flash("ไม่สามารถลบได้ — มีการยืมหนังสือเล่มนี้อยู่", "error")
            return redirect(url_for("books"))
        # borrow_log -> books is a FOREIGN KEY (data-integrity phase): a book
        # with returned history may never be deleted while in circulation.
        # Point the librarian at the sanctioned path: retire first.
        history = borrow_repo.fetchone(
            "SELECT COUNT(*) FROM borrow_log WHERE book_id=?", (bid,))[0]
        if history > 0:
            flash("ไม่สามารถลบได้ — เล่มนี้มีประวัติการยืม "
                  "กดปุ่ม หยุด (หยุดให้บริการ) ก่อน แล้วจึงลบถาวรได้", "error")
            return redirect(url_for("books"))
        try:
            book_repo.delete(bid)
        except sqlite3.IntegrityError:
            logger.info("web book delete: DB refused id=%s", bid)
            flash("ลบไม่สำเร็จ — ข้อมูลยังถูกอ้างอิงจากประวัติการใช้งาน", "error")
            return redirect(url_for("books"))
        logger.info("Web: book deleted: id=%s", bid)
        flash("ลบหนังสือสำเร็จ", "success")
        return redirect(url_for("books"))

    @app.route("/books/<int:bid>/retire", methods=["POST"])
    @login_required
    def book_retire(bid):
        """Withdraw a book from circulation — history is kept, deletion stays
        blocked by FK, but the book leaves every borrow flow until restored."""
        book = book_repo.fetchone("SELECT status FROM books WHERE id=?", (bid,))
        if book is None:
            flash("ไม่พบหนังสือในระบบ", "error")
        elif book[0] == BOOK_STATUS_AVAILABLE:
            book_repo.update_status(bid, BOOK_STATUS_RETIRED)
            logger.info("Web: book retired: id=%s", bid)
            flash("หยุดให้บริการแล้ว — ประวัติการยืมยังถูกเก็บไว้", "success")
        elif book[0] == BOOK_STATUS_RETIRED:
            flash("หนังสือเล่มนี้หยุดให้บริการอยู่แล้ว", "error")
        else:
            flash("หนังสือกำลังถูกยืมอยู่ กรุณารอคืนก่อน แล้วจึงหยุดให้บริการ", "error")
        return redirect(url_for("books"))

    @app.route("/books/<int:bid>/restore", methods=["POST"])
    @login_required
    def book_restore(bid):
        """Put a retired book back into circulation (status BOOK_STATUS_AVAILABLE)."""
        book = book_repo.fetchone("SELECT status FROM books WHERE id=?", (bid,))
        if book is None:
            flash("ไม่พบหนังสือในระบบ", "error")
        elif book[0] == BOOK_STATUS_RETIRED:
            book_repo.update_status(bid, BOOK_STATUS_AVAILABLE)
            logger.info("Web: book restored: id=%s", bid)
            flash("หนังสือกลับมาให้บริการตามปกติแล้ว", "success")
        else:
            flash("หนังสือเล่มนี้กำลังให้บริการอยู่แล้ว", "error")
        return redirect(url_for("books"))

    # ── Borrow ────────────────────────────────────────────
    @app.route("/borrow", methods=["GET", "POST"])
    @login_required
    def borrow():
        max_books = settings.get("max_books_per_member")
        if request.method == "POST":
            try:
                member_id = int(request.form.get("member_id", 0))
                book_id = int(request.form.get("book_id", 0))
            except (TypeError, ValueError):
                flash("กรุณาเลือกสมาชิกและหนังสือ", "error")
                return redirect(url_for("borrow"))

            due = request.form.get("due_date", "")
            try:
                due_dt = datetime.strptime(due, "%Y-%m-%d")
                if due_dt.date() < datetime.now().date():
                    flash("กำหนดคืนต้องไม่เป็นวันที่ผ่านมาแล้ว", "error")
                    return redirect(url_for("borrow"))
            except ValueError:
                flash("รูปแบบวันที่ไม่ถูกต้อง (YYYY-MM-DD)", "error")
                return redirect(url_for("borrow"))

            try:
                new_id = borrow_repo.try_borrow(member_id, book_id,
                                                datetime.now().strftime("%Y-%m-%d"), due)
                logger.info("Web: borrowed book %s to member %s (record %s)",
                            book_id, member_id, new_id)
                flash("ยืมหนังสือสำเร็จ", "success")
            except BorrowError as exc:
                flash(str(exc), "error")
            return redirect(url_for("borrow"))

        raw_members = member_repo.fetchall(
            "SELECT id, name, grade, number FROM members ORDER BY name")
        enriched = []
        for m in raw_members:
            mid = m[0]
            active = borrow_repo.fetchone(
                "SELECT COUNT(*) FROM borrow_log WHERE member_id=? AND returned=0", (mid,))[0]
            overdue = borrow_repo.fetchone(
                "SELECT COUNT(*) FROM borrow_log WHERE member_id=? AND returned=0 AND return_due < date('now')",
                (mid,))[0]
            enriched.append({"id": mid, "name": m[1], "grade": m[2], "number": m[3],
                             "active_borrows": active, "overdue_borrows": overdue})
        books_ = book_repo.fetchall(
            "SELECT id, code, title FROM books WHERE status = ? ORDER BY title",
            (BOOK_STATUS_AVAILABLE,))
        books_ = _rows_to_dicts(books_, ["id", "code", "title"])
        default_due = (datetime.now() + timedelta(days=settings.get("default_loan_days"))).strftime("%Y-%m-%d")
        return render_template("borrow.html", members=enriched, books=books_,
                               default_due=default_due, max_books=max_books)

    # ── Return ────────────────────────────────────────────
    @app.route("/return", methods=["GET", "POST"])
    @login_required
    def return_book():
        if request.method == "POST":
            try:
                record_id = int(request.form.get("record_id", 0))
            except (TypeError, ValueError):
                record_id = 0
            # NOTE: the submitted book_id is never trusted — the book id is
            # derived from the borrow_log row itself inside try_return().
            try:
                borrow_repo.try_return(record_id)
                logger.info("Web: returned record_id=%s", record_id)
                flash("คืนหนังสือสำเร็จ", "success")
            except BorrowError as exc:
                flash(str(exc), "error")
            return redirect(url_for("return_book"))

        member_filter = request.args.get("member_id", type=int)
        sql = """SELECT bl.id, m.name, m.grade, m.number, b.id as book_id, b.code, b.title,
                        bl.borrow_date, bl.return_due
                 FROM borrow_log bl
                 JOIN members m ON bl.member_id=m.id
                 JOIN books b ON bl.book_id=b.id
                 WHERE bl.returned=0"""
        params = []
        if member_filter:
            sql += " AND bl.member_id=?"
            params.append(member_filter)
        sql += " ORDER BY bl.borrow_date DESC"
        borrowed = borrow_repo.fetchall(sql, tuple(params))
        borrowed = _rows_to_dicts(borrowed, ["id", "name", "grade", "number", "book_id",
                                             "code", "title", "borrow_date", "return_due"])
        today = datetime.now().strftime("%Y-%m-%d")
        for b in borrowed:
            b["overdue"] = b["return_due"] < today
        filter_member = None
        if member_filter:
            frow = member_repo.fetchone(
                "SELECT id, name, grade, number FROM members WHERE id=?", (member_filter,))
            if frow:
                filter_member = {"id": frow[0], "name": frow[1],
                                 "grade": frow[2], "number": frow[3]}
        return render_template("return_book.html", borrowed=borrowed,
                               filter_member=filter_member)

    # ── History ───────────────────────────────────────────
    @app.route("/history")
    @login_required
    def history():
        total = borrow_repo.fetchone("SELECT COUNT(*) FROM borrow_log")[0]
        return render_template("history.html", total=total)

    @app.route("/api/history")
    @login_required
    def api_history():
        """DataTables server-side feed for the borrow-history table."""
        dt = _dt_args()
        base = ("borrow_log bl JOIN members m ON bl.member_id=m.id "
                "JOIN books b ON bl.book_id=b.id")
        where, params = "", ()
        if dt["q"]:
            like = f"%{dt['q']}%"
            where = ("m.name LIKE ? OR m.grade LIKE ? "
                     "OR b.code LIKE ? OR b.title LIKE ?")
            params = (like, like, like, like)
        return _dt_response(
            dt, borrow_repo, base, where, params,
            "SELECT bl.id, m.name, m.grade, b.code, b.title, "
            "bl.borrow_date, bl.return_due, bl.returned "
            "FROM borrow_log bl "
            "JOIN members m ON bl.member_id=m.id "
            "JOIN books b ON bl.book_id=b.id",
            [("bl.id", "id"), ("m.name", "name"), ("m.grade", "grade"),
             ("b.code", "code"), ("b.title", "title"),
             ("bl.borrow_date", "borrow_date"), ("bl.return_due", "return_due"),
             ("bl.returned", "returned")],
        )

    # ── Access History + Recording ──────────────────────────
    @app.route("/access")
    @login_required
    def access():
        rows = access_repo.get_recent_access(limit=100)
        records = _rows_to_dicts(rows, ["access_time", "name", "grade", "action"])
        members_ = member_repo.fetchall("SELECT id, name, grade, number FROM members ORDER BY name")
        members_ = _rows_to_dicts(members_, ["id", "name", "grade", "number"])
        return render_template("access.html", records=records, members=members_)

    @app.route("/access/record", methods=["POST"])
    @login_required
    def access_record():
        try:
            member_id = int(request.form.get("member_id", 0))
        except (TypeError, ValueError):
            member_id = 0
        if not member_id:
            flash("กรุณาเลือกสมาชิก", "error")
            return redirect(url_for("access"))
        last = access_repo.fetchone(
            "SELECT action FROM access_log WHERE member_id=? ORDER BY id DESC LIMIT 1",
            (member_id,))
        action = "ออก" if (last and last[0] == "เข้า") else "เข้า"
        access_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        access_repo.log_access(member_id, access_time, action)
        logger.info("Web: access recorded: member_id=%s -> %s", member_id, action)
        flash(f"บันทึก{action}สำเร็จ", "success")
        return redirect(url_for("access"))

    # ── Card scan (QR badge): shared by borrow/return/access pages ────────
    # The camera itself never decrypts anything: POST /scan/resolve takes the
    # QR payload text (AES-sealed) and decrypts server-side with the .env key.
    # Each attempt costs one PBKDF2 derivation, and the route is fully wrapped
    # by login_required + CSRF before_request — brute force is impractical.
    @app.route("/scan/resolve", methods=["POST"])
    @login_required
    def scan_resolve():
        payload = _clean(request.form.get("payload"), 800)
        if not payload:
            return jsonify(ok=False, error="invalid"), 400
        try:
            decoded = QRService.decode_qr(payload)
        except Exception:
            # wrong key / tampered / not a token — outcome only, never payload
            logger.info("web scan: decode rejected (expired key, tampered, or foreign token)")
            return jsonify(ok=False, error="invalid"), 400
        parts = [p.strip() for p in decoded.split("|")]
        if len(parts) != 5 or not parts[2]:
            logger.info("web scan: decrypted payload is not a member badge shape")
            return jsonify(ok=False, error="invalid"), 400
        name, grade, number, _reg, _exp = parts
        rows = member_repo.fetchall(
            "SELECT id, name, grade, number, register_date, expire_date "
            "FROM members WHERE number=?",
            (number,))
        if not rows:
            logger.info("web scan: no member with that badge number")
            return jsonify(ok=False, error="not_found"), 404
        exact = [r for r in rows if r[1] == name and r[2] == grade]
        row = (exact or rows)[0]
        today = datetime.now().strftime("%Y-%m-%d")
        member = {
            "id": row[0], "name": row[1], "grade": row[2], "number": row[3],
            "expire_date": row[5],
            "expired": bool(row[5] and row[5] < today),
        }
        return jsonify(ok=True, member=member)

    # ── Settings ──────────────────────────────────────────
    @app.route("/settings", methods=["GET", "POST"])
    @login_required
    def settings_page():
        if request.method == "POST":
            settings.set("school_name", _clean(request.form.get("school_name"),
                                               LIMITS["school_name"]) or "DSNPRU")
            settings.set("default_loan_days", _int_in(request.form.get("default_loan_days"), 1, 90, 7))
            settings.set("max_books_per_member", _int_in(request.form.get("max_books_per_member"), 1, 20, 3))
            settings.set("member_expiry_days", _int_in(request.form.get("member_expiry_days"), 1, 3650, 365))
            settings.set("school_logo_path", _clean(request.form.get("school_logo_path"),
                                                    LIMITS["logo_path"]) or "assets/logos/school_logo.png")
            settings.save()
            flash("บันทึกการตั้งค่าสำเร็จ", "success")
            return redirect(url_for("settings_page"))
        return render_template("settings.html", settings=settings.all())

    # ── API endpoints ─────────────────────────────────────
    @app.route("/api/stats")
    @login_required
    def api_stats():
        members = member_repo.fetchall("SELECT COUNT(*) FROM members")[0][0]
        books = book_repo.fetchall("SELECT COUNT(*) FROM books")[0][0]
        borrowed = borrow_repo.fetchall(
            "SELECT COUNT(*) FROM borrow_log WHERE returned=0")[0][0]
        return jsonify({"members": members, "books": books, "borrowed": borrowed})

    # Dashboard analytics for the Chart.js widgets (index.html chart section).
    # All series are zero-filled so gaps read as zero, not skipped days.
    @app.route("/api/analytics")
    @login_required
    def api_analytics():
        from datetime import timedelta
        today = datetime.now().date()

        def zero_filled_days(number_of_days, counts_by_day):
            days = []
            for i in range(number_of_days):
                d = (today - timedelta(days=number_of_days - 1 - i)).strftime("%Y-%m-%d")
                days.append({"date": d, "count": counts_by_day.get(d, 0)})
            return days

        start14 = (today - timedelta(days=13)).strftime("%Y-%m-%d")
        rows = borrow_repo.fetchall(
            "SELECT borrow_date, COUNT(*) FROM borrow_log "
            "WHERE borrow_date >= ? GROUP BY borrow_date", (start14,))
        borrow_by_day = {r[0]: r[1] for r in rows}

        start7 = (today - timedelta(days=6)).strftime("%Y-%m-%d")
        rows = access_repo.fetchall(
            "SELECT substr(access_time, 1, 10) AS d, COUNT(*) FROM access_log "
            "WHERE substr(access_time, 1, 10) >= ? GROUP BY d", (start7,))
        access_by_day = {r[0]: r[1] for r in rows}

        status_counts = {r[0]: r[1] for r in book_repo.fetchall(
            "SELECT status, COUNT(*) FROM books GROUP BY status")}
        known = {BOOK_STATUS_AVAILABLE, BOOK_STATUS_BORROWED, BOOK_STATUS_RETIRED}
        status_breakdown = {
            "available": status_counts.get(BOOK_STATUS_AVAILABLE, 0),
            "borrowed": status_counts.get(BOOK_STATUS_BORROWED, 0),
            "retired": status_counts.get(BOOK_STATUS_RETIRED, 0),
            "other": sum(v for k, v in status_counts.items() if k not in known),
        }
        return jsonify({
            "borrow_last_14_days":
                zero_filled_days(14, borrow_by_day),
            "access_last_7_days":
                zero_filled_days(7, access_by_day),
            "book_status": status_breakdown,
        })

    return app
