# web/app.py - Flask web UI for THE PY-LIB
import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from flask import Flask, render_template, request, redirect, url_for, jsonify, session, flash
from config import settings
from config.constants import DB_PATH
from config.env_store import ensure_secrets
from db.schema import init_db
from db.repositories.member_repo import MemberRepository
from db.repositories.book_repo import BookRepository
from db.repositories.borrow_repo import BorrowRepository, BorrowError
from db.repositories.access_repo import AccessRepository
from services.qr_service import QRService
from services.pdf_service import PDFService
from core.security import verify_password
from core.logger import logger

from web import csrf as csrf_mod
from web import throttle
from web.csrf import csrf_field_markup

from functools import wraps
from markupsafe import Markup
import bcrypt
import sqlite3


def _rows_to_dicts(rows, columns):
    return [dict(zip(columns, row, strict=False)) for row in rows]


# Timing-equalization hash for unknown usernames (avoid bcrypt compares of
# existing users vs instant rejection of missing users).
_DUMMY_HASH = bcrypt.hashpw(b"pylib-timing-equalizer", bcrypt.gensalt())

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
        MAX_CONTENT_LENGTH=1024 * 1024,
    )

    @app.route("/assets/<path:filename>")
    def serve_assets(filename):
        from flask import send_from_directory
        return send_from_directory(os.path.join(app.root_path, ".."), "assets/" + filename)

    member_repo = MemberRepository()
    book_repo = BookRepository()
    borrow_repo = BorrowRepository()
    access_repo = AccessRepository()

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
        if not csrf_mod.validate(session, request.form):
            logger.warning("csrf: rejected POST (missing/invalid token) %s", request.path)
            flash("เซสชันหมดอายุหรือโทเคนไม่ถูกต้อง กรุณาลองอีกครั้ง", "error")
            return redirect(request.referrer or url_for("dashboard"))
        return None

    # ── Auth ──────────────────────────────────────────────
    @app.route("/login", methods=["GET", "POST"])
    def login():
        if request.method == "POST":
            username = _clean(request.form.get("username"), 40)
            password = request.form.get("password", "")
            ip = request.remote_addr or "-local-"
            key = (ip, username)

            if throttle.is_locked(key):
                remaining = throttle.lock_remaining_seconds(key)
                minutes = max(1, (remaining + 59) // 60)
                logger.warning("web login: throttled attempt (locked %ss remaining)", remaining)
                # one generic message for throttled AND wrong creds (no user enumeration)
                flash(f"ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง (ลองใหม่ได้ในอีก ~{minutes} นาที)", "error")
                return render_template("login.html", config=settings.all()), 429

            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            user = conn.execute(
                "SELECT * FROM admin_users WHERE username=?", (username,)).fetchone()
            conn.close()

            hashed = (user["password"] if user else _DUMMY_HASH)
            ok = verify_password(password, hashed)
            if ok:
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
        rows = member_repo.fetchall(
            "SELECT id, name, grade, number, register_date, expire_date, qrcode_path "
            "FROM members ORDER BY id DESC")
        members_ = _rows_to_dicts(rows, ["id", "name", "grade", "number",
                                         "register_date", "expire_date", "qrcode_path"])
        return render_template("members.html", members=members_)

    @app.route("/members/add", methods=["POST"])
    @login_required
    def member_add():
        name = _clean(request.form.get("name"), LIMITS["member_name"])
        grade = _clean(request.form.get("grade"), LIMITS["member_grade"])
        number = _clean(request.form.get("number"), LIMITS["member_number"])
        if not all([name, grade, number]):
            flash("กรุณากรอกข้อมูลให้ครบ", "error")
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
        row = member_repo.fetchone("SELECT qrcode_path FROM members WHERE id=?", (mid,))
        if row and row[0]:
            try:
                os.remove(row[0])
            except OSError:
                pass
        member_repo.delete(mid)
        logger.info("Web: member deleted: id=%s", mid)
        flash("ลบสมาชิกสำเร็จ", "success")
        return redirect(url_for("members"))

    # ── Books ─────────────────────────────────────────────
    @app.route("/books")
    @login_required
    def books():
        rows = book_repo.fetchall(
            "SELECT id, code, title, status FROM books ORDER BY id DESC")
        books_ = _rows_to_dicts(rows, ["id", "code", "title", "status"])
        return render_template("books.html", books=books_)

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
        book_repo.add(code, title, status="ว่าง")
        logger.info("Web: book added: code=%s", code)
        flash("เพิ่มหนังสือสำเร็จ", "success")
        return redirect(url_for("books"))

    @app.route("/books/<int:bid>/delete", methods=["POST"])
    @login_required
    def book_delete(bid):
        active = borrow_repo.fetchone(
            "SELECT COUNT(*) FROM borrow_log WHERE book_id=? AND returned=0", (bid,))[0]
        if active > 0:
            flash("ไม่สามารถลบได้ — มีการยืมหนังสือเล่มนี้อยู่", "error")
            return redirect(url_for("books"))
        book_repo.delete(bid)
        logger.info("Web: book deleted: id=%s", bid)
        flash("ลบหนังสือสำเร็จ", "success")
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
            "SELECT id, code, title FROM books WHERE status='ว่าง' ORDER BY title")
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
        rows = borrow_repo.fetchall(
            """SELECT m.name, m.grade, b.code, b.title,
                      bl.borrow_date, bl.return_due, bl.returned
               FROM borrow_log bl
               JOIN members m ON bl.member_id=m.id
               JOIN books b ON bl.book_id=b.id
               ORDER BY bl.id DESC""")
        records = _rows_to_dicts(rows, ["name", "grade", "code", "title",
                                        "borrow_date", "return_due", "returned"])
        return render_template("history.html", records=records)

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

    return app
