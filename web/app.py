# web/app.py - Flask web UI for THE PY-LIB
import os
import sys
import uuid
import secrets
from datetime import datetime, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from flask import Flask, render_template, request, redirect, url_for, jsonify, session, flash
from config import settings
from config.constants import DB_PATH
from db.repositories.member_repo import MemberRepository
from db.repositories.book_repo import BookRepository
from db.repositories.borrow_repo import BorrowRepository
from db.repositories.access_repo import AccessRepository
from services.qr_service import QRService
from services.pdf_service import PDFService
from core.security import verify_password
from core.logger import logger

import sqlite3


def create_app():
    app = Flask(__name__, static_folder="static", template_folder="templates")
    app.secret_key = secrets.token_hex(32)

    member_repo = MemberRepository()
    book_repo = BookRepository()
    borrow_repo = BorrowRepository()
    access_repo = AccessRepository()

    def login_required(f):
        from functools import wraps
        @wraps(f)
        def decorated(*args, **kwargs):
            if not session.get("logged_in"):
                return redirect(url_for("login"))
            return f(*args, **kwargs)
        return decorated

    # ── Auth ──────────────────────────────────────────────
    @app.route("/login", methods=["GET", "POST"])
    def login():
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            user = conn.execute("SELECT * FROM admin_users WHERE username=?", (username,)).fetchone()
            conn.close()
            if user and verify_password(password, user["password"]):
                session["logged_in"] = True
                session["username"] = username
                logger.info(f"Web login: {username}")
                return redirect(url_for("dashboard"))
            flash("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง", "error")
        return render_template("login.html")

    @app.route("/logout")
    def logout():
        session.clear()
        return redirect(url_for("login"))

    # ── Dashboard ─────────────────────────────────────────
    @app.route("/")
    @login_required
    def dashboard():
        members = member_repo.fetchall("SELECT COUNT(*) as c FROM members")[0][0]
        books = book_repo.fetchall("SELECT COUNT(*) as c FROM books")[0][0]
        borrowed = borrow_repo.fetchall(
            "SELECT COUNT(*) as c FROM borrow_log WHERE returned=0")[0][0]
        overdue = borrow_repo.fetchall(
            "SELECT COUNT(*) as c FROM borrow_log WHERE returned=0 AND return_due < date('now')"
        )[0][0]
        return render_template("index.html",
                               members=members, books=books,
                               borrowed=borrowed, overdue=overdue)

    # ── Members ───────────────────────────────────────────
    @app.route("/members")
    @login_required
    def members():
        rows = member_repo.fetchall(
            "SELECT id, name, grade, number, reg_date, exp_date, qr_path "
            "FROM members ORDER BY id DESC")
        return render_template("members.html", members=rows)

    @app.route("/members/add", methods=["POST"])
    @login_required
    def member_add():
        name = request.form.get("name", "").strip()
        grade = request.form.get("grade", "").strip()
        number = request.form.get("number", "").strip()
        if not all([name, grade, number]):
            flash("กรุณากรอกข้อมูลให้ครบ", "error")
            return redirect(url_for("members"))
        reg_date = datetime.now().strftime("%Y-%m-%d")
        exp_date = (datetime.now() + timedelta(days=settings.get("member_expiry_days"))).strftime("%Y-%m-%d")
        qr_data = f"{name}|{grade}|{number}|{reg_date}|{exp_date}"
        qr_filename = f"assets/qrcodes/{uuid.uuid4()}.png"
        os.makedirs("assets/qrcodes", exist_ok=True)
        os.makedirs("assets/cards", exist_ok=True)
        encrypted = QRService.generate_qr(qr_data, qr_filename)
        card_pdf = f"assets/cards/{uuid.uuid4()}.pdf"
        PDFService.generate_card(name, grade, number, reg_date, exp_date, encrypted, card_pdf)
        member_repo.add(name, grade, number, reg_date, exp_date, "qr", qr_filename)
        logger.info(f"Web: member added: {name}")
        flash("เพิ่มสมาชิกสำเร็จ", "success")
        return redirect(url_for("members"))

    @app.route("/members/<int:mid>/delete", methods=["POST"])
    @login_required
    def member_delete(mid):
        member_repo.delete(mid)
        logger.info(f"Web: member deleted: id={mid}")
        flash("ลบสมาชิกสำเร็จ", "success")
        return redirect(url_for("members"))

    # ── Books ─────────────────────────────────────────────
    @app.route("/books")
    @login_required
    def books():
        rows = book_repo.fetchall(
            "SELECT id, code, title, status FROM books ORDER BY id DESC")
        return render_template("books.html", books=rows)

    @app.route("/books/add", methods=["POST"])
    @login_required
    def book_add():
        code = request.form.get("code", "").strip()
        title = request.form.get("title", "").strip()
        if not code or not title:
            flash("กรุณากรอกข้อมูลให้ครบ", "error")
            return redirect(url_for("books"))
        book_repo.add(code, title, status="ว่าง")
        logger.info(f"Web: book added: {code} {title}")
        flash("เพิ่มหนังสือสำเร็จ", "success")
        return redirect(url_for("books"))

    @app.route("/books/<int:bid>/delete", methods=["POST"])
    @login_required
    def book_delete(bid):
        book_repo.delete(bid)
        logger.info(f"Web: book deleted: id={bid}")
        flash("ลบหนังสือสำเร็จ", "success")
        return redirect(url_for("books"))

    # ── Borrow ────────────────────────────────────────────
    @app.route("/borrow", methods=["GET", "POST"])
    @login_required
    def borrow():
        if request.method == "POST":
            member_id = request.form.get("member_id")
            book_id = request.form.get("book_id")
            due = request.form.get("due_date", "")
            if not member_id or not book_id:
                flash("กรุณาเลือกสมาชิกและหนังสือ", "error")
                return redirect(url_for("borrow"))
            borrow_date = datetime.now().strftime("%Y-%m-%d")
            borrow_repo.add(int(member_id), int(book_id), borrow_date, due)
            book_repo.update_status(int(book_id), "ยืมแล้ว")
            logger.info(f"Web: borrowed book {book_id} to member {member_id}")
            flash("ยืมหนังสือสำเร็จ", "success")
            return redirect(url_for("borrow"))

        members = member_repo.fetchall(
            "SELECT id, name, grade, number FROM members ORDER BY name")
        books = book_repo.fetchall(
            "SELECT id, code, title FROM books WHERE status='ว่าง' ORDER BY title")
        default_due = (datetime.now() + timedelta(days=settings.get("default_loan_days"))).strftime("%Y-%m-%d")
        return render_template("borrow.html", members=members, books=books,
                               default_due=default_due)

    # ── Return ────────────────────────────────────────────
    @app.route("/return", methods=["GET", "POST"])
    @login_required
    def return_book():
        if request.method == "POST":
            record_id = request.form.get("record_id")
            book_id = request.form.get("book_id")
            if record_id and book_id:
                borrow_repo.mark_returned(int(record_id))
                book_repo.update_status(int(book_id), "ว่าง")
                logger.info(f"Web: returned book_id={book_id} record_id={record_id}")
                flash("คืนหนังสือสำเร็จ", "success")
            return redirect(url_for("return_book"))

        borrowed = borrow_repo.fetchall(
            """SELECT bl.id, m.name, m.grade, m.number, b.code, b.title,
                      bl.borrow_date, bl.return_due
               FROM borrow_log bl
               JOIN members m ON bl.member_id=m.id
               JOIN books b ON bl.book_id=b.id
               WHERE bl.returned=0
               ORDER BY bl.borrow_date DESC""")
        return render_template("return_book.html", borrowed=borrowed)

    # ── History ───────────────────────────────────────────
    @app.route("/history")
    @login_required
    def history():
        rows = borrow_repo.fetchall(
            """SELECT m.name, m.grade, b.code, b.title,
                      bl.borrow_date, bl.return_due, bl.return_date, bl.returned
               FROM borrow_log bl
               JOIN members m ON bl.member_id=m.id
               JOIN books b ON bl.book_id=b.id
               ORDER BY bl.id DESC""")
        return render_template("history.html", records=rows)

    # ── Access History ────────────────────────────────────
    @app.route("/access")
    @login_required
    def access():
        rows = access_repo.get_recent_access(limit=100)
        return render_template("access.html", records=rows)

    # ── Settings ──────────────────────────────────────────
    @app.route("/settings", methods=["GET", "POST"])
    @login_required
    def settings_page():
        if request.method == "POST":
            settings.set("school_name", request.form.get("school_name", "").strip())
            settings.set("default_loan_days", int(request.form.get("default_loan_days", 7)))
            settings.set("max_books_per_member", int(request.form.get("max_books_per_member", 3)))
            settings.set("member_expiry_days", int(request.form.get("member_expiry_days", 365)))
            settings.save()
            flash("บันทึกการตั้งค่าสำเร็จ", "success")
            return redirect(url_for("settings_page"))
        return render_template("settings.html", settings=settings.all())

    # ── API endpoints (for AJAX/SPA features) ────────────
    @app.route("/api/stats")
    @login_required
    def api_stats():
        members = member_repo.fetchall("SELECT COUNT(*) FROM members")[0][0]
        books = book_repo.fetchall("SELECT COUNT(*) FROM books")[0][0]
        borrowed = borrow_repo.fetchall(
            "SELECT COUNT(*) FROM borrow_log WHERE returned=0")[0][0]
        return jsonify({"members": members, "books": books, "borrowed": borrowed})

    return app
