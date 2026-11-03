# web/app.py - Flask web UI for THE PY-LIB
import io
import os
import sys
import time
from collections import deque
from datetime import datetime, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from flask import (Flask, render_template, request, redirect, url_for,
                   jsonify, session, flash, send_file, abort)
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
from services.qr_service import QRService
from services.pdf_service import PDFService
from core.logger import logger
from core import security
from core import throttle
from services import member_identity

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

# Member portal constants
MEMBER_ACTIVATION_TTL_SECONDS = 10 * 60    # card/code → PIN-choice window
MEMBER_PIN_MIN, MEMBER_PIN_MAX = 6, 64     # 6+ keeps trivial pads out

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

    def _member_portal_enabled():
        """The Settings switch can 404 the whole /member/* page family."""
        return bool(settings.get("member_portal_enabled"))

    def member_required(f):
        """Gate for the member's own status-only portal session.

        Authorization is symmetric with the admin side: a member session
        carries no admin rights (the `logged_in` key it simply lacks), so
        member pages and admin pages are disjoint. Members never see, and
        never POST to, anything that changes library state.
        """
        @wraps(f)
        def decorated(*args, **kwargs):
            if not _member_portal_enabled():
                abort(404)
            if not session.get("member_id"):
                return redirect(url_for("member_login"))
            return f(*args, **kwargs)
        return decorated

    def _member_from_session():
        """The logged-in member's row, or None (row deleted → treat as off)."""
        mid = session.get("member_id")
        if not mid:
            return None
        return member_repo.fetchone(
            "SELECT id, name, grade, number, expire_date "
            "FROM members WHERE id = ?", (mid,))

    def _activation_member_fresh():
        """Active mid-activation member id, or None if the window elapsed.

        The activation window (10 min) bounds the time a half-open
        session exists between card/code verification and PIN choice.
        """
        mid = session.get("activation_member_id")
        at = session.get("activation_at")
        if not mid or not at:
            return None
        if time.time() - float(at) > MEMBER_ACTIVATION_TTL_SECONDS:
            session.pop("activation_member_id", None)
            session.pop("activation_at", None)
            return None
        return mid

    def _drop_activation_session():
        session.pop("activation_member_id", None)
        session.pop("activation_at", None)

    def _member_page_row():
        """Member row for portal PAGE handlers, or None after cleanup.

        None means the member row vanished while logged in (deleted by an
        admin): the dangling session is dropped here, so page handlers can
        simply redirect to the login. /member/qr.png keeps its own JSON
        guard instead because it answers the live countdown script.
        """
        row = _member_from_session()
        if not row:
            session.pop("member_id", None)
            # same generic message as a failed login — nothing echoes back
            flash("ไม่พบสมาชิกนี้ในระบบ กรุณาเข้าสู่ระบบอีกครั้ง", "error")
        return row

    # ---- /member/qr.png rate limiting (per member session; ids only) ----
    QR_RATE_LIMIT = 30        # images per window, generous for a human page
    QR_RATE_WINDOW = 60.0     # seconds
    _qr_rate = {}             # member_id -> deque of request timestamps

    def _qr_rate_ok(mid):
        now = time.time()
        dq = _qr_rate.get(mid)
        if dq is None:
            dq = _qr_rate[mid] = deque()
        while dq and now - dq[0] > QR_RATE_WINDOW:
            dq.popleft()
        if len(dq) >= QR_RATE_LIMIT:
            return False
        dq.append(now)
        if len(_qr_rate) > 4096:      # bound the table; empties first
            for stale in [k for k, v in _qr_rate.items() if not v][:100]:
                _qr_rate.pop(stale, None)
        return True

    def _qr_png(payload):
        """Render a ticket payload as PNG bytes, no filesystem touch."""
        import qrcode
        img = qrcode.QRCode(
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=8, border=2)
        img.add_data(payload)
        img.make(fit=True)
        out = io.BytesIO()
        img.make_image(fill_color="black", back_color="white").save(
            out, format="PNG")
        out.seek(0)
        return out

    @app.context_processor
    def _inject_csrf():
        try:
            return {
                "csrf_token": csrf_mod._get_token(session),
                "csrf_field": csrf_field_markup(session),
            }
        except RuntimeError:  # outside request context
            return {"csrf_token": "", "csrf_field": Markup("")}

    # member endpoints grouped by their session state at POST time
    _MEMBER_PUBLIC_POSTS = {
        # no member session yet: first login + first activation steps
        "member_login", "member_activate_scan",
        "member_activate_code", "member_activate_setpin",
    }
    _MEMBER_SESSION_POSTS = {"member_logout"}

    @app.before_request
    def _csrf_protect():
        """Reject state-changing POSTs without a valid session CSRF token.

        Zones, checked in order:
        * `login` (admin) — rendered without login_required; token required;
        * member public endpoints — token required, bounce to the member
          page that rendered the form (portal may be off → 404 anyway);
        * member-session endpoints (portal/logout) — a logged-in member's
          POSTs use the same token, but bounce inside the portal;
        * everything else — admin area: requires BOTH an admin session and
          the token, so a member-only session can never pass here.
        """
        if request.method != "POST":
            return None
        endpoint = request.endpoint
        if endpoint == "login":
            # login is rendered without login_required; session token still required
            if not csrf_mod.validate(session, request.form):
                flash("ระดับเซสชันไม่ถูกต้อง กรุณาลองอีกครั้ง", "error")
                return redirect(url_for("login"))
            return None
        if endpoint in _MEMBER_PUBLIC_POSTS:
            if not csrf_mod.validate(session, request.form):
                logger.warning("csrf: rejected member POST %s", request.path)
                flash("หมดเวลาการทำรายการ กรุณาลองอีกครั้ง", "error")
                back = "member_activate" if endpoint != "member_login" \
                    else "member_login"
                return redirect(url_for(back))
            return None
        if endpoint in _MEMBER_SESSION_POSTS and not session.get("logged_in"):
            # member-only session (an admin account posting member_logout is
            # still inside the admin zone below and passes the same check)
            token = (request.form.get("csrf_token")
                     or request.headers.get("X-CSRF-Token", ""))
            if not csrf_mod.validate(session, token):
                logger.warning("csrf: rejected member POST %s", request.path)
                flash("หมดเวลาการทำรายการ กรุณาลองอีกครั้ง", "error")
                return redirect(url_for("member_portal"))
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
        total = repo.fetchone(f"SELECT COUNT(*) FROM {base}")[0]  # noqa: S608 - base is a fixed table-name literal
        where_sql = f" WHERE {where}" if where else ""
        filtered = repo.fetchone(
            f"SELECT COUNT(*) FROM {base}{where_sql}",  # noqa: S608 - base/where are caller-fixed literals; user input is always bound
            tuple(params))[0]
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

    # ── Member portal (rotating QR identity, status-only) ──
    #
    # Members log in with (ชั้น + เลขที่) + their own PIN and see exactly
    # their own data. The portal is one module per page — /member (home:
    # profile + live rotating QR), /member/loans, /member/returns,
    # /member/access — joined by the nav bar from member_base.html. Each
    # page re-checks the session row and nothing here changes library
    # state — borrow/return/delete remain admin-side on web AND desktop.

    @app.route("/member/login", methods=["GET", "POST"])
    def member_login():
        if not _member_portal_enabled():
            abort(404)
        if request.method == "GET" and _member_from_session():
            return redirect(url_for("member_portal"))
        if request.method == "POST":
            grade = _clean(request.form.get("grade"), LIMITS["member_grade"])
            number = _clean(request.form.get("number"), LIMITS["member_number"])
            pin = (request.form.get("pin") or "").strip()
            # throttle key ties the lock to the attempted badge, so a
            # stalker guessing one badge can't freeze others and vice versa
            ip = request.remote_addr or "-local-"
            key = (ip, f"member:{grade}:{number}")
            if throttle.is_locked(key):
                minutes = throttle.lock_remaining_minutes(key)
                logger.warning("web member login: throttled "
                               "(locked %ss remaining)",
                               throttle.lock_remaining_seconds(key))
                flash("รหัส PIN ไม่ถูกต้อง "
                      f"(ลองใหม่ได้ในอีก ~{minutes} นาที)", "error")
                return render_template("member_login.html"), 429
            row = None
            if grade and number and pin and len(pin) < 200:
                row = member_repo.get_by_badge_for_login(grade, number)
            ok = bool(row and row[1]
                      and member_repo.verify_member_pin(row[0], pin))
            if ok:
                throttle.record_success(key)
                session["member_id"] = row[0]
                _drop_activation_session()
                logger.info("web member login: successful")  # no badge/PII in logs
                return redirect(url_for("member_portal"))
            throttle.record_failure(key)
            logger.warning("web member login: failed attempt")
            flash("ชั้น เลขที่ หรือรหัส PIN ไม่ถูกต้อง", "error")
        return render_template("member_login.html")

    @app.route("/member/logout", methods=["POST"])
    @member_required
    def member_logout():
        session.pop("member_id", None)
        _drop_activation_session()
        return redirect(url_for("member_login"))

    @app.route("/member")
    @member_required
    def member_portal():
        """Home page: profile + live rotating QR + summary module links."""
        row = _member_page_row()
        if not row:
            return redirect(url_for("member_login"))
        mid, name, grade, number, expire_date = row
        today = datetime.now().strftime("%Y-%m-%d")
        loan_count = member_repo.fetchone(
            "SELECT COUNT(*) FROM borrow_log "
            "WHERE member_id = ? AND returned = 0", (mid,))[0]
        overdue_count = member_repo.fetchone(
            "SELECT COUNT(*) FROM borrow_log "
            "WHERE member_id = ? AND returned = 0 AND return_due < ?",
            (mid, today))[0]
        return_count = member_repo.fetchone(
            "SELECT COUNT(*) FROM borrow_log "
            "WHERE member_id = ? AND returned = 1", (mid,))[0]
        access_count = member_repo.fetchone(
            "SELECT COUNT(*) FROM access_log WHERE member_id = ?", (mid,))[0]
        hours = member_identity.rotation_hours()
        grace = member_identity.grace_steps()
        seconds_left = security.ticket_epoch_seconds_left(hours)
        expired = bool(expire_date and expire_date < today)
        return render_template(
            "member_home.html",
            member_name=name, member_grade=grade, member_number=number,
            expire_date=expire_date, member_expired=expired,
            loan_count=loan_count, overdue_count=overdue_count,
            return_count=return_count, access_count=access_count,
            qr_seconds=seconds_left, qr_rotation_hours=hours, qr_grace=grace,
        )

    @app.route("/member/loans")
    @member_required
    def member_loans():
        """Books the member currently borrows — its own module page."""
        row = _member_page_row()
        if not row:
            return redirect(url_for("member_login"))
        today = datetime.now().strftime("%Y-%m-%d")
        loans = member_repo.fetchall(
            "SELECT bo.title, bo.code, bl.borrow_date, bl.return_due "
            "FROM borrow_log bl JOIN books bo ON bo.id = bl.book_id "
            "WHERE bl.member_id = ? AND bl.returned = 0 "
            "ORDER BY bl.return_due ASC", (row[0],))
        return render_template("member_loans.html", loans=loans, today=today)

    @app.route("/member/returns")
    @member_required
    def member_returns():
        """Recently returned books, newest first — its own module page.

        As a dedicated page it can afford a longer window than the old
        shared portal card, so the cap went 10 → 50 rows.
        """
        row = _member_page_row()
        if not row:
            return redirect(url_for("member_login"))
        recent_returns = member_repo.fetchall(
            "SELECT bo.title, bo.code, bl.borrow_date, bl.return_due "
            "FROM borrow_log bl JOIN books bo ON bo.id = bl.book_id "
            "WHERE bl.member_id = ? AND bl.returned = 1 "
            "ORDER BY bl.id DESC LIMIT 50", (row[0],))
        return render_template("member_returns.html",
                               recent_returns=recent_returns)

    @app.route("/member/access")
    @member_required
    def member_access():
        """Own door-scan history, newest first — its own module page.

        Limit raised 10 → 50 for the same reason as /member/returns.
        """
        row = _member_page_row()
        if not row:
            return redirect(url_for("member_login"))
        access_history = member_repo.fetchall(
            "SELECT access_time, action FROM access_log "
            "WHERE member_id = ? ORDER BY id DESC LIMIT 50", (row[0],))
        return render_template("member_access.html",
                               access_history=access_history)

    @app.route("/member/qr.png")
    @member_required
    def member_qr():
        row = _member_from_session()
        if not row:
            session.pop("member_id", None)
            return jsonify(ok=False, error="session"), 401
        if not _qr_rate_ok(row[0]):
            logger.info("web member qr: rate limited (id kept out of logs)")
            return jsonify(ok=False, error="too_fast"), 429
        hours = member_identity.rotation_hours()
        # ?next=1 renders the NEXT window: the portal shows it as the
        # reserve QR so a desk scan never dead-ends mid-flip
        epoch = security.ticket_epoch(hours) + (
            1 if request.args.get("next") == "1" else 0)
        payload = security.encrypt_member_ticket(row[0], hours, epoch=epoch)
        resp = send_file(_qr_png(payload), mimetype="image/png",
                         download_name="member-qr.png", conditional=False)
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
        resp.headers["Pragma"] = "no-cache"
        return resp

    # ---- first-time activation (card QR at the desk or admin code) ----

    @app.route("/member/activate")
    def member_activate():
        if not _member_portal_enabled():
            abort(404)
        if _member_from_session():
            return redirect(url_for("member_portal"))
        mid = _activation_member_fresh()
        member_label = None
        if mid:
            row = member_repo.fetchone(
                "SELECT name, grade, number FROM members WHERE id = ?",
                (mid,))
            member_label = f"{row[0]} · {row[1]} · {row[2]}" if row else None
        return render_template("member_activate.html",
                               activation_member=mid and member_label)

    @app.route("/member/activate/scan", methods=["POST"])
    def member_activate_scan():
        if not _member_portal_enabled():
            abort(404)
        if _member_from_session():
            return redirect(url_for("member_portal"))
        payload = _clean(request.form.get("payload"), 800)
        try:
            ident = member_identity.resolve_scanned(payload)
        except member_identity.MemberIdentityError as exc:
            # the Thai message is the outcome — no payload ever echoes back
            logger.info("web activation: scan rejected")
            flash(str(exc), "error")
            return redirect(url_for("member_activate"))
        except Exception:
            logger.info("web activation: scan decode error")
            flash("QR ไม่ถูกต้อง กรุณาลองใหม่", "error")
            return redirect(url_for("member_activate"))
        if ident.get("kind") != "static":
            # a v3 ticket here proves nothing useful — whoever has it
            # already has a working login; first activation needs the
            # physical card or an admin-issued code
            logger.info("web activation: non-card token rejected")
            flash("เปิดใช้งานครั้งแรกต้องใช้ QR จากบัตรประจำตัว หรือรหัสที่ผู้ดูแลออกให้",
                  "error")
            return redirect(url_for("member_activate"))
        mid = ident["member_id"]
        if member_repo.pin_state(mid) == "set":
            # takeover guard: an activated member cannot be re-activated,
            # not even by presenting the original card
            logger.info("web activation: refused, member already has pin id=%s", mid)
            flash("สมาชิกนี้เปิดใช้งานหน้าของตนเองอยู่แล้ว — "
                  "กรุณาเข้าสู่ระบบด้วยรหัส PIN", "error")
            return redirect(url_for("member_login"))
        session["activation_member_id"] = mid
        session["activation_at"] = time.time()
        logger.info("web activation: card verified id=%s", mid)
        return redirect(url_for("member_activate"))

    @app.route("/member/activate/code", methods=["POST"])
    def member_activate_code():
        if not _member_portal_enabled():
            abort(404)
        if _member_from_session():
            return redirect(url_for("member_portal"))
        code = (request.form.get("code") or "").strip()[:80]
        mid = member_repo.find_activation_member(code) if code else None
        if not mid:
            logger.info("web activation: code rejected")
            flash("รหัสเปิดใช้งานไม่ถูกต้องหรือหมดอายุ กรุณาติดต่อผู้ดูแล",
                  "error")
            return redirect(url_for("member_activate"))
        if member_repo.pin_state(mid) == "set":
            logger.info("web activation: refused, already active id=%s", mid)
            flash("สมาชิกนี้เปิดใช้งานหน้าของตนเองอยู่แล้ว — "
                  "กรุณาเข้าสู่ระบบด้วยรหัส PIN", "error")
            return redirect(url_for("member_login"))
        # claim the code NOW: the 12-char plaintext never rides in the
        # session cookie, and double-entry burns it for everyone (fine —
        # the activation guard still protects the account itself)
        member_repo.claim_activation_code(code, mid)
        session["activation_member_id"] = mid
        session["activation_at"] = time.time()
        logger.info("web activation: code accepted id=%s", mid)
        return redirect(url_for("member_activate"))

    @app.route("/member/activate/set-pin", methods=["POST"])
    def member_activate_setpin():
        if not _member_portal_enabled():
            abort(404)
        if _member_from_session():
            return redirect(url_for("member_portal"))
        mid = _activation_member_fresh()
        if not mid:
            flash("หมดเวลาเปิดใช้งาน กรุณาเริ่มขั้นตอนใหม่", "error")
            return redirect(url_for("member_activate"))
        pin1 = request.form.get("pin1") or ""
        pin2 = request.form.get("pin2") or ""
        if pin1 != pin2:
            flash("รหัส PIN สองช่องไม่ตรงกัน", "error")
            return redirect(url_for("member_activate"))
        pin = pin1.strip()
        if (not MEMBER_PIN_MIN <= len(pin) <= MEMBER_PIN_MAX
                or any(ord(c) < 32 for c in pin)):
            flash(f"รหัส PIN ต้องมีความยาว {MEMBER_PIN_MIN}–"
                  f"{MEMBER_PIN_MAX} ตัวอักษร", "error")
            return redirect(url_for("member_activate"))
        if member_repo.pin_state(mid) == "set":
            logger.info("web activation: refused, already active id=%s", mid)
            flash("สมาชิกนี้เปิดใช้งานหน้าของตนเองอยู่แล้ว — "
                  "กรุณาเข้าสู่ระบบด้วยรหัส PIN", "error")
            return redirect(url_for("member_login"))
        if not member_repo.activate_with_pin(mid, pin):
            _drop_activation_session()
            logger.info("web activation: atomic activation refused id=%s", mid)
            flash("เปิดใช้งานไม่สำเร็จ กรุณาเริ่มขั้นตอนใหม่", "error")
            return redirect(url_for("member_activate"))
        session["member_id"] = mid
        _drop_activation_session()
        logger.info("web activation: pin set id=%s", mid)
        flash("เปิดใช้งานสำเร็จ ยินดีต้อนรับเข้าสู่หน้าของคุณ", "success")
        return redirect(url_for("member_portal"))

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
        """DataTables server-side feed for the members table.

        The extra `pin_set` column tells the UI whether the member has
        activated their portal login (has a stored bcrypt PIN) — it adjusts
        the row actions (issue activation code vs clear PIN). The hash
        itself is never selected.
        """
        dt = _dt_args()
        base = "members"
        where, params = "", ()
        if dt["q"]:
            like = f"%{dt['q']}%"
            where = "name LIKE ? OR grade LIKE ? OR number LIKE ?"
            params = (like, like, like)
        pin_expr = "(CASE WHEN pin_hash IS NULL THEN 0 ELSE 1 END)"
        return _dt_response(
            dt, member_repo, base, where, params,
            "SELECT id, name, grade, number, register_date, expire_date, "  # noqa: S608 - pin_expr is a fixed CASE literal from the allowlisted columns below
            f"qrcode_path, {pin_expr} FROM members",
            [("id", "id"), ("name", "name"), ("grade", "grade"),
             ("number", "number"), ("register_date", "register_date"),
             ("expire_date", "expire_date"), ("qrcode_path", "qrcode_path"),
             (pin_expr, "pin_set")],
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

    @app.route("/members/<int:mid>/activation-code", methods=["POST"])
    @login_required
    def member_issue_code(mid):
        """Mint a one-time portal activation code for THIS member.

        Only offered while the member has no portal PIN yet — an issued
        code always means 'set a NEW pin', so it must not exist for an
        account that is already active. The plaintext is returned exactly
        once: rendered in the page flash for a classic form POST, or as
        {ok, code} JSON for the members-table row button (which must
        surface it in a dialog, then it exists nowhere else). Only the
        SHA-256 hash reaches the database, and it is never logged.
        """
        issued = None
        if member_repo.pin_state(mid) != "set":
            issued = member_repo.issue_activation_code(mid)
            logger.info("Web: member activation code issued: id=%s", mid)
        # the table engine POSTs header-only CSRF; a browser form posts the
        # token as a field — use that to pick the response shape
        as_json = bool(request.headers.get("X-CSRF-Token")) \
            and not request.form.get("csrf_token")
        if as_json:
            if issued is None:
                return jsonify(ok=False, message=
                               "สมาชิกนี้เปิดใช้งานหน้าของตนเองอยู่แล้ว "
                               "— หากต้องการรีเซ็ต ให้ล้างรหัสก่อน")
            return jsonify(ok=True, code=issued)
        if issued is None:
            flash("สมาชิกเปิดใช้งานหน้าของตนเองอยู่แล้ว — "
                  "หากต้องการรีเซ็ต ให้กดปุ่ม ล้างรหัส ก่อน", "error")
            return redirect(url_for("members"))
        flash(f"รหัสเปิดใช้งานสำหรับสมาชิกรหัสที่ {mid}: {issued} "
              "(ใช้ได้ครั้งเดียว ภายใน 15 นาที — สำเนาของรหัสจะไม่แสดงอีก)",
              "success")
        return redirect(url_for("members"))

    @app.route("/members/<int:mid>/reset-pin", methods=["POST"])
    @login_required
    def member_reset_pin(mid):
        """Clear a member's portal PIN (forces re-activation).

        Outstanding activation codes die with the reset (repo handles
        this), so a leaked code can't resurrect the account silently.
        """
        member_repo.clear_pin(mid)
        logger.info("Web: member portal pin cleared: id=%s", mid)
        flash("ล้างรหัสผ่านหน้า (portal) สมาชิกเรียบร้อย — "
              "สมาชิกต้องเปิดใช้งานใหม่ด้วยบัตร หรือรหัสที่ผู้ดูแลสร้างให้",
              "success")
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
    # Every accepted format lives in services.member_identity — the same
    # module the desktop scanner uses, so a QR times out identically on both
    # sides. Each attempt costs one PBKDF2 derivation, and the route is
    # fully wrapped by login_required + CSRF before_request — brute force
    # is impractical.
    @app.route("/scan/resolve", methods=["POST"])
    @login_required
    def scan_resolve():
        payload = _clean(request.form.get("payload"), 800)
        if not payload:
            return jsonify(ok=False, error="invalid"), 400
        try:
            ident = member_identity.resolve_scanned(payload, member_repo)
        except member_identity.MemberIdentityError as exc:
            # classify outcome for the desk UI without echoing any payload
            if "QR หมดเวลา" in str(exc):
                logger.info("web scan: ticket window expired")
                return jsonify(ok=False, error="expired",
                               message=str(exc)), 400
            if "ไม่พบสมาชิก" in str(exc):
                logger.info("web scan: no member matches this scan")
                return jsonify(ok=False, error="not_found",
                               message=str(exc)), 404
            logger.info("web scan: payload rejected "
                        "(expired key, tampered, foreign token, or policy)")
            return jsonify(ok=False, error="invalid",
                           message="QR ไม่ถูกต้อง บัตรเก่า หรือไฟล์โทเคนถูกเปลี่ยน"), 400
        except Exception:
            logger.error("web scan: resolver crashed", exc_info=True)
            return jsonify(ok=False, error="invalid"), 400
        member = {"id": ident["member_id"], "name": ident["name"],
                  "grade": ident["grade"], "number": ident["number"],
                  "expire_date": ident["expire_date"],
                  "expired": ident["expired"]}
        return jsonify(ok=True, member=member)

    # ── Settings ──────────────────────────────────────────
    @app.route("/settings", methods=["GET", "POST"])
    @login_required
    def settings_page():
        if request.method == "POST":
            # the page carries two independent forms (general + member
            # portal); only the group that submitted is written, so
            # checkbox-absent semantics can't wipe the untouched group
            if request.form.get("settings_group") == "member_portal":
                # member portal / rotating QR identity
                settings.set("member_portal_enabled",
                             bool(request.form.get("member_portal_enabled")))
                # clamps match the member_identity read-side (1..72 h, 0..2)
                settings.set("member_qr_rotation_hours",
                             _int_in(request.form.get("member_qr_rotation_hours"),
                                     1, 72, 12))
                settings.set("member_qr_grace_steps",
                             _int_in(request.form.get("member_qr_grace_steps"),
                                     0, 2, 1))
                settings.set("accept_static_card_qr",
                             bool(request.form.get("accept_static_card_qr")))
            else:
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
