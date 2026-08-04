# app.py - LibraryApp shell + navigation controller
import sqlite3

import customtkinter as ctk

import db.schema
import ui.theme as theme
from config.constants import DB_PATH
from core.logger import logger

from ui.views.login_view import LoginView
from ui.views.dashboard_view import DashboardView
from ui.views.member_view import MemberView
from ui.views.book_view import BookView
from ui.views.borrow_view import BorrowView
from ui.views.return_view import ReturnView
from ui.views.history_view import HistoryView
from ui.views.access_view import AccessScannerView, AccessHistoryView
from ui.views.settings_view import SettingsView
from ui.views.webui_view import WebUIView


class LibraryApp(ctk.CTk):
    # (key, sidebar label, page title, view class)
    PAGES = [
        ("home", "หน้าหลัก", "หน้าหลัก", DashboardView),
        ("members", "จัดการสมาชิก", "จัดการสมาชิก", MemberView),
        ("books", "จัดการหนังสือ", "จัดการหนังสือ", BookView),
        ("borrow", "ยืมหนังสือ", "ยืมหนังสือ", BorrowView),
        ("return", "คืนหนังสือ", "คืนหนังสือ", ReturnView),
        ("history", "ประวัติยืม-คืน", "ประวัติการยืม-คืน", HistoryView),
        ("access", "บันทึกเข้าออก", "บันทึกการเข้าออก", AccessScannerView),
        ("access_history", "ประวัติเข้าออก", "ประวัติการเข้าออก", AccessHistoryView),
        ("settings", "ตั้งค่า", "ตั้งค่า", SettingsView),
        ("webui", "Web UI", "Web UI", WebUIView),
    ]
    SIDEBAR_GROUPS = [
        ("", ["home"]),
        ("จัดการ", ["members", "books"]),
        ("ยืม-คืน", ["borrow", "return", "history"]),
        ("ห้องสมุด", ["access", "access_history"]),
        ("เพิ่มเติม", ["webui"]),
        ("ระบบ", ["settings"]),
    ]

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        theme.register_fonts()

        self.title("THE PY-LIB — ระบบห้องสมุดอัจฉริยะ")
        self.geometry("1180x740")
        self.minsize(1024, 640)

        db.schema.init_db()

        self.login_conn = sqlite3.connect(DB_PATH)
        self.nav_frame = None
        self.header_frame = None
        self.content_frame = None
        self.current_view = None
        self.active_key = None
        self.nav_buttons = {}

        self._show_login()

    # ================= login =================
    def _show_login(self):
        self.content_frame = ctk.CTkFrame(self, fg_color=theme.content_bg())
        self.content_frame.pack(fill="both", expand=True)
        self.current_view = LoginView(self.content_frame, self.login_conn, self._on_login_success)

    def _on_login_success(self):
        logger.info("Login successful, building main UI")
        if self.login_conn is not None:
            try:
                self.login_conn.close()
            except Exception as exc:
                logger.warning(f"Error closing login connection: {exc}")
            self.login_conn = None
        if self.content_frame is not None:
            self.content_frame.destroy()
        self._build_main_ui()

    def _logout(self):
        logger.info("Logging out")
        for frame in (self.nav_frame, self.header_frame, self.content_frame):
            if frame is not None:
                try:
                    frame.destroy()
                except Exception:
                    pass
        self.nav_frame = self.header_frame = self.content_frame = None
        self.current_view = None
        self.nav_buttons = {}
        self.active_key = None
        self.login_conn = sqlite3.connect(DB_PATH)
        self._show_login()

    # ================= main shell =================
    def _build_main_ui(self):
        self.nav_frame = ctk.CTkFrame(self, width=232, corner_radius=0,
                                      fg_color=theme.SIDEBAR)
        self.nav_frame.pack(side="left", fill="y")
        self.nav_frame.pack_propagate(False)

        # Brand
        brand = ctk.CTkFrame(self.nav_frame, fg_color="transparent")
        brand.pack(fill="x", padx=20, pady=(24, 8))
        ctk.CTkLabel(brand, text="THE PY-LIB", font=theme.font(22, bold=True),
                     text_color="#FFFFFF", anchor="w").pack(fill="x")
        ctk.CTkLabel(brand, text="ระบบห้องสมุดอัจฉริยะ", font=theme.font(12),
                     text_color=theme.SIDEBAR_MUTED, anchor="w").pack(fill="x")

        sep = ctk.CTkFrame(self.nav_frame, height=1, fg_color=theme.SIDEBAR_HOVER)
        sep.pack(fill="x", padx=20, pady=(12, 8))

        # Groups
        for group_label, keys in self.SIDEBAR_GROUPS:
            if group_label:
                ctk.CTkLabel(self.nav_frame, text=group_label, font=theme.font(11, bold=True),
                             text_color=theme.SIDEBAR_MUTED, anchor="w").pack(
                    fill="x", padx=20, pady=(14, 4))
            for key in keys:
                self.nav_buttons[key] = self._nav_button(key)

        # Header
        self.header_frame = ctk.CTkFrame(self, height=60, corner_radius=0,
                                         fg_color=theme.card_bg())
        self.header_frame.pack(side="top", fill="x")
        self.header_frame.pack_propagate(False)

        self.header_title = ctk.CTkLabel(self.header_frame, text="", font=theme.font(19, bold=True),
                                         text_color=theme.text_color(), anchor="w")
        self.header_title.pack(side="left", padx=24)

        ctk.CTkButton(self.header_frame, text="ออกจากระบบ", command=self._logout,
                      font=theme.font(13, bold=True), width=110, height=32,
                      fg_color="transparent", text_color=theme.DANGER,
                      hover_color=theme.DANGER_SOFT, corner_radius=8).pack(
            side="right", padx=20)

        # Content
        self.content_frame = ctk.CTkFrame(self, fg_color=theme.content_bg(),
                                          corner_radius=0)
        self.content_frame.pack(side="top", fill="both", expand=True)

        self.show_dashboard()

    def _nav_button(self, key):
        entry = next(p for p in self.PAGES if p[0] == key)
        btn = ctk.CTkButton(self.nav_frame, text=entry[1], command=lambda: self._go(key),
                            anchor="w", height=38,
                            font=theme.font(14), corner_radius=8,
                            fg_color="transparent", hover_color=theme.SIDEBAR_HOVER,
                            text_color=theme.SIDEBAR_TEXT)
        btn.pack(fill="x", padx=12, pady=2)
        return btn

    def _go(self, key):
        entry = next(p for p in self.PAGES if p[0] == key)
        self._highlight(key)
        self.header_title.configure(text=entry[2])
        self._show_view(entry[3])

    def _highlight(self, key):
        for k, btn in self.nav_buttons.items():
            if k == key:
                btn.configure(fg_color=theme.SIDEBAR_ACTIVE, text_color="#FFFFFF")
            else:
                btn.configure(fg_color="transparent", text_color=theme.SIDEBAR_TEXT)
        self.active_key = key

    def _show_view(self, view_cls):
        if self.current_view is not None:
            try:
                self.current_view.destroy()
            except Exception as exc:
                logger.warning(f"Error destroying view: {exc}")
        self.current_view = view_cls(self.content_frame, self)

    # ================= navigation targets =================
    def show_dashboard(self):
        self._go("home")

    def show_members(self):
        self._go("members")

    def show_books(self):
        self._go("books")

    def show_borrow(self):
        self._go("borrow")

    def show_return(self):
        self._go("return")

    def show_history(self):
        self._go("history")

    def show_access_scanner(self):
        self._go("access")

    def show_access_history(self):
        self._go("access_history")

    def show_settings(self):
        self._go("settings")

    def show_about(self):
        self._go("settings")

    # ================= run =================
    def run(self):
        self.mainloop()
