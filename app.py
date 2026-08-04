# app.py - LibraryApp shell + navigation controller
import sqlite3
import customtkinter as ctk

import db.schema
from config.constants import COLORS, DB_PATH
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


class LibraryApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme(COLORS.get("primary", "blue"))

        self.title("ระบบห้องสมุดอัจฉริยะ THE PY-LIB")
        self.geometry("1100x700")
        self.minsize(900, 600)

        db.schema.init_db()

        self.login_conn = sqlite3.connect(DB_PATH)

        self.nav_frame = None
        self.content_frame = None
        self.current_view = None

        self._show_login()

    # ---------------- login ----------------
    def _show_login(self):
        self.content_frame = ctk.CTkFrame(self, fg_color=COLORS.get("bg", "#f0f0f0"))
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

    # ---------------- navigation shell ----------------
    def _build_main_ui(self):
        self.nav_frame = ctk.CTkFrame(self, width=200, corner_radius=0, fg_color="#14375e")
        self.nav_frame.pack(side="left", fill="y")
        self.nav_frame.pack_propagate(False)

        ctk.CTkLabel(self.nav_frame, text="THE PY-LIB",
                     font=("Sarabun-Bold", 18), text_color="#ffffff").pack(pady=(20, 10))

        self.nav_buttons = {}

        self.content_frame = ctk.CTkFrame(self, fg_color=COLORS.get("bg", "#f0f0f0"))
        self.content_frame.pack(side="left", fill="both", expand=True)

        self.nav_buttons["home"] = self._nav_button("หน้าหลัก", self.show_dashboard)
        self._show_view(DashboardView)

    def _nav_button(self, text, command):
        btn = ctk.CTkButton(self.nav_frame, text=text, command=command,
                            anchor="w", fg_color="transparent", hover_color="#1f538d",
                            corner_radius=0, height=40)
        btn.pack(fill="x", pady=(2, 0))
        return btn

    def _ensure_nav(self, label, command):
        if label not in self.nav_buttons:
            self.nav_buttons[label] = self._nav_button(label, command)

    def _show_view(self, view_cls):
        if self.current_view is not None:
            try:
                self.current_view.destroy()
            except Exception as exc:
                logger.warning(f"Error destroying view: {exc}")
        self.current_view = view_cls(self.content_frame, self)

    # ----- navigation targets -----
    def show_dashboard(self):
        for label, btn in list(self.nav_buttons.items()):
            if label != "home":
                btn.destroy()
                self.nav_buttons.pop(label, None)
        self._show_view(DashboardView)

    def show_members(self):
        self._ensure_nav("จัดการสมาชิก", self.show_members)
        self._show_view(MemberView)

    def show_books(self):
        self._ensure_nav("จัดการหนังสือ", self.show_books)
        self._show_view(BookView)

    def show_borrow(self):
        self._ensure_nav("ยืมหนังสือ", self.show_borrow)
        self._show_view(BorrowView)

    def show_return(self):
        self._ensure_nav("คืนหนังสือ", self.show_return)
        self._show_view(ReturnView)

    def show_history(self):
        self._ensure_nav("ประวัติยืม-คืน", self.show_history)
        self._show_view(HistoryView)

    def show_access_scanner(self):
        self._ensure_nav("บันทึกเข้าออก", self.show_access_scanner)
        self._show_view(AccessScannerView)

    def show_access_history(self):
        self._ensure_nav("ประวัติเข้าออก", self.show_access_history)
        self._show_view(AccessHistoryView)

    def show_settings(self):
        self._ensure_nav("ตั้งค่า", self.show_settings)
        self._show_view(SettingsView)

    def show_about(self):
        logger.info("About requested from dashboard")
        self._ensure_nav("เกี่ยวกับ", self.show_about)
        self._show_view(DashboardView)

    # ---------------- run ----------------
    def run(self):
        self.mainloop()