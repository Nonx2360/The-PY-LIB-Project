# ui/views/history_view.py
import customtkinter as ctk
from datetime import datetime

import ui.theme as theme
from db.repositories.borrow_repo import BorrowRepository
from services.pdf_service import PDFService
from ui.widgets.toast import show_toast
from core.logger import logger


class HistoryView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self.repo = BorrowRepository()
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)
        page = ctk.CTkFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True, padx=28, pady=24)

        theme.heading(page, "ประวัติการยืม-คืน", size=26).pack(anchor="w")
        theme.subheading(page, "ค้นหาประวัติการยืม-คืน และส่งออกเป็น PDF").pack(
            anchor="w", pady=(2, 18))

        # ---- search bar ----
        search_card = theme.card(page)
        search_card.pack(fill="x", pady=(0, 14))
        bar = ctk.CTkFrame(search_card, fg_color="transparent")
        bar.pack(fill="x", padx=18, pady=14)
        bar.grid_columnconfigure((0, 1, 2), weight=1)

        self.member_entry = theme.entry(bar, placeholder_text="ค้นหาชื่อสมาชิก / ชั้น / เลขที่")
        self.member_entry.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        self.book_entry = theme.entry(bar, placeholder_text="ค้นหารหัส / ชื่อหนังสือ")
        self.book_entry.grid(row=0, column=1, padx=8, sticky="ew")
        self.date_entry = theme.entry(bar, placeholder_text="วันที่ (YYYY-MM-DD)")
        self.date_entry.grid(row=0, column=2, padx=(8, 8), sticky="ew")

        self.member_entry.bind("<Return>", lambda e: self._search())
        self.book_entry.bind("<Return>", lambda e: self._search())
        self.date_entry.bind("<Return>", lambda e: self._search())

        theme.primary_button(bar, "ค้นหา", self._search, width=110, height=38,
                             font=theme.font(14, bold=True)).grid(row=0, column=3, padx=(8, 4))
        theme.secondary_button(bar, "ล้าง", self._clear_filters, width=80,
                               height=38).grid(row=0, column=4, padx=4)
        theme.secondary_button(bar, "Export PDF", self._export_pdf, width=120,
                               height=38).grid(row=0, column=5, padx=(4, 0))

        # ---- table ----
        table_card = theme.card(page)
        table_card.pack(fill="both", expand=True)

        header = ctk.CTkFrame(table_card, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(12, 4))
        self.record_count = ctk.CTkLabel(header, text="ประวัติทั้งหมด", font=theme.font(16, bold=True),
                                         text_color=theme.text_color(), anchor="w")
        self.record_count.pack(side="left")

        cols = ["สมาชิก", "หนังสือ", "วันที่ยืม", "กำหนดคืน", "สถานะ"]
        col_frame = ctk.CTkFrame(table_card, fg_color=theme.content_bg(), corner_radius=8)
        col_frame.pack(fill="x", padx=10, pady=(0, 4))
        for i, c in enumerate(cols):
            w = {0: 0.24, 1: 0.34, 2: 0.12, 3: 0.13, 4: 0.10}[i]
            ctk.CTkLabel(col_frame, text=c, font=theme.font(12, bold=True),
                         text_color=theme.muted_color(), anchor="w").pack(
                side="left", expand=True, fill="x", padx=(i * 2, 0))

        self.display_frame = ctk.CTkScrollableFrame(table_card, fg_color="transparent",
                                                    corner_radius=0)
        self.display_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._load_records()

    # ================= data =================
    def _clear_filters(self):
        self.member_entry.delete(0, "end")
        self.book_entry.delete(0, "end")
        self.date_entry.delete(0, "end")
        self._load_records()

    def _load_records(self, member_f="", book_f="", date_f=""):
        for w in self.display_frame.winfo_children():
            w.destroy()

        query = '''
            SELECT m.name, m.grade, m.number, b.code, b.title,
                   bl.borrow_date, bl.return_due, bl.returned
            FROM borrow_log bl
            JOIN members m ON bl.member_id = m.id
            JOIN books b ON bl.book_id = b.id
            WHERE 1=1
        '''
        params = []
        if member_f:
            query += " AND (m.name LIKE ? OR m.grade LIKE ? OR m.number LIKE ?)"
            params.extend([f"%{member_f}%"] * 3)
        if book_f:
            query += " AND (b.code LIKE ? OR b.title LIKE ?)"
            params.extend([f"%{book_f}%"] * 2)
        if date_f:
            query += " AND (bl.borrow_date = ? OR bl.return_due = ?)"
            params.extend([date_f] * 2)
        query += " ORDER BY bl.borrow_date DESC LIMIT 300"

        records = self.repo.fetchall(query, params)
        self.record_count.configure(text=f"ประวัติทั้งหมด ({len(records)})")

        if not records:
            ctk.CTkLabel(self.display_frame, text="ไม่พบประวัติ",
                         font=theme.font(14), text_color=theme.muted_color()).pack(pady=24)
            return

        today = datetime.now().date()
        for r in records:
            name, grade, number, code, title, borrow_date, due, returned = r
            row = ctk.CTkFrame(self.display_frame, fg_color="transparent")
            row.pack(fill="x", padx=2, pady=2)

            overdue = False
            try:
                overdue = (not returned) and datetime.strptime(due, "%Y-%m-%d").date() < today
            except ValueError:
                pass

            member_txt = f"{name} ({grade}/{number})"
            book_txt = f"{code} — {title}"
            ctk.CTkLabel(row, text=member_txt, font=theme.font(13),
                         text_color=theme.text_color(), anchor="w").pack(
                side="left", expand=True, fill="x")
            ctk.CTkLabel(row, text=book_txt, font=theme.font(13),
                         text_color=theme.text_color(), anchor="w").pack(
                side="left", expand=True, fill="x")
            ctk.CTkLabel(row, text=borrow_date, font=theme.font(13),
                         text_color=theme.muted_color(), anchor="w").pack(
                side="left", expand=True, fill="x")
            ctk.CTkLabel(row, text=due, font=theme.font(13),
                         text_color=theme.muted_color(), anchor="w").pack(
                side="left", expand=True, fill="x")

            if returned:
                badge = theme.badge_green(row, "คืนแล้ว")
            elif overdue:
                badge = theme.badge_red(row, "ค้างส่ง")
            else:
                badge = theme.badge_amber(row, "ยังไม่คืน")
            badge.pack(side="left", expand=True, fill="x", padx=8)

            ctk.CTkFrame(row, height=1, fg_color=theme.border_color()).pack(
                side="bottom", fill="x", pady=(6, 0))

    def _search(self):
        self._load_records(self.member_entry.get().strip(),
                           self.book_entry.get().strip(),
                           self.date_entry.get().strip())

    # ================= export =================
    def _export_pdf(self):
        try:
            records = self.repo.fetchall('''
                SELECT m.name, m.grade, m.number, b.code, b.title,
                       bl.borrow_date, bl.return_due, bl.returned
                FROM borrow_log bl
                JOIN members m ON bl.member_id = m.id
                JOIN books b ON bl.book_id = b.id
                ORDER BY bl.borrow_date DESC
            ''')
            import os
            os.makedirs("reports", exist_ok=True)
            filename = f"reports/history_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            PDFService.generate_borrow_history_pdf(filename, records)
            show_toast(self, f"Export สำเร็จ: {filename}", "green")
        except Exception as e:
            logger.error(f"export_history error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")
