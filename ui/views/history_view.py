# ui/views/history_view.py
import customtkinter as ctk
from datetime import datetime

from db.repositories.borrow_repo import BorrowRepository
from services.pdf_service import PDFService
from core.logger import logger


class HistoryView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent)
        self.nav = nav
        self.repo = BorrowRepository()
        self._build()

    def _show_toast(self, msg, color="green"):
        lbl = ctk.CTkLabel(self.winfo_toplevel(), text=msg, text_color=color)
        lbl.pack(pady=5)
        self.after(2500, lbl.destroy)

    def _build(self):
        self.pack(pady=20, padx=40, fill="both", expand=True)
        ctk.CTkLabel(self, text="ประวัติการยืม-คืน", font=("Helvetica", 24)).pack(pady=20)

        search = ctk.CTkFrame(self)
        search.pack(pady=10, padx=20, fill="x")

        self.member_entry = ctk.CTkEntry(search, placeholder_text="ค้นหาตามชื่อสมาชิก")
        self.member_entry.pack(side="left", padx=5, fill="x", expand=True)

        self.book_entry = ctk.CTkEntry(search, placeholder_text="ค้นหาตามรหัสหนังสือ")
        self.book_entry.pack(side="left", padx=5, fill="x", expand=True)

        self.date_entry = ctk.CTkEntry(search, placeholder_text="ค้นหาตามวันที่ (YYYY-MM-DD)")
        self.date_entry.pack(side="left", padx=5, fill="x", expand=True)

        ctk.CTkButton(search, text="ค้นหา", command=self._search).pack(side="left", padx=5)
        ctk.CTkButton(search, text="Export PDF", command=self._export_pdf).pack(side="left", padx=5)

        self.display_frame = ctk.CTkScrollableFrame(self)
        self.display_frame.pack(pady=10, padx=20, fill="both", expand=True)

        ctk.CTkButton(self, text="กลับ", command=self.nav.show_dashboard).pack(pady=10)

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
        query += " ORDER BY bl.borrow_date DESC"

        records = self.repo.fetchall(query, params)
        if not records:
            ctk.CTkLabel(self.display_frame, text="ไม่พบประวัติ").pack(pady=10)
            return

        for r in records:
            status = "คืนแล้ว" if r[7] else "ยังไม่คืน"
            row = ctk.CTkFrame(self.display_frame)
            row.pack(pady=5, padx=10, fill="x")
            ctk.CTkLabel(row,
                         text=(f"สมาชิก: {r[0]} ({r[1]}/{r[2]}) | "
                               f"หนังสือ: {r[3]} - {r[4]} | "
                               f"ยืม: {r[5]} | คืน: {r[6]} | {status}")).pack(pady=5, padx=10)

    def _search(self):
        self._load_records(self.member_entry.get(), self.book_entry.get(), self.date_entry.get())

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
            filename = f"history_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            PDFService.generate_borrow_history_pdf(filename, records)
            self._show_toast(f"Export สำเร็จ: {filename}")
        except Exception as e:
            logger.error(f"export_history error: {e}", exc_info=True)
            self._show_toast(f"เกิดข้อผิดพลาด: {e}", "red")
