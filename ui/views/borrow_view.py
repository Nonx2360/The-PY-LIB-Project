# ui/views/borrow_view.py
import customtkinter as ctk
from datetime import datetime, timedelta

import ui.theme as theme
from config import settings
from db.repositories.book_repo import BookRepository
from db.repositories.borrow_repo import BorrowRepository
from db.repositories.member_repo import MemberRepository
from ui.widgets.scan_window import ScanWindow
from ui.widgets.toast import show_toast
from core.logger import logger


class BorrowView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self.member_repo = MemberRepository()
        self.book_repo = BookRepository()
        self.borrow_repo = BorrowRepository()
        self.current_member = None
        self.current_book = None
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)
        page = ctk.CTkFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True, padx=28, pady=24)

        theme.heading(page, "ยืมหนังสือ", size=26).pack(anchor="w")
        theme.subheading(page, "สแกนบัตรสมาชิก แล้วค้นหาหนังสือที่ต้องการยืม").pack(
            anchor="w", pady=(2, 18))

        # ============ step 1: member ============
        step1 = theme.card(page)
        step1.pack(fill="x", pady=(0, 12))
        head = ctk.CTkFrame(step1, fg_color="transparent")
        head.pack(fill="x", padx=18, pady=(14, 8))
        ctk.CTkLabel(head, text="ขั้นตอนที่ 1", font=theme.font(12, bold=True),
                     text_color=theme.PRIMARY).pack(side="left")
        ctk.CTkLabel(head, text="สแกนบัตรสมาชิก", font=theme.font(16, bold=True),
                     text_color=theme.text_color()).pack(side="left", padx=(8, 0))

        body1 = ctk.CTkFrame(step1, fg_color="transparent")
        body1.pack(fill="x", padx=18, pady=(0, 14))
        body1.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(body1, text="ยังไม่ได้สแกนบัตร",
                                         font=theme.font(14),
                                         text_color=theme.muted_color(), anchor="w")
        self.status_label.grid(row=0, column=0, sticky="w")

        self.member_info_label = ctk.CTkLabel(body1, text="", font=theme.font(14, bold=True),
                                              text_color=theme.SUCCESS, anchor="w")
        self.member_info_label.grid(row=1, column=0, sticky="w", pady=(4, 0))

        theme.primary_button(body1, "สแกนบัตรสมาชิก", self._open_scan_window, height=42,
                             width=220).grid(row=0, column=1, rowspan=2, padx=(16, 0))

        # ============ step 2: book ============
        step2 = theme.card(page)
        step2.pack(fill="x", pady=(0, 12))
        head2 = ctk.CTkFrame(step2, fg_color="transparent")
        head2.pack(fill="x", padx=18, pady=(14, 8))
        ctk.CTkLabel(head2, text="ขั้นตอนที่ 2", font=theme.font(12, bold=True),
                     text_color=theme.INFO).pack(side="left")
        ctk.CTkLabel(head2, text="กรอกรหัสหนังสือ", font=theme.font(16, bold=True),
                     text_color=theme.text_color()).pack(side="left", padx=(8, 0))

        body2 = ctk.CTkFrame(step2, fg_color="transparent")
        body2.pack(fill="x", padx=18, pady=(0, 14))
        body2.grid_columnconfigure(0, weight=1)
        self.book_code_entry = theme.entry(body2, placeholder_text="เช่น B-001")
        self.book_code_entry.grid(row=0, column=0, sticky="ew")
        theme.primary_button(body2, "ค้นหา", self._search_book, width=110,
                             height=38, font=theme.font(14, bold=True)).grid(
            row=0, column=1, padx=(10, 0))
        self.book_code_entry.bind("<Return>", lambda e: self._search_book())

        # ============ step 3: confirm ============
        step3 = theme.card(page)
        step3.pack(fill="x")
        head3 = ctk.CTkFrame(step3, fg_color="transparent")
        head3.pack(fill="x", padx=18, pady=(14, 8))
        ctk.CTkLabel(head3, text="ขั้นตอนที่ 3", font=theme.font(12, bold=True),
                     text_color=theme.SUCCESS).pack(side="left")
        ctk.CTkLabel(head3, text="ยืนยันการยืม", font=theme.font(16, bold=True),
                     text_color=theme.text_color()).pack(side="left", padx=(8, 0))

        body3 = ctk.CTkFrame(step3, fg_color="transparent")
        body3.pack(fill="x", padx=18, pady=(0, 16))
        body3.grid_columnconfigure(0, weight=1)

        self.book_info_label = ctk.CTkLabel(body3, text="—", font=theme.font(14),
                                            text_color=theme.muted_color(), anchor="w")
        self.book_info_label.grid(row=0, column=0, sticky="w", pady=(0, 10))

        due_row = ctk.CTkFrame(body3, fg_color="transparent")
        due_row.grid(row=1, column=0, sticky="w")
        ctk.CTkLabel(due_row, text="กำหนดคืน:", font=theme.font(14, bold=True),
                     text_color=theme.text_color()).pack(side="left", padx=(0, 8))
        default_due = (datetime.now() + timedelta(days=settings.get("default_loan_days"))).strftime("%Y-%m-%d")
        self.due_date_entry = theme.entry(due_row, width=120)
        self.due_date_entry.insert(0, default_due)
        self.due_date_entry.pack(side="left")

        quick_frame = ctk.CTkFrame(due_row, fg_color="transparent")
        quick_frame.pack(side="left", padx=10)
        for days in (7, 14, 21):
            theme.secondary_button(quick_frame, f"{days} วัน",
                                   lambda d=days: self._set_quick_date(d),
                                   width=64, height=30).pack(side="left", padx=3)

        self.borrow_button = theme.success_button(body3, "ยืมหนังสือ", self._process_borrow,
                                                  height=44, width=220)
        self.borrow_button.grid(row=2, column=0, sticky="w", pady=(14, 0))

    # ================= helpers =================
    def _set_quick_date(self, days):
        self.due_date_entry.delete(0, "end")
        self.due_date_entry.insert(0,
            (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d"))

    def _open_scan_window(self):
        ScanWindow(self, self._handle_member_qr,
                   on_error=lambda msg: show_toast(self, msg, "red"))

    def _handle_member_qr(self, decoded):
        try:
            name, grade, number, *_ = decoded.split("|")
            row = self.member_repo.fetchone(
                "SELECT id FROM members WHERE name=? AND grade=? AND number=?",
                (name, grade, number))
            if row:
                self.current_member = {"id": row[0], "name": name,
                                       "grade": grade, "number": number}
                self.status_label.configure(text="พบข้อมูลสมาชิก ✓",
                                            text_color=theme.SUCCESS)
                self.member_info_label.configure(
                    text=f"{name}  •  ชั้น {grade}  •  เลขที่ {number}")
                self.book_code_entry.focus()
                show_toast(self, f"ยินดีต้อนรับ {name}", "green")
            else:
                self.status_label.configure(text="ไม่พบสมาชิกในระบบ",
                                            text_color=theme.DANGER)
                self.member_info_label.configure(text="")
                self.current_member = None
                show_toast(self, "ไม่พบสมาชิกในระบบ", "red")
        except Exception as e:
            logger.error(f"QR decode error: {e}", exc_info=True)
            show_toast(self, "QR Code ไม่ถูกต้อง", "red")

    def _search_book(self):
        if not self.current_member:
            show_toast(self, "กรุณาสแกนบัตรสมาชิกก่อน", "red")
            return
        code = self.book_code_entry.get().strip()
        if not code:
            show_toast(self, "กรุณากรอกรหัสหนังสือ", "red")
            return
        book = self.book_repo.get_by_code(code)
        if not book:
            self.book_info_label.configure(text=f"ไม่พบหนังสือรหัส '{code}'")
            self.current_book = None
            self.borrow_button.pack_forget()
            show_toast(self, "ไม่พบหนังสือในระบบ", "red")
            return
        if book[3] != "ว่าง":
            self.book_info_label.configure(text=f"หนังสือ '{book[2]}' ไม่พร้อมให้ยืม (สถานะ: {book[3]})")
            self.current_book = None
            self.borrow_button.pack_forget()
            show_toast(self, "หนังสือถูกยืมอยู่", "red")
            return
        self.book_info_label.configure(
            text=f"รหัส: {book[1]}    ชื่อ: {book[2]}    สถานะ: พร้อมให้ยืม")
        self.book_info_label.configure(text_color=theme.text_color())
        self.current_book = book
        self.borrow_button.grid()

    def _process_borrow(self):
        if not self.current_member or not self.current_book:
            show_toast(self, "ข้อมูลไม่ครบถ้วน", "red")
            return
        due_str = self.due_date_entry.get().strip()
        try:
            due_dt = datetime.strptime(due_str, "%Y-%m-%d")
            if due_dt.date() < datetime.now().date():
                show_toast(self, "กำหนดคืนต้องไม่เป็นวันที่ผ่านมาแล้ว", "red")
                return
        except ValueError:
            show_toast(self, "รูปแบบวันที่ไม่ถูกต้อง (YYYY-MM-DD)", "red")
            return

        mid = self.current_member["id"]
        overdue = self.borrow_repo.fetchone(
            "SELECT COUNT(*) FROM borrow_log WHERE member_id=? AND returned=0 AND return_due < date('now')",
            (mid,))[0]
        if overdue > 0:
            show_toast(self, "ไม่สามารถยืมได้ เนื่องจากมีหนังสือค้างส่ง", "red")
            return
        current = self.borrow_repo.fetchone(
            "SELECT COUNT(*) FROM borrow_log WHERE member_id=? AND returned=0", (mid,))[0]
        max_books = settings.get("max_books_per_member")
        if current >= max_books:
            show_toast(self, f"ยืมครบจำนวนที่กำหนดแล้ว ({max_books} เล่ม)", "red")
            return

        borrow_date = datetime.now().strftime("%Y-%m-%d")
        try:
            self.borrow_repo.add(mid, self.current_book[0], borrow_date, due_str)
            self.book_repo.update_status(self.current_book[0], "ยืมแล้ว")
            show_toast(
                self, f"ยืมสำเร็จ: {self.current_book[2]}\n"
                      f"สมาชิก: {self.current_member['name']}   คืนภายใน: {due_str}",
                "green")
            self.book_code_entry.delete(0, "end")
            self.book_info_label.configure(text="—")
            self.current_book = None
            self.borrow_button.grid_forget()
        except Exception as e:
            logger.error(f"process_borrow error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")
