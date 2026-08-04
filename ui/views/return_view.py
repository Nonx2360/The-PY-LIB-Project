# ui/views/return_view.py
import customtkinter as ctk
from datetime import datetime

import ui.theme as theme
from db.repositories.book_repo import BookRepository
from db.repositories.borrow_repo import BorrowRepository
from db.repositories.member_repo import MemberRepository
from ui.widgets.scan_window import ScanWindow
from ui.widgets.toast import show_toast
from core.logger import logger


class ReturnView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self.member_repo = MemberRepository()
        self.book_repo = BookRepository()
        self.borrow_repo = BorrowRepository()
        self._current_member = None
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)
        page = ctk.CTkFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True, padx=28, pady=24)

        theme.heading(page, "คืนหนังสือ", size=26).pack(anchor="w")
        theme.subheading(page, "สแกนบัตรสมาชิก เพื่อดูหนังสือที่ยืมอยู่").pack(
            anchor="w", pady=(2, 18))

        # ---- member card ----
        member_card = theme.card(page)
        member_card.pack(fill="x", pady=(0, 14))
        head = ctk.CTkFrame(member_card, fg_color="transparent")
        head.pack(fill="x", padx=18, pady=(14, 8))
        ctk.CTkLabel(head, text="สมาชิก", font=theme.font(16, bold=True),
                     text_color=theme.text_color()).pack(side="left")
        theme.primary_button(head, "สแกนบัตรสมาชิก", self._open_scan_window,
                             width=200, height=32, font=theme.font(13, bold=True)).pack(
            side="right")

        body = ctk.CTkFrame(member_card, fg_color="transparent")
        body.pack(fill="x", padx=18, pady=(0, 14))
        self.status_label = ctk.CTkLabel(body, text="ยังไม่ได้สแกนบัตร",
                                         font=theme.font(14), text_color=theme.muted_color())
        self.status_label.pack(anchor="w")
        self.member_info_label = ctk.CTkLabel(body, text="", font=theme.font(14, bold=True),
                                              text_color=theme.SUCCESS)
        self.member_info_label.pack(anchor="w", pady=(4, 0))

        # ---- borrowed books ----
        list_card = theme.card(page)
        list_card.pack(fill="both", expand=True)
        top = ctk.CTkFrame(list_card, fg_color="transparent")
        top.pack(fill="x", padx=18, pady=(14, 6))
        self.books_title = ctk.CTkLabel(top, text="หนังสือที่ยืมอยู่",
                                        font=theme.font(16, bold=True),
                                        text_color=theme.text_color(), anchor="w")
        self.books_title.pack(side="left")

        self.books_frame = ctk.CTkScrollableFrame(list_card, fg_color="transparent",
                                                  corner_radius=0)
        self.books_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        ctk.CTkLabel(self.books_frame, text="สแกนบัตรสมาชิกเพื่อดูรายการ",
                     font=theme.font(14), text_color=theme.muted_color()).pack(pady=24)

    # ================= scan =================
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
                self._current_member = {"id": row[0], "name": name,
                                        "grade": grade, "number": number}
                self.status_label.configure(text="พบข้อมูลสมาชิก ✓",
                                            text_color=theme.SUCCESS)
                self.member_info_label.configure(
                    text=f"{name}  •  ชั้น {grade}  •  เลขที่ {number}")
                self._load_borrowed_books(row[0])
            else:
                self.status_label.configure(text="ไม่พบสมาชิกในระบบ",
                                            text_color=theme.DANGER)
                self.member_info_label.configure(text="")
                self._clear_books("ไม่พบสมาชิกในระบบ")
                show_toast(self, "ไม่พบสมาชิกในระบบ", "red")
        except Exception as e:
            logger.error(f"QR decode error: {e}", exc_info=True)
            show_toast(self, "QR Code ไม่ถูกต้อง", "red")

    # ================= data =================
    def _clear_books(self, msg):
        for w in self.books_frame.winfo_children():
            w.destroy()
        ctk.CTkLabel(self.books_frame, text=msg, font=theme.font(14),
                     text_color=theme.muted_color()).pack(pady=24)

    def _load_borrowed_books(self, member_id):
        for w in self.books_frame.winfo_children():
            w.destroy()

        books = self.borrow_repo.fetchall('''
            SELECT bl.id, b.code, b.title, bl.borrow_date, bl.return_due
            FROM books b
            JOIN borrow_log bl ON b.id = bl.book_id
            WHERE bl.member_id = ? AND bl.returned = 0
            ORDER BY bl.return_due ASC
        ''', (member_id,))

        if not books:
            self.books_title.configure(text="หนังสือที่ยืมอยู่ (0)")
            ctk.CTkLabel(self.books_frame, text="ไม่มีหนังสือที่ยืมอยู่",
                         font=theme.font(14), text_color=theme.muted_color()).pack(pady=24)
            return

        self.books_title.configure(text=f"หนังสือที่ยืมอยู่ ({len(books)})")
        today = datetime.now().date()

        for book in books:
            bl_id, code, title, borrow_date, due = book
            row = theme.card(self.books_frame, corner_radius=10)
            row.pack(fill="x", padx=4, pady=4)
            row.grid_columnconfigure(1, weight=1)

            overdue = False
            try:
                overdue = datetime.strptime(due, "%Y-%m-%d").date() < today
            except ValueError:
                pass

            stripe = ctk.CTkFrame(row, width=4, height=40,
                                  fg_color=theme.DANGER if overdue else theme.SUCCESS,
                                  corner_radius=2)
            stripe.grid(row=0, column=0, rowspan=2, padx=(10, 12), pady=8)
            ctk.CTkLabel(row, text=title, font=theme.font(15, bold=True),
                         text_color=theme.text_color(), anchor="w").grid(
                row=0, column=1, sticky="w", pady=(8, 0))
            ctk.CTkLabel(row, text=f"รหัส {code}  •  ยืมเมื่อ {borrow_date}",
                         font=theme.font(12), text_color=theme.muted_color()).grid(
                row=1, column=1, sticky="w", pady=(0, 8))

            (theme.badge_red(row, f"ค้างส่ง • กำหนด {due}")
             if overdue else theme.badge_amber(row, f"กำหนดคืน {due}")).grid(
                row=0, column=3, rowspan=2, padx=4)

            # fetch the book_id for this borrow record so we can update book status
            book_row = self.borrow_repo.fetchone(
                "SELECT book_id FROM borrow_log WHERE id = ?", (bl_id,))
            book_id = book_row[0] if book_row else None
            theme.success_button(
                row, "คืนหนังสือ",
                lambda b=bl_id, bk=book_id: self._return_book(b, bk),
                width=110, height=32, font=theme.font(13, bold=True)).grid(
                row=0, column=4, rowspan=2, padx=(6, 10))

    def _return_book(self, record_id, book_id):
        try:
            logger.debug(f"Returning borrow_log id={record_id}, book_id={book_id}")
            self.borrow_repo.mark_returned(record_id)
            if book_id is not None:
                self.book_repo.update_status(book_id, "ว่าง")
            show_toast(self, "คืนหนังสือสำเร็จ", "green")
            if self._current_member:
                logger.debug(f"Reloading books for member id={self._current_member['id']}")
                self._load_borrowed_books(self._current_member["id"])
        except Exception as e:
            logger.error(f"return_book error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")