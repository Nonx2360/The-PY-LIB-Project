# ui/views/book_view.py
import os
import customtkinter as ctk
from tkinter import filedialog

import ui.theme as theme
from db.repositories.base_repo import BaseRepository
from db.repositories.book_repo import BookRepository
from services.excel_service import ExcelService
from ui.widgets.confirm_dialog import show_confirm_dialog
from ui.widgets.toast import show_toast
from core.logger import logger


class BookView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self.repo = BookRepository()
        self.base = BaseRepository()
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)
        page = ctk.CTkFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True, padx=28, pady=24)

        theme.heading(page, "จัดการหนังสือ", size=26).pack(anchor="w")
        theme.subheading(page, "เพิ่มหนังสือ หรือนำเข้าจากไฟล์ Excel").pack(anchor="w", pady=(2, 18))

        body = ctk.CTkFrame(page, fg_color="transparent")
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(1, weight=1)

        # ---- left: add form ----
        form_card = theme.card(body)
        form_card.grid(row=0, column=0, padx=(0, 14), sticky="nsw")
        theme.card_title(form_card, "เพิ่มหนังสือ").pack(anchor="w", padx=18, pady=(16, 14))
        form = ctk.CTkFrame(form_card, fg_color="transparent")
        form.pack(fill="x", padx=18, pady=(0, 16))

        ctk.CTkLabel(form, text="รหัสหนังสือ", font=theme.font(12, bold=True),
                     text_color=theme.text_color(), anchor="w").pack(fill="x")
        self.code_entry = theme.entry(form, placeholder_text="เช่น B-001", width=300)
        self.code_entry.pack(fill="x", pady=(4, 10))

        ctk.CTkLabel(form, text="ชื่อเรื่อง", font=theme.font(12, bold=True),
                     text_color=theme.text_color(), anchor="w").pack(fill="x")
        self.title_entry = theme.entry(form, placeholder_text="ชื่อหนังสือ")
        self.title_entry.pack(fill="x", pady=(4, 4))
        self.title_entry.bind("<Return>", lambda e: self._add_book())

        theme.primary_button(form, "เพิ่มหนังสือ", self._add_book, height=42).pack(
            fill="x", pady=(12, 6))

        sep = ctk.CTkFrame(form, height=1, fg_color=theme.border_color())
        sep.pack(fill="x", pady=12)

        theme.subheading(form, "นำเข้าจาก Excel", size=12).pack(anchor="w", pady=(0, 8))
        theme.secondary_button(form, "เลือกไฟล์ Excel เพื่อนำเข้า", self._import_excel,
                               height=38).pack(fill="x", pady=(0, 6))
        theme.ghost_button(form, "ดาวน์โหลดแม่แบบ Excel", self._export_template,
                           height=34).pack(fill="x")

        # ---- right: book list ----
        list_card = theme.card(body)
        list_card.grid(row=0, column=1, sticky="nsew")

        top = ctk.CTkFrame(list_card, fg_color="transparent")
        top.pack(fill="x", padx=18, pady=(14, 10))
        self.count_label = ctk.CTkLabel(top, text="รายการหนังสือ", font=theme.font(17, bold=True),
                                        text_color=theme.text_color(), anchor="w")
        self.count_label.pack(side="left")
        self.search_entry = theme.entry(top, placeholder_text="ค้นหา รหัส / ชื่อ",
                                        width=240)
        self.search_entry.pack(side="right")
        self.search_entry.bind("<KeyRelease>", lambda e: self._display_books())

        self.list_frame = ctk.CTkScrollableFrame(list_card, fg_color="transparent",
                                                 corner_radius=0)
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._display_books()

    # ================= data =================
    def _display_books(self):
        for w in self.list_frame.winfo_children():
            w.destroy()

        books = self.repo.get_all()
        q = self.search_entry.get().strip().lower()
        if q:
            books = [b for b in books if q in (b[1] or "").lower() or q in (b[2] or "").lower()]

        self.count_label.configure(text=f"รายการหนังสือ ({len(books)})")

        if not books:
            ctk.CTkLabel(self.list_frame, text="ไม่พบหนังสือ", font=theme.font(14),
                         text_color=theme.muted_color()).pack(pady=24)
            return

        for b in books:
            self._book_row(b)

    def _book_row(self, b):
        row = theme.card(self.list_frame, corner_radius=10)
        row.pack(fill="x", padx=4, pady=4)
        row.grid_columnconfigure(1, weight=1)

        ctk.CTkFrame(row, width=4, height=40, fg_color=theme.INFO,
                     corner_radius=2).grid(row=0, column=0, rowspan=2, padx=(10, 12), pady=8)
        ctk.CTkLabel(row, text=b[2], font=theme.font(15, bold=True),
                     text_color=theme.text_color(), anchor="w").grid(
            row=0, column=1, sticky="w", pady=(8, 0))
        ctk.CTkLabel(row, text=f"รหัส {b[1]}", font=theme.font(12),
                     text_color=theme.muted_color(), anchor="w").grid(
            row=1, column=1, sticky="w", pady=(0, 8))

        status = (b[3] or "").strip()
        status_badge = (theme.badge_green(row, "ว่าง") if status in ("ว่าง", "Available", "")
                        else theme.badge_red(row, "ยืมแล้ว") if "ยืม" in status
                        else theme.badge_amber(row, status))
        status_badge.grid(row=0, column=2, rowspan=2, padx=6)

        theme.ghost_button(row, "ลบ", lambda bb=b: self._delete_book(bb),
                           width=56, height=30, text_color=theme.DANGER).grid(
            row=0, column=3, rowspan=2, padx=(6, 10))

    # ================= actions =================
    def _add_book(self):
        code = self.code_entry.get().strip()
        title = self.title_entry.get().strip()
        if not all([code, title]):
            show_toast(self, "กรุณากรอกข้อมูลให้ครบ", "red")
            return
        if self.repo.get_by_code(code):
            show_toast(self, f"รหัส {code} มีอยู่แล้วในระบบ", "red")
            return
        try:
            self.repo.add(code, title, status="ว่าง")
            show_toast(self, f"เพิ่มหนังสือสำเร็จ: {title}", "green")
            self.code_entry.delete(0, "end")
            self.title_entry.delete(0, "end")
            self._display_books()
        except Exception as e:
            logger.error(f"add_book error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")

    def _delete_book(self, book):
        active = self.base.fetchone(
            "SELECT id FROM borrow_log WHERE book_id = ? AND returned = 0", (book[0],))
        if active:
            show_toast(self, "ไม่สามารถลบหนังสือที่กำลังถูกยืมอยู่ได้", "red")
            return
        if not show_confirm_dialog(self, "ยืนยันการลบ",
                                   f"ลบหนังสือ '{book[2]}' (รหัส {book[1]}) ใช่หรือไม่?"):
            return
        try:
            self.repo.delete(book[0])
            show_toast(self, "ลบหนังสือสำเร็จ", "green")
            self._display_books()
        except Exception as e:
            logger.error(f"delete_book error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")

    def _export_template(self):
        try:
            df = ExcelService.generate_template_df()
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="book_import_template.xlsx")
            if filename:
                df.to_excel(filename, index=False)
                show_toast(self, f"บันทึกแม่แบบ Excel สำเร็จ", "green")
                if os.name == "nt":
                    os.startfile(filename)
        except Exception as e:
            logger.error(f"export_template error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")

    def _import_excel(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="เลือกไฟล์ Excel ที่ต้องการนำเข้า")
        if not filename:
            return
        try:
            books = ExcelService.read_books_from_excel(filename)
        except ValueError as e:
            show_toast(self, str(e), "red")
            return

        ok, err = 0, []
        for b in books:
            existing = self.repo.get_by_code(b["code"])
            if existing:
                err.append(f"รหัส {b['code']} มีอยู่แล้ว")
                continue
            try:
                self.repo.add(b["code"], b["title"], status="ว่าง")
                ok += 1
            except Exception as e:
                err.append(str(e))

        msg = f"นำเข้าสำเร็จ: {ok} รายการ" + (f"\n{chr(10).join(err[:5])}" if err else "")
        show_toast(self, msg, "green" if ok > 0 else "red")
        if ok > 0:
            self._display_books()
