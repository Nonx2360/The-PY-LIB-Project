# ui/views/book_view.py
import os
import customtkinter as ctk
from tkinter import filedialog

from db.repositories.book_repo import BookRepository
from services.excel_service import ExcelService
from ui.widgets.confirm_dialog import show_confirm_dialog
from core.logger import logger


class BookView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent)
        self.nav = nav
        self.repo = BookRepository()
        self._build()

    def _show_toast(self, msg, color="green"):
        lbl = ctk.CTkLabel(self.winfo_toplevel(), text=msg, text_color=color)
        lbl.pack(pady=5)
        self.after(2500, lbl.destroy)

    def _build(self):
        self.pack(pady=20, padx=40, fill="both", expand=True)
        ctk.CTkLabel(self, text="จัดการหนังสือ", font=("Helvetica", 24)).pack(pady=20)

        form = ctk.CTkFrame(self)
        form.pack(pady=10, padx=20, fill="x")

        self.code_entry = ctk.CTkEntry(form, placeholder_text="รหัสหนังสือ")
        self.code_entry.pack(pady=5, padx=20, fill="x")

        self.title_entry = ctk.CTkEntry(form, placeholder_text="ชื่อเรื่อง")
        self.title_entry.pack(pady=5, padx=20, fill="x")

        ctk.CTkButton(form, text="เพิ่มหนังสือ", command=self._add_book).pack(pady=10)
        ctk.CTkButton(form, text="นำเข้าจาก Excel", command=self._import_excel).pack(pady=5)
        ctk.CTkButton(form, text="ดาวน์โหลดแม่แบบ Excel", command=self._export_template).pack(pady=5)

        ctk.CTkButton(self, text="กลับ", command=self.nav.show_dashboard).pack(pady=10)

        self._display_books()

    def _add_book(self):
        code = self.code_entry.get().strip()
        title = self.title_entry.get().strip()
        if not all([code, title]):
            self._show_toast("กรุณากรอกข้อมูลให้ครบ", "red")
            return
        try:
            self.repo.add(code, title, status="ว่าง")
            self._show_toast("เพิ่มหนังสือสำเร็จ")
            self.nav.show_books()
        except Exception as e:
            logger.error(f"add_book error: {e}", exc_info=True)
            self._show_toast(f"เกิดข้อผิดพลาด: {e}", "red")

    def _display_books(self):
        scroll = ctk.CTkScrollableFrame(self)
        scroll.pack(pady=10, padx=20, fill="both", expand=True)

        for b in self.repo.get_all():
            row = ctk.CTkFrame(scroll)
            row.pack(pady=5, padx=10, fill="x")
            ctk.CTkLabel(row, text=f"รหัส: {b[1]} | ชื่อ: {b[2]} | สถานะ: {b[3]}").pack(
                side="left", padx=10)
            ctk.CTkButton(row, text="ลบ",
                          command=lambda bb=b: self._delete_book(bb)).pack(side="right", padx=5)

    def _delete_book(self, book):
        # Check currently borrowed
        from db.repositories.base_repo import BaseRepository
        base = BaseRepository()
        active = base.fetchone(
            "SELECT id FROM borrow_log WHERE book_id = ? AND returned = 0", (book[0],))
        if active:
            self._show_toast("ไม่สามารถลบหนังสือที่กำลังถูกยืมอยู่ได้", "red")
            return
        if not show_confirm_dialog(self, "ยืนยันการลบ",
                                   f"ลบหนังสือ '{book[2]}' ใช่หรือไม่?"):
            return
        try:
            self.repo.delete(book[0])
            self._show_toast("ลบหนังสือสำเร็จ")
            self.nav.show_books()
        except Exception as e:
            logger.error(f"delete_book error: {e}", exc_info=True)
            self._show_toast(f"เกิดข้อผิดพลาด: {e}", "red")

    def _export_template(self):
        try:
            df = ExcelService.generate_template_df()
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="book_import_template.xlsx")
            if filename:
                df.to_excel(filename, index=False)
                self._show_toast(f"บันทึกแม่แบบ Excel สำเร็จ: {filename}")
                if os.name == "nt":
                    os.startfile(filename)
        except Exception as e:
            self._show_toast(f"เกิดข้อผิดพลาด: {e}", "red")

    def _import_excel(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="เลือกไฟล์ Excel ที่ต้องการนำเข้า")
        if not filename:
            return
        try:
            books = ExcelService.read_books_from_excel(filename)
        except ValueError as e:
            self._show_toast(str(e), "red")
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

        msg = f"นำเข้าสำเร็จ: {ok} รายการ"
        if err:
            msg += "\n" + "\n".join(err[:5])
        color = "green" if ok > 0 else "red"
        self._show_toast(msg, color)
        if ok > 0:
            self.nav.show_books()
