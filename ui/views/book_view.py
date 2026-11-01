# ui/views/book_view.py
import os
import queue
import threading

import customtkinter as ctk
from tkinter import filedialog

import ui.theme as theme
from config.constants import (BOOK_STATUS_AVAILABLE, BOOK_STATUS_BORROWED,
                              BOOK_STATUS_RETIRED)
from db.repositories.base_repo import BaseRepository
from db.repositories.book_repo import BookRepository
from services.excel_service import ExcelService
from ui.widgets.confirm_dialog import show_confirm_dialog
from ui.widgets.row_window import RowWindow
from ui.widgets.toast import show_toast
from core.logger import logger


class BookView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self.repo = BookRepository()
        self.base = BaseRepository()
        self._search_job = None
        self._import_running = False
        self._selected = set()          # bulk-management selection (ids)
        self._bulk_count_label = None
        self._bulk_built = False
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
        self.import_btn = theme.secondary_button(
            form, "เลือกไฟล์ Excel เพื่อนำเข้า", self._import_excel, height=38)
        self.import_btn.pack(fill="x", pady=(0, 6))
        self.import_status = ctk.CTkLabel(form, text="", font=theme.font(11),
                                          text_color=theme.muted_color(),
                                          anchor="center")
        self.import_status.pack(fill="x", pady=(0, 4))
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
        self.search_entry.bind("<KeyRelease>", self._on_search_key)

        # bulk action bar — hidden until rows are ticked; sits between the
        # header and the bounded list (selection survives paging/search)
        self.bulk_bar = ctk.CTkFrame(list_card, fg_color=theme.PRIMARY,
                                     corner_radius=10)

        # bounded list window: renders only the first 50 rows plus a
        # "show more" button — building a row costs ~10 Tk windows, so
        # rendering all 1000+ rows at once froze the app and could exhaust
        # Windows GDI handles (Tk then panics — uncatchable).
        self.list_frame = RowWindow(list_card, empty_note="ไม่พบหนังสือ")
        self.list_frame.on_status = self._update_count
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._display_books()

    # ================= data =================
    def _update_count(self, _shown, total):
        if not self._alive():
            return
        self.count_label.configure(text=f"รายการหนังสือ ({total})")

    def _alive(self):
        """False once the view has been destroyed (user navigated away)."""
        try:
            return bool(self.winfo_exists())
        except Exception:
            return False

    def _on_search_key(self, _e=None):
        # rebuilds are bounded but cost a few hundred ms — debounce so a
        # burst of keystrokes doesn't queue several rebuilds
        if self._search_job is not None:
            self.after_cancel(self._search_job)
        self._search_job = self.after(350, self._search_now)

    def _search_now(self):
        self._search_job = None
        self._display_books()

    def _display_books(self):
        self._search_job = None
        # newest first, so a freshly imported batch sits on the first page
        books = self.repo.fetchall("SELECT * FROM books ORDER BY id DESC")
        q = self.search_entry.get().strip().lower()
        if q:
            books = [b for b in books if q in (b[1] or "").lower() or q in (b[2] or "").lower()]

        self.list_frame.render(books, self._book_row)

    def _book_row(self, b):
        row = theme.card(self.list_frame, corner_radius=10)
        row.pack(fill="x", padx=4, pady=4)
        row.grid_columnconfigure(2, weight=1)

        check = ctk.CTkCheckBox(row, text="", width=24,
                                checkbox_width=18, checkbox_height=18)
        check.grid(row=0, column=0, rowspan=2, padx=(10, 4), pady=8)
        if b[0] in self._selected:
            check.select()
        # command is wired after assignment so the lambda can close over the
        # very checkbox it belongs to (no forward-reference at creation)
        check.configure(command=lambda bid=b[0], c=check: self._toggle_select(bid, c))

        ctk.CTkFrame(row, width=4, height=40, fg_color=theme.INFO,
                     corner_radius=2).grid(row=0, column=1, rowspan=2, padx=(4, 10), pady=8)
        ctk.CTkLabel(row, text=b[2], font=theme.font(15, bold=True),
                     text_color=theme.text_color(), anchor="w").grid(
            row=0, column=2, sticky="w", pady=(8, 0))
        ctk.CTkLabel(row, text=f"รหัส {b[1]}", font=theme.font(12),
                     text_color=theme.muted_color(), anchor="w").grid(
            row=1, column=2, sticky="w", pady=(0, 8))

        status = (b[3] or "").strip()
        status_badge = (theme.badge_green(row, BOOK_STATUS_AVAILABLE) if status in (BOOK_STATUS_AVAILABLE, "Available", "")
                        else theme.badge_blue(row, BOOK_STATUS_RETIRED) if status == BOOK_STATUS_RETIRED
                        else theme.badge_red(row, BOOK_STATUS_BORROWED) if "ยืม" in status
                        else theme.badge_amber(row, status))
        status_badge.grid(row=0, column=3, rowspan=2, padx=6)

        if status == BOOK_STATUS_AVAILABLE:
            theme.ghost_button(row, "หยุด", lambda bb=b: self._retire_book(bb),
                               width=56, height=30).grid(row=0, column=4, rowspan=2, padx=(6, 0))
            theme.ghost_button(row, "ลบ", lambda bb=b: self._delete_book(bb),
                               width=56, height=30, text_color=theme.DANGER).grid(
                row=0, column=5, rowspan=2, padx=(6, 10))
        elif status == BOOK_STATUS_RETIRED:
            theme.ghost_button(row, "กลับมา", lambda bb=b: self._restore_book(bb),
                               width=56, height=30).grid(row=0, column=4, rowspan=2, padx=(6, 0))
            theme.ghost_button(row, "ลบถาวร", lambda bb=b: self._delete_book(bb),
                               width=76, height=30, text_color=theme.DANGER).grid(
                row=0, column=5, rowspan=2, padx=(6, 10))
        else:
            theme.ghost_button(row, "ลบ", lambda bb=b: self._delete_book(bb),
                               width=56, height=30, text_color=theme.DANGER).grid(
                row=0, column=4, rowspan=2, padx=(6, 10))

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
            self.repo.add(code, title, status=BOOK_STATUS_AVAILABLE)
            show_toast(self, f"เพิ่มหนังสือสำเร็จ: {title}", "green")
            self.code_entry.delete(0, "end")
            self.title_entry.delete(0, "end")
            self._display_books()
        except Exception as e:
            logger.error(f"add_book error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")

    def _retire_book(self, book):
        # try_return() sets the book back to 'ว่าง' on every return, so a
        # borrowed book must be returned BEFORE it can be withdrawn —
        # otherwise the return would silently un-retire it.
        if (book[3] or "").strip() != "ว่าง":
            show_toast(self, "หนังสือไม่อยู่ที่ชั้น ต้องคืนก่อนจึงหยุดให้บริการได้", "red")
            return
        try:
            self.repo.update_status(book[0], "หยุดให้บริการ")
            show_toast(self, "หยุดให้บริการแล้ว — ประวัติการยืมยังถูกเก็บไว้", "green")
            self._display_books()
        except Exception as e:
            logger.error(f"retire_book error: {e}", exc_info=True)
            show_toast(self, "เกิดข้อผิดพลาด ไม่สามารถหยุดให้บริการได้", "red")

    def _restore_book(self, book):
        try:
            self.repo.update_status(book[0], "ว่าง")
            show_toast(self, "หนังสือกลับมาให้บริการปกติแล้ว", "green")
            self._display_books()
        except Exception as e:
            logger.error(f"restore_book error: {e}", exc_info=True)
            show_toast(self, "เกิดข้อผิดพลาด ไม่สามารถกลับมาให้บริการได้", "red")

    def _delete_book(self, book):
        status = (book[3] or "").strip()
        if status == "หยุดให้บริการ":
            # two-step purge: retired copies may be erased permanently,
            # together with their borrow_log rows, in one transaction
            history = self.base.fetchone(
                "SELECT COUNT(*) FROM borrow_log WHERE book_id = ?", (book[0],))[0]
            if not show_confirm_dialog(
                    self, "ยืนยันการลบถาวร",
                    f"ลบ '{book[2]}' พร้อมประวัติการยืม {history} รายการแบบถาวร?\n"
                    "ลบแล้วไม่สามารถย้อนกลับได้"):
                return

            def _purge(conn):
                conn.execute("DELETE FROM borrow_log WHERE book_id=?", (book[0],))
                conn.execute("DELETE FROM books WHERE id=?", (book[0],))

            try:
                self.base.transact(_purge)
            except Exception as e:
                logger.error(f"purge_book error: {e}", exc_info=True)
                show_toast(self, "ลบไม่สำเร็จ กรุณาลองใหม่", "red")
                return
            show_toast(self, "ลบหนังสือและประวัติการยืมถาวรแล้ว", "green")
            self._display_books()
            return
        active = self.base.fetchone(
            "SELECT id FROM borrow_log WHERE book_id = ? AND returned = 0", (book[0],))
        if active:
            show_toast(self, "ไม่สามารถลบหนังสือที่กำลังถูกยืมอยู่ได้", "red")
            return
        # borrow_log -> books is a FOREIGN KEY: books with borrow history
        # cannot be deleted while in circulation — point to the retire path
        history = self.base.fetchone(
            "SELECT id FROM borrow_log WHERE book_id = ? LIMIT 1", (book[0],))
        if history:
            show_toast(self, "ลบไม่ได้ — เล่มนี้มีประวัติการยืม กดปุ่ม หยุด ก่อน แล้วจึงลบถาวรได้", "red")
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
            show_toast(self, "ลบไม่สำเร็จ — ข้อมูลยังถูกอ้างอิงจากประวัติการใช้งาน", "red")

    # ================= bulk management =================
    def _toggle_select(self, book_id, box):
        if box.get():
            self._selected.add(book_id)
        else:
            self._selected.discard(book_id)
        self._sync_bulk_bar()

    def _sync_bulk_bar(self):
        """Show the bar when rows are ticked, hide it when none remain."""
        n = len(self._selected)
        if n == 0:
            self._bulk_count_label = None
            if self.bulk_bar.winfo_manager():
                self.bulk_bar.pack_forget()
            return
        if not self.bulk_bar.winfo_manager():
            if not self._bulk_built:
                self._build_bulk_bar()
                self._bulk_built = True
            # keep the bar visually above the list even though it was
            # created before the list frame got packed
            self.bulk_bar.pack(fill="x", padx=10, pady=(0, 6),
                               before=self.list_frame)
        if self._bulk_count_label is not None:
            self._bulk_count_label.configure(text=f"เลือกไว้ {n} รายการ")

    def _build_bulk_bar(self):
        label = ctk.CTkLabel(self.bulk_bar, text="",
                             font=theme.font(13, bold=True),
                             text_color="#FFFFFF")
        label.pack(side="left", padx=(12, 6), pady=8)
        self._bulk_count_label = label
        theme.secondary_button(self.bulk_bar, "หยุดให้บริการ",
                               self._bulk_retire, height=30).pack(side="left", padx=3)
        theme.secondary_button(self.bulk_bar, "ให้บริการกลับมา",
                               self._bulk_restore, height=30).pack(side="left", padx=3)
        theme.secondary_button(self.bulk_bar, "ลบที่เลือก",
                               self._bulk_delete, height=30,
                               text_color=theme.DANGER).pack(side="left", padx=3)
        theme.secondary_button(self.bulk_bar, "เคลียร์การเลือก",
                               self._clear_selection, height=30).pack(side="left", padx=(3, 10))

    def _clear_selection(self):
        self._selected.clear()
        self._sync_bulk_bar()
        if self._alive():
            self._display_books()    # rebuild rows so no checkbox stays ticked

    def _bulk_retire(self):
        ids = sorted(self._selected)
        if not ids or not show_confirm_dialog(
                self, "ยืนยันการหยุดให้บริการ",
                f"หยุดให้บริการหนังสือที่เลือก {len(ids)} เล่ม?\n"
                "เล่มที่กำลังถูกยืมอยู่จะถูกข้าม"):
            return
        changed = self.repo.bulk_retire(ids)
        skipped = len(ids) - changed
        show_toast(self, f"สำเร็จ {changed} / ข้าม {skipped} รายการ",
                   "green" if changed and not skipped else "amber")
        logger.info("desktop books bulk: retire requested=%d changed=%d",
                    len(ids), changed)
        self._clear_selection()

    def _bulk_restore(self):
        ids = sorted(self._selected)
        if not ids or not show_confirm_dialog(
                self, "ยืนยันการให้บริการกลับมา",
                f"ให้บริการหนังสือที่เลือกกลับมา {len(ids)} เล่ม?\n"
                "เล่มที่ไม่ได้หยุดให้บริการจะถูกข้าม"):
            return
        changed = self.repo.bulk_restore(ids)
        skipped = len(ids) - changed
        show_toast(self, f"สำเร็จ {changed} / ข้าม {skipped} รายการ",
                   "green" if changed and not skipped else "amber")
        logger.info("desktop books bulk: restore requested=%d changed=%d",
                    len(ids), changed)
        self._clear_selection()

    def _bulk_delete(self):
        ids = sorted(self._selected)
        if not ids or not show_confirm_dialog(
                self, "ยืนยันการลบ",
                f"ลบหนังสือที่เลือก {len(ids)} เล่ม?\n"
                "เล่มที่หยุดให้บริการจะถูกลบถาวรพร้อมประวัติการยืม "
                "(ลบแล้วไม่สามารถย้อนกลับ)\n"
                "เล่มที่มีประวัติการยืมแต่ยังให้บริการอยู่จะถูกข้าม"):
            return
        deleted, skipped = self.repo.delete_many(ids)
        show_toast(self, f"ลบสำเร็จ {deleted} / ข้าม {skipped} รายการ",
                   "green" if deleted and not skipped else "amber")
        logger.info("desktop books bulk: delete requested=%d deleted=%d "
                    "skipped=%d", len(ids), deleted, skipped)
        self._clear_selection()

    def _export_template(self):
        try:
            df = ExcelService.generate_template_df()
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="book_import_template.xlsx")
            if filename:
                df.to_excel(filename, index=False)
                show_toast(self, "บันทึกแม่แบบ Excel สำเร็จ", "green")
                if os.name == "nt":
                    os.startfile(filename)
        except Exception as e:
            logger.error(f"export_template error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")

    def _import_excel(self):
        if self._import_running:
            show_toast(self, "กำลังนำเข้าอยู่ กรุณารอจนเสร็จก่อน", "amber")
            return
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="เลือกไฟล์ Excel ที่ต้องการนำเข้า")
        if not filename:
            return

        # Reading the workbook + inserting it (one transaction) can take a
        # few seconds for a 1000-row file; run it on a worker thread so the
        # window never freezes. Results travel back through a queue that
        # this view polls with after() on the Tk thread — tkinter widgets
        # are never touched from the worker thread.
        self._import_running = True
        self.import_btn.configure(state="disabled")
        self.import_status.configure(text="กำลังอ่านและนำเข้าข้อมูล...")

        results = queue.Queue()
        threading.Thread(target=self._import_worker, args=(filename, results),
                         daemon=True, name="excel-import").start()
        self._poll_import(results)

    def _import_worker(self, filename, results):
        try:
            logger.info("import excel: reading workbook (background thread)")
            books = ExcelService.read_books_from_excel(filename)
            logger.info("import excel: parsed rows=%d", len(books))
            ok, skipped = 0, 0
            if books:
                # single BEGIN IMMEDIATE transaction via add_many —
                # one commit for the whole batch, duplicates skipped
                ok, skipped = self.repo.add_many(
                    [(b["code"], b["title"]) for b in books])
        except ValueError as e:
            logger.warning("import excel rejected: %s", e)
            results.put(("bad-template", str(e)))
        except Exception as e:
            logger.error("import excel failed: %s", e, exc_info=True)
            results.put(("error", "นำเข้าไม่สำเร็จ กรุณาลองใหม่"))
        else:
            logger.info("import excel: inserted=%d skipped=%d", ok, skipped)
            results.put(("done", (ok, skipped)))

    def _poll_import(self, results):
        if not self._alive():
            return
        try:
            kind, payload = results.get_nowait()
        except queue.Empty:
            self.after(80, lambda: self._poll_import(results))
            return

        self._import_running = False
        if not self._alive():
            return
        self.import_btn.configure(state="normal")
        self.import_status.configure(text="")

        if kind == "done":
            ok, skipped = payload
            msg = f"นำเข้าสำเร็จ: {ok} รายการ"
            if skipped:
                msg += f" (ข้ามซ้ำ/ไม่สมบูรณ์ {skipped} รายการ)"
            logger.info("import excel: refresh list after insert=%d", ok)
            show_toast(self, msg, "green" if ok > 0 else "red")
            if ok > 0:
                self._display_books()
        elif kind == "bad-template":
            show_toast(self, payload, "red")
        else:
            show_toast(self, payload, "red")
