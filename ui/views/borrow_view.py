# ui/views/borrow_view.py
import threading, time
import customtkinter as ctk
from PIL import Image
from datetime import datetime, timedelta
from queue import Queue

from db.repositories.member_repo import MemberRepository
from db.repositories.book_repo import BookRepository
from db.repositories.borrow_repo import BorrowRepository
from services.camera_service import CameraService
from core.logger import logger


class BorrowView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent)
        self.nav = nav
        self.member_repo = MemberRepository()
        self.book_repo = BookRepository()
        self.borrow_repo = BorrowRepository()
        self.current_member = None
        self.current_book = None
        self._build()

    def _show_toast(self, msg, color="green"):
        lbl = ctk.CTkLabel(self.winfo_toplevel(), text=msg, text_color=color)
        lbl.pack(pady=5)
        self.after(2500, lbl.destroy)

    def _build(self):
        self.pack(pady=20, padx=40, fill="both", expand=True)
        ctk.CTkLabel(self, text="ยืมหนังสือ", font=("Helvetica", 24)).pack(pady=20)

        # Member section
        member_frame = ctk.CTkFrame(self)
        member_frame.pack(pady=10, padx=20, fill="x")

        self.status_label = ctk.CTkLabel(member_frame, text="กรุณาสแกนบัตรสมาชิก",
                                         font=("Helvetica", 18))
        self.status_label.pack(pady=10)

        self.member_info_label = ctk.CTkLabel(member_frame, text="", font=("Helvetica", 14))
        self.member_info_label.pack(pady=5)

        ctk.CTkButton(member_frame, text="สแกนบัตร",
                      command=self._open_scan_window,
                      height=60, font=("Helvetica", 20)).pack(pady=20)

        # Book section
        book_frame = ctk.CTkFrame(self)
        book_frame.pack(pady=10, padx=20, fill="x")

        self.book_code_entry = ctk.CTkEntry(book_frame, placeholder_text="กรอกรหัสหนังสือ")
        self.book_code_entry.pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkButton(book_frame, text="ค้นหา", command=self._search_book,
                      width=100).pack(side="right", padx=5)

        # Book info + due date
        self.book_info_frame = ctk.CTkFrame(self)
        self.book_info_frame.pack(pady=10, padx=20, fill="x")

        self.book_info_label = ctk.CTkLabel(self.book_info_frame, text="", font=("Helvetica", 14))
        self.book_info_label.pack(pady=10)

        # Due date row
        due_row = ctk.CTkFrame(self.book_info_frame)
        due_row.pack(pady=5, fill="x")
        ctk.CTkLabel(due_row, text="กำหนดคืน:", font=("Helvetica", 14)).pack(side="left", padx=5)

        default_due = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")
        self.due_date_entry = ctk.CTkEntry(due_row, width=120)
        self.due_date_entry.insert(0, default_due)
        self.due_date_entry.pack(side="left", padx=5)

        quick_frame = ctk.CTkFrame(due_row)
        quick_frame.pack(side="left", padx=5)
        ctk.CTkButton(quick_frame, text="7 วัน", width=60,
                      command=lambda: self._set_quick_date(7)).pack(side="left", padx=2)
        ctk.CTkButton(quick_frame, text="14 วัน", width=60,
                      command=lambda: self._set_quick_date(14)).pack(side="left", padx=2)
        ctk.CTkLabel(due_row, text="(YYYY-MM-DD)", font=("Helvetica", 12),
                     text_color="gray").pack(side="left", padx=5)

        self.borrow_button = ctk.CTkButton(self.book_info_frame, text="ยืมหนังสือ",
                                           command=self._process_borrow,
                                           font=("Helvetica", 16))
        # Initially hidden; shown after book search

        ctk.CTkButton(self, text="กลับ", command=self.nav.show_dashboard).pack(pady=10)
        self.book_code_entry.bind("<Return>", lambda e: self._search_book())

    def _set_quick_date(self, days):
        self.due_date_entry.delete(0, "end")
        self.due_date_entry.insert(0, (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d"))

    def _open_scan_window(self):
        scan_win = ctk.CTkToplevel(self)
        scan_win.title("สแกน QR Code")
        scan_win.geometry("800x600")

        video_frame = ctk.CTkFrame(scan_win)
        video_frame.pack(pady=10, padx=10, fill="both", expand=True)
        video_label = ctk.CTkLabel(video_frame, text="")
        video_label.pack(fill="both", expand=True)

        ctk.CTkLabel(scan_win, text="นำ QR Code มาวางตรงกล้อง\nรอสักครู่ระบบจะสแกนอัตโนมัติ",
                     font=("Helvetica", 16)).pack(pady=10)
        status_label = ctk.CTkLabel(scan_win, text="กำลังรอสแกน...",
                                    font=("Helvetica", 14), text_color="yellow")
        status_label.pack(pady=5)

        cap = CameraService.get_camera()
        if cap is None:
            self._show_toast("ไม่สามารถเปิดกล้องได้", "red")
            scan_win.destroy()
            return

        qr_queue = Queue()

        def scan_loop():
            while True:
                if not scan_win.winfo_exists():
                    break
                ret, frame = cap.read()
                if not ret:
                    break
                processed, decoded = CameraService.process_frame(frame)
                if decoded:
                    qr_queue.put(decoded)
                    scan_win.after(1000, lambda: [scan_win.destroy(), cap.release()])
                    return
                rgb = __import__("cv2").cvtColor(processed, __import__("cv2").COLOR_BGR2RGB)
                pil = Image.fromarray(rgb)
                ctk_img = ctk.CTkImage(light_image=pil, dark_image=pil, size=(640, 480))
                video_label.configure(image=ctk_img)
                time.sleep(0.033)
            cap.release()

        t = threading.Thread(target=scan_loop, daemon=True)
        t.start()

        def poll_qr():
            if not qr_queue.empty():
                decoded = qr_queue.get()
                try:
                    name, grade, number, *_ = decoded.split("|")
                    row = self.member_repo.fetchone(
                        "SELECT id FROM members WHERE name=? AND grade=? AND number=?",
                        (name, grade, number))
                    if row:
                        self.current_member = {"id": row[0], "name": name,
                                               "grade": grade, "number": number}
                        self.status_label.configure(text="พบข้อมูลสมาชิก", text_color="green")
                        self.member_info_label.configure(
                            text=f"สมาชิก: {name}\nชั้น: {grade}\nเลขที่: {number}")
                        self.book_code_entry.focus()
                    else:
                        self.status_label.configure(text="ไม่พบสมาชิกในระบบ", text_color="red")
                except Exception as e:
                    logger.error(f"QR decode error: {e}")
                return
            if scan_win.winfo_exists():
                scan_win.after(100, poll_qr)

        scan_win.after(100, poll_qr)
        scan_win.protocol("WM_DELETE_WINDOW", lambda: [cap.release(), scan_win.destroy()])

    def _search_book(self):
        if not self.current_member:
            self._show_toast("กรุณาสแกนบัตรสมาชิกก่อน", "red")
            return
        code = self.book_code_entry.get().strip()
        if not code:
            self._show_toast("กรุณากรอกรหัสหนังสือ", "red")
            return
        book = self.book_repo.get_by_code(code)
        if not book:
            self.book_info_label.configure(text="ไม่พบหนังสือในระบบ")
            self.current_book = None
            self.borrow_button.pack_forget()
            return
        if book[3] != "ว่าง":
            self.book_info_label.configure(text=f"หนังสือไม่พร้อมให้ยืม สถานะ: {book[3]}")
            self.current_book = None
            self.borrow_button.pack_forget()
            return
        self.book_info_label.configure(text=f"รหัส: {book[1]}\nชื่อ: {book[2]}\nสถานะ: {book[3]}")
        self.current_book = book
        self.borrow_button.pack(pady=10)

    def _process_borrow(self):
        if not self.current_member or not self.current_book:
            self._show_toast("ข้อมูลไม่ครบถ้วน", "red")
            return
        due_str = self.due_date_entry.get().strip()
        try:
            due_dt = datetime.strptime(due_str, "%Y-%m-%d")
            if due_dt.date() < datetime.now().date():
                self._show_toast("กำหนดคืนต้องไม่เป็นวันที่ผ่านมาแล้ว", "red")
                return
        except ValueError:
            self._show_toast("รูปแบบวันที่ไม่ถูกต้อง (YYYY-MM-DD)", "red")
            return

        mid = self.current_member["id"]
        # Overdue check
        overdue = self.borrow_repo.fetchone(
            "SELECT COUNT(*) FROM borrow_log WHERE member_id=? AND returned=0 AND return_due < date('now')",
            (mid,))[0]
        if overdue > 0:
            self._show_toast("ไม่สามารถยืมได้ เนื่องจากมีหนังสือค้างส่ง", "red")
            return
        # Limit check
        current = self.borrow_repo.fetchone(
            "SELECT COUNT(*) FROM borrow_log WHERE member_id=? AND returned=0", (mid,))[0]
        if current >= 3:
            self._show_toast("ยืมครบจำนวนที่กำหนดแล้ว (3 เล่ม)", "red")
            return

        borrow_date = datetime.now().strftime("%Y-%m-%d")
        try:
            self.borrow_repo.add(mid, self.current_book[0], borrow_date, due_str)
            self.book_repo.update_status(self.current_book[0], "ยืมแล้ว")
            self._show_toast(
                f"ยืมสำเร็จ: {self.current_book[2]}\nสมาชิก: {self.current_member['name']}\nคืน: {due_str}")
            self.book_code_entry.delete(0, "end")
            self.book_info_label.configure(text="")
            self.current_book = None
            self.borrow_button.pack_forget()
        except Exception as e:
            logger.error(f"process_borrow error: {e}", exc_info=True)
            self._show_toast(f"เกิดข้อผิดพลาด: {e}", "red")
