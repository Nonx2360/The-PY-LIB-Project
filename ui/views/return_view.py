# ui/views/return_view.py
import threading, time
import customtkinter as ctk
from PIL import Image
from queue import Queue

from db.repositories.member_repo import MemberRepository
from db.repositories.book_repo import BookRepository
from db.repositories.borrow_repo import BorrowRepository
from services.camera_service import CameraService
from core.logger import logger


class ReturnView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent)
        self.nav = nav
        self.member_repo = MemberRepository()
        self.book_repo = BookRepository()
        self.borrow_repo = BorrowRepository()
        self._current_member_id = None
        self._build()

    def _show_toast(self, msg, color="green"):
        lbl = ctk.CTkLabel(self.winfo_toplevel(), text=msg, text_color=color)
        lbl.pack(pady=5)
        self.after(2500, lbl.destroy)

    def _build(self):
        self.pack(pady=20, padx=40, fill="both", expand=True)
        ctk.CTkLabel(self, text="คืนหนังสือ", font=("Helvetica", 24)).pack(pady=20)

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

        self.books_frame = ctk.CTkScrollableFrame(self)
        self.books_frame.pack(pady=10, padx=20, fill="both", expand=True)

        ctk.CTkButton(self, text="กลับ", command=self.nav.show_dashboard).pack(pady=10)

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

        threading.Thread(target=scan_loop, daemon=True).start()

        def poll_qr():
            if not qr_queue.empty():
                decoded = qr_queue.get()
                try:
                    name, grade, number, *_ = decoded.split("|")
                    row = self.member_repo.fetchone(
                        "SELECT id FROM members WHERE name=? AND grade=? AND number=?",
                        (name, grade, number))
                    if row:
                        self._current_member_id = row[0]
                        self.status_label.configure(text="พบข้อมูลสมาชิก", text_color="green")
                        self.member_info_label.configure(
                            text=f"สมาชิก: {name}\nชั้น: {grade}\nเลขที่: {number}")
                        self._load_borrowed_books(row[0])
                    else:
                        self.status_label.configure(text="ไม่พบสมาชิกในระบบ", text_color="red")
                except Exception as e:
                    logger.error(f"QR decode error: {e}")
                return
            if scan_win.winfo_exists():
                scan_win.after(100, poll_qr)

        scan_win.after(100, poll_qr)
        scan_win.protocol("WM_DELETE_WINDOW", lambda: [cap.release(), scan_win.destroy()])

    def _load_borrowed_books(self, member_id):
        for w in self.books_frame.winfo_children():
            w.destroy()

        books = self.borrow_repo.fetchall('''
            SELECT b.id, b.code, b.title, bl.borrow_date, bl.return_due
            FROM books b
            JOIN borrow_log bl ON b.id = bl.book_id
            WHERE bl.member_id = ? AND bl.returned = 0
        ''', (member_id,))

        if not books:
            ctk.CTkLabel(self.books_frame, text="ไม่มีหนังสือที่ยืมอยู่",
                         font=("Helvetica", 14)).pack(pady=10)
            return

        ctk.CTkLabel(self.books_frame, text="รายการหนังสือที่ยืม",
                     font=("Helvetica", 16, "bold")).pack(pady=(0, 10))

        for book in books:
            row = ctk.CTkFrame(self.books_frame)
            row.pack(pady=5, padx=10, fill="x")
            ctk.CTkLabel(row,
                         text=f"รหัส: {book[1]}\nชื่อ: {book[2]}\nยืม: {book[3]} | คืน: {book[4]}",
                         font=("Helvetica", 12)).pack(side="left", padx=10, pady=5)
            ctk.CTkButton(row, text="คืน", width=80, height=32,
                          command=lambda bb=book: self._return_book(bb[0], member_id)
                          ).pack(side="right", padx=10, pady=5)

    def _return_book(self, book_id, member_id):
        try:
            self.borrow_repo.mark_returned(book_id)
            self.book_repo.update_status(book_id, "ว่าง")
            self._show_toast("คืนหนังสือสำเร็จ")
            self._load_borrowed_books(member_id)
        except Exception as e:
            logger.error(f"return_book error: {e}", exc_info=True)
            self._show_toast(f"เกิดข้อผิดพลาด: {e}", "red")
