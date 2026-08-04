# ui/views/access_view.py
import threading, time
import customtkinter as ctk
from PIL import Image
from datetime import datetime
from queue import Queue

from db.repositories.access_repo import AccessRepository
from db.repositories.member_repo import MemberRepository
from services.camera_service import CameraService
from services.pdf_service import PDFService
from core.logger import logger


class AccessScannerView(ctk.CTkFrame):
    """Records entry/exit of library members via QR scan."""
    def __init__(self, parent, nav):
        super().__init__(parent)
        self.nav = nav
        self.access_repo = AccessRepository()
        self.member_repo = MemberRepository()
        self._build()

    def _show_toast(self, msg, color="green"):
        lbl = ctk.CTkLabel(self.winfo_toplevel(), text=msg, text_color=color)
        lbl.pack(pady=5)
        self.after(2500, lbl.destroy)

    def _build(self):
        self.pack(pady=20, padx=40, fill="both", expand=True)
        ctk.CTkLabel(self, text="บันทึกการเข้าออกห้องสมุด",
                     font=("Helvetica", 24)).pack(pady=20)

        self.status_label = ctk.CTkLabel(self, text="กรุณาสแกนบัตรสมาชิก",
                                         font=("Helvetica", 18))
        self.status_label.pack(pady=10)

        self.last_scan_label = ctk.CTkLabel(self, text="", font=("Helvetica", 14))
        self.last_scan_label.pack(pady=10)

        ctk.CTkButton(self, text="สแกนบัตร", command=self._open_scan_window,
                      height=60, font=("Helvetica", 20)).pack(pady=20)

        ctk.CTkButton(self, text="กลับ", command=self.nav.show_dashboard).pack(pady=10)

    def _open_scan_window(self):
        scan_win = ctk.CTkToplevel(self)
        scan_win.title("สแกน QR Code")
        scan_win.geometry("800x600")

        vid_frame = ctk.CTkFrame(scan_win)
        vid_frame.pack(pady=10, padx=10, fill="both", expand=True)
        vid_label = ctk.CTkLabel(vid_frame, text="")
        vid_label.pack(fill="both", expand=True)

        ctk.CTkLabel(scan_win,
                     text="นำ QR Code มาวางตรงกล้อง\nรอสักครู่ระบบจะสแกนอัตโนมัติ",
                     font=("Helvetica", 16)).pack(pady=10)
        status = ctk.CTkLabel(scan_win, text="กำลังรอสแกน...",
                              font=("Helvetica", 14), text_color="yellow")
        status.pack(pady=5)

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
                vid_label.configure(image=ctk_img)
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
                        mid = row[0]
                        last = self.access_repo.fetchone(
                            "SELECT action FROM access_log WHERE member_id=? ORDER BY id DESC LIMIT 1",
                            (mid,))
                        action = "ออก" if (last and last[0] == "เข้า") else "เข้า"
                        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        self.access_repo.log_access(mid, ts, action)
                        self.status_label.configure(
                            text=f"บันทึกการ{action}ห้องสมุดสำเร็จ", text_color="green")
                        self.last_scan_label.configure(
                            text=f"สมาชิก: {name}\nชั้น: {grade}\nเลขที่: {number}\nเวลา: {ts}")
                    else:
                        self.status_label.configure(text="ไม่พบสมาชิกในระบบ", text_color="red")
                except Exception as e:
                    logger.error(f"Access scan error: {e}", exc_info=True)
                return
            if scan_win.winfo_exists():
                scan_win.after(100, poll_qr)

        scan_win.after(100, poll_qr)
        scan_win.protocol("WM_DELETE_WINDOW", lambda: [cap.release(), scan_win.destroy()])


class AccessHistoryView(ctk.CTkFrame):
    """Displays and exports entry/exit history."""
    def __init__(self, parent, nav):
        super().__init__(parent)
        self.nav = nav
        self.repo = AccessRepository()
        self._build()

    def _show_toast(self, msg, color="green"):
        lbl = ctk.CTkLabel(self.winfo_toplevel(), text=msg, text_color=color)
        lbl.pack(pady=5)
        self.after(2500, lbl.destroy)

    def _build(self):
        self.pack(pady=20, padx=40, fill="both", expand=True)
        ctk.CTkLabel(self, text="ประวัติการเข้าออก", font=("Helvetica", 24)).pack(pady=20)

        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=20, pady=10)
        ctk.CTkButton(top, text="Export PDF", command=self._export_pdf).pack(side="right", padx=5)

        self.table_frame = ctk.CTkScrollableFrame(self)
        self.table_frame.pack(pady=10, padx=20, fill="both", expand=True)

        ctk.CTkButton(self, text="กลับ", command=self.nav.show_dashboard).pack(pady=10)

        self._load()

    def _load(self):
        for w in self.table_frame.winfo_children():
            w.destroy()

        records = self.repo.get_recent_access(limit=100)
        if not records:
            ctk.CTkLabel(self.table_frame, text="ไม่พบประวัติ").pack(pady=10)
            return

        for r in records:
            row = ctk.CTkFrame(self.table_frame)
            row.pack(fill="x", padx=5, pady=2)
            ctk.CTkLabel(row, text=f"{r[0]}  ชื่อ: {r[1]} ({r[2]})  {r[3]}").pack(
                side="left", padx=10)

    def _export_pdf(self):
        try:
            records = self.repo.get_recent_access(limit=1000)
            # Convert tuple (access_time, name, grade, action) to the 5-col expected
            rows = [(r[1], r[2], "", r[0], r[3]) for r in records]
            from datetime import datetime
            filename = f"access_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            import os
            os.makedirs("reports", exist_ok=True)
            filename = f"reports/{filename}"
            PDFService.generate_access_history_pdf(filename, rows)
            self._show_toast(f"Export สำเร็จ: {filename}")
        except Exception as e:
            logger.error(f"export_access error: {e}", exc_info=True)
            self._show_toast(f"เกิดข้อผิดพลาด: {e}", "red")
