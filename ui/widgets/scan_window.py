# ui/widgets/scan_window.py - shared camera QR scan window
import threading
import time
from queue import Queue

import customtkinter as ctk
from PIL import Image

from services.camera_service import CameraService
from core.logger import logger
from ui.theme import font, muted_color, text_color


class ScanWindow:
    """Opens a camera window, scans a QR code, then calls on_decode(text).

    The window closes itself automatically once a code is decoded.
    If the camera cannot be opened, on_error() is called (if provided).
    """

    def __init__(self, parent, on_decode, on_error=None, title="สแกน QR Code"):
        self.on_decode = on_decode
        self.on_error = on_error
        self._closed = False

        self.win = ctk.CTkToplevel(parent)
        self.win.title(title)
        self.win.geometry("760x640")
        self.win.minsize(640, 520)
        self.win.grab_set()

        title_lbl = ctk.CTkLabel(self.win, text="สแกน QR Code",
                                 font=font(22, bold=True), text_color=text_color())
        title_lbl.pack(pady=(18, 4))

        hint = ctk.CTkLabel(self.win,
                            text="นำบัตรสมาชิกมาไว้ตรงกล้อง ระบบจะสแกนให้อัตโนมัติ",
                            font=font(14), text_color=muted_color())
        hint.pack(pady=(0, 10))

        video_frame = ctk.CTkFrame(self.win, corner_radius=14,
                                   border_width=2, border_color="#22C55E")
        video_frame.pack(padx=24, pady=6, fill="both", expand=True)
        self.video_label = ctk.CTkLabel(video_frame, text="กำลังเปิดกล้อง...",
                                        font=font(16), text_color=muted_color())
        self.video_label.pack(fill="both", expand=True)

        self.status_label = ctk.CTkLabel(self.win, text="กำลังรอสแกน...",
                                         font=font(15, bold=True), text_color="#F59E0B")
        self.status_label.pack(pady=12)

        close_btn = ctk.CTkButton(self.win, text="ปิดหน้าต่าง", command=self.close,
                                  font=font(13), fg_color="transparent",
                                  border_width=1, text_color=muted_color(),
                                  corner_radius=8)
        close_btn.pack(pady=(0, 16))

        self.cap = CameraService.get_camera()
        if self.cap is None:
            self.status_label.configure(text="ไม่สามารถเปิดกล้องได้", text_color="#EF4444")
            if self.on_error:
                self.on_error("ไม่สามารถเปิดกล้องได้")
            self.win.after(1500, self.close)
            return

        self.qr_queue = Queue()
        threading.Thread(target=self._scan_loop, daemon=True).start()
        self.win.after(100, self._poll_qr)
        self.win.protocol("WM_DELETE_WINDOW", self.close)

    # ---- camera thread ----
    def _scan_loop(self):
        try:
            while True:
                if self._closed or not self.win.winfo_exists():
                    break
                ret, frame = self.cap.read()
                if not ret:
                    break
                processed, decoded = CameraService.process_frame(frame)
                if decoded:
                    self.qr_queue.put(decoded)
                    return
                try:
                    import cv2
                    rgb = cv2.cvtColor(processed, cv2.COLOR_BGR2RGB)
                    pil = Image.fromarray(rgb)
                    img = ctk.CTkImage(light_image=pil, dark_image=pil, size=(640, 480))
                    self.win.after(0, lambda i=img: self._update_video(i))
                except Exception:
                    pass
                time.sleep(0.033)
        except Exception as exc:
            logger.error(f"scan loop error: {exc}", exc_info=True)
        finally:
            self._release()

    def _update_video(self, img):
        try:
            if not self._closed and self.win.winfo_exists():
                self.video_label.configure(image=img, text="")
        except Exception:
            pass

    # ---- main-thread polling ----
    def _poll_qr(self):
        if self._closed:
            return
        if not self.qr_queue.empty():
            decoded = self.qr_queue.get()
            self.status_label.configure(text="พบ QR Code แล้ว", text_color="#22C55E")
            try:
                cb = self.on_decode
            except Exception:
                cb = None
            self.win.after(400, self.close)
            if cb:
                try:
                    cb(decoded)
                except Exception as exc:
                    logger.error(f"scan callback error: {exc}", exc_info=True)
            return
        if self.win.winfo_exists():
            self.win.after(100, self._poll_qr)

    def _release(self):
        try:
            if self.cap is not None:
                self.cap.release()
        except Exception:
            pass

    def close(self):
        if self._closed:
            return
        self._closed = True
        self._release()
        try:
            if self.win.winfo_exists():
                self.win.destroy()
        except Exception:
            pass
