# ui/widgets/pdf_viewer.py - built-in PDF viewer window
import os
import customtkinter as ctk
from PIL import Image
import fitz

import ui.theme as theme
from core.logger import logger


class PDFViewerWindow:
    """Opens a Toplevel window to view a PDF file with page navigation."""

    def __init__(self, parent, pdf_path, title="PDF Viewer"):
        if not os.path.exists(pdf_path):
            logger.error(f"PDF not found: {pdf_path}")
            return

        self.win = ctk.CTkToplevel(parent)
        self.win.title(title)
        self.win.geometry("820x700")
        self.win.minsize(640, 520)
        self.win.grab_set()

        self.pdf_path = pdf_path
        self.current_page = 0
        self.doc = None
        self._closed = False

        try:
            self.doc = fitz.open(pdf_path)
        except Exception as e:
            logger.error(f"Failed to open PDF: {e}")
            ctk.CTkLabel(self.win, text=f"ไม่สามารถเปิด PDF ได้\n{e}",
                         font=theme.font(14), text_color=theme.DANGER).pack(pady=40)
            return

        self.total_pages = len(self.doc)

        # ---- header ----
        header = ctk.CTkFrame(self.win, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(12, 4))
        ctk.CTkLabel(header, text=os.path.basename(pdf_path),
                     font=theme.font(15, bold=True), text_color=theme.text_color(),
                     anchor="w").pack(side="left", fill="x", expand=True)
        ctk.CTkButton(header, text="ปิด", command=self.close, width=70, height=30,
                      font=theme.font(12), fg_color="transparent",
                      border_width=1, text_color=theme.text_color(),
                      border_color=theme.border_color(),
                      hover_color=theme.border_color(),
                      corner_radius=6).pack(side="right")

        # ---- page label ----
        self.page_label = ctk.CTkLabel(self.win, text="", font=theme.font(13),
                                       text_color=theme.muted_color())
        self.page_label.pack(pady=(0, 4))

        # ---- canvas for PDF page ----
        self.canvas_frame = ctk.CTkFrame(self.win, fg_color="transparent")
        self.canvas_frame.pack(fill="both", expand=True, padx=16)

        self.page_label_img = ctk.CTkLabel(self.canvas_frame, text="กำลังโหลด...",
                                           font=theme.font(14), text_color=theme.muted_color())
        self.page_label_img.pack(fill="both", expand=True)

        # ---- navigation ----
        nav = ctk.CTkFrame(self.win, fg_color="transparent")
        nav.pack(fill="x", padx=16, pady=(8, 14))

        self.prev_btn = theme.secondary_button(nav, "← ก่อนหน้า", self._prev_page,
                                               width=120, height=34)
        self.prev_btn.pack(side="left")

        self.next_btn = theme.primary_button(nav, "ถัดไป →", self._next_page,
                                             width=120, height=34)
        self.next_btn.pack(side="right")

        self.page_info = ctk.CTkLabel(nav, text="", font=theme.font(13),
                                      text_color=theme.text_color())
        self.page_info.pack(side="left", expand=True)

        self._render_page()
        self.win.protocol("WM_DELETE_WINDOW", self.close)

    def _render_page(self):
        if self._closed or self.doc is None:
            return
        try:
            page = self.doc[self.current_page]
            # render at 2x for crisp display
            mat = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # fit to window width
            canvas_w = self.canvas_frame.winfo_width()
            if canvas_w < 10:
                canvas_w = 760
            scale = canvas_w / img.width
            new_h = int(img.height * scale)
            img = img.resize((canvas_w, new_h), Image.LANCZOS)

            ctk_img = ctk.CTkImage(light_image=img, dark_image=img,
                                   size=(canvas_w, new_h))
            self.page_label_img.configure(image=ctk_img, text="")
            self.page_label_img.image = ctk_img

            self.page_label.configure(
                text=f"หน้า {self.current_page + 1} / {self.total_pages}")
            self.page_info.configure(
                text=f"หน้า {self.current_page + 1} / {self.total_pages}")

            self.prev_btn.configure(state="normal" if self.current_page > 0 else "disabled")
            self.next_btn.configure(state="normal" if self.current_page < self.total_pages - 1 else "disabled")
        except Exception as e:
            logger.error(f"PDF render error: {e}", exc_info=True)
            self.page_label_img.configure(text=f"แสดงหน้าไม่ได้: {e}", image=None)

    def _prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._render_page()

    def _next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self._render_page()

    def close(self):
        if self._closed:
            return
        self._closed = True
        try:
            if self.doc:
                self.doc.close()
        except Exception:
            pass
        try:
            if self.win.winfo_exists():
                self.win.destroy()
        except Exception:
            pass
