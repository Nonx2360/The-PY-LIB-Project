# ui/views/webui_view.py — Web UI launcher control panel
import threading
import webbrowser
import customtkinter as ctk

import ui.theme as theme
from core.logger import logger
from ui.widgets.toast import show_toast


class WebUIView(ctk.CTkFrame):
    """Sidebar panel to start/stop the Flask Web UI server."""

    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self._server = None
        self._thread = None
        self._running = False
        self._port = 5000
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)
        page = ctk.CTkFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True, padx=28, pady=24)

        theme.heading(page, "Web UI", size=26).pack(anchor="w")
        theme.subheading(page, "เปิดเว็บไซต์สำหรับใช้งานผ่านเบราว์เซอร์").pack(anchor="w", pady=(2, 18))

        # ── status card ──
        self._status_card = theme.card(page)
        self._status_card.pack(fill="x", pady=(0, 14))
        theme.card_title(self._status_card, "สถานะเซิร์ฟเวอร์").pack(anchor="w", padx=18, pady=(16, 10))

        body = ctk.CTkFrame(self._status_card, fg_color="transparent")
        body.pack(fill="x", padx=18, pady=(0, 18))

        self._status_label = ctk.CTkLabel(body, text="● ปิดอยู่", font=theme.font(15, bold=True),
                                          text_color="#EF4444")
        self._status_label.pack(anchor="w", pady=(0, 6))

        self._url_label = ctk.CTkLabel(body, text="", font=theme.font(13),
                                       text_color=theme.muted_color())
        self._url_label.pack(anchor="w", pady=(0, 12))

        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(fill="x")

        self._start_btn = theme.primary_button(btn_row, "▶  เริ่ม Web UI", self._start_server,
                                               height=40, width=160)
        self._start_btn.pack(side="left", padx=(0, 8))

        self._stop_btn = theme.danger_button(btn_row, "■  หยุด", self._stop_server,
                                             height=40, width=120, state="disabled")
        self._stop_btn.pack(side="left", padx=(0, 8))

        self._open_btn = theme.secondary_button(btn_row, "🌐  เปิดในเบราว์เซอร์",
                                                self._open_browser, height=40, width=180,
                                                state="disabled")
        self._open_btn.pack(side="left")

        # ── info card ──
        info_card = theme.card(page)
        info_card.pack(fill="x")
        theme.card_title(info_card, "ข้อมูล").pack(anchor="w", padx=18, pady=(16, 10))

        info_body = ctk.CTkFrame(info_card, fg_color="transparent")
        info_body.pack(fill="x", padx=18, pady=(0, 18))

        rows = [
            ("พอร์ต", str(self._port)),
            ("ฟีเจอร์", "สมาชิก, หนังสือ, ยืม-คืน, ประวัติ, ตั้งค่า"),
            ("ล็อกอิน", "ใช้รหัสผ่านเดียวกับแอปเดสท็อป"),
            ("ข้อจำกัด", "ไม่รองรับ QR สแกนจากกล้อง (ต้องใช้เดสท็อป)"),
        ]
        for k, v in rows:
            r = ctk.CTkFrame(info_body, fg_color="transparent")
            r.pack(fill="x", pady=2)
            ctk.CTkLabel(r, text=k, font=theme.font(13, bold=True),
                         text_color=theme.muted_color(), width=100, anchor="w").pack(side="left")
            ctk.CTkLabel(r, text=v, font=theme.font(13),
                         text_color=theme.text_color(), anchor="w").pack(
                side="left", fill="x", expand=True)

    def _start_server(self):
        if self._running:
            return
        try:
            from web.app import create_app
            app = create_app()

            def run():
                try:
                    app.run(host="0.0.0.0", port=self._port, debug=False, use_reloader=False)
                except OSError as e:
                    logger.error(f"Web UI port {self._port} error: {e}")

            self._thread = threading.Thread(target=run, daemon=True)
            self._thread.start()
            self._running = True

            self._status_label.configure(text="● กำลังทำงาน", text_color="#22C55E")
            self._url_label.configure(text=f"WebUI Running at : http://127.0.0.1:{self._port}")
            self._start_btn.configure(state="disabled")
            self._stop_btn.configure(state="normal")
            self._open_btn.configure(state="normal")

            logger.info(f"Web UI started on port {self._port}")
            show_toast(self, f"Web UI เริ่มทำงานแล้ว — http://127.0.0.1:{self._port}", "green")
        except Exception as e:
            logger.error(f"Failed to start Web UI: {e}", exc_info=True)
            show_toast(self, f"ไม่สามารถเริ่ม Web UI: {e}", "red")

    def _stop_server(self):
        if not self._running:
            return
        self._running = False
        self._status_label.configure(text="● ปิดอยู่", text_color="#EF4444")
        self._url_label.configure(text="")
        self._start_btn.configure(state="normal")
        self._stop_btn.configure(state="disabled")
        self._open_btn.configure(state="disabled")
        logger.info("Web UI stopped")
        show_toast(self, "Web UI หยุดแล้ว", "orange")

    def _open_browser(self):
        url = f"http://127.0.0.1:{self._port}"
        webbrowser.open(url)
        logger.info(f"Opened browser: {url}")
