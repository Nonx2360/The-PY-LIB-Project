# ui/views/settings_view.py
import customtkinter as ctk
from core.logger import logger


class SettingsView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent)
        self.nav = nav
        self._build()

    def _build(self):
        self.pack(pady=20, padx=40, fill="both", expand=True)

        ctk.CTkLabel(self, text="ตั้งค่า",
                     font=("Sarabun-Bold", 28), text_color="#1f538d").pack(pady=(10, 20))

        form = ctk.CTkFrame(self)
        form.pack(fill="x", padx=40)

        ctk.CTkLabel(form, text="ธีมการแสดงผล:", font=("Sarabun", 14)).pack(pady=(10, 0))

        themes = ["System", "Dark", "Light"]
        selected = ctk.StringVar(value=ctk.get_appearance_mode())
        theme_row = ctk.CTkFrame(form, fg_color="transparent")
        theme_row.pack(pady=5)
        for theme in themes:
            rb = ctk.CTkRadioButton(
                theme_row, text=theme, variable=selected, value=theme,
                command=lambda t=theme: self._apply_theme(t),
                font=("Sarabun", 13),
            )
            rb.pack(side="left", padx=15)

        info = ctk.CTkLabel(
            self,
            text="THE PY-LIB (ระบบห้องสมุดอัจฉริยะ)\n"
                 "เวอร์ชัน 2.0 — โครงสร้างแบบแยกชั้น (layered)\n"
                 "ความปลอดภัย: bcrypt + AES-256-CBC (QR)",
            font=("Sarabun", 14),
            justify="center",
        )
        info.pack(pady=(30, 0))

        ctk.CTkButton(self, text="กลับหน้าหลัก", width=200, height=40,
                      font=("Sarabun", 14, "bold"),
                      command=self.nav.show_dashboard).pack(pady=30)

    def _apply_theme(self, theme):
        try:
            ctk.set_appearance_mode(theme)
            logger.info(f"Appearance theme changed to {theme}")
        except Exception as exc:
            logger.error(f"Failed to apply theme: {exc}")