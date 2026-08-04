# ui/views/settings_view.py
import customtkinter as ctk

import ui.theme as theme
from core.logger import logger


class SettingsView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)
        page = ctk.CTkFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True, padx=28, pady=24)

        theme.heading(page, "ตั้งค่า", size=26).pack(anchor="w")
        theme.subheading(page, "ปรับแต่งการแสดงผล และข้อมูลระบบ").pack(anchor="w", pady=(2, 18))

        # ---- appearance ----
        card1 = theme.card(page)
        card1.pack(fill="x", pady=(0, 14))
        theme.card_title(card1, "การแสดงผล").pack(anchor="w", padx=18, pady=(16, 6))

        body1 = ctk.CTkFrame(card1, fg_color="transparent")
        body1.pack(fill="x", padx=18, pady=(0, 18))
        ctk.CTkLabel(body1, text="ธีมการแสดงผล:", font=theme.font(14),
                     text_color=theme.text_color()).pack(anchor="w", pady=(0, 8))

        themes = [("ระบบ (ตาม Windows)", "System"), ("สว่าง", "Light"), ("มืด", "Dark")]
        selected = ctk.StringVar(value=ctk.get_appearance_mode())
        for label, value in themes:
            rb = ctk.CTkRadioButton(
                body1, text=label, variable=selected, value=value,
                command=lambda v=value: self._apply_theme(v),
                font=theme.font(13), text_color=theme.text_color())
            rb.pack(anchor="w", pady=4)

        # ---- about ----
        card2 = theme.card(page)
        card2.pack(fill="x")
        theme.card_title(card2, "เกี่ยวกับระบบ").pack(anchor="w", padx=18, pady=(16, 10))

        body2 = ctk.CTkFrame(card2, fg_color="transparent")
        body2.pack(fill="x", padx=18, pady=(0, 18))

        rows = [
            ("ชื่อระบบ", "THE PY-LIB — ระบบห้องสมุดอัจฉริยะ"),
            ("เวอร์ชัน", "2.0 (โครงสร้างแบบแยกชั้น)"),
            ("ความปลอดภัย", "รหัสผ่าน bcrypt • QR AES-256-CBC"),
            ("รายงาน", "บัตรสมาชิก / ประวัติยืม-คืน / ประวัติเข้าออก (PDF)"),
            ("โรงเรียน", "DSNPRU"),
        ]
        for k, v in rows:
            r = ctk.CTkFrame(body2, fg_color="transparent")
            r.pack(fill="x", pady=2)
            ctk.CTkLabel(r, text=k, font=theme.font(13, bold=True),
                         text_color=theme.muted_color(), width=120, anchor="w").pack(side="left")
            ctk.CTkLabel(r, text=v, font=theme.font(13),
                         text_color=theme.text_color(), anchor="w").pack(
                side="left", fill="x", expand=True)

    def _apply_theme(self, theme_name):
        try:
            ctk.set_appearance_mode(theme_name)
            logger.info(f"Appearance theme changed to {theme_name}")
        except Exception as exc:
            logger.error(f"Failed to apply theme: {exc}")
