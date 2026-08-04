# ui/views/login_view.py
import customtkinter as ctk

import ui.theme as theme
from core.security import verify_password
from core.logger import logger


class LoginView(ctk.CTkFrame):
    def __init__(self, parent, db_conn, on_login_success):
        super().__init__(parent, fg_color="transparent")
        self.db_conn = db_conn
        self.on_login_success = on_login_success
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)

        # ---- left brand panel ----
        left = ctk.CTkFrame(self, fg_color="#0F172A", corner_radius=0, width=420)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        brand = ctk.CTkFrame(left, fg_color="transparent")
        brand.pack(pady=(70, 10), padx=44, anchor="w")
        ctk.CTkLabel(brand, text="THE PY-LIB", font=theme.font(34, bold=True),
                     text_color="#FFFFFF", anchor="w").pack(anchor="w")
        ctk.CTkLabel(brand, text="ระบบห้องสมุดอัจฉริยะ", font=theme.font(16),
                     text_color="#94A3B8", anchor="w").pack(anchor="w", pady=(6, 0))

        features = [
            ("สมาชิกและหนังสือ", "จัดการข้อมูลสมาชิกและหนังสือได้ในที่เดียว"),
            ("สแกน QR Code", "ยืม-คืนและบันทึกเข้าออกด้วยกล้อง"),
            ("รายงาน PDF", "ส่งออกบัตรสมาชิกและประวัติเป็น PDF"),
            ("ปลอดภัย AES-256", "ข้อมูลเข้ารหัส, รหัสผ่าน bcrypt"),
        ]
        for title, desc in features:
            row = ctk.CTkFrame(left, fg_color="transparent")
            row.pack(fill="x", padx=44, pady=10, anchor="w")
            ctk.CTkLabel(row, text="•", font=theme.font(18, bold=True),
                         text_color="#2563EB").pack(side="left", padx=(0, 12))
            col = ctk.CTkFrame(row, fg_color="transparent")
            col.pack(side="left", fill="x", expand=True)
            ctk.CTkLabel(col, text=title, font=theme.font(14, bold=True),
                         text_color="#E2E8F0", anchor="w").pack(anchor="w")
            ctk.CTkLabel(col, text=desc, font=theme.font(12),
                         text_color="#64748B", anchor="w").pack(anchor="w")

        # ---- right form panel ----
        right = ctk.CTkFrame(self, fg_color=theme.content_bg(), corner_radius=0)
        right.pack(side="left", fill="both", expand=True)

        card = ctk.CTkFrame(right, fg_color=theme.card_bg(), corner_radius=16,
                            border_width=1, border_color=theme.border_color())
        card.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.42, relheight=0.62)

        ctk.CTkLabel(card, text="เข้าสู่ระบบ", font=theme.font(26, bold=True),
                     text_color=theme.text_color()).pack(pady=(36, 4))
        ctk.CTkLabel(card, text="ผู้ดูแลระบบห้องสมุด", font=theme.font(13),
                     text_color=theme.muted_color()).pack(pady=(0, 22))

        form = ctk.CTkFrame(card, fg_color="transparent")
        form.pack(fill="x", padx=40)

        ctk.CTkLabel(form, text="ชื่อผู้ใช้", font=theme.font(13, bold=True),
                     text_color=theme.text_color(), anchor="w").pack(fill="x")
        self.username_entry = theme.entry(form, placeholder_text="กรอกชื่อผู้ใช้")
        self.username_entry.pack(fill="x", pady=(4, 12))

        ctk.CTkLabel(form, text="รหัสผ่าน", font=theme.font(13, bold=True),
                     text_color=theme.text_color(), anchor="w").pack(fill="x")
        self.password_entry = theme.entry(form, placeholder_text="กรอกรหัสผ่าน", show="*")
        self.password_entry.pack(fill="x", pady=(4, 4))

        self.error_label = ctk.CTkLabel(form, text="", font=theme.font(12),
                                        text_color=theme.DANGER, anchor="w")
        self.error_label.pack(fill="x", pady=(6, 0))

        theme.primary_button(card, "เข้าสู่ระบบ", self._do_login, height=46,
                             font=theme.font(15, bold=True)).pack(fill="x", padx=40, pady=(22, 14))

        ctk.CTkLabel(card, text="ติดต่อผู้ดูแลระบบหากลืมรหัสผ่าน",
                     font=theme.font(11), text_color=theme.muted_color()).pack(pady=(0, 28))

        self.username_entry.bind("<Return>", lambda e: self.password_entry.focus())
        self.password_entry.bind("<Return>", lambda e: self._do_login())
        self.username_entry.focus()

    def _do_login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get()
        if not username or not password:
            self.error_label.configure(text="กรุณากรอกชื่อผู้ใช้และรหัสผ่าน")
            return
        try:
            cursor = self.db_conn.cursor()
            cursor.execute("SELECT password FROM admin_users WHERE username = ?", (username,))
            row = cursor.fetchone()
            if row and verify_password(password, row[0]):
                logger.info(f"Login successful: {username}")
                self.on_login_success()
            else:
                logger.warning(f"Failed login attempt for: {username}")
                self.error_label.configure(text="ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")
        except Exception as e:
            logger.error(f"Login DB error: {e}")
            self.error_label.configure(text=f"เกิดข้อผิดพลาด: {str(e)}")
