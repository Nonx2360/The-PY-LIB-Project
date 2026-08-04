# ui/views/login_view.py
import customtkinter as ctk
from core.security import verify_password
from core.logger import logger

class LoginView(ctk.CTkFrame):
    def __init__(self, parent, db_conn, on_login_success):
        super().__init__(parent)
        self.db_conn = db_conn
        self.on_login_success = on_login_success
        self._build()

    def _build(self):
        self.pack(pady=20, padx=40, fill="both", expand=True)

        ctk.CTkLabel(self, text="ระบบห้องสมุดโรงเรียน",
                     font=("Helvetica", 24, "bold")).pack(pady=20)

        form = ctk.CTkFrame(self)
        form.pack(pady=20, padx=40, fill="x")

        ctk.CTkLabel(form, text="ชื่อผู้ใช้:", font=("Helvetica", 14)).pack(pady=(10, 0))
        self.username_entry = ctk.CTkEntry(form, placeholder_text="กรอกชื่อผู้ใช้", width=300)
        self.username_entry.pack(pady=5)

        ctk.CTkLabel(form, text="รหัสผ่าน:", font=("Helvetica", 14)).pack(pady=(10, 0))
        self.password_entry = ctk.CTkEntry(form, placeholder_text="กรอกรหัสผ่าน",
                                           show="*", width=300)
        self.password_entry.pack(pady=5)

        self.error_label = ctk.CTkLabel(form, text="", text_color="red")
        self.error_label.pack(pady=5)

        ctk.CTkButton(form, text="เข้าสู่ระบบ", width=200, height=40,
                      font=("Helvetica", 14, "bold"),
                      command=self._do_login).pack(pady=20)

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
