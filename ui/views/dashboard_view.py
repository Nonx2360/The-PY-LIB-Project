# ui/views/dashboard_view.py
import customtkinter as ctk

class DashboardView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True, padx=20, pady=20)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", pady=(0, 20))
        ctk.CTkLabel(header, text="ระบบห้องสมุดโรงเรียน",
                     font=("Sarabun-Bold", 32), text_color="#1f538d").pack(pady=10)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="both", expand=True)
        btn_frame.grid_columnconfigure((0, 1, 2), weight=1)
        btn_frame.grid_rowconfigure((0, 1, 2), weight=1)

        buttons = [
            ("จัดการสมาชิก",       self.nav.show_members,        0, 0),
            ("จัดการหนังสือ",      self.nav.show_books,          0, 1),
            ("ยืมหนังสือ",          self.nav.show_borrow,         0, 2),
            ("คืนหนังสือ",          self.nav.show_return,         1, 0),
            ("ประวัติการยืม-คืน",  self.nav.show_history,        1, 1),
            ("บันทึกเข้าออก",      self.nav.show_access_scanner, 1, 2),
            ("ประวัติเข้าออก",     self.nav.show_access_history, 2, 0),
            ("ตั้งค่า",             self.nav.show_settings,       2, 1),
            ("เกี่ยวกับ",           self.nav.show_about,          2, 2),
        ]

        for text, cmd, row, col in buttons:
            btn = ctk.CTkButton(
                btn_frame, text=text, command=cmd,
                height=80, font=("Sarabun", 20, "bold"),
                fg_color="#1f538d", hover_color="#14375e",
                corner_radius=15, border_width=2, border_color="#14375e"
            )
            btn.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")

        self._btn_frame = btn_frame

        # Responsive font resize
        self.winfo_toplevel().bind("<Configure>", self._on_resize)

    def _on_resize(self, event):
        w = event.width
        size = 16 if w < 800 else (20 if w < 1200 else 24)
        for child in self._btn_frame.winfo_children():
            if isinstance(child, ctk.CTkButton):
                child.configure(font=("Sarabun", size, "bold"))
