# ui/views/dashboard_view.py
import customtkinter as ctk

import ui.theme as theme
from db.repositories.base_repo import BaseRepository


class DashboardView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self.base = BaseRepository()
        self._resizable_labels = []
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)
        page = ctk.CTkFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True, padx=28, pady=24)

        theme.heading(page, "หน้าหลัก", size=26).pack(anchor="w")
        theme.subheading(page, "ภาพรวมระบบห้องสมุด").pack(anchor="w", pady=(2, 18))

        # ---- stat cards ----
        stats = [
            ("สมาชิกทั้งหมด", self._count("SELECT COUNT(*) FROM members"), theme.PRIMARY),
            ("หนังสือทั้งหมด", self._count("SELECT COUNT(*) FROM books"), theme.INFO),
            ("กำลังยืม", self._count("SELECT COUNT(*) FROM borrow_log WHERE returned = 0"), theme.WARNING),
            ("ค้างส่ง", self._count("SELECT COUNT(*) FROM borrow_log WHERE returned = 0 AND return_due < date('now')"), theme.DANGER),
        ]
        stat_row = ctk.CTkFrame(page, fg_color="transparent")
        stat_row.pack(fill="x", pady=(0, 22))
        for i, (label, value, color) in enumerate(stats):
            stat_row.grid_columnconfigure(i, weight=1)
            card = theme.card(stat_row)
            card.grid(row=0, column=i, padx=6, sticky="nsew")
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="x", padx=18, pady=16)
            stripe = ctk.CTkFrame(inner, width=6, height=44, fg_color=color, corner_radius=3)
            stripe.pack(side="left", padx=(0, 14))
            stripe.pack_propagate(False)
            col = ctk.CTkFrame(inner, fg_color="transparent")
            col.pack(side="left", fill="x", expand=True)
            num = ctk.CTkLabel(col, text=str(value), font=theme.font(28, bold=True),
                               text_color=color, anchor="w")
            num.pack(anchor="w")
            self._resizable_labels.append(num)
            ctk.CTkLabel(col, text=label, font=theme.font(13),
                         text_color=theme.muted_color(), anchor="w").pack(anchor="w")

        # ---- action tiles ----
        tiles = [
            ("จัดการสมาชิก", "เพิ่ม ค้นหา และพิมพ์บัตร", theme.PRIMARY, self.nav.show_members),
            ("จัดการหนังสือ", "เพิ่มหนังสือ หรือนำเข้า Excel", theme.INFO, self.nav.show_books),
            ("ยืมหนังสือ", "สแกนบัตรและยืมหนังสือ", theme.SUCCESS, self.nav.show_borrow),
            ("คืนหนังสือ", "สแกนบัตรและรับคืน", theme.WARNING, self.nav.show_return),
            ("ประวัติยืม-คืน", "ค้นหาประวัติและส่งออก PDF", theme.PRIMARY, self.nav.show_history),
            ("บันทึกเข้าออก", "สแกนบัตรเมื่อเข้าออกห้องสมุด", theme.INFO, self.nav.show_access_scanner),
            ("ประวัติเข้าออก", "ดูและส่งออกประวัติ", theme.SUCCESS, self.nav.show_access_history),
            ("ตั้งค่า", "ธีมการแสดงผลและข้อมูลระบบ", theme.WARNING, self.nav.show_settings),
        ]
        grid = ctk.CTkFrame(page, fg_color="transparent")
        grid.pack(fill="x")
        for i in range(4):
            grid.grid_columnconfigure(i, weight=1)

        for i, (title, desc, color, cmd) in enumerate(tiles):
            row, col = divmod(i, 4)
            tile = theme.card(grid)
            tile.grid(row=row, column=col, padx=6, pady=6, sticky="nsew")
            tile.bind("<Button-1>", lambda e, c=cmd: c())
            inner = ctk.CTkFrame(tile, fg_color="transparent")
            inner.pack(fill="x", padx=16, pady=14)
            dot = ctk.CTkFrame(inner, width=10, height=10, fg_color=color, corner_radius=5)
            dot.pack(side="left", padx=(0, 12))
            dot.pack_propagate(False)
            col_f = ctk.CTkFrame(inner, fg_color="transparent")
            col_f.pack(side="left", fill="x", expand=True)
            title_lbl = ctk.CTkLabel(col_f, text=title, font=theme.font(16, bold=True),
                                     text_color=theme.text_color(), anchor="w")
            title_lbl.pack(anchor="w")
            self._resizable_labels.append(title_lbl)
            ctk.CTkLabel(col_f, text=desc, font=theme.font(12),
                         text_color=theme.muted_color(), anchor="w").pack(anchor="w")
            ctk.CTkLabel(inner, text="→", font=theme.font(16, bold=True),
                         text_color=color).pack(side="right")

        self.winfo_toplevel().bind("<Configure>", self._on_resize)

    def _count(self, query):
        try:
            row = self.base.fetchone(query)
            return row[0] if row else 0
        except Exception:
            return 0

    def _on_resize(self, event):
        try:
            size = 15 if event.width < 1200 else 17
            for lbl in self._resizable_labels:
                lbl.configure(font=theme.font(size, bold=True))
        except Exception:
            pass
