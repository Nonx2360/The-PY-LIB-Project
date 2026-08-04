# ui/views/settings_view.py
import os
import shutil
import customtkinter as ctk
from datetime import datetime

import ui.theme as theme
from config import settings
from core.logger import logger
from ui.widgets.toast import show_toast


class SettingsView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self._vars = {}
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)
        page = ctk.CTkScrollableFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True, padx=28, pady=24)

        theme.heading(page, "ตั้งค่า", size=26).pack(anchor="w")
        theme.subheading(page, "ปรับแต่งระบบและข้อมูลทั่วไป").pack(anchor="w", pady=(2, 18))

        # ---- 1. appearance ----
        self._build_appearance(page)

        # ---- 2. ยืม-คืน ----
        self._build_loan_settings(page)

        # ---- 3. ข้อมูลโรงเรียน ----
        self._build_school(page)

        # ---- 4. ฐานข้อมูล ----
        self._build_database(page)

        # ---- 5. เกี่ยวกับ ----
        self._build_about(page)

        # ---- save button ----
        save_frame = ctk.CTkFrame(page, fg_color="transparent")
        save_frame.pack(fill="x", pady=(18, 0))
        theme.primary_button(save_frame, "บันทึกการตั้งค่าทั้งหมด",
                             self._save_all, height=44,
                             font=theme.font(15, bold=True)).pack(side="right")

    # ================================================================
    # 1. appearance
    # ================================================================
    def _build_appearance(self, parent):
        card = theme.card(parent)
        card.pack(fill="x", pady=(0, 14))
        theme.card_title(card, "การแสดงผล").pack(anchor="w", padx=18, pady=(16, 10))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=18, pady=(0, 18))

        ctk.CTkLabel(body, text="ธีม:", font=theme.font(14, bold=True),
                     text_color=theme.text_color()).pack(anchor="w", pady=(0, 6))
        self._vars["theme"] = ctk.StringVar(value=settings.get("theme"))
        for label, value in [("ระบบ (ตาม Windows)", "System"), ("สว่าง", "Light"), ("มืด", "Dark")]:
            ctk.CTkRadioButton(
                body, text=label, variable=self._vars["theme"], value=value,
                command=lambda v=value: self._apply_theme(v),
                font=theme.font(13), text_color=theme.text_color()).pack(anchor="w", pady=3)

    # ================================================================
    # 2. loan settings
    # ================================================================
    def _build_loan_settings(self, parent):
        card = theme.card(parent)
        card.pack(fill="x", pady=(0, 14))
        theme.card_title(card, "การยืม-คืน").pack(anchor="w", padx=18, pady=(16, 10))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=18, pady=(0, 18))
        body.grid_columnconfigure((1, 3), weight=1)

        fields = [
            ("ระยะเวลาการยืม (วัน):", "default_loan_days", 7, 1, 0),
            ("จำนวนหนังสือสูงสุดต่อคน:", "max_books_per_member", 3, 3, 0),
            ("อายุบัตรสมาชิก (วัน):", "member_expiry_days", 365, 1, 2),
        ]
        for label_text, key, default, col, row in fields:
            ctk.CTkLabel(body, text=label_text, font=theme.font(14, bold=True),
                         text_color=theme.text_color(), anchor="w").grid(
                row=row, column=col * 2, sticky="w", padx=(0, 8), pady=(10, 2))
            var = ctk.StringVar(value=str(settings.get(key)))
            self._vars[key] = var
            theme.entry(body, placeholder_text=str(default), width=120,
                        textvariable=var).grid(
                row=row + 1, column=col * 2, sticky="w", pady=(0, 8))

        hints = [
            "กำหนดจำนวนวันเริ่มต้นเมื่อยืมหนังสือ (เช่น 7, 14, 21)",
            "สมาชิกยืมได้ไม่เกินจำนวนนี้พร้อมกัน",
            "กำหนดวันหมดอายุของบัตรสมาชิก (เช่น 365 = 1 ปี)",
        ]
        for i, hint in enumerate(hints):
            ctk.CTkLabel(body, text=hint, font=theme.font(11),
                         text_color=theme.muted_color(), anchor="w").grid(
                row=i * 2 + 1, column=1 if i < 2 else 5, columnspan=3,
                sticky="w", padx=(4, 0), pady=(0, 8))

    # ================================================================
    # 3. school info
    # ================================================================
    def _build_school(self, parent):
        card = theme.card(parent)
        card.pack(fill="x", pady=(0, 14))
        theme.card_title(card, "ข้อมูลโรงเรียน").pack(anchor="w", padx=18, pady=(16, 10))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=18, pady=(0, 18))

        ctk.CTkLabel(body, text="ชื่อโรงเรียน:", font=theme.font(14, bold=True),
                     text_color=theme.text_color(), anchor="w").pack(anchor="w", pady=(0, 4))
        self._vars["school_name"] = ctk.StringVar(value=settings.get("school_name"))
        theme.entry(body, placeholder_text="DSNPRU", width=400,
                    textvariable=self._vars["school_name"]).pack(anchor="w", pady=(0, 10))

        ctk.CTkLabel(body, text="ที่อยู่ไฟล์โลโก้:", font=theme.font(14, bold=True),
                     text_color=theme.text_color(), anchor="w").pack(anchor="w", pady=(0, 4))
        logo_frame = ctk.CTkFrame(body, fg_color="transparent")
        logo_frame.pack(fill="x", pady=(0, 10))
        self._vars["school_logo_path"] = ctk.StringVar(value=settings.get("school_logo_path"))
        theme.entry(logo_frame, placeholder_text="assets/logos/school_logo.png", width=340,
                    textvariable=self._vars["school_logo_path"]).pack(side="left")
        theme.secondary_button(logo_frame, "เลือกไฟล์", self._pick_logo, width=90).pack(
            side="left", padx=(8, 0))

        logo_path = settings.get("school_logo_path")
        if logo_path and os.path.exists(logo_path):
            ctk.CTkLabel(body, text=f"ไฟล์ปัจจุบัน: {logo_path}",
                         font=theme.font(12), text_color=theme.SUCCESS).pack(anchor="w")
        else:
            ctk.CTkLabel(body, text="ไม่พบไฟล์โลโก้ — บัตรสมาชิกจะไม่มีรูป",
                         font=theme.font(12), text_color=theme.WARNING).pack(anchor="w")

    def _pick_logo(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            filetypes=[("PNG files", "*.png"), ("All files", "*.*")],
            title="เลือกไฟล์โลโก้โรงเรียน")
        if path:
            self._vars["school_logo_path"].set(path)

    # ================================================================
    # 4. database
    # ================================================================
    def _build_database(self, parent):
        card = theme.card(parent)
        card.pack(fill="x", pady=(0, 14))
        theme.card_title(card, "ฐานข้อมูล").pack(anchor="w", padx=18, pady=(16, 10))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=18, pady=(0, 18))

        db_path = os.path.abspath("db/library.db")
        size_kb = os.path.getsize(db_path) / 1024 if os.path.exists(db_path) else 0
        ctk.CTkLabel(body, text=f"ไฟล์ฐานข้อมูล: {db_path}",
                     font=theme.font(12), text_color=theme.muted_color()).pack(anchor="w")
        ctk.CTkLabel(body, text=f"ขนาด: {size_kb:.1f} KB",
                     font=theme.font(12), text_color=theme.muted_color()).pack(anchor="w", pady=(0, 10))

        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(fill="x")
        theme.secondary_button(btn_row, "สำรองฐานข้อมูล (Backup)", self._backup_db,
                               width=220).pack(side="left", padx=(0, 8))

    def _backup_db(self):
        from tkinter import filedialog
        db_path = "db/library.db"
        if not os.path.exists(db_path):
            show_toast(self, "ไม่พบไฟล์ฐานข้อมูล", "red")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("SQLite DB", "*.db")],
            initialfile=f"library_backup_{ts}.db")
        if dest:
            try:
                shutil.copy2(db_path, dest)
                show_toast(self, f"สำรองสำเร็จ: {os.path.basename(dest)}", "green")
                logger.info(f"Database backed up to {dest}")
            except Exception as e:
                logger.error(f"Backup failed: {e}", exc_info=True)
                show_toast(self, f"สำรองไม่สำเร็จ: {e}", "red")

    # ================================================================
    # 5. about
    # ================================================================
    def _build_about(self, parent):
        card = theme.card(parent)
        card.pack(fill="x")
        theme.card_title(card, "เกี่ยวกับระบบ").pack(anchor="w", padx=18, pady=(16, 10))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=18, pady=(0, 18))

        rows = [
            ("ชื่อระบบ", "THE PY-LIB — ระบบห้องสมุดอัจฉริยะ"),
            ("เวอร์ชัน", "2.0 (โมดูลาร์ + UI ใหม่)"),
            ("ความปลอดภัย", "รหัสผ่าน bcrypt • QR AES-256-CBC • PBKDF2"),
            ("รายงาน", "บัตรสมาชิก / ประวัติยืม-คืน / ประวัติเข้าออก (PDF)"),
            ("โครงสร้าง", "config / core / db / services / ui (แยกชั้น)"),
            ("โรงเรียน", settings.get("school_name")),
        ]
        for k, v in rows:
            r = ctk.CTkFrame(body, fg_color="transparent")
            r.pack(fill="x", pady=2)
            ctk.CTkLabel(r, text=k, font=theme.font(13, bold=True),
                         text_color=theme.muted_color(), width=110, anchor="w").pack(side="left")
            ctk.CTkLabel(r, text=v, font=theme.font(13),
                         text_color=theme.text_color(), anchor="w").pack(
                side="left", fill="x", expand=True)

    # ================================================================
    # actions
    # ================================================================
    def _apply_theme(self, theme_name):
        try:
            ctk.set_appearance_mode(theme_name)
            self._vars["theme"].set(theme_name)
            logger.info(f"Theme changed to {theme_name}")
        except Exception as exc:
            logger.error(f"Failed to apply theme: {exc}")

    def _save_all(self):
        try:
            settings.set("theme", self._vars["theme"].get())
            settings.set("school_name", self._vars["school_name"].get().strip())
            settings.set("school_logo_path", self._vars["school_logo_path"].get().strip())

            for key in ("default_loan_days", "max_books_per_member", "member_expiry_days"):
                raw = self._vars[key].get().strip()
                try:
                    settings.set(key, int(raw))
                except ValueError:
                    pass

            settings.save()
            show_toast(self, "บันทึกการตั้งค่าสำเร็จ", "green")
            logger.info(f"Settings saved: {settings.all()}")
        except Exception as e:
            logger.error(f"Save settings error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")
