# ui/views/member_view.py
import os
import uuid
import customtkinter as ctk
from datetime import datetime, timedelta
from tkinter import filedialog

from PIL import Image

import ui.theme as theme
from config import settings
from db.repositories.member_repo import MemberRepository
from services.qr_service import QRService
from services.pdf_service import PDFService
from ui.widgets.confirm_dialog import show_confirm_dialog
from ui.widgets.toast import show_toast
from ui.widgets.pdf_viewer import PDFViewerWindow
from core.logger import logger


class MemberView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self.repo = MemberRepository()
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)
        page = ctk.CTkFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True, padx=28, pady=24)

        theme.heading(page, "จัดการสมาชิก", size=26).pack(anchor="w")
        theme.subheading(page, "เพิ่มสมาชิก พิมพ์บัตร และดู QR Code").pack(anchor="w", pady=(2, 18))

        body = ctk.CTkFrame(page, fg_color="transparent")
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(1, weight=1)

        # ---- left: add form ----
        form_card = theme.card(body)
        form_card.grid(row=0, column=0, padx=(0, 14), sticky="nsw")
        theme.card_title(form_card, "เพิ่มสมาชิกใหม่").pack(anchor="w", padx=18, pady=(16, 14))
        form = ctk.CTkFrame(form_card, fg_color="transparent")
        form.pack(fill="x", padx=18, pady=(0, 16))

        ctk.CTkLabel(form, text="ชื่อ-นามสกุล", font=theme.font(12, bold=True),
                     text_color=theme.text_color(), anchor="w").pack(fill="x")
        self.name_entry = theme.entry(form, placeholder_text="เช่น สมชาย ใจดี", width=300)
        self.name_entry.pack(fill="x", pady=(4, 10))

        row2 = ctk.CTkFrame(form, fg_color="transparent")
        row2.pack(fill="x", pady=(0, 10))
        row2.grid_columnconfigure(0, weight=1)
        row2.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row2, text="ชั้น", font=theme.font(12, bold=True),
                     text_color=theme.text_color(), anchor="w").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(row2, text="เลขที่", font=theme.font(12, bold=True),
                     text_color=theme.text_color(), anchor="w").grid(row=0, column=1, sticky="w", padx=(12, 0))
        self.grade_entry = theme.entry(row2, placeholder_text="เช่น ม.1/1")
        self.grade_entry.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.number_entry = theme.entry(row2, placeholder_text="เช่น 15")
        self.number_entry.grid(row=1, column=1, sticky="ew", pady=(4, 0), padx=(12, 0))

        theme.primary_button(form, "เพิ่มสมาชิก + สร้าง QR", self._add_member, height=42).pack(
            fill="x", pady=(8, 0))
        theme.subheading(form, "ระบบจะสร้าง QR Code และบัตรสมาชิก PDF ให้อัตโนมัติ",
                         size=11).pack(anchor="w", pady=(8, 0))

        # ---- right: member list ----
        list_card = theme.card(body)
        list_card.grid(row=0, column=1, sticky="nsew")

        top = ctk.CTkFrame(list_card, fg_color="transparent")
        top.pack(fill="x", padx=18, pady=(14, 10))
        self.count_label = ctk.CTkLabel(top, text="รายชื่อสมาชิก", font=theme.font(17, bold=True),
                                        text_color=theme.text_color(), anchor="w")
        self.count_label.pack(side="left")
        self.search_entry = theme.entry(top, placeholder_text="ค้นหา ชื่อ / ชั้น / เลขที่",
                                        width=240)
        self.search_entry.pack(side="right")
        self.search_entry.bind("<KeyRelease>", lambda e: self._display_members())

        self.list_frame = ctk.CTkScrollableFrame(list_card, fg_color="transparent",
                                                 corner_radius=0)
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._display_members()

    # ================= data =================
    def _display_members(self):
        for w in self.list_frame.winfo_children():
            w.destroy()

        members = self.repo.get_all()
        q = self.search_entry.get().strip().lower()
        if q:
            members = [m for m in members if q in (m[1] or "").lower()
                       or q in (m[2] or "").lower() or q in (m[3] or "").lower()]

        self.count_label.configure(text=f"รายชื่อสมาชิก ({len(members)})")

        if not members:
            ctk.CTkLabel(self.list_frame, text="ไม่พบสมาชิก", font=theme.font(14),
                         text_color=theme.muted_color()).pack(pady=24)
            return

        for m in members:
            self._member_row(m)

    def _member_row(self, m):
        mid, name, grade, number, reg, exp, qr_data, qr_path = m
        row = theme.card(self.list_frame, corner_radius=10)
        row.pack(fill="x", padx=4, pady=4)
        row.grid_columnconfigure(1, weight=1)

        ctk.CTkFrame(row, width=4, height=40, fg_color=theme.PRIMARY,
                     corner_radius=2).grid(row=0, column=0, rowspan=2, padx=(10, 12), pady=8)
        ctk.CTkLabel(row, text=name, font=theme.font(15, bold=True),
                     text_color=theme.text_color(), anchor="w").grid(
            row=0, column=1, sticky="w", pady=(8, 0))
        info = f"ชั้น {grade}  •  เลขที่ {number}"
        ctk.CTkLabel(row, text=info, font=theme.font(12),
                     text_color=theme.muted_color(), anchor="w").grid(
            row=1, column=1, sticky="w", pady=(0, 8))

        expire_lbl = theme.badge_amber(row, f"หมดอายุ {exp}") if exp >= datetime.now().strftime("%Y-%m-%d") \
            else theme.badge_red(row, f"หมดอายุ {exp}")
        expire_lbl.grid(row=0, column=2, rowspan=2, padx=6)

        btn_row = ctk.CTkFrame(row, fg_color="transparent")
        btn_row.grid(row=0, column=3, rowspan=2, padx=(6, 10))
        theme.secondary_button(btn_row, "ดู QR", lambda mm=m: self._view_qr(mm),
                               width=72, height=30).pack(side="left", padx=2)
        theme.secondary_button(btn_row, "บัตร PDF", lambda mm=m: self._view_card(mm),
                               width=80, height=30).pack(side="left", padx=2)
        theme.ghost_button(btn_row, "ลบ", lambda mm=m: self._delete_member(mm),
                           width=56, height=30, text_color=theme.DANGER).pack(side="left", padx=2)

    # ================= actions =================
    def _add_member(self):
        name = self.name_entry.get().strip()
        grade = self.grade_entry.get().strip()
        number = self.number_entry.get().strip()
        if not all([name, grade, number]):
            show_toast(self, "กรุณากรอกข้อมูลให้ครบ", "red")
            return

        reg_date = datetime.now().strftime("%Y-%m-%d")
        exp_date = (datetime.now() + timedelta(days=settings.get("member_expiry_days"))).strftime("%Y-%m-%d")
        qr_data = f"{name}|{grade}|{number}|{reg_date}|{exp_date}"

        try:
            os.makedirs("assets/qrcodes", exist_ok=True)
            os.makedirs("assets/cards", exist_ok=True)
            qr_filename = f"assets/qrcodes/{uuid.uuid4()}.png"
            encrypted = QRService.generate_qr(qr_data, qr_filename)

            member_id = self.repo.add(name, grade, number, reg_date, exp_date,
                                      encrypted, qr_filename)
            card_path = f"assets/cards/member_{member_id}.pdf"
            PDFService.generate_member_card_pdf(name, number,
                                                "assets/logos/school_logo.png",
                                                card_path, encrypted)
            pil_img = Image.open(qr_filename)
            self._show_qr_window(pil_img, name, grade, number, reg_date, exp_date)
            show_toast(self, f"เพิ่มสมาชิกสำเร็จ: {name}", "green")
            self.name_entry.delete(0, "end")
            self.grade_entry.delete(0, "end")
            self.number_entry.delete(0, "end")
            self._display_members()
        except Exception as e:
            logger.error(f"add_member error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")

    def _show_qr_window(self, pil_image, name, grade, number, reg, exp):
        win = ctk.CTkToplevel(self)
        win.title("QR Code บัตรสมาชิก")
        win.geometry("420x560")
        win.resizable(False, False)
        win.grab_set()
        frame = theme.card(win)
        frame.pack(padx=16, pady=16, fill="both", expand=True)

        ctk.CTkLabel(frame, text="QR Code สมาชิก", font=theme.font(19, bold=True),
                     text_color=theme.text_color()).pack(pady=(16, 6))
        ctk.CTkLabel(frame, text=f"ชื่อ: {name}   ชั้น: {grade}   เลขที่: {number}\n"
                     f"สมัคร: {reg}   หมดอายุ: {exp}",
                     font=theme.font(13), text_color=theme.muted_color(),
                     justify="center").pack(pady=(0, 12))

        ctk_img = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(240, 240))
        lbl = ctk.CTkLabel(frame, image=ctk_img, text="")
        lbl.image = ctk_img
        lbl.pack(pady=6)

        theme.primary_button(frame, "บันทึก QR Code",
                             lambda: self._save_qr(pil_image, name), width=200).pack(pady=8)
        theme.secondary_button(frame, "ปิด", win.destroy, width=120).pack(pady=(0, 16))

    def _save_qr(self, pil_image, name):
        path = filedialog.asksaveasfilename(defaultextension=".png",
                                            filetypes=[("PNG files", "*.png")],
                                            initialfile=f"qrcode_{name}.png")
        if path:
            pil_image.save(path)
            show_toast(self, f"บันทึก QR Code แล้ว: {path}", "green")

    def _view_qr(self, member):
        try:
            if member[7] and os.path.exists(member[7]):
                pil_img = Image.open(member[7])
            else:
                qr_data = member[6] or f"{member[1]}|{member[2]}|{member[3]}|{member[4]}|{member[5]}"
                qr_path = f"assets/qrcodes/{uuid.uuid4()}.png"
                QRService.generate_qr(qr_data, qr_path)
                pil_img = Image.open(qr_path)
            self._show_qr_window(pil_img, member[1], member[2], member[3],
                                 member[4], member[5])
        except Exception as e:
            logger.error(f"view_qr error: {e}", exc_info=True)
            show_toast(self, f"ไม่สามารถเปิด QR Code ได้: {e}", "red")

    def _view_card(self, member):
        try:
            os.makedirs("assets/cards", exist_ok=True)
            card_path = os.path.abspath(f"assets/cards/member_{member[0]}.pdf")
            if not os.path.exists(card_path):
                PDFService.generate_member_card_pdf(
                    member[1], member[3], "assets/logos/school_logo.png",
                    card_path, member[6])
            PDFViewerWindow(self, card_path, title=f"บัตรสมาชิก — {member[1]}")
        except Exception as e:
            logger.error(f"view_card error: {e}", exc_info=True)
            show_toast(self, f"ไม่สามารถเปิดบัตรได้: {e}", "red")

    def _delete_member(self, member):
        if not show_confirm_dialog(self, "ยืนยันการลบ",
                                   f"ลบสมาชิก '{member[1]}' ใช่หรือไม่?\nการกระทำนี้ไม่สามารถย้อนกลับได้"):
            return
        try:
            if member[7] and os.path.exists(member[7]):
                os.remove(member[7])
            self.repo.delete(member[0])
            show_toast(self, "ลบสมาชิกสำเร็จ", "green")
            self._display_members()
        except Exception as e:
            logger.error(f"delete_member error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")
