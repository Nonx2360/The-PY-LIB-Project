# ui/views/member_view.py
import os, uuid, io
import customtkinter as ctk
from tkinter import filedialog
from PIL import Image
from datetime import datetime, timedelta

from db.repositories.member_repo import MemberRepository
from services.qr_service import QRService
from services.pdf_service import PDFService
from ui.widgets.confirm_dialog import show_confirm_dialog
from core.logger import logger


class MemberView(ctk.CTkFrame):
    def __init__(self, parent, nav):
        super().__init__(parent)
        self.nav = nav
        self.repo = MemberRepository()
        self._build()

    def _build(self):
        self.pack(pady=20, padx=40, fill="both", expand=True)
        ctk.CTkLabel(self, text="จัดการสมาชิก", font=("Helvetica", 24)).pack(pady=20)

        form = ctk.CTkFrame(self)
        form.pack(pady=10, padx=20, fill="x")

        self.name_entry = ctk.CTkEntry(form, placeholder_text="ชื่อ-นามสกุล")
        self.name_entry.pack(pady=5, padx=20, fill="x")

        self.grade_entry = ctk.CTkEntry(form, placeholder_text="ชั้น")
        self.grade_entry.pack(pady=5, padx=20, fill="x")

        self.number_entry = ctk.CTkEntry(form, placeholder_text="เลขที่")
        self.number_entry.pack(pady=5, padx=20, fill="x")

        ctk.CTkButton(form, text="เพิ่มสมาชิก",
                      command=self._add_member).pack(pady=10)

        ctk.CTkButton(self, text="กลับ", command=self.nav.show_dashboard).pack(pady=10)

        self._display_members()

    def _show_toast(self, msg, color="green"):
        lbl = ctk.CTkLabel(self.winfo_toplevel(), text=msg, text_color=color)
        lbl.pack(pady=5)
        self.after(2500, lbl.destroy)

    def _add_member(self):
        name = self.name_entry.get().strip()
        grade = self.grade_entry.get().strip()
        number = self.number_entry.get().strip()
        if not all([name, grade, number]):
            self._show_toast("กรุณากรอกข้อมูลให้ครบ", "red")
            return

        reg_date = datetime.now().strftime("%Y-%m-%d")
        exp_date = (datetime.now() + timedelta(days=365)).strftime("%Y-%m-%d")
        qr_data = f"{name}|{grade}|{number}|{reg_date}|{exp_date}"

        try:
            qr_filename = f"assets/qrcodes/{uuid.uuid4()}.png"
            encrypted = QRService.generate_qr(qr_data, qr_filename)

            member_id = self.repo.add(name, grade, number, reg_date, exp_date, encrypted, qr_filename)

            card_path = f"assets/cards/member_{member_id}.pdf"
            PDFService.generate_member_card_pdf(name, number, "assets/logos/school_logo.png",
                                                card_path, encrypted)
            pil_img = Image.open(qr_filename)
            self._show_qr_window(pil_img, name, grade, number, reg_date, exp_date)
            self._show_toast("เพิ่มสมาชิกสำเร็จ")
            self.nav.show_members()
        except Exception as e:
            logger.error(f"add_member error: {e}", exc_info=True)
            self._show_toast(f"เกิดข้อผิดพลาด: {e}", "red")

    def _show_qr_window(self, pil_image, name, grade, number, reg, exp):
        win = ctk.CTkToplevel(self)
        win.title("QR Code บัตรสมาชิก")
        win.geometry("400x600")
        frame = ctk.CTkFrame(win)
        frame.pack(pady=20, padx=20, fill="both", expand=True)

        ctk.CTkLabel(frame, text=f"ชื่อ: {name}\nชั้น: {grade}\nเลขที่: {number}\n"
                     f"สมัคร: {reg}\nหมดอายุ: {exp}", font=("Helvetica", 14)).pack(pady=10)

        ctk_img = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(256, 256))
        lbl = ctk.CTkLabel(frame, image=ctk_img, text="")
        lbl.image = ctk_img
        lbl.pack(pady=10)

        ctk.CTkButton(frame, text="บันทึก QR Code",
                      command=lambda: self._save_qr(pil_image, name)).pack(pady=5)
        ctk.CTkButton(frame, text="ปิด", command=win.destroy).pack(pady=5)

    def _save_qr(self, pil_image, name):
        path = filedialog.asksaveasfilename(defaultextension=".png",
                                            filetypes=[("PNG files", "*.png")],
                                            initialfile=f"qrcode_{name}.png")
        if path:
            pil_image.save(path)
            self._show_toast(f"บันทึก QR Code แล้ว: {path}")

    def _display_members(self):
        scroll = ctk.CTkScrollableFrame(self)
        scroll.pack(pady=10, padx=20, fill="both", expand=True)

        members = self.repo.get_all()
        for m in members:
            row = ctk.CTkFrame(scroll)
            row.pack(pady=5, padx=10, fill="x")

            info = (f"ชื่อ: {m[1]}  ชั้น: {m[2]}  เลขที่: {m[3]}\n"
                    f"สมัคร: {m[4]}  หมดอายุ: {m[5]}")
            ctk.CTkLabel(row, text=info, justify="left").pack(side="left", padx=10)

            btn_row = ctk.CTkFrame(row)
            btn_row.pack(side="right", padx=5)

            ctk.CTkButton(btn_row, text="ดู QR",
                          command=lambda mm=m: self._view_qr(mm)).pack(side="left", padx=2)
            ctk.CTkButton(btn_row, text="ดูบัตร",
                          command=lambda mm=m: self._view_card(mm)).pack(side="left", padx=2)
            ctk.CTkButton(btn_row, text="ลบ",
                          command=lambda mm=m: self._delete_member(mm)).pack(side="left", padx=2)

    def _view_qr(self, member):
        try:
            pil_img = Image.open(member[7])
            self._show_qr_window(pil_img, member[1], member[2], member[3], member[4], member[5])
        except Exception as e:
            self._show_toast(f"ไม่สามารถเปิด QR Code ได้: {e}", "red")

    def _view_card(self, member):
        card_path = os.path.abspath(f"assets/cards/member_{member[0]}.pdf")
        if not os.path.exists(card_path):
            try:
                PDFService.generate_member_card_pdf(
                    member[1], member[3], "assets/logos/school_logo.png",
                    card_path, member[6])
            except Exception as e:
                self._show_toast(f"ไม่สามารถสร้างบัตร: {e}", "red")
                return
        if os.name == "nt":
            os.startfile(card_path)
        else:
            import subprocess
            subprocess.run(["xdg-open", card_path])

    def _delete_member(self, member):
        if not show_confirm_dialog(self, "ยืนยันการลบ",
                                   f"ลบสมาชิก '{member[1]}' ใช่หรือไม่?"):
            return
        try:
            if member[7] and os.path.exists(member[7]):
                os.remove(member[7])
            self.repo.delete(member[0])
            self._show_toast("ลบสมาชิกสำเร็จ")
            self.nav.show_members()
        except Exception as e:
            logger.error(f"delete_member error: {e}", exc_info=True)
            self._show_toast(f"เกิดข้อผิดพลาด: {e}", "red")
