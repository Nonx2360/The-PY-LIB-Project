# ui/views/access_view.py
import customtkinter as ctk
from datetime import datetime

import ui.theme as theme
from db.repositories.access_repo import AccessRepository
from db.repositories.member_repo import MemberRepository
from ui.widgets.scan_window import ScanWindow
from ui.widgets.toast import show_toast
from core.logger import logger


class AccessScannerView(ctk.CTkFrame):
    """Records entry/exit of library members via QR scan."""

    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self.access_repo = AccessRepository()
        self.member_repo = MemberRepository()
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)
        page = ctk.CTkFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True, padx=28, pady=24)

        theme.heading(page, "บันทึกการเข้าออก", size=26).pack(anchor="w")
        theme.subheading(page, "สแกนบัตรสมาชิก เมื่อเข้าหรือออกห้องสมุด").pack(
            anchor="w", pady=(2, 24))

        center = ctk.CTkFrame(page, fg_color="transparent")
        center.place(relx=0.5, rely=0.45, anchor="center")

        card = theme.card(center, corner_radius=18)
        card.pack(padx=40, pady=20)

        ctk.CTkLabel(card, text="สแกนบัตรสมาชิก", font=theme.font(20, bold=True),
                     text_color=theme.text_color()).pack(pady=(26, 6))

        self.status_label = ctk.CTkLabel(card, text="พร้อมใช้งาน",
                                         font=theme.font(15),
                                         text_color=theme.muted_color())
        self.status_label.pack(pady=(0, 14))

        theme.primary_button(card, "เปิดกล้องสแกน", self._open_scan_window,
                             height=56, width=280,
                             font=theme.font(16, bold=True)).pack(pady=(0, 10))

        self.last_scan_label = ctk.CTkLabel(card, text="ยังไม่มีการสแกนในรอบนี้",
                                            font=theme.font(13),
                                            text_color=theme.muted_color(),
                                            justify="center")
        self.last_scan_label.pack(pady=(6, 22))

        # ================= scan =================
    def _open_scan_window(self):
        ScanWindow(self, self._handle_member_qr,
                   on_error=lambda msg: show_toast(self, msg, "red"))

    def _handle_member_qr(self, decoded):
        try:
            name, grade, number, *_ = decoded.split("|")
            row = self.member_repo.fetchone(
                "SELECT id FROM members WHERE name=? AND grade=? AND number=?",
                (name, grade, number))
            if row:
                mid = row[0]
                last = self.access_repo.fetchone(
                    "SELECT action FROM access_log WHERE member_id=? ORDER BY id DESC LIMIT 1",
                    (mid,))
                action = "ออก" if (last and last[0] == "เข้า") else "เข้า"
                ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.access_repo.log_access(mid, ts, action)
                color = theme.SUCCESS if action == "เข้า" else theme.INFO
                self.status_label.configure(text=f"บันทึกการ{action}ห้องสมุดสำเร็จ ✓",
                                            text_color=color)
                self.last_scan_label.configure(
                    text=f"{name}  •  ชั้น {grade}  •  เลขที่ {number}\nเวลา: {ts}",
                    text_color=theme.text_color())
                show_toast(self, f"บันทึกการ{action}ห้องสมุด: {name}", "green" if action == "เข้า" else "blue")
            else:
                self.status_label.configure(text="ไม่พบสมาชิกในระบบ", text_color=theme.DANGER)
                show_toast(self, "ไม่พบสมาชิกในระบบ", "red")
        except Exception as e:
            logger.error(f"Access scan error: {e}", exc_info=True)
            self.status_label.configure(text="QR Code ไม่ถูกต้อง", text_color=theme.DANGER)
            show_toast(self, "QR Code ไม่ถูกต้อง", "red")


class AccessHistoryView(ctk.CTkFrame):
    """Displays and exports entry/exit history."""

    def __init__(self, parent, nav):
        super().__init__(parent, fg_color="transparent")
        self.nav = nav
        self.repo = AccessRepository()
        self._build()

    def _build(self):
        self.pack(fill="both", expand=True)
        page = ctk.CTkFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True, padx=28, pady=24)

        top = ctk.CTkFrame(page, fg_color="transparent")
        top.pack(fill="x", pady=(0, 14))
        theme.heading(top, "ประวัติการเข้าออก", size=26).pack(side="left")
        theme.secondary_button(top, "Export PDF", self._export_pdf, width=130,
                               height=38).pack(side="right")

        table_card = theme.card(page)
        table_card.pack(fill="both", expand=True)

        header = ctk.CTkFrame(table_card, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(12, 4))
        self.count_label = ctk.CTkLabel(header, text="ประวัติการเข้าออก",
                                        font=theme.font(16, bold=True),
                                        text_color=theme.text_color(), anchor="w")
        self.count_label.pack(side="left")

        self.table_frame = ctk.CTkScrollableFrame(table_card, fg_color="transparent",
                                                  corner_radius=0)
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._load()

    # ================= data =================
    def _load(self):
        for w in self.table_frame.winfo_children():
            w.destroy()

        records = self.repo.get_recent_access(limit=200)
        self.count_label.configure(text=f"ประวัติการเข้าออก ({len(records)})")

        if not records:
            ctk.CTkLabel(self.table_frame, text="ไม่พบประวัติ", font=theme.font(14),
                         text_color=theme.muted_color()).pack(pady=24)
            return

        for r in records:
            ts, name, grade, action = r
            row = ctk.CTkFrame(self.table_frame, fg_color="transparent")
            row.pack(fill="x", padx=2, pady=2)

            ctk.CTkLabel(row, text=f"{name}  •  ชั้น {grade}",
                         font=theme.font(13), text_color=theme.text_color(),
                         anchor="w").pack(side="left", expand=True, fill="x")
            ctk.CTkLabel(row, text=ts, font=theme.font(13),
                         text_color=theme.muted_color(), anchor="w").pack(
                side="left", expand=True, fill="x")

            badge = theme.badge_green(row, "เข้า") if action == "เข้า" \
                else theme.badge_blue(row, "ออก")
            badge.pack(side="left", expand=True, fill="x", padx=8)

            ctk.CTkFrame(row, height=1, fg_color=theme.border_color()).pack(
                side="bottom", fill="x", pady=(6, 0))

    # ================= export =================
    def _export_pdf(self):
        try:
            records = self.repo.get_recent_access(limit=1000)
            rows = [(r[1], r[2], "", r[0], r[3]) for r in records]
            import os
            os.makedirs("reports", exist_ok=True)
            filename = f"reports/access_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            from services.pdf_service import PDFService
            PDFService.generate_access_history_pdf(filename, rows)
            show_toast(self, f"Export สำเร็จ: {filename}", "green")
        except Exception as e:
            logger.error(f"export_access error: {e}", exc_info=True)
            show_toast(self, f"เกิดข้อผิดพลาด: {e}", "red")
