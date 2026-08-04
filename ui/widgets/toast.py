# ui/widgets/toast.py
import customtkinter as ctk

from core.logger import logger
from ui.theme import font


def show_toast(parent, message, color="green"):
    """Floating toast at the bottom-center of the main window; auto-dismisses."""
    colors = {
        "green": ("#16A34A", "#DCFCE7"),
        "red": ("#DC2626", "#FEE2E2"),
        "amber": ("#D97706", "#FEF3C7"),
        "blue": ("#0284C7", "#E0F2FE"),
        "dark": ("#334155", "#E2E8F0"),
    }
    accent, soft = colors.get(color, colors["dark"])
    try:
        root = parent.winfo_toplevel()
    except Exception:
        return
    try:
        win = ctk.CTkToplevel(root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(fg_color=accent)

        frame = ctk.CTkFrame(win, fg_color=accent, corner_radius=12)
        frame.pack(fill="both", expand=True)
        lbl = ctk.CTkLabel(frame, text=message, font=font(14, bold=True),
                           text_color="#FFFFFF", padx=24, pady=10, wraplength=520)
        lbl.pack()

        win.update_idletasks()
        w = max(lbl.winfo_reqwidth() + 48, 220)
        h = lbl.winfo_reqheight() + 20
        x = root.winfo_rootx() + (root.winfo_width() - w) // 2
        y = root.winfo_rooty() + root.winfo_height() - h - 48
        win.geometry(f"{w}x{h}+{x}+{y}")
        win.after(2800, win.destroy)
    except Exception as exc:
        logger.warning(f"Toast failed: {exc}")
