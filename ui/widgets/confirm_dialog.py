# ui/widgets/confirm_dialog.py
import customtkinter as ctk

from ui.theme import (DANGER, border_color, card_bg, font, text_color,
                      muted_color, secondary_button)


def show_confirm_dialog(parent, title: str, message: str, ok_text: str = "ยืนยัน",
                        ok_color=None) -> bool:
    """Shows a modal yes/no dialog. Returns True if user clicked ok."""
    dialog = ctk.CTkToplevel(parent)
    dialog.title(title)
    dialog.resizable(False, False)
    dialog.configure(fg_color=card_bg())
    dialog.overrideredirect(False)
    dialog.grab_set()

    card = ctk.CTkFrame(dialog, fg_color="transparent",
                        border_width=1, border_color=border_color(), corner_radius=14)
    card.pack(padx=24, pady=20, fill="both", expand=True)

    ctk.CTkLabel(card, text=title, font=font(19, bold=True),
                 text_color=text_color()).pack(pady=(18, 6), padx=24, anchor="w")
    ctk.CTkLabel(card, text=message, font=font(14), text_color=muted_color(),
                 justify="left", wraplength=360).pack(pady=6, padx=24, anchor="w")

    result = {"value": False}

    def on_ok():
        result["value"] = True
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    btn_frame = ctk.CTkFrame(card, fg_color="transparent")
    btn_frame.pack(pady=(16, 18), padx=24, fill="x")
    secondary_button(btn_frame, "ยกเลิก", on_cancel, width=110).pack(side="right", padx=(8, 0))
    ctk.CTkButton(btn_frame, text=ok_text, command=on_ok, width=110,
                  font=font(13, bold=True), fg_color=ok_color or DANGER,
                  hover_color="#DC2626", corner_radius=8).pack(side="right")

    dialog.update_idletasks()
    w, h = dialog.winfo_reqwidth() + 40, dialog.winfo_reqheight() + 30
    x = (dialog.winfo_screenwidth() - w) // 2
    y = (dialog.winfo_screenheight() - h) // 2
    dialog.geometry(f"{w}x{h}+{x}+{y}")
    dialog.wait_window()
    return result["value"]