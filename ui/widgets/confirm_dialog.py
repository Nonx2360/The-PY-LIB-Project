# ui/widgets/confirm_dialog.py
import customtkinter as ctk

def show_confirm_dialog(parent, title: str, message: str) -> bool:
    """Shows a modal yes/no dialog. Returns True if user clicked yes."""
    dialog = ctk.CTkToplevel(parent)
    dialog.title(title)
    dialog.geometry("400x200")
    dialog.grab_set()

    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() - 400) // 2
    y = (dialog.winfo_screenheight() - 200) // 2
    dialog.geometry(f"400x200+{x}+{y}")

    ctk.CTkLabel(dialog, text=message, font=("Sarabun", 14)).pack(pady=20, padx=20)

    result = {"value": False}

    def on_yes():
        result["value"] = True
        dialog.destroy()

    def on_no():
        dialog.destroy()

    btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    btn_frame.pack(pady=10)
    ctk.CTkButton(btn_frame, text="ใช่", command=on_yes,
                  fg_color="#FF4B4B", hover_color="#FF3333").pack(side="left", padx=10)
    ctk.CTkButton(btn_frame, text="ไม่", command=on_no).pack(side="left", padx=10)

    dialog.wait_window()
    return result["value"]
