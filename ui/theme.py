# ui/theme.py - shared design tokens + helpers for all views
import ctypes
import os
import tkinter.font as tkfont

import customtkinter as ctk

from config.constants import FONTS_DIR

# ---- palette (fixed accents) ----
PRIMARY = "#2563EB"
PRIMARY_HOVER = "#1D4ED8"
PRIMARY_SOFT = "#DBEAFE"
SUCCESS = "#22C55E"
SUCCESS_SOFT = "#DCFCE7"
WARNING = "#F59E0B"
WARNING_SOFT = "#FEF3C7"
DANGER = "#EF4444"
DANGER_SOFT = "#FEE2E2"
INFO = "#0EA5E9"
INFO_SOFT = "#E0F2FE"

SIDEBAR = "#0F172A"
SIDEBAR_HOVER = "#1E293B"
SIDEBAR_ACTIVE = PRIMARY
SIDEBAR_TEXT = "#CBD5E1"
SIDEBAR_MUTED = "#64748B"

# ---- theme-aware colors ----
CONTENT_LIGHT = "#F1F5F9"
CONTENT_DARK = "#0B1220"
CARD_LIGHT = "#FFFFFF"
CARD_DARK = "#1E293B"
TEXT_LIGHT = "#0F172A"
TEXT_DARK = "#F8FAFC"
MUTED_LIGHT = "#64748B"
MUTED_DARK = "#94A3B8"
BORDER_LIGHT = "#E2E8F0"
BORDER_DARK = "#334155"
INPUT_LIGHT = "#F8FAFC"
INPUT_DARK = "#0F172A"

# ---- fonts ----
FONT = "Sarabun"
FONT_BOLD = "Sarabun-Bold"
FALLBACK_FONT = "Segoe UI"


def register_fonts():
    """Load bundled Sarabun TTFs privately (FR_PRIVATE) so Tk can render Thai text."""
    global FONT, FONT_BOLD
    if os.name != "nt":
        return
    for name in ("Sarabun-Regular.ttf", "Sarabun-Bold.ttf"):
        path = os.path.abspath(os.path.join(FONTS_DIR, name))
        if os.path.exists(path):
            try:
                ctypes.windll.gdi32.AddFontResourceExW(path, 0x10, 0)  # FR_PRIVATE
            except Exception:
                pass
    try:
        tkfont.families()  # force Tk to refresh its font table
    except Exception:
        pass
    try:
        families = set(tkfont.families())
        if "Sarabun" not in families:
            FONT = FALLBACK_FONT
        if "Sarabun-Bold" not in families:
            FONT_BOLD = FONT
    except Exception:
        FONT = FALLBACK_FONT
        FONT_BOLD = FALLBACK_FONT


def is_light():
    return ctk.get_appearance_mode() == "Light"


def content_bg():
    return CONTENT_LIGHT if is_light() else CONTENT_DARK


def card_bg():
    return CARD_LIGHT if is_light() else CARD_DARK


def text_color():
    return TEXT_LIGHT if is_light() else TEXT_DARK


def muted_color():
    return MUTED_LIGHT if is_light() else MUTED_DARK


def border_color():
    return BORDER_LIGHT if is_light() else BORDER_DARK


def input_bg():
    return INPUT_LIGHT if is_light() else INPUT_DARK


def font(size=13, bold=False):
    return ctk.CTkFont(family=FONT_BOLD if bold else FONT, size=size,
                       weight="bold" if bold else "normal")


# ---- widget builders ----
def heading(parent, text, size=24):
    return ctk.CTkLabel(parent, text=text, font=font(size, bold=True),
                        text_color=text_color(), anchor="w")


def subheading(parent, text, size=14):
    return ctk.CTkLabel(parent, text=text, font=font(size),
                        text_color=muted_color(), anchor="w")


def card(parent, **kwargs):
    kwargs.setdefault("corner_radius", 14)
    kwargs.setdefault("border_width", 1)
    kwargs.setdefault("fg_color", card_bg())
    return ctk.CTkFrame(parent, corner_radius=kwargs.pop("corner_radius"),
                        border_width=kwargs.pop("border_width"),
                        fg_color=kwargs.pop("fg_color"),
                        border_color=kwargs.pop("border_color", border_color()),
                        **kwargs)


def card_title(parent, text, size=17):
    return ctk.CTkLabel(parent, text=text, font=font(size, bold=True),
                        text_color=text_color(), anchor="w")


def primary_button(parent, text, command, height=40, **kwargs):
    kwargs.setdefault("font", font(14, bold=True))
    return ctk.CTkButton(parent, text=text, command=command, height=height,
                         fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
                         corner_radius=8, **kwargs)


def secondary_button(parent, text, command, height=36, **kwargs):
    kwargs.setdefault("font", font(13))
    kwargs.setdefault("text_color", text_color())
    return ctk.CTkButton(parent, text=text, command=command, height=height,
                         fg_color="transparent",
                         border_width=1, border_color=border_color(),
                         hover_color=border_color(),
                         corner_radius=8, **kwargs)


def ghost_button(parent, text, command, height=36, **kwargs):
    kwargs.setdefault("font", font(13))
    kwargs.setdefault("text_color", muted_color())
    return ctk.CTkButton(parent, text=text, command=command, height=height,
                         fg_color="transparent",
                         hover_color=border_color(),
                         corner_radius=8, **kwargs)


def danger_button(parent, text, command, height=36, **kwargs):
    kwargs.setdefault("font", font(13, bold=True))
    return ctk.CTkButton(parent, text=text, command=command, height=height,
                         fg_color=DANGER, hover_color="#DC2626",
                         corner_radius=8, **kwargs)


def success_button(parent, text, command, height=40, **kwargs):
    kwargs.setdefault("font", font(14, bold=True))
    return ctk.CTkButton(parent, text=text, command=command, height=height,
                         fg_color=SUCCESS, hover_color="#16A34A",
                         corner_radius=8, **kwargs)


def badge(parent, text, color, soft_color):
    lbl = ctk.CTkLabel(parent, text=text, font=font(12, bold=True),
                       fg_color=soft_color, text_color=color,
                       corner_radius=6, padx=10, pady=2)
    return lbl


def badge_green(parent, text):
    return badge(parent, text, "#15803D" if is_light() else "#4ADE80", SUCCESS_SOFT)


def badge_red(parent, text):
    return badge(parent, text, "#B91C1C" if is_light() else "#F87171", DANGER_SOFT)


def badge_amber(parent, text):
    return badge(parent, text, "#B45309" if is_light() else "#FBBF24", WARNING_SOFT)


def badge_blue(parent, text):
    return badge(parent, text, "#0369A1" if is_light() else "#38BDF8", INFO_SOFT)


def entry(parent, placeholder="", width=280, **kwargs):
    kwargs.setdefault("placeholder_text", placeholder)
    e = ctk.CTkEntry(parent, width=width,
                     font=font(14), fg_color=input_bg(), border_color=border_color(),
                     corner_radius=8, **kwargs)
    return e


def label(parent, text, size=13, color=None, bold=False):
    return ctk.CTkLabel(parent, text=text, font=font(size, bold=bold),
                        text_color=color or text_color())
