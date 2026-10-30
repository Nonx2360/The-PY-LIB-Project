# ui/theme.py - shared design tokens + helpers for all views
import ctypes
import os
import tkinter.font as tkfont

import customtkinter as ctk

from config.constants import FONTS_DIR

# ---- palette - "ห้องอ่านหนังสือ" Reading Room ----
# Primary carried from assets/logos/school_logo.png (petrol navy);
# status colors are rubber-stamp inks on due-date cards.
PRIMARY = "#1E4B63"
PRIMARY_HOVER = "#173C50"
PRIMARY_SOFT = "#E3ECF1"
SUCCESS = "#276749"
SUCCESS_HOVER = "#1E5237"
SUCCESS_SOFT = "#E8F3EB"
WARNING = "#B45309"          # marigold - the single accent, used sparingly
WARNING_BRIGHT = "#F59E0B"
WARNING_SOFT = "#FBF0DD"
DANGER = "#B3261E"
DANGER_HOVER = "#9A221C"
DANGER_SOFT = "#F9E9E7"
INFO = "#1E4B63"             # info collapses into the navy family
INFO_SOFT = "#E3ECF1"

SIDEBAR = "#14374A"
SIDEBAR_HOVER = "#102C3E"
SIDEBAR_ACTIVE = PRIMARY
SIDEBAR_TEXT = "#A9BFCB"
SIDEBAR_MUTED = "#7E97A5"

# ---- theme-aware colors ----
CONTENT_LIGHT = "#F6F5F1"    # warm near-white stationery paper
CONTENT_DARK = "#0F1A23"
CARD_LIGHT = "#FFFFFF"
CARD_DARK = "#182531"
TEXT_LIGHT = "#20303B"
TEXT_DARK = "#EDF3F6"
MUTED_LIGHT = "#647782"
MUTED_DARK = "#93A6B1"
BORDER_LIGHT = "#E3E0D8"
BORDER_DARK = "#2A3A46"
INPUT_LIGHT = "#F8F7F3"
INPUT_DARK = "#101B24"
INPUT_BORDER_LIGHT = "#CFCBBF"
INPUT_BORDER_DARK = "#33475A"

# ---- fonts ----
FONT = "Sarabun"
FONT_BOLD = "Sarabun-Bold"
FONT_DISPLAY = "Prompt SemiBold"      # headings / stat numbers / tiles
FONT_DISPLAY_MED = "Prompt Medium"    # buttons
FALLBACK_FONT = "Segoe UI"
DISPLAY_OK = False                    # set by register_fonts()
PROMPT_FILES = ("Prompt-Medium.ttf", "Prompt-SemiBold.ttf")


def register_fonts():
    """Load bundled Sarabun + Prompt TTFs privately (FR_PRIVATE) so Tk can
    render Thai text and Reading Room display type."""
    global FONT, FONT_BOLD, FONT_DISPLAY, FONT_DISPLAY_MED, DISPLAY_OK
    if os.name != "nt":
        DISPLAY_OK = True
        return
    ttf_names = ("Sarabun-Regular.ttf", "Sarabun-Bold.ttf") + PROMPT_FILES
    for name in ttf_names:
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
        DISPLAY_OK = FONT_DISPLAY in families and FONT_DISPLAY_MED in families
        if not DISPLAY_OK:
            FONT_DISPLAY = FONT_BOLD
            FONT_DISPLAY_MED = FONT
    except Exception:
        FONT = FALLBACK_FONT
        FONT_BOLD = FALLBACK_FONT
        DISPLAY_OK = False


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


def font(size=13, bold=False, display=False):
    """Widget text.
    display=True  -> Prompt display face (falls back to Sarabun bold if
                     the bundled TTFs are missing).
    display=False -> Sarabun body face.
    """
    if not display:
        fam = FONT_BOLD if bold else FONT
        return ctk.CTkFont(family=fam, size=size,
                           weight="bold" if bold else "normal")
    if DISPLAY_OK:
        fam = FONT_DISPLAY if bold else FONT_DISPLAY_MED
        weight = "normal"
    else:
        fam = FONT_BOLD if bold else FONT
        weight = "bold" if bold else "normal"
    return ctk.CTkFont(family=fam, size=size, weight=weight)


# ---- widget builders ----
def heading(parent, text, size=24):
    return ctk.CTkLabel(parent, text=text, font=font(size, bold=True, display=True),
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
    return ctk.CTkLabel(parent, text=text, font=font(size, bold=True, display=True),
                        text_color=text_color(), anchor="w")


def primary_button(parent, text, command, height=40, **kwargs):
    kwargs.setdefault("font", font(15, display=True))
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
                         fg_color=DANGER, hover_color=DANGER_HOVER,
                         corner_radius=8, **kwargs)


def success_button(parent, text, command, height=40, **kwargs):
    kwargs.setdefault("font", font(15, display=True))
    return ctk.CTkButton(parent, text=text, command=command, height=height,
                         fg_color=SUCCESS, hover_color=SUCCESS_HOVER,
                         corner_radius=8, **kwargs)


def badge(parent, text, color, soft_color):
    """Rubber-stamp style status: 2px ink border on a tinted wash.

    CTkLabel has no border_* arguments (they exist on frames), so the
    stamp is a CTkFrame shell that draws the border, with a transparent
    label inside. Returns the shell; callers may pack/grid it as before.
    """
    shell = ctk.CTkFrame(parent, fg_color=soft_color, border_width=2,
                         border_color=color, corner_radius=4)
    ctk.CTkLabel(shell, text=text, font=font(12, bold=True),
                 fg_color="transparent", text_color=color,
                 padx=10, pady=2).pack()
    return shell


def badge_green(parent, text):
    return badge(parent, text, "#276749" if is_light() else "#7FD4A3", SUCCESS_SOFT)


def badge_red(parent, text):
    return badge(parent, text, "#9A221C" if is_light() else "#EFA6A0", DANGER_SOFT)


def badge_amber(parent, text):
    return badge(parent, text, "#8A4B06" if is_light() else "#E9B57C", WARNING_SOFT)


def badge_blue(parent, text):
    return badge(parent, text, PRIMARY if is_light() else "#9DC0D2", INFO_SOFT)


def entry(parent, placeholder="", width=280, **kwargs):
    kwargs.setdefault("placeholder_text", placeholder)
    e = ctk.CTkEntry(parent, width=width,
                     font=font(14),
                     fg_color=input_bg(),
                     border_color=INPUT_BORDER_LIGHT if is_light() else INPUT_BORDER_DARK,
                     corner_radius=8, **kwargs)
    return e


def label(parent, text, size=13, color=None, bold=False):
    return ctk.CTkLabel(parent, text=text, font=font(size, bold=bold),
                        text_color=color or text_color())
