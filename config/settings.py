# config/settings.py - load/save user settings to config/settings.json
import json
import os

SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "settings.json")

DEFAULTS = {
    "theme": "System",
    "default_loan_days": 7,
    "max_books_per_member": 3,
    "member_expiry_days": 365,
    "school_name": "DSNPRU",
    "school_logo_path": "assets/logos/school_logo.png",
    "webui_host": "127.0.0.1",  # loopback by default; set to 0.0.0.0 for LAN (with HTTPS)

    # member portal (rotating QR identity + status-only login)
    "member_portal_enabled": True,     # off = /member/* pages 404 for everyone
    "member_qr_rotation_hours": 12,    # each QR window lasts this many hours
    "member_qr_grace_steps": 1,        # accepted window drift (current ± grace)
    "accept_static_card_qr": True,     # printed cards keep working until off'd
}

_settings = dict(DEFAULTS)


def load():
    """Load settings from disk; missing keys get defaults."""
    global _settings
    _settings = dict(DEFAULTS)
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            for k in DEFAULTS:
                if k in data:
                    _settings[k] = data[k]
        except Exception:
            pass
    return _settings


def save():
    """Persist current settings to disk."""
    try:
        os.makedirs(os.path.dirname(SETTINGS_FILE) or ".", exist_ok=True)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(_settings, f, indent=2, ensure_ascii=False)
    except Exception as exc:
        from core.logger import logger
        logger.error(f"Failed to save settings: {exc}")


def get(key):
    return _settings.get(key, DEFAULTS.get(key))


def set(key, value):
    _settings[key] = value


def all():
    return dict(_settings)
