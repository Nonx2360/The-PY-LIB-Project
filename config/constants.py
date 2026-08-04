# config/constants.py
import os

# Load settings on import
from config.settings import load as _load_settings
_settings = _load_settings()

DEFAULT_DUE_DAYS = _settings.get("default_loan_days", 7)
DB_PATH = 'db/library.db'
QR_DIR = 'assets/qrcodes'
CARD_DIR = 'assets/cards'
LOGOS_DIR = 'assets/logos'
FONTS_DIR = 'assets/fonts'

FONT_REGULAR = 'assets/fonts/Sarabun-Regular.ttf'
FONT_BOLD = 'assets/fonts/Sarabun-Bold.ttf'

COLORS = {
    'primary': 'blue',
    'bg': '#f0f0f0',
    'text': '#000000',
    'error': '#ff0000'
}
