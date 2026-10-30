# config/constants.py

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

# ------------------------- book circulation states -------------------------
# Stored verbatim in books.status. The Thai strings ARE the storage format:
# legacy databases and every borrow/return/retire flow match on the exact
# wording, so they must never be edited casually. Use these constants
# instead of string literals in code.
BOOK_STATUS_AVAILABLE = 'ว่าง'
BOOK_STATUS_BORROWED = 'ยืมแล้ว'
BOOK_STATUS_RETIRED = 'หยุดให้บริการ'
