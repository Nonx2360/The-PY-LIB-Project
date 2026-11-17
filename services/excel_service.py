# services/excel_service.py — shared book-catalog Excel import (desktop + web)
import os

import pandas as pd

# Hard field-length guard (the web layer applies its own, tighter LIMITS).
_CODE_MAX_LEN = 64
_TITLE_MAX_LEN = 200


class ExcelService:
    # Import ceiling: keeps memory/latency bounded when a catalog drop is
    # huge; anything larger should be split into several files and imported
    # in batches. MAX_FILE_BYTES mirrors the web layer's MAX_CONTENT_LENGTH
    # (which the desktop has no HTTP equivalent of): a legit 5,000-row
    # Thai-template workbook is well under 1 MB, so an 8 MB candidate is
    # not a catalog, it's something to split.
    MAX_ROWS = 5000
    MAX_FILE_BYTES = 8 * 1024 * 1024

    @staticmethod
    def generate_template_df() -> pd.DataFrame:
        sample_data = {
            'รหัสหนังสือ': ['001', '002'],
            'ชื่อเรื่อง': ['ตัวอย่างหนังสือ 1', 'ตัวอย่างหนังสือ 2']
        }
        return pd.DataFrame(sample_data)

    @staticmethod
    def read_books_from_excel(source) -> list[dict]:
        """Read the Thai-template workbook into [{code, title}, ...].

        `source` is a file path or any pandas-readable file object (e.g. an
        uploaded .xlsx stream from the web form). Designed for bulk imports:
          * rows missing code or title are dropped quietly (no half-records);
          * duplicate codes are KEPT here so callers can count/merge them;
          * control characters are stripped and fields length-capped.
        Raises ValueError when the column layout doesn't match the template
        or when the file exceeds MAX_ROWS.
        """
        # Bounded parse: read at most MAX_ROWS + 1 rows so an oversized file
        # is rejected WITHOUT materializing the whole sheet in RAM. The
        # previous full-sheet read parsed 100k rows (~4.6 s / ~43 MB) just
        # to reject them; this caps rejection at ~0.4 s / a few MB and
        # bounds the memory of every accepted import to 5,000 rows too.
        df = pd.read_excel(source, nrows=ExcelService.MAX_ROWS + 1)
        required_columns = ['รหัสหนังสือ', 'ชื่อเรื่อง']
        if not all(col in df.columns for col in required_columns):
            raise ValueError("รูปแบบไฟล์ Excel ไม่ถูกต้อง กรุณาใช้แม่แบบที่กำหนด")
        if len(df) > ExcelService.MAX_ROWS:
            raise ValueError(
                f"ไฟล์มีมากกว่า {ExcelService.MAX_ROWS:,} แถว — "
                "กรุณาแบ่งเป็นหลายไฟล์ นำเข้าทีละส่วน")

        books = []
        for _, row in df.iterrows():
            code = ExcelService._text(row['รหัสหนังสือ'], _CODE_MAX_LEN)
            title = ExcelService._text(row['ชื่อเรื่อง'], _TITLE_MAX_LEN)
            if not code or not title:
                continue                                   # blank/partial row
            books.append({'code': code, 'title': title})
        return books

    @staticmethod
    def file_exceeds_cap(path) -> bool:
        """Desktop pre-parse guard: True when the file is larger than
        MAX_FILE_BYTES (the desktop stand-in for the web's
        MAX_CONTENT_LENGTH upload cap). Kept on the service instead of
        inline in the view so the wiring is unit-testable without Tk."""
        return os.path.getsize(path) > ExcelService.MAX_FILE_BYTES

    @staticmethod
    def _text(value, max_len) -> str:
        """Cell -> bounded single-line text; NaN/None -> empty string."""
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
        text = "".join(c for c in str(value) if ord(c) >= 32).strip()
        return text[:max_len]


# Module-level aliases: callers (and tests) import these names directly;
# the class attributes above are the single source of truth.
MAX_ROWS = ExcelService.MAX_ROWS
MAX_FILE_BYTES = ExcelService.MAX_FILE_BYTES
