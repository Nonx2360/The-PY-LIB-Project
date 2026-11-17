# tests/test_desktop_import_guard.py — the desktop 8 MB pre-parse guard.
#
# Regression guard: book_view._import_excel reads the size cap THROUGH the
# ExcelService class (ExcelService.MAX_FILE_BYTES). It once referenced a
# constant that only existed at module level, so every desktop import died
# with AttributeError before parsing anything — and the web-only tests
# never executed the desktop path. These tests pin the class-level
# constants and the guard helper's behavior.
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from services.excel_service import ExcelService, MAX_FILE_BYTES, MAX_ROWS


class TestDesktopImportGuard(unittest.TestCase):
    def test_constants_reachable_through_the_class(self):
        # the exact regression: the desktop guard reads these via the class
        self.assertEqual(ExcelService.MAX_FILE_BYTES, 8 * 1024 * 1024)
        self.assertEqual(ExcelService.MAX_ROWS, 5000)
        # module aliases stay in lockstep (single source of truth)
        self.assertEqual(ExcelService.MAX_FILE_BYTES, MAX_FILE_BYTES)
        self.assertEqual(ExcelService.MAX_ROWS, MAX_ROWS)

    def test_small_file_passes_the_guard(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "catalog.xlsx")
            with open(path, "wb") as fh:
                fh.write(b"\x00" * 1024)          # 1 KB, far under 8 MB
            self.assertFalse(ExcelService.file_exceeds_cap(path))

    def test_oversized_file_is_refused(self):
        # shrink the cap so the test writes KBs, not 8 MB
        with mock.patch.object(ExcelService, "MAX_FILE_BYTES", 1024):
            with tempfile.TemporaryDirectory() as tmp:
                path = os.path.join(tmp, "huge.xlsx")
                with open(path, "wb") as fh:
                    fh.write(b"\x00" * 2048)
                self.assertTrue(ExcelService.file_exceeds_cap(path))
        # cap restored after the patch context
        self.assertEqual(ExcelService.MAX_FILE_BYTES, 8 * 1024 * 1024)

    def test_missing_file_raises(self):
        # a vanished selection must not be silently treated as "small"
        with self.assertRaises(OSError):
            ExcelService.file_exceeds_cap(
                os.path.join(tempfile.gettempdir(), "no-such-file-xyz.xlsx"))


if __name__ == "__main__":
    unittest.main()
