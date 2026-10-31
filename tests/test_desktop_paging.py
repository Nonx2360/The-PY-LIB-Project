# tests/test_desktop_paging.py — the capped list-window math that keeps the
# desktop UI alive with big catalogs.
#
# Background: rendering a book row costs ~10 Tk windows, and every Tk window
# consumes Windows GDI/USER handles. Rendering a 1000-book catalog row-by-row
# in the books page froze the app ("not responding") and eventually crashed
# the process with Tk's uncatchable
# Tcl_Panic("Failed to create the embedded menu window") — the panic fires
# when Windows refuses to create Tk's hidden menu windows because the
# process ran out of handle quota. ui/paging.py slices what gets rendered.
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui import paging  # noqa: E402


class TestVisibleSlice(unittest.TestCase):
    def test_small_catalog_renders_everything(self):
        self.assertEqual(paging.visible_slice(40, 50), (0, 40, False))

    def test_initial_window_is_capped_and_offers_more(self):
        self.assertEqual(paging.visible_slice(1100, paging.INITIAL_ROWS),
                         (0, paging.INITIAL_ROWS, True))

    def test_never_exceeds_hard_cap(self):
        # a teacher clicking "show more" 100 times still stops at HARD_CAP
        limit = paging.INITIAL_ROWS
        for _ in range(100):
            limit = paging.next_limit(limit)
        self.assertEqual(limit, paging.HARD_CAP)
        start, count, more = paging.visible_slice(1100, limit)
        self.assertEqual((start, count), (0, paging.HARD_CAP))
        self.assertFalse(more)          # no button past the cap
        self.assertTrue(paging.reached_cap(1100, limit))

    def test_exact_boundary_show_more_disappears(self):
        # 130 rows: at limit 150 everything fits, no "show more"
        self.assertEqual(paging.visible_slice(130, 150), (0, 130, False))
        # 199 rows at the same limit still shows the button
        _start, _count, more = paging.visible_slice(199, 150)
        self.assertTrue(more)

    def test_step_growth_reaches_cap_in_bounded_clicks(self):
        clicks = 0
        limit = paging.INITIAL_ROWS
        while not paging.reached_cap(10_000, limit) and clicks < 1000:
            limit = paging.next_limit(limit)
            clicks += 1
        self.assertEqual(limit, paging.HARD_CAP)
        self.assertLessEqual(clicks, 10)   # a handful of clicks, not thousands

    def test_empty_and_invalid_inputs_are_safe(self):
        self.assertEqual(paging.visible_slice(0, 50), (0, 0, False))
        # limit 0 = render nothing but rows remain -> button stays available
        self.assertEqual(paging.visible_slice(10, 0), (0, 0, True))
        # a cleared/invalid limit floors at zero, one click brings the
        # initial window back
        self.assertEqual(paging.next_limit(-5), paging.STEP_ROWS)


if __name__ == "__main__":
    unittest.main()
