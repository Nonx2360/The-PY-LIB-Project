# ui/paging.py — pure paging math for the desktop capped-list windows.
#
# Kept free of tkinter so it stays trivially unit-testable (see
# tests/test_desktop_paging.py). Sizing rationale: an idle measure on this
# machine showed ~0.62 ms/widget to create and ~1.5 ms/widget to destroy a
# CTk widget — a book row costs ~10 widgets. Rendering only the visible
# page keeps a 1000-row catalog at <= ~0.4s of UI work where rendering all
# rows froze the desktop app for many seconds and blew past Windows' GDI
# handle quota (Tk then died with "Failed to create the embedded menu
# window" — a Tcl_Panic, uncatchable).

INITIAL_ROWS = 50      # rendered as soon as the page is shown
STEP_ROWS = 80         # appended per "show more" click
HARD_CAP = 300         # absolute widget-count guard (~3.5k Tk windows)


def visible_slice(total, limit):
    """Return the (start, count, more) slice of `total` rows to render.

    `more` is True when rows still exist beyond the slice AND the hard cap
    (HARD_CAP) allows growing further — once at the cap the caller shows a
    "use search" note instead of a button.
    """
    total = max(0, int(total))
    count = min(total, max(0, int(limit)), HARD_CAP)
    more = count < total and count < HARD_CAP
    return (0, count, more)


def next_limit(limit):
    """Grow the window by STEP_ROWS, clamped to HARD_CAP."""
    return min(max(0, int(limit)) + STEP_ROWS, HARD_CAP)


def reached_cap(total, limit):
    """True once the window itself sits at the hard cap and rows still
    remain, so a "use search" note replaces the button."""
    return int(limit) >= HARD_CAP and int(total) > HARD_CAP
