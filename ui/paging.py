# ui/paging.py — pure paging math for the desktop capped-list windows.
#
# Kept free of tkinter so it stays trivially unit-testable (see
# tests/test_desktop_paging.py).
#
# Re-measured 2026-11-14 (this machine, Python 3.10 / Tk 8.6): a book row
# costs ~15-35 ms to create and ~17-26 ms to destroy (rounded CTk card +
# checkbox + labels + badge + buttons ≈ 10 widgets) — a 50-row page is
# 2.5-3 s of solid UI work, and rendering all rows froze the app for many
# seconds and blew past Windows' GDI handle quota (Tk then died with
# "Failed to create the embedded menu window" — a Tcl_Panic, uncatchable).
# Page sizes below keep one page ≈ 1 s of streamed work and a hard cap of
# ~1000 Tk windows; RowWindow streams both phases chunk-by-chunk so no
# single event-loop bite blocks input for long.

INITIAL_ROWS = 25      # rendered as soon as the page is shown
STEP_ROWS = 25         # appended per "show more" click
HARD_CAP = 100         # absolute widget-count guard (~1000 Tk windows)


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
