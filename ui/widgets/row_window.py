# ui/widgets/row_window.py — scrollable list window with a bounded widget count
#
# A CTk widget costs ~0.6 ms to create and ~1.5 ms to destroy, and every
# Tk window occupies Windows GDI/USER handles (hard per-process quotas).
# Rendering a 1000-row catalog row-by-row therefore froze the desktop app
# for many seconds and could kill the process with Tk's uncatchable
# Tk_Panic("Failed to create the embedded menu window") — see
# ui/paging.py for the measured numbers.
#
# RowWindow renders only the first ui.paging.INITIAL_ROWS rows and offers
# a "show more" button until ui.paging.HARD_CAP; beyond that it asks the
# user to search. Views pass their existing row builder unchanged.
#
# Rendering is CHUNKED (measured 2026-11-14: replacing 50 book rows is
# ~2.5-3 s of solid Tk work — half destroying, half creating). Destroying
# and building both run a few widgets per event-loop tick, so the page
# never hard-blocks input while the list refreshes. "Show more" appends
# the next slice in place instead of rebuilding what is already on screen.
import customtkinter as ctk

import ui.theme as theme
from ui.paging import (INITIAL_ROWS, STEP_ROWS,
                       visible_slice, next_limit, reached_cap)

CHUNK_ROWS = 2          # rows built per event-loop tick (~70 ms measured)
DRAIN_WIDGETS = 12      # stale widgets destroyed per tick (~30 ms measured)
CHUNK_DELAY = 15        # ms between ticks — a real yield for input + paint


class RowWindow(ctk.CTkScrollableFrame):

    def __init__(self, master, empty_note="ไม่พบรายการ", **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self._items = []
        self._builder = None
        self._limit = INITIAL_ROWS
        self._empty_note = empty_note
        self.on_status = None      # optional callback(shown, total)
        self._footer = None        # trailing button/note widget (if any)
        self._tick_after = None    # scheduled next-tick id
        self._render_gen = 0       # invalidates stale scheduled ticks
        self._built_count = 0      # items actually rendered so far

    # ------------- public -------------

    def render(self, items, builder):
        """Replace the contents, rendering at most the currently shown rows.

        `builder(item)` builds ONE row into this window (the exact function
        the view already used with the old loop). A fresh render resets the
        window to INITIAL_ROWS so searches don't grow it indefinitely.
        """
        self._items = list(items)
        self._builder = builder
        self._limit = INITIAL_ROWS
        self._start_render()

    def refresh(self):
        """Re-render with the current limit (data identity unchanged)."""
        self._start_render()

    # ------------- internals -------------

    def _cancel_ticks(self):
        if self._tick_after is not None:
            try:
                self.after_cancel(self._tick_after)
            except Exception:
                pass  # view already destroyed
            self._tick_after = None
        self._render_gen += 1

    def _start_render(self):
        self._cancel_ticks()
        self._footer = None
        self._built_count = 0
        stale = self.winfo_children()
        total = len(self._items)
        _start, count, _more = visible_slice(total, self._limit)
        if self.on_status is not None:
            self.on_status(count, total)
        gen = self._render_gen
        self._tick(gen, stale, 0)

    def _tick(self, gen, stale, si):
        """One event-loop bite: drain stale widgets, then fill new rows."""
        if gen != self._render_gen or not self.winfo_exists():
            return  # a newer render superseded this batch (or view gone)
        stop = min(len(stale), si + DRAIN_WIDGETS)
        while si < stop:
            stale[si].destroy()
            si += 1
        if si < len(stale):
            self._tick_after = self.after(CHUNK_DELAY,
                                          lambda: self._tick(gen, stale, si))
            return
        self._fill(0, gen)

    def _fill(self, start, gen):
        """Build rows [start, shown) in CHUNK_ROWS bites, then the footer."""
        if gen != self._render_gen or not self.winfo_exists():
            return
        total = len(self._items)
        _start, count, _more = visible_slice(total, self._limit)
        if count == 0:
            self._set_footer(ctk.CTkLabel(
                self, text=self._empty_note, font=theme.font(14),
                text_color=theme.muted_color()), 24)
            return
        end = min(count, start + CHUNK_ROWS)
        for item in self._items[start:end]:
            self._builder(item)
        self._built_count = end
        if end < count:
            self._tick_after = self.after(CHUNK_DELAY,
                                          lambda: self._fill(end, gen))
            return
        self._add_footer(count, total)

    def _add_footer(self, count, total):
        remaining = total - count
        # hard cap reached: offer the search box instead of ever growing
        if reached_cap(total, self._limit):
            self._set_footer(ctk.CTkLabel(
                self,
                text=(f"แสดง {count} รายการแรกจากทั้งหมด {total} — "
                      "ใช้ช่องค้นหาเพื่อหารายการอื่น"),
                font=theme.font(12), text_color=theme.muted_color()), (10, 6))
            return
        step = min(STEP_ROWS, remaining)
        self._set_footer(ctk.CTkButton(
            self, text=f"แสดงเพิ่มอีก {step} รายการ "
                       f"(แสดงอยู่ {count} จาก {total})",
            command=self._show_more, height=32,
            fg_color="transparent",
            border_width=1,
            border_color=theme.border_color(),
            text_color=theme.muted_color()), (10, 6))

    def _set_footer(self, widget, pady):
        self._footer = widget
        widget.pack(pady=pady)

    def _show_more(self):
        """Extend the window by STEP_ROWS without rebuilding existing rows."""
        self._cancel_ticks()
        total = len(self._items)
        from_count = self._built_count          # what is REALLY on screen
        self._limit = next_limit(self._limit)
        if self._footer is not None:
            self._footer.destroy()
            self._footer = None
        _start, new_count, _more = visible_slice(total, self._limit)
        if from_count < new_count:
            self._fill(from_count, self._render_gen)
        else:
            self._add_footer(new_count, total)
