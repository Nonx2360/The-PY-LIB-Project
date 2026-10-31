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
import customtkinter as ctk

import ui.theme as theme
from ui.paging import (INITIAL_ROWS, STEP_ROWS, HARD_CAP,
                       visible_slice, next_limit, reached_cap)


class RowWindow(ctk.CTkScrollableFrame):

    def __init__(self, master, empty_note="ไม่พบรายการ", **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self._items = []
        self._builder = None
        self._limit = INITIAL_ROWS
        self._empty_note = empty_note
        self.on_status = None      # optional callback(shown, total)

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
        self._render()

    def refresh(self):
        """Re-render with the current limit (data identity unchanged)."""
        self._render()

    def _render(self):
        for w in self.winfo_children():
            w.destroy()

        total = len(self._items)
        _start, count, more = visible_slice(total, self._limit)
        if self.on_status is not None:
            self.on_status(count, total)

        if count == 0:
            ctk.CTkLabel(self, text=self._empty_note, font=theme.font(14),
                         text_color=theme.muted_color()).pack(pady=24)
            return

        for item in self._items[:count]:
            self._builder(item)

        remaining = total - count
        # hard cap reached: offer the search box instead of ever growing
        if reached_cap(total, self._limit):
            ctk.CTkLabel(
                self,
                text=(f"แสดง {count} รายการแรกจากทั้งหมด {total} — "
                      "ใช้ช่องค้นหาเพื่อหารายการอื่น"),
                font=theme.font(12), text_color=theme.muted_color()
            ).pack(pady=(10, 6))
            return
        if more:
            step = min(STEP_ROWS, remaining)
            ctk.CTkButton(self, text=f"แสดงเพิ่มอีก {step} รายการ "
                                     f"(แสดงอยู่ {count} จาก {total})",
                          command=self._show_more, height=32,
                          fg_color="transparent",
                          border_width=1,
                          border_color=theme.border_color(),
                          text_color=theme.muted_color()).pack(pady=(10, 6))

    def _show_more(self):
        self._limit = next_limit(self._limit)
        self._render()
