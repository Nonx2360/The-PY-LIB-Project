# vendored frontend stack

All files here are third-party libraries pinned at fixed versions and served
locally (school LAN / offline-friendly). Source: cdn.jsdelivr.net or the
upstream CDN at download time. Update deliberately: change the pinned version,
re-download, and note it here.

| File | Package | Version | License | Purpose in THE PY-LIB |
|---|---|---|---|---|
| `alpine.min.js` | alpinejs | 3.17.5 | MIT | light reactive state (borrow lock notes) |
| `jquery.min.js` | jquery | 3.7.1 | MIT | DOM engine for DataTables (list tables) |
| `datatables.min.js` | datatables.net-dt (core JS) | 1.13.8 | Apache-2.0 | server-side paging/search/sort for /books /members /history |
| `sweetalert2.min.js` + `.css` | sweetalert2 | 11.26.25 | MIT | confirm dialogs + flash toasts + scanner notices |
| `chart.umd.min.js` | chart.js | 4.6.1 | MIT | dashboard analytics (index) |
| `flatpickr.min.js` + `.css` | flatpickr | 4.6.13 | MIT | due-date picker (borrow) |
| `flatpickr-th.js` | flatpickr | 4.6.13 (l10n/th) | MIT | Thai locale for Flatpickr |
| `htmx.min.js` | htmx.org | 2.0.11 | BSD-2-Clause | RETAINED BUT UNUSED — list tables moved to DataTables; kept in case a future page wants partial swaps |

Hand-made by this project (no official asset is vendored):

- `../lang/datatables-th.json` — our own Thai translation of the DataTables
  language keys (cdn.datatables.net refused to serve its i18n JSON at fetch
  time), loaded by `../js/tables.js` via `language: {url: ...}`;
- the entire DataTables surface (wrapper, search box, length menu, sort
  arrows, paginate buttons) is styled inside `../style.css` against the
  Reading Room design tokens — the official `dataTables.dataTables.min.css`
  404'd on both cdn.datatables.net and jsdelivr at vendoring time.

Rules: every reference in HTML/JS is local (`url_for('static', ...)`), never
a CDN URL; the app must keep working on an offline LAN.
