# vendored frontend stack

All files here are third-party libraries pinned at fixed versions and served
locally (school LAN / offline-friendly). Source: cdn.jsdelivr.net at download
time. Update deliberately: change the pinned version, re-download, and note it
here.

| File | Package | Version | License | Purpose in THE PY-LIB |
|---|---|---|---|---|
| `alpine.min.js` | alpinejs | 3.17.5 | MIT | light reactive state (borrow lock notes) |
| `htmx.min.js` | htmx.org | 2.0.11 | BSD-2-Clause | server-driven partial updates (book/member/history tables) |
| `sweetalert2.min.js` + `.css` | sweetalert2 | 11.26.25 | MIT | confirm dialogs + flash toasts + scanner notices |
| `chart.umd.min.js` | chart.js | 4.6.1 | MIT | dashboard analytics (index) |
| `flatpickr.min.js` + `.css` | flatpickr | 4.6.13 | MIT | due-date picker (borrow) |
| `flatpickr-th.js` | flatpickr | 4.6.13 (l10n/th) | MIT | Thai locale for Flatpickr |

Not vendored on purpose:

- **Lucide icons** — the app already ships the Lucide icon set as the static
  sprite `../templates/icons.svg` (zero-JS variant of the same library).
- **html5-qrcode** — camera scanning uses the vendored `js/jsQR.js`
  with the tailored `js/scan.js` flow (Thai error messages, /scan/resolve).
- **Axios** — the codebase uses native `fetch` everywhere.
- **DOMPurify** — HTMX exchanges complete server-rendered HTML; no client-side
  HTML string assembly happens in this app.

Wiring order matters (see `templates/base.html`):
`sweetalert2.min.js → htmx.min.js → js/ui.js → alpine.min.js` (ui.js must run
before Alpine so its `alpine:init` components are registered).
