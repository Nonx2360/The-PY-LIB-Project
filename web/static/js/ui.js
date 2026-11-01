/* THE PY-LIB — shared browser helpers for the WebUI
 *
 * Works with the vendored stack loaded in base.html, in this exact order:
 *   sweetalert2 → ui.js (this file) → [page scripts: jQuery + DataTables]
 *   → alpine
 *
 * Responsibilities:
 *   - themed confirm()/alert()/toast() built on SweetAlert2
 *   - converts all server-rendered .alert flash messages into toasts
 *   - intercepts <form data-confirm="..."> submissions (delegated, so it
 *     also covers rows rendered dynamically by the list-table engine)
 *   - auto-inits Flatpickr on [data-flatpickr] inputs (JS + Thai locale
 *     load in the pages that need them, e.g. borrow)
 *   - registers the Alpine "pylibBorrow" component used by borrow.html
 *
 * Everything degrades to the pre-alpine behavior when the vendor stack is
 * absent (no-JS or old server build): plain confirm(), plain alerts,
 * native date input, forms posting the classic way.
 */
(function () {
    "use strict";

    var CONFIRM = { confirmButtonColor: "#1E4B63", cancelButtonColor: "#647782" };

    function swalAvailable() { return typeof window.Swal === "object" && window.Swal; }

    var PyLib = {
        csrfToken: function () {
            var el = document.querySelector('input[name="csrf_token"]');
            return el ? el.value : "";
        },

        /* Promise<confirmed:boolean> */
        confirm: function (opts) {
            opts = opts || {};
            var message = opts.text || "ยืนยันการทำรายการนี้?";
            if (!swalAvailable()) return Promise.resolve(window.confirm(message));
            return window.Swal.fire({
                title: opts.title || "",
                text: message,
                icon: opts.icon || "question",
                showCancelButton: true,
                confirmButtonText: opts.confirmText || "ยืนยัน",
                cancelButtonText: opts.cancelText || "ยกเลิก",
                confirmButtonColor: opts.danger ? "#B3261E" : CONFIRM.confirmButtonColor,
                cancelButtonColor: CONFIRM.cancelButtonColor,
                showClass: { popup: "" },   /* keep it snappy on school hardware */
                hideClass: { popup: "" }
            }).then(function (r) { return !!r.isConfirmed; });
        },

        alert: function (opts) {
            opts = opts || {};
            if (!swalAvailable()) { window.alert(opts.text || ""); return; }
            window.Swal.fire({
                title: opts.title || "",
                text: opts.text || "",
                icon: opts.icon || "info",
                confirmButtonText: "ตกลง",
                confirmButtonColor: CONFIRM.confirmButtonColor,
                showClass: { popup: "" },
                hideClass: { popup: "" }
            });
        },

        toast: function (message, type) {
            if (!swalAvailable()) return;
            window.Swal.fire({
                toast: true, position: "top-end", timer: 3500, timerProgressBar: true,
                showConfirmButton: false, title: message, icon: type === "success" ? "success" : (type === "warning" ? "warning" : "error"),
                showClass: { popup: "" }, hideClass: { popup: "" }
            });
        },

        /* Quick due-date buttons (borrow page): pick #dueDate, keep the
         * server king — the value it receives stays a plain YYYY-MM-DD. */
        setDueDate: function (days) {
            var el = document.getElementById("dueDate");
            if (!el) return;
            var d = new Date();
            d.setDate(d.getDate() + days);
            var iso = d.toISOString().split("T")[0];
            if (el._flatpickr) el._flatpickr.setDate(iso, true);
            else el.value = iso;
        },

        initFlatpickr: function (el, extra) {
            if (!window.flatpickr) return null;
            var localeTh = window.flatpickr.l10ns && window.flatpickr.l10ns.th;
            if (localeTh) window.flatpickr.localize(localeTh);
            var inst = window.flatpickr(el, Object.assign({
                dateFormat: "Y-m-d",           /* submitted value: YYYY-MM-DD always */
                altInput: true,
                altFormat: "j F Y",            /* Thai month names via th locale */
                allowInput: false,
                disableMobile: true            /* consistent UI + quick buttons */
            }, extra || {}));
            return (el._flatpickr = inst);
        }
    };
    window.PyLib = PyLib;

    /* ---- flash messages -> toasts (drop the originals so they appear once) */
    document.addEventListener("DOMContentLoaded", function () {
        if (!swalAvailable()) return;
        var alerts = document.querySelectorAll(".alert");
        alerts.forEach(function (el) {
            var ok = el.classList.contains("alert-success");
            PyLib.toast(el.textContent.trim(), ok ? "success" : "error");
            el.remove();
        });
    });

    /* ---- data-confirm: replaces inline onclick="return confirm(...)" ----
     * capture phase so the native submit is stopped before anything else.
     * data-confirm may sit on the <form> or on the triggering submit button
     * (per-row messages live on the button for dynamically rendered rows). */
    document.addEventListener("submit", function (e) {
        var form = e.target;
        if (!form || !form.getAttribute) return;
        var msg = form.getAttribute("data-confirm");
        if (!msg) {
            var btn = e.submitter
                || form.querySelector('button[data-confirm][type="submit"], [type="submit"][data-confirm]');
            if (btn) msg = btn.getAttribute("data-confirm");
        }
        if (!msg || form.dataset.confirmed === "1") return;
        e.preventDefault();
        var danger = (e.submitter && e.submitter.hasAttribute("data-confirm-danger"))
            || form.hasAttribute("data-confirm-danger");
        if (!swalAvailable()) {
            if (window.confirm(msg)) { form.dataset.confirmed = "1"; form.submit(); }
            return;
        }
        PyLib.confirm({ text: msg, icon: danger ? "warning" : "question", danger: danger })
            .then(function (ok) {
                if (ok) { form.dataset.confirmed = "1"; form.submit(); }
            });
    }, true);

    /* ---- Flatpickr auto-init (page loads its JS/CSS first) ---- */
    document.addEventListener("DOMContentLoaded", function () {
        if (!window.flatpickr) return;
        document.querySelectorAll("[data-flatpickr]").forEach(function (el) {
            if (el._flatpickr) return;
            var extra = {};
            if (el.dataset.flatpickrMin === "today") extra.minDate = "today";
            PyLib.initFlatpickr(el, extra);
        });
    });

    /* ---- Alpine components (registered before Alpine initializes) ---- */
    document.addEventListener("alpine:init", function () {
        if (!window.Alpine || !window.Alpine.data) return;
        window.Alpine.data("pylibBorrow", function (maxBooks) {
            return {
                memberId: "",
                active: 0,
                overdue: 0,
                maxBooks: maxBooks || 3,
                get canBorrow() {
                    return !!this.memberId && this.overdue === 0 && this.active < this.maxBooks;
                },
                get reason() {
                    if (!this.memberId) return "";
                    if (this.overdue > 0)
                        return "สมาชิกคนนี้มีหนังสือค้างส่ง " + this.overdue + " เล่ม — ต้องคืนก่อนจึงจะยืมได้";
                    if (this.active >= this.maxBooks)
                        return "ยืมครบจำนวนที่กำหนดแล้ว (" + this.maxBooks + " เล่ม)";
                    return "";
                },
                syncMember: function (event) {
                    var opt = event.target.options[event.target.selectedIndex];
                    this.memberId = opt.value;
                    this.active = parseInt(opt.dataset.active || 0, 10);
                    this.overdue = parseInt(opt.dataset.overdue || 0, 10);
                }
            };
        });
    });
})();
