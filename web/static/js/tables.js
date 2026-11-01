/* THE PY-LIB — list-table engine for /books /members /history
 *
 * Backbone: jQuery 3.7.1 + DataTables 1.13.8 (both vendored), server-side
 * mode. The page ships a skeleton <table>; DataTables drives paging,
 * searching and sorting against the /api/{books,members,history} feeds, so
 * a 10,000-row catalog costs the same as a 10-row one.
 *
 * Loaded by books.html / members.html / history.html:
 *   <script src=".../js/tables.js" data-table="books"></script>
 *
 * What sits on top of stock DataTables:
 *   - Thai language file served from /static/lang/datatables-th.json
 *     (hand-authored — no official i18n asset is vendored);
 *   - the server's ORDER BY whitelist is a fixed (SQL identifier, dict key)
 *     list, so the browser's table-column index is translated into that
 *     whitelist index here (checkbox/action columns are not sortable);
 *   - bulk management: per-row checkboxes keep their selection while
 *     paging through the table, and the bulk bar POSTs {action, ids} to
 *     /books/bulk or /members/bulk with the session CSRF token in the
 *     X-CSRF-Token header;
 *   - per-row buttons reuse the classic endpoints (POST → redirect) via
 *     fetch, then redraw the table from the server's answer.
 *
 * The CSRF token comes from the add-form rendered by the same page
 * (input[name="csrf_token"]) — see PyLib.csrfToken() in ui.js.
 */
(function () {
    "use strict";

    var currentScript = document.currentScript;
    var kind = currentScript && currentScript.getAttribute("data-table");
    if (!kind) return;

    var $ = window.jQuery;
    if (!$ || !$.fn.DataTable) return;
    if (!window.PyLib) return;

    /* ------------------------------------------------------------------ *
     *  Rendering helpers (escape on display only — sorting uses raw data)
     * ------------------------------------------------------------------ */
    function esc(v) {
        return String(v == null ? "" : v).replace(/[&<>"']/g, function (c) {
            return { "&": "&amp;", "<": "&lt;", ">": "&gt;",
                     '"': "&quot;", "'": "&#39;" }[c];
        });
    }

    function escDisplay(v, type) { return type === "display" ? esc(v) : v; }

    function statusBadge(status, type) {
        if (type !== "display") return status;
        if (status === "ว่าง")
            return '<span class="badge badge-success">ว่าง</span>';
        if (status === "หยุดให้บริการ")
            return '<span class="badge badge-muted">หยุดให้บริการ</span>';
        return '<span class="badge badge-warning">' + esc(status) + "</span>";
    }

    function returnedBadge(returned, type) {
        if (type !== "display") return returned;
        return returned
            ? '<span class="badge badge-success">คืนแล้ว</span>'
            : '<span class="badge badge-warning">กำลังยืม</span>';
    }

    function dateCell(v, type) { return type === "display" && !v ? "—" : v; }

    /* ------------------------------------------------------------------ *
     *  Table definitions — keyed by the script tag's data-table value
     * ------------------------------------------------------------------ */
    var TABLES = {
        books: {
            url: "/api/books",
            bulkUrl: "/books/bulk",
            bulk: true,
            /* display-column index -> server ORDER-BY whitelist index;
             * column 0 (checkbox) and 5 (actions) are not sortable */
            colMap: { 1: 1, 2: 2, 3: 3, 4: 4 },
            extraData: function () {
                var sel = document.getElementById("booksStatus");
                return sel && sel.value ? { status: sel.value } : {};
            },
            columns: [
                { data: null, orderable: false, className: "dt-check",
                  render: function (data, type, row) {
                      return '<input type="checkbox" class="row-check" data-id="'
                          + esc(row.id)
                          + '" aria-label="เลือกแถว ' + esc(row.code) + '">';
                  } },
                { data: "code", orderable: true, render: escDisplay },
                { data: "title", orderable: true, render: escDisplay },
                { data: "status", orderable: true, render: statusBadge },
                { data: "history", orderable: true,
                  render: function (v) {
                      return v > 0 ? esc(v) + " ครั้ง" : "—";
                  } },
                { data: null, orderable: false, className: "actions-cell",
                  render: function (data, type, row) { return rowActions(row); } },
            ],
        },

        members: {
            url: "/api/members",
            bulkUrl: "/members/bulk",
            bulk: true,
            colMap: { 1: 1, 2: 2, 3: 3, 4: 4, 5: 5 },
            extraData: null,
            columns: [
                { data: null, orderable: false, className: "dt-check",
                  render: function (data, type, row) {
                      return '<input type="checkbox" class="row-check" data-id="'
                          + esc(row.id)
                          + '" aria-label="เลือกสมาชิก ' + esc(row.name) + '">';
                  } },
                { data: "name", orderable: true, render: escDisplay },
                { data: "grade", orderable: true, render: escDisplay },
                { data: "number", orderable: true, render: escDisplay },
                { data: "register_date", orderable: true, render: dateCell },
                { data: "expire_date", orderable: true, render: dateCell },
                { data: null, orderable: false, className: "actions-cell",
                  render: function (data, type, row) { return rowActions(row); } },
            ],
        },

        history: {
            /* read-only; display col N maps to whitelist N+1 (whitelist
             * keeps bl.id at 0 for the default newest-first order) */
            url: "/api/history",
            bulk: false,
            colMap: { 0: 1, 1: 2, 2: 3, 3: 4, 4: 5, 5: 6, 6: 7 },
            extraData: null,
            columns: [
                { data: "name", orderable: true, render: escDisplay },
                { data: "grade", orderable: true, render: escDisplay },
                { data: "code", orderable: true, render: escDisplay },
                { data: "title", orderable: true, render: escDisplay },
                { data: "borrow_date", orderable: true, render: dateCell },
                { data: "return_due", orderable: true, render: dateCell },
                { data: "returned", orderable: true, render: returnedBadge },
            ],
        },
    };
    var cfg = TABLES[kind];
    if (!cfg) return;

    var tableSel = "#" + kind + "Table";

    /* ------------------------------------------------------------------ *
     *  Per-row actions (classic POST → redirect endpoints)
     * ------------------------------------------------------------------ */
    const ACTIONS = {
        books: {
            retire:  { url: function (id) { return "/books/" + id + "/retire"; },
                       confirm: null },
            restore: { url: function (id) { return "/books/" + id + "/restore"; },
                       confirm: null },
            delete:  { url: function (id) { return "/books/" + id + "/delete"; },
                       confirm: function (btn) {
                           return "ลบหนังสือ " + btn.getAttribute("data-title")
                               + "?";
                       } },
            purge:   { url: function (id) { return "/books/" + id + "/delete"; },
                       confirm: function (btn) {
                           return "ลบถาวร: หนังสือ " + btn.getAttribute("data-title")
                               + " พร้อมประวัติการยืม "
                               + btn.getAttribute("data-history")
                               + " รายการ — ลบแล้วไม่สามารถย้อนกลับ"
                               + " ยืนยันหรือไม่?";
                       } },
        },
        members: {
            delete: { url: function (id) { return "/members/" + id + "/delete"; },
                      confirm: function (btn) {
                          return "ลบสมาชิก " + btn.getAttribute("data-title") + "?";
                      } },
        },
    };

    function trashBtn(label) {
        var icon = "<svg width='12' height='12'>"
            + "<use href='#icon-trash'/></svg>";
        return label ? icon + " " + esc(label) : icon;
    }

    function rowActions(row) {
        var id = esc(row.id);
        if (kind === "books") {
            var parts = [];
            if (row.status === "ว่าง") {
                parts.push(
                    "<button type='button' class='btn btn-secondary btn-sm'"
                    + " data-row-action='retire' data-id='" + id + "'"
                    + " title='หยุดให้บริการ — ออกจากรายการยืม"
                    + " ประวัติถูกเก็บไว้'>หยุด</button>");
            } else if (row.status === "หยุดให้บริการ") {
                parts.push(
                    "<button type='button' class='btn btn-success btn-sm'"
                    + " data-row-action='restore' data-id='" + id + "'"
                    + " title='กลับมาให้บริการตามปกติ'>กลับมา</button>");
            }
            if (row.status === "หยุดให้บริการ") {
                parts.push(
                    "<button type='button' class='btn btn-danger btn-sm'"
                    + " data-row-action='purge' data-id='" + id + "'"
                    + " data-title='" + esc(row.title) + "'"
                    + " data-history='" + esc(row.history) + "'"
                    + " title='ลบถาวร พร้อมประวัติการยืม " + esc(row.history)
                    + " รายการ'>" + trashBtn("ลบถาวร") + "</button>");
            } else if (row.status === "ว่าง" && !row.history) {
                parts.push(
                    "<button type='button' class='btn btn-danger btn-sm'"
                    + " data-row-action='delete' data-id='" + id + "'"
                    + " data-title='" + esc(row.title) + "'"
                    + " title='ลบ'>" + trashBtn() + "</button>");
            }
            if (!parts.length) {
                return "<span class='text-muted'"
                    + " title='ยืมอยู่/มีประวัติ — ต้องคืนและหยุดให้บริการ"
                    + "ก่อนจึงจัดการได้'>—</span>";
            }
            return parts.join(" ");
        }
        if (kind === "members") {
            var qr = "";
            if (row.qrcode_path) {
                qr = "<a href='/assets/" + encodeURIComponent(
                        String(row.qrcode_path).replace(/^assets\//, ""))
                    + "' target='_blank' class='btn btn-ghost btn-sm'"
                    + " title='ดู QR'><svg width='14' height='14'>"
                    + "<use href='#icon-image'/></svg></a>";
            }
            return qr
                + "<button type='button' class='btn btn-danger btn-sm'"
                + " data-row-action='delete' data-id='" + id + "'"
                + " data-title='" + esc(row.name) + "'"
                + " title='ลบ'>" + trashBtn() + "</button>";
        }
        return "";
    }

    /* ------------------------------------------------------------------ *
     *  Bulk selection state (persists across pagination)
     * ------------------------------------------------------------------ */
    var selected = new Set();
    var busy = false;

    function renderedIds() {
        var out = [];
        $(tableSel + " tbody input.row-check").each(function () {
            var id = parseInt(this.getAttribute("data-id"), 10);
            if (id) out.push(id);
        });
        return out;
    }

    function syncBulkBar() {
        var bar = document.querySelector(".bulk-bar[data-table='" + kind + "']");
        if (bar) bar.hidden = selected.size === 0;
        var countEl = document.getElementById(kind + "BulkCount");
        if (countEl) countEl.textContent = String(selected.size);
        var all = document.getElementById(kind + "CheckAll");
        if (all) {
            var ids = renderedIds();
            var n = ids.filter(function (id) { return selected.has(id); }).length;
            all.checked = ids.length > 0 && n === ids.length;
            all.indeterminate = n > 0 && n < ids.length;
        }
    }

    /* ------------------------------------------------------------------ *
     *  Networking
     * ------------------------------------------------------------------ */
    function buildAjax() {
        return function (d, cb, settings) {
            var extra = cfg.extraData ? cfg.extraData() : null;
            if (extra) Object.assign(d, extra);
            /* display-column index -> server ORDER-BY whitelist index;
             * no order yet (or an unsortable column) = id, newest first */
            var o = d.order && d.order[0];
            var serverCol = o ? cfg.colMap[Number(o.column)] : 0;
            if (serverCol == null) serverCol = 0;
            var dir = (o && (o.dir === "asc" || o.dir === "desc"))
                ? o.dir : "desc";
            d.order = [{ column: serverCol, dir: dir }];
            var drawNum = settings.iDraw;
            $.ajax({
                url: cfg.url, type: "GET", data: d, dataType: "json",
            }).done(cb).fail(function (xhr) {
                /* A failure (session expired → login redirect, 500, offline)
                 * must not leave a broken page — collapse to an empty set. */
                PyLib.toast("เซสชันหมดอายุหรือเกิดข้อผิดพลาด"
                            + " โปรดล็อกอินใหม่แล้วลองอีกครั้ง", "error");
                cb({ draw: drawNum, data: [],
                     recordsTotal: 0, recordsFiltered: 0 });
            });
        };
    }

    function postJson(url, body) {
        return fetch(url, {
            method: "POST",
            credentials: "same-origin",
            headers: {
                "Content-Type": "application/json",
                "X-CSRF-Token": PyLib.csrfToken() || "",
            },
            body: JSON.stringify(body),
        }).then(function (r) {
            return r.text().then(function (text) {
                var j = null;
                try { j = JSON.parse(text); } catch (e) { /* not JSON */ }
                if (!r.ok || !j || j.status === "error") {
                    throw new Error((j && j.message)
                        || "เซสชันหมดอายุ โปรดล็อกอินใหม่แล้วลองอีกครั้ง");
                }
                return j;
            });
        });
    }

    /* Per-row buttons hit the classic endpoints which answer with a
     * redirect → page HTML. A success flash inside that HTML never shows
     * (we stay on the page), so this mirrors the flash categories. */
    function rowPost(url) {
        return fetch(url, {
            method: "POST",
            credentials: "same-origin",
            headers: { "X-CSRF-Token": PyLib.csrfToken() || "" },
        }).then(function (r) {
            if (r.status === 403) {
                PyLib.toast("เซสชันหมดอายุ โปรดล็อกอินใหม่", "error");
                return;
            }
            if (!r.ok) {
                PyLib.toast("ดำเนินการไม่สำเร็จ", "error");
                return;
            }
            return r.text().then(function (t) {
                var failed = /class=["']alert alert-error/.test(t);
                PyLib.toast(failed ? "ดำเนินการไม่สำเร็จ" : "ดำเนินการสำเร็จ",
                            failed ? "error" : "success");
                tableEl.ajax.reload(null, false);
            });
        }).catch(function () {
            PyLib.toast("เกิดข้อผิดพลาดเครือข่าย", "error");
        });
    }

    /* ------------------------------------------------------------------ *
     *  Bulk handlers
     * ------------------------------------------------------------------ */
    function bulkConfirmText(action, n) {
        if (kind === "books") {
            if (action === "retire")
                return "หยุดให้บริการหนังสือที่เลือก " + n + " เล่ม?"
                    + " เล่มที่ถูกยืมอยู่จะถูกข้าม";
            if (action === "restore")
                return "ให้บริการหนังสือที่เลือกกลับมา " + n + " เล่ม?"
                    + " เล่มที่ไม่ได้หยุดให้บริการจะถูกข้าม";
            return "ลบหนังสือที่เลือก " + n + " เล่ม?"
                + " เล่มที่หยุดให้บริการจะลบถาวรพร้อมประวัติการยืม"
                + " เล่มที่มีประวัติการยืมแต่ยังให้บริการอยู่จะถูกข้าม";
        }
        return "ลบสมาชิกที่เลือก " + n + " คน?"
            + " สมาชิกที่มีประวัติการยืมหรือการเข้า-ออกจะถูกข้าม";
    }

    function runBulk(action) {
        if (busy || !selected.size) return;
        var ids = Array.from(selected);
        var n = ids.length;
        PyLib.confirm({
            text: bulkConfirmText(action, n),
            icon: "warning",
            danger: true,
            confirmText: action === "delete" ? "ลบ" : "ยืนยัน",
        }).then(function (ok) {
            if (!ok) return;
            busy = true;
            postJson(cfg.bulkUrl, { action: action, ids: ids })
                .then(function (out) {
                    var skipped = out.skipped || 0;
                    var msg = (action === "delete")
                        ? "ลบสำเร็จ " + out.deleted + " / ข้าม " + skipped
                        : "สำเร็จ " + out.changed + " / ข้าม " + skipped;
                    PyLib.toast(msg, skipped ? "warning" : "success");
                    selected.clear();
                    tableEl.ajax.reload(null, false);
                    syncBulkBar();
                })
                .catch(function (err) {
                    PyLib.toast(err.message || "เกิดข้อผิดพลาด", "error");
                })
                .then(function () { busy = false; });
        });
    }

    /* ------------------------------------------------------------------ *
     *  Init
     * ------------------------------------------------------------------ */
    var tableEl = $(tableSel).DataTable({
        processing: true,
        serverSide: true,
        searchDelay: 400,     /* every keystroke is a server COUNT — wait */
        lengthMenu: [[10, 25, 50, 100], [10, 25, 50, 100]],
        pageLength: 50,
        order: [],            /* server default: newest (id) first */
        language: { url: "/static/lang/datatables-th.json" },
        ajax: buildAjax(),
        columns: cfg.columns,
    });

    /* keep the header checkbox honest after every redraw */
    tableEl.on("draw", syncBulkBar);

    $(tableSel + " tbody").on("change", "input.row-check", function () {
        var id = parseInt(this.getAttribute("data-id"), 10);
        if (!id) return;
        if (this.checked) selected.add(id); else selected.delete(id);
        syncBulkBar();
    });

    var checkAll = document.getElementById(kind + "CheckAll");
    if (checkAll) {
        checkAll.addEventListener("change", function () {
            var ids = renderedIds();
            ids.forEach(function (id) {
                if (checkAll.checked) selected.add(id);
                else selected.delete(id);
            });
            $(tableSel + " tbody input.row-check").each(function () {
                this.checked = checkAll.checked;
            });
            syncBulkBar();
        });
    }

    var actionDef = ACTIONS[kind] || {};
    $(tableSel + " tbody").on("click", "[data-row-action]", function () {
        if (busy) return;
        var btn = this;
        var def = actionDef[btn.getAttribute("data-row-action")];
        if (!def) return;
        var proceed = def.confirm
            ? PyLib.confirm({ text: def.confirm(btn),
                              icon: "warning", danger: true })
            : Promise.resolve(true);
        proceed.then(function (ok) {
            if (ok) rowPost(def.url(btn.getAttribute("data-id")));
        });
    });

    var barEl = cfg.bulk
        ? document.querySelector(".bulk-bar[data-table='" + kind + "']")
        : null;
    if (barEl) {
        barEl.addEventListener("click", function (e) {
            var button = e.target.closest("[data-bulk]");
            if (button) runBulk(button.getAttribute("data-bulk"));
        });
    }

    /* books page: a new status filter restarts from page 1 */
    var statusSel = document.getElementById("booksStatus");
    if (statusSel) {
        statusSel.addEventListener("change", function () {
            tableEl.page(0);
            tableEl.ajax.reload(null, false);
        });
    }
})();
