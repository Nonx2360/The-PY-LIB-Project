/* THE PY-LIB — member portal live updates over WebSocket
 *
 * Opens /member/ws and applies server-computed snapshots: the home-page
 * stat cards/counters, and on the module pages the table body itself.
 * The server pushes whenever its DB markers move (borrow/return/access
 * changes from the web UI *or* the desktop app — see web/live.py).
 *
 * The QR panel and its countdown are owned by member_qr.js and are NOT
 * touched here. Without this script the pages still work as plain
 * server-rendered documents; reconnects back off and stop, so a closed
 * portal or lost session simply leaves the page as-is.
 */
(function () {
    "use strict";

    // only pages inside the logged-in portal layout carry these hooks
    var liveTable = document.querySelector("[data-live-table]");
    var chipCounts = document.querySelectorAll("[data-live-count]");
    if (!liveTable && chipCounts.length === 0) return;

    var attempts = 0;
    var ws = null;
    var pingTimer = null;

    function esc(v) {
        return String(v == null ? "" : v).replace(/[&<>"']/g, function (c) {
            return { "&": "&amp;", "<": "&lt;", ">": "&gt;",
                     '"': "&quot;", "'": "&#39;" }[c];
        });
    }

    /* ---------------- DOM helpers (no innerHTML for data) --------------- */
    function badge(text, cls) {
        var span = document.createElement("span");
        span.className = "badge " + (cls || "badge-danger");
        span.textContent = text;
        return span;
    }

    function emptyRow(cols, iconHref, title, note) {
        var tr = document.createElement("tr");
        tr.className = "member-empty";
        var td = document.createElement("td");
        td.colSpan = cols;
        var svg = document.createElementNS(
            "http://www.w3.org/2000/svg", "svg");
        svg.setAttribute("width", "24");
        svg.setAttribute("height", "24");
        svg.setAttribute("aria-hidden", "true");
        var use = document.createElementNS(
            "http://www.w3.org/2000/svg", "use");
        use.setAttribute("href", iconHref);
        svg.appendChild(use);
        var t = document.createElement("span");
        t.className = "member-empty-title";
        t.textContent = title;
        var n = document.createElement("span");
        n.className = "member-empty-note";
        n.textContent = note;
        td.appendChild(svg); td.appendChild(t); td.appendChild(n);
        tr.appendChild(td);
        return tr;
    }

    function textRow(cells, today, mode) {
        // mode "loans": cells = [title, code, borrow_date, due]
        // mode "returns": same payload, no due badge
        var tr = document.createElement("tr");
        for (var i = 0; i < 4; i++) {
            var td = document.createElement("td");
            td.textContent = cells[i];
            if (mode === "loans" && i === 3) {
                var due = cells[3];
                if (due && due < today) td.appendChild(badge("เกินกำหนด"));
                else if (due && due === today) td.appendChild(badge("วันนี้", "badge-warning"));
            }
            tr.appendChild(td);
        }
        return tr;
    }

    function renderTbody(table, rows, today) {
        var mode = table.getAttribute("data-live-table");
        var tbody = table.tBodies[0];
        if (!tbody) return;
        while (tbody.firstChild) tbody.removeChild(tbody.firstChild);
        if (mode === "loans") {
            rows.forEach(function (r) {
                tbody.appendChild(textRow(r, today, "loans"));
            });
            if (rows.length === 0) tbody.appendChild(emptyRow(
                4, "#icon-book", "ไม่มีหนังสือที่กำลังยืม",
                "รายการจะแสดงที่นี่หลังจากยืมหนังสือที่เคาน์เตอร์"));
        } else if (mode === "returns") {
            rows.forEach(function (r) {
                tbody.appendChild(textRow(r, today, "returns"));
            });
            if (rows.length === 0) tbody.appendChild(emptyRow(
                4, "#icon-history", "ยังไม่มีประวัติการคืนหนังสือ",
                "จะแสดงเมื่อคุณคืนหนังสือที่เคาน์เตอร์แล้ว"));
        } else if (mode === "access") {
            rows.forEach(function (r) {
                var tr = document.createElement("tr"),
                    t1 = document.createElement("td"),
                    t2 = document.createElement("td");
                t1.textContent = r[0]; t2.textContent = r[1];
                tr.appendChild(t1); tr.appendChild(t2);
                tbody.appendChild(tr);
            });
            if (rows.length === 0) tbody.appendChild(emptyRow(
                2, "#icon-clock", "ยังไม่มีประวัติการเข้าใช้",
                "จะแสดงเมื่อคุณสแกน QR ของตัวเองที่หน้าเคาน์เตอร์"));
        }
    }

    /* ---------------- snapshot application ----------------------------- */
    function apply(data) {
        if (!data || data.type !== "snapshot") return;
        var c = data.counts || {};
        var map = { loans: c.loans, overdue: c.overdue,
                    returns: c.returns, access: c.access };
        chipCounts.forEach(function (el) {
            var v = map[el.getAttribute("data-live-count")];
            if (v != null) el.textContent = String(v);
        });
        // overdue badge lives inside the loans stat card; server renders
        // it once — this only re-sets count visibility
        var badgeEl = document.querySelector("[data-live-badge-overdue]");
        if (badgeEl) {
            badgeEl.hidden = !c.overdue;
            var n = badgeEl.querySelector("[data-live-badge-overdue-n]");
            if (n) n.textContent = String(c.overdue || 0);
        }
        if (liveTable) {
            var mode = liveTable.getAttribute("data-live-table");
            var rows = mode === "loans" ? data.loans
                     : mode === "returns" ? data.returns : data.access;
            renderTbody(liveTable, rows || [], data.today);
        }
    }

    /* ---------------- connection --------------------------------------- */
    function connect() {
        if (attempts >= 5) return;   // server off / socket down: give up
        attempts += 1;
        var proto = location.protocol === "https:" ? "wss:" : "ws:";
        try {
            ws = new WebSocket(proto + "//" + location.host + "/member/ws");
        } catch (e) { return; }

        ws.onopen = function () {
            attempts = 0;                    // recovered: reset the backoff
            if (pingTimer) clearInterval(pingTimer);
            // keepalives: the server closes after 30s of silence
            pingTimer = setInterval(function () {
                try { ws.send('{"op":"ping"}'); } catch (e) { /* closing */ }
            }, 20000);
        };
        ws.onmessage = function (ev) {
            try { apply(JSON.parse(ev.data)); } catch (e) { /* ignore junk */ }
        };
        ws.onclose = function () {
            if (pingTimer) { clearInterval(pingTimer); pingTimer = null; }
            var delay = Math.min(1000 * attempts, 6000);   // 1s..6s backoff
            setTimeout(connect, delay);
        };
    }

    connect();
})();
