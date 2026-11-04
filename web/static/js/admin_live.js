/* THE PY-LIB — admin live updates over WebSocket
 *
 * Tiny client for /admin/ws: the server only says WHICH feed topics
 * changed (borrow_log, books, members, access_log — see web/live.py);
 * the page reacts with the data path it already has:
 *   - table pages (books/members/history): in-place DataTables reload
 *     via the hook exposed by tables.js;
 *   - the dashboard: a throttled full reload (read-only page).
 * Action pages (borrow / return / scan) and forms deliberately do not
 * include this script — a mid-typing reload would be destructive.
 */
(function () {
    "use strict";

    var currentScript = document.currentScript;
    var topicsFilter = currentScript && currentScript.getAttribute("data-topic");
    if (!topicsFilter) return;

    var CONNECTS = {
        books: ["books", "borrow_log"],     // book status flips on lend/return
        members: ["members", "borrow_log"], // pin/expire flags live in members
        history: ["borrow_log"],
        dashboard: ["members", "books", "borrow_log", "access_log"],
    };
    var watch = CONNECTS[topicsFilter];
    if (!watch) return;

    var attempts = 0;
    var ws = null;
    var pingTimer = null;
    var lastReloadAt = 0;

    function isTablePage() { return !!window.PyLib.dataTablesReload; }

    function react(changed) {
        var hit = false;
        for (var i = 0; i < watch.length; i++) {
            if (changed.indexOf(watch[i]) !== -1) { hit = true; break; }
        }
        if (!hit) return;
        if (topicsFilter === "dashboard") {
            var now = Date.now();
            if (now - lastReloadAt >= 6000) {
                lastReloadAt = now;
                location.reload();
            }
        } else if (isTablePage()) {
            window.PyLib.dataTablesReload();
        }
    }

    function connect() {
        if (attempts >= 5) return;
        attempts += 1;
        var proto = location.protocol === "https:" ? "wss:" : "ws:";
        try {
            ws = new WebSocket(proto + "//" + location.host + "/admin/ws");
        } catch (e) { return; }

        ws.onopen = function () {
            attempts = 0;
            if (pingTimer) clearInterval(pingTimer);
            pingTimer = setInterval(function () {
                try { ws.send('{"op":"ping"}'); } catch (e) { /* closing */ }
            }, 20000);
        };
        ws.onmessage = function (ev) {
            try {
                var msg = JSON.parse(ev.data);
                if (msg && msg.type === "topics") react(msg.topics || []);
            } catch (e) { /* ignore junk */ }
        };
        ws.onclose = function () {
            if (pingTimer) { clearInterval(pingTimer); pingTimer = null; }
            var delay = Math.min(1000 * attempts, 6000);
            setTimeout(connect, delay);
        };
    }

    connect();
})();
