/* THE PY-LIB — member portal live QR countdown
 *
 * The server renders <img id="memberQr"> for the CURRENT rotation window
 * plus the whole seconds-left. This file ticks that down, and when the
 * window flips it pulls a fresh PNG (?r= cache-buster) from
 * /member/qr.png — session-bound, rate-limited server-side (30/min), and
 * no-store, so both sides agree caching must never serve a stale window.
 */
(function () {
    "use strict";
    var img = document.getElementById("memberQr"),
        box = document.getElementById("qrCountdown"),
        txt = document.getElementById("qrCountText");
    if (!img || !box) return;
    var left = parseInt(box.getAttribute("data-seconds"), 10) || 0;
    var rotation = parseInt(box.getAttribute("data-rotation"), 10) || 43200;
    if (rotation < 60) rotation = 60;
    var lastFetch = 0;

    function fmt(s) {
        var p = function (n) { return (n < 10 ? "0" : "") + n; };
        var h = (s / 3600) | 0, m = ((s % 3600) / 60) | 0, sec = s % 60;
        return h + ":" + p(m) + ":" + p(sec);
    }

    function refetch() {
        var now = Date.now();
        if (now - lastFetch < 5000) return;      // never hammer the endpoint
        lastFetch = now;
        img.src = img.src.split("?")[0] + "?r=" + now;
    }

    function tick() {
        left -= 1;
        if (left <= 0) {
            left = rotation;    // new window began — server will render it
            refetch();
        }
        if (txt) txt.textContent = fmt(left);
        setTimeout(tick, 1000);
    }

    if (txt) txt.textContent = fmt(left);
    setTimeout(tick, 1000);
})();
