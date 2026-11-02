/* THE PY-LIB — first-activation badge scanner (member pages)
 *
 * Camera via getUserMedia; QR decoding with the vendored jsQR 1.4.0.
 * Mirrors web/static/js/scan.js but targets the member pages: on a read
 * the raw (AES-sealed) payload goes into a hidden input and the form
 * submits normally, so the server's redirect + Thai flash message simply
 * render — no per-endpoint JS contract to keep in sync.
 */
(function () {
    "use strict";
    var btn = document.getElementById("scanBtn"),
        video = document.getElementById("scanVideo"),
        canvas = document.getElementById("scanCanvas"),
        status = document.getElementById("scanStatus"),
        payload = document.getElementById("scanPayload"),
        form = document.getElementById("scanForm");
    if (!btn || !video || !canvas || !payload || !form) return;

    var stream = null, raf = null, active = false;

    function setStatus(msg, isError) {
        if (!status) return;
        status.textContent = msg;
        status.className = "scan-status" + (isError ? " scan-status-error" : "");
        status.style.display = msg ? "inline" : "none";
    }

    function stop() {
        active = false;
        if (raf) { cancelAnimationFrame(raf); raf = null; }
        if (stream) {
            stream.getTracks().forEach(function (t) { t.stop(); });
            stream = null;
        }
        video.style.display = "none";
        btn.hidden = false;
    }

    function tick() {
        if (!active) return;
        if (video.readyState === video.HAVE_ENOUGH_DATA) {
            var w = video.videoWidth, h = video.videoHeight;
            if (w && h) {
                canvas.width = w;
                canvas.height = h;
                var ctx = canvas.getContext("2d", { willReadFrequently: true });
                ctx.drawImage(video, 0, 0, w, h);
                var frame = ctx.getImageData(0, 0, w, h);
                var code = window.jsQR(frame.data, w, h,
                                       { inversionAttempts: "dontInvert" });
                if (code && code.data) {
                    stop();
                    payload.value = code.data;
                    setStatus("อ่านบัตรได้แล้ว — กำลังตรวจสอบ...", false);
                    form.submit();      // server replies with redirect + flash
                    return;
                }
            }
        }
        raf = requestAnimationFrame(tick);
    }

    btn.addEventListener("click", function () {
        if (active) return;
        if (!window.isSecureContext || !navigator.mediaDevices
                || !navigator.mediaDevices.getUserMedia) {
            setStatus("กล้องใช้ไม่ได้: ต้องเปิดผ่าน localhost หรือ https:// — "
                      + "ใช้รหัสจากผู้ดูแลด้านล่างแทนได้", true);
            return;
        }
        btn.hidden = true;
        video.style.display = "block";
        setStatus("นำบัตรสมาชิกมาที่กล้อง...", false);
        navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } })
            .then(function (s) {
                stream = s;
                active = true;
                video.srcObject = s;
                video.play();
                raf = requestAnimationFrame(tick);
            })
            .catch(function () {
                stop();
                setStatus("ไม่ได้รับอนุญาตให้ใช้กล้อง — ใช้รหัสจากผู้ดูแลด้านล่างแทนได้",
                          true);
            });
    });
})();
