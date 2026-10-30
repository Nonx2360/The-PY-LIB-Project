/* THE PY-LIB — member badge scanner for the WebUI
 *
 * Camera via getUserMedia; QR decoding with the vendored jsQR 1.4.0
 * (Apache-2.0, (c) Cosmo Wolfe) in jsQR.js next to this file.
 *
 * The camera only *reads* the QR text: the AES-sealed payload is POSTed to
 * /scan/resolve (session + CSRF required) and decrypted server-side with the
 * same key as the desktop app.
 *
 * Cameras need a secure context: https:// or localhost. On plain LAN HTTP
 * browsers refuse getUserMedia — we show a notice and the manual form still
 * works, so nothing bricks.
 */
(function () {
    "use strict";
    function $(id) { return document.getElementById(id); }
    function csrfToken() {
        var input = document.querySelector('input[name="csrf_token"]');
        return input ? input.value : "";
    }

    function createScanner(opts) {
        var btn = $(opts.buttonId),
            video = $(opts.videoId),
            canvas = $(opts.canvasId),
            status = $(opts.statusId);
        if (!btn || !video || !canvas) return null;

        var stream = null, raf = null, active = false;
        var BASE = opts.endpoint || "/scan/resolve";

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

        function resolvePayload(payload) {
            setStatus("กำลังตรวจสอบบัตร...", false);
            var body = new FormData();
            body.append("payload", payload);
            body.append("csrf_token", csrfToken());
            fetch(BASE, { method: "POST", body: body })
                .then(function (r) {
                    return r.json().then(function (j) { return { code: r.status, json: j }; });
                })
                .then(function (res) {
                    var j = res.json || {};
                    if (j.ok && j.member) {
                        stop();
                        setStatus("", false);
                        opts.onResolve(j.member);
                        return;
                    }
                    var msg = j.error === "not_found"
                        ? "ไม่พบสมาชิกนี้ในระบบ"
                        : j.error === "invalid"
                          ? "QR ไม่ถูกต้อง บัตรเก่า หรือไฟล์โทเคนถูกเปลี่ยน"
                          : "ไม่สามารถอ่านบัตรได้";
                    if (res.code === 401 || res.code === 403) {
                        msg = "เซสชันหมดอายุ กรุณาล็อกอินใหม่";
                    }
                    setStatus(msg, true);
                })
                .catch(function () { setStatus("การเชื่อมต่อล้มเหลว กรุณาลองใหม่", true); });
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
                    var code = window.jsQR(frame.data, w, h, { inversionAttempts: "dontInvert" });
                    if (code && code.data) {
                        stop();
                        resolvePayload(code.data);
                        return;
                    }
                }
            }
            raf = requestAnimationFrame(tick);
        }

        btn.addEventListener("click", function () {
            if (active) return;
            if (!window.isSecureContext || !navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
                setStatus("กล้องใช้ไม่ได้: ต้องเปิดผ่าน localhost หรือ https:// (เมื่อแชร์ผ่านเครือข่ายแบบ http ธรรมดา เบราว์เซอร์จะบล็อกกล้อง) — ใช้การเลือกจากลิสต์ต่อไปได้", true);
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
                    setStatus("ไม่ได้รับอนุญาตให้ใช้กล้อง — ใช้การเลือกจากลิสต์ต่อไปได้", true);
                });
        });
        return { stop: stop };
    }

    window.PyLibScan = { createScanner: createScanner };
})();
