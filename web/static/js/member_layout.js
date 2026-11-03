/* THE PY-LIB — member portal layout: collapsible left sidebar
 *
 * The hamburger button in the portal topbar folds the sidebar away on
 * wide screens (per-device choice remembered in localStorage) and opens
 * it as an overlay drawer on narrow ones (backdrop + Escape close it).
 * Purely presentational: nothing is sent to the server.
 */
(function () {
    "use strict";
    var shell = document.getElementById("memberShell"),
        btn = document.getElementById("memberSideToggle"),
        backdrop = document.getElementById("memberBackdrop"),
        KEY = "pylibMemberSidebar";
    if (!shell || !btn) return;

    function isMobile() { return window.innerWidth <= 900; }

    function stored() {
        try { return localStorage.getItem(KEY); } catch (e) { return null; }
    }

    function store(value) {
        try { localStorage.setItem(KEY, value); } catch (e) { /* private mode */ }
    }

    function sync() {
        var open = !isMobile() ? !shell.classList.contains("side-collapsed")
                               : shell.classList.contains("side-open");
        btn.setAttribute("aria-expanded", open ? "true" : "false");
        if (backdrop) backdrop.classList.toggle("show", isMobile() && open);
    }

    // initial state: CSS keeps the drawer hidden on mobile; on desktop the
    // remembered choice applies (open by default)
    if (!isMobile() && stored() === "collapsed") {
        shell.classList.add("side-collapsed");
    }
    sync();

    btn.addEventListener("click", function () {
        if (isMobile()) {
            shell.classList.toggle("side-open");
        } else {
            var collapsed = shell.classList.toggle("side-collapsed");
            store(collapsed ? "collapsed" : "open");
        }
        sync();
    });

    if (backdrop) backdrop.addEventListener("click", function () {
        shell.classList.remove("side-open");
        sync();
    });

    document.addEventListener("keydown", function (e) {
        if (e.key === "Escape" && isMobile()) {
            shell.classList.remove("side-open");
            sync();
        }
    });

    window.addEventListener("resize", function () {
        // crossing back to desktop always leaves the overlay closed
        if (!isMobile()) shell.classList.remove("side-open");
        sync();
    });
})();
