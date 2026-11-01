# core/logger.py — structured logging setup
import logging
import sys

def _console_stream():
    """Return a UTF-8-safe stdout for the console handler.

    A plain Windows console defaults to cp1252/cp874, which cannot encode
    Thai log messages ("หยุดให้บริการ..." etc.). logging swallows the
    UnicodeEncodeError with a stderr "--- Logging error ---" notice, so
    those messages are silently LOST instead of crashing the app.Python
    3.7+ TextIOWrapper.reconfigure() fixes this process-wide; errors=
    "replace" is the belt-and-braces fallback for redirected streams.
    """
    stream = sys.stdout
    try:
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass  # exotic stream objects keep their default behaviour
    return stream

def setup_logger(debug: bool = False) -> logging.Logger:
    level = logging.DEBUG if debug else logging.INFO
    root = logging.getLogger()
    root.setLevel(level)

    # remove old handlers so reconfigure works
    for h in list(root.handlers):
        root.removeHandler(h)

    root.addHandler(logging.StreamHandler(_console_stream()))
    if debug:
        root.addHandler(logging.FileHandler("debug.log", encoding="utf-8", mode="w"))

    fmt = "%(asctime)s [%(levelname)s] (%(threadName)s) %(message)s"
    for h in root.handlers:
        h.setFormatter(logging.Formatter(fmt))

    return logging.getLogger("PY-LIB")

logger = setup_logger()
