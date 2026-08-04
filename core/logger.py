# core/logger.py — structured logging setup
import logging
import sys

def setup_logger(debug: bool = False) -> logging.Logger:
    level = logging.DEBUG if debug else logging.INFO
    root = logging.getLogger()
    root.setLevel(level)

    # remove old handlers so reconfigure works
    for h in list(root.handlers):
        root.removeHandler(h)

    root.addHandler(logging.StreamHandler(sys.stdout))
    if debug:
        root.addHandler(logging.FileHandler("debug.log", encoding="utf-8", mode="w"))

    fmt = "%(asctime)s [%(levelname)s] (%(threadName)s) %(message)s"
    for h in root.handlers:
        h.setFormatter(logging.Formatter(fmt))

    return logging.getLogger("PY-LIB")

logger = setup_logger()
