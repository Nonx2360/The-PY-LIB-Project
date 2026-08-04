# core/logger.py — structured logging setup
import logging
import sys

def setup_logger(debug: bool = False) -> logging.Logger:
    level = logging.DEBUG if debug else logging.INFO
    handlers = [logging.StreamHandler(sys.stdout)]
    if debug:
        handlers.append(logging.FileHandler("debug.log", encoding="utf-8"))

    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] (%(threadName)s) %(message)s",
        handlers=handlers,
    )
    return logging.getLogger("PY-LIB")

logger = setup_logger()
