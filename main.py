# main.py - entry point only (creates the app shell and runs it)
import sys
import traceback

import core.logger as logger_mod
import db.schema
from app import LibraryApp


def _log_tk_crash(exc, val, tb):
    """sys.excepthook: never let an exception die silently."""
    logger_mod.logger.critical("Uncaught exception", exc_info=(exc, val, tb))
    traceback.print_exception(exc, val, tb)


def _log_versions():
    try:
        import tkinter
        import customtkinter
        logger_mod.logger.info(
            "THE PY-LIB desktop starting: python %s, tkinter %.2f, "
            "customtkinter %s",
            sys.version.split()[0], tkinter.TkVersion,
            customtkinter.__version__)
    except Exception:
        logger_mod.logger.warning("could not resolve runtime versions")


def main():
    debug = "--debug" in sys.argv

    # Reconfigure logger before any app imports use it
    logger_mod.setup_logger(debug=debug)
    _log_versions()

    sys.excepthook = _log_tk_crash

    db.schema.init_db()

    app = LibraryApp()
    app.run()


if __name__ == "__main__":
    main()
