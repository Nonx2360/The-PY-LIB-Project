# main.py - entry point only (creates the app shell and runs it)
import sys

import core.logger as logger_mod
import db.schema
from app import LibraryApp


def main():
    debug = "--debug" in sys.argv

    # Reconfigure logger before any app imports use it
    logger_mod.setup_logger(debug=debug)

    db.schema.init_db()

    app = LibraryApp()
    app.run()


if __name__ == "__main__":
    main()
