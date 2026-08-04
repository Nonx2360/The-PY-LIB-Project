# main.py - entry point only (creates the app shell and runs it)
import sys
import traceback

import db.schema
from app import LibraryApp


def main():
    # Ensure every table exists and admin is seeded/migrated before launching.
    db.schema.init_db()

    app = LibraryApp()
    app.run()


if __name__ == "__main__":
    main()