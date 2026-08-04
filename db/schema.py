# db/schema.py
import sqlite3
import os
from config.constants import DB_PATH


def seed_admin():
    """Seed default admin on first boot and migrate legacy plain-text passwords to bcrypt."""
    from core.security import hash_password
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("SELECT password FROM admin_users WHERE username = 'admin'")
    if cursor.fetchone() is None:
        hashed_default = hash_password('admin123')
        cursor.execute("INSERT INTO admin_users VALUES (?, ?)", ('admin', hashed_default))

    cursor.execute("SELECT username, password FROM admin_users")
    for username, pwd in cursor.fetchall():
        if not (pwd.startswith('$2a$') or pwd.startswith('$2b$') or pwd.startswith('$2y$')):
            new_hash = hash_password(pwd)
            cursor.execute("UPDATE admin_users SET password = ? WHERE username = ?", (new_hash, username))

    conn.commit()
    conn.close()


def init_db():
    if not os.path.exists('db'):
        os.makedirs('db')
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS admin_users (
            username TEXT PRIMARY KEY,
            password TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS members (
            id INTEGER PRIMARY KEY,
            name TEXT,
            grade TEXT,
            number TEXT,
            register_date TEXT,
            expire_date TEXT,
            qrcode_data TEXT,
            qrcode_path TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS books (
            id INTEGER PRIMARY KEY,
            code TEXT,
            title TEXT,
            status TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS borrow_log (
            id INTEGER PRIMARY KEY,
            member_id INTEGER,
            book_id INTEGER,
            borrow_date TEXT,
            return_due TEXT,
            returned INTEGER,
            FOREIGN KEY (member_id) REFERENCES members (id),
            FOREIGN KEY (book_id) REFERENCES books (id)
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS access_log (
            id INTEGER PRIMARY KEY,
            member_id INTEGER,
            access_time DATETIME,
            action TEXT,
            FOREIGN KEY (member_id) REFERENCES members (id)
        )
    ''')
    
    conn.commit()
    conn.close()

    seed_admin()
