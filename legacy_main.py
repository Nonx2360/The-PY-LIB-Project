import customtkinter as ctk
import sqlite3
import os
import base64
import qrcode
from PIL import Image
from datetime import datetime, timedelta
import uuid
from PIL import ImageTk
from tkinter import filedialog
import cv2
from pyzbar.pyzbar import decode
import threading
import time
import numpy as np
from queue import Queue
from reportlab.lib.pagesizes import A4, A8
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
import io
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
import pandas as pd

# Set appearance mode and default color theme
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class LibraryApp:
    def __init__(self):
        self.app = ctk.CTk()
        self.app.title("ระบบห้องสมุดโรงเรียน")
        self.app.geometry("800x600")
        
        # Create necessary directories
        self.create_directories()
        
        # Initialize database
        self.init_database()
        
        # Start with login window
        self.show_login()
        
    def create_directories(self):
        # Create necessary directories if they don't exist
        directories = [
            'db',
            'assets',
            'assets/qrcodes',
            'assets/cards',
            'assets/logos',
            'assets/fonts'
        ]
        for directory in directories:
            if not os.path.exists(directory):
                os.makedirs(directory)
        
    def init_database(self):
        # Create database directory if it doesn't exist
        if not os.path.exists('db'):
            os.makedirs('db')
            
        # Create QR codes directory if it doesn't exist
        if not os.path.exists('assets/qrcodes'):
            os.makedirs('assets/qrcodes')
            
        # Connect to SQLite database
        self.conn = sqlite3.connect('db/library.db')
        self.cursor = self.conn.cursor()
        
        # Create tables
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS admin_users (
                username TEXT PRIMARY KEY,
                password TEXT
            )
        ''')
        
        self.cursor.execute('''
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
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS books (
                id INTEGER PRIMARY KEY,
                code TEXT,
                title TEXT,
                status TEXT
            )
        ''')
        
        self.cursor.execute('''
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
        
        # Create access log table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS access_log (
                id INTEGER PRIMARY KEY,
                member_id INTEGER,
                access_time DATETIME,
                action TEXT,
                FOREIGN KEY (member_id) REFERENCES members (id)
            )
        ''')
        
        # Insert default admin user if not exists (with bcrypt hash)
        from core.security import hash_password
        self.cursor.execute("SELECT password FROM admin_users WHERE username = 'admin'")
        admin_row = self.cursor.fetchone()
        if not admin_row:
            hashed_default = hash_password('admin123')
            self.cursor.execute("INSERT INTO admin_users VALUES (?, ?)", ('admin', hashed_default))
        
        # Check and migrate any plain-text passwords to bcrypt hashes
        self.cursor.execute("SELECT username, password FROM admin_users")
        all_users = self.cursor.fetchall()
        for username, pwd in all_users:
            if not (pwd.startswith('$2a$') or pwd.startswith('$2b$') or pwd.startswith('$2y$')):
                new_hash = hash_password(pwd)
                self.cursor.execute("UPDATE admin_users SET password = ? WHERE username = ?", (new_hash, username))
            
        self.conn.commit()
        
    def show_login(self):
        # Clear any existing widgets
        for widget in self.app.winfo_children():
            widget.destroy()
            
        # Create login frame
        login_frame = ctk.CTkFrame(self.app)
        login_frame.pack(pady=20, padx=40, fill="both", expand=True)
        
        # Login title
        title_label = ctk.CTkLabel(login_frame, text="ระบบห้องสมุดโรงเรียน", font=("Helvetica", 24, "bold"))
        title_label.pack(pady=20)
        
        # Create form frame
        form_frame = ctk.CTkFrame(login_frame)
        form_frame.pack(pady=20, padx=40, fill="x")
        
        # Username entry
        username_label = ctk.CTkLabel(form_frame, text="ชื่อผู้ใช้:", font=("Helvetica", 14))
        username_label.pack(pady=(10,0))
        username_entry = ctk.CTkEntry(form_frame, placeholder_text="กรอกชื่อผู้ใช้", width=300)
        username_entry.pack(pady=5)
        
        # Password entry
        password_label = ctk.CTkLabel(form_frame, text="รหัสผ่าน:", font=("Helvetica", 14))
        password_label.pack(pady=(10,0))
        password_entry = ctk.CTkEntry(form_frame, placeholder_text="กรอกรหัสผ่าน", show="*", width=300)
        password_entry.pack(pady=5)
        
        # Error message label (initially empty)
        self.login_error_label = ctk.CTkLabel(form_frame, text="", text_color="red")
        self.login_error_label.pack(pady=5)
        
        # Login button
        login_button = ctk.CTkButton(
            form_frame, 
            text="เข้าสู่ระบบ",
            width=200,
            height=40,
            font=("Helvetica", 14, "bold"),
            command=lambda: self.login(username_entry.get(), password_entry.get())
        )
        login_button.pack(pady=20)
        
        # Bind Enter key to login
        username_entry.bind('<Return>', lambda e: password_entry.focus())
        password_entry.bind('<Return>', lambda e: self.login(username_entry.get(), password_entry.get()))
        
        # Set focus to username entry
        username_entry.focus()

    def login(self, username, password):
        if not username or not password:
            self.login_error_label.configure(text="กรุณากรอกชื่อผู้ใช้และรหัสผ่าน")
            return
            
        try:
            self.cursor.execute("SELECT password FROM admin_users WHERE username = ?", (username,))
            row = self.cursor.fetchone()
            from core.security import verify_password
            if row and verify_password(password, row[0]):
                self.show_dashboard()
            else:
                self.login_error_label.configure(text="ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")
        except sqlite3.Error as e:
            self.login_error_label.configure(text=f"เกิดข้อผิดพลาด: {str(e)}")
            
    def show_dashboard(self):
        # Clear any existing widgets
        for widget in self.app.winfo_children():
            widget.destroy()
            
        # Create main container with padding
        main_frame = ctk.CTkFrame(self.app, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Header section with logo and title
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))
        
        # Title with modern font and color
        title_label = ctk.CTkLabel(
            header_frame,
            text="ระบบห้องสมุดโรงเรียน",
            font=("Sarabun-Bold", 32),
            text_color="#1f538d"  # สีน้ำเงินเดิม
        )
        title_label.pack(pady=10)
        
        # Create grid layout for buttons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="both", expand=True)
        
        # Configure grid layout
        button_frame.grid_columnconfigure((0,1,2), weight=1)
        button_frame.grid_rowconfigure((0,1,2,3), weight=1)
        
        # Define buttons with their positions
        buttons = [
            ("จัดการสมาชิก", self.show_member_management, 0, 0),
            ("จัดการหนังสือ", self.show_book_management, 0, 1),
            ("ยืมหนังสือ", self.show_borrow, 0, 2),
            ("คืนหนังสือ", self.show_return, 1, 0),
            ("ประวัติการยืม-คืน", self.show_history, 1, 1),
            ("บันทึกเข้าออก", self.show_access_scanner, 1, 2),
            ("ประวัติเข้าออก", self.show_access_history, 2, 0),
            ("ตั้งค่า", self.show_settings, 2, 1),
            ("เกี่ยวกับ", self.show_about, 2, 2)
        ]
        
        # Create and place buttons
        for text, command, row, col in buttons:
            btn = ctk.CTkButton(
                button_frame,
                text=text,
                command=command,
                height=80,  # Taller buttons for touch
                font=("Sarabun", 20, "bold"),
                fg_color="#1f538d",  # สีน้ำเงินเดิม
                hover_color="#14375e",  # สีน้ำเงินเข้มเมื่อ hover
                corner_radius=15,
                border_width=2,
                border_color="#14375e"
            )
            btn.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        
        # Add responsive behavior
        def on_resize(event):
            # Adjust font size based on window width
            width = event.width
            if width < 800:
                font_size = 16
            elif width < 1200:
                font_size = 20
            else:
                font_size = 24
            
            # Update all button fonts
            for child in button_frame.winfo_children():
                if isinstance(child, ctk.CTkButton):
                    child.configure(font=("Sarabun", font_size, "bold"))
        
        # Bind resize event
        self.app.bind("<Configure>", on_resize)
        
    def show_member_management(self):
        # Clear any existing widgets
        for widget in self.app.winfo_children():
            widget.destroy()
            
        # Create member management frame
        member_frame = ctk.CTkFrame(self.app)
        member_frame.pack(pady=20, padx=40, fill="both", expand=True)
        
        # Title
        title_label = ctk.CTkLabel(member_frame, text="จัดการสมาชิก", font=("Helvetica", 24))
        title_label.pack(pady=20)
        
        # Create form frame
        form_frame = ctk.CTkFrame(member_frame)
        form_frame.pack(pady=10, padx=20, fill="x")
        
        # Form fields
        name_entry = ctk.CTkEntry(form_frame, placeholder_text="ชื่อ-นามสกุล")
        name_entry.pack(pady=5, padx=20, fill="x")
        
        grade_entry = ctk.CTkEntry(form_frame, placeholder_text="ชั้น")
        grade_entry.pack(pady=5, padx=20, fill="x")
        
        number_entry = ctk.CTkEntry(form_frame, placeholder_text="เลขที่")
        number_entry.pack(pady=5, padx=20, fill="x")
        
        # Add member button
        add_button = ctk.CTkButton(form_frame, text="เพิ่มสมาชิก", 
                                 command=lambda: self.add_member(
                                     name_entry.get(),
                                     grade_entry.get(),
                                     number_entry.get()
                                 ))
        add_button.pack(pady=10)
        
        # Back button
        back_button = ctk.CTkButton(member_frame, text="กลับ", 
                                  command=self.show_dashboard)
        back_button.pack(pady=10)
        
        # Display existing members
        self.display_members(member_frame)
        
    def add_member(self, name, grade, number):
        if not all([name, grade, number]):
            self.show_error("กรุณากรอกข้อมูลให้ครบ")
            return
            
        # Generate QR code data
        register_date = datetime.now().strftime("%Y-%m-%d")
        expire_date = (datetime.now() + timedelta(days=365)).strftime("%Y-%m-%d")
        qr_data = f"{name}|{grade}|{number}|{register_date}|{expire_date}"
        from core.security import encrypt_qr_data
        encoded_data = encrypt_qr_data(qr_data)
        
        # Generate QR code
        qr = qrcode.QRCode(version=1, box_size=10, border=5)
        qr.add_data(encoded_data)
        qr.make(fit=True)
        if hasattr(qr, 'get_image'):
            pil_image = qr.get_image()
        else:
            pil_image = qr.make_image()
        
        # Save QR code
        qr_filename = f"assets/qrcodes/{uuid.uuid4()}.png"
        pil_image.save(qr_filename)
        
        # Add to database
        try:
            self.cursor.execute('''
                INSERT INTO members (name, grade, number, register_date, expire_date, qrcode_data, qrcode_path)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (name, grade, number, register_date, expire_date, encoded_data, qr_filename))
            self.conn.commit()
            
            # Show QR code window
            self.show_qr_code_window(pil_image, name, grade, number, register_date, expire_date)
            
            # Generate member card PDF
            member_id = self.cursor.lastrowid
            card_pdf_path = f"assets/cards/member_{member_id}.pdf"
            self.generate_member_card_pdf(name, number, "assets/logos/school_logo.png", card_pdf_path, encoded_data)
            
            # Refresh member list
            self.show_member_management()
            
        except sqlite3.Error as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")
            
    def show_qr_code_window(self, qr_image, name, grade, number, register_date, expire_date):
        # Create new window
        qr_window = ctk.CTkToplevel(self.app)
        qr_window.title("QR Code บัตรสมาชิก")
        qr_window.geometry("400x600")
        
        # Create frame
        qr_frame = ctk.CTkFrame(qr_window)
        qr_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # Member info
        info_text = f"""
ชื่อ-นามสกุล: {name}
ชั้น: {grade}
เลขที่: {number}
วันที่สมัคร: {register_date}
วันหมดอายุ: {expire_date}
        """
        info_label = ctk.CTkLabel(qr_frame, text=info_text, font=("Helvetica", 14))
        info_label.pack(pady=10)
        
        # Convert QR code to PhotoImage
        if hasattr(qr_image, 'get_image'):
            pil_image = qr_image.get_image()
        else:
            pil_image = qr_image
        qr_ctk_image = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(256, 256))
        qr_label = ctk.CTkLabel(qr_frame, image=qr_ctk_image, text="")
        qr_label.image = qr_ctk_image
        qr_label.pack(pady=10)
        
        # Save button
        save_button = ctk.CTkButton(qr_frame, text="บันทึก QR Code", 
                                  command=lambda: self.save_qr_code(qr_image, name))
        save_button.pack(pady=10)
        
        # Close button
        close_button = ctk.CTkButton(qr_frame, text="ปิด", 
                                   command=qr_window.destroy)
        close_button.pack(pady=10)
        
    def save_qr_code(self, qr_image, name):
        # Create save dialog
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png")],
            initialfile=f"qrcode_{name}.png"
        )
        
        if file_path:
            qr_image.save(file_path)
            self.show_success(f"บันทึก QR Code เรียบร้อยแล้ว: {file_path}")
            
    def display_members(self, parent_frame):
        # Create scrollable frame for members
        scroll_frame = ctk.CTkScrollableFrame(parent_frame)
        scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Get all members
        self.cursor.execute("SELECT * FROM members ORDER BY id DESC")
        members = self.cursor.fetchall()
        
        # Display each member
        for member in members:
            member_frame = ctk.CTkFrame(scroll_frame)
            member_frame.pack(pady=5, padx=10, fill="x")
            
            # Member info with complete details
            info_text = f"""
ชื่อ-นามสกุล: {member[1]}
ชั้น: {member[2]}
เลขที่: {member[3]}
วันที่สมัคร: {member[4]}
วันหมดอายุ: {member[5]}
            """
            info_label = ctk.CTkLabel(member_frame, text=info_text, justify="left")
            info_label.pack(side="left", padx=10)
            
            # Buttons frame
            button_frame = ctk.CTkFrame(member_frame)
            button_frame.pack(side="right", padx=5)
            
            # View QR Code button
            view_qr_button = ctk.CTkButton(button_frame, text="ดู QR Code", 
                                         command=lambda m=member: self.view_member_qr(m))
            view_qr_button.pack(side="left", padx=2)
            
            # View Card button
            view_card_button = ctk.CTkButton(button_frame, text="ดูบัตรสมาชิก", 
                                           command=lambda m=member: self.view_member_card(m))
            view_card_button.pack(side="left", padx=2)
            
            # Delete button
            delete_button = ctk.CTkButton(button_frame, text="ลบ", 
                                        command=lambda m=member: self.delete_member(m))
            delete_button.pack(side="left", padx=2)
            
    def view_member_qr(self, member):
        try:
            # Load QR code image
            qr_image = Image.open(member[7])
            
            # Show QR code window
            self.show_qr_code_window(qr_image, member[1], member[2], member[3], 
                                   member[4], member[5])
        except Exception as e:
            self.show_error(f"ไม่สามารถเปิด QR Code ได้: {str(e)}")
            
    def view_member_card(self, member):
        try:
            card_path = os.path.abspath(f"assets/cards/member_{member[0]}.pdf")
            print("PDF path:", card_path)
            print("File exists:", os.path.exists(card_path))
            if os.path.exists(card_path):
                if os.name == 'nt':  # Windows
                    os.startfile(card_path)
                else:
                    import subprocess
                    subprocess.run(['xdg-open', card_path])
            else:
                # If card doesn't exist, generate it
                try:
                    # Generate QR code
                    qr_data = member[6]  # qrcode_data from database
                    qr = qrcode.make(qr_data)
                    qr_bytes = io.BytesIO()
                    qr.save(qr_bytes, format='PNG')
                    qr_bytes.seek(0)

                    # Generate member card
                    self.generate_member_card_pdf(
                        member[1],  # name
                        member[3],  # number
                        "assets/logos/school_logo.png",  # logo path
                        card_path,
                        qr_data
                    )
                    
                    # Open the newly generated card
                    if os.name == 'nt':  # Windows
                        os.startfile(card_path)
                    else:  # Linux/Mac
                        import subprocess
                        subprocess.run(['xdg-open', card_path])
                except Exception as e:
                    self.show_error(f"ไม่สามารถสร้างบัตรสมาชิกได้: {str(e)}")
        except Exception as e:
            self.show_error(f"ไม่สามารถเปิดบัตรสมาชิกได้: {str(e)}")
            
    def delete_member(self, member):
        try:
            # Delete QR code file
            if os.path.exists(member[7]):  # qrcode_path
                os.remove(member[7])
                
            # Delete from database
            self.cursor.execute("DELETE FROM members WHERE id = ?", (member[0],))
            self.conn.commit()
            
            self.show_success("ลบสมาชิกสำเร็จ")
            self.show_member_management()  # Refresh the view
        except Exception as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")
            
    def show_error(self, message):
        error_label = ctk.CTkLabel(self.app, text=message, text_color="red")
        error_label.pack(pady=10)
        self.app.after(2000, error_label.destroy)
        
    def show_success(self, message):
        success_label = ctk.CTkLabel(self.app, text=message, text_color="green")
        success_label.pack(pady=10)
        self.app.after(2000, success_label.destroy)
        
    def show_book_management(self):
        # Clear any existing widgets
        for widget in self.app.winfo_children():
            widget.destroy()
            
        # Create book management frame
        book_frame = ctk.CTkFrame(self.app)
        book_frame.pack(pady=20, padx=40, fill="both", expand=True)
        
        # Title
        title_label = ctk.CTkLabel(book_frame, text="จัดการหนังสือ", font=("Helvetica", 24))
        title_label.pack(pady=20)
        
        # Create form frame
        form_frame = ctk.CTkFrame(book_frame)
        form_frame.pack(pady=10, padx=20, fill="x")
        
        # Form fields
        code_entry = ctk.CTkEntry(form_frame, placeholder_text="รหัสหนังสือ")
        code_entry.pack(pady=5, padx=20, fill="x")
        
        title_entry = ctk.CTkEntry(form_frame, placeholder_text="ชื่อเรื่อง")
        title_entry.pack(pady=5, padx=20, fill="x")
        
        # Add book button
        add_button = ctk.CTkButton(form_frame, text="เพิ่มหนังสือ", 
                                 command=lambda: self.add_book(
                                     code_entry.get(),
                                     title_entry.get()
                                 ))
        add_button.pack(pady=10)
        
        # Import Excel button
        import_button = ctk.CTkButton(form_frame, text="นำเข้าจาก Excel", 
                                    command=self.import_books_from_excel)
        import_button.pack(pady=10)
        
        # Export template button
        template_button = ctk.CTkButton(form_frame, text="ดาวน์โหลดแม่แบบ Excel", 
                                      command=self.export_excel_template)
        template_button.pack(pady=10)
        
        # Back button
        back_button = ctk.CTkButton(book_frame, text="กลับ", 
                                  command=self.show_dashboard)
        back_button.pack(pady=10)
        
        # Display existing books
        self.display_books(book_frame)
        
    def add_book(self, code, title):
        if not all([code, title]):
            self.show_error("กรุณากรอกข้อมูลให้ครบ")
            return
            
        try:
            self.cursor.execute('''
                INSERT INTO books (code, title, status)
                VALUES (?, ?, ?)
            ''', (code, title, "ว่าง"))
            self.conn.commit()
            self.show_success("เพิ่มหนังสือสำเร็จ")
            self.show_book_management()  # Refresh the view
        except sqlite3.Error as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")
            
    def display_books(self, parent_frame):
        # Create scrollable frame for books
        scroll_frame = ctk.CTkScrollableFrame(parent_frame)
        scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Get all books
        self.cursor.execute("SELECT * FROM books ORDER BY id DESC")
        books = self.cursor.fetchall()
        
        # Display each book
        for book in books:
            book_frame = ctk.CTkFrame(scroll_frame)
            book_frame.pack(pady=5, padx=10, fill="x")
            
            # Book info
            info_text = f"รหัส: {book[1]} | ชื่อ: {book[2]} | สถานะ: {book[3]}"
            info_label = ctk.CTkLabel(book_frame, text=info_text)
            info_label.pack(side="left", padx=10)
            
            # Delete button
            delete_button = ctk.CTkButton(book_frame, text="ลบ", 
                                        command=lambda b=book: self.delete_book(b))
            delete_button.pack(side="right", padx=5)
            
    def delete_book(self, book):
        try:
            # Check if book is borrowed
            self.cursor.execute("SELECT * FROM borrow_log WHERE book_id = ? AND returned = 0", (book[0],))
            if self.cursor.fetchone():
                self.show_error("ไม่สามารถลบหนังสือที่กำลังถูกยืมอยู่ได้")
                return

            # Show confirmation dialog
            confirm = self.show_confirm_dialog(
                "ยืนยันการลบ",
                f"คุณต้องการลบหนังสือ '{book[2]}' ใช่หรือไม่?"
            )
            
            if not confirm:
                return
                
            # Delete from database
            self.cursor.execute("DELETE FROM books WHERE id = ?", (book[0],))
            self.conn.commit()
            
            self.show_success("ลบหนังสือสำเร็จ")
            self.show_book_management()  # Refresh the view
        except Exception as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")

    def show_confirm_dialog(self, title, message):
        dialog = ctk.CTkToplevel(self.app)
        dialog.title(title)
        dialog.geometry("400x200")
        dialog.grab_set()  # Make dialog modal
        
        # Center dialog on screen
        dialog.update_idletasks()
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        x = (screen_width - 400) // 2
        y = (screen_height - 200) // 2
        dialog.geometry(f"400x200+{x}+{y}")
        
        # Message
        label = ctk.CTkLabel(dialog, text=message, font=("Sarabun", 14))
        label.pack(pady=20, padx=20)
        
        # Variable to store result
        result = {'value': False}
        
        def on_yes():
            result['value'] = True
            dialog.destroy()
            
        def on_no():
            result['value'] = False
            dialog.destroy()
        
        # Buttons frame
        button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        button_frame.pack(pady=20)
        
        # Yes button
        yes_button = ctk.CTkButton(
            button_frame,
            text="ใช่",
            command=on_yes,
            fg_color="#FF4B4B",
            hover_color="#FF3333"
        )
        yes_button.pack(side="left", padx=10)
        
        # No button
        no_button = ctk.CTkButton(
            button_frame,
            text="ไม่",
            command=on_no
        )
        no_button.pack(side="left", padx=10)
        
        # Wait for dialog to close
        dialog.wait_window()
        
        return result['value']
        
    def show_borrow(self):
        # Clear any existing widgets
        for widget in self.app.winfo_children():
            widget.destroy()
            
        # Create borrow frame
        borrow_frame = ctk.CTkFrame(self.app)
        borrow_frame.pack(pady=20, padx=40, fill="both", expand=True)
        
        # Title
        title_label = ctk.CTkLabel(borrow_frame, text="ยืมหนังสือ", font=("Helvetica", 24))
        title_label.pack(pady=20)
        
        # Member info frame
        member_frame = ctk.CTkFrame(borrow_frame)
        member_frame.pack(pady=10, padx=20, fill="x")
        
        # Status display
        self.borrow_status_label = ctk.CTkLabel(member_frame, text="กรุณาสแกนบัตรสมาชิก", font=("Helvetica", 18))
        self.borrow_status_label.pack(pady=10)
        
        # Member info display
        self.member_info_label = ctk.CTkLabel(member_frame, text="", font=("Helvetica", 14))
        self.member_info_label.pack(pady=5)
        
        # Scan button
        scan_button = ctk.CTkButton(member_frame, text="สแกนบัตร", 
                                  command=lambda: self.start_borrow_scan(),
                                  height=60,
                                  font=("Helvetica", 20))
        scan_button.pack(pady=20)
        
        # Book search frame
        book_frame = ctk.CTkFrame(borrow_frame)
        book_frame.pack(pady=10, padx=20, fill="x")
        
        # Book search entry
        self.book_code_entry = ctk.CTkEntry(book_frame, placeholder_text="กรอกรหัสหนังสือ")
        self.book_code_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # Search button
        search_button = ctk.CTkButton(book_frame, text="ค้นหา", 
                                    command=self.search_book,
                                    width=100)
        search_button.pack(side="right", padx=5)
        
        # Book info frame
        self.book_info_frame = ctk.CTkFrame(borrow_frame)
        self.book_info_frame.pack(pady=10, padx=20, fill="x")
        
        # Book info label (initially empty)
        self.book_info_label = ctk.CTkLabel(self.book_info_frame, text="", font=("Helvetica", 14))
        self.book_info_label.pack(pady=10)

        # Due date frame
        self.due_date_frame = ctk.CTkFrame(self.book_info_frame)
        self.due_date_frame.pack(pady=10, fill="x")

        # Due date label
        due_date_label = ctk.CTkLabel(self.due_date_frame, text="กำหนดคืน:", font=("Helvetica", 14))
        due_date_label.pack(side="left", padx=5)

        # Default due date (7 days from now)
        default_due_date = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")
        self.due_date_entry = ctk.CTkEntry(self.due_date_frame, width=120)
        self.due_date_entry.pack(side="left", padx=5)
        self.due_date_entry.insert(0, default_due_date)

        # Calendar button
        calendar_button = ctk.CTkButton(self.due_date_frame, 
                                      text="📅", 
                                      width=40,
                                      command=lambda: DatePicker(self.app, self.due_date_entry).show())
        calendar_button.pack(side="left", padx=5)

        # Quick buttons frame
        quick_buttons_frame = ctk.CTkFrame(self.due_date_frame)
        quick_buttons_frame.pack(side="left", padx=5)

        # Quick selection buttons
        ctk.CTkButton(quick_buttons_frame, 
                     text="7 วัน",
                     width=60,
                     command=lambda: self.set_quick_date(7)).pack(side="left", padx=2)
        
        ctk.CTkButton(quick_buttons_frame,
                     text="14 วัน",
                     width=60,
                     command=lambda: self.set_quick_date(14)).pack(side="left", padx=2)

        # Due date info
        due_date_info = ctk.CTkLabel(self.due_date_frame, 
                                   text="(รูปแบบ: YYYY-MM-DD)", 
                                   font=("Helvetica", 12),
                                   text_color="gray")
        due_date_info.pack(side="left", padx=5)
        
        # Borrow button (initially hidden)
        self.borrow_button = ctk.CTkButton(self.book_info_frame, text="ยืมหนังสือ", 
                                         command=self.process_borrow,
                                         font=("Helvetica", 16))
        
        # Back button
        back_button = ctk.CTkButton(borrow_frame, text="กลับ", 
                                  command=self.show_dashboard)
        back_button.pack(pady=10)
        
        # Store current member and book
        self.current_member = None
        self.current_book = None
        
        # Bind Enter key to search
        self.book_code_entry.bind('<Return>', lambda e: self.search_book())

    def start_borrow_scan(self):
        # Create new window for scanning
        scan_window = ctk.CTkToplevel(self.app)
        scan_window.title("สแกน QR Code")
        scan_window.geometry("800x600")
        
        # Create frame for video
        video_frame = ctk.CTkFrame(scan_window)
        video_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Create label for video
        video_label = ctk.CTkLabel(video_frame, text="")
        video_label.pack(fill="both", expand=True)
        
        # Add instruction label
        instruction_label = ctk.CTkLabel(scan_window, 
            text="นำ QR Code มาวางตรงกล้อง\nรอสักครู่ระบบจะสแกนอัตโนมัติ",
            font=("Helvetica", 16))
        instruction_label.pack(pady=10)
        
        # Add status label
        status_label = ctk.CTkLabel(scan_window,
            text="กำลังรอสแกน...",
            font=("Helvetica", 14),
            text_color="yellow")
        status_label.pack(pady=5)
        
        # Start webcam
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            self.show_error("ไม่สามารถเปิดกล้องได้")
            scan_window.destroy()
            return
            
        # Set camera resolution
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        
        # Create a queue for thread communication
        qr_queue = Queue()
        
        def scan_qr():
            while True:
                try:
                    ret, frame = cap.read()
                    if not ret:
                        print("Failed to grab frame")
                        break
                        
                    # Resize frame for better display
                    frame = cv2.resize(frame, (640, 480))
                    
                    # Convert frame to grayscale
                    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                    
                    # Apply threshold to make QR code more visible
                    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                    
                    # Try to decode QR code from both original and thresholded image
                    decoded_objects = decode(frame)
                    if not decoded_objects:
                        decoded_objects = decode(thresh)
                    
                    for obj in decoded_objects:
                        try:
                            # Draw rectangle around QR code
                            points = obj.polygon
                            if len(points) > 4:
                                hull = cv2.convexHull(np.array([point for point in points], dtype=np.float32))
                                cv2.polylines(frame, [hull], True, (0, 255, 0), 2)
                            else:
                                cv2.polylines(frame, [np.array(points, dtype=np.int32)], True, (0, 255, 0), 2)
                            
                            # Decode QR data
                            qr_data = obj.data.decode('utf-8')
                            from core.security import decrypt_qr_data
                            decoded_data = decrypt_qr_data(qr_data)
                            name, grade, number, register_date, expire_date = decoded_data.split("|")
                            
                            # Update status
                            status_label.configure(text="✓ สแกนสำเร็จ!", text_color="green")
                            
                            # Put data in queue for main thread to process
                            qr_queue.put((name, grade, number))
                            
                            # Close scan window after a short delay
                            scan_window.after(1000, lambda: [scan_window.destroy(), cap.release()])
                            return
                            
                        except Exception as e:
                            print(f"Error decoding QR: {str(e)}")
                            status_label.configure(text="❌ ไม่สามารถอ่าน QR Code ได้", text_color="red")
                    
                    # Convert frame to CTkImage
                    frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    frame_pil = Image.fromarray(frame_rgb)
                    frame_ctk = ctk.CTkImage(light_image=frame_pil, dark_image=frame_pil, size=(640, 480))
                    
                    # Update video label
                    video_label.configure(image=frame_ctk)
                    
                except Exception as e:
                    print(f"Error in scan loop: {str(e)}")
                    break
                    
                # Check if window is closed
                if not scan_window.winfo_exists():
                    break
                    
                time.sleep(0.03)  # 30 FPS
                
            # Clean up
            cap.release()
                
        # Start scanning in separate thread
        scan_thread = threading.Thread(target=scan_qr)
        scan_thread.daemon = True
        scan_thread.start()
        
        def process_qr_data():
            try:
                if not qr_queue.empty():
                    name, grade, number = qr_queue.get()
                    self.cursor.execute("""
                        SELECT id FROM members 
                        WHERE name = ? AND grade = ? AND number = ?
                    """, (name, grade, number))
                    member = self.cursor.fetchone()
                    if member:
                        self.current_member = {
                            'id': member[0],
                            'name': name,
                            'grade': grade,
                            'number': number
                        }
                        # Update status and member info
                        self.borrow_status_label.configure(
                            text="พบข้อมูลสมาชิก",
                            text_color="green"
                        )
                        self.member_info_label.configure(
                            text=f"สมาชิก: {name}\nชั้น: {grade}\nเลขที่: {number}"
                        )
                        # Enable book search
                        self.book_code_entry.configure(state="normal")
                        self.book_code_entry.focus()
                    else:
                        self.borrow_status_label.configure(
                            text="ไม่พบข้อมูลสมาชิกในระบบ",
                            text_color="red"
                        )
                if scan_window.winfo_exists():
                    scan_window.after(100, process_qr_data)
            except Exception as e:
                print(f"Error processing QR data: {str(e)}")
        
        # Start processing QR data
        scan_window.after(100, process_qr_data)
        
        # Handle window close
        def on_closing():
            cap.release()
            scan_window.destroy()
            
        scan_window.protocol("WM_DELETE_WINDOW", on_closing)

    def search_book(self):
        if not self.current_member:
            self.show_error("กรุณาสแกนบัตรสมาชิกก่อน")
            return
            
        book_code = self.book_code_entry.get().strip()
        if not book_code:
            self.show_error("กรุณากรอกรหัสหนังสือ")
            return
            
        try:
            self.cursor.execute("""
                SELECT id, code, title, status
                FROM books
                WHERE code = ?
            """, (book_code,))
            book = self.cursor.fetchone()
            
            if not book:
                self.book_info_label.configure(text="ไม่พบหนังสือในระบบ")
                self.current_book = None
                self.borrow_button.pack_forget()
                return
                
            if book[3] != "ว่าง":
                self.book_info_label.configure(text=f"หนังสือรหัส {book_code} ไม่พร้อมให้ยืม\nสถานะ: {book[3]}")
                self.current_book = None
                self.borrow_button.pack_forget()
                return
                
            # Show book info and enable borrow button
            self.book_info_label.configure(text=f"หนังสือที่พบ:\nรหัส: {book[1]}\nชื่อ: {book[2]}\nสถานะ: {book[3]}")
            self.current_book = book
            self.borrow_button.pack(pady=10)
            
        except Exception as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")

    def process_borrow(self):
        if not self.current_member or not self.current_book:
            self.show_error("ข้อมูลไม่ครบถ้วน")
            return

        # Validate due date
        try:
            due_date = self.due_date_entry.get().strip()
            if not due_date:
                self.show_error("กรุณากำหนดวันคืนหนังสือ")
                return

            # Validate date format
            datetime.strptime(due_date, "%Y-%m-%d")

            # Check if due date is in the past
            if datetime.strptime(due_date, "%Y-%m-%d").date() < datetime.now().date():
                self.show_error("กำหนดคืนต้องไม่เป็นวันที่ผ่านมาแล้ว")
                return

        except ValueError:
            self.show_error("รูปแบบวันที่ไม่ถูกต้อง กรุณาใช้รูปแบบ YYYY-MM-DD")
            return
            
        try:
            # Check if member has overdue books
            self.cursor.execute("""
                SELECT COUNT(*) FROM borrow_log
                WHERE member_id = ? AND returned = 0 AND return_due < date('now')
            """, (self.current_member['id'],))
            overdue_count = self.cursor.fetchone()[0]
            
            if overdue_count > 0:
                self.show_error("ไม่สามารถยืมได้ เนื่องจากมีหนังสือค้างส่ง")
                return
                
            # Check if member has reached borrow limit
            self.cursor.execute("""
                SELECT COUNT(*) FROM borrow_log
                WHERE member_id = ? AND returned = 0
            """, (self.current_member['id'],))
            current_borrows = self.cursor.fetchone()[0]
            
            if current_borrows >= 3:  # Maximum 3 books per member
                self.show_error("ไม่สามารถยืมได้ เนื่องจากยืมครบจำนวนที่กำหนด (3 เล่ม)")
                return
                
            # Process borrow
            borrow_date = datetime.now().strftime("%Y-%m-%d")
            
            self.cursor.execute("""
                INSERT INTO borrow_log (member_id, book_id, borrow_date, return_due, returned)
                VALUES (?, ?, ?, ?, 0)
            """, (self.current_member['id'], self.current_book[0], borrow_date, due_date))
            
            self.cursor.execute("""
                UPDATE books SET status = 'ยืมแล้ว'
                WHERE id = ?
            """, (self.current_book[0],))
            
            self.conn.commit()
            
            # Show success message
            self.show_success(f"""ยืมหนังสือสำเร็จ
สมาชิก: {self.current_member['name']}
หนังสือ: {self.current_book[2]}
วันที่ยืม: {borrow_date}
กำหนดคืน: {due_date}""")
            
            # Clear book info
            self.book_code_entry.delete(0, 'end')
            self.book_info_label.configure(text="")
            self.current_book = None
            self.borrow_button.pack_forget()
            
        except Exception as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")

    def show_return(self):
        # Clear any existing widgets
        for widget in self.app.winfo_children():
            widget.destroy()
            
        # Create return frame
        return_frame = ctk.CTkFrame(self.app)
        return_frame.pack(pady=20, padx=40, fill="both", expand=True)
        
        # Title
        title_label = ctk.CTkLabel(return_frame, text="คืนหนังสือ", font=("Helvetica", 24))
        title_label.pack(pady=20)
        
        # Member info frame
        member_frame = ctk.CTkFrame(return_frame)
        member_frame.pack(pady=10, padx=20, fill="x")
        
        # Status display
        self.return_status_label = ctk.CTkLabel(member_frame, text="กรุณาสแกนบัตรสมาชิก", font=("Helvetica", 18))
        self.return_status_label.pack(pady=10)
        
        # Member info display
        self.return_member_info_label = ctk.CTkLabel(member_frame, text="", font=("Helvetica", 14))
        self.return_member_info_label.pack(pady=5)
        
        # Scan button
        scan_button = ctk.CTkButton(member_frame, text="สแกนบัตร", 
                                  command=lambda: self.start_return_scan(),
                                  height=60,
                                  font=("Helvetica", 20))
        scan_button.pack(pady=20)
        
        # Create borrowed books frame
        self.borrowed_books_frame = ctk.CTkScrollableFrame(return_frame)
        self.borrowed_books_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Back button
        back_button = ctk.CTkButton(return_frame, text="กลับ", 
                                  command=self.show_dashboard)
        back_button.pack(pady=10)

    def start_return_scan(self):
        # Create new window for scanning
        scan_window = ctk.CTkToplevel(self.app)
        scan_window.title("สแกน QR Code")
        scan_window.geometry("800x600")
        
        # Create frame for video
        video_frame = ctk.CTkFrame(scan_window)
        video_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Create label for video
        video_label = ctk.CTkLabel(video_frame, text="")
        video_label.pack(fill="both", expand=True)
        
        # Add instruction label
        instruction_label = ctk.CTkLabel(scan_window, 
            text="นำ QR Code มาวางตรงกล้อง\nรอสักครู่ระบบจะสแกนอัตโนมัติ",
            font=("Helvetica", 16))
        instruction_label.pack(pady=10)
        
        # Add status label
        status_label = ctk.CTkLabel(scan_window,
            text="กำลังรอสแกน...",
            font=("Helvetica", 14),
            text_color="yellow")
        status_label.pack(pady=5)
        
        # Start webcam
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            self.show_error("ไม่สามารถเปิดกล้องได้")
            scan_window.destroy()
            return
            
        # Set camera resolution
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        
        # Create a queue for thread communication
        qr_queue = Queue()
        
        def scan_qr():
            while True:
                try:
                    ret, frame = cap.read()
                    if not ret:
                        print("Failed to grab frame")
                        break
                        
                    # Resize frame for better display
                    frame = cv2.resize(frame, (640, 480))
                    
                    # Convert frame to grayscale
                    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                    
                    # Apply threshold to make QR code more visible
                    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                    
                    # Try to decode QR code from both original and thresholded image
                    decoded_objects = decode(frame)
                    if not decoded_objects:
                        decoded_objects = decode(thresh)
                    
                    for obj in decoded_objects:
                        try:
                            # Draw rectangle around QR code
                            points = obj.polygon
                            if len(points) > 4:
                                hull = cv2.convexHull(np.array([point for point in points], dtype=np.float32))
                                cv2.polylines(frame, [hull], True, (0, 255, 0), 2)
                            else:
                                cv2.polylines(frame, [np.array(points, dtype=np.int32)], True, (0, 255, 0), 2)
                            
                            # Decode QR data
                            qr_data = obj.data.decode('utf-8')
                            from core.security import decrypt_qr_data
                            decoded_data = decrypt_qr_data(qr_data)
                            name, grade, number, register_date, expire_date = decoded_data.split("|")
                            
                            # Update status
                            status_label.configure(text="✓ สแกนสำเร็จ!", text_color="green")
                            
                            # Put data in queue for main thread to process
                            qr_queue.put((name, grade, number))
                            
                            # Close scan window after a short delay
                            scan_window.after(1000, lambda: [scan_window.destroy(), cap.release()])
                            return
                            
                        except Exception as e:
                            print(f"Error decoding QR: {str(e)}")
                            status_label.configure(text="❌ ไม่สามารถอ่าน QR Code ได้", text_color="red")
                    
                    # Convert frame to CTkImage
                    frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    frame_pil = Image.fromarray(frame_rgb)
                    frame_ctk = ctk.CTkImage(light_image=frame_pil, dark_image=frame_pil, size=(640, 480))
                    
                    # Update video label
                    video_label.configure(image=frame_ctk)
                    
                except Exception as e:
                    print(f"Error in scan loop: {str(e)}")
                    break
                    
                # Check if window is closed
                if not scan_window.winfo_exists():
                    break
                    
                time.sleep(0.03)  # 30 FPS
                
            # Clean up
            cap.release()
                
        # Start scanning in separate thread
        scan_thread = threading.Thread(target=scan_qr)
        scan_thread.daemon = True
        scan_thread.start()
        
        def process_qr_data():
            try:
                if not qr_queue.empty():
                    name, grade, number = qr_queue.get()
                    self.cursor.execute("""
                        SELECT id FROM members 
                        WHERE name = ? AND grade = ? AND number = ?
                    """, (name, grade, number))
                    member = self.cursor.fetchone()
                    if member:
                        member_id = member[0]
                        # Update status and member info
                        self.return_status_label.configure(
                            text="พบข้อมูลสมาชิก",
                            text_color="green"
                        )
                        self.return_member_info_label.configure(
                            text=f"สมาชิก: {name}\nชั้น: {grade}\nเลขที่: {number}"
                        )
                        # Display borrowed books
                        self.display_borrowed_books(member_id)
                    else:
                        self.return_status_label.configure(
                            text="ไม่พบข้อมูลสมาชิกในระบบ",
                            text_color="red"
                        )
                if scan_window.winfo_exists():
                    scan_window.after(100, process_qr_data)
            except Exception as e:
                print(f"Error processing QR data: {str(e)}")
        
        # Start processing QR data
        scan_window.after(100, process_qr_data)
        
        # Handle window close
        def on_closing():
            cap.release()
            scan_window.destroy()
            
        scan_window.protocol("WM_DELETE_WINDOW", on_closing)

    def display_borrowed_books(self, member_id):
        # Clear existing books
        for widget in self.borrowed_books_frame.winfo_children():
            widget.destroy()
            
        # Get borrowed books
        self.cursor.execute('''
            SELECT b.id, b.code, b.title, bl.borrow_date, bl.return_due
            FROM books b
            JOIN borrow_log bl ON b.id = bl.book_id
            WHERE bl.member_id = ? AND bl.returned = 0
        ''', (member_id,))
        books = self.cursor.fetchall()
        
        if not books:
            no_books_label = ctk.CTkLabel(
                self.borrowed_books_frame, 
                text="ไม่มีหนังสือที่ยืมอยู่",
                font=("Helvetica", 14)
            )
            no_books_label.pack(pady=10)
            return
            
        # Add header
        header_label = ctk.CTkLabel(
            self.borrowed_books_frame,
            text="รายการหนังสือที่ยืม",
            font=("Helvetica", 16, "bold")
        )
        header_label.pack(pady=(0, 10))
            
        # Display each book
        for book in books:
            book_frame = ctk.CTkFrame(self.borrowed_books_frame)
            book_frame.pack(pady=5, padx=10, fill="x")
            
            # Book info
            info_text = f"รหัส: {book[1]}\nชื่อ: {book[2]}\nวันที่ยืม: {book[3]}\nกำหนดคืน: {book[4]}"
            info_label = ctk.CTkLabel(book_frame, text=info_text, font=("Helvetica", 12))
            info_label.pack(side="left", padx=10, pady=5)
            
            # Return button
            return_button = ctk.CTkButton(
                book_frame,
                text="คืน",
                command=lambda b=book: self.return_book(b[0], member_id),
                width=80,
                height=32,
                font=("Helvetica", 12)
            )
            return_button.pack(side="right", padx=10, pady=5)
        
    def return_book(self, book_id, member_id):
        try:
            self.cursor.execute('''
                UPDATE borrow_log 
                SET returned = 1 
                WHERE book_id = ? AND returned = 0
            ''', (book_id,))
            self.cursor.execute("UPDATE books SET status = 'ว่าง' WHERE id = ?", (book_id,))
            self.conn.commit()
            self.show_success("คืนหนังสือสำเร็จ")
            self.display_borrowed_books(member_id)  # Refresh list
        except Exception as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")
        
    def show_history(self):
        # Clear any existing widgets
        for widget in self.app.winfo_children():
            widget.destroy()
            
        # Create history frame
        history_frame = ctk.CTkFrame(self.app)
        history_frame.pack(pady=20, padx=40, fill="both", expand=True)
        
        # Title
        title_label = ctk.CTkLabel(history_frame, text="ประวัติการยืม-คืน", font=("Helvetica", 24))
        title_label.pack(pady=20)
        
        # Search frame
        search_frame = ctk.CTkFrame(history_frame)
        search_frame.pack(pady=10, padx=20, fill="x")
        
        # Search fields
        member_entry = ctk.CTkEntry(search_frame, placeholder_text="ค้นหาตามชื่อสมาชิก")
        member_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        book_entry = ctk.CTkEntry(search_frame, placeholder_text="ค้นหาตามรหัสหนังสือ")
        book_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        date_entry = ctk.CTkEntry(search_frame, placeholder_text="ค้นหาตามวันที่ (YYYY-MM-DD)")
        date_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # Search button
        search_button = ctk.CTkButton(search_frame, text="ค้นหา", 
                                    command=lambda: self.search_history(
                                        member_entry.get(),
                                        book_entry.get(),
                                        date_entry.get()
                                    ))
        search_button.pack(side="left", padx=5)
        
        # Export button
        export_button = ctk.CTkButton(search_frame, text="Export PDF", 
                                    command=self.export_history)
        export_button.pack(side="left", padx=5)
        
        # History display frame
        self.history_display_frame = ctk.CTkScrollableFrame(history_frame)
        self.history_display_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Back button
        back_button = ctk.CTkButton(history_frame, text="กลับ", 
                                  command=self.show_dashboard)
        back_button.pack(pady=10)
        
        # Display initial history
        self.display_history()
        
    def display_history(self, member_filter="", book_filter="", date_filter=""):
        # Clear existing history
        for widget in self.history_display_frame.winfo_children():
            widget.destroy()
            
        # Build query
        query = '''
            SELECT m.name, m.grade, m.number, b.code, b.title, bl.borrow_date, bl.return_due, bl.returned
            FROM borrow_log bl
            JOIN members m ON bl.member_id = m.id
            JOIN books b ON bl.book_id = b.id
            WHERE 1=1
        '''
        params = []
        
        if member_filter:
            query += " AND (m.name LIKE ? OR m.grade LIKE ? OR m.number LIKE ?)"
            params.extend([f"%{member_filter}%"] * 3)
            
        if book_filter:
            query += " AND (b.code LIKE ? OR b.title LIKE ?)"
            params.extend([f"%{book_filter}%"] * 2)
            
        if date_filter:
            query += " AND (bl.borrow_date = ? OR bl.return_due = ?)"
            params.extend([date_filter] * 2)
            
        query += " ORDER BY bl.borrow_date DESC"
        
        # Execute query
        self.cursor.execute(query, params)
        records = self.cursor.fetchall()
        
        if not records:
            no_records_label = ctk.CTkLabel(self.history_display_frame, text="ไม่พบประวัติ")
            no_records_label.pack(pady=10)
            return
            
        # Display each record
        for record in records:
            record_frame = ctk.CTkFrame(self.history_display_frame)
            record_frame.pack(pady=5, padx=10, fill="x")
            
            # Record info
            status = "คืนแล้ว" if record[7] else "ยังไม่คืน"
            info_text = (f"สมาชิก: {record[0]} ({record[1]}/{record[2]}) | "
                        f"หนังสือ: {record[3]} - {record[4]} | "
                        f"ยืม: {record[5]} | คืน: {record[6]} | "
                        f"สถานะ: {status}")
            info_label = ctk.CTkLabel(record_frame, text=info_text)
            info_label.pack(pady=5, padx=10)
            
    def search_history(self, member_filter, book_filter, date_filter):
        self.display_history(member_filter, book_filter, date_filter)
        
    def export_history(self):
        try:
            # ลงทะเบียนฟอนต์ Sarabun
            font_dir = "assets/fonts"
            pdfmetrics.registerFont(TTFont('Sarabun', os.path.join(font_dir, 'Sarabun-Regular.ttf')))
            pdfmetrics.registerFont(TTFont('Sarabun-Bold', os.path.join(font_dir, 'Sarabun-Bold.ttf')))

            # Get all history records
            self.cursor.execute('''
                SELECT m.name, m.grade, m.number, b.code, b.title, 
                       bl.borrow_date, bl.return_due, bl.returned
                FROM borrow_log bl
                JOIN members m ON bl.member_id = m.id
                JOIN books b ON bl.book_id = b.id
                ORDER BY bl.borrow_date DESC
            ''')
            records = self.cursor.fetchall()
            
            # Create PDF file
            from reportlab.lib.pagesizes import A4
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
            from reportlab.lib.styles import ParagraphStyle

            filename = f"history_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
            
            # Create table data
            data = [["ชื่อสมาชิก", "ชั้น", "เลขที่", "รหัสหนังสือ", "ชื่อหนังสือ", "วันที่ยืม", "วันที่คืน", "สถานะ"]]
            for record in records:
                status = "คืนแล้ว" if record[7] else "ยังไม่คืน"
                data.append([
                    record[0], record[1], record[2], record[3], record[4],
                    record[5], record[6], status
                ])
            
            # สร้าง style สำหรับฟอนต์ Sarabun
            style = ParagraphStyle(name='Sarabun', fontName='Sarabun', fontSize=12, leading=14)
            style_bold = ParagraphStyle(name='Sarabun-Bold', fontName='Sarabun-Bold', fontSize=13, leading=15)

            # แปลงข้อมูลเป็น Paragraph เพื่อรองรับภาษาไทย
            data = [
                [Paragraph(cell, style_bold) if i == 0 else Paragraph(cell, style) for i, cell in enumerate(row)]
                if idx == 0 else
                [Paragraph(str(cell), style) for cell in row]
                for idx, row in enumerate(data)
            ]

            # กำหนดความกว้างคอลัมน์
            col_widths = [70, 40, 40, 60, 120, 60, 60, 50]

            # Create table
            table = Table(data, colWidths=col_widths)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4B0082')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Sarabun-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 13),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#E6E6FA')),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#222222')),
                ('FONTNAME', (0, 1), (-1, -1), 'Sarabun'),
                ('FONTSIZE', (0, 1), (-1, -1), 12),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#4B0082'))
            ]))
            
            # Build PDF
            elements = [table]
            doc.build(elements)
            
            self.show_success(f"Export สำเร็จ: {filename}")
        except Exception as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")
        
    def generate_member_card_pdf(self, member_name, member_number, school_logo_path, output_pdf_path, qr_base64):
        # ลงทะเบียนฟอนต์ (ทำครั้งเดียวต่อโปรแกรม)
        font_dir = "assets/fonts"
        pdfmetrics.registerFont(TTFont('Sarabun', os.path.join(font_dir, 'Sarabun-Regular.ttf')))
        pdfmetrics.registerFont(TTFont('Sarabun-Bold', os.path.join(font_dir, 'Sarabun-Bold.ttf')))

        # ขนาดบัตร
        card_width = 85.6 * mm
        card_height = 53.98 * mm

        # สร้าง QR Code และบันทึกเป็นไฟล์ชั่วคราว
        qr = qrcode.QRCode(version=1, box_size=10, border=5)
        qr.add_data(qr_base64)
        qr.make(fit=True)
        pil_image = qr.get_image() if hasattr(qr, 'get_image') else qr.make_image()
        temp_qr_path = f"assets/qrcodes/temp_{member_number}.png"
        pil_image.save(temp_qr_path)

        c = canvas.Canvas(output_pdf_path, pagesize=(card_width, card_height))

        # สี
        purple = colors.HexColor('#4B0082')
        light_purple = colors.HexColor('#E6E6FA')

        # พื้นหลัง
        c.setFillColor(light_purple)
        c.rect(0, 0, card_width, card_height, fill=1)

        # ขอบบัตร
        c.setStrokeColor(purple)
        c.setLineWidth(2)
        c.rect(2*mm, 2*mm, card_width-4*mm, card_height-4*mm, stroke=1, fill=0)

        # หัวข้อบัตร
        c.setFillColor(purple)
        c.setFont("Sarabun-Bold", 16)
        c.drawString(28*mm, card_height-8*mm, "บัตรสมาชิกห้องสมุด")

        # โลโก้โรงเรียน
        if os.path.exists(school_logo_path):
            c.drawImage(school_logo_path, 5*mm, card_height-20*mm, width=18*mm, height=18*mm, mask='auto')
        else:
            c.setFillColor(purple)
            c.rect(5*mm, card_height-20*mm, 18*mm, 18*mm, stroke=1, fill=0)

        # QR Code
        c.drawImage(temp_qr_path, card_width-23*mm, 5*mm, width=18*mm, height=18*mm, mask='auto')

        # ข้อมูลสมาชิก
        c.setFillColor(purple)
        c.setFont("Sarabun-Bold", 14)
        c.drawString(28*mm, card_height-16*mm, f"ชื่อ: {member_name}")
        c.setFont("Sarabun", 13)
        c.drawString(28*mm, card_height-24*mm, f"เลขที่: {member_number}")

        # เส้นคั่น
        c.setStrokeColor(purple)
        c.setLineWidth(1)
        c.line(28*mm, card_height-28*mm, card_width-28*mm, card_height-28*mm)

        c.save()

        # ลบไฟล์ QR Code ชั่วคราว
        if os.path.exists(temp_qr_path):
            os.remove(temp_qr_path)

    def show_popup(self, title, message):
        popup = ctk.CTkToplevel(self.app)
        popup.title(title)
        popup.geometry("320x200")
        popup.grab_set()  # focus modal

        label = ctk.CTkLabel(popup, text=message, font=("Sarabun", 14))
        label.pack(pady=30)

        def close_popup(event=None):
            popup.destroy()

        ok_btn = ctk.CTkButton(popup, text="OK", command=close_popup)
        ok_btn.pack(pady=10)
        ok_btn.focus_set()
        popup.bind('<Return>', close_popup)

    def show_settings(self):
        for widget in self.app.winfo_children():
            widget.destroy()

        settings_frame = ctk.CTkFrame(self.app)
        settings_frame.pack(pady=20, padx=40, fill="both", expand=True)

        title_label = ctk.CTkLabel(settings_frame, text="การตั้งค่า", font=("Sarabun-Bold", 24))
        title_label.pack(pady=20)

        # --- User Table ---
        table_frame = ctk.CTkFrame(settings_frame)
        table_frame.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(table_frame, text="รายชื่อผู้ใช้ทั้งหมด", font=("Sarabun-Bold", 16)).pack(pady=2)

        self.cursor.execute("SELECT username, password FROM admin_users")
        users = self.cursor.fetchall()
        for user in users:
            user_row = ctk.CTkFrame(table_frame)
            user_row.pack(fill="x", padx=5, pady=1)
            ctk.CTkLabel(user_row, text=user[0], font=("Sarabun", 13), width=120, anchor="w").pack(side="left", padx=5)
            ctk.CTkLabel(user_row, text="********", font=("Sarabun", 13), width=120, anchor="w").pack(side="left", padx=5)

            def make_delete_user(username):
                return lambda: self.delete_user(username)

            def make_change_pass(username):
                return lambda: self.change_password_popup(username)

            del_btn = ctk.CTkButton(user_row, text="ลบ", fg_color="red", command=make_delete_user(user[0]), width=60)
            del_btn.pack(side="left", padx=5)
            change_btn = ctk.CTkButton(user_row, text="เปลี่ยนรหัส", command=make_change_pass(user[0]), width=100)
            change_btn.pack(side="left", padx=5)

        # --- User Management Form ---
        user_frame = ctk.CTkFrame(settings_frame)
        user_frame.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(user_frame, text="ชื่อผู้ใช้ใหม่:", font=("Sarabun", 14)).pack(pady=2)
        username_entry = ctk.CTkEntry(user_frame)
        username_entry.pack(pady=2)

        ctk.CTkLabel(user_frame, text="รหัสผ่านใหม่:", font=("Sarabun", 14)).pack(pady=2)
        password_entry = ctk.CTkEntry(user_frame, show="*")
        password_entry.pack(pady=2)

        def save_user():
            username = username_entry.get()
            password = password_entry.get()
            if not username or not password:
                self.show_error("กรุณากรอกชื่อผู้ใช้และรหัสผ่าน")
                return
            try:
                from core.security import hash_password
                hashed_pw = hash_password(password)
                self.cursor.execute("INSERT OR REPLACE INTO admin_users (username, password) VALUES (?, ?)", (username, hashed_pw))
                self.conn.commit()
                self.show_success("บันทึกผู้ใช้สำเร็จ")
                self.show_settings()  # refresh table
            except Exception as e:
                self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")

        save_user_btn = ctk.CTkButton(user_frame, text="เพิ่มผู้ใช้", command=save_user)
        save_user_btn.pack(pady=5)

        # --- Theme Switch ---
        theme_frame = ctk.CTkFrame(settings_frame)
        theme_frame.pack(pady=20, padx=20, fill="x")

        ctk.CTkLabel(theme_frame, text="โหมดสี:", font=("Sarabun", 14)).pack(pady=2)
        theme_var = ctk.StringVar(value=ctk.get_appearance_mode())
        def change_theme():
            ctk.set_appearance_mode(theme_var.get())
        theme_option = ctk.CTkOptionMenu(theme_frame, variable=theme_var, values=["Light", "Dark", "System"], command=lambda _: change_theme())
        theme_option.pack(pady=5)

        # --- Back Button ---
        back_btn = ctk.CTkButton(settings_frame, text="กลับ", command=self.show_dashboard)
        back_btn.pack(pady=20)

    def delete_user(self, username):
        if username == "admin":
            self.show_error("ไม่สามารถลบผู้ใช้ admin ได้")
            return
        try:
            self.cursor.execute("DELETE FROM admin_users WHERE username = ?", (username,))
            self.conn.commit()
            self.show_success("ลบผู้ใช้สำเร็จ")
            self.show_settings()
        except Exception as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")

    def change_password_popup(self, username):
        popup = ctk.CTkToplevel(self.app)
        popup.title("เปลี่ยนรหัสผ่าน")
        popup.geometry("300x180")
        popup.grab_set()

        ctk.CTkLabel(popup, text=f"เปลี่ยนรหัสผ่านสำหรับ {username}", font=("Sarabun-Bold", 14)).pack(pady=10)
        new_pass_entry = ctk.CTkEntry(popup, show="*")
        new_pass_entry.pack(pady=5)
        new_pass_entry.focus_set()

        def save_new_pass():
            new_pass = new_pass_entry.get()
            if not new_pass:
                self.show_error("กรุณากรอกรหัสผ่านใหม่")
                return
            try:
                from core.security import hash_password
                hashed_pw = hash_password(new_pass)
                self.cursor.execute("UPDATE admin_users SET password = ? WHERE username = ?", (hashed_pw, username))
                self.conn.commit()
                self.show_success("เปลี่ยนรหัสผ่านสำเร็จ")
                popup.destroy()
                self.show_settings()
            except Exception as e:
                self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")

        ctk.CTkButton(popup, text="บันทึก", command=save_new_pass).pack(pady=10)
        popup.bind('<Return>', lambda e: save_new_pass())

    def show_about(self):
        for widget in self.app.winfo_children():
            widget.destroy()

        about_frame = ctk.CTkFrame(self.app)
        about_frame.pack(pady=20, padx=40, fill="both", expand=True)

        title_label = ctk.CTkLabel(about_frame, text="เกี่ยวกับโปรแกรม", font=("Sarabun-Bold", 24))
        title_label.pack(pady=20)

        # ข้อมูลสมาชิก
        members = [
            "เด็กหญิง ขวัญชนก อุ่นศิริ เลขที่ 4",
            "เด็กชาย ณัฐชนน รอดน้อย เลขที่ 30",
            "เด็กชาย ปริญญา บุญเทียน เลขที่ 38"
        ]
        member_text = "\n".join(members)
        member_label = ctk.CTkLabel(about_frame, text="สมาชิกผู้พัฒนา:\n" + member_text, font=("Sarabun", 16), justify="left")
        member_label.pack(pady=10)

        # ปุ่มกลับ
        back_btn = ctk.CTkButton(about_frame, text="กลับ", command=self.show_dashboard)
        back_btn.pack(pady=20)

    def show_access_scanner(self):
        # Clear any existing widgets
        for widget in self.app.winfo_children():
            widget.destroy()
            
        # Create scanner frame
        scanner_frame = ctk.CTkFrame(self.app)
        scanner_frame.pack(pady=20, padx=40, fill="both", expand=True)
        
        # Title
        title_label = ctk.CTkLabel(scanner_frame, text="บันทึกการเข้าออกห้องสมุด", font=("Helvetica", 24))
        title_label.pack(pady=20)
        
        # Status display
        self.access_status_label = ctk.CTkLabel(scanner_frame, text="กรุณาสแกนบัตรสมาชิก", font=("Helvetica", 18))
        self.access_status_label.pack(pady=10)
        
        # Last scan info
        self.last_scan_label = ctk.CTkLabel(scanner_frame, text="", font=("Helvetica", 14))
        self.last_scan_label.pack(pady=10)
        
        # Scan button
        scan_button = ctk.CTkButton(scanner_frame, text="สแกนบัตร", 
                                  command=lambda: self.start_access_scan(),
                                  height=60,
                                  font=("Helvetica", 20))
        scan_button.pack(pady=20)
        
        # Back button
        back_button = ctk.CTkButton(scanner_frame, text="กลับ", 
                                  command=self.show_dashboard)
        back_button.pack(pady=10)
        
    def start_access_scan(self):
        # Create new window for scanning
        scan_window = ctk.CTkToplevel(self.app)
        scan_window.title("สแกน QR Code")
        scan_window.geometry("800x600")
        
        # Create frame for video
        video_frame = ctk.CTkFrame(scan_window)
        video_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Create label for video
        video_label = ctk.CTkLabel(video_frame, text="")
        video_label.pack(fill="both", expand=True)
        
        # Add instruction label
        instruction_label = ctk.CTkLabel(scan_window, 
            text="นำ QR Code มาวางตรงกล้อง\nรอสักครู่ระบบจะสแกนอัตโนมัติ",
            font=("Helvetica", 16))
        instruction_label.pack(pady=10)
        
        # Add status label
        status_label = ctk.CTkLabel(scan_window,
            text="กำลังรอสแกน...",
            font=("Helvetica", 14),
            text_color="yellow")
        status_label.pack(pady=5)
        
        # Start webcam
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            self.show_error("ไม่สามารถเปิดกล้องได้")
            scan_window.destroy()
            return
            
        # Set camera resolution
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        
        # Create a queue for thread communication
        qr_queue = Queue()
        
        def scan_qr():
            while True:
                try:
                    ret, frame = cap.read()
                    if not ret:
                        print("Failed to grab frame")
                        break
                        
                    # Resize frame for better display
                    frame = cv2.resize(frame, (640, 480))
                    
                    # Convert frame to grayscale
                    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                    
                    # Apply threshold to make QR code more visible
                    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                    
                    # Try to decode QR code from both original and thresholded image
                    decoded_objects = decode(frame)
                    if not decoded_objects:
                        decoded_objects = decode(thresh)
                    
                    for obj in decoded_objects:
                        try:
                            # Draw rectangle around QR code
                            points = obj.polygon
                            if len(points) > 4:
                                hull = cv2.convexHull(np.array([point for point in points], dtype=np.float32))
                                cv2.polylines(frame, [hull], True, (0, 255, 0), 2)
                            else:
                                cv2.polylines(frame, [np.array(points, dtype=np.int32)], True, (0, 255, 0), 2)
                            
                            # Decode QR data
                            qr_data = obj.data.decode('utf-8')
                            from core.security import decrypt_qr_data
                            decoded_data = decrypt_qr_data(qr_data)
                            name, grade, number, register_date, expire_date = decoded_data.split("|")
                            
                            # Update status
                            status_label.configure(text="✓ สแกนสำเร็จ!", text_color="green")
                            
                            # Put data in queue for main thread to process
                            qr_queue.put((name, grade, number))
                            
                            # Close scan window after a short delay
                            scan_window.after(1000, lambda: [scan_window.destroy(), cap.release()])
                            return
                            
                        except Exception as e:
                            print(f"Error decoding QR: {str(e)}")
                            status_label.configure(text="❌ ไม่สามารถอ่าน QR Code ได้", text_color="red")
                    
                    # Convert frame to CTkImage
                    frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    frame_pil = Image.fromarray(frame_rgb)
                    frame_ctk = ctk.CTkImage(light_image=frame_pil, dark_image=frame_pil, size=(640, 480))
                    
                    # Update video label
                    video_label.configure(image=frame_ctk)
                    
                except Exception as e:
                    print(f"Error in scan loop: {str(e)}")
                    break
                    
                # Check if window is closed
                if not scan_window.winfo_exists():
                    break
                    
                time.sleep(0.03)  # 30 FPS
                
            # Clean up
            cap.release()
                
        # Start scanning in separate thread
        scan_thread = threading.Thread(target=scan_qr)
        scan_thread.daemon = True
        scan_thread.start()
        
        def process_qr_data():
            try:
                if not qr_queue.empty():
                    name, grade, number = qr_queue.get()
                    self.cursor.execute("""
                        SELECT id FROM members 
                        WHERE name = ? AND grade = ? AND number = ?
                    """, (name, grade, number))
                    member = self.cursor.fetchone()
                    if member:
                        member_id = member[0]
                        # Get last access record for this member
                        self.cursor.execute("""
                            SELECT action FROM access_log 
                            WHERE member_id = ? 
                            ORDER BY access_time DESC LIMIT 1
                        """, (member_id,))
                        last_action = self.cursor.fetchone()
                        
                        # Determine action (if last action was 'เข้า', then 'ออก', and vice versa)
                        action = "ออก" if last_action and last_action[0] == "เข้า" else "เข้า"
                        
                        # Record access
                        self.cursor.execute("""
                            INSERT INTO access_log (member_id, access_time, action)
                            VALUES (?, datetime('now', 'localtime'), ?)
                        """, (member_id, action))
                        self.conn.commit()
                        
                        # Update status labels
                        self.access_status_label.configure(
                            text=f"บันทึกการ{action}ห้องสมุดสำเร็จ",
                            text_color="green"
                        )
                        self.last_scan_label.configure(
                            text=f"สมาชิก: {name}\nชั้น: {grade}\nเลขที่: {number}\nเวลา: {datetime.now().strftime('%H:%M:%S')}"
                        )
                    else:
                        self.access_status_label.configure(
                            text="ไม่พบข้อมูลสมาชิกในระบบ",
                            text_color="red"
                        )
                if scan_window.winfo_exists():
                    scan_window.after(100, process_qr_data)
            except Exception as e:
                print(f"Error processing QR data: {str(e)}")
        
        # Start processing QR data
        scan_window.after(100, process_qr_data)
        
        # Handle window close
        def on_closing():
            cap.release()
            scan_window.destroy()
            
        scan_window.protocol("WM_DELETE_WINDOW", on_closing)
        
    def show_access_history(self):
        # Clear any existing widgets
        for widget in self.app.winfo_children():
            widget.destroy()
            
        # Create history frame
        history_frame = ctk.CTkFrame(self.app)
        history_frame.pack(pady=20, padx=40, fill="both", expand=True)
        
        # Title
        title_label = ctk.CTkLabel(history_frame, text="ประวัติการเข้าออกห้องสมุด", font=("Helvetica", 24))
        title_label.pack(pady=20)
        
        # Search frame
        search_frame = ctk.CTkFrame(history_frame)
        search_frame.pack(pady=10, padx=20, fill="x")
        
        # Search fields
        member_entry = ctk.CTkEntry(search_frame, placeholder_text="ค้นหาตามชื่อสมาชิก")
        member_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        date_entry = ctk.CTkEntry(search_frame, placeholder_text="ค้นหาตามวันที่ (YYYY-MM-DD)")
        date_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # Search button
        search_button = ctk.CTkButton(search_frame, text="ค้นหา", 
                                    command=lambda: self.search_access_history(
                                        member_entry.get(),
                                        date_entry.get()
                                    ))
        search_button.pack(side="left", padx=5)
        
        # Export button
        export_button = ctk.CTkButton(search_frame, text="Export PDF", 
                                    command=self.export_access_history)
        export_button.pack(side="left", padx=5)
        
        # History display frame
        self.access_history_frame = ctk.CTkScrollableFrame(history_frame)
        self.access_history_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Back button
        back_button = ctk.CTkButton(history_frame, text="กลับ", 
                                  command=self.show_dashboard)
        back_button.pack(pady=10)
        
        # Display initial history
        self.display_access_history()
        
    def display_access_history(self, member_filter="", date_filter=""):
        # Clear existing history
        for widget in self.access_history_frame.winfo_children():
            widget.destroy()
            
        # Build query
        query = '''
            SELECT m.name, m.grade, m.number, al.access_time, al.action
            FROM access_log al
            JOIN members m ON al.member_id = m.id
            WHERE 1=1
        '''
        params = []
        
        if member_filter:
            query += " AND (m.name LIKE ? OR m.grade LIKE ? OR m.number LIKE ?)"
            params.extend([f"%{member_filter}%"] * 3)
            
        if date_filter:
            query += " AND date(al.access_time) = ?"
            params.append(date_filter)
            
        query += " ORDER BY al.access_time DESC"
        
        # Execute query
        self.cursor.execute(query, params)
        records = self.cursor.fetchall()
        
        if not records:
            no_records_label = ctk.CTkLabel(self.access_history_frame, text="ไม่พบประวัติ")
            no_records_label.pack(pady=10)
            return
            
        # Display each record
        for record in records:
            record_frame = ctk.CTkFrame(self.access_history_frame)
            record_frame.pack(pady=5, padx=10, fill="x")
            
            # Record info
            info_text = f"สมาชิก: {record[0]} ({record[1]}/{record[2]}) | เวลา: {record[3]} | การทำรายการ: {record[4]}"
            info_label = ctk.CTkLabel(record_frame, text=info_text)
            info_label.pack(pady=5, padx=10)
            
    def search_access_history(self, member_filter, date_filter):
        self.display_access_history(member_filter, date_filter)
        
    def export_access_history(self):
        try:
            # ตรวจสอบว่ามีโฟลเดอร์สำหรับเก็บรายงานหรือไม่
            if not os.path.exists('reports'):
                os.makedirs('reports')

            # Get all history records
            self.cursor.execute('''
                SELECT m.name, m.grade, m.number, al.access_time, al.action
                FROM access_log al
                JOIN members m ON al.member_id = m.id
                ORDER BY al.access_time DESC
            ''')
            records = self.cursor.fetchall()
            
            if not records:
                self.show_error("ไม่พบข้อมูลสำหรับการ export")
                return

            # Create PDF file
            filename = os.path.join('reports', f"access_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
            doc = SimpleDocTemplate(
                filename,
                pagesize=A4,
                rightMargin=30,
                leftMargin=30,
                topMargin=30,
                bottomMargin=30
            )
            
            # Create table data
            data = [["ชื่อสมาชิก", "ชั้น", "เลขที่", "เวลา", "การทำรายการ"]]
            for record in records:
                data.append([
                    str(record[0]),  # name
                    str(record[1]),  # grade
                    str(record[2]),  # number
                    str(record[3]),  # time
                    str(record[4])   # action
                ])
            
            # Register Thai font
            font_path = os.path.join('assets', 'fonts', 'Sarabun-Regular.ttf')
            font_path_bold = os.path.join('assets', 'fonts', 'Sarabun-Bold.ttf')
            
            if not os.path.exists(font_path) or not os.path.exists(font_path_bold):
                self.show_error("ไม่พบไฟล์ฟอนต์ที่จำเป็น กรุณาตรวจสอบโฟลเดอร์ assets/fonts")
                return
                
            pdfmetrics.registerFont(TTFont('Sarabun', font_path))
            pdfmetrics.registerFont(TTFont('Sarabun-Bold', font_path_bold))
            
            # Create styles
            styles = getSampleStyleSheet()
            header_style = ParagraphStyle(
                'HeaderStyle',
                parent=styles['Normal'],
                fontName='Sarabun-Bold',
                fontSize=13,
                textColor=colors.white,
                alignment=1
            )
            
            cell_style = ParagraphStyle(
                'CellStyle',
                parent=styles['Normal'],
                fontName='Sarabun',
                fontSize=12,
                alignment=1
            )
            
            # Format data with Thai font support
            formatted_data = [[Paragraph(str(cell), header_style) for cell in data[0]]]
            for row in data[1:]:
                formatted_data.append([Paragraph(str(cell), cell_style) for cell in row])
            
            # Set column widths
            col_widths = [120, 60, 60, 120, 80]
            
            # Create table
            table = Table(formatted_data, colWidths=col_widths, repeatRows=1)
            table.setStyle(TableStyle([
                # Header style
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4B0082')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Sarabun-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 13),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('TOPPADDING', (0, 0), (-1, 0), 12),
                
                # Content style
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#E6E6FA')),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#222222')),
                ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 1), (-1, -1), 'Sarabun'),
                ('FONTSIZE', (0, 1), (-1, -1), 12),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 10),
                ('TOPPADDING', (0, 1), (-1, -1), 10),
                
                # Grid
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#4B0082')),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#E6E6FA'), colors.HexColor('#F0F0FF')])
            ]))
            
            # Add title
            title_style = ParagraphStyle(
                'TitleStyle',
                parent=styles['Normal'],
                fontName='Sarabun-Bold',
                fontSize=16,
                alignment=1,
                spaceAfter=30
            )
            title = Paragraph("รายงานประวัติการเข้าออกห้องสมุด", title_style)
            
            # Add date
            date_style = ParagraphStyle(
                'DateStyle',
                parent=styles['Normal'],
                fontName='Sarabun',
                fontSize=12,
                alignment=1,
                spaceAfter=20
            )
            current_date = Paragraph(
                f"วันที่ออกรายงาน: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                date_style
            )
            
            # Build PDF
            elements = [title, current_date, table]
            doc.build(elements)
            
            self.show_success(f"Export สำเร็จ: {filename}")
            
            # เปิดไฟล์ PDF หลังจาก export เสร็จ
            if os.name == 'nt':  # Windows
                os.startfile(filename)
            else:  # Linux/Mac
                import subprocess
                subprocess.run(['xdg-open', filename])
                
        except Exception as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")
            print(f"Error details: {str(e)}")  # For debugging

    def export_excel_template(self):
        try:
            # Create sample data
            sample_data = {
                'รหัสหนังสือ': ['001', '002'],
                'ชื่อเรื่อง': ['ตัวอย่างหนังสือ 1', 'ตัวอย่างหนังสือ 2']
            }
            df = pd.DataFrame(sample_data)
            
            # Save to Excel
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="book_import_template.xlsx"
            )
            
            if filename:
                df.to_excel(filename, index=False)
                self.show_success(f"บันทึกแม่แบบไฟล์ Excel สำเร็จ: {filename}")
                
                # เปิดไฟล์หลังจากบันทึก
                if os.name == 'nt':  # Windows
                    os.startfile(filename)
                else:  # Linux/Mac
                    import subprocess
                    subprocess.run(['xdg-open', filename])
        except Exception as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")

    def import_books_from_excel(self):
        try:
            filename = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx")],
                title="เลือกไฟล์ Excel ที่ต้องการนำเข้า"
            )
            
            if not filename:
                return
                
            # Read Excel file
            df = pd.read_excel(filename)
            
            # Validate columns
            required_columns = ['รหัสหนังสือ', 'ชื่อเรื่อง']
            if not all(col in df.columns for col in required_columns):
                self.show_error("รูปแบบไฟล์ Excel ไม่ถูกต้อง กรุณาใช้แม่แบบที่กำหนด")
                return
                
            # Start import
            success_count = 0
            error_count = 0
            error_messages = []
            
            for index, row in df.iterrows():
                try:
                    # Check if book code already exists
                    self.cursor.execute("SELECT id FROM books WHERE code = ?", (str(row['รหัสหนังสือ']),))
                    if self.cursor.fetchone():
                        error_count += 1
                        error_messages.append(f"รหัส {row['รหัสหนังสือ']} มีอยู่ในระบบแล้ว")
                        continue
                        
                    # Insert new book
                    self.cursor.execute('''
                        INSERT INTO books (code, title, status)
                        VALUES (?, ?, ?)
                    ''', (str(row['รหัสหนังสือ']), str(row['ชื่อเรื่อง']), "ว่าง"))
                    success_count += 1
                    
                except Exception as e:
                    error_count += 1
                    error_messages.append(f"ผิดพลาด: {str(e)}")
                    
            self.conn.commit()
            
            # Show result
            result_message = f"นำเข้าสำเร็จ: {success_count} รายการ\n"
            if error_count > 0:
                result_message += f"ผิดพลาด: {error_count} รายการ\n"
                result_message += "\n".join(error_messages[:5])  # Show first 5 errors
                if len(error_messages) > 5:
                    result_message += f"\n... และอีก {len(error_messages) - 5} ข้อผิดพลาด"
                    
            if success_count > 0:
                self.show_success(result_message)
            else:
                self.show_error(result_message)
                
            self.show_book_management()  # Refresh view
            
        except Exception as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")

    def run(self):
        self.app.mainloop()

class DatePicker:
    def __init__(self, parent, entry_widget):
        self.parent = parent
        self.entry_widget = entry_widget
        self.top = None
        
        # Thai month names
        self.thai_months = [
            "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
        ]
        
    def show(self):
        if self.top is not None and self.top.winfo_exists():
            return
            
        self.top = ctk.CTkToplevel(self.parent)
        self.top.title("เลือกวันที่")
        
        # Get current date from entry or use today
        try:
            current_date = datetime.strptime(self.entry_widget.get(), "%Y-%m-%d")
        except:
            current_date = datetime.now()
            
        self.year = current_date.year
        self.month = current_date.month
        
        # Month and Year Selection Frame
        nav_frame = ctk.CTkFrame(self.top)
        nav_frame.pack(pady=5, padx=5, fill="x")
        
        # Previous Month Button
        prev_month = ctk.CTkButton(nav_frame, text="◀", width=30, 
                                 command=self.prev_month)
        prev_month.pack(side="left", padx=5)
        
        # Month Label
        self.month_label = ctk.CTkLabel(nav_frame, 
                                      text=f"{self.thai_months[self.month-1]} {self.year}",
                                      font=("Helvetica", 14))
        self.month_label.pack(side="left", expand=True)
        
        # Next Month Button
        next_month = ctk.CTkButton(nav_frame, text="▶", width=30,
                                 command=self.next_month)
        next_month.pack(side="right", padx=5)
        
        # Days Header
        days_frame = ctk.CTkFrame(self.top)
        days_frame.pack(pady=5, padx=5)
        
        # Thai day names
        days = ["อา", "จ", "อ", "พ", "พฤ", "ศ", "ส"]
        for i, day in enumerate(days):
            ctk.CTkLabel(days_frame, text=day, width=30).grid(row=0, column=i, padx=1, pady=1)
        
        # Create calendar
        self.cal_frame = ctk.CTkFrame(self.top)
        self.cal_frame.pack(pady=5, padx=5)
        
        self.update_calendar()
        
        # Quick buttons frame
        quick_frame = ctk.CTkFrame(self.top)
        quick_frame.pack(pady=5, padx=5, fill="x")
        
        # Today button
        today_btn = ctk.CTkButton(quick_frame, text="วันนี้", 
                                command=lambda: self.set_date(datetime.now()))
        today_btn.pack(side="left", padx=5, expand=True)
        
        # +7 days button
        week_btn = ctk.CTkButton(quick_frame, text="7 วัน", 
                               command=lambda: self.set_date(datetime.now() + timedelta(days=7)))
        week_btn.pack(side="left", padx=5, expand=True)
        
        # +14 days button
        two_week_btn = ctk.CTkButton(quick_frame, text="14 วัน",
                                   command=lambda: self.set_date(datetime.now() + timedelta(days=14)))
        two_week_btn.pack(side="left", padx=5, expand=True)
        
        # Center the window
        self.top.update_idletasks()
        width = self.top.winfo_width()
        height = self.top.winfo_height()
        x = (self.top.winfo_screenwidth() // 2) - (width // 2)
        y = (self.top.winfo_screenheight() // 2) - (height // 2)
        self.top.geometry(f'+{x}+{y}')
        
    def update_calendar(self):
        # Clear existing calendar
        for widget in self.cal_frame.winfo_children():
            widget.destroy()
            
        # Get first day of month and number of days
        first_day = datetime(self.year, self.month, 1)
        last_day = (first_day + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        num_days = last_day.day
        
        # Update month label
        self.month_label.configure(text=f"{self.thai_months[self.month-1]} {self.year}")
        
        # Calculate starting position
        start_pos = first_day.weekday()
        start_pos = (start_pos + 1) % 7  # Adjust for Sunday start
        
        # Create calendar buttons
        day = 1
        for i in range(6):  # 6 rows max
            for j in range(7):  # 7 days per week
                if (i == 0 and j < start_pos) or (day > num_days):
                    # Empty cell
                    ctk.CTkLabel(self.cal_frame, text="", width=30).grid(row=i+1, column=j, padx=1, pady=1)
                else:
                    # Date button
                    date = datetime(self.year, self.month, day)
                    btn = ctk.CTkButton(self.cal_frame, text=str(day), width=30,
                                     command=lambda d=date: self.set_date(d))
                    
                    # Highlight today
                    if date.date() == datetime.now().date():
                        btn.configure(fg_color="green")
                    
                    btn.grid(row=i+1, column=j, padx=1, pady=1)
                    day += 1
                    
    def prev_month(self):
        self.month -= 1
        if self.month < 1:
            self.month = 12
            self.year -= 1
        self.update_calendar()
        
    def next_month(self):
        self.month += 1
        if self.month > 12:
            self.month = 1
            self.year += 1
        self.update_calendar()
        
    def set_date(self, date):
        self.entry_widget.delete(0, 'end')
        self.entry_widget.insert(0, date.strftime("%Y-%m-%d"))
        if self.top:
            self.top.destroy()

    def show_borrow(self):
        # ... existing code ...

        # Due date frame
        self.due_date_frame = ctk.CTkFrame(self.book_info_frame)
        self.due_date_frame.pack(pady=10, fill="x")

        # Due date label
        due_date_label = ctk.CTkLabel(self.due_date_frame, text="กำหนดคืน:", font=("Helvetica", 14))
        due_date_label.pack(side="left", padx=5)

        # Default due date (7 days from now)
        default_due_date = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")
        self.due_date_entry = ctk.CTkEntry(self.due_date_frame, width=120)
        self.due_date_entry.pack(side="left", padx=5)
        self.due_date_entry.insert(0, default_due_date)

        # Calendar button
        calendar_button = ctk.CTkButton(self.due_date_frame, 
                                      text="📅", 
                                      width=40,
                                      command=lambda: DatePicker(self.app, self.due_date_entry).show())
        calendar_button.pack(side="left", padx=5)

        # Quick buttons frame
        quick_buttons_frame = ctk.CTkFrame(self.due_date_frame)
        quick_buttons_frame.pack(side="left", padx=5)

        # Quick selection buttons
        ctk.CTkButton(quick_buttons_frame, 
                     text="7 วัน",
                     width=60,
                     command=lambda: self.set_quick_date(7)).pack(side="left", padx=2)
        
        ctk.CTkButton(quick_buttons_frame,
                     text="14 วัน",
                     width=60,
                     command=lambda: self.set_quick_date(14)).pack(side="left", padx=2)

        # ... rest of existing code ...

    def set_quick_date(self, days):
        future_date = (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d")
        self.due_date_entry.delete(0, 'end')
        self.due_date_entry.insert(0, future_date)

if __name__ == "__main__":
    import sys
    
    if "--debug" in sys.argv:
        import logging
        import inspect
        import functools
        import traceback
        
        # 1. Setup Logging
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s [%(levelname)s] (%(threadName)s) %(message)s',
            handlers=[
                logging.StreamHandler(sys.stdout),
                logging.FileHandler("debug.log", encoding="utf-8")
            ]
        )
        logger = logging.getLogger("DebugLogger")
        logger.info("=== Debug Mode Enabled ===")

        # 2. Hook global unhandled exceptions
        def exception_hook(exctype, value, tb):
            logger.error("Unhandled main thread exception:", exc_info=(exctype, value, tb))
            sys.__excepthook__(exctype, value, tb)
        sys.excepthook = exception_hook

        # 3. Hook thread exceptions
        def thread_exception_hook(args):
            logger.error(f"Unhandled exception in thread '{args.thread.name}':", 
                         exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
        threading.excepthook = thread_exception_hook

        # 4. Decorator to trace methods
        def trace_methods(instance):
            for attr_name in dir(instance):
                # Skip magic/private methods
                if attr_name.startswith('__') and attr_name.endswith('__'):
                    continue
                try:
                    attr = getattr(instance, attr_name)
                    if inspect.ismethod(attr) or inspect.isfunction(attr):
                        def make_wrapper(func_name, func_obj):
                            @functools.wraps(func_obj)
                            def wrapper(*args, **kwargs):
                                safe_args = args[1:] if len(args) > 0 else ()
                                logger.debug(f"▶ [ENTER] {func_name} | Args: {safe_args} | Kwargs: {kwargs}")
                                start_t = time.time()
                                try:
                                    res = func_obj(*args, **kwargs)
                                    duration = time.time() - start_t
                                    logger.debug(f"◀ [EXIT]  {func_name} | Duration: {duration:.4f}s")
                                    return res
                                except Exception as exc:
                                    duration = time.time() - start_t
                                    logger.error(f"❌ [ERROR] {func_name} | Duration: {duration:.4f}s | Msg: {exc}\n{traceback.format_exc()}")
                                    raise exc
                            return wrapper
                        setattr(instance, attr_name, make_wrapper(attr_name, attr))
                except Exception as e:
                    pass

        # Instantiate app
        app = LibraryApp()
        
        # Apply tracing to the created app instance methods
        trace_methods(app)
        
        # Run app
        app.run()
    else:
        app = LibraryApp()
        app.run()
 
