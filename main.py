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
from reportlab.lib.pagesizes import A8
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
import io
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

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
                author TEXT,
                category TEXT,
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
        
        # Insert default admin user if not exists
        self.cursor.execute("SELECT * FROM admin_users WHERE username = 'admin'")
        if not self.cursor.fetchone():
            self.cursor.execute("INSERT INTO admin_users VALUES (?, ?)", ('admin', 'admin123'))
            
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
            self.cursor.execute("SELECT * FROM admin_users WHERE username = ? AND password = ?", 
                              (username, password))
            if self.cursor.fetchone():
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
            ("ตั้งค่า", self.show_settings, 1, 2),
            ("เกี่ยวกับ", self.show_about, 2, 0),
            ("ออกจากระบบ", self.show_login, 2, 1)
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
        encoded_data = base64.b64encode(qr_data.encode()).decode()
        
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
        
        author_entry = ctk.CTkEntry(form_frame, placeholder_text="ผู้แต่ง")
        author_entry.pack(pady=5, padx=20, fill="x")
        
        category_entry = ctk.CTkEntry(form_frame, placeholder_text="หมวดหมู่")
        category_entry.pack(pady=5, padx=20, fill="x")
        
        # Add book button
        add_button = ctk.CTkButton(form_frame, text="เพิ่มหนังสือ", 
                                 command=lambda: self.add_book(
                                     code_entry.get(),
                                     title_entry.get(),
                                     author_entry.get(),
                                     category_entry.get()
                                 ))
        add_button.pack(pady=10)
        
        # Back button
        back_button = ctk.CTkButton(book_frame, text="กลับ", 
                                  command=self.show_dashboard)
        back_button.pack(pady=10)
        
        # Display existing books
        self.display_books(book_frame)
        
    def add_book(self, code, title, author, category):
        if not all([code, title, author, category]):
            self.show_error("กรุณากรอกข้อมูลให้ครบ")
            return
            
        try:
            self.cursor.execute('''
                INSERT INTO books (code, title, author, category, status)
                VALUES (?, ?, ?, ?, ?)
            ''', (code, title, author, category, "ว่าง"))
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
            info_text = f"รหัส: {book[1]} | ชื่อ: {book[2]} | ผู้แต่ง: {book[3]} | หมวดหมู่: {book[4]} | สถานะ: {book[5]}"
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
                
            # Delete from database
            self.cursor.execute("DELETE FROM books WHERE id = ?", (book[0],))
            self.conn.commit()
            
            self.show_success("ลบหนังสือสำเร็จ")
            self.show_book_management()  # Refresh the view
        except Exception as e:
            self.show_error(f"เกิดข้อผิดพลาด: {str(e)}")
            
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
        
        # Scan button only (no QR entry)
        scan_button = ctk.CTkButton(member_frame, text="สแกน QR Code", command=self.start_scan)
        scan_button.pack(pady=5)
        
        # Member info display
        self.member_info_label = ctk.CTkLabel(member_frame, text="กรุณาสแกน QR Code สมาชิก")
        self.member_info_label.pack(pady=5)
        
        # Book selection frame
        book_frame = ctk.CTkFrame(borrow_frame)
        book_frame.pack(pady=10, padx=20, fill="x")
        
        # Book selection dropdown
        self.book_var = ctk.StringVar(value="เลือกหนังสือ")
        self.book_dropdown = ctk.CTkOptionMenu(book_frame, variable=self.book_var)
        self.book_dropdown.pack(pady=5, padx=20, fill="x")
        self.update_book_dropdown()
        
        # Due date selection
        due_date_frame = ctk.CTkFrame(borrow_frame)
        due_date_frame.pack(pady=10, padx=20, fill="x")
        
        due_date_label = ctk.CTkLabel(due_date_frame, text="กำหนดคืน:")
        due_date_label.pack(side="left", padx=10)
        
        self.due_date = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")
        due_date_entry = ctk.CTkEntry(due_date_frame, placeholder_text=self.due_date)
        due_date_entry.pack(side="left", padx=10)
        
        # Borrow button (no qr_data param)
        borrow_button = ctk.CTkButton(borrow_frame, text="ยืมหนังสือ", 
                                    command=lambda: self.borrow_book(
                                        self.book_var.get(),
                                        due_date_entry.get() or self.due_date
                                    ))
        borrow_button.pack(pady=10)
        
        # Back button
        back_button = ctk.CTkButton(borrow_frame, text="กลับ", command=self.show_dashboard)
        back_button.pack(pady=10)
        
    def start_scan(self, mode="borrow"):
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
                            decoded_data = base64.b64decode(qr_data).decode()
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
                        SELECT id, register_date, expire_date FROM members 
                        WHERE name = ? AND grade = ? AND number = ?
                    """, (name, grade, number))
                    member = self.cursor.fetchone()
                    if member:
                        member_id, register_date, expire_date = member
                        if mode == "return":
                            self.return_member_info_label.configure(
                                text=f"สมาชิก: {name}\nชั้น: {grade}\nเลขที่: {number}"
                            )
                            self.display_borrowed_books(member_id)
                            self.show_popup("พบข้อมูลสมาชิกในระบบ", f"ชื่อ: {name}\nชั้น: {grade}\nเลขที่: {number}")
                        else:
                            self.scanned_member = (name, grade, number, register_date, expire_date)
                            self.member_info_label.configure(
                                text=f"สมาชิก: {name}\nชั้น: {grade}\nเลขที่: {number}"
                            )
                            self.show_popup("พบข้อมูลสมาชิกในระบบ", f"ชื่อ: {name}\nชั้น: {grade}\nเลขที่: {number}")
                    else:
                        self.show_popup("ไม่พบข้อมูลสมาชิกในระบบ", f"ชื่อ: {name}\nชั้น: {grade}\nเลขที่: {number}")
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
        
    def update_book_dropdown(self):
        self.cursor.execute("SELECT id, code, title FROM books WHERE status = 'ว่าง'")
        books = self.cursor.fetchall()
        self.book_options = {f"{book[1]} - {book[2]}": book[0] for book in books}
        if hasattr(self, 'book_dropdown'):
            self.book_dropdown.configure(values=list(self.book_options.keys()))
            if list(self.book_options.keys()):
                self.book_dropdown.set(list(self.book_options.keys())[0])
            else:
                self.book_dropdown.set("เลือกหนังสือ")
            
    def borrow_book(self, book_text, due_date):
        print("เลือก:", book_text)
        print("ตัวเลือก:", list(self.book_options.keys()))
        if not hasattr(self, 'scanned_member') or not self.scanned_member:
            self.show_error("กรุณาสแกน QR Code สมาชิกก่อน")
            return
        if book_text == "เลือกหนังสือ":
            self.show_error("กรุณาเลือกหนังสือ")
            return
        if book_text not in self.book_options:
            self.show_error("ไม่พบหนังสือในระบบ กรุณาลองใหม่")
            return
        try:
            name, grade, number, register_date, expire_date = self.scanned_member
            self.cursor.execute("SELECT id FROM members WHERE name = ? AND grade = ? AND number = ?", (name, grade, number))
            member = self.cursor.fetchone()
            if not member:
                self.show_error("ไม่พบข้อมูลสมาชิก")
                return
            member_id = member[0]
            book_id = self.book_options[book_text]
            borrow_date = datetime.now().strftime("%Y-%m-%d")
            self.cursor.execute('''
                INSERT INTO borrow_log (member_id, book_id, borrow_date, return_due, returned)
                VALUES (?, ?, ?, ?, ?)
            ''', (member_id, book_id, borrow_date, due_date, 0))
            self.cursor.execute("UPDATE books SET status = 'ยืมแล้ว' WHERE id = ?", (book_id,))
            self.conn.commit()
            self.show_success("ยืมหนังสือสำเร็จ")
            self.show_dashboard()
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
        
        # QR code entry frame
        qr_frame = ctk.CTkFrame(member_frame)
        qr_frame.pack(pady=5, padx=20, fill="x")
        
        # QR code entry
        qr_entry = ctk.CTkEntry(qr_frame, placeholder_text="สแกน QR Code หรือวางข้อมูล")
        qr_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # Scan button
        scan_button = ctk.CTkButton(qr_frame, text="สแกน QR Code", 
                                    command=lambda: self.start_scan(mode="return"))
        scan_button.pack(side="right", padx=5)
        
        # Member info display
        self.return_member_info_label = ctk.CTkLabel(member_frame, text="")
        self.return_member_info_label.pack(pady=5)
        
        # Borrowed books frame
        self.borrowed_books_frame = ctk.CTkFrame(return_frame)
        self.borrowed_books_frame.pack(pady=10, padx=20, fill="x")
        
        # Back button
        back_button = ctk.CTkButton(return_frame, text="กลับ", 
                                  command=self.show_dashboard)
        back_button.pack(pady=10)
        
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
            no_books_label = ctk.CTkLabel(self.borrowed_books_frame, text="ไม่มีหนังสือที่ยืมอยู่")
            no_books_label.pack(pady=10)
            return
            
        # Display each book
        for book in books:
            book_frame = ctk.CTkFrame(self.borrowed_books_frame)
            book_frame.pack(pady=5, padx=10, fill="x")
            
            # Book info
            info_text = f"รหัส: {book[1]} | ชื่อ: {book[2]} | ยืม: {book[3]} | คืน: {book[4]}"
            info_label = ctk.CTkLabel(book_frame, text=info_text)
            info_label.pack(side="left", padx=10)
            
            # Return button
            return_button = ctk.CTkButton(
                book_frame, text="คืน",
                command=lambda b=book: self.return_book(b[0], member_id)
            )
            return_button.pack(side="right", padx=5)
            
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
            ctk.CTkLabel(user_row, text=user[1], font=("Sarabun", 13), width=120, anchor="w").pack(side="left", padx=5)

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
                self.cursor.execute("INSERT OR REPLACE INTO admin_users (username, password) VALUES (?, ?)", (username, password))
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
                self.cursor.execute("UPDATE admin_users SET password = ? WHERE username = ?", (new_pass, username))
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
            "เด็กชาย ปริญญา บุญเทียน เลขที่ 38",
            "เด็กชาย เอกวิญญ์ ลี้ไพบูลย์ เลขที่ 45"
        ]
        member_text = "\n".join(members)
        member_label = ctk.CTkLabel(about_frame, text="สมาชิกผู้พัฒนา:\n" + member_text, font=("Sarabun", 16), justify="left")
        member_label.pack(pady=10)

        # ปุ่มกลับ
        back_btn = ctk.CTkButton(about_frame, text="กลับ", command=self.show_dashboard)
        back_btn.pack(pady=20)

    def run(self):
        self.app.mainloop()

if __name__ == "__main__":
    app = LibraryApp()
    app.run() 