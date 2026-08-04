# services/pdf_service.py
import os
import qrcode
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

class PDFService:
    @staticmethod
    def register_fonts():
        font_dir = "assets/fonts"
        regular_path = os.path.join(font_dir, 'Sarabun-Regular.ttf')
        bold_path = os.path.join(font_dir, 'Sarabun-Bold.ttf')
        
        # Only register if they exist and aren't registered
        try:
            pdfmetrics.registerFont(TTFont('Sarabun', regular_path))
            pdfmetrics.registerFont(TTFont('Sarabun-Bold', bold_path))
        except Exception:
            pass

    @classmethod
    def generate_member_card_pdf(cls, member_name, member_number, school_logo_path, output_pdf_path, qr_base64):
        cls.register_fonts()
        
        # Card Size
        card_width = 85.6 * mm
        card_height = 53.98 * mm

        # Temp QR Code
        qr = qrcode.QRCode(version=1, box_size=10, border=5)
        qr.add_data(qr_base64)
        qr.make(fit=True)
        pil_image = qr.get_image() if hasattr(qr, 'get_image') else qr.make_image()
        temp_qr_path = f"assets/qrcodes/temp_{member_number}.png"
        pil_image.save(temp_qr_path)

        c = canvas.Canvas(output_pdf_path, pagesize=(card_width, card_height))

        # Colors
        purple = colors.HexColor('#4B0082')
        light_purple = colors.HexColor('#E6E6FA')

        # Background
        c.setFillColor(light_purple)
        c.rect(0, 0, card_width, card_height, fill=1)

        # Card Border
        c.setStrokeColor(purple)
        c.setLineWidth(2)
        c.rect(2*mm, 2*mm, card_width-4*mm, card_height-4*mm, stroke=1, fill=0)

        # Card Header
        c.setFillColor(purple)
        c.setFont("Sarabun-Bold", 16)
        c.drawString(28*mm, card_height-8*mm, "บัตรสมาชิกห้องสมุด")

        # School Logo
        if os.path.exists(school_logo_path):
            c.drawImage(school_logo_path, 5*mm, card_height-20*mm, width=18*mm, height=18*mm, mask='auto')
        else:
            c.setFillColor(purple)
            c.rect(5*mm, card_height-20*mm, 18*mm, 18*mm, stroke=1, fill=0)

        # QR Code
        c.drawImage(temp_qr_path, card_width-23*mm, 5*mm, width=18*mm, height=18*mm, mask='auto')

        # Member Details
        c.setFillColor(purple)
        c.setFont("Sarabun-Bold", 14)
        c.drawString(28*mm, card_height-16*mm, f"ชื่อ: {member_name}")
        c.setFont("Sarabun", 13)
        c.drawString(28*mm, card_height-24*mm, f"เลขที่: {member_number}")

        # Split Line
        c.setStrokeColor(purple)
        c.setLineWidth(1)
        c.line(28*mm, card_height-28*mm, card_width-28*mm, card_height-28*mm)

        c.save()

        # Clean temp QR file
        if os.path.exists(temp_qr_path):
            os.remove(temp_qr_path)

    @classmethod
    def generate_borrow_history_pdf(cls, filename, records):
        cls.register_fonts()
        doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        
        # Create table headers & rows
        data = [["ชื่อสมาชิก", "ชั้น", "เลขที่", "รหัสหนังสือ", "ชื่อหนังสือ", "วันที่ยืม", "วันที่คืน", "สถานะ"]]
        for record in records:
            status = "คืนแล้ว" if record[7] else "ยังไม่คืน"
            data.append([
                record[0], record[1], record[2], record[3], record[4],
                record[5], record[6], status
            ])
            
        style = ParagraphStyle(name='Sarabun', fontName='Sarabun', fontSize=10, leading=12)
        style_bold = ParagraphStyle(name='Sarabun-Bold', fontName='Sarabun-Bold', fontSize=10, leading=12)

        paragraph_data = [
            [Paragraph(cell, style_bold) if i == 0 else Paragraph(cell, style) for i, cell in enumerate(row)]
            if idx == 0 else
            [Paragraph(str(cell), style) for cell in row]
            for idx, row in enumerate(data)
        ]

        table = Table(paragraph_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.grey),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        
        elements = [table]
        doc.build(elements)

    @classmethod
    def generate_access_history_pdf(cls, filename, records):
        cls.register_fonts()
        doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        
        data = [["ชื่อสมาชิก", "ชั้น", "เลขที่", "เวลา", "การทำรายการ"]]
        for record in records:
            data.append([
                str(record[0]),  # name
                str(record[1]),  # grade
                str(record[2]),  # number
                str(record[3]),  # time
                str(record[4])   # action
            ])
            
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'ThaiTitle',
            parent=styles['Heading1'],
            fontName='Sarabun-Bold',
            fontSize=20,
            leading=24,
            alignment=1, # Center
            spaceAfter=20
        )
        
        style = ParagraphStyle(name='Sarabun', fontName='Sarabun', fontSize=11, leading=14)
        style_bold = ParagraphStyle(name='Sarabun-Bold', fontName='Sarabun-Bold', fontSize=11, leading=14)

        paragraph_data = [
            [Paragraph(cell, style_bold) if i == 0 else Paragraph(cell, style) for i, cell in enumerate(row)]
            if idx == 0 else
            [Paragraph(str(cell), style) for cell in row]
            for idx, row in enumerate(data)
        ]

        table = Table(paragraph_data, colWidths=[120, 50, 50, 150, 100])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1f538d')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('BOTTOMPADDING', (0,0), (-1,0), 8),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))

        title_p = Paragraph("รายงานประวัติการเข้าใช้งานห้องสมุด", title_style)
        elements = [title_p, table]
        doc.build(elements)
