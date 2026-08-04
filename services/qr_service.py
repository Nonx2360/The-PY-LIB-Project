# services/qr_service.py
import qrcode
from PIL import Image
from core.security import encrypt_qr_data, decrypt_qr_data
import os

class QRService:
    @staticmethod
    def generate_qr(data: str, save_path: str):
        encrypted = encrypt_qr_data(data)
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(encrypted)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        img.save(save_path)
        return encrypted
        
    @staticmethod
    def decode_qr(encrypted_data: str) -> str:
        return decrypt_qr_data(encrypted_data)
