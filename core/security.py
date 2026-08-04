import os
import bcrypt
import base64
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.backends import default_backend
from dotenv import load_dotenv

# Load environmental variables
load_dotenv()

# Secure default or loaded key passphrase
SECRET_PASSPHRASE = os.getenv("SECRET_PASSPHRASE", "default_secret_py_lib_key_change_me_in_dotenv")
SALT_FOR_KEY = os.getenv("SECRET_SALT", "pylibsalt123").encode('utf-8')

# Derive key once using PBKDF2
kdf = PBKDF2HMAC(
    algorithm=hashes.SHA256(),
    length=32,  # AES-256 requires 32 bytes
    salt=SALT_FOR_KEY,
    iterations=100000,
    backend=default_backend()
)
AES_KEY = kdf.derive(SECRET_PASSPHRASE.encode('utf-8'))

def hash_password(password: str) -> str:
    """Hash password using bcrypt."""
    salt = bcrypt.gensalt()
    hashed = bcrypt.hashpw(password.encode('utf-8'), salt)
    return hashed.decode('utf-8')

def verify_password(password: str, hashed: str) -> bool:
    """Verify password using bcrypt."""
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))
    except Exception:
        return False

def encrypt_qr_data(plain_text: str) -> str:
    """Encrypt plain text using AES-256-CBC and return a URL-safe Base64 encoded string containing IV + Ciphertext."""
    # Generate 16 bytes IV
    iv = os.urandom(16)
    
    # Pad plain_text to block size (16 bytes)
    pad_len = 16 - (len(plain_text.encode('utf-8')) % 16)
    padded_data = plain_text.encode('utf-8') + bytes([pad_len] * pad_len)
    
    cipher = Cipher(algorithms.AES(AES_KEY), modes.CBC(iv), backend=default_backend())
    encryptor = cipher.encryptor()
    ciphertext = encryptor.update(padded_data) + encryptor.finalize()
    
    # Prepend IV to ciphertext and encode
    combined = iv + ciphertext
    return base64.urlsafe_b64encode(combined).decode('utf-8')

def decrypt_qr_data(encrypted_text: str) -> str:
    """Decrypt using AES-256-CBC with fallback to standard Base64 decoding for legacy QR codes."""
    try:
        # Try AES decryption first
        combined = base64.urlsafe_b64decode(encrypted_text.encode('utf-8'))
        if len(combined) >= 16:
            iv = combined[:16]
            ciphertext = combined[16:]
            cipher = Cipher(algorithms.AES(AES_KEY), modes.CBC(iv), backend=default_backend())
            decryptor = cipher.decryptor()
            padded_data = decryptor.update(ciphertext) + decryptor.finalize()
            pad_len = padded_data[-1]
            if 1 <= pad_len <= 16:
                plain_bytes = padded_data[:-pad_len]
                return plain_bytes.decode('utf-8')
    except Exception:
        pass
        
    # Fallback to standard base64 decoding for legacy support
    try:
        return base64.b64decode(encrypted_text).decode('utf-8')
    except Exception as e:
        raise ValueError(f"Failed to decrypt QR data: {str(e)}")

