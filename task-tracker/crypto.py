import os
from cryptography.fernet import Fernet

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SECRET_KEY_PATH = os.path.join(BASE_DIR, "secret.key")


def load_or_create_key():
    if os.path.exists(SECRET_KEY_PATH):
        with open(SECRET_KEY_PATH, "rb") as f:
            return f.read()
    key = Fernet.generate_key()
    with open(SECRET_KEY_PATH, "wb") as f:
        f.write(key)
    return key


_key = load_or_create_key()
_cipher = Fernet(_key)


def encrypt(text: str) -> str:
    if not text:
        return ""
    return _cipher.encrypt(text.encode()).decode()


def decrypt(token: str) -> str:
    if not token:
        return ""
    try:
        return _cipher.decrypt(token.encode()).decode()
    except Exception:
        return "***"
