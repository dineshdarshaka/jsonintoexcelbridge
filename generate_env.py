"""Quick script to generate a .env file with secrets."""
import secrets
from cryptography.fernet import Fernet

lines = [
    f"API_KEY={secrets.token_hex(32)}",
    f"FERNET_KEY={Fernet.generate_key().decode()}",
    "ALLOWED_ORIGINS=http://localhost:3000",
    "EXCEL_DIR=data",
    "EXCEL_SHEET_NAME=Sheet1",
    "HOST=127.0.0.1",
    "PORT=8000",
]

with open(".env", "w", encoding="utf-8") as f:
    f.write("\n".join(lines) + "\n")

print("✅ .env created with generated secrets.")
print("\n".join(lines))
