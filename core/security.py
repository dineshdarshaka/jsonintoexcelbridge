"""
core/security.py
----------------
Authentication and encryption utilities for the Local Bridge Service.

- API key validation via the X-API-KEY header.
- Payload encryption / decryption using Fernet (symmetric AES-128-CBC under
  the hood, with HMAC-SHA256 authentication).
"""

from __future__ import annotations

import base64
import secrets

from cryptography.fernet import Fernet, InvalidToken
from fastapi import Request, HTTPException, status

from core.config import settings

# ---------------------------------------------------------------------------
# Fernet cipher (instantiated once at import time)
# ---------------------------------------------------------------------------
_fernet: Fernet = Fernet(settings.FERNET_KEY.encode("utf-8"))


# ===================================================================
# Authentication
# ===================================================================

def verify_api_key(request: Request) -> None:
    """
    Validate the X-API-KEY header against the server-side secret.

    Raises
    ------
    HTTPException(403)
        If the header is missing, empty, or does not match.
    """
    client_key: str | None = request.headers.get("X-API-KEY")
    if not client_key:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Missing X-API-KEY header",
        )

    # Constant-time comparison to mitigate timing attacks
    if not secrets.compare_digest(client_key, settings.API_KEY):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Invalid API key",
        )


# ===================================================================
# Encryption / Decryption
# ===================================================================

def decrypt_payload(encrypted_b64: str) -> str:
    """
    Decrypt a Base64-encoded Fernet ciphertext and return the plaintext
    UTF-8 string.

    Parameters
    ----------
    encrypted_b64 : str
        Base64-encoded ciphertext produced by Fernet.encrypt().

    Returns
    -------
    str
        The decrypted plaintext.

    Raises
    ------
    HTTPException(400)
        If the payload cannot be decrypted or is malformed.
    """
    try:
        # Fernet tokens are already URL-safe base64 strings.  fernet.decrypt()
        # expects the raw token bytes, so we pass the UTF-8 encoded string
        # directly without attempting an intermediate base64 decode.
        token: bytes = encrypted_b64.encode("utf-8")
        plaintext: bytes = _fernet.decrypt(token)
        return plaintext.decode("utf-8")

    except InvalidToken:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Decryption failed — invalid or tampered payload",
        )
    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Decryption error: {exc!s}",
        )


def encrypt_payload(plaintext: str) -> str:
    """
    Encrypt a plaintext string and return the Base64-encoded Fernet token.

    This is a convenience helper — useful for testing or for encrypting
    response data if needed later.
    """
    token: bytes = _fernet.encrypt(plaintext.encode("utf-8"))
    return token.decode("utf-8")


# ===================================================================
# Key generation helper (run manually once)
# ===================================================================

def generate_fernet_key() -> str:
    """
    Generate a fresh Fernet key.

    Run this once from the Python REPL or a script and store the output
    in your .env file as FERNET_KEY.
    """
    return Fernet.generate_key().decode("utf-8")
