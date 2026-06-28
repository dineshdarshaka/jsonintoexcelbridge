"""
core/config.py
-------------
Centralized configuration using environment variables.
All secrets are sourced from a local .env file — never hardcoded.

Supports runtime updates: the admin UI can POST new values, this module
writes them back to the .env file on disk.
"""

from __future__ import annotations

import os
import re
import uuid
from pathlib import Path
from typing import Any

from dotenv import load_dotenv, set_key

# ---------------------------------------------------------------------------
# Load .env from the project root
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BASE_DIR / ".env"

load_dotenv(ENV_PATH)

# re-load after potential set_key writes
def _reload_env():
    """Re-read .env so set_key changes are visible to os.getenv."""
    load_dotenv(ENV_PATH, override=True)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _mask_secret(value: str, show: int = 4) -> str:
    """Return a masked version of a secret, showing only first/last N chars."""
    if len(value) <= show * 2:
        return "*" * len(value)
    return value[:show] + "*" * (len(value) - show * 2) + value[-show:]


def _serialize_list(lst: list[str]) -> str:
    """Serialize a list of strings as comma-separated."""
    return ",".join(lst)


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
class Settings:
    """
    Application-wide settings pulled exclusively from the environment.
    """

    # -- Authentication ----------------------------------------------------
    API_KEY: str = os.getenv("API_KEY", "")
    if not API_KEY:
        raise RuntimeError("API_KEY is not set in the .env file")

    # -- Encryption ---------------------------------------------------------
    FERNET_KEY: str = os.getenv("FERNET_KEY", "")
    if not FERNET_KEY:
        raise RuntimeError("FERNET_KEY is not set in the .env file")

    # -- CORS ---------------------------------------------------------------
    ALLOWED_ORIGINS: list[str] = [
        origin.strip()
        for origin in os.getenv("ALLOWED_ORIGINS", "http://localhost:3000").split(",")
        if origin.strip()
    ]

    # -- File paths ---------------------------------------------------------
    EXCEL_FILE_PATH: Path = Path(
        os.getenv("EXCEL_FILE_PATH", str(BASE_DIR / "data.xlsx"))
    ).resolve()

    EXCEL_SHEET_NAME: str = os.getenv("EXCEL_SHEET_NAME", "Sheet1")

    # -- Server -------------------------------------------------------------
    HOST: str = os.getenv("HOST", "127.0.0.1")
    PORT: int = int(os.getenv("PORT", "8000"))

    # -- Bridge Identity ----------------------------------------------------
    # Permanent unique ID — generated once on first startup, saved to .env,
    # and reused forever regardless of IP/hostname changes.
    BRIDGE_ID: str = os.getenv("BRIDGE_ID", "").strip().strip("'\"").strip()
    if not BRIDGE_ID:
        BRIDGE_ID = str(uuid.uuid4())
        set_key(str(ENV_PATH), "BRIDGE_ID", BRIDGE_ID)
        _reload_env()
        print(f"  🔑 Generated permanent Bridge ID: {BRIDGE_ID}")


# Singleton instance
settings = Settings()


# ---------------------------------------------------------------------------
# Runtime config read / write (used by the admin API)
# ---------------------------------------------------------------------------

def get_config_snapshot() -> dict[str, Any]:
    """
    Return a JSON-safe snapshot of all current settings.
    Secrets are *masked* — never returned in full over the API.
    """
    return {
        "API_KEY": _mask_secret(settings.API_KEY),
        "API_KEY_length": len(settings.API_KEY),
        "FERNET_KEY": _mask_secret(settings.FERNET_KEY),
        "FERNET_KEY_length": len(settings.FERNET_KEY),
        "ALLOWED_ORIGINS": _serialize_list(settings.ALLOWED_ORIGINS),
        "EXCEL_FILE_PATH": str(settings.EXCEL_FILE_PATH),
        "EXCEL_SHEET_NAME": settings.EXCEL_SHEET_NAME,
        "HOST": settings.HOST,
        "PORT": settings.PORT,
        "BRIDGE_ID": settings.BRIDGE_ID,
    }


def update_env_file(updates: dict[str, str | int]) -> dict[str, str]:
    """
    Persist new configuration values to the .env file on disk.

    Only recognised keys are written; unknown keys are silently ignored.
    Secrets (API_KEY, FERNET_KEY) are written as-is (no masking).

    Returns a dict of {key: message} for each key that was updated.
    """
    ALLOWED_KEYS = {
        "API_KEY",
        "FERNET_KEY",
        "ALLOWED_ORIGINS",
        "EXCEL_FILE_PATH",
        "EXCEL_SHEET_NAME",
        "HOST",
        "PORT",
        "BRIDGE_ID",
    }

    results: dict[str, str] = {}

    for key, raw_value in updates.items():
        if key not in ALLOWED_KEYS:
            results[key] = "skipped — unknown key"
            continue

        value = str(raw_value).strip()
        if not value:
            results[key] = "skipped — empty value not allowed"
            continue

        # Validate specific keys
        if key == "PORT":
            try:
                int(value)
            except ValueError:
                results[key] = "skipped — must be an integer"
                continue

        set_key(str(ENV_PATH), key, value)
        results[key] = "updated"

    # Reload os.environ so future Settings reads see the new values
    load_dotenv(ENV_PATH, override=True)

    return results
