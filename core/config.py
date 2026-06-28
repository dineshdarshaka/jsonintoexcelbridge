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
import sys
import uuid
from pathlib import Path
from typing import Any

from dotenv import load_dotenv, set_key

# ---------------------------------------------------------------------------
# Load .env from the project root (auto-create if missing)
# ---------------------------------------------------------------------------
# When frozen (PyInstaller .exe), use the exe's directory for .env / data files.
# Otherwise use the project root (parent of core/).
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BASE_DIR / ".env"


# ---------------------------------------------------------------------------
# Default Excel file path — safe data partition detection
# ---------------------------------------------------------------------------
def _detect_safe_data_partition() -> Path | None:
    """
    On Windows, find a non-system partition with > 100 GB total space
    to use as the default storage location for the Excel data file.

    The data stays safe even if Windows is reformatted (system drive wiped).

    Returns a Path to the recommended Excel file, or None if no suitable
    partition is found.
    """
    if sys.platform != "win32":
        return None

    # Get the system drive letter
    system_drive = os.environ.get("SystemDrive", "C:").rstrip(":\\").upper()

    # Read the bridge location from env (may have been set via approval callback)
    bridge_location = os.getenv("BRIDGE_LOCATION", "").strip()

    try:
        import ctypes
        import string

        # Get all logical drives
        drives = []
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        for letter in string.ascii_uppercase:
            if bitmask & 1:
                drives.append(f"{letter}:\\")
            bitmask >>= 1

        # Check each drive — prefer non-system, fixed, > 100 GB
        for drive in drives:
            drive_letter = drive[0].upper()
            if drive_letter == system_drive:
                continue

            drv_type = ctypes.windll.kernel32.GetDriveTypeW(drive)
            if drv_type != 3:  # DRIVE_FIXED = 3
                continue

            try:
                total_bytes = ctypes.c_ulonglong(0)
                ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                    ctypes.c_wchar_p(drive), None,
                    ctypes.byref(total_bytes), None,
                )
                total_gb = total_bytes.value / (1024 ** 3)
                if total_gb >= 100:
                    # Found a safe partition
                    safe_dir = Path(drive) / "important files(dont delete)" / "app data"
                    safe_dir.mkdir(parents=True, exist_ok=True)

                    if bridge_location:
                        filename = f"{bridge_location}.xlsx"
                    else:
                        filename = "data.xlsx"

                    return safe_dir / filename
            except Exception:
                continue
    except Exception:
        pass

    return None


def _default_excel_path() -> str:
    """
    Compute the default Excel file path.

    Priority:
      1. If a safe non-system partition (>100 GB) exists → use it
         at  D:/important files(dont delete)/app data/{location}.xlsx
      2. Otherwise fall back to BASE_DIR / data.xlsx
    """
    safe = _detect_safe_data_partition()
    if safe is not None:
        return str(safe)
    return str(BASE_DIR / "data.xlsx")


if not ENV_PATH.is_file():
    import secrets as _secrets
    from cryptography.fernet import Fernet

    _excel_path = _default_excel_path()
    _lines = [
        f"API_KEY={_secrets.token_hex(32)}",
        f"FERNET_KEY={Fernet.generate_key().decode()}",
        "ALLOWED_ORIGINS=http://localhost:3000",
        f"EXCEL_FILE_PATH={_excel_path}",
        "EXCEL_SHEET_NAME=Sheet1",
        "HOST=127.0.0.1",
        "PORT=8000",
    ]
    with open(str(ENV_PATH), "w", encoding="utf-8") as _f:
        _f.write("\n".join(_lines) + "\n")
    print("✅ Auto-generated .env with secrets.")
    _excel_info = _excel_path
else:
    _excel_info = os.getenv("EXCEL_FILE_PATH", "data.xlsx")

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

    # -- Bridge Status ------------------------------------------------------
    # Set by the central app via POST /api/approval when admin approves/rejects.
    BRIDGE_IS_APPROVED: bool = os.getenv("BRIDGE_IS_APPROVED", "false").lower() in ("true", "1", "yes")
    BRIDGE_IS_MAIN: bool = os.getenv("BRIDGE_IS_MAIN", "false").lower() in ("true", "1", "yes")
    BRIDGE_LOCATION: str = os.getenv("BRIDGE_LOCATION", "")
    BRIDGE_OFFICE: str = os.getenv("BRIDGE_OFFICE", "")


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
        "BRIDGE_IS_MAIN",
        "BRIDGE_IS_APPROVED",
        "BRIDGE_LOCATION",
        "BRIDGE_OFFICE",
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

        # python-dotenv's set_key wraps values in single quotes by default.
        # Use quote_mode='never' to write bare values (e.g. true, not 'true').
        set_key(str(ENV_PATH), key, value, quote_mode="never")
        results[key] = "updated"

    # Reload os.environ so future Settings reads see the new values
    load_dotenv(ENV_PATH, override=True)

    return results


# ---------------------------------------------------------------------------
# Public helper — safe path detection (used by admin API)
# ---------------------------------------------------------------------------
def detect_safe_path() -> dict[str, str]:
    """
    Public entry-point for the admin UI to query the recommended safe
    data path.  Returns a dict with `path` and `reason` keys.
    """
    safe = _detect_safe_data_partition()
    if safe is not None:
        drive = safe.drive
        return {
            "path": str(safe),
            "reason": f"Using non-system drive {drive} (>100 GB). Data survives Windows reinstall.",
        }
    return {
        "path": str(BASE_DIR / "data.xlsx"),
        "reason": "No suitable non-system partition found. Using app directory (⚠ not reformat-safe).",
    }
