"""
core/config.py
-------------
Centralized configuration using environment variables.
All secrets are sourced from a local .env file — never hardcoded.

Supports runtime updates: the admin UI can POST new values, this module
writes them back to the .env file on disk.
"""

from __future__ import annotations

import json
import os
import re
import sys
import uuid
from datetime import datetime, timezone
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
# Default data root — safe data partition detection
# ---------------------------------------------------------------------------
def _safe_folder_name(name: str) -> str:
    """Sanitize a name for use as a folder/file name."""
    return "".join(c for c in name if c.isalnum() or c in (" ", "-", "_", ".")).strip()


def _detect_safe_data_root() -> Path | None:
    """
    On Windows, find a non-system partition with > 100 GB total space
    to use as the default storage location for bridge data.

    The data stays safe even if Windows is reformatted (system drive wiped).

    Returns a Path to the recommended data root directory, or None if no
    suitable partition is found.
    """
    if sys.platform != "win32":
        return None

    system_drive = os.environ.get("SystemDrive", "C:").rstrip(":\\").upper()

    try:
        import ctypes
        import string

        drives = []
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        for letter in string.ascii_uppercase:
            if bitmask & 1:
                drives.append(f"{letter}:\\")
            bitmask >>= 1

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
                    safe_dir = Path(drive) / "important files(dont delete)" / "app data"
                    safe_dir.mkdir(parents=True, exist_ok=True)
                    return safe_dir
            except Exception:
                continue
    except Exception:
        pass

    return None


def _default_data_root() -> str:
    """
    Compute the default data root directory.

    Priority:
      1. If a safe non-system partition (>100 GB) exists → use it
         at  D:/important files(dont delete)/app data/
      2. Otherwise fall back to BASE_DIR / data/

    Each location gets a subfolder, each office gets a workbook:
      {DATA_ROOT}/{location}/{office}.xlsx
    """
    safe = _detect_safe_data_root()
    if safe is not None:
        return str(safe)
    return str(BASE_DIR / "data")


if not ENV_PATH.is_file():
    import secrets as _secrets
    from cryptography.fernet import Fernet

    _data_root = _default_data_root()
    _lines = [
        f"API_KEY={_secrets.token_hex(32)}",
        f"FERNET_KEY={Fernet.generate_key().decode()}",
        "ALLOWED_ORIGINS=http://localhost:3000",
        f"DATA_ROOT={_data_root}",
        "EXCEL_SHEET_NAME=Sheet1",
        "HOST=127.0.0.1",
        "PORT=8000",
    ]
    with open(str(ENV_PATH), "w", encoding="utf-8") as _f:
        _f.write("\n".join(_lines) + "\n")
    print("✅ Auto-generated .env with secrets.")
    _data_root_info = _data_root
else:
    _data_root_info = os.getenv("DATA_ROOT", os.getenv("EXCEL_DIR", str(BASE_DIR / "data")))

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
    # Root directory where all bridge data is stored.
    # Folder structure: {DATA_ROOT}/{location}/{office}.xlsx
    # Each .xlsx workbook contains one sheet per section.
    _raw_data_root = Path(
        os.getenv("DATA_ROOT", os.getenv("EXCEL_DIR", os.getenv("EXCEL_FILE_PATH", str(BASE_DIR / "data"))))
    ).resolve()
    # If the resolved path is a file (e.g., legacy EXCEL_FILE_PATH pointing to
    # a single .xlsx), use its parent directory as the data root instead.
    if _raw_data_root.is_file():
        DATA_ROOT: Path = _raw_data_root.parent
    else:
        DATA_ROOT: Path = _raw_data_root

    EXCEL_SHEET_NAME: str = os.getenv("EXCEL_SHEET_NAME", "Sheet1")

    @staticmethod
    def get_office_excel_path(office_name: str, location: str = "") -> Path:
        """Return the Excel file path for a given office and location.

        Folder structure: {DATA_ROOT}/{location}/{office}.xlsx

        If location is empty, uses the bridge's configured BRIDGE_LOCATION.
        Falls back to placing the file directly in DATA_ROOT if no location
        is available.
        """
        safe_office = _safe_folder_name(office_name)
        if not safe_office:
            safe_office = "default"

        loc = location.strip() if location else Settings.BRIDGE_LOCATION.strip()
        if loc:
            safe_loc = _safe_folder_name(loc)
            return Settings.DATA_ROOT / safe_loc / f"{safe_office}.xlsx"
        else:
            return Settings.DATA_ROOT / f"{safe_office}.xlsx"

    @staticmethod
    def get_location_dir(location: str = "") -> Path:
        """Return the directory for a given location's data files."""
        loc = location.strip() if location else Settings.BRIDGE_LOCATION.strip()
        if loc:
            safe_loc = _safe_folder_name(loc)
            return Settings.DATA_ROOT / safe_loc
        return Settings.DATA_ROOT

    @staticmethod
    def link_data_root(new_root: str) -> dict:
        """
        Link the bridge to an existing data folder (e.g., from a backup).
        Updates DATA_ROOT in .env and returns status info.

        This is used when installing a bridge on a new PC to restore
        data from a previous installation.
        """
        new_path = Path(new_root).resolve()
        if not new_path.exists():
            return {"status": "error", "message": f"Path does not exist: {new_path}"}
        if not new_path.is_dir():
            return {"status": "error", "message": f"Path is not a directory: {new_path}"}

        old_root = str(Settings.DATA_ROOT)
        set_key(str(ENV_PATH), "DATA_ROOT", str(new_path), quote_mode="auto")
        _reload_env()
        # Update the class-level setting
        Settings.DATA_ROOT = new_path

        # Log the change
        log_path_change(old_root, str(new_path), "link")

        # Count location folders and office workbooks in the new root
        loc_count = 0
        office_count = 0
        try:
            for item in new_path.iterdir():
                if item.is_dir():
                    loc_count += 1
                    office_count += len(list(item.glob("*.xlsx")))
        except Exception:
            pass

        return {
            "status": "ok",
            "message": f"Data root linked to: {new_path}",
            "data_root": str(new_path),
            "location_folders_found": loc_count,
            "office_workbooks_found": office_count,
        }

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
        "DATA_ROOT": str(settings.DATA_ROOT),
        "EXCEL_SHEET_NAME": settings.EXCEL_SHEET_NAME,
        "HOST": settings.HOST,
        "PORT": settings.PORT,
        "BRIDGE_ID": settings.BRIDGE_ID,
        "BRIDGE_LOCATION": settings.BRIDGE_LOCATION,
        "BRIDGE_OFFICE": settings.BRIDGE_OFFICE,
        "BRIDGE_IS_MAIN": settings.BRIDGE_IS_MAIN,
        "BRIDGE_IS_APPROVED": settings.BRIDGE_IS_APPROVED,
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
        "DATA_ROOT",
        "EXCEL_DIR",
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

        # python-dotenv's set_key. Use quote_mode='auto' — paths with spaces
        # need quoting, but booleans like 'true' don't.
        set_key(str(ENV_PATH), key, value, quote_mode="auto")
        results[key] = "updated"

    # Reload os.environ so future Settings reads see the new values
    load_dotenv(ENV_PATH, override=True)

    return results


# ---------------------------------------------------------------------------
# Path-change log
# ---------------------------------------------------------------------------
LOG_FILE_PATH: Path = BASE_DIR / "path_change_log.json"


def log_path_change(
    old_path: str,
    new_path: str,
    action: str,
    extra: dict[str, Any] | None = None,
) -> None:
    """
    Append a change-log entry to the JSON log file.

    Parameters
    ----------
    old_path : str
        The previous Excel file path.
    new_path : str
        The new Excel file path.
    action : str
        One of 'move', 'fresh', or 'same' — describes what happened to the data.
    extra : dict | None
        Optional additional context (e.g. errors).
    """
    entry = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "old_path": old_path,
        "new_path": new_path,
        "action": action,
    }
    if extra:
        entry["extra"] = extra

    # Load existing log (or start fresh)
    log_entries: list[dict[str, Any]] = []
    if LOG_FILE_PATH.is_file():
        try:
            log_entries = json.loads(LOG_FILE_PATH.read_text(encoding="utf-8"))
            if not isinstance(log_entries, list):
                log_entries = []
        except Exception:
            log_entries = []

    log_entries.append(entry)

    # Keep the log manageable — retain last 500 entries
    if len(log_entries) > 500:
        log_entries = log_entries[-500:]

    LOG_FILE_PATH.parent.mkdir(parents=True, exist_ok=True)
    LOG_FILE_PATH.write_text(
        json.dumps(log_entries, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# Public helper — safe path detection (used by admin API)
# ---------------------------------------------------------------------------
def detect_safe_path() -> dict[str, str]:
    """
    Public entry-point for the admin UI to query the recommended safe
    data path.  Returns a dict with `path` and `reason` keys.
    """
    safe = _detect_safe_data_root()
    if safe is not None:
        drive = safe.drive
        return {
            "path": str(safe),
            "reason": f"Using non-system drive {drive} (>100 GB). Data survives Windows reinstall.",
        }
    return {
        "path": str(BASE_DIR / "data"),
        "reason": "No suitable non-system partition found. Using app directory (⚠ not reformat-safe).",
    }
