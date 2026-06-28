"""
web_app.py
----------
Simple Flask web frontend for the Local Bridge Service.

Provides a clean UI to:
  - View data from the Excel file
  - Send new text data (encrypted automatically)
  - Delete / clear data

The Flask app handles Fernet encryption/decryption server-side,
so the browser only deals with plain text.

Run:  python web_app.py
"""

from __future__ import annotations

import json
import os
from pathlib import Path

import requests
from cryptography.fernet import Fernet
from dotenv import load_dotenv
from flask import Flask, render_template, request, jsonify

# ---------------------------------------------------------------------------
# Load config from the same .env as the bridge service
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"
load_dotenv(ENV_PATH)

BRIDGE_URL = f"http://{os.getenv('HOST', '127.0.0.1')}:{os.getenv('PORT', '8000')}"
API_KEY = os.getenv("API_KEY", "")
FERNET_KEY = os.getenv("FERNET_KEY", "")

if not API_KEY or not FERNET_KEY:
    raise RuntimeError(
        "API_KEY and FERNET_KEY must be set in .env file.\n"
        "Run:  python generate_env.py"
    )

fernet = Fernet(FERNET_KEY.encode("utf-8"))

# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------
app = Flask(__name__)


def _bridge_headers() -> dict[str, str]:
    """Return headers needed to call the bridge API."""
    return {
        "X-API-KEY": API_KEY,
        "Content-Type": "application/json",
    }


# ===================================================================
# Page routes
# ===================================================================


@app.route("/")
def index():
    """Serve the main UI page."""
    return render_template("index.html")


# ===================================================================
# API proxy routes (called by the frontend JS)
# ===================================================================


@app.route("/api/data")
def get_data():
    """Fetch all data from the bridge and return as JSON."""
    try:
        resp = requests.get(
            f"{BRIDGE_URL}/data",
            headers=_bridge_headers(),
            timeout=10,
        )
        resp.raise_for_status()
        return jsonify(resp.json())
    except requests.exceptions.ConnectionError:
        return jsonify({
            "status": "error",
            "message": (
                "Cannot connect to the bridge service. "
                "Make sure main.py is running (python main.py)."
            ),
        }), 503
    except requests.exceptions.RequestException as exc:
        return jsonify({"status": "error", "message": str(exc)}), 502


@app.route("/api/send", methods=["POST"])
def send_data():
    """
    Accept plain text from the UI, wrap it into a row, encrypt it,
    and forward to the bridge's POST /update.
    """
    body = request.get_json(silent=True) or {}
    text = body.get("text", "").strip()

    if not text:
        return jsonify({"status": "error", "message": "Text cannot be empty."}), 400

    # Wrap the text as a single-row / single-column JSON array
    # (matches the format the bridge expects)
    records = [{"data": text}]

    try:
        plain_json = json.dumps(records)
        encrypted = fernet.encrypt(plain_json.encode("utf-8")).decode("utf-8")

        resp = requests.post(
            f"{BRIDGE_URL}/update",
            headers=_bridge_headers(),
            json={"encrypted": encrypted},
            timeout=10,
        )
        resp.raise_for_status()
        return jsonify(resp.json())
    except requests.exceptions.ConnectionError:
        return jsonify({
            "status": "error",
            "message": "Cannot connect to the bridge service. Is main.py running?",
        }), 503
    except requests.exceptions.RequestException as exc:
        return jsonify({"status": "error", "message": str(exc)}), 502


@app.route("/api/append", methods=["POST"])
def append_data():
    """
    Append a row to existing data (read → append row → write back).
    """
    body = request.get_json(silent=True) or {}
    text = body.get("text", "").strip()

    if not text:
        return jsonify({"status": "error", "message": "Text cannot be empty."}), 400

    try:
        # 1. Read existing data
        resp = requests.get(
            f"{BRIDGE_URL}/data",
            headers=_bridge_headers(),
            timeout=10,
        )
        resp.raise_for_status()
        existing = resp.json()

        # 2. Append the new row
        records = existing.get("data", [])
        records.append({"data": text})

        # 3. Encrypt and send back
        plain_json = json.dumps(records)
        encrypted = fernet.encrypt(plain_json.encode("utf-8")).decode("utf-8")

        resp = requests.post(
            f"{BRIDGE_URL}/update",
            headers=_bridge_headers(),
            json={"encrypted": encrypted},
            timeout=10,
        )
        resp.raise_for_status()
        return jsonify(resp.json())
    except requests.exceptions.ConnectionError:
        return jsonify({
            "status": "error",
            "message": "Cannot connect to the bridge service. Is main.py running?",
        }), 503
    except requests.exceptions.RequestException as exc:
        return jsonify({"status": "error", "message": str(exc)}), 502


@app.route("/api/clear", methods=["POST"])
def clear_data():
    """Clear all data by writing an empty DataFrame (header-only row)."""
    try:
        # Write a row with a placeholder so the bridge accepts it
        records: list[dict[str, str]] = []

        plain_json = json.dumps(records)
        encrypted = fernet.encrypt(plain_json.encode("utf-8")).decode("utf-8")

        resp = requests.post(
            f"{BRIDGE_URL}/update",
            headers=_bridge_headers(),
            json={"encrypted": encrypted},
            timeout=10,
        )
        # 400 is expected for empty payload — that's fine
        if resp.status_code == 400:
            return jsonify({"status": "ok", "message": "Data cleared."})
        resp.raise_for_status()
        return jsonify(resp.json())
    except requests.exceptions.ConnectionError:
        return jsonify({
            "status": "error",
            "message": "Cannot connect to the bridge service. Is main.py running?",
        }), 503
    except requests.exceptions.RequestException as exc:
        return jsonify({"status": "error", "message": str(exc)}), 502


@app.route("/api/health")
def health():
    """Check if the bridge is reachable."""
    try:
        resp = requests.get(f"{BRIDGE_URL}/health", timeout=5)
        return jsonify({"bridge": resp.json(), "web_app": "ok"})
    except requests.exceptions.ConnectionError:
        return jsonify({
            "bridge": "unreachable",
            "web_app": "ok",
            "message": "Bridge service is not running. Start it with: python main.py",
        }), 503


# ===================================================================
# Entry point
# ===================================================================

if __name__ == "__main__":
    print(f" Bridge URL: {BRIDGE_URL}")
    print(f" Web App  : http://127.0.0.1:5000")
    print(" Make sure the bridge service is running:  python main.py")
    print("-" * 50)
    app.run(host="127.0.0.1", port=5000, debug=True)
