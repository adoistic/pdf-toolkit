#!/usr/bin/env python3
# ─────────────────────────────────────────────────────────────────────
# Copyright (c) 2025-2026 Thothica Private Limited, Delhi, India.
# All rights reserved.  Proprietary and confidential.
# Unauthorized copying or distribution is strictly prohibited.
# ─────────────────────────────────────────────────────────────────────
"""
license.py — License validation module for PDF Toolkit.

Validates license keys against Firebase Firestore, binds keys to a
machine fingerprint, and caches validation state locally (Fernet-encrypted)
for offline use.

No firebase-admin dependency — uses the Firestore REST API with an API key.
"""

import os
import sys
import json
import time
import base64
import hashlib
import platform

from pathlib import Path
from datetime import datetime, timezone

import requests
from cryptography.fernet import Fernet, InvalidToken

# ── Firebase Configuration ───────────────────────────────────────────
# Reads from .env file (copy .env.example → .env and fill in values).
# API key is safe to embed — it only identifies the project.
# Security is enforced by Firestore rules, not the key.

# Hardcoded for frozen builds — filled automatically by build.py
_EMBEDDED_PROJECT_ID = ""
_EMBEDDED_API_KEY    = ""


def _load_dotenv():
    """Load .env file from the script directory into os.environ."""
    if getattr(sys, "frozen", False):
        env_path = Path(sys.executable).parent / ".env"
    else:
        env_path = Path(__file__).parent / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip())


if _EMBEDDED_PROJECT_ID and _EMBEDDED_API_KEY:
    # Frozen build: use baked-in credentials
    FIREBASE_PROJECT_ID = _EMBEDDED_PROJECT_ID
    FIREBASE_API_KEY    = _EMBEDDED_API_KEY
else:
    # Development: read from .env file
    _load_dotenv()
    FIREBASE_PROJECT_ID = os.getenv("FIREBASE_PROJECT_ID", "")
    FIREBASE_API_KEY    = os.getenv("FIREBASE_API_KEY", "")

FIRESTORE_BASE = (
    f"https://firestore.googleapis.com/v1/"
    f"projects/{FIREBASE_PROJECT_ID}/databases/(default)/documents"
)

# ── Cache Configuration ──────────────────────────────────────────────

APPDATA_DIR = Path(os.getenv("APPDATA", "")) / "PDFToolkit"
CACHE_FILE  = APPDATA_DIR / "license.dat"

RECHECK_INTERVAL = 5 * 60      # Phone home every 5 minutes during active use
OFFLINE_MAX      = 72 * 3600   # Hard lock after 72 hours (3 days) offline

REQUEST_TIMEOUT = 10           # Seconds for HTTP requests


# ── Machine Fingerprinting ───────────────────────────────────────────

def get_machine_id() -> str:
    """Return a stable SHA-256 hash identifying this machine.

    Primary: Windows MachineGuid from the registry.
    Fallback: MAC address + OS identifier.
    """
    try:
        if platform.system() == "Windows":
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Cryptography",
            )
            guid, _ = winreg.QueryValueEx(key, "MachineGuid")
            winreg.CloseKey(key)
            return hashlib.sha256(guid.encode()).hexdigest()
    except Exception:
        pass

    # Fallback
    import uuid as _uuid
    mac = _uuid.getnode()
    raw = f"{mac}-{platform.system()}-{platform.node()}"
    return hashlib.sha256(raw.encode()).hexdigest()


# ── Internal Helpers ─────────────────────────────────────────────────

def _derive_fernet_key(machine_id: str) -> bytes:
    """Derive a Fernet key from the machine ID (deterministic)."""
    digest = hashlib.sha256(machine_id.encode()).digest()
    return base64.urlsafe_b64encode(digest)


def _hash_license_key(license_key: str) -> str:
    """Hash the license key to get the Firestore document ID."""
    return hashlib.sha256(license_key.strip().upper().encode()).hexdigest()


def _parse_firestore_value(field: dict):
    """Extract a Python value from a Firestore REST field dict."""
    if "stringValue" in field:
        return field["stringValue"]
    if "booleanValue" in field:
        return field["booleanValue"]
    if "timestampValue" in field:
        return field["timestampValue"]
    if "integerValue" in field:
        return int(field["integerValue"])
    if "nullValue" in field:
        return None
    return None


def _is_expired(expires_at_str: str) -> bool:
    """Check if an ISO-8601 timestamp is in the past."""
    if not expires_at_str:
        return False  # No expiration means perpetual
    try:
        exp = datetime.fromisoformat(expires_at_str.replace("Z", "+00:00"))
        return exp < datetime.now(timezone.utc)
    except (ValueError, TypeError):
        return False


# ── Local Cache ──────────────────────────────────────────────────────

def load_cache() -> dict | None:
    """Load and decrypt the license cache from disk.

    Returns the cache dict, or None if missing / corrupt / wrong machine.
    """
    if not CACHE_FILE.exists():
        return None
    try:
        machine_id = get_machine_id()
        fernet = Fernet(_derive_fernet_key(machine_id))
        encrypted = CACHE_FILE.read_bytes()
        decrypted = fernet.decrypt(encrypted)
        data = json.loads(decrypted)

        # Verify machine binding (prevents copying the file to another PC)
        if data.get("machine_id") != machine_id:
            return None

        return data
    except (InvalidToken, json.JSONDecodeError, Exception):
        return None


def save_cache(data: dict) -> None:
    """Encrypt and write the license cache to disk."""
    APPDATA_DIR.mkdir(parents=True, exist_ok=True)
    machine_id = get_machine_id()
    data["machine_id"] = machine_id
    fernet = Fernet(_derive_fernet_key(machine_id))
    encrypted = fernet.encrypt(json.dumps(data).encode())
    CACHE_FILE.write_bytes(encrypted)


def clear_cache() -> None:
    """Delete the license cache file."""
    try:
        if CACHE_FILE.exists():
            CACHE_FILE.unlink()
    except OSError:
        pass


# ── Online Validation ────────────────────────────────────────────────

def validate_online(license_key: str) -> dict | None:
    """Validate a license key against Firestore.

    Returns:
        dict  — ``{valid, message, expires_at, key_hash}``
        None  — if the network is unavailable (treat as offline)
    """
    key_hash = _hash_license_key(license_key)
    machine_id = get_machine_id()

    url = f"{FIRESTORE_BASE}/licenses/{key_hash}?key={FIREBASE_API_KEY}"

    try:
        resp = requests.get(url, timeout=REQUEST_TIMEOUT)

        if resp.status_code == 404:
            return {"valid": False, "message": "Invalid license key."}
        if resp.status_code != 200:
            return None  # Server error → treat as offline

        doc = resp.json()
        fields = doc.get("fields", {})

        # Check revocation
        revoked = _parse_firestore_value(fields.get("revoked", {}))
        if revoked:
            return {"valid": False, "message": "This license has been revoked."}

        # Check expiration
        expires_at = _parse_firestore_value(fields.get("expires_at", {}))
        if _is_expired(expires_at):
            return {"valid": False, "message": "This license has expired."}

        # Check machine binding
        bound_machine = _parse_firestore_value(
            fields.get("machine_id", {})
        ) or ""
        if bound_machine and bound_machine != machine_id:
            return {
                "valid": False,
                "message": "This license is bound to another machine.",
            }

        # Bind if unbound
        if not bound_machine:
            _bind_machine(key_hash, machine_id)

        return {
            "valid": True,
            "message": "License valid.",
            "expires_at": expires_at or "",
            "key_hash": key_hash,
        }

    except requests.RequestException:
        return None  # Network error → offline


def _bind_machine(key_hash: str, machine_id: str) -> None:
    """Bind a license to this machine (one-time PATCH)."""
    url = (
        f"{FIRESTORE_BASE}/licenses/{key_hash}"
        f"?key={FIREBASE_API_KEY}"
        f"&updateMask.fieldPaths=machine_id"
    )
    body = {
        "fields": {
            "machine_id": {"stringValue": machine_id},
        }
    }
    try:
        requests.patch(url, json=body, timeout=REQUEST_TIMEOUT)
    except requests.RequestException:
        pass  # Best-effort; the key is still valid even if binding fails


# ── Public API ───────────────────────────────────────────────────────

def check_license() -> tuple[bool, str, bool]:
    """Check license validity — called on every gated request.

    Phones home every RECHECK_INTERVAL (5 min).  If the server is
    unreachable, the cached validation stands for up to OFFLINE_MAX
    (72 h / 3 days).  After that the app locks until connectivity
    is restored.

    Returns:
        (valid, message, needs_key)
        - valid:     True if the app should run
        - message:   Human-readable status text
        - needs_key: True if the UI should show the key entry dialog
    """
    cache = load_cache()

    # ── No cache → prompt for key ────────────────────────────────
    if cache is None:
        return False, "Please enter your license key.", True

    license_key = cache.get("license_key", "")
    expires_at  = cache.get("expires_at", "")

    # ── Cache shows expired → prompt for new key ─────────────────
    if _is_expired(expires_at):
        clear_cache()
        return False, "Your license has expired.", True

    last_validated = cache.get("last_validated", 0)
    age = time.time() - last_validated

    # ── Time to phone home? (every 5 min) ────────────────────────
    if age > RECHECK_INTERVAL:
        result = validate_online(license_key)

        if result is not None:
            # Got an online response
            if result["valid"]:
                cache["last_validated"] = time.time()
                cache["expires_at"] = result.get("expires_at", expires_at)
                save_cache(cache)
                return True, "License valid.", False
            else:
                # Revoked / expired / bound elsewhere → lock NOW
                clear_cache()
                return False, result["message"], True

        # ── Offline: within 72h → allow ──────────────────────────
        if age <= OFFLINE_MAX:
            return True, "License valid (offline mode).", False

        # ── Offline: past 72h → lock ─────────────────────────────
        return (
            False,
            "License check required — please connect to the internet.",
            False,
        )

    # ── Recently validated → allow ───────────────────────────────
    return True, "License valid.", False


def activate_key(license_key: str) -> tuple[bool, str]:
    """Attempt to activate a license key.

    Returns:
        (success, message)
    """
    license_key = license_key.strip()
    if not license_key:
        return False, "Please enter a license key."

    result = validate_online(license_key)

    if result is None:
        return (
            False,
            "Cannot reach the license server. "
            "Please check your internet connection.",
        )

    if not result["valid"]:
        return False, result["message"]

    # Activation succeeded — cache it
    save_cache({
        "license_key": license_key,
        "key_hash": result["key_hash"],
        "expires_at": result.get("expires_at", ""),
        "last_validated": time.time(),
    })

    return True, "License activated successfully!"
