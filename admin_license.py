#!/usr/bin/env python3
# ─────────────────────────────────────────────────────────────────────
# Copyright (c) 2025-2026 Thothica Private Limited, Delhi, India.
# All rights reserved.  Proprietary and confidential.
# Unauthorized copying or distribution is strictly prohibited.
# ─────────────────────────────────────────────────────────────────────
"""
admin_license.py — Admin CLI for managing PDF Toolkit licenses.

Usage:
    python admin_license.py generate --days 365 --note "Customer name"
    python admin_license.py revoke  PDFT-XXXX-XXXX-XXXX
    python admin_license.py unbind  PDFT-XXXX-XXXX-XXXX
    python admin_license.py extend  PDFT-XXXX-XXXX-XXXX --days 180
    python admin_license.py list

Requires:
    pip install firebase-admin
    Place your service_account.json in the same directory.
"""

import argparse
import hashlib
import secrets
import string
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import firebase_admin
from firebase_admin import credentials, firestore


# ── Firebase Init ────────────────────────────────────────────────────

SERVICE_ACCOUNT = Path(__file__).parent / "service_account.json"


def init_firebase():
    if not firebase_admin._apps:
        if not SERVICE_ACCOUNT.exists():
            print(f"ERROR: {SERVICE_ACCOUNT} not found.")
            print("Download it from Firebase Console → Project Settings → Service Accounts.")
            sys.exit(1)
        cred = credentials.Certificate(str(SERVICE_ACCOUNT))
        firebase_admin.initialize_app(cred)
    return firestore.client()


# ── Helpers ──────────────────────────────────────────────────────────

def generate_key_string() -> str:
    """Generate a random license key: PDFT-XXXX-XXXX-XXXX."""
    chars = string.ascii_uppercase + string.digits
    groups = ["PDFT"]
    for _ in range(3):
        groups.append("".join(secrets.choice(chars) for _ in range(4)))
    return "-".join(groups)


def hash_key(key: str) -> str:
    """SHA-256 hash of the uppercase key → Firestore document ID."""
    return hashlib.sha256(key.strip().upper().encode()).hexdigest()


def format_ts(val) -> str:
    """Format a Firestore timestamp or string for display."""
    if val is None:
        return "N/A"
    if hasattr(val, "strftime"):
        return val.strftime("%Y-%m-%d %H:%M UTC")
    return str(val)


# ── Commands ─────────────────────────────────────────────────────────

def cmd_generate(args):
    db = init_firebase()
    key = generate_key_string()
    key_hash = hash_key(key)
    expires_at = datetime.now(timezone.utc) + timedelta(days=args.days)

    db.collection("licenses").document(key_hash).set({
        "key_preview": key[-4:],
        "machine_id": "",
        "expires_at": expires_at,
        "revoked": False,
        "created_at": firestore.SERVER_TIMESTAMP,
        "note": args.note or "",
    })

    print()
    print(f"  License Key : {key}")
    print(f"  Expires     : {expires_at.strftime('%Y-%m-%d')}")
    print(f"  Note        : {args.note or '(none)'}")
    print(f"  Doc ID      : {key_hash}")
    print()
    print("  Send this key to the customer. They enter it in the app on first launch.")
    print()


def cmd_revoke(args):
    db = init_firebase()
    key_hash = hash_key(args.key)
    doc = db.collection("licenses").document(key_hash).get()
    if not doc.exists:
        print(f"ERROR: No license found for key {args.key}")
        sys.exit(1)
    db.collection("licenses").document(key_hash).update({"revoked": True})
    print(f"  Revoked: {args.key}")
    print("  The app will lock the next time it connects to the internet.")


def cmd_unbind(args):
    db = init_firebase()
    key_hash = hash_key(args.key)
    doc = db.collection("licenses").document(key_hash).get()
    if not doc.exists:
        print(f"ERROR: No license found for key {args.key}")
        sys.exit(1)
    db.collection("licenses").document(key_hash).update({"machine_id": ""})
    print(f"  Unbound: {args.key}")
    print("  The key can now be activated on a different machine.")


def cmd_extend(args):
    db = init_firebase()
    key_hash = hash_key(args.key)
    doc = db.collection("licenses").document(key_hash).get()
    if not doc.exists:
        print(f"ERROR: No license found for key {args.key}")
        sys.exit(1)

    data = doc.to_dict()
    old_exp = data.get("expires_at")
    base = old_exp if old_exp and hasattr(old_exp, "date") else datetime.now(timezone.utc)
    new_exp = base + timedelta(days=args.days)

    db.collection("licenses").document(key_hash).update({
        "expires_at": new_exp,
        "revoked": False,  # Unrevoke if extending
    })
    print(f"  Extended: {args.key}")
    print(f"  New expiration: {new_exp.strftime('%Y-%m-%d')}")


def cmd_list(args):
    db = init_firebase()
    docs = db.collection("licenses").stream()

    count = 0
    print()
    print(f"  {'Preview':>8}  {'Status':<10} {'Bound':<8} {'Expires':<22} Note")
    print(f"  {'─' * 8}  {'─' * 10} {'─' * 8} {'─' * 22} {'─' * 20}")

    for doc in docs:
        d = doc.to_dict()
        preview = f"...{d.get('key_preview', '????')}"
        status = "REVOKED" if d.get("revoked") else "ACTIVE"
        bound = "BOUND" if d.get("machine_id") else "FREE"
        exp = format_ts(d.get("expires_at"))
        note = d.get("note", "")
        print(f"  {preview:>8}  {status:<10} {bound:<8} {exp:<22} {note}")
        count += 1

    print()
    print(f"  Total: {count} license(s)")
    print()


# ── Main ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="PDF Toolkit — License Manager (Admin)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command")

    # generate
    gen = sub.add_parser("generate", help="Generate a new license key")
    gen.add_argument("--days", type=int, default=365, help="Validity in days (default: 365)")
    gen.add_argument("--note", type=str, default="", help="Customer name or note")

    # revoke
    rev = sub.add_parser("revoke", help="Revoke a license key")
    rev.add_argument("key", help="The license key (e.g. PDFT-XXXX-XXXX-XXXX)")

    # unbind
    ub = sub.add_parser("unbind", help="Remove machine binding from a key")
    ub.add_argument("key", help="The license key")

    # extend
    ext = sub.add_parser("extend", help="Extend expiration of a key")
    ext.add_argument("key", help="The license key")
    ext.add_argument("--days", type=int, default=365, help="Additional days (default: 365)")

    # list
    sub.add_parser("list", help="List all licenses")

    args = parser.parse_args()

    commands = {
        "generate": cmd_generate,
        "revoke":   cmd_revoke,
        "unbind":   cmd_unbind,
        "extend":   cmd_extend,
        "list":     cmd_list,
    }

    if args.command in commands:
        commands[args.command](args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
