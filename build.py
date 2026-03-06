#!/usr/bin/env python3
# ─────────────────────────────────────────────────────────────────────
# build.py — Build PDF Toolkit as a Windows EXE
#
# Reads .env for Firebase credentials, temporarily embeds them into
# license.py, runs PyInstaller, then restores the original file.
#
# Usage:  python build.py
# Output: dist/PDFToolkit/
# ─────────────────────────────────────────────────────────────────────

import os
import re
import sys
import shutil
import subprocess
from pathlib import Path

ROOT = Path(__file__).parent
LICENSE_PY = ROOT / "license.py"
SPEC_FILE = ROOT / "pdf_toolkit.spec"
ENV_FILE = ROOT / ".env"


def load_env(env_path: Path) -> dict:
    """Parse a .env file into a dict."""
    env = {}
    if not env_path.exists():
        return env
    for line in env_path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        env[key.strip()] = value.strip()
    return env


def embed_credentials(license_path: Path, project_id: str, api_key: str) -> str:
    """Replace the empty embedded constants in license.py with real values.

    Returns the original file content (for restoration).
    """
    original = license_path.read_text(encoding="utf-8")

    patched = re.sub(
        r'^_EMBEDDED_PROJECT_ID\s*=\s*""',
        f'_EMBEDDED_PROJECT_ID = "{project_id}"',
        original,
        flags=re.MULTILINE,
    )
    patched = re.sub(
        r'^_EMBEDDED_API_KEY\s*=\s*""',
        f'_EMBEDDED_API_KEY    = "{api_key}"',
        patched,
        flags=re.MULTILINE,
    )

    if patched == original:
        print("WARNING: Could not find empty credential placeholders in license.py")
        print("         Make sure _EMBEDDED_PROJECT_ID and _EMBEDDED_API_KEY are set to \"\"")
        sys.exit(1)

    license_path.write_text(patched, encoding="utf-8")
    return original


def restore_file(license_path: Path, original_content: str):
    """Restore license.py to its original state."""
    license_path.write_text(original_content, encoding="utf-8")


def run_pyinstaller():
    """Run PyInstaller with the spec file."""
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        str(SPEC_FILE),
    ]
    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=str(ROOT))
    return result.returncode


def main():
    print("=" * 60)
    print("  PDF Toolkit — Build Script")
    print("=" * 60)

    # 1. Read .env
    env = load_env(ENV_FILE)
    project_id = env.get("FIREBASE_PROJECT_ID", "")
    api_key = env.get("FIREBASE_API_KEY", "")

    if not project_id or not api_key:
        print(f"\nERROR: Missing Firebase credentials in {ENV_FILE}")
        print("       FIREBASE_PROJECT_ID and FIREBASE_API_KEY must be set.")
        sys.exit(1)

    print(f"\n  Project ID: {project_id}")
    print(f"  API Key:    {api_key[:10]}...")

    # 2. Embed credentials into license.py
    print("\n  Embedding credentials into license.py...")
    original_license = embed_credentials(LICENSE_PY, project_id, api_key)

    try:
        # 3. Run PyInstaller
        print("\n  Running PyInstaller...\n")
        rc = run_pyinstaller()

        if rc != 0:
            print(f"\n  PyInstaller FAILED (exit code {rc})")
            sys.exit(rc)

        print("\n" + "=" * 60)
        print("  BUILD SUCCESSFUL")
        print(f"  Output: {ROOT / 'dist' / 'PDFToolkit'}")
        print("=" * 60)

    finally:
        # 4. Always restore license.py
        print("\n  Restoring license.py to original state...")
        restore_file(LICENSE_PY, original_license)

    # Summary
    dist_dir = ROOT / "dist" / "PDFToolkit"
    if dist_dir.exists():
        main_exe = dist_dir / "PDFToolkit.exe"
        gui_exe = dist_dir / "PDFToolkitGUI.exe"
        helper_exe = dist_dir / "_run_one.exe"
        internal = dist_dir / "_internal"
        templates = internal / "templates"
        static = internal / "static"
        print(f"\n  PDFToolkit.exe:    {'OK' if main_exe.exists() else 'MISSING'}")
        print(f"  PDFToolkitGUI.exe: {'OK' if gui_exe.exists() else 'MISSING'}")
        print(f"  _run_one.exe:      {'OK' if helper_exe.exists() else 'MISSING'}")
        print(f"  templates/:        {'OK' if templates.exists() else 'MISSING'}")
        print(f"  static/:           {'OK' if static.exists() else 'MISSING'}")


if __name__ == "__main__":
    main()
