# -*- mode: python ; coding: utf-8 -*-
# ─────────────────────────────────────────────────────────────────────
# PyInstaller spec for PDF Toolkit
# Builds two EXEs: PDFToolkit.exe (main app) + _run_one.exe (helper)
# ─────────────────────────────────────────────────────────────────────

import os
from pathlib import Path

block_cipher = None
ROOT = os.path.dirname(os.path.abspath(SPEC))

# ── Main app: PDFToolkit.exe ─────────────────────────────────────────

main_a = Analysis(
    [os.path.join(ROOT, 'app.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[
        (os.path.join(ROOT, 'templates'), 'templates'),
        (os.path.join(ROOT, 'static'), 'static'),
    ],
    hiddenimports=[
        'license',
        'flask',
        'flask.json',
        'jinja2',
        'markupsafe',
        'werkzeug',
        'requests',
        'cryptography',
        'cryptography.fernet',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['unittest', 'test'],
    noarchive=False,
    cipher=block_cipher,
)

main_pyz = PYZ(main_a.pure, cipher=block_cipher)

main_exe = EXE(
    main_pyz,
    main_a.scripts,
    [],
    exclude_binaries=True,
    name='PDFToolkit',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,   # Keep console for debugging; change to False for release
    icon=None,
)

# ── Helper: _run_one.exe ────────────────────────────────────────────

helper_a = Analysis(
    [os.path.join(ROOT, '_run_one.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[],
    hiddenimports=[
        'add_toc',
        'pdf_to_docx',
        'fitz',
        'fitz.fitz',
        'docx',
        'docx.shared',
        'docx.enum.text',
        'docx.enum.table',
        'docx.oxml',
        'docx.oxml.ns',
        'lxml',
        'lxml.etree',
        'PIL',
        'PIL.Image',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['unittest', 'test'],
    noarchive=False,
    cipher=block_cipher,
)

helper_pyz = PYZ(helper_a.pure, cipher=block_cipher)

helper_exe = EXE(
    helper_pyz,
    helper_a.scripts,
    [],
    exclude_binaries=True,
    name='_run_one',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,   # Needs console for stdout communication
    icon=None,
)

# ── GUI app: PDFToolkitGUI.exe ─────────────────────────────────────

gui_a = Analysis(
    [os.path.join(ROOT, 'gui.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['unittest', 'test'],
    noarchive=False,
    cipher=block_cipher,
)

gui_pyz = PYZ(gui_a.pure, cipher=block_cipher)

gui_exe = EXE(
    gui_pyz,
    gui_a.scripts,
    [],
    exclude_binaries=True,
    name='PDFToolkitGUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,   # GUI app — no console window
    icon=None,
)

# ── Collect into single folder ──────────────────────────────────────

coll = COLLECT(
    main_exe, main_a.binaries, main_a.datas,
    helper_exe, helper_a.binaries, helper_a.datas,
    gui_exe, gui_a.binaries, gui_a.datas,
    strip=False,
    upx=False,
    name='PDFToolkit',
)
