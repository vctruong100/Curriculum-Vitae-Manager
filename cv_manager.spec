# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for CV Research Experience Manager.

Build with:
    pyinstaller cv_manager.spec

This produces a single-file executable per OS.
On macOS/Linux where Calibri is unavailable, the app falls back to
the system default proportional font (see docx_handler.py FONT_NAME).
"""

import sys
import os

block_cipher = None

a = Analysis(
    ['src/launcher.pyw'],
    pathex=['src'],
    binaries=[],
    datas=[
        # Bundle default config if present
        ('data/config.json', 'data') if os.path.exists('data/config.json') else (None, None),
        ('src', 'src'),  # Bundle source modules
    ],
    hiddenimports=[
        'docx',
        'openpyxl',
        'rapidfuzz',
        'rapidfuzz.fuzz',
        'rapidfuzz.process',
        'rapidfuzz.distance',
        'sqlite3',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.simpledialog',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Explicitly exclude network modules
        'requests',
        'urllib3',
        'httpx',
        'aiohttp',
        'pip',
        'setuptools',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Filter out None entries from datas
a.datas = [(d, s, t) for d, s, t in a.datas if d is not None]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CV_Manager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Windowed app (no console)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',  # Uncomment and provide icon file
)
