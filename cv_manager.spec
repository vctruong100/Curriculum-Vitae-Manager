# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for CV Research Experience Manager.

Build:
    pyinstaller --clean --noconfirm cv_manager.spec

The executable is placed in the project root next to CV_Manager.bat.

Toggle console visibility with CONSOLE_MODE env var:
    SET CONSOLE_MODE=1
    pyinstaller --clean --noconfirm cv_manager.spec
"""

import sys
import os

block_cipher = None

SHOW_CONSOLE = os.environ.get("CONSOLE_MODE", "0").strip() == "1"
DIST_PATH = os.path.abspath(".")

datas_list = []
if os.path.exists("data/config.json"):
    datas_list.append(("data/config.json", "data"))
if os.path.exists("build/assets/app.ico"):
    datas_list.append(("build/assets/app.ico", "assets"))

icon_file = "build/assets/app.ico" if os.path.exists("build/assets/app.ico") else None

a = Analysis(
    ["src/main.py"],
    pathex=["src"],
    binaries=[],
    datas=datas_list,
    hiddenimports=[
        "docx",
        "docx.opc",
        "docx.opc.constants",
        "docx.oxml",
        "docx.oxml.ns",
        "openpyxl",
        "rapidfuzz",
        "rapidfuzz.fuzz",
        "rapidfuzz.process",
        "rapidfuzz.distance",
        "rapidfuzz.distance.Levenshtein",
        "rapidfuzz.distance.DamerauLevenshtein",
        "rapidfuzz.utils",
        "sqlite3",
        "tkinter",
        "tkinter.ttk",
        "tkinter.filedialog",
        "tkinter.messagebox",
        "tkinter.simpledialog",
        "config",
        "gui",
        "processor",
        "docx_handler",
        "excel_parser",
        "database",
        "models",
        "normalizer",
        "import_export",
        "logger",
        "validators",
        "migrations",
        "permissions",
        "offline_guard",
        "error_handler",
        "progress_dialog",
        "tooltip_text",
        "resource_path",
        "instance_lock",
        "undo_buffer",
        "update_checker",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "requests",
        "urllib3",
        "httpx",
        "aiohttp",
        "pip",
        "setuptools",
        "numpy",
        "pandas",
        "matplotlib",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe_args = dict(
    name="CV_Manager",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=SHOW_CONSOLE,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
if icon_file is not None:
    exe_args["icon"] = [icon_file]

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    **exe_args,
)

import shutil
_src_exe = os.path.join("dist", "CV_Manager.exe")
_dst_exe = os.path.join(DIST_PATH, "CV_Manager.exe")
if os.path.exists(_src_exe):
    shutil.copy2(_src_exe, _dst_exe)
    print(f"Copied {_src_exe} -> {_dst_exe}")
