# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for CV Research Experience Manager.

Build (one-file, default):
    pyinstaller --clean --noconfirm cv_manager.spec

Build (one-folder):
    set BUILD_MODE=onedir
    pyinstaller --clean --noconfirm cv_manager.spec

The executable is placed at the project root next to CV_Manager.bat (via --distpath .).

Environment variables:
    CONSOLE_MODE  — "1" to show a console window (debug builds).
    BUILD_MODE    — "onedir" for one-folder build; anything else is one-file.
"""

import sys
import os
import time

block_cipher = None

SHOW_CONSOLE = os.environ.get("CONSOLE_MODE", "0").strip() == "1"
BUILD_MODE = os.environ.get("BUILD_MODE", "onefile").strip().lower()

# ---------------------------------------------------------------------------
# Version resource — read build number so every rebuild gets a new version,
# which forces Windows to invalidate its cached icon for the .exe.
# ---------------------------------------------------------------------------
_build_number_path = os.path.join("build", "build_number.txt")
if os.path.exists(_build_number_path):
    with open(_build_number_path, encoding="utf-8") as _f:
        _BUILD_NUMBER = int(_f.read().strip() or "0")
else:
    _BUILD_NUMBER = int(time.time()) % 65535

_APP_VERSION_STR = f"1.2.0.{_BUILD_NUMBER}"
print(f"[spec] Build number: {_BUILD_NUMBER}  version: {_APP_VERSION_STR}")

from PyInstaller.utils.win32.versioninfo import (
    VSVersionInfo,
    FixedFileInfo,
    StringFileInfo,
    StringTable,
    StringStruct,
    VarFileInfo,
    VarStruct,
)

_version_info = VSVersionInfo(
    ffi=FixedFileInfo(
        filevers=(1, 2, 0, _BUILD_NUMBER),
        prodvers=(1, 2, 0, _BUILD_NUMBER),
        mask=0x3F,
        flags=0x0,
        OS=0x40004,
        fileType=0x1,
        subtype=0x0,
        date=(0, 0),
    ),
    kids=[
        StringFileInfo(
            [
                StringTable(
                    "040904B0",
                    [
                        StringStruct("CompanyName", "CenExel"),
                        StringStruct(
                            "FileDescription",
                            "CV Research Experience Manager",
                        ),
                        StringStruct("FileVersion", _APP_VERSION_STR),
                        StringStruct("InternalName", "CV_Manager"),
                        StringStruct("OriginalFilename", "CV_Manager.exe"),
                        StringStruct(
                            "ProductName",
                            "CV Research Experience Manager",
                        ),
                        StringStruct("ProductVersion", _APP_VERSION_STR),
                    ],
                )
            ]
        ),
        VarFileInfo([VarStruct("Translation", [0x0409, 0x04B0])]),
    ],
)

# ---------------------------------------------------------------------------
# Data files
# ---------------------------------------------------------------------------
datas_list = []
if os.path.exists("data/config.json"):
    datas_list.append(("data/config.json", "data"))
if os.path.exists("build/assets/app.ico"):
    datas_list.append(("build/assets/app.ico", "assets"))

icon_file = "build/assets/app.ico" if os.path.exists("build/assets/app.ico") else None

# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------
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
        "appid",
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

# ---------------------------------------------------------------------------
# EXE — common arguments
# ---------------------------------------------------------------------------
exe_args = dict(
    name="CV_Manager",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=SHOW_CONSOLE,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    uac_admin=False,
    version=_version_info,
)

if icon_file is not None:
    exe_args["icon"] = [icon_file]

# For one-file builds, use a stable runtime tmpdir next to the exe so that
# the extraction path does not change between runs.  This stabilises taskbar
# grouping because Windows sees the same child-process path every time.
if BUILD_MODE != "onedir":
    exe_args["runtime_tmpdir"] = "_cv_manager_runtime"

# ---------------------------------------------------------------------------
# One-file build (default) — everything bundled into the EXE
# ---------------------------------------------------------------------------
if BUILD_MODE != "onedir":
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        [],
        **exe_args,
    )


# ---------------------------------------------------------------------------
# One-folder build — exe + supporting files in CV_Manager/ (project root)
# ---------------------------------------------------------------------------
else:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        **exe_args,
    )

    coll = COLLECT(
        exe,
        a.binaries,
        a.zipfiles,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name="CV_Manager",
    )
