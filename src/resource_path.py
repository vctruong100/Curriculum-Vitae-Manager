"""
Resource path resolution for frozen (PyInstaller) and source execution modes.

When running from source, paths resolve relative to the project root (parent of src/).
When running as a frozen .exe, paths resolve relative to the directory containing
the executable, ensuring ./data/ and ./assets/ sit next to the .exe.
"""

import sys
import os
from pathlib import Path


def is_frozen() -> bool:
    return getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS")


def get_bundle_dir() -> Path:
    if is_frozen():
        return Path(sys._MEIPASS)
    return Path(__file__).parent.parent.resolve()


def get_application_dir() -> Path:
    if is_frozen():
        return Path(sys.executable).parent.resolve()
    return Path(__file__).parent.parent.resolve()


def resource_path(relative: str) -> Path:
    return get_bundle_dir() / relative


def writable_path(relative: str) -> Path:
    return get_application_dir() / relative
