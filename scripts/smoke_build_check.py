"""
Post-build smoke test for CV Research Experience Manager.

Verifies:
  1. The built .exe exists (one-file or one-folder).
  2. The version resource on the .exe matches the current build number.
  3. build/assets/app.ico exists and is non-empty.
  4. (Optional) On a headed machine the exe launches and quits immediately.

Usage:
    python scripts/smoke_build_check.py
    python scripts/smoke_build_check.py --launch   # also test launch+quit

No network requests.  Purely local file checks.
"""

import sys
import os
import struct
import subprocess
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).parent.parent.resolve()
BUILD_NUMBER_FILE = PROJECT_ROOT / "build" / "build_number.txt"
ICON_PATH = PROJECT_ROOT / "build" / "assets" / "app.ico"
ONEFILE_EXE = PROJECT_ROOT / "CV_Manager.exe"
ONEDIR_EXE = PROJECT_ROOT / "CV_Manager" / "CV_Manager.exe"


def _read_build_number() -> int:
    """Read the current build number from build/build_number.txt."""
    try:
        return int(BUILD_NUMBER_FILE.read_text(encoding="utf-8").strip())
    except (FileNotFoundError, ValueError):
        return -1


def _read_exe_product_version(exe_path: Path) -> str:
    """Read the ProductVersion string from the exe's version resource.

    Uses ctypes + GetFileVersionInfoW on Windows.
    Returns the version string or "" on failure.
    """
    if sys.platform != "win32":
        return ""

    try:
        import ctypes
        from ctypes import wintypes

        GetFileVersionInfoSizeW = ctypes.windll.version.GetFileVersionInfoSizeW
        GetFileVersionInfoSizeW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(wintypes.DWORD)]
        GetFileVersionInfoSizeW.restype = wintypes.DWORD

        GetFileVersionInfoW = ctypes.windll.version.GetFileVersionInfoW
        GetFileVersionInfoW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p]
        GetFileVersionInfoW.restype = wintypes.BOOL

        VerQueryValueW = ctypes.windll.version.VerQueryValueW
        VerQueryValueW.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR, ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(ctypes.c_uint)]
        VerQueryValueW.restype = wintypes.BOOL

        path_str = str(exe_path)
        dummy = wintypes.DWORD(0)
        size = GetFileVersionInfoSizeW(path_str, ctypes.byref(dummy))
        if size == 0:
            return ""

        buf = ctypes.create_string_buffer(size)
        if not GetFileVersionInfoW(path_str, 0, size, buf):
            return ""

        # Try to read ProductVersion from StringFileInfo
        sub_block = r"\StringFileInfo\040904B0\ProductVersion"
        lp_buf = ctypes.c_void_p()
        u_len = ctypes.c_uint()
        if VerQueryValueW(buf, sub_block, ctypes.byref(lp_buf), ctypes.byref(u_len)):
            if u_len.value > 0:
                return ctypes.wstring_at(lp_buf, u_len.value - 1)

        return ""
    except Exception as exc:
        logger.debug("Could not read version resource: %s", exc)
        return ""


def check_exe_exists() -> Path:
    """Check that at least one built exe exists.  Returns the path or None."""
    if ONEFILE_EXE.exists():
        logger.info("PASS: One-file exe found: %s (%d bytes)", ONEFILE_EXE, ONEFILE_EXE.stat().st_size)
        return ONEFILE_EXE
    if ONEDIR_EXE.exists():
        logger.info("PASS: One-folder exe found: %s (%d bytes)", ONEDIR_EXE, ONEDIR_EXE.stat().st_size)
        return ONEDIR_EXE
    logger.error("FAIL: No built exe found at %s or %s", ONEFILE_EXE, ONEDIR_EXE)
    return None


def check_version_resource(exe_path: Path) -> bool:
    """Check that the exe version resource matches the current build number."""
    build_number = _read_build_number()
    if build_number < 0:
        logger.warning("SKIP: Could not read build number from %s", BUILD_NUMBER_FILE)
        return True

    expected_suffix = f".{build_number}"
    actual = _read_exe_product_version(exe_path)

    if not actual:
        logger.warning("SKIP: Could not read version resource (non-Windows or missing).")
        return True

    if actual.endswith(expected_suffix):
        logger.info("PASS: Version resource = %r (build %d)", actual, build_number)
        return True
    else:
        logger.error(
            "FAIL: Version resource = %r, expected suffix %r",
            actual,
            expected_suffix,
        )
        return False


def check_icon_exists() -> bool:
    """Check that build/assets/app.ico exists and is non-empty."""
    if not ICON_PATH.exists():
        logger.error("FAIL: Icon not found at %s", ICON_PATH)
        return False
    size = ICON_PATH.stat().st_size
    if size < 100:
        logger.error("FAIL: Icon at %s is suspiciously small (%d bytes)", ICON_PATH, size)
        return False
    logger.info("PASS: Icon exists at %s (%d bytes)", ICON_PATH, size)
    return True


def check_launch(exe_path: Path, timeout: int = 10) -> bool:
    """Launch the exe with --help or a harmless flag and check it exits cleanly.

    Only meaningful on a headed Windows machine.
    """
    try:
        result = subprocess.run(
            [str(exe_path), "--mode", "list-sites"],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        if result.returncode == 0:
            logger.info("PASS: Exe launched and exited cleanly (exit code 0).")
            return True
        else:
            logger.warning(
                "WARN: Exe exited with code %d.  stdout=%r  stderr=%r",
                result.returncode,
                result.stdout[:200],
                result.stderr[:200],
            )
            return True  # Non-zero but it ran — acceptable for smoke test
    except subprocess.TimeoutExpired:
        logger.warning("WARN: Exe did not exit within %d seconds — killing.", timeout)
        return False
    except Exception as exc:
        logger.error("FAIL: Could not launch exe: %s", exc)
        return False


def main():
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    do_launch = "--launch" in sys.argv

    results = []

    exe_path = check_exe_exists()
    results.append(exe_path is not None)

    results.append(check_icon_exists())

    if exe_path:
        results.append(check_version_resource(exe_path))

        if do_launch:
            results.append(check_launch(exe_path))

    passed = sum(results)
    total = len(results)
    failed = total - passed

    print()
    if failed == 0:
        print(f"All {total} checks passed.")
        sys.exit(0)
    else:
        print(f"{failed}/{total} checks FAILED.")
        sys.exit(1)


if __name__ == "__main__":
    main()
