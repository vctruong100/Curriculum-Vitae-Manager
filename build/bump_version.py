"""
Build-time version bumper for CV Research Experience Manager.

Reads the current build number from build/build_number.txt, increments it,
writes it back, and prints the new value.  The PyInstaller .spec reads this
file to embed a fresh file_version / product_version in the .exe, which
forces Windows to invalidate its icon cache for the executable.

Usage (called by build scripts, not by users):
    python build/bump_version.py

No network requests.  Purely local file I/O.
"""

import sys
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

BUILD_NUMBER_FILE = Path(__file__).parent / "build_number.txt"


def read_build_number() -> int:
    """Read the current build number from disk.  Returns 0 if missing."""
    try:
        text = BUILD_NUMBER_FILE.read_text(encoding="utf-8").strip()
        return int(text)
    except (FileNotFoundError, ValueError):
        return 0


def write_build_number(number: int) -> None:
    """Write the build number to disk."""
    BUILD_NUMBER_FILE.write_text(str(number) + "\n", encoding="utf-8")
    logger.info("[Version] Wrote build number %d to %s", number, BUILD_NUMBER_FILE)


def next_build_number() -> int:
    """Increment the build number and return the new value."""
    current = read_build_number()
    new_number = current + 1
    write_build_number(new_number)
    return new_number


def main():
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    new_number = next_build_number()
    print(f"Build number bumped to {new_number}")


if __name__ == "__main__":
    main()
