"""
Security and permissions hardening for the CV Research Experience Manager.

Provides:
- Directory permission enforcement (owner-only on Unix, restricted ACL hints on Windows).
- Log sanitization for Redact mode (mask protocol strings).
- Backup retention / pruning.
"""

import os
import sys
import stat
import logging
import re
from pathlib import Path
from datetime import datetime, timedelta
from typing import List, Optional

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Directory permissions
# ---------------------------------------------------------------------------

def set_owner_only_permissions(path: Path) -> bool:
    """
    Set owner-read/write-only permissions on a directory or file.

    On Unix: chmod 700 (directory) or 600 (file).
    On Windows: logs a warning — use icacls manually (documented in README).

    Returns True if permissions were set successfully.
    """
    if not path.exists():
        logger.warning("Cannot set permissions: path does not exist: %s", path)
        return False

    if sys.platform == "win32":
        # Windows: best-effort using icacls is not safe to automate silently.
        # Log guidance instead.
        logger.info(
            "Windows detected. To restrict '%s' to owner-only, run:\n"
            '  icacls "%s" /inheritance:r /grant:r "%%USERNAME%%:(OI)(CI)F"',
            path,
            path,
        )
        return True  # Not a failure — just informational

    try:
        if path.is_dir():
            os.chmod(path, stat.S_IRWXU)  # 0o700
        else:
            os.chmod(path, stat.S_IRUSR | stat.S_IWUSR)  # 0o600
        logger.info("Set owner-only permissions on %s", path)
        return True
    except OSError as exc:
        logger.error("Failed to set permissions on %s: %s", path, exc)
        return False


def secure_user_directory(user_data_path: Path) -> None:
    """
    Apply owner-only permissions to the user data directory and its
    immediate children (logs, backups, exports, imports, db file).
    """
    if not user_data_path.exists():
        logger.debug("User data path does not exist yet: %s", user_data_path)
        return

    set_owner_only_permissions(user_data_path)

    for child in user_data_path.iterdir():
        set_owner_only_permissions(child)


# ---------------------------------------------------------------------------
# Log sanitization (Redact mode)
# ---------------------------------------------------------------------------

# Regex that matches common protocol-like tokens (e.g. PF-12345, NVS-789)
_PROTOCOL_RE = re.compile(r"\b[A-Za-z]{1,10}-?\d[\w-]*\b")


def sanitize_log_text(text: str, mode: str = "") -> str:
    """
    Sanitize text for logging.

    In Redact mode, replace protocol-like tokens with [REDACTED].
    In other modes, return text as-is.
    """
    if "redact" not in mode.lower():
        return text

    return _PROTOCOL_RE.sub("[REDACTED]", text)


def sanitize_log_entry(entry_dict: dict, mode: str = "") -> dict:
    """
    Sanitize a log entry dict for Redact mode.
    Masks the 'protocol' and 'details' fields.
    """
    if "redact" not in mode.lower():
        return entry_dict

    sanitized = dict(entry_dict)
    if "protocol" in sanitized and sanitized["protocol"]:
        sanitized["protocol"] = "[REDACTED]"
    if "details" in sanitized and sanitized["details"]:
        sanitized["details"] = sanitize_log_text(sanitized["details"], mode)
    return sanitized


# ---------------------------------------------------------------------------
# Backup retention / pruning
# ---------------------------------------------------------------------------

def prune_backups(
    backup_dir: Path,
    retention_days: int = 90,
    dry_run: bool = False,
) -> List[Path]:
    """
    Remove backup files older than *retention_days*.

    Scans for .json and .db files matching the backup naming pattern.
    Returns list of deleted (or would-delete) paths.

    Args:
        backup_dir: Directory containing backup files.
        retention_days: Keep files newer than this many days.
        dry_run: If True, only report — don't delete.
    """
    if not backup_dir.exists():
        return []

    cutoff = datetime.now() - timedelta(days=retention_days)
    deleted: List[Path] = []

    for fp in backup_dir.iterdir():
        if not fp.is_file():
            continue
        if fp.suffix not in (".json", ".db", ".sqlite"):
            continue

        mtime = datetime.fromtimestamp(fp.stat().st_mtime)
        if mtime < cutoff:
            if dry_run:
                logger.info("Would prune backup: %s (modified %s)", fp.name, mtime)
            else:
                try:
                    fp.unlink()
                    logger.info("Pruned backup: %s (modified %s)", fp.name, mtime)
                except OSError as exc:
                    logger.error("Failed to prune %s: %s", fp.name, exc)
            deleted.append(fp)

    if deleted:
        logger.info(
            "Backup pruning: %d file(s) %s from %s",
            len(deleted),
            "would be removed" if dry_run else "removed",
            backup_dir,
        )
    return deleted


def prune_user_backups(
    user_data_path: Path,
    retention_days: int = 90,
    dry_run: bool = False,
) -> List[Path]:
    """Convenience: prune backups under a user's backups/ subdirectory."""
    backup_dir = user_data_path / "backups"
    return prune_backups(backup_dir, retention_days, dry_run)


# ---------------------------------------------------------------------------
# Log retention / pruning
# ---------------------------------------------------------------------------

def prune_logs(
    logs_dir: Path,
    retention_days: int = 90,
    dry_run: bool = False,
) -> List[Path]:
    """
    Remove log files older than *retention_days*.

    Scans for .json, .csv, and .log files in the logs directory.
    Returns list of deleted (or would-delete) paths.

    Args:
        logs_dir: Directory containing log files.
        retention_days: Keep files newer than this many days.
        dry_run: If True, only report — don't delete.
    """
    if not logs_dir.exists():
        return []

    cutoff = datetime.now() - timedelta(days=retention_days)
    deleted: List[Path] = []

    for fp in logs_dir.iterdir():
        if not fp.is_file():
            continue
        if fp.suffix not in (".json", ".csv", ".log"):
            continue

        mtime = datetime.fromtimestamp(fp.stat().st_mtime)
        if mtime < cutoff:
            if dry_run:
                logger.info("Would prune log: %s (modified %s)", fp.name, mtime)
            else:
                try:
                    fp.unlink()
                    logger.info("Pruned log: %s (modified %s)", fp.name, mtime)
                except OSError as exc:
                    logger.error("Failed to prune log %s: %s", fp.name, exc)
            deleted.append(fp)

    if deleted:
        logger.info(
            "Log pruning: %d file(s) %s from %s",
            len(deleted),
            "would be removed" if dry_run else "removed",
            logs_dir,
        )
    return deleted


def prune_user_logs(
    user_data_path: Path,
    retention_days: int = 90,
    dry_run: bool = False,
) -> List[Path]:
    """Convenience: prune logs under a user's logs/ subdirectory."""
    logs_dir = user_data_path / "logs"
    return prune_logs(logs_dir, retention_days, dry_run)
