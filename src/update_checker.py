"""
Optional version update checker for CV Research Experience Manager.

All network code is isolated in this module. This module is NEVER imported
at startup unless both config.network_enabled and config.check_updates_on_startup
are True, or the --check-updates CLI flag is used.

No other module in the application performs network requests.
"""

import io
import json
import logging
import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

from config import APP_VERSION, UPDATE_CHECK_URL, get_app_root


logger = logging.getLogger(__name__)

_TIMEOUT = 10
_VERSION_RE = re.compile(r'^v?(\d+)\.(\d+)\.(\d+)')


def parse_semver(version_str: str) -> Optional[Tuple[int, int, int]]:
    m = _VERSION_RE.match(version_str.strip())
    if m is None:
        return None
    return (int(m.group(1)), int(m.group(2)), int(m.group(3)))


def is_newer(remote: str, local: str = APP_VERSION) -> bool:
    r = parse_semver(remote)
    l = parse_semver(local)
    if r is None or l is None:
        return False
    return r > l


def check_for_update() -> Optional[dict]:
    """Check GitHub releases for a newer version.

    Returns a dict with keys 'tag_name', 'html_url', 'zipball_url'
    if a newer version is found, or None otherwise.
    Raises on network errors.
    """
    logger.info("[UpdateChecker] Checking %s", UPDATE_CHECK_URL)
    req = Request(
        UPDATE_CHECK_URL,
        headers={"Accept": "application/vnd.github.v3+json", "User-Agent": "CV-Manager"},
    )
    try:
        resp = urlopen(req, timeout=_TIMEOUT)
        data = json.loads(resp.read().decode("utf-8"))
    except HTTPError as exc:
        logger.warning("[UpdateChecker] HTTP error: %s", exc)
        raise
    except URLError as exc:
        logger.warning("[UpdateChecker] Network error: %s", exc)
        raise
    except Exception as exc:
        logger.warning("[UpdateChecker] Unexpected error: %s", exc)
        raise

    tag = data.get("tag_name", "")
    if not tag:
        logger.info("[UpdateChecker] No tag_name in release response")
        return None

    if is_newer(tag):
        logger.info("[UpdateChecker] Newer version found: %s (current: %s)", tag, APP_VERSION)
        return {
            "tag_name": tag,
            "html_url": data.get("html_url", ""),
            "zipball_url": data.get("zipball_url", ""),
        }

    logger.info("[UpdateChecker] Current version %s is up to date (remote: %s)", APP_VERSION, tag)
    return None


def download_and_apply(zipball_url: str, target_dir: Optional[Path] = None) -> Path:
    """Download a release ZIP, back up the current app, and extract.

    Args:
        zipball_url: URL to the release ZIP archive.
        target_dir: Directory to extract into. Defaults to app root.

    Returns:
        Path to the backup directory.
    """
    if target_dir is None:
        target_dir = get_app_root()

    logger.info("[UpdateChecker] Downloading %s", zipball_url)
    req = Request(zipball_url, headers={"User-Agent": "CV-Manager"})
    resp = urlopen(req, timeout=60)
    zip_bytes = resp.read()

    if len(zip_bytes) == 0:
        raise ValueError("Downloaded ZIP is empty")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = target_dir / f"backup_{timestamp}"
    backup_dir.mkdir(parents=True, exist_ok=True)

    for item in target_dir.iterdir():
        if item.name.startswith("backup_"):
            continue
        if item.name in ("data", ".git", "__pycache__"):
            continue
        dest = backup_dir / item.name
        if item.is_dir():
            shutil.copytree(str(item), str(dest))
        else:
            shutil.copy2(str(item), str(dest))

    logger.info("[UpdateChecker] Backup created at %s", backup_dir)

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        members = zf.namelist()
        prefix = ""
        if members:
            first = members[0]
            if "/" in first:
                prefix = first.split("/")[0] + "/"

        for member in members:
            if member.endswith("/"):
                continue
            rel_path = member
            if prefix and rel_path.startswith(prefix):
                rel_path = rel_path[len(prefix):]
            if not rel_path:
                continue

            out_path = target_dir / rel_path
            out_path.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(member) as src, open(str(out_path), "wb") as dst:
                dst.write(src.read())

    logger.info("[UpdateChecker] Update extracted to %s", target_dir)
    return backup_dir
