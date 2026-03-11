"""
Offline-only enforcement for the CV Research Experience Manager.

Provides:
1. Environment proxy detection (HTTP_PROXY, HTTPS_PROXY, etc.)
2. Import scanner for disallowed network modules (requests, urllib, httpx, etc.)
3. Runtime socket monkeypatch to block all outbound connections.

All guards are controlled by config and default to ON.
"""

import os
import sys
import logging
import socket as _socket_module
from typing import List, Tuple

logger = logging.getLogger(__name__)

# Modules that should never be imported in an offline-only app
DISALLOWED_MODULES = frozenset({
    "requests",
    "urllib3",
    "httpx",
    "aiohttp",
    "grpc",
    "websocket",
    "websockets",
    "socketio",
})

# Environment variables that indicate proxy configuration
PROXY_ENV_VARS = (
    "HTTP_PROXY",
    "HTTPS_PROXY",
    "http_proxy",
    "https_proxy",
    "ALL_PROXY",
    "all_proxy",
    "FTP_PROXY",
    "ftp_proxy",
)

# Store original socket so we can restore if needed
_original_socket_connect = None
_guard_active = False


class OfflineViolationError(RuntimeError):
    """Raised when an offline-only constraint is violated."""
    pass


def check_proxy_environment() -> List[str]:
    """
    Check if any proxy environment variables are set.

    Returns:
        List of (var_name) strings that are set. Empty list means clean.
    """
    violations = []
    for var in PROXY_ENV_VARS:
        value = os.environ.get(var)
        if value:
            violations.append(var)
            logger.warning(
                "Offline guard: proxy env var '%s' is set. "
                "This app is offline-only.",
                var,
            )
    return violations


def scan_loaded_modules() -> List[str]:
    """
    Scan sys.modules for disallowed network modules.

    Returns:
        List of module names that should not be imported.
    """
    found = []
    for mod_name in list(sys.modules.keys()):
        base = mod_name.split(".")[0]
        if base in DISALLOWED_MODULES:
            found.append(mod_name)
            logger.warning(
                "Offline guard: disallowed module '%s' is loaded.",
                mod_name,
            )
    return found


def _blocked_connect(self, *args, **kwargs):
    """Replacement for socket.connect that raises on any call."""
    raise OfflineViolationError(
        "Network connection attempted but this application is offline-only. "
        "All socket.connect() calls are blocked by the offline guard."
    )


def activate_socket_guard():
    """
    Monkeypatch socket.socket.connect to block all outbound connections.
    This is reversible via deactivate_socket_guard().
    """
    global _original_socket_connect, _guard_active
    if _guard_active:
        logger.debug("Offline guard: socket guard already active.")
        return

    _original_socket_connect = _socket_module.socket.connect
    _socket_module.socket.connect = _blocked_connect
    _guard_active = True
    logger.info("Offline guard: socket.connect monkeypatched — all network blocked.")


def deactivate_socket_guard():
    """
    Restore original socket.connect. Useful for testing.
    """
    global _original_socket_connect, _guard_active
    if not _guard_active:
        return
    if _original_socket_connect is not None:
        _socket_module.socket.connect = _original_socket_connect
        _original_socket_connect = None
    _guard_active = False
    logger.info("Offline guard: socket.connect restored.")


def enforce_offline(
    fail_fast: bool = True,
    block_sockets: bool = True,
) -> Tuple[bool, List[str]]:
    """
    Run all offline checks and optionally activate the socket guard.

    Args:
        fail_fast: If True, raise OfflineViolationError on first problem.
        block_sockets: If True, monkeypatch socket.connect.

    Returns:
        (all_clean, list_of_violation_strings)
    """
    violations = []

    # 1. Proxy environment
    proxy_vars = check_proxy_environment()
    for var in proxy_vars:
        violations.append(f"Proxy env var set: {var}")

    # 2. Disallowed modules
    bad_modules = scan_loaded_modules()
    for mod in bad_modules:
        violations.append(f"Disallowed module loaded: {mod}")

    # 3. Socket guard
    if block_sockets:
        activate_socket_guard()

    if violations and fail_fast:
        msg = "Offline-only violations detected:\n" + "\n".join(
            f"  - {v}" for v in violations
        )
        logger.error(msg)
        raise OfflineViolationError(msg)

    if not violations:
        logger.info("Offline guard: all checks passed.")

    return len(violations) == 0, violations
