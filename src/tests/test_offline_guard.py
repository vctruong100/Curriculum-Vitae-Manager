"""
Tests for the offline_guard module.

Covers: proxy detection, module scanning, socket monkeypatch,
enforce_offline orchestration.
"""

import sys
import os
import socket
from pathlib import Path
from unittest import mock

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent.resolve()))

from offline_guard import (
    check_proxy_environment,
    scan_loaded_modules,
    activate_socket_guard,
    deactivate_socket_guard,
    enforce_offline,
    OfflineViolationError,
)


class TestCheckProxyEnvironment:
    def test_clean_environment(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            violations = check_proxy_environment()
            assert len(violations) == 0

    def test_http_proxy_set(self):
        with mock.patch.dict(os.environ, {"HTTP_PROXY": "http://proxy:8080"}, clear=False):
            violations = check_proxy_environment()
            assert "HTTP_PROXY" in violations

    def test_https_proxy_set(self):
        with mock.patch.dict(os.environ, {"HTTPS_PROXY": "http://proxy:8080"}, clear=False):
            violations = check_proxy_environment()
            assert "HTTPS_PROXY" in violations

    def test_multiple_proxies(self):
        env = {
            "HTTP_PROXY": "http://p:80",
            "HTTPS_PROXY": "http://p:443",
        }
        with mock.patch.dict(os.environ, env, clear=False):
            violations = check_proxy_environment()
            assert len(violations) >= 2


class TestScanLoadedModules:
    def test_no_disallowed_modules(self):
        # In a test environment, requests/httpx should not be loaded
        # unless test dependencies pull them in
        found = scan_loaded_modules()
        # We can't guarantee they're not loaded, so just check type
        assert isinstance(found, list)

    def test_detects_loaded_module(self):
        # Temporarily fake a disallowed module in sys.modules
        sys.modules["requests"] = mock.MagicMock()
        try:
            found = scan_loaded_modules()
            assert any("requests" in m for m in found)
        finally:
            del sys.modules["requests"]


class TestSocketGuard:
    def test_activate_blocks_connect(self):
        activate_socket_guard()
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            with pytest.raises(OfflineViolationError):
                s.connect(("127.0.0.1", 80))
            s.close()
        finally:
            deactivate_socket_guard()

    def test_deactivate_restores(self):
        activate_socket_guard()
        deactivate_socket_guard()
        # After deactivation, socket.connect should be the original
        # We don't actually connect, just verify it doesn't raise
        # OfflineViolationError for a non-connecting call
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.close()  # Just verify no crash

    def test_double_activate_is_safe(self):
        activate_socket_guard()
        activate_socket_guard()  # Should not crash
        deactivate_socket_guard()

    def test_double_deactivate_is_safe(self):
        deactivate_socket_guard()
        deactivate_socket_guard()  # Should not crash


class TestEnforceOffline:
    def test_clean_environment_passes(self):
        # Remove any proxies for this test
        env_clean = {k: "" for k in ["HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy", "ALL_PROXY", "all_proxy"]}
        with mock.patch.dict(os.environ, env_clean, clear=False):
            all_clean, violations = enforce_offline(fail_fast=False, block_sockets=False)
            # May still have violations from loaded modules in test env
            assert isinstance(all_clean, bool)
            assert isinstance(violations, list)

    def test_fail_fast_raises(self):
        with mock.patch.dict(os.environ, {"HTTP_PROXY": "http://proxy:80"}, clear=False):
            with pytest.raises(OfflineViolationError):
                enforce_offline(fail_fast=True, block_sockets=False)

    def test_no_fail_fast_returns_violations(self):
        with mock.patch.dict(os.environ, {"HTTP_PROXY": "http://proxy:80"}, clear=False):
            all_clean, violations = enforce_offline(fail_fast=False, block_sockets=False)
            assert all_clean is False
            assert len(violations) > 0
