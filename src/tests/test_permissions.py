"""
Tests for the permissions module.

Covers: log sanitization, backup pruning, directory permissions.
"""

import sys
import time
from pathlib import Path
from datetime import datetime, timedelta

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent.resolve()))

from permissions import (
    sanitize_log_text,
    sanitize_log_entry,
    prune_backups,
    prune_logs,
    set_owner_only_permissions,
)


class TestSanitizeLogText:
    def test_redact_mode_masks_protocol(self):
        result = sanitize_log_text("Matched PF-12345 in study", mode="redact")
        assert "PF-12345" not in result
        assert "[REDACTED]" in result

    def test_non_redact_mode_preserves(self):
        text = "Matched PF-12345 in study"
        result = sanitize_log_text(text, mode="update")
        assert result == text

    def test_empty_string(self):
        assert sanitize_log_text("", mode="redact") == ""

    def test_no_protocol_in_text(self):
        result = sanitize_log_text("A plain text message", mode="redact")
        assert "[REDACTED]" not in result

    def test_multiple_protocols(self):
        result = sanitize_log_text("PF-123 and NVS-456 matched", mode="redact")
        assert "PF-123" not in result
        assert "NVS-456" not in result


class TestSanitizeLogEntry:
    def test_redact_mode(self):
        entry = {
            "operation": "replaced",
            "protocol": "PF-12345",
            "details": "Matched PF-12345 (score=95)",
            "sponsor": "Pfizer",
        }
        sanitized = sanitize_log_entry(entry, mode="redact")
        assert sanitized["protocol"] == "[REDACTED]"
        assert "PF-12345" not in sanitized["details"]
        assert sanitized["sponsor"] == "Pfizer"  # Sponsor preserved

    def test_non_redact_mode(self):
        entry = {
            "protocol": "PF-12345",
            "details": "Some details",
        }
        sanitized = sanitize_log_entry(entry, mode="update")
        assert sanitized["protocol"] == "PF-12345"


class TestPruneBackups:
    def test_prunes_old_files(self, tmp_dir):
        backup_dir = tmp_dir / "backups"
        backup_dir.mkdir()

        # Create an "old" backup file
        old_file = backup_dir / "site_1_old_20200101_000000.json"
        old_file.write_text("{}")
        # Set mtime to 200 days ago
        import os
        old_time = time.time() - (200 * 86400)
        os.utime(old_file, (old_time, old_time))

        # Create a "new" backup file
        new_file = backup_dir / "site_1_new_20240101_000000.json"
        new_file.write_text("{}")

        deleted = prune_backups(backup_dir, retention_days=90)
        assert old_file in deleted
        assert new_file not in deleted
        assert not old_file.exists()
        assert new_file.exists()

    def test_dry_run(self, tmp_dir):
        backup_dir = tmp_dir / "backups"
        backup_dir.mkdir()
        old_file = backup_dir / "site_1_old.json"
        old_file.write_text("{}")
        import os
        old_time = time.time() - (200 * 86400)
        os.utime(old_file, (old_time, old_time))

        deleted = prune_backups(backup_dir, retention_days=90, dry_run=True)
        assert len(deleted) == 1
        assert old_file.exists()  # Not actually deleted

    def test_empty_directory(self, tmp_dir):
        backup_dir = tmp_dir / "backups"
        backup_dir.mkdir()
        deleted = prune_backups(backup_dir, retention_days=90)
        assert len(deleted) == 0

    def test_nonexistent_directory(self, tmp_dir):
        deleted = prune_backups(tmp_dir / "nope", retention_days=90)
        assert len(deleted) == 0


class TestPruneLogs:
    def test_prunes_old_json_and_csv(self, tmp_dir):
        logs_dir = tmp_dir / "logs"
        logs_dir.mkdir()

        old_json = logs_dir / "update_20200101_000000.json"
        old_json.write_text("{}")
        old_csv = logs_dir / "update_20200101_000000.csv"
        old_csv.write_text("header")
        import os
        old_time = time.time() - (200 * 86400)
        os.utime(old_json, (old_time, old_time))
        os.utime(old_csv, (old_time, old_time))

        new_json = logs_dir / "update_20240601_120000.json"
        new_json.write_text("{}")

        deleted = prune_logs(logs_dir, retention_days=90)
        assert old_json in deleted
        assert old_csv in deleted
        assert new_json not in deleted
        assert not old_json.exists()
        assert not old_csv.exists()
        assert new_json.exists()

    def test_prunes_old_log_file(self, tmp_dir):
        logs_dir = tmp_dir / "logs"
        logs_dir.mkdir()
        old_log = logs_dir / "access_denied.log"
        old_log.write_text("entry")
        import os
        old_time = time.time() - (200 * 86400)
        os.utime(old_log, (old_time, old_time))

        deleted = prune_logs(logs_dir, retention_days=90)
        assert old_log in deleted
        assert not old_log.exists()

    def test_dry_run(self, tmp_dir):
        logs_dir = tmp_dir / "logs"
        logs_dir.mkdir()
        old_file = logs_dir / "old.json"
        old_file.write_text("{}")
        import os
        old_time = time.time() - (200 * 86400)
        os.utime(old_file, (old_time, old_time))

        deleted = prune_logs(logs_dir, retention_days=90, dry_run=True)
        assert len(deleted) == 1
        assert old_file.exists()  # Not actually deleted

    def test_ignores_non_log_extensions(self, tmp_dir):
        logs_dir = tmp_dir / "logs"
        logs_dir.mkdir()
        txt_file = logs_dir / "notes.txt"
        txt_file.write_text("keep me")
        import os
        old_time = time.time() - (200 * 86400)
        os.utime(txt_file, (old_time, old_time))

        deleted = prune_logs(logs_dir, retention_days=90)
        assert len(deleted) == 0
        assert txt_file.exists()

    def test_empty_directory(self, tmp_dir):
        logs_dir = tmp_dir / "logs"
        logs_dir.mkdir()
        deleted = prune_logs(logs_dir, retention_days=90)
        assert len(deleted) == 0

    def test_nonexistent_directory(self, tmp_dir):
        deleted = prune_logs(tmp_dir / "nope", retention_days=90)
        assert len(deleted) == 0


class TestSetOwnerOnlyPermissions:
    def test_nonexistent_path(self, tmp_dir):
        result = set_owner_only_permissions(tmp_dir / "nope")
        assert result is False

    def test_directory(self, tmp_dir):
        test_dir = tmp_dir / "secure"
        test_dir.mkdir()
        result = set_owner_only_permissions(test_dir)
        assert result is True

    def test_file(self, tmp_dir):
        test_file = tmp_dir / "secure.txt"
        test_file.write_text("secret")
        result = set_owner_only_permissions(test_file)
        assert result is True
