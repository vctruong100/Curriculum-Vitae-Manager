"""
Tests for the config module.

Covers: defaults, load/save, type validation, network override,
path resolution, user directories.
"""

import sys
import json
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent.resolve()))

from config import AppConfig, get_config, set_config


class TestAppConfigDefaults:
    def test_default_values(self):
        c = AppConfig()
        assert c.fuzzy_threshold_full == 92
        assert c.fuzzy_threshold_masked == 90
        assert c.benchmark_min_count == 4
        assert c.network_enabled is False
        assert c.font_name == "Calibri"
        assert c.font_size == 11
        assert c.phase_order == ["Phase I", "Phase II\u2013IV"]

    def test_data_root_default(self):
        c = AppConfig()
        assert c.data_root != ""
        assert "data" in c.data_root


class TestAppConfigSaveLoad:
    def test_save_and_load(self, tmp_dir):
        path = tmp_dir / "config.json"
        c = AppConfig(fuzzy_threshold_full=85, data_root=str(tmp_dir))
        c.save(path)

        loaded = AppConfig.load(path)
        assert loaded.fuzzy_threshold_full == 85

    def test_load_missing_file(self, tmp_dir):
        loaded = AppConfig.load(tmp_dir / "nope.json")
        # Should return defaults
        assert loaded.fuzzy_threshold_full == 92

    def test_load_corrupt_json(self, tmp_dir):
        path = tmp_dir / "bad.json"
        path.write_text("not valid json {{{")
        loaded = AppConfig.load(path)
        assert loaded.fuzzy_threshold_full == 92  # Defaults

    def test_network_forced_false(self, tmp_dir):
        path = tmp_dir / "config.json"
        data = {"network_enabled": True, "fuzzy_threshold_full": 92}
        path.write_text(json.dumps(data))
        loaded = AppConfig.load(path)
        assert loaded.network_enabled is False


class TestAppConfigFromDict:
    def test_known_fields_only(self):
        data = {
            "fuzzy_threshold_full": 80,
            "unknown_field": "ignored",
        }
        c = AppConfig.from_dict(data)
        assert c.fuzzy_threshold_full == 80
        assert not hasattr(c, "unknown_field")


class TestAppConfigPaths:
    def test_user_data_path(self, tmp_dir):
        c = AppConfig(data_root=str(tmp_dir))
        path = c.get_user_data_path("testuser")
        assert "testuser" in str(path)
        assert str(tmp_dir) in str(path)

    def test_user_db_path(self, tmp_dir):
        c = AppConfig(data_root=str(tmp_dir))
        path = c.get_user_db_path("testuser")
        assert path.name == "sites.db"

    def test_ensure_user_directories(self, tmp_dir):
        c = AppConfig(data_root=str(tmp_dir))
        c.ensure_user_directories("testuser")
        assert (tmp_dir / "users" / "testuser").exists()
        assert (tmp_dir / "users" / "testuser" / "exports").exists()
        assert (tmp_dir / "users" / "testuser" / "imports").exists()
        assert (tmp_dir / "users" / "testuser" / "backups").exists()
        assert (tmp_dir / "users" / "testuser" / "logs").exists()
        assert (tmp_dir / "tmp").exists()


class TestGlobalConfig:
    def test_set_and_get(self, tmp_dir):
        c = AppConfig(data_root=str(tmp_dir), fuzzy_threshold_full=77)
        set_config(c)
        loaded = get_config()
        assert loaded.fuzzy_threshold_full == 77
        # Reset
        set_config(None)
