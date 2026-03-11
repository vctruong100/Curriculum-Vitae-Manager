"""
Tests for the migrations module.

Covers: schema version tracking, auto-migration, rollback, backup creation.
"""

import sys
import sqlite3
import tempfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent.resolve()))

from migrations import (
    get_schema_version,
    set_schema_version,
    ensure_schema_info_table,
    auto_migrate,
    rollback_one,
    check_and_migrate,
    backup_database,
    LATEST_VERSION,
    MIGRATIONS,
)


@pytest.fixture
def db_conn(tmp_path):
    """Provide an in-memory SQLite connection with schema_info table."""
    db_path = tmp_path / "test.db"
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    # Create the base tables that migration v1 expects
    conn.execute("""
        CREATE TABLE IF NOT EXISTS sites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            owner_user_id TEXT NOT NULL,
            name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS studies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id INTEGER NOT NULL,
            phase TEXT NOT NULL,
            subcategory TEXT NOT NULL,
            year INTEGER NOT NULL,
            sponsor TEXT NOT NULL,
            protocol TEXT,
            description_full TEXT NOT NULL,
            description_masked TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (site_id) REFERENCES sites(id) ON DELETE CASCADE
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS site_versions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            note TEXT,
            data TEXT NOT NULL,
            FOREIGN KEY (site_id) REFERENCES sites(id) ON DELETE CASCADE
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS category_order (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id INTEGER NOT NULL,
            order_data TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (site_id) REFERENCES sites(id) ON DELETE CASCADE,
            UNIQUE(site_id)
        )
    """)
    conn.commit()
    yield conn, db_path
    conn.close()


class TestSchemaVersion:
    def test_no_table_returns_zero(self, tmp_path):
        conn = sqlite3.connect(":memory:")
        assert get_schema_version(conn) == 0
        conn.close()

    def test_set_and_get(self, db_conn):
        conn, _ = db_conn
        ensure_schema_info_table(conn)
        set_schema_version(conn, 5)
        assert get_schema_version(conn) == 5

    def test_overwrite_version(self, db_conn):
        conn, _ = db_conn
        ensure_schema_info_table(conn)
        set_schema_version(conn, 1)
        set_schema_version(conn, 2)
        assert get_schema_version(conn) == 2


class TestAutoMigrate:
    def test_migrate_from_zero(self, db_conn):
        conn, db_path = db_conn
        applied = auto_migrate(conn, db_path)
        assert len(applied) == len(MIGRATIONS)
        assert get_schema_version(conn) == LATEST_VERSION

    def test_already_current(self, db_conn):
        conn, db_path = db_conn
        auto_migrate(conn, db_path)
        # Run again — should be no-op
        applied = auto_migrate(conn, db_path)
        assert len(applied) == 0

    def test_migrate_to_specific_version(self, db_conn):
        conn, db_path = db_conn
        applied = auto_migrate(conn, db_path, target_version=1)
        assert get_schema_version(conn) == 1
        assert len(applied) == 1

    def test_dry_run(self, db_conn):
        conn, db_path = db_conn
        applied = auto_migrate(conn, db_path, dry_run=True)
        assert len(applied) == len(MIGRATIONS)
        # Version should NOT have changed
        assert get_schema_version(conn) == 0

    def test_creates_backup(self, db_conn):
        conn, db_path = db_conn
        # Write some data so db file exists
        conn.commit()
        auto_migrate(conn, db_path)
        # Check backup files were created
        backups = list(db_path.parent.glob("test_pre_v*"))
        assert len(backups) >= 1


class TestRollback:
    def test_rollback_one_step(self, db_conn):
        conn, db_path = db_conn
        auto_migrate(conn, db_path)
        assert get_schema_version(conn) == LATEST_VERSION

        desc = rollback_one(conn, db_path)
        assert desc is not None
        assert get_schema_version(conn) == LATEST_VERSION - 1

    def test_rollback_at_zero(self, db_conn):
        conn, db_path = db_conn
        ensure_schema_info_table(conn)
        desc = rollback_one(conn, db_path)
        assert desc is None


class TestCheckAndMigrate:
    def test_auto_migrates(self, db_conn):
        conn, db_path = db_conn
        check_and_migrate(db_path, conn)
        assert get_schema_version(conn) == LATEST_VERSION


class TestBackupDatabase:
    def test_creates_file(self, db_conn):
        _, db_path = db_conn
        backup_path = backup_database(db_path, "test_backup")
        assert backup_path.exists()
        assert backup_path.stat().st_size > 0
