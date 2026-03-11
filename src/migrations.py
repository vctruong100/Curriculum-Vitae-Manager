"""
SQLite schema migration system for the CV Research Experience Manager.

Provides:
- A schema_version table for tracking the current version.
- An auto-migration function that applies pending migrations in order.
- Rollback/backup before each migration step.
- Version checks at startup.

Each migration is a simple (version, description, up_sql, down_sql) tuple.
"""

import sqlite3
import shutil
import logging
from pathlib import Path
from datetime import datetime
from typing import List, Optional, Tuple

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Migration definitions
# ---------------------------------------------------------------------------
# Each entry: (version: int, description: str, up_sql: str, down_sql: str)
# Version numbers MUST be sequential starting at 1.

MIGRATIONS: List[Tuple[int, str, str, str]] = [
    (
        1,
        "Initial schema — sites, studies, site_versions, schema_info, category_order",
        # up — these tables already exist from _init_schema; this migration
        # is a no-op marker so the version table records we are at v1.
        "",
        # down — drop everything (destructive, emergency only)
        """
        DROP TABLE IF EXISTS category_order;
        DROP TABLE IF EXISTS site_versions;
        DROP TABLE IF EXISTS studies;
        DROP TABLE IF EXISTS sites;
        DROP TABLE IF EXISTS schema_info;
        """,
    ),
    (
        2,
        "Add backup_retention_days to schema_info defaults",
        """
        INSERT OR IGNORE INTO schema_info (key, value)
        VALUES ('backup_retention_days', '90');
        """,
        """
        DELETE FROM schema_info WHERE key = 'backup_retention_days';
        """,
    ),
    (
        3,
        "Add description index for faster fuzzy lookups",
        """
        CREATE INDEX IF NOT EXISTS idx_studies_sponsor
            ON studies(sponsor);
        CREATE INDEX IF NOT EXISTS idx_studies_phase_subcat
            ON studies(phase, subcategory);
        """,
        """
        DROP INDEX IF EXISTS idx_studies_sponsor;
        DROP INDEX IF EXISTS idx_studies_phase_subcat;
        """,
    ),
]

LATEST_VERSION = MIGRATIONS[-1][0] if MIGRATIONS else 0


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_schema_version(conn: sqlite3.Connection) -> int:
    """
    Read current schema version from the database.
    Returns 0 if the schema_info table doesn't exist or has no version row.
    """
    try:
        cursor = conn.execute(
            "SELECT value FROM schema_info WHERE key = 'schema_version'"
        )
        row = cursor.fetchone()
        if row is not None:
            return int(row[0])
    except sqlite3.OperationalError:
        # Table doesn't exist yet
        pass
    return 0


def set_schema_version(conn: sqlite3.Connection, version: int) -> None:
    """Write the current schema version."""
    conn.execute(
        """
        INSERT OR REPLACE INTO schema_info (key, value)
        VALUES ('schema_version', ?)
        """,
        (str(version),),
    )
    conn.commit()
    logger.info("Schema version set to %d", version)


def ensure_schema_info_table(conn: sqlite3.Connection) -> None:
    """Create the schema_info table if it does not exist."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS schema_info (
            key TEXT PRIMARY KEY,
            value TEXT
        )
        """
    )
    conn.commit()


def backup_database(db_path: Path, label: str = "pre_migration") -> Path:
    """
    Create a timestamped copy of the database file.

    Returns the backup path.
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = db_path.parent / f"{db_path.stem}_{label}_{ts}{db_path.suffix}"
    shutil.copy2(db_path, backup_path)
    logger.info("Database backed up to %s", backup_path)
    return backup_path


def auto_migrate(
    conn: sqlite3.Connection,
    db_path: Optional[Path] = None,
    target_version: Optional[int] = None,
    dry_run: bool = False,
) -> List[str]:
    """
    Apply all pending migrations up to *target_version* (default: latest).

    Args:
        conn: Open SQLite connection.
        db_path: Path to the .db file (for backups). If None, skip backup.
        target_version: Stop after reaching this version. None = latest.
        dry_run: If True, log what would happen but don't execute.

    Returns:
        List of applied migration descriptions.
    """
    ensure_schema_info_table(conn)
    current = get_schema_version(conn)
    target = target_version if target_version is not None else LATEST_VERSION

    if current >= target:
        logger.info(
            "Schema is up-to-date (version %d, target %d).", current, target
        )
        return []

    applied = []
    for version, description, up_sql, _down_sql in MIGRATIONS:
        if version <= current:
            continue
        if version > target:
            break

        logger.info(
            "Applying migration v%d: %s%s",
            version,
            description,
            " [DRY RUN]" if dry_run else "",
        )

        if not dry_run:
            # Backup before each step
            if db_path is not None and db_path.exists():
                backup_database(db_path, f"pre_v{version}")

            if up_sql.strip():
                for statement in up_sql.strip().split(";"):
                    stmt = statement.strip()
                    if stmt:
                        conn.execute(stmt)

            set_schema_version(conn, version)

        applied.append(f"v{version}: {description}")

    if applied:
        logger.info("Applied %d migration(s).", len(applied))
    return applied


def rollback_one(
    conn: sqlite3.Connection,
    db_path: Optional[Path] = None,
) -> Optional[str]:
    """
    Roll back the most recent migration (one step down).

    Returns the description of the rolled-back migration, or None if at v0.
    """
    ensure_schema_info_table(conn)
    current = get_schema_version(conn)

    if current <= 0:
        logger.info("Already at version 0, nothing to roll back.")
        return None

    migration = None
    for v, desc, _up, down_sql in MIGRATIONS:
        if v == current:
            migration = (v, desc, down_sql)
            break

    if migration is None:
        logger.error("No migration found for version %d", current)
        return None

    v, desc, down_sql = migration
    logger.info("Rolling back migration v%d: %s", v, desc)

    if db_path is not None and db_path.exists():
        backup_database(db_path, f"pre_rollback_v{v}")

    if down_sql.strip():
        for statement in down_sql.strip().split(";"):
            stmt = statement.strip()
            if stmt:
                conn.execute(stmt)

    set_schema_version(conn, v - 1)
    return f"v{v}: {desc}"


def check_and_migrate(
    db_path: Path,
    conn: sqlite3.Connection,
) -> None:
    """
    Convenience function for startup: check version and apply pending
    migrations automatically.
    """
    ensure_schema_info_table(conn)
    current = get_schema_version(conn)
    if current < LATEST_VERSION:
        logger.info(
            "Database at v%d, latest is v%d — running auto-migrate.",
            current,
            LATEST_VERSION,
        )
        applied = auto_migrate(conn, db_path)
        for desc in applied:
            logger.info("  Applied: %s", desc)
    else:
        logger.debug("Database schema is current (v%d).", current)
