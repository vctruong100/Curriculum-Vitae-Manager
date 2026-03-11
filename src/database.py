"""
SQLite database layer for multi-site management.
Per-user private, local-only storage.
"""

import sqlite3
import os
from pathlib import Path
from datetime import datetime
from typing import List, Optional, Tuple
import json
import shutil

import logging

from models import Study, Site, SiteVersion
from config import get_config, AppConfig
from migrations import check_and_migrate

db_logger = logging.getLogger(__name__)


class DatabaseManager:
    """Manages SQLite database operations for a user's sites."""
    
    SCHEMA_VERSION = 1
    
    def __init__(self, user_id: Optional[str] = None, config: Optional[AppConfig] = None):
        self.config = config or get_config()
        self.user_id = user_id or self.config.get_user_id()
        self.db_path = self.config.get_user_db_path(self.user_id)
        self._connection: Optional[sqlite3.Connection] = None
        
        # Ensure directories exist
        self.config.ensure_user_directories(self.user_id)
    
    def _get_connection(self) -> sqlite3.Connection:
        """Get or create database connection."""
        if self._connection is None:
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
            self._connection = sqlite3.connect(str(self.db_path))
            self._connection.row_factory = sqlite3.Row
            # Enable WAL mode for safe single-user concurrency
            self._connection.execute('PRAGMA journal_mode=WAL')
            # Enable foreign key enforcement
            self._connection.execute('PRAGMA foreign_keys=ON')
            self._init_schema()
            # Run auto-migrations
            try:
                check_and_migrate(self.db_path, self._connection)
            except Exception as exc:
                db_logger.error("Migration failed: %s", exc)
        return self._connection
    
    def _init_schema(self) -> None:
        """Initialize database schema if not exists."""
        conn = self._connection
        cursor = conn.cursor()
        
        # Sites table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sites (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                owner_user_id TEXT NOT NULL,
                name TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        ''')
        
        # Studies table
        cursor.execute('''
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
        ''')
        
        # Site versions table (for snapshots)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS site_versions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                site_id INTEGER NOT NULL,
                created_at TEXT NOT NULL,
                note TEXT,
                data TEXT NOT NULL,
                FOREIGN KEY (site_id) REFERENCES sites(id) ON DELETE CASCADE
            )
        ''')
        
        # Schema version table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS schema_info (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        ''')
        
        # Category order table (for custom sorting)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS category_order (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                site_id INTEGER NOT NULL,
                order_data TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY (site_id) REFERENCES sites(id) ON DELETE CASCADE,
                UNIQUE(site_id)
            )
        ''')
        
        # Create indices
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sites_owner ON sites(owner_user_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_studies_site ON studies(site_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_studies_year ON studies(year)')
        
        conn.commit()
    
    def close(self) -> None:
        """Close database connection."""
        if self._connection:
            self._connection.close()
            self._connection = None
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
    
    # ==================== Site Operations ====================
    
    def create_site(self, name: str) -> Site:
        """Create a new site for the current user."""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        now = datetime.now().isoformat()
        cursor.execute('''
            INSERT INTO sites (owner_user_id, name, created_at, updated_at)
            VALUES (?, ?, ?, ?)
        ''', (self.user_id, name, now, now))
        
        conn.commit()
        
        return Site(
            id=cursor.lastrowid,
            owner_user_id=self.user_id,
            name=name,
            created_at=datetime.fromisoformat(now),
            updated_at=datetime.fromisoformat(now),
        )
    
    def get_sites(self) -> List[Site]:
        """Get all sites owned by the current user."""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM sites WHERE owner_user_id = ? ORDER BY name
        ''', (self.user_id,))
        
        sites = []
        for row in cursor.fetchall():
            sites.append(Site(
                id=row['id'],
                owner_user_id=row['owner_user_id'],
                name=row['name'],
                created_at=datetime.fromisoformat(row['created_at']),
                updated_at=datetime.fromisoformat(row['updated_at']),
            ))
        
        return sites
    
    def get_site(self, site_id: int) -> Optional[Site]:
        """Get a site by ID, verifying ownership."""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM sites WHERE id = ? AND owner_user_id = ?
        ''', (site_id, self.user_id))
        
        row = cursor.fetchone()
        if row is None:
            return None
        
        return Site(
            id=row['id'],
            owner_user_id=row['owner_user_id'],
            name=row['name'],
            created_at=datetime.fromisoformat(row['created_at']),
            updated_at=datetime.fromisoformat(row['updated_at']),
        )
    
    def rename_site(self, site_id: int, new_name: str) -> bool:
        """Rename a site, verifying ownership."""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        now = datetime.now().isoformat()
        cursor.execute('''
            UPDATE sites SET name = ?, updated_at = ?
            WHERE id = ? AND owner_user_id = ?
        ''', (new_name, now, site_id, self.user_id))
        
        conn.commit()
        return cursor.rowcount > 0
    
    def delete_site(self, site_id: int) -> bool:
        """Delete a site and all its studies, verifying ownership."""
        # Create backup first
        self.create_site_backup(site_id, "Pre-deletion backup")
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        # Delete studies first (cascade)
        cursor.execute('''
            DELETE FROM studies WHERE site_id = ? AND EXISTS (
                SELECT 1 FROM sites WHERE id = ? AND owner_user_id = ?
            )
        ''', (site_id, site_id, self.user_id))
        
        # Delete site
        cursor.execute('''
            DELETE FROM sites WHERE id = ? AND owner_user_id = ?
        ''', (site_id, self.user_id))
        
        conn.commit()
        return cursor.rowcount > 0
    
    def save_category_order(self, site_id: int, order: List[str]) -> bool:
        """Save custom category order for a site."""
        if not self._verify_site_ownership(site_id):
            return False
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        now = datetime.now().isoformat()
        order_json = json.dumps(order)
        
        cursor.execute('''
            INSERT OR REPLACE INTO category_order (site_id, order_data, updated_at)
            VALUES (?, ?, ?)
        ''', (site_id, order_json, now))
        
        conn.commit()
        return True
    
    def get_category_order(self, site_id: int) -> Optional[List[str]]:
        """Get custom category order for a site."""
        if not self._verify_site_ownership(site_id):
            return None
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT order_data FROM category_order WHERE site_id = ?
        ''', (site_id,))
        
        row = cursor.fetchone()
        if row is None:
            return None
        
        return json.loads(row['order_data'])
    
    # ==================== Study Operations ====================
    
    def add_study(self, site_id: int, study: Study) -> Optional[Study]:
        """Add a study to a site, verifying ownership."""
        # Verify site ownership
        if not self._verify_site_ownership(site_id):
            return None
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        now = datetime.now().isoformat()
        cursor.execute('''
            INSERT INTO studies (
                site_id, phase, subcategory, year, sponsor, protocol,
                description_full, description_masked, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            site_id, study.phase, study.subcategory, study.year,
            study.sponsor, study.protocol, study.description_full,
            study.description_masked, now, now
        ))
        
        # Update site timestamp
        cursor.execute('''
            UPDATE sites SET updated_at = ? WHERE id = ?
        ''', (now, site_id))
        
        conn.commit()
        
        study.id = cursor.lastrowid
        study.site_id = site_id
        study.created_at = datetime.fromisoformat(now)
        study.updated_at = datetime.fromisoformat(now)
        
        return study
    
    def get_studies(self, site_id: int) -> List[Study]:
        """Get all studies for a site, verifying ownership."""
        if not self._verify_site_ownership(site_id):
            return []
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM studies WHERE site_id = ?
            ORDER BY phase, subcategory, year DESC, sponsor, protocol
        ''', (site_id,))
        
        studies = []
        for row in cursor.fetchall():
            studies.append(Study(
                id=row['id'],
                site_id=row['site_id'],
                phase=row['phase'],
                subcategory=row['subcategory'],
                year=row['year'],
                sponsor=row['sponsor'],
                protocol=row['protocol'] or '',
                description_full=row['description_full'],
                description_masked=row['description_masked'],
                created_at=datetime.fromisoformat(row['created_at']),
                updated_at=datetime.fromisoformat(row['updated_at']),
            ))
        
        return studies
    
    def update_study(self, study: Study) -> bool:
        """Update a study, verifying ownership."""
        if study.id is None or study.site_id is None:
            return False
        
        if not self._verify_site_ownership(study.site_id):
            return False
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        now = datetime.now().isoformat()
        cursor.execute('''
            UPDATE studies SET
                phase = ?, subcategory = ?, year = ?, sponsor = ?,
                protocol = ?, description_full = ?, description_masked = ?,
                updated_at = ?
            WHERE id = ? AND site_id = ?
        ''', (
            study.phase, study.subcategory, study.year, study.sponsor,
            study.protocol, study.description_full, study.description_masked,
            now, study.id, study.site_id
        ))
        
        # Update site timestamp
        cursor.execute('''
            UPDATE sites SET updated_at = ? WHERE id = ?
        ''', (now, study.site_id))
        
        conn.commit()
        return cursor.rowcount > 0
    
    def delete_study(self, study_id: int, site_id: int) -> bool:
        """Delete a study, verifying ownership."""
        if not self._verify_site_ownership(site_id):
            return False
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            DELETE FROM studies WHERE id = ? AND site_id = ?
        ''', (study_id, site_id))
        
        now = datetime.now().isoformat()
        cursor.execute('''
            UPDATE sites SET updated_at = ? WHERE id = ?
        ''', (now, site_id))
        
        conn.commit()
        return cursor.rowcount > 0
    
    def bulk_add_studies(self, site_id: int, studies: List[Study]) -> int:
        """Add multiple studies to a site. Returns count added."""
        if not self._verify_site_ownership(site_id):
            return 0
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        now = datetime.now().isoformat()
        count = 0
        
        for study in studies:
            cursor.execute('''
                INSERT INTO studies (
                    site_id, phase, subcategory, year, sponsor, protocol,
                    description_full, description_masked, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                site_id, study.phase, study.subcategory, study.year,
                study.sponsor, study.protocol, study.description_full,
                study.description_masked, now, now
            ))
            count += 1
        
        # Update site timestamp
        cursor.execute('''
            UPDATE sites SET updated_at = ? WHERE id = ?
        ''', (now, site_id))
        
        conn.commit()
        return count
    
    # ==================== Backup & Version Operations ====================
    
    def create_site_backup(self, site_id: int, note: str = "") -> Optional[int]:
        """Create a versioned backup of a site."""
        if not self._verify_site_ownership(site_id):
            return None
        
        site = self.get_site(site_id)
        if not site:
            return None
        
        studies = self.get_studies(site_id)
        
        # Create snapshot data
        snapshot = {
            'site_name': site.name,
            'studies': [
                {
                    'phase': s.phase,
                    'subcategory': s.subcategory,
                    'year': s.year,
                    'sponsor': s.sponsor,
                    'protocol': s.protocol,
                    'description_full': s.description_full,
                    'description_masked': s.description_masked,
                }
                for s in studies
            ]
        }
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        now = datetime.now().isoformat()
        cursor.execute('''
            INSERT INTO site_versions (site_id, created_at, note, data)
            VALUES (?, ?, ?, ?)
        ''', (site_id, now, note, json.dumps(snapshot)))
        
        conn.commit()
        
        # Also create a file backup
        self._create_file_backup(site_id, site.name, snapshot)
        
        return cursor.lastrowid
    
    def _create_file_backup(self, site_id: int, site_name: str, data: dict) -> Path:
        """Create a timestamped file backup."""
        backup_dir = self.config.get_user_backups_path(self.user_id)
        backup_dir.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = "".join(c if c.isalnum() or c in '-_' else '_' for c in site_name)
        backup_file = backup_dir / f"site_{site_id}_{safe_name}_{timestamp}.json"
        
        with open(backup_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
        
        return backup_file
    
    def get_site_versions(self, site_id: int) -> List[SiteVersion]:
        """Get all versions/snapshots of a site."""
        if not self._verify_site_ownership(site_id):
            return []
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, site_id, created_at, note FROM site_versions
            WHERE site_id = ? ORDER BY created_at DESC
        ''', (site_id,))
        
        versions = []
        for row in cursor.fetchall():
            versions.append(SiteVersion(
                id=row['id'],
                site_id=row['site_id'],
                created_at=datetime.fromisoformat(row['created_at']),
                note=row['note'] or '',
            ))
        
        return versions
    
    # ==================== Helper Methods ====================
    
    def _verify_site_ownership(self, site_id: int) -> bool:
        """Verify that the current user owns the site."""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT 1 FROM sites WHERE id = ? AND owner_user_id = ?
        ''', (site_id, self.user_id))
        
        return cursor.fetchone() is not None
    
    def get_study_count(self, site_id: int) -> int:
        """Get count of studies in a site."""
        if not self._verify_site_ownership(site_id):
            return 0
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT COUNT(*) FROM studies WHERE site_id = ?', (site_id,))
        return cursor.fetchone()[0]
    
    def clear_studies(self, site_id: int) -> bool:
        """Clear all studies from a site (with backup)."""
        if not self._verify_site_ownership(site_id):
            return False
        
        # Backup first
        self.create_site_backup(site_id, "Pre-clear backup")
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute('DELETE FROM studies WHERE site_id = ?', (site_id,))
        
        now = datetime.now().isoformat()
        cursor.execute('UPDATE sites SET updated_at = ? WHERE id = ?', (now, site_id))
        
        conn.commit()
        return True


def verify_database_access(user_id: str, target_user_id: str) -> Tuple[bool, str]:
    """
    Verify that a user can access another user's data.
    Returns (allowed, reason).
    """
    if user_id != target_user_id:
        return False, f"Access denied: User '{user_id}' cannot access data owned by '{target_user_id}'"
    return True, ""
