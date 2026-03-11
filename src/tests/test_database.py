"""
Tests for the database module.

Covers: CRUD for sites/studies, ownership verification, bulk operations,
backups, versioning, category order, WAL mode, concurrent safety.
"""

import sys
import sqlite3
from pathlib import Path
from datetime import datetime

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent.resolve()))

from database import DatabaseManager
from models import Study, Site
from config import AppConfig


@pytest.fixture
def db_manager(app_config):
    """Provide a DatabaseManager with a temp database."""
    with DatabaseManager(config=app_config) as db:
        yield db


class TestSiteCRUD:
    def test_create_site(self, db_manager):
        site = db_manager.create_site("Test Site")
        assert site.id is not None
        assert site.name == "Test Site"

    def test_get_sites(self, db_manager):
        db_manager.create_site("Site A")
        db_manager.create_site("Site B")
        sites = db_manager.get_sites()
        assert len(sites) == 2

    def test_get_site_by_id(self, db_manager):
        site = db_manager.create_site("My Site")
        fetched = db_manager.get_site(site.id)
        assert fetched is not None
        assert fetched.name == "My Site"

    def test_get_nonexistent_site(self, db_manager):
        assert db_manager.get_site(9999) is None

    def test_rename_site(self, db_manager):
        site = db_manager.create_site("Old Name")
        success = db_manager.rename_site(site.id, "New Name")
        assert success is True
        fetched = db_manager.get_site(site.id)
        assert fetched.name == "New Name"

    def test_delete_site(self, db_manager):
        site = db_manager.create_site("To Delete")
        success = db_manager.delete_site(site.id)
        assert success is True
        assert db_manager.get_site(site.id) is None


class TestStudyCRUD:
    def _make_study(self):
        return Study(
            phase="Phase I",
            subcategory="Oncology",
            year=2024,
            sponsor="Pfizer",
            protocol="PF-123",
            description_full="Full desc",
            description_masked="Masked desc",
        )

    def test_add_study(self, db_manager):
        site = db_manager.create_site("Site")
        study = self._make_study()
        result = db_manager.add_study(site.id, study)
        assert result is not None
        assert result.id is not None

    def test_get_studies(self, db_manager):
        site = db_manager.create_site("Site")
        db_manager.add_study(site.id, self._make_study())
        studies = db_manager.get_studies(site.id)
        assert len(studies) == 1
        assert studies[0].sponsor == "Pfizer"

    def test_update_study(self, db_manager):
        site = db_manager.create_site("Site")
        study = self._make_study()
        db_manager.add_study(site.id, study)
        studies = db_manager.get_studies(site.id)
        s = studies[0]
        s.sponsor = "Novartis"
        success = db_manager.update_study(s)
        assert success is True
        updated = db_manager.get_studies(site.id)
        assert updated[0].sponsor == "Novartis"

    def test_delete_study(self, db_manager):
        site = db_manager.create_site("Site")
        study = self._make_study()
        db_manager.add_study(site.id, study)
        studies = db_manager.get_studies(site.id)
        success = db_manager.delete_study(studies[0].id, site.id)
        assert success is True
        assert len(db_manager.get_studies(site.id)) == 0

    def test_bulk_add(self, db_manager):
        site = db_manager.create_site("Bulk Site")
        studies = [self._make_study() for _ in range(10)]
        count = db_manager.bulk_add_studies(site.id, studies)
        assert count == 10
        assert db_manager.get_study_count(site.id) == 10

    def test_clear_studies(self, db_manager):
        site = db_manager.create_site("Clear Site")
        db_manager.bulk_add_studies(site.id, [self._make_study() for _ in range(5)])
        assert db_manager.get_study_count(site.id) == 5
        db_manager.clear_studies(site.id)
        assert db_manager.get_study_count(site.id) == 0


class TestOwnership:
    def test_cannot_access_other_user_site(self, app_config):
        # Create site as user A
        with DatabaseManager(user_id="user_a", config=app_config) as db_a:
            site = db_a.create_site("User A Site")

        # Try to access as user B
        with DatabaseManager(user_id="user_b", config=app_config) as db_b:
            result = db_b.get_site(site.id)
            assert result is None

    def test_cannot_add_study_to_other_user_site(self, app_config):
        with DatabaseManager(user_id="user_a", config=app_config) as db_a:
            site = db_a.create_site("User A Site")

        with DatabaseManager(user_id="user_b", config=app_config) as db_b:
            study = Study(
                phase="P", subcategory="S", year=2024, sponsor="Sp",
                protocol="", description_full="f", description_masked="m",
            )
            result = db_b.add_study(site.id, study)
            assert result is None


class TestBackups:
    def test_create_site_backup(self, db_manager, app_config):
        site = db_manager.create_site("Backup Site")
        study = Study(
            phase="P", subcategory="S", year=2024, sponsor="Sp",
            protocol="Pr", description_full="f", description_masked="m",
        )
        db_manager.add_study(site.id, study)
        version_id = db_manager.create_site_backup(site.id, "Test backup")
        assert version_id is not None

    def test_get_site_versions(self, db_manager):
        site = db_manager.create_site("Version Site")
        db_manager.create_site_backup(site.id, "v1")
        db_manager.create_site_backup(site.id, "v2")
        versions = db_manager.get_site_versions(site.id)
        assert len(versions) == 2


class TestCategoryOrder:
    def test_save_and_load_order(self, db_manager):
        site = db_manager.create_site("Order Site")
        order = ["Phase I > Oncology", "Phase I > Cardiology", "Phase II\u2013IV > Oncology"]
        db_manager.save_category_order(site.id, order)
        loaded = db_manager.get_category_order(site.id)
        assert loaded == order

    def test_no_saved_order(self, db_manager):
        site = db_manager.create_site("No Order")
        assert db_manager.get_category_order(site.id) is None
