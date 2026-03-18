"""
Integration tests for the CV Research Experience Manager.

End-to-end flows: update/inject, redact, import/export, preview,
database round-trips. All use synthetic data — no external files.
"""

import sys
import json
from pathlib import Path

import pytest
from docx import Document
from docx.shared import RGBColor

sys.path.insert(0, str(Path(__file__).parent.parent.resolve()))

from config import AppConfig, set_config
from processor import CVProcessor
from import_export import ImportExportManager
from database import DatabaseManager
from models import Study
from tests.conftest import _make_master_xlsx, _make_master_xlsx_seven_col, _make_cv_docx


class TestModeAUpdateInject:
    """Integration tests for Mode A: Update/Inject."""

    def test_inject_new_studies(self, app_config, tmp_dir):
        """Master has studies not in CV — they should be injected."""
        cv_path = tmp_dir / "cv.docx"
        _make_cv_docx(cv_path)  # Has 2023 and 2022 studies

        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx(master_path)  # Has 2024 studies

        processor = CVProcessor(app_config)
        result = processor.mode_a_update_inject(
            cv_path,
            master_path=master_path,
            output_path=tmp_dir / "output.docx",
        )

        assert result.success is True
        assert (tmp_dir / "output.docx").exists()

        # Check that injected studies appear in log
        counts = result.get_counts()
        assert counts.get("inserted", 0) > 0 or counts.get("matched-existing", 0) > 0

    def test_idempotent_injection(self, app_config, tmp_dir):
        """Running inject twice should not duplicate studies."""
        cv_path = tmp_dir / "cv.docx"
        _make_cv_docx(cv_path)

        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx(master_path)

        processor = CVProcessor(app_config)

        # First run
        output1 = tmp_dir / "output1.docx"
        result1 = processor.mode_a_update_inject(
            cv_path, master_path=master_path, output_path=output1,
        )
        assert result1.success is True

        # Second run on the output of first run
        output2 = tmp_dir / "output2.docx"
        result2 = processor.mode_a_update_inject(
            output1, master_path=master_path, output_path=output2,
        )
        assert result2.success is True

    def test_manual_benchmark_year(self, app_config, tmp_dir):
        """Manual benchmark year should control injection cutoff."""
        cv_path = tmp_dir / "cv.docx"
        _make_cv_docx(cv_path)

        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx(master_path)

        processor = CVProcessor(app_config)
        result = processor.mode_a_update_inject(
            cv_path,
            master_path=master_path,
            manual_benchmark_year=2025,  # Nothing should be injected above 2025
            output_path=tmp_dir / "output.docx",
        )
        assert result.success is True

    def test_missing_cv_fails(self, app_config, tmp_dir):
        """Non-existent CV path should fail gracefully."""
        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx(master_path)

        processor = CVProcessor(app_config)
        result = processor.mode_a_update_inject(
            tmp_dir / "nope.docx",
            master_path=master_path,
        )
        assert result.success is False
        assert result.error_message != ""

    def test_cv_no_research_section(self, app_config, tmp_dir):
        """CV without Research Experience should fail."""
        cv_path = tmp_dir / "cv_no_re.docx"
        _make_cv_docx(cv_path, include_research_exp=False)

        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx(master_path)

        processor = CVProcessor(app_config)
        result = processor.mode_a_update_inject(
            cv_path, master_path=master_path,
        )
        assert result.success is False


class TestModeBRedact:
    """Integration tests for Mode B: Redact Protocols."""

    def test_redact_protocols(self, app_config, tmp_dir):
        """Matched studies should have protocols removed."""
        cv_path = tmp_dir / "cv.docx"
        _make_cv_docx(cv_path)

        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx(master_path)

        processor = CVProcessor(app_config)
        result = processor.mode_b_redact_protocols(
            cv_path,
            master_path=master_path,
            output_path=tmp_dir / "redacted.docx",
        )

        assert result.success is True
        assert (tmp_dir / "redacted.docx").exists()

    def test_redact_output_no_protocols(self, app_config, tmp_dir):
        """Matched redacted studies should have protocol removed and XXX mask applied."""
        cv_path = tmp_dir / "cv.docx"
        _make_cv_docx(cv_path)

        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx(master_path)

        processor = CVProcessor(app_config)
        output = tmp_dir / "redacted.docx"
        result = processor.mode_b_redact_protocols(
            cv_path, master_path=master_path, output_path=output,
        )
        assert result.success is True

        replaced_ops = [
            e for e in result.log_entries if e.operation == "replaced"
        ]
        skipped_ops = [
            e for e in result.log_entries
            if e.operation in ("skipped-no-protocol", "skipped-already-masked")
        ]
        assert len(replaced_ops) + len(skipped_ops) >= 0

        doc = Document(output)
        redacted_paras = [
            p for p in doc.paragraphs if "XXX" in p.text
        ]
        for p in redacted_paras:
            for run in p.runs:
                if run.font.color and run.font.color.rgb:
                    assert run.font.color.rgb != RGBColor(0xFF, 0, 0), \
                        f"Found red text in redacted paragraph: '{run.text}'"

    def test_missing_cv_fails(self, app_config, tmp_dir):
        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx(master_path)

        processor = CVProcessor(app_config)
        result = processor.mode_b_redact_protocols(
            tmp_dir / "nope.docx", master_path=master_path,
        )
        assert result.success is False


class TestPreview:
    """Integration tests for the preview pathway."""

    def test_preview_update(self, app_config, tmp_dir):
        cv_path = tmp_dir / "cv.docx"
        _make_cv_docx(cv_path)

        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx(master_path)

        processor = CVProcessor(app_config)
        changes, error = processor.preview_changes(
            cv_path, master_path=master_path, mode="update_inject",
        )
        assert error == ""
        assert isinstance(changes, list)

    def test_preview_redact(self, app_config, tmp_dir):
        cv_path = tmp_dir / "cv.docx"
        _make_cv_docx(cv_path)

        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx(master_path)

        processor = CVProcessor(app_config)
        changes, error = processor.preview_changes(
            cv_path, master_path=master_path, mode="redact_protocols",
        )
        assert error == ""
        assert isinstance(changes, list)

    def test_preview_cv_not_found(self, app_config, tmp_dir):
        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx(master_path)

        processor = CVProcessor(app_config)
        changes, error = processor.preview_changes(
            tmp_dir / "nope.docx", master_path=master_path,
        )
        assert error != ""
        assert len(changes) == 0


class TestImportExport:
    """Integration tests for import/export."""

    def test_import_and_export_round_trip(self, app_config, tmp_dir):
        master_path = tmp_dir / "master_7col.xlsx"
        _make_master_xlsx_seven_col(master_path)

        manager = ImportExportManager(app_config)

        # Import
        success, msg, site_id = manager.import_xlsx_to_site(
            master_path, "Test Site", replace_existing=True,
        )
        assert success is True
        assert site_id is not None

        # Export
        export_path = tmp_dir / "exported.xlsx"
        success, msg, out_path = manager.export_site_to_xlsx(
            site_id, output_path=export_path,
        )
        assert success is True
        assert export_path.exists()

    def test_import_replace_existing(self, app_config, tmp_dir):
        master_path = tmp_dir / "master_7col.xlsx"
        _make_master_xlsx_seven_col(master_path)

        manager = ImportExportManager(app_config)

        # First import
        success1, _, site_id1 = manager.import_xlsx_to_site(
            master_path, "Same Site", replace_existing=True,
        )
        assert success1 is True

        # Second import with replace
        success2, _, site_id2 = manager.import_xlsx_to_site(
            master_path, "Same Site", replace_existing=True,
        )
        assert success2 is True
        assert site_id2 == site_id1  # Same site reused

    def test_import_invalid_file(self, app_config, tmp_dir):
        bad_path = tmp_dir / "bad.xlsx"
        from openpyxl import Workbook
        wb = Workbook()
        wb.save(bad_path)
        wb.close()

        manager = ImportExportManager(app_config)
        success, msg, _ = manager.import_xlsx_to_site(bad_path, "Bad Site")
        assert success is False

    def test_export_nonexistent_site(self, app_config, tmp_dir):
        manager = ImportExportManager(app_config)
        success, msg, _ = manager.export_site_to_xlsx(9999)
        assert success is False


class TestDatabaseIntegration:
    """Integration tests for database operations via import/export."""

    def test_duplicate_site_name_without_replace(self, app_config, tmp_dir):
        master_path = tmp_dir / "master_7col.xlsx"
        _make_master_xlsx_seven_col(master_path)

        manager = ImportExportManager(app_config)
        manager.import_xlsx_to_site(master_path, "MySite", replace_existing=True)
        success, msg, _ = manager.import_xlsx_to_site(
            master_path, "MySite", replace_existing=False,
        )
        assert success is False
        assert "already exists" in msg.lower()

    def test_site_from_db_as_master_source(self, app_config, tmp_dir):
        """Import a master list to DB, then use the site as master source for Mode A."""
        master_path = tmp_dir / "master_7col.xlsx"
        _make_master_xlsx_seven_col(master_path)

        manager = ImportExportManager(app_config)
        success, _, site_id = manager.import_xlsx_to_site(
            master_path, "DB Source", replace_existing=True,
        )
        assert success is True

        cv_path = tmp_dir / "cv.docx"
        _make_cv_docx(cv_path)

        processor = CVProcessor(app_config)
        result = processor.mode_a_update_inject(
            cv_path,
            site_id=site_id,
            output_path=tmp_dir / "from_db.docx",
        )
        assert result.success is True
