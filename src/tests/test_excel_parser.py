"""
Tests for the excel_parser module.

Covers: parsing master xlsx, column B parsing, hierarchy detection,
year/subcategory/phase parsing, export round-trip, validation.
"""

import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).parent.parent.resolve()))

from excel_parser import (
    parse_master_xlsx,
    parse_column_b,
    studies_to_research_experience,
    export_studies_to_xlsx,
    validate_master_xlsx,
)
from models import Study


class TestParseMasterXlsx:
    def test_parse_default_sample(self, sample_master_xlsx):
        studies = parse_master_xlsx(sample_master_xlsx)
        assert len(studies) >= 5
        # All should have phases
        phases = {s.phase for s in studies}
        assert "Phase I" in phases

    def test_phase_assignment(self, sample_master_xlsx):
        studies = parse_master_xlsx(sample_master_xlsx)
        # First studies should be Phase I
        phase_i_studies = [s for s in studies if s.phase == "Phase I"]
        assert len(phase_i_studies) >= 2

    def test_year_parsing(self, sample_master_xlsx):
        studies = parse_master_xlsx(sample_master_xlsx)
        for s in studies:
            assert 1900 <= s.year <= 2100

    def test_sponsor_extraction(self, sample_master_xlsx):
        studies = parse_master_xlsx(sample_master_xlsx)
        sponsors = {s.sponsor for s in studies}
        assert "Pfizer" in sponsors

    def test_protocol_extraction(self, sample_master_xlsx):
        studies = parse_master_xlsx(sample_master_xlsx)
        protocols = {s.protocol for s in studies if s.protocol}
        assert len(protocols) > 0

    def test_masked_description(self, sample_master_xlsx):
        studies = parse_master_xlsx(sample_master_xlsx)
        for s in studies:
            if s.description_masked:
                # Masked descriptions should not contain protocol tokens
                # (they should have XXX instead)
                assert "XXX" in s.description_masked or s.description_masked == s.description_full


class TestParseColumnB:
    def test_with_protocol(self):
        sponsor, protocol, desc = parse_column_b("Pfizer PF-123: A Phase 1 study")
        assert sponsor == "Pfizer"
        assert protocol == "PF-123"
        # Description may be phase-normalized (Phase 1 -> Phase I)
        assert "study" in desc.lower()

    def test_no_protocol(self):
        sponsor, protocol, desc = parse_column_b("Pfizer: A study description")
        assert sponsor == "Pfizer"
        assert protocol == ""
        assert "study description" in desc

    def test_empty(self):
        sponsor, protocol, desc = parse_column_b("")
        assert sponsor == ""
        assert protocol == ""
        assert desc == ""

    def test_no_colon(self):
        sponsor, protocol, desc = parse_column_b("Pfizer PF-123")
        assert sponsor == "Pfizer"
        assert desc == ""


class TestValidateMasterXlsx:
    def test_valid_file(self, sample_master_xlsx):
        is_valid, error = validate_master_xlsx(sample_master_xlsx)
        assert is_valid is True
        assert error == ""

    def test_empty_file(self, empty_master_xlsx):
        is_valid, error = validate_master_xlsx(empty_master_xlsx)
        assert is_valid is False
        assert "empty" in error.lower() or "phase" in error.lower() or "study" in error.lower()

    def test_nonexistent_file(self, tmp_dir):
        is_valid, error = validate_master_xlsx(tmp_dir / "nonexistent.xlsx")
        assert is_valid is False
        assert "not found" in error.lower()

    def test_wrong_extension(self, tmp_dir):
        bad_file = tmp_dir / "test.csv"
        bad_file.write_text("data")
        is_valid, error = validate_master_xlsx(bad_file)
        assert is_valid is False


class TestExportStudies:
    def test_round_trip(self, tmp_dir, sample_master_xlsx):
        """Parse, then export, then parse again — counts should match."""
        studies = parse_master_xlsx(sample_master_xlsx)
        export_path = tmp_dir / "exported.xlsx"
        export_studies_to_xlsx(studies, export_path)

        assert export_path.exists()
        # Re-parse
        re_parsed = parse_master_xlsx(export_path)
        assert len(re_parsed) == len(studies)


class TestStudiesToResearchExperience:
    def test_structure(self, sample_master_xlsx):
        studies = parse_master_xlsx(sample_master_xlsx)
        re_exp = studies_to_research_experience(studies)
        assert len(re_exp.phases) > 0
        all_studies = re_exp.get_all_studies()
        assert len(all_studies) == len(studies)
