"""
Tests for the models module.

Covers: Study identity/dedup, sorting, ResearchExperience benchmark calculation,
phase ordering, custom sorting.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent.resolve()))

from models import (
    Study, Subcategory, Phase, ResearchExperience,
    LogEntry, OperationResult, Site, SiteVersion,
)
from datetime import datetime


class TestStudy:
    def test_identity_tuple_with_protocol(self):
        s = Study(
            phase="Phase I", subcategory="Oncology", year=2024,
            sponsor="Pfizer", protocol="PF-123",
            description_full="desc full", description_masked="desc masked",
        )
        t = s.get_identity_tuple()
        assert t == ("Phase I", "Oncology", 2024, "Pfizer", "PF-123", "desc masked")

    def test_identity_tuple_no_protocol(self):
        s = Study(
            phase="Phase I", subcategory="Oncology", year=2024,
            sponsor="Pfizer", protocol="",
            description_full="desc", description_masked="desc",
        )
        t = s.get_identity_tuple()
        assert t[4] == ""

    def test_equality(self):
        s1 = Study(
            phase="Phase I", subcategory="Onc", year=2024,
            sponsor="Pfizer", protocol="PF-1",
            description_full="f", description_masked="m",
        )
        s2 = Study(
            phase="Phase I", subcategory="Onc", year=2024,
            sponsor="Pfizer", protocol="PF-1",
            description_full="f", description_masked="m",
        )
        assert s1 == s2

    def test_inequality_different_year(self):
        s1 = Study(
            phase="Phase I", subcategory="Onc", year=2024,
            sponsor="Pfizer", protocol="PF-1",
            description_full="f", description_masked="m",
        )
        s2 = Study(
            phase="Phase I", subcategory="Onc", year=2023,
            sponsor="Pfizer", protocol="PF-1",
            description_full="f", description_masked="m",
        )
        assert s1 != s2

    def test_hash_consistency(self):
        s1 = Study(
            phase="Phase I", subcategory="Onc", year=2024,
            sponsor="Pfizer", protocol="PF-1",
            description_full="f", description_masked="m",
        )
        s2 = Study(
            phase="Phase I", subcategory="Onc", year=2024,
            sponsor="Pfizer", protocol="PF-1",
            description_full="f", description_masked="m",
        )
        assert hash(s1) == hash(s2)

    def test_format_for_cv_with_protocol(self):
        s = Study(
            phase="Phase I", subcategory="Onc", year=2024,
            sponsor="Pfizer", protocol="PF-123",
            description_full="A Phase 1 study", description_masked="A Phase 1 study of XXX",
        )
        fmt = s.format_for_cv(include_protocol=True)
        assert "2024\t" in fmt
        assert "Pfizer PF-123:" in fmt

    def test_format_for_cv_without_protocol(self):
        s = Study(
            phase="Phase I", subcategory="Onc", year=2024,
            sponsor="Pfizer", protocol="PF-123",
            description_full="A Phase 1 study", description_masked="A Phase 1 study of XXX",
        )
        fmt = s.format_for_cv(include_protocol=False)
        assert "PF-123" not in fmt
        assert "Pfizer:" in fmt


class TestSubcategory:
    def test_sort_studies(self):
        studies = [
            Study(phase="P", subcategory="S", year=2022, sponsor="B",
                  protocol="", description_full="", description_masked=""),
            Study(phase="P", subcategory="S", year=2024, sponsor="A",
                  protocol="", description_full="", description_masked=""),
            Study(phase="P", subcategory="S", year=2024, sponsor="B",
                  protocol="", description_full="", description_masked=""),
        ]
        sc = Subcategory(name="Test", studies=studies)
        sc.sort_studies()
        assert sc.studies[0].year == 2024
        assert sc.studies[0].sponsor == "A"
        assert sc.studies[1].year == 2024
        assert sc.studies[1].sponsor == "B"
        assert sc.studies[2].year == 2022


class TestPhase:
    def test_sort_subcategories(self):
        p = Phase(name="Phase I")
        p.subcategories = [
            Subcategory(name="Neurology"),
            Subcategory(name="Cardiology"),
            Subcategory(name="Oncology"),
        ]
        p.sort_subcategories()
        names = [sc.name for sc in p.subcategories]
        assert names == ["Cardiology", "Neurology", "Oncology"]

    def test_get_or_create_subcategory_existing(self):
        p = Phase(name="Phase I")
        sc1 = p.get_or_create_subcategory("Oncology")
        sc2 = p.get_or_create_subcategory("oncology")  # Case-insensitive
        assert sc1 is sc2

    def test_get_or_create_subcategory_new(self):
        p = Phase(name="Phase I")
        sc = p.get_or_create_subcategory("Oncology")
        assert sc.name == "Oncology"
        assert len(p.subcategories) == 1


class TestResearchExperience:
    def test_phase_order(self):
        re_exp = ResearchExperience()
        re_exp.phases = [
            Phase(name="Phase II\u2013IV"),
            Phase(name="Phase I"),
            Phase(name="Uncategorized"),
        ]
        re_exp.sort_all()
        names = [p.name for p in re_exp.phases]
        assert names[0] == "Phase I"
        assert names[-1] == "Uncategorized"

    def test_calculate_benchmark_year_normal(self):
        re_exp = ResearchExperience()
        phase = re_exp.get_or_create_phase("Phase I")
        sc = phase.get_or_create_subcategory("Onc")
        for i in range(5):
            sc.studies.append(
                Study(phase="Phase I", subcategory="Onc", year=2024,
                      sponsor="S", protocol="", description_full="",
                      description_masked="")
            )
        # 5 studies in 2024 -> benchmark = 2024
        assert re_exp.calculate_benchmark_year(min_count=4) == 2024

    def test_calculate_benchmark_year_few_studies(self):
        re_exp = ResearchExperience()
        phase = re_exp.get_or_create_phase("Phase I")
        sc = phase.get_or_create_subcategory("Onc")
        # 2 studies in 2024 -> benchmark steps back to 2023
        for _ in range(2):
            sc.studies.append(
                Study(phase="Phase I", subcategory="Onc", year=2024,
                      sponsor="S", protocol="", description_full="",
                      description_masked="")
            )
        sc.studies.append(
            Study(phase="Phase I", subcategory="Onc", year=2023,
                  sponsor="S", protocol="", description_full="",
                  description_masked="")
        )
        assert re_exp.calculate_benchmark_year(min_count=4) == 2023

    def test_get_all_studies(self):
        re_exp = ResearchExperience()
        phase = re_exp.get_or_create_phase("Phase I")
        sc = phase.get_or_create_subcategory("Onc")
        sc.studies.append(
            Study(phase="Phase I", subcategory="Onc", year=2024,
                  sponsor="S", protocol="", description_full="",
                  description_masked="")
        )
        assert len(re_exp.get_all_studies()) == 1


class TestOperationResult:
    def test_get_counts(self):
        entries = [
            LogEntry(datetime.now(), "inserted", "P", "S", 2024, "Sp", "Pr", "d"),
            LogEntry(datetime.now(), "inserted", "P", "S", 2024, "Sp", "Pr", "d"),
            LogEntry(datetime.now(), "skipped-duplicate", "P", "S", 2024, "Sp", "Pr", "d"),
        ]
        result = OperationResult(success=True, log_entries=entries)
        counts = result.get_counts()
        assert counts["inserted"] == 2
        assert counts["skipped-duplicate"] == 1
