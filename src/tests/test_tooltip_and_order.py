import sys
import os
import json
import shutil
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

APP_ROOT = Path(__file__).parent.parent.resolve()
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from config import AppConfig, set_config
from models import Study, ResearchExperience
from database import DatabaseManager
from tooltip_text import (
    TOOLTIP_TEXT,
    TOOLTIP_DEFAULT,
    TOOLTIP_MAX_WIDTH,
    get_tooltip_text,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def tmp_dir(tmp_path):
    return tmp_path


@pytest.fixture
def app_config(tmp_dir):
    cfg = AppConfig(data_root=str(tmp_dir / "data"))
    cfg.ensure_user_directories()
    set_config(cfg)
    return cfg


@pytest.fixture
def db(app_config):
    mgr = DatabaseManager(config=app_config)
    yield mgr
    mgr.close()


def _make_study(phase="Phase I", subcategory="Oncology", year=2024,
                sponsor="Pfizer", protocol="PF-001"):
    return Study(
        phase=phase,
        subcategory=subcategory,
        year=year,
        sponsor=sponsor,
        protocol=protocol,
        description_full=f"{sponsor} {protocol}: test study",
        description_masked=f"{sponsor}: test study XXX",
    )


# ===========================================================================
# 1. Tooltip text mapping and retrieval
# ===========================================================================

class TestTooltipText:

    def test_known_keys_have_text(self):
        required_keys = [
            "fuzzy_threshold_full",
            "fuzzy_threshold_masked",
            "benchmark_min_count",
            "highlight_inserted",
            "font_name",
            "font_size",
            "uncategorized_label",
        ]
        for key in required_keys:
            text = get_tooltip_text(key)
            assert text != TOOLTIP_DEFAULT, f"Key '{key}' should have real tooltip text"
            assert len(text) > 10, f"Tooltip for '{key}' too short"

    def test_unknown_key_returns_default(self):
        assert get_tooltip_text("nonexistent_key_xyz") == TOOLTIP_DEFAULT

    def test_tooltip_text_dict_is_not_empty(self):
        assert len(TOOLTIP_TEXT) >= 7

    def test_all_values_are_strings(self):
        for key, val in TOOLTIP_TEXT.items():
            assert isinstance(val, str), f"Value for '{key}' is not a string"

    def test_max_width_is_positive_int(self):
        assert isinstance(TOOLTIP_MAX_WIDTH, int)
        assert TOOLTIP_MAX_WIDTH > 0

    def test_highlight_inserted_mentions_yellow(self):
        text = get_tooltip_text("highlight_inserted")
        assert "yellow" in text.lower()

    def test_uncategorized_label_mentions_label(self):
        text = get_tooltip_text("uncategorized_label")
        assert "label" in text.lower()

    def test_font_name_mentions_font(self):
        text = get_tooltip_text("font_name")
        assert "font" in text.lower()

    def test_backup_retention_days_exists(self):
        text = get_tooltip_text("backup_retention_days")
        assert text != TOOLTIP_DEFAULT

    def test_log_retention_days_exists(self):
        text = get_tooltip_text("log_retention_days")
        assert text != TOOLTIP_DEFAULT


# ===========================================================================
# 2. ConfigToolTip widget (headless — no display server needed)
# ===========================================================================

class TestConfigToolTipWidget:

    @pytest.fixture(autouse=True)
    def _skip_no_display(self):
        try:
            import tkinter as tk
            root = tk.Tk()
            root.withdraw()
            self._root = root
            yield
            root.destroy()
        except Exception:
            pytest.skip("No display available for tkinter tests")

    def test_tooltip_icon_created(self):
        import tkinter as tk
        from gui import ConfigToolTip
        frame = tk.Frame(self._root, bg="#ffffff")
        frame.pack()
        tip = ConfigToolTip(frame, "Test tooltip text")
        assert tip.icon_widget is not None
        assert tip.icon_widget.cget("text") == "\u24d8"

    def test_tooltip_icon_is_focusable(self):
        import tkinter as tk
        from gui import ConfigToolTip
        frame = tk.Frame(self._root, bg="#ffffff")
        frame.pack()
        tip = ConfigToolTip(frame, "Focus test")
        assert str(tip.icon_widget.cget("takefocus")) == "1"

    def test_tooltip_show_and_hide(self):
        import tkinter as tk
        from gui import ConfigToolTip
        frame = tk.Frame(self._root, bg="#ffffff")
        frame.pack()
        tip = ConfigToolTip(frame, "Show/hide test")
        assert tip._tip_window is None
        tip._show()
        assert tip._tip_window is not None
        tip._hide_now()
        assert tip._tip_window is None

    def test_tooltip_toggle(self):
        import tkinter as tk
        from gui import ConfigToolTip
        frame = tk.Frame(self._root, bg="#ffffff")
        frame.pack()
        tip = ConfigToolTip(frame, "Toggle test")
        tip._toggle()
        assert tip._tip_window is not None
        tip._toggle()
        assert tip._tip_window is None

    def test_tooltip_escape_hides(self):
        import tkinter as tk
        from gui import ConfigToolTip
        frame = tk.Frame(self._root, bg="#ffffff")
        frame.pack()
        tip = ConfigToolTip(frame, "Escape test")
        tip._show()
        assert tip._tip_window is not None
        tip._hide_now()
        assert tip._tip_window is None

    def test_double_show_is_safe(self):
        import tkinter as tk
        from gui import ConfigToolTip
        frame = tk.Frame(self._root, bg="#ffffff")
        frame.pack()
        tip = ConfigToolTip(frame, "Double show")
        tip._show()
        tip._show()
        assert tip._tip_window is not None
        tip._hide_now()

    def test_double_hide_is_safe(self):
        import tkinter as tk
        from gui import ConfigToolTip
        frame = tk.Frame(self._root, bg="#ffffff")
        frame.pack()
        tip = ConfigToolTip(frame, "Double hide")
        tip._hide_now()
        tip._hide_now()
        assert tip._tip_window is None


# ===========================================================================
# 3. Category Order auto-update — add_study
# ===========================================================================

class TestCategoryOrderAddStudy:

    def test_add_study_creates_order_entry(self, db):
        site = db.create_site("Test Site")
        study = _make_study(phase="Phase I", subcategory="Oncology")
        db.add_study(site.id, study)

        order = db.get_category_order(site.id)
        assert order is not None
        assert "Phase I > Oncology" in order

    def test_add_study_second_category_appended(self, db):
        site = db.create_site("Test Site")
        db.add_study(site.id, _make_study(phase="Phase I", subcategory="Oncology"))
        db.add_study(site.id, _make_study(phase="Phase I", subcategory="Cardiology",
                                          sponsor="Novartis", protocol="NVS-001"))

        order = db.get_category_order(site.id)
        assert order == ["Phase I > Oncology", "Phase I > Cardiology"]

    def test_add_study_new_phase_appended(self, db):
        site = db.create_site("Test Site")
        db.add_study(site.id, _make_study(phase="Phase I", subcategory="Oncology"))
        db.add_study(site.id, _make_study(
            phase="Phase II\u2013IV", subcategory="Oncology",
            sponsor="Roche", protocol="RO-001",
        ))

        order = db.get_category_order(site.id)
        assert len(order) == 2
        assert "Phase II\u2013IV > Oncology" in order

    def test_idempotency_same_category_no_duplicate(self, db):
        site = db.create_site("Test Site")
        db.add_study(site.id, _make_study(phase="Phase I", subcategory="Oncology"))
        db.add_study(site.id, _make_study(phase="Phase I", subcategory="Oncology",
                                          sponsor="AZ", protocol="AZ-001"))

        order = db.get_category_order(site.id)
        assert order.count("Phase I > Oncology") == 1

    def test_idempotency_repeated_calls_stable(self, db):
        site = db.create_site("Test Site")
        for i in range(5):
            db.add_study(site.id, _make_study(
                phase="Phase I", subcategory="Oncology",
                sponsor=f"Sponsor{i}", protocol=f"P-{i}",
            ))

        order = db.get_category_order(site.id)
        assert order == ["Phase I > Oncology"]


# ===========================================================================
# 4. Category Order auto-update — bulk_add_studies
# ===========================================================================

class TestCategoryOrderBulkAdd:

    def test_bulk_add_creates_all_entries(self, db):
        site = db.create_site("Bulk Site")
        studies = [
            _make_study(phase="Phase I", subcategory="Oncology"),
            _make_study(phase="Phase I", subcategory="Cardiology",
                        sponsor="Novartis", protocol="NVS-001"),
            _make_study(phase="Phase II\u2013IV", subcategory="Oncology",
                        sponsor="Roche", protocol="RO-001"),
        ]
        db.bulk_add_studies(site.id, studies)

        order = db.get_category_order(site.id)
        assert order is not None
        assert len(order) == 3
        assert "Phase I > Oncology" in order
        assert "Phase I > Cardiology" in order
        assert "Phase II\u2013IV > Oncology" in order

    def test_bulk_add_no_duplicates(self, db):
        site = db.create_site("Bulk Site")
        studies = [
            _make_study(phase="Phase I", subcategory="Oncology", sponsor="A", protocol="A-1"),
            _make_study(phase="Phase I", subcategory="Oncology", sponsor="B", protocol="B-1"),
            _make_study(phase="Phase I", subcategory="Oncology", sponsor="C", protocol="C-1"),
        ]
        db.bulk_add_studies(site.id, studies)

        order = db.get_category_order(site.id)
        assert order == ["Phase I > Oncology"]

    def test_bulk_add_then_single_add_appends(self, db):
        site = db.create_site("Mixed Site")
        db.bulk_add_studies(site.id, [
            _make_study(phase="Phase I", subcategory="Oncology"),
        ])
        db.add_study(site.id, _make_study(
            phase="Phase I", subcategory="Dermatology",
            sponsor="Lilly", protocol="LY-001",
        ))

        order = db.get_category_order(site.id)
        assert order == ["Phase I > Oncology", "Phase I > Dermatology"]


# ===========================================================================
# 5. Category Order auto-update — update_study
# ===========================================================================

class TestCategoryOrderUpdateStudy:

    def test_update_study_adds_new_category(self, db):
        site = db.create_site("Update Site")
        study = _make_study(phase="Phase I", subcategory="Oncology")
        db.add_study(site.id, study)

        study.phase = "Phase II\u2013IV"
        study.subcategory = "Neurology"
        db.update_study(study)

        order = db.get_category_order(site.id)
        assert "Phase I > Oncology" in order
        assert "Phase II\u2013IV > Neurology" in order

    def test_update_study_same_category_no_change(self, db):
        site = db.create_site("Update Site")
        study = _make_study(phase="Phase I", subcategory="Oncology")
        db.add_study(site.id, study)

        study.description_full = "Updated description"
        db.update_study(study)

        order = db.get_category_order(site.id)
        assert order == ["Phase I > Oncology"]


# ===========================================================================
# 6. Category Order — import path via bulk_add_studies
# ===========================================================================

class TestCategoryOrderImport:

    def test_import_creates_order_entries(self, db):
        site = db.create_site("Import Site")
        studies = [
            _make_study(phase="Phase I", subcategory="Oncology"),
            _make_study(phase="Phase I", subcategory="Hematology",
                        sponsor="Amgen", protocol="AM-001"),
            _make_study(phase="Phase II\u2013IV", subcategory="Oncology",
                        sponsor="BMS", protocol="BMS-001"),
        ]
        db.bulk_add_studies(site.id, studies)

        order = db.get_category_order(site.id)
        assert len(order) == 3

    def test_import_with_existing_order_preserves_then_appends(self, db):
        site = db.create_site("Import Site")
        db.save_category_order(site.id, ["Phase I > Oncology"])

        studies = [
            _make_study(phase="Phase I", subcategory="Oncology"),
            _make_study(phase="Phase I", subcategory="Cardiology",
                        sponsor="AZ", protocol="AZ-001"),
        ]
        db.bulk_add_studies(site.id, studies)

        order = db.get_category_order(site.id)
        assert order[0] == "Phase I > Oncology"
        assert "Phase I > Cardiology" in order
        assert len(order) == 2


# ===========================================================================
# 7. Category Order — _ensure_category_order_entries directly
# ===========================================================================

class TestEnsureCategoryOrderEntries:

    def test_empty_entries_is_noop(self, db):
        site = db.create_site("Test")
        db._ensure_category_order_entries(site.id, [])
        assert db.get_category_order(site.id) is None

    def test_creates_order_from_scratch(self, db):
        site = db.create_site("Test")
        db._ensure_category_order_entries(site.id, [
            ("Phase I", "Oncology"),
            ("Phase I", "Cardiology"),
        ])
        order = db.get_category_order(site.id)
        assert order == ["Phase I > Oncology", "Phase I > Cardiology"]

    def test_idempotent_on_repeated_calls(self, db):
        site = db.create_site("Test")
        entries = [("Phase I", "Oncology")]
        db._ensure_category_order_entries(site.id, entries)
        db._ensure_category_order_entries(site.id, entries)
        db._ensure_category_order_entries(site.id, entries)
        order = db.get_category_order(site.id)
        assert order == ["Phase I > Oncology"]

    def test_preserves_existing_order(self, db):
        site = db.create_site("Test")
        db.save_category_order(site.id, [
            "Phase II\u2013IV > Oncology",
            "Phase I > Oncology",
        ])
        db._ensure_category_order_entries(site.id, [
            ("Phase I", "Oncology"),
            ("Phase I", "Cardiology"),
        ])
        order = db.get_category_order(site.id)
        assert order[0] == "Phase II\u2013IV > Oncology"
        assert order[1] == "Phase I > Oncology"
        assert order[2] == "Phase I > Cardiology"

    def test_same_subcategory_different_phase(self, db):
        site = db.create_site("Test")
        db._ensure_category_order_entries(site.id, [
            ("Phase I", "Oncology"),
            ("Phase II\u2013IV", "Oncology"),
        ])
        order = db.get_category_order(site.id)
        assert len(order) == 2
        assert "Phase I > Oncology" in order
        assert "Phase II\u2013IV > Oncology" in order


# ===========================================================================
# 8. Category Order — sorting uses DB order
# ===========================================================================

class TestSortingUsesDBOrder:

    def test_sort_all_custom_respects_order(self):
        re_exp = ResearchExperience()
        p1 = re_exp.get_or_create_phase("Phase I")
        sc_onc = p1.get_or_create_subcategory("Oncology")
        sc_onc.studies.append(_make_study())
        sc_card = p1.get_or_create_subcategory("Cardiology")
        sc_card.studies.append(_make_study(subcategory="Cardiology",
                                           sponsor="AZ", protocol="AZ-1"))

        custom_order = ["Phase I > Cardiology", "Phase I > Oncology"]
        re_exp.sort_all_custom(custom_order)

        names = [sc.name for p in re_exp.phases for sc in p.subcategories]
        assert names == ["Cardiology", "Oncology"]

    def test_new_category_appears_last_in_custom_sort(self):
        re_exp = ResearchExperience()
        p1 = re_exp.get_or_create_phase("Phase I")
        sc_onc = p1.get_or_create_subcategory("Oncology")
        sc_onc.studies.append(_make_study())
        sc_card = p1.get_or_create_subcategory("Cardiology")
        sc_card.studies.append(_make_study(subcategory="Cardiology",
                                           sponsor="AZ", protocol="AZ-1"))
        sc_new = p1.get_or_create_subcategory("Neurology")
        sc_new.studies.append(_make_study(subcategory="Neurology",
                                          sponsor="BMS", protocol="BMS-1"))

        custom_order = ["Phase I > Oncology", "Phase I > Cardiology"]
        re_exp.sort_all_custom(custom_order)

        names = [sc.name for p in re_exp.phases for sc in p.subcategories]
        assert names[-1] == "Neurology"

    def test_deterministic_sorting_across_runs(self, db):
        site = db.create_site("Deterministic")
        studies = [
            _make_study(phase="Phase I", subcategory="Zebra",
                        sponsor="Z", protocol="Z-1"),
            _make_study(phase="Phase I", subcategory="Alpha",
                        sponsor="A", protocol="A-1"),
            _make_study(phase="Phase I", subcategory="Middle",
                        sponsor="M", protocol="M-1"),
        ]
        db.bulk_add_studies(site.id, studies)

        order1 = db.get_category_order(site.id)

        db._ensure_category_order_entries(site.id, [
            ("Phase I", "Zebra"),
            ("Phase I", "Alpha"),
            ("Phase I", "Middle"),
        ])
        order2 = db.get_category_order(site.id)

        assert order1 == order2


# ===========================================================================
# 9. Category Order — delete does not crash
# ===========================================================================

class TestCategoryOrderDelete:

    def test_delete_study_does_not_remove_order_entry(self, db):
        site = db.create_site("Delete Test")
        study = _make_study(phase="Phase I", subcategory="Oncology")
        db.add_study(site.id, study)
        order_before = db.get_category_order(site.id)

        db.delete_study(study.id, site.id)
        order_after = db.get_category_order(site.id)

        assert order_before == order_after

    def test_delete_site_cascades_order(self, db):
        site = db.create_site("Cascade Test")
        db.add_study(site.id, _make_study())
        db.delete_site(site.id)
        assert db.get_category_order(site.id) is None


# ===========================================================================
# 10. Integration: Mode A with site_id picks up new categories
# ===========================================================================

class TestModeAOrderIntegration:

    def test_mode_a_with_site_ensures_order(self, app_config, tmp_dir):
        from tests.conftest import _make_cv_docx, _make_master_xlsx_seven_col

        cfg = app_config
        set_config(cfg)

        cv_path = _make_cv_docx(tmp_dir / "cv.docx")
        master_path = tmp_dir / "master.xlsx"
        _make_master_xlsx_seven_col(master_path)

        with DatabaseManager(config=cfg) as db:
            site = db.create_site("Order Test Site")
            from excel_parser import parse_master_xlsx_seven_col
            studies = parse_master_xlsx_seven_col(master_path)
            db.bulk_add_studies(site.id, studies)

            order = db.get_category_order(site.id)
            assert order is not None
            assert len(order) > 0

        from processor import CVProcessor
        processor = CVProcessor(cfg)
        result = processor.mode_a_update_inject(
            cv_path,
            site_id=site.id,
            output_path=tmp_dir / "output.docx",
        )
        assert result.success

        with DatabaseManager(config=cfg) as db:
            order_after = db.get_category_order(site.id)
            assert order_after is not None
            assert len(order_after) >= len(order)


# ===========================================================================
# 11. Cleanup verification
# ===========================================================================

class TestCleanup:

    def test_no_artifacts_in_real_result_folder(self, app_config, tmp_dir):
        cfg = app_config
        result_root = cfg.get_result_root()
        assert str(tmp_dir) in str(result_root)
