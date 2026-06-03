"""
Tests for Mode D (Study Browser) and User Preferences.
"""

import pytest
import json
from pathlib import Path
from user_preferences import UserPreferences, UserPreferencesManager
from config import AppConfig


class TestUserPreferences:
    """Test user preferences data structure and persistence."""
    
    def test_default_preferences(self):
        """Test default preference values."""
        prefs = UserPreferences()
        assert prefs.fullscreen is False
        assert prefs.window_width == 1000
        assert prefs.window_height == 900
        assert prefs.mode_d_selected_site_id is None
        assert prefs.mode_d_show_protocol is True
        assert prefs.mode_d_hanging_indent == 0.5
        assert prefs.mode_d_font_name == "Calibri"
        assert prefs.mode_d_font_size == 11
        assert prefs.mode_d_sponsor_bold is True
        assert prefs.mode_d_sponsor_color == "#000000"
        assert prefs.mode_d_protocol_color == "#000000"
        assert prefs.mode_d_filter_years == []
        assert prefs.mode_d_filter_phases == []
        assert prefs.mode_d_filter_subcategories == []
        assert prefs.mode_c_column_widths == {}
    
    def test_to_dict(self):
        """Test conversion to dictionary."""
        prefs = UserPreferences(fullscreen=True, window_width=1200)
        data = prefs.to_dict()
        assert isinstance(data, dict)
        assert data['fullscreen'] is True
        assert data['window_width'] == 1200
    
    def test_from_dict(self):
        """Test creation from dictionary."""
        data = {
            'fullscreen': True,
            'window_width': 1200,
            'mode_d_font_size': 14,
            'unknown_field': 'ignored'
        }
        prefs = UserPreferences.from_dict(data)
        assert prefs.fullscreen is True
        assert prefs.window_width == 1200
        assert prefs.mode_d_font_size == 14
        assert not hasattr(prefs, 'unknown_field')
    
    def test_save_and_load(self, tmp_path):
        """Test saving and loading preferences."""
        config = AppConfig(data_root=str(tmp_path))
        config.ensure_user_directories()
        
        manager = UserPreferencesManager(config=config)
        
        # Create and save preferences
        prefs = UserPreferences(
            fullscreen=True,
            window_width=1400,
            mode_d_selected_site_id=5,
            mode_d_font_size=12
        )
        manager.save(prefs)
        
        # Load and verify
        loaded_prefs = manager.load()
        assert loaded_prefs.fullscreen is True
        assert loaded_prefs.window_width == 1400
        assert loaded_prefs.mode_d_selected_site_id == 5
        assert loaded_prefs.mode_d_font_size == 12
    
    def test_load_nonexistent_returns_defaults(self, tmp_path):
        """Test loading when file doesn't exist returns defaults."""
        config = AppConfig(data_root=str(tmp_path))
        manager = UserPreferencesManager(config=config)
        
        prefs = manager.load()
        assert isinstance(prefs, UserPreferences)
        assert prefs.fullscreen is False
    
    def test_load_corrupted_file_returns_defaults(self, tmp_path):
        """Test loading corrupted file returns defaults."""
        config = AppConfig(data_root=str(tmp_path))
        config.ensure_user_directories()
        manager = UserPreferencesManager(config=config)
        
        # Write corrupted JSON
        prefs_path = config.get_user_data_path() / "preferences.json"
        prefs_path.write_text("{ invalid json }")
        
        prefs = manager.load()
        assert isinstance(prefs, UserPreferences)
        assert prefs.fullscreen is False
    
    def test_mode_d_filter_persistence(self, tmp_path):
        """Test Mode D filter preferences persist correctly."""
        config = AppConfig(data_root=str(tmp_path))
        config.ensure_user_directories()
        manager = UserPreferencesManager(config=config)
        
        prefs = UserPreferences(
            mode_d_filter_years=[2020, 2021, 2022],
            mode_d_filter_phases=["Phase I", "Phase II–IV"],
            mode_d_filter_subcategories=["Healthy Adults", "Vaccines"]
        )
        manager.save(prefs)
        
        loaded = manager.load()
        assert loaded.mode_d_filter_years == [2020, 2021, 2022]
        assert loaded.mode_d_filter_phases == ["Phase I", "Phase II–IV"]
        assert loaded.mode_d_filter_subcategories == ["Healthy Adults", "Vaccines"]
    
    def test_mode_c_column_widths_persistence(self, tmp_path):
        """Test Mode C column widths persist correctly."""
        config = AppConfig(data_root=str(tmp_path))
        config.ensure_user_directories()
        manager = UserPreferencesManager(config=config)
        
        prefs = UserPreferences(
            mode_c_column_widths={
                'phase': 100,
                'subcategory': 150,
                'year': 60,
                'sponsor': 140
            }
        )
        manager.save(prefs)
        
        loaded = manager.load()
        assert loaded.mode_c_column_widths['phase'] == 100
        assert loaded.mode_c_column_widths['subcategory'] == 150
        assert loaded.mode_c_column_widths['year'] == 60
        assert loaded.mode_c_column_widths['sponsor'] == 140
    
    def test_search_terms_persistence(self, tmp_path):
        """Test search terms persist correctly."""
        config = AppConfig(data_root=str(tmp_path))
        config.ensure_user_directories()
        manager = UserPreferencesManager(config=config)
        
        multi_search = "Healthy\nVaccine\nPhase I"
        prefs = UserPreferences(
            mode_d_single_search="diabetes",
            mode_d_multi_search=multi_search
        )
        manager.save(prefs)
        
        loaded = manager.load()
        assert loaded.mode_d_single_search == "diabetes"
        assert loaded.mode_d_multi_search == multi_search
    
    def test_formatting_preferences(self, tmp_path):
        """Test formatting preferences persist correctly."""
        config = AppConfig(data_root=str(tmp_path))
        config.ensure_user_directories()
        manager = UserPreferencesManager(config=config)
        
        prefs = UserPreferences(
            mode_d_hanging_indent=0.75,
            mode_d_font_name="Times New Roman",
            mode_d_font_size=14,
            mode_d_sponsor_bold=False,
            mode_d_sponsor_color="#FF0000",
            mode_d_protocol_color="#0000FF"
        )
        manager.save(prefs)
        
        loaded = manager.load()
        assert loaded.mode_d_hanging_indent == 0.75
        assert loaded.mode_d_font_name == "Times New Roman"
        assert loaded.mode_d_font_size == 14
        assert loaded.mode_d_sponsor_bold is False
        assert loaded.mode_d_sponsor_color == "#FF0000"
        assert loaded.mode_d_protocol_color == "#0000FF"
    
    def test_fullscreen_toggle_persistence(self, tmp_path):
        """Test fullscreen state persists across sessions."""
        config = AppConfig(data_root=str(tmp_path))
        config.ensure_user_directories()
        manager = UserPreferencesManager(config=config)
        
        # Start not fullscreen
        prefs = UserPreferences(fullscreen=False, window_width=1000, window_height=900)
        manager.save(prefs)
        
        # Toggle to fullscreen
        prefs.fullscreen = True
        manager.save(prefs)
        
        loaded = manager.load()
        assert loaded.fullscreen is True
        
        # Toggle back
        prefs.fullscreen = False
        prefs.window_width = 1200
        prefs.window_height = 1000
        manager.save(prefs)
        
        loaded = manager.load()
        assert loaded.fullscreen is False
        assert loaded.window_width == 1200
        assert loaded.window_height == 1000
    
    def test_get_method_caches(self, tmp_path):
        """Test that get() method caches loaded preferences."""
        config = AppConfig(data_root=str(tmp_path))
        config.ensure_user_directories()
        manager = UserPreferencesManager(config=config)
        
        prefs1 = manager.get()
        prefs2 = manager.get()
        
        assert prefs1 is prefs2  # Same object
    
    def test_multiple_users_separate_preferences(self, tmp_path):
        """Test that different users have separate preferences."""
        config = AppConfig(data_root=str(tmp_path))
        
        # User 1
        manager1 = UserPreferencesManager(user_id="user1", config=config)
        prefs1 = UserPreferences(fullscreen=True, mode_d_font_size=12)
        manager1.save(prefs1)
        
        # User 2
        manager2 = UserPreferencesManager(user_id="user2", config=config)
        prefs2 = UserPreferences(fullscreen=False, mode_d_font_size=14)
        manager2.save(prefs2)
        
        # Verify separation
        loaded1 = manager1.load()
        loaded2 = manager2.load()
        
        assert loaded1.fullscreen is True
        assert loaded1.mode_d_font_size == 12
        assert loaded2.fullscreen is False
        assert loaded2.mode_d_font_size == 14


class TestModeDIntegration:
    """Integration tests for Mode D functionality."""
    
    def test_mode_d_preferences_initialized(self, tmp_path):
        """Test that Mode D preferences are properly initialized."""
        config = AppConfig(data_root=str(tmp_path))
        manager = UserPreferencesManager(config=config)
        prefs = manager.load()
        
        # Verify all Mode D fields exist
        assert hasattr(prefs, 'mode_d_selected_site_id')
        assert hasattr(prefs, 'mode_d_multi_search')
        assert hasattr(prefs, 'mode_d_single_search')
        assert hasattr(prefs, 'mode_d_show_protocol')
        assert hasattr(prefs, 'mode_d_hanging_indent')
        assert hasattr(prefs, 'mode_d_font_name')
        assert hasattr(prefs, 'mode_d_font_size')
        assert hasattr(prefs, 'mode_d_sponsor_bold')
        assert hasattr(prefs, 'mode_d_sponsor_color')
        assert hasattr(prefs, 'mode_d_protocol_color')
        assert hasattr(prefs, 'mode_d_filter_years')
        assert hasattr(prefs, 'mode_d_filter_phases')
        assert hasattr(prefs, 'mode_d_filter_subcategories')
    
    def test_mode_c_preferences_initialized(self, tmp_path):
        """Test that Mode C preferences are properly initialized."""
        config = AppConfig(data_root=str(tmp_path))
        manager = UserPreferencesManager(config=config)
        prefs = manager.load()
        
        assert hasattr(prefs, 'mode_c_column_widths')
        assert isinstance(prefs.mode_c_column_widths, dict)
