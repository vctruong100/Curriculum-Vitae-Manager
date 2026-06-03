"""
User preferences management for GUI state persistence.
Stores per-user UI preferences like window state, selected site, filters, etc.
"""

import json
import logging
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Optional, List, Dict, Any

from config import get_config, AppConfig

pref_logger = logging.getLogger(__name__)


@dataclass
class UserPreferences:
    """User-specific GUI preferences."""
    
    # Window state
    fullscreen: bool = False
    window_width: int = 1000
    window_height: int = 900
    
    # Mode D preferences
    mode_d_selected_site_id: Optional[int] = None
    mode_d_multi_search: str = ""
    mode_d_single_search: str = ""
    mode_d_show_protocol: bool = True
    mode_d_hanging_indent: float = 0.5
    mode_d_font_name: str = "Calibri"
    mode_d_font_size: int = 11
    mode_d_sponsor_bold: bool = True
    mode_d_sponsor_color: str = "#000000"
    mode_d_protocol_color: str = "#000000"
    mode_d_filter_years: List[int] = field(default_factory=list)
    mode_d_filter_phases: List[str] = field(default_factory=list)
    mode_d_filter_subcategories: List[str] = field(default_factory=list)
    
    # Mode C preferences
    mode_c_column_widths: Dict[str, int] = field(default_factory=dict)
    
    def to_dict(self) -> dict:
        """Convert preferences to dictionary."""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: dict) -> "UserPreferences":
        """Create preferences from dictionary."""
        known_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered = {k: v for k, v in data.items() if k in known_fields}
        return cls(**filtered)


class UserPreferencesManager:
    """Manages loading and saving user preferences."""
    
    def __init__(self, user_id: Optional[str] = None, config: Optional[AppConfig] = None):
        self.config = config or get_config()
        self.user_id = user_id or self.config.get_user_id()
        self.prefs_path = self.config.get_user_data_path(self.user_id) / "preferences.json"
        self._preferences: Optional[UserPreferences] = None
    
    def load(self) -> UserPreferences:
        """Load user preferences from file, or return defaults."""
        if self.prefs_path.exists():
            try:
                with open(self.prefs_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self._preferences = UserPreferences.from_dict(data)
                pref_logger.info("[Preferences] Loaded from %s", self.prefs_path)
                return self._preferences
            except (json.JSONDecodeError, IOError) as exc:
                pref_logger.warning("[Preferences] Failed to load: %s", exc)
        
        self._preferences = UserPreferences()
        pref_logger.info("[Preferences] Using defaults")
        return self._preferences
    
    def save(self, preferences: UserPreferences) -> None:
        """Save user preferences to file."""
        try:
            self.prefs_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.prefs_path, 'w', encoding='utf-8') as f:
                json.dump(preferences.to_dict(), f, indent=2)
            self._preferences = preferences
            pref_logger.debug("[Preferences] Saved to %s", self.prefs_path)
        except IOError as exc:
            pref_logger.error("[Preferences] Failed to save: %s", exc)
    
    def get(self) -> UserPreferences:
        """Get current preferences (loads if not already loaded)."""
        if self._preferences is None:
            return self.load()
        return self._preferences
