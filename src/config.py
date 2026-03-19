"""
Configuration management for the CV Research Experience Manager.
All settings are local-only with no network capabilities.
"""

import sys
import os
import json
import getpass
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import List, Optional


APP_NAME = "CV Research Experience Manager"
APP_VERSION = "1.2.0"

DEFAULT_UNCATEGORIZED_LABEL = "Uncategorized"

HANGING_INDENT_MIN = 0.0
HANGING_INDENT_MAX = 2.0
HANGING_INDENT_DEFAULT = 0.5

UNDO_TIMEOUT_SECONDS = 300

DEFAULT_ICON_PATH = "build/assets/app.ico"
BUILD_ICON_PATH = "build/assets/app.ico"

UPDATE_CHECK_URL = "https://api.github.com/repos/vctruong100/Curriculum-Vitae-Manager/releases/latest"


ALLOWED_FONTS = [
    "Calibri",
    "Times New Roman",
    "Garamond",
    "Helvetica",
    "Roboto",
    "Open Sans",
    "Lato",
    "Didot",
]


def get_app_root() -> Path:
    """Get the application root directory (project root).

    When frozen (PyInstaller .exe at project root), returns the directory
    containing the executable.  When running from source, returns the
    parent of src/.  Both resolve to the same project root where ./data/
    lives.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent.resolve()
    return Path(__file__).parent.parent.resolve()


def get_default_data_root() -> Path:
    """Get the default data root directory."""
    return get_app_root() / "data"


def get_os_username() -> str:
    """Get the current OS username."""
    return getpass.getuser()


@dataclass
class AppConfig:
    """Application configuration - all local, no network."""
    
    # Fuzzy matching thresholds
    fuzzy_threshold_full: int = 92
    fuzzy_threshold_masked: int = 90
    
    # Benchmark calculation
    benchmark_min_count: int = 4  # ≤3 triggers step-back
    
    # Formatting options
    highlight_inserted: bool = False
    use_track_changes: bool = False
    
    # Phase ordering
    phase_order: List[str] = field(default_factory=lambda: ["Phase I", "Phase II–IV"])
    
    # Security - network disabled
    network_enabled: bool = False
    allow_redaction_without_full_match: bool = False
    
    # Storage paths (local only)
    data_root: str = ""
    user_id_strategy: str = "os_username"  # "os_username" or "app_username"
    
    # Font settings
    font_name: str = "Calibri"
    font_size: int = 11
    
    # Year inference thresholds
    year_inference_full_threshold: int = 88
    year_inference_masked_threshold: int = 85
    
    # Benchmark settings
    auto_find_benchmark: bool = True
    manual_benchmark_year: Optional[int] = None

    # Backup retention
    backup_retention_days: int = 90

    # Log retention
    log_retention_days: int = 90

    # Sorting behavior for Update/Inject mode
    enable_sort_existing: bool = True

    # Configurable label for the "Uncategorized" subcategory
    uncategorized_label: str = "Uncategorized"

    # Offline guard (default ON)
    offline_guard_enabled: bool = True

    # Hanging indentation (inches) for study paragraphs
    hanging_indent_inches: float = 0.5

    # Optional update checker (disabled by default)
    check_updates_on_startup: bool = False

    def __post_init__(self):
        if not self.data_root:
            self.data_root = str(get_default_data_root())
        self._validate()
    
    def _validate(self) -> None:
        """Validate config values. Raises ValueError on invalid config."""
        _errors = []
        if not isinstance(self.fuzzy_threshold_full, int) or not (0 <= self.fuzzy_threshold_full <= 100):
            _errors.append(f"fuzzy_threshold_full must be int 0-100, got {self.fuzzy_threshold_full!r}")
        if not isinstance(self.fuzzy_threshold_masked, int) or not (0 <= self.fuzzy_threshold_masked <= 100):
            _errors.append(f"fuzzy_threshold_masked must be int 0-100, got {self.fuzzy_threshold_masked!r}")
        if not isinstance(self.benchmark_min_count, int) or self.benchmark_min_count < 1:
            _errors.append(f"benchmark_min_count must be int >= 1, got {self.benchmark_min_count!r}")
        if not isinstance(self.font_size, int) or self.font_size < 1:
            _errors.append(f"font_size must be int >= 1, got {self.font_size!r}")
        if not isinstance(self.font_name, str) or self.font_name not in ALLOWED_FONTS:
            _errors.append(
                f"font_name must be one of {ALLOWED_FONTS}, got {self.font_name!r}"
            )
        if not isinstance(self.backup_retention_days, int) or self.backup_retention_days < 1:
            _errors.append(f"backup_retention_days must be int >= 1, got {self.backup_retention_days!r}")
        if not isinstance(self.log_retention_days, int) or self.log_retention_days < 1:
            _errors.append(f"log_retention_days must be int >= 1, got {self.log_retention_days!r}")
        if not isinstance(self.enable_sort_existing, bool):
            _errors.append(f"enable_sort_existing must be bool, got {self.enable_sort_existing!r}")
        if not isinstance(self.uncategorized_label, str) or not self.uncategorized_label.strip():
            _errors.append(f"uncategorized_label must be a non-empty string, got {self.uncategorized_label!r}")
        if not isinstance(self.hanging_indent_inches, (int, float)):
            _errors.append(f"hanging_indent_inches must be a number, got {self.hanging_indent_inches!r}")
        elif not (HANGING_INDENT_MIN <= float(self.hanging_indent_inches) <= HANGING_INDENT_MAX):
            _errors.append(
                f"hanging_indent_inches must be {HANGING_INDENT_MIN}–{HANGING_INDENT_MAX}, "
                f"got {self.hanging_indent_inches!r}"
            )
        if not isinstance(self.check_updates_on_startup, bool):
            _errors.append(f"check_updates_on_startup must be bool, got {self.check_updates_on_startup!r}")
        if self.manual_benchmark_year is not None:
            if not isinstance(self.manual_benchmark_year, int) or not (1900 <= self.manual_benchmark_year <= 2100):
                _errors.append(f"manual_benchmark_year must be int 1900-2100 or None, got {self.manual_benchmark_year!r}")
        if _errors:
            raise ValueError(
                "Invalid configuration:\n" + "\n".join(f"  - {e}" for e in _errors)
            )

    @property
    def data_path(self) -> Path:
        return Path(self.data_root)
    
    def get_user_id(self) -> str:
        """Get the current user ID based on strategy."""
        if self.user_id_strategy == "os_username":
            return get_os_username()
        return get_os_username()  # Default fallback
    
    def get_user_data_path(self, user_id: Optional[str] = None) -> Path:
        """Get the data path for a specific user."""
        uid = user_id or self.get_user_id()
        return self.data_path / "users" / uid
    
    def get_user_db_path(self, user_id: Optional[str] = None) -> Path:
        """Get the SQLite database path for a user."""
        return self.get_user_data_path(user_id) / "sites.db"
    
    def get_user_exports_path(self, user_id: Optional[str] = None) -> Path:
        """Get the exports directory for a user."""
        return self.get_user_data_path(user_id) / "exports"
    
    def get_user_imports_path(self, user_id: Optional[str] = None) -> Path:
        """Get the imports directory for a user."""
        return self.get_user_data_path(user_id) / "imports"
    
    def get_user_backups_path(self, user_id: Optional[str] = None) -> Path:
        """Get the backups directory for a user."""
        return self.get_user_data_path(user_id) / "backups"
    
    def get_user_logs_path(self, user_id: Optional[str] = None) -> Path:
        """Get the logs directory for a user."""
        return self.get_user_data_path(user_id) / "logs"
    
    def get_user_results_path(self, user_id: Optional[str] = None) -> Path:
        """Get the results directory for a user."""
        return self.get_user_data_path(user_id) / "results"

    def get_result_root(self) -> Path:
        """Get the root directory for output result files (.docx, .xlsx exports).

        In production (default data_root), this returns ``<project_root>/result/``.
        When ``data_root`` is overridden (e.g. in tests), returns ``<data_root>/result/``
        so tests never pollute the real result folder.
        """
        default_root = str(get_default_data_root())
        if self.data_root == default_root:
            return get_app_root() / "result"
        return Path(self.data_root) / "result"
    
    def get_temp_path(self) -> Path:
        """Get the temporary files directory."""
        return self.data_path / "tmp"
    
    def ensure_user_directories(self, user_id: Optional[str] = None) -> None:
        """Create all necessary user directories with restrictive permissions."""
        paths = [
            self.get_user_data_path(user_id),
            self.get_user_exports_path(user_id),
            self.get_user_imports_path(user_id),
            self.get_user_backups_path(user_id),
            self.get_user_logs_path(user_id),
            self.get_user_results_path(user_id),
            self.get_temp_path(),
        ]
        for path in paths:
            path.mkdir(parents=True, exist_ok=True)
            # On Windows, permissions are handled differently
            # On Unix, we would use os.chmod(path, 0o700)
    
    def to_dict(self) -> dict:
        """Convert config to dictionary."""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: dict) -> "AppConfig":
        """Create config from dictionary."""
        # Filter only known fields
        known_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered = {k: v for k, v in data.items() if k in known_fields}
        return cls(**filtered)
    
    def save(self, path: Optional[Path] = None) -> None:
        """Save configuration to JSON file."""
        if path is None:
            path = Path(self.data_root) / "config.json"
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, indent=2)
    
    @classmethod
    def load(cls, path: Optional[Path] = None) -> "AppConfig":
        """Load configuration from JSON file, or return defaults."""
        if path is None:
            path = get_default_data_root() / "config.json"
        
        if path.exists():
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                config = cls.from_dict(data)
                # Ensure network is always disabled
                config.network_enabled = False
                return config
            except (json.JSONDecodeError, IOError):
                pass
        
        return cls()


# Global config instance
_config: Optional[AppConfig] = None


def get_config() -> AppConfig:
    """Get the global configuration instance."""
    global _config
    if _config is None:
        _config = AppConfig.load()
    return _config


def set_config(config: AppConfig) -> None:
    """Set the global configuration instance."""
    global _config
    _config = config
