"""
Data models for the CV Research Experience Manager.
"""

import logging
from dataclasses import dataclass, field
from typing import Optional, List
from datetime import datetime

from normalizer import normalize_heading_key, normalize_subcat_key, is_uncategorized_key


@dataclass
class Study:
    """Represents a single study/research entry."""
    
    phase: str
    subcategory: str
    year: int
    sponsor: str
    protocol: str  # May be empty
    description_full: str
    description_masked: str
    
    # Database fields
    id: Optional[int] = None
    site_id: Optional[int] = None
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    
    def get_identity_tuple(self, normalized_masked: str = "") -> tuple:
        """
        Get the identity tuple for deduplication.
        Phase and Subcategory are normalized so that casing, spacing, and
        Roman-numeral variants are treated as identical.
        (PhaseKey, SubcatKey, Year, Sponsor, Protocol?, DescriptionMaskedNormalized)
        """
        masked = normalized_masked or self.description_masked
        return (
            normalize_heading_key(self.phase),
            normalize_subcat_key(self.subcategory),
            self.year,
            self.sponsor,
            self.protocol or "",
            masked
        )
    
    def format_for_cv(self, include_protocol: bool = True) -> str:
        """
        Format the study for CV display.
        Format: {Year}<TAB>{Sponsor}{[ SPACE ]{Protocol}}: {Description}
        """
        if include_protocol and self.protocol:
            return f"{self.year}\t{self.sponsor} {self.protocol}: {self.description_full}"
        else:
            return f"{self.year}\t{self.sponsor}: {self.description_masked}"
    
    def __hash__(self):
        return hash(self.get_identity_tuple())
    
    def __eq__(self, other):
        if not isinstance(other, Study):
            return False
        return self.get_identity_tuple() == other.get_identity_tuple()


@dataclass
class Subcategory:
    """Represents a subcategory within a phase."""
    
    name: str
    studies: List[Study] = field(default_factory=list)
    
    def sort_studies(self):
        """Sort studies: Year desc, then Sponsor, then Protocol."""
        self.studies.sort(key=lambda s: (-s.year, s.sponsor.lower(), s.protocol.lower()))


@dataclass
class Phase:
    """Represents a phase (Phase I or Phase II-IV)."""
    
    name: str
    subcategories: List[Subcategory] = field(default_factory=list)
    
    def sort_subcategories(self):
        """Sort subcategories alphabetically and sort studies within each."""
        self.subcategories.sort(key=lambda sc: sc.name.lower())
        for sc in self.subcategories:
            sc.sort_studies()
    
    def get_or_create_subcategory(self, name: str) -> Subcategory:
        """Get existing subcategory or create new one.

        Matching uses normalize_subcat_key (NFC, casefold, whitespace
        collapse, dash/quote canonicalization).  The *first* variant
        seen becomes the persisted display name.
        """
        key = normalize_subcat_key(name)
        for sc in self.subcategories:
            if normalize_subcat_key(sc.name) == key:
                logging.debug(
                    "[Models] Matched existing subcategory '%s' "
                    "(key='%s') for requested '%s'",
                    sc.name,
                    key,
                    name,
                )
                return sc
        new_sc = Subcategory(name=name)
        self.subcategories.append(new_sc)
        logging.debug(
            "[Models] Created new subcategory '%s' (key='%s')",
            name,
            key,
        )
        return new_sc


@dataclass
class ResearchExperience:
    """Represents the entire Research Experience section."""
    
    phases: List[Phase] = field(default_factory=list)
    
    def get_phase_order_key(self, phase_name: str) -> int:
        """Get sort key for phase ordering (Phase I first, then Phase II-IV, Uncategorized last)."""
        key = normalize_heading_key(phase_name)
        if is_uncategorized_key(phase_name):
            return 99
        if key == "Phase I":
            return 0
        return 1
    
    def sort_all(self):
        """Sort phases, subcategories, and studies."""
        self.phases.sort(key=lambda p: self.get_phase_order_key(p.name))
        for phase in self.phases:
            phase.sort_subcategories()
    
    def sort_all_custom(self, custom_order: List[str]):
        """Sort phases and subcategories according to custom order.
        
        Args:
            custom_order: List of "Phase > Subcategory" strings in desired order
        """
        # Build order lookup: "Phase > Subcategory" -> index
        order_lookup = {key: idx for idx, key in enumerate(custom_order)}
        default_order = len(custom_order)  # Items not in list go last
        
        # Collect all phase/subcategory combinations with their studies
        all_combos = []
        for phase in self.phases:
            for subcat in phase.subcategories:
                key = f"{phase.name} > {subcat.name}"
                order_idx = order_lookup.get(key, default_order)
                # Keep all studies - don't filter any out
                all_combos.append((order_idx, phase.name, subcat.name, subcat.studies[:]))  # Copy list
        
        # Sort by custom order
        all_combos.sort(key=lambda x: (x[0], x[1], x[2]))
        
        # Rebuild phases structure
        self.phases = []
        for _, phase_name, subcat_name, studies in all_combos:
            # Skip empty subcategories
            if not studies:
                continue
                
            phase = self.get_or_create_phase(phase_name)
            subcat = phase.get_or_create_subcategory(subcat_name)
            subcat.studies = studies
            # Sort studies within subcategory by year descending
            subcat.studies.sort(key=lambda s: (-s.year, s.sponsor, s.protocol))
    
    def get_or_create_phase(self, name: str) -> Phase:
        """Get existing phase or create new one.

        Matching uses normalize_heading_key (NFC, casefold, whitespace
        collapse, dash/quote canonicalization, Roman-numeral equivalence
        via PHASE_SYNONYMS).  The *first* variant seen becomes the
        persisted display name.
        """
        key = normalize_heading_key(name)
        for p in self.phases:
            if normalize_heading_key(p.name) == key:
                logging.debug(
                    "[Models] Matched existing phase '%s' "
                    "(key='%s') for requested '%s'",
                    p.name,
                    key,
                    name,
                )
                return p
        new_phase = Phase(name=name)
        self.phases.append(new_phase)
        logging.debug(
            "[Models] Created new phase '%s' (key='%s')",
            name,
            key,
        )
        return new_phase
    
    def get_all_studies(self) -> List[Study]:
        """Get all studies across all phases and subcategories."""
        studies = []
        for phase in self.phases:
            for subcat in phase.subcategories:
                studies.extend(subcat.studies)
        return studies
    
    def get_all_years(self) -> List[int]:
        """Get all unique years across all studies."""
        return list(set(s.year for s in self.get_all_studies()))
    
    def calculate_benchmark_year(self, min_count: int = 4) -> int:
        """
        Calculate the benchmark year for injection.
        Latest year, but if count <= 3, use latest - 1.
        """
        studies = self.get_all_studies()
        if not studies:
            return datetime.now().year
        
        years = [s.year for s in studies]
        latest = max(years)
        count_in_latest = sum(1 for y in years if y == latest)
        
        if count_in_latest <= (min_count - 1):  # ≤3 when min_count=4
            return latest - 1
        return latest


@dataclass
class Site:
    """Represents a site database."""
    
    id: Optional[int] = None
    owner_user_id: str = ""
    name: str = ""
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    studies: List[Study] = field(default_factory=list)


@dataclass
class SiteVersion:
    """Represents a version/snapshot of a site."""
    
    id: Optional[int] = None
    site_id: int = 0
    created_at: Optional[datetime] = None
    note: str = ""


@dataclass
class LogEntry:
    """Represents a log entry for operations."""
    
    timestamp: datetime
    operation: str  # inserted, matched-existing, skipped-duplicate, no-changes, replaced, skipped-no-match, ambiguous-below-threshold
    phase: str
    subcategory: str
    year: int
    sponsor: str
    protocol: str
    details: str
    
    def to_dict(self) -> dict:
        return {
            "timestamp": self.timestamp.isoformat(),
            "operation": self.operation,
            "phase": self.phase,
            "subcategory": self.subcategory,
            "year": self.year,
            "sponsor": self.sponsor,
            "protocol": self.protocol,
            "details": self.details,
        }


@dataclass
class OperationResult:
    """Result of an update/redact operation."""
    
    success: bool
    output_path: Optional[str] = None
    log_entries: List[LogEntry] = field(default_factory=list)
    summary: dict = field(default_factory=dict)
    error_message: str = ""
    
    def get_counts(self) -> dict:
        """Get counts by operation type."""
        counts = {}
        for entry in self.log_entries:
            op = entry.operation
            counts[op] = counts.get(op, 0) + 1
        return counts
