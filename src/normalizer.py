"""
Text normalization and matching utilities for the CV Research Experience Manager.
"""

import re
import unicodedata
from typing import Optional, Tuple
from rapidfuzz import fuzz


# Phase normalization patterns
PHASE_PATTERNS = [
    (r'\bPhase\s*1\b', 'Phase I'),
    (r'\bPhase\s*I\b(?!\s*[IV-])', 'Phase I'),
    (r'\bPhase\s*2\b', 'Phase II'),
    (r'\bPhase\s*II\b(?!\s*[IV-])', 'Phase II'),
    (r'\bPhase\s*3\b', 'Phase III'),
    (r'\bPhase\s*III\b', 'Phase III'),
    (r'\bPhase\s*4\b', 'Phase IV'),
    (r'\bPhase\s*IV\b', 'Phase IV'),
    (r'\bPhase\s*II\s*[-–—]\s*IV\b', 'Phase II–IV'),
    (r'\bPhase\s*2\s*[-–—]\s*4\b', 'Phase II–IV'),
]

# PHASE_SYNONYMS — maps every known casefolded variant to the canonical name.
# Used exclusively for matching keys; persisted text is never altered.
# Add new synonyms here if additional phase names are introduced.
PHASE_SYNONYMS = {
    # Phase I
    "phase i": "Phase I",
    "phase 1": "Phase I",
    "phase i.": "Phase I",
    "phase1": "Phase I",
    # Phase II
    "phase ii": "Phase II",
    "phase 2": "Phase II",
    "phase ii.": "Phase II",
    "phase2": "Phase II",
    # Phase III
    "phase iii": "Phase III",
    "phase 3": "Phase III",
    "phase iii.": "Phase III",
    "phase3": "Phase III",
    # Phase IV
    "phase iv": "Phase IV",
    "phase 4": "Phase IV",
    "phase iv.": "Phase IV",
    "phase4": "Phase IV",
    # Phase II-IV (various dash forms collapse to en-dash canonical)
    "phase ii-iv": "Phase II\u2013IV",
    "phase ii\u2013iv": "Phase II\u2013IV",
    "phase ii\u2014iv": "Phase II\u2013IV",
    "phase 2-4": "Phase II\u2013IV",
    "phase 2\u20134": "Phase II\u2013IV",
    "phase 2\u20144": "Phase II\u2013IV",
    # Uncategorized
    "uncategorized": "Uncategorized",
}

def normalize_heading_key(text: str) -> str:
    """Normalize a heading key for matching purposes."""
    text = unicodedata.normalize('NFC', text)
    text = text.casefold()
    text = re.sub(r'[\s\t]+', ' ', text).strip()
    text = re.sub(r'[–—−‐‑‒―]', '-', text)
    text = re.sub(r'["\u2018\u2019\u201a\u201b]', "'", text)
    text = re.sub(r'["\u201c\u201d\u201e\u201f]', '"', text)
    return PHASE_SYNONYMS.get(text, text)

def normalize_subcat_key(text: str) -> str:
    """Normalize a subcategory key for matching purposes."""
    text = unicodedata.normalize('NFC', text)
    text = text.casefold()
    text = re.sub(r'[\s\t]+', ' ', text).strip()
    text = re.sub(r'[–—−‐‑‒―]', '-', text)
    text = re.sub(r'["\u2018\u2019\u201a\u201b]', "'", text)
    text = re.sub(r'["\u201c\u201d\u201e\u201f]', '"', text)
    return PHASE_SYNONYMS.get(text, text)


def is_uncategorized_key(text: str) -> bool:
    """Return True if *text* normalizes to the canonical 'Uncategorized' key.

    This matches both the hard-coded default and any configured label whose
    casefold equals 'uncategorized'.
    """
    key = normalize_heading_key(text)
    return key == "Uncategorized" or text.strip().casefold() == "uncategorized"

# Protocol detection regex
PROTOCOL_REGEX = re.compile(r'([A-Za-z]{1,10}-?\d[\w-]*)')

# All-caps alphanumeric pattern for protocol detection
ALLCAPS_PROTOCOL_REGEX = re.compile(r'\b([A-Z]{2,}[A-Z0-9-]*\d+[A-Z0-9-]*)\b')

# Numeric-only protocol pattern (e.g., "12345", "2023-001")
NUMERIC_PROTOCOL_REGEX = re.compile(r'\b(\d{3,}(?:-\d+)?)\b')


def normalize_whitespace(text: str) -> str:
    """Normalize whitespace: collapse multiple spaces/tabs to single space."""
    return re.sub(r'[\s\t]+', ' ', text).strip()


def normalize_dashes(text: str) -> str:
    """Unify all dash variants to standard hyphen."""
    return re.sub(r'[–—−‐‑‒―]', '-', text)


def normalize_quotes(text: str) -> str:
    """Normalize curly quotes to straight quotes."""
    text = re.sub('[\u2018\u2019\u201a\u201b]', "'", text)
    text = re.sub('[\u201c\u201d\u201e\u201f]', '"', text)
    return text


def normalize_colon_spacing(text: str) -> str:
    """Canonicalize colon spacing: ensure single space after colon."""
    return re.sub(r'\s*:\s*', ': ', text)


def normalize_phase(text: str) -> str:
    """Normalize phase names (e.g., 'Phase 1' -> 'Phase I')."""
    result = text
    for pattern, replacement in PHASE_PATTERNS:
        result = re.sub(pattern, replacement, result, flags=re.IGNORECASE)
    return result


def collapse_x_runs(text: str) -> str:
    """Collapse runs of X to XXX for matching purposes."""
    return re.sub(r'X{2,}', 'XXX', text)


def normalize_for_matching(text: str) -> str:
    """
    Full normalization for matching purposes:
    - lowercase
    - normalize whitespace/tabs
    - unify dashes
    - normalize quotes
    - canonicalize colon spacing
    - normalize phases
    - collapse X runs to XXX
    """
    result = unicodedata.normalize('NFC', text)
    result = result.lower()
    result = normalize_whitespace(result)
    result = normalize_dashes(result)
    result = normalize_quotes(result)
    result = normalize_colon_spacing(result)
    result = normalize_phase(result)
    result = collapse_x_runs(result)
    return result


def normalize_for_display(text: str) -> str:
    """
    Normalization for display (preserves case and X runs):
    - normalize whitespace
    - unify dashes
    - normalize quotes
    - canonicalize colon spacing
    - normalize phases
    """
    result = unicodedata.normalize('NFC', text)
    result = normalize_whitespace(result)
    result = normalize_dashes(result)
    result = normalize_quotes(result)
    result = normalize_colon_spacing(result)
    result = normalize_phase(result)
    return result


def extract_protocol(text: str) -> Optional[str]:
    """
    Extract protocol from text.
    Returns the protocol if found, None otherwise.
    """
    # Try the standard protocol pattern first
    match = PROTOCOL_REGEX.search(text)
    if match:
        potential = match.group(1)
        # Ensure it's not just a sponsor with numbers (e.g., "23andMe")
        # Protocol should have letter-number pattern, not number-letter at start
        if not re.match(r'^\d+[a-zA-Z]', potential):
            return potential
    
    # Try all-caps pattern
    match = ALLCAPS_PROTOCOL_REGEX.search(text)
    if match:
        return match.group(1)
    
    # Try numeric-only pattern (at least 3 digits)
    # This should be last to avoid false positives with years
    match = NUMERIC_PROTOCOL_REGEX.search(text)
    if match:
        potential = match.group(1)
        # Avoid matching years (1900-2099)
        if not (1900 <= int(potential.split('-')[0]) <= 2099):
            return potential
    
    return None


def is_protocol_like(text: str) -> bool:
    """Check if text contains a protocol-like token."""
    return extract_protocol(text) is not None


def parse_sponsor_protocol(text: str) -> Tuple[str, str]:
    """
    Parse sponsor and protocol from the left side of a study line.
    
    Input: "Sponsor Protocol" or "Sponsor" (no protocol)
    Returns: (sponsor, protocol)
    """
    text = text.strip()
    
    # Look for protocol pattern
    protocol = extract_protocol(text)
    
    if protocol:
        # Find where protocol starts and extract sponsor
        idx = text.find(protocol)
        if idx > 0:
            sponsor = text[:idx].strip()
            return sponsor, protocol
        else:
            # Protocol at start - unusual, treat whole thing as sponsor
            return text, ""
    
    # No protocol found
    return text, ""

def parse_study_line(line: str) -> Optional[Tuple[int, str, str, str]]:
    """
    Parse a study line in format: {Year}<TAB>{Sponsor}{[ SPACE ]{Protocol}}: {Description}
    
    Returns: (year, sponsor, protocol, description) or None if not a valid study line.
    """
    line = line.strip()
    if not line:
        return None
    
    # Try to find year at the start (with or without tab)
    year_match = re.match(r'^(\d{4})[\t\s]+', line)
    if not year_match:
        return None
    
    year = int(year_match.group(1))
    rest = line[year_match.end():]
    
    # Find the colon or semicolon that separates sponsor/protocol from description
    # Some CVs use semicolon instead of colon
    colon_idx = rest.find(':')
    semicolon_idx = rest.find(';')
    
    # Use whichever delimiter comes first (if both exist)
    if colon_idx == -1 and semicolon_idx == -1:
        # No delimiter, might still be valid - treat rest as sponsor
        sponsor, protocol = parse_sponsor_protocol(rest)
        return year, sponsor, protocol, ""
    elif colon_idx == -1:
        delimiter_idx = semicolon_idx
    elif semicolon_idx == -1:
        delimiter_idx = colon_idx
    else:
        delimiter_idx = min(colon_idx, semicolon_idx)
    
    sponsor_protocol_part = rest[:delimiter_idx].strip()
    description = rest[delimiter_idx + 1:].strip()
    
    # Strip role label from description if present
    description = strip_role_label(description)
    
    sponsor, protocol = parse_sponsor_protocol(sponsor_protocol_part)
    
    return year, sponsor, protocol, description

def fuzzy_match(text1: str, text2: str, threshold: int = 90) -> Tuple[bool, int]:
    """
    Perform fuzzy matching between two texts.
    
    Returns: (is_match, score)
    """
    norm1 = normalize_for_matching(text1)
    norm2 = normalize_for_matching(text2)
    
    # Try exact match first
    if norm1 == norm2:
        return True, 100
    
    # Fuzzy match
    score = fuzz.ratio(norm1, norm2)
    return score >= threshold, score


def exact_match(text1: str, text2: str) -> bool:
    """Check for exact match after normalization."""
    return normalize_for_matching(text1) == normalize_for_matching(text2)


def match_study_to_master(
    cv_year: int,
    cv_sponsor: str,
    cv_protocol: str,
    cv_description: str,
    master_studies: list,
    fuzzy_threshold_full: int = 92,
    fuzzy_threshold_masked: int = 90,
) -> Optional[Tuple[object, str, int]]:
    """
    Match a CV study line to a master study.
    
    Returns: (matched_study, match_type, score) or None
    match_type: 'exact_full', 'fuzzy_full', 'exact_masked', 'fuzzy_masked'
    """
    from models import Study
    
    # Filter by year first (allow 0 for unknown years)
    if cv_year > 0:
        candidates = [s for s in master_studies if s.year == cv_year]
    else:
        candidates = list(master_studies)
    
    if not candidates:
        return None
    
    # Build the full CV text for matching (normalize sponsor case)
    cv_sponsor_norm = cv_sponsor.upper()
    if cv_protocol:
        cv_full_text = f"{cv_sponsor} {cv_protocol}: {cv_description}"
    else:
        cv_full_text = f"{cv_sponsor}: {cv_description}"
    
    # Strategy 1: Try matching against full descriptions (Column B)
    # Try exact match first
    for study in candidates:
        if study.sponsor.upper() != cv_sponsor_norm:
            continue
        master_full = f"{study.sponsor} {study.protocol}: {study.description_full}" if study.protocol else f"{study.sponsor}: {study.description_full}"
        if exact_match(cv_full_text, master_full):
            return study, 'exact_full', 100
    
    # Fuzzy match against full
    best_match = None
    best_score = 0
    best_type = None
    
    for study in candidates:
        # Sponsor fuzzy match
        sponsor_match = study.sponsor.upper() == cv_sponsor_norm or fuzz.ratio(study.sponsor.lower(), cv_sponsor.lower()) >= 85
        if not sponsor_match:
            continue
            
        master_full = f"{study.sponsor} {study.protocol}: {study.description_full}" if study.protocol else f"{study.sponsor}: {study.description_full}"
        is_match, score = fuzzy_match(cv_full_text, master_full, fuzzy_threshold_full)
        if is_match and score > best_score:
            best_match = study
            best_score = score
            best_type = 'fuzzy_full'
    
    # Strategy 2: Try matching against masked descriptions (Column C)
    # This is important for CVs that are already masked with XXXX
    for study in candidates:
        if study.sponsor.upper() != cv_sponsor_norm:
            continue
        master_masked = f"{study.sponsor}: {study.description_masked}"
        if exact_match(cv_full_text, master_masked):
            return study, 'exact_masked', 100
    
    # Fuzzy match against masked
    for study in candidates:
        # Sponsor fuzzy match
        sponsor_match = study.sponsor.upper() == cv_sponsor_norm or fuzz.ratio(study.sponsor.lower(), cv_sponsor.lower()) >= 85
        if not sponsor_match:
            continue
            
        master_masked = f"{study.sponsor}: {study.description_masked}"
        is_match, score = fuzzy_match(cv_full_text, master_masked, fuzzy_threshold_masked)
        if is_match and score > best_score:
            best_match = study
            best_score = score
            best_type = 'fuzzy_masked'
    
    # Strategy 3: Match by description alone (ignoring sponsor mismatch)
    # Useful when sponsor name formatting differs (e.g., "BRISTOL-MYERS SQUIBB" vs "Bristol-Myers Squibb")
    if not best_match:
        for study in candidates:
            # Try matching just the description parts
            cv_desc_norm = normalize_for_matching(cv_description)
            master_masked_norm = normalize_for_matching(study.description_masked)
            master_full_norm = normalize_for_matching(study.description_full)
            
            # Check masked description
            is_match, score = fuzzy_match(cv_description, study.description_masked, fuzzy_threshold_masked - 5)
            if is_match and score > best_score:
                best_match = study
                best_score = score
                best_type = 'fuzzy_masked'
            
            # Check full description
            is_match, score = fuzzy_match(cv_description, study.description_full, fuzzy_threshold_full - 5)
            if is_match and score > best_score:
                best_match = study
                best_score = score
                best_type = 'fuzzy_full'
    
    if best_match:
        return best_match, best_type, best_score
    
    return None


def infer_year_from_master(
    cv_sponsor: str,
    cv_description: str,
    master_studies: list,
    heading_year_bound: Optional[int] = None,
    full_match_min_score: int = 88,
    masked_match_min_score: int = 85,
) -> Tuple[Optional[int], Optional[object], str]:
    """
    Infer missing year for a CV study by matching to master database.
    
    Args:
        cv_sponsor: Sponsor name from CV
        cv_description: Description from CV
        master_studies: List of master studies
        heading_year_bound: Upper bound from heading like "Pre 2022" (exclusive)
        full_match_min_score: Minimum score for Column B matching
        masked_match_min_score: Minimum score for Column C matching
    
    Returns: (inferred_year, matched_study, reason)
        - inferred_year: The inferred year or None
        - matched_study: The matched Study object or None
        - reason: String explaining the inference or why it failed
    """
    from models import Study
    
    # Normalize sponsor for matching
    cv_sponsor_norm = cv_sponsor.upper()
    
    # Filter candidates by sponsor
    candidates = []
    for study in master_studies:
        sponsor_match = (
            study.sponsor.upper() == cv_sponsor_norm or 
            fuzz.ratio(study.sponsor.lower(), cv_sponsor.lower()) >= 85
        )
        if sponsor_match:
            candidates.append(study)
    
    if not candidates:
        return None, None, f"No master studies found for sponsor '{cv_sponsor}'"
    
    # Two-pass matching: exact then fuzzy
    best_match = None
    best_score = 0
    best_type = None
    
    # Pass 1: Exact matches
    for study in candidates:
        # Try full description (Column B)
        master_full = f"{study.sponsor} {study.protocol}: {study.description_full}" if study.protocol else f"{study.sponsor}: {study.description_full}"
        cv_full = f"{cv_sponsor}: {cv_description}"
        
        if exact_match(cv_full, master_full):
            # Check heading bound
            if heading_year_bound is None or study.year < heading_year_bound:
                return study.year, study, f"Exact match (full) with year {study.year}"
        
        # Try masked description (Column C)
        master_masked = f"{study.sponsor}: {study.description_masked}"
        if exact_match(cv_full, master_masked):
            if heading_year_bound is None or study.year < heading_year_bound:
                return study.year, study, f"Exact match (masked) with year {study.year}"
    
    # Pass 2: Fuzzy matches
    for study in candidates:
        # Try full description (Column B)
        master_full = f"{study.sponsor} {study.protocol}: {study.description_full}" if study.protocol else f"{study.sponsor}: {study.description_full}"
        cv_full = f"{cv_sponsor}: {cv_description}"
        
        is_match, score = fuzzy_match(cv_full, master_full, full_match_min_score)
        if is_match and score > best_score:
            # Check heading bound
            if heading_year_bound is None or study.year < heading_year_bound:
                best_match = study
                best_score = score
                best_type = 'fuzzy_full'
        
        # Try masked description (Column C)
        master_masked = f"{study.sponsor}: {study.description_masked}"
        is_match, score = fuzzy_match(cv_full, master_masked, masked_match_min_score)
        if is_match and score > best_score:
            if heading_year_bound is None or study.year < heading_year_bound:
                best_match = study
                best_score = score
                best_type = 'fuzzy_masked'
    
    if best_match:
        return best_match.year, best_match, f"Fuzzy match ({best_type}, score={best_score}) with year {best_match.year}"
    
    # No match above threshold
    bound_msg = f" (within bound < {heading_year_bound})" if heading_year_bound else ""
    return None, None, f"No match above threshold{bound_msg} - ambiguous-old-format"


def strip_role_label(description: str) -> str:
    """
    Remove leading role label from description.
    
    Example:
        "Research Assistant, A Phase 2 study..." -> "A Phase 2 study..."
        "Laboratory Technician, A randomized..." -> "A randomized..."
    
    Only removes the first comma-delimited token after the colon.
    """
    # Common role patterns at clinical trial companies
    role_patterns = [
        r'^Research\s+Assistant,\s*',
        r'^Laboratory\s+Technician\s+I+,\s*',  # Includes I, II, III, etc.
        r'^Laboratory\s+Technician,\s*',
        r'^Laboratory\s+Manager,\s*',
        r'^Lab\s+Manager,\s*',
        r'^Lab\s+Technician,\s*',
        r'^Study\s+Coordinator,\s*',
        r'^Clinical\s+Research\s+Coordinator,\s*',
        r'^Clinical\s+Research\s+Associate,\s*',
        r'^Research\s+Coordinator,\s*',
        r'^Research\s+Associate,\s*',
        r'^Research\s+Scientist,\s*',
        r'^Senior\s+Research\s+Assistant,\s*',
        r'^Senior\s+Laboratory\s+Technician,\s*',
        r'^Project\s+Manager,\s*',
        r'^Clinical\s+Trial\s+Manager,\s*',
    ]
    
    result = description.strip()
    
    # Try each pattern
    for pattern in role_patterns:
        result = re.sub(pattern, '', result, flags=re.IGNORECASE)
    
    return result.strip()

def validate_year(year_str: str) -> Optional[int]:
    """Validate and parse a year string. Returns int or None if invalid."""
    try:
        year = int(year_str)
        if 1900 <= year <= 2100:
            return year
    except (ValueError, TypeError):
        pass
    return None


def is_phase_heading(text: str) -> Optional[str]:
    """
    Check if text is a phase heading and return normalized phase name.
    Returns None if not a phase heading.
    
    Uses normalize_heading_key for robust matching against PHASE_SYNONYMS,
    then falls back to regex-based detection for partial matches.
    """
    text = text.strip()
    if not text:
        return None

    # Fast path: exact synonym lookup via normalize_heading_key
    key = normalize_heading_key(text)
    if key in (
        "Phase I", "Phase II", "Phase III", "Phase IV",
        "Phase II\u2013IV", "Uncategorized",
    ):
        return key

    # Fallback: apply normalize_phase regex patterns then check
    normalized = normalize_phase(text)
    if re.match(r'^Phase\s+I(?:\s|$)', normalized, re.IGNORECASE):
        return "Phase I"
    if re.match(r'^Phase\s+II[-\u2013]IV', normalized, re.IGNORECASE):
        return "Phase II\u2013IV"
    if re.match(r'^Phase\s+II(?:\s|$)', normalized, re.IGNORECASE):
        return "Phase II"
    if re.match(r'^Phase\s+III(?:\s|$)', normalized, re.IGNORECASE):
        return "Phase III"
    if re.match(r'^Phase\s+IV(?:\s|$)', normalized, re.IGNORECASE):
        return "Phase IV"

    if text.lower().strip() == "uncategorized":
        return "Uncategorized"

    return None


SPONSOR_PROTOCOL_RE = re.compile(
    r'^(?P<sponsor>[A-Z][A-Za-z&\s\-\.]+?)\s+'
    r'(?P<protocol>[A-Za-z]{1,10}[\-]?\d[\w\-]*)',
    re.UNICODE,
)


def contains_protocol_token(text: str) -> bool:
    """Return True if *text* contains a sponsor+protocol token.

    Uses combined heuristics:
      1. NFC + casefold + whitespace collapse + dash canonicalization.
      2. ``extract_protocol`` for the protocol part.
      3. Cross-check that a sponsor-like prefix precedes the protocol.

    Does NOT rely on font colour or bold style.
    """
    normed = unicodedata.normalize('NFC', text)
    normed = re.sub(r'[\s\t]+', ' ', normed).strip()
    normed = re.sub(r'[–—−‐‑‒―]', '-', normed)
    normed = re.sub('[\u2018\u2019\u201a\u201b]', "'", normed)
    normed = re.sub('[\u201c\u201d\u201e\u201f]', '"', normed)

    proto = extract_protocol(normed)
    if proto is None:
        return False

    proto_idx = normed.find(proto)
    if proto_idx <= 0:
        return False

    prefix = normed[:proto_idx].strip()
    if len(prefix) < 2:
        return False

    return True


def is_already_masked(text: str) -> bool:
    """Return True if *text* looks like an already-masked study line.

    A masked line uses XXX (or multiple Xs) in place of the protocol
    and treatment names, and does NOT contain a protocol token.
    """
    if 'XXX' not in text.upper() and 'xxx' not in text.lower():
        return False
    return not contains_protocol_token(text)


def is_year_line(text: str) -> bool:
    """Check if a line starts with a 4-digit year."""
    return bool(re.match(r'^\d{4}[\t\s]', text.strip()))
