"""
Core processing logic for CV modification modes.
Mode A: Update/Inject
Mode B: Redact Protocols
"""

import re
import logging
from pathlib import Path
from typing import List, Optional, Tuple, Set
from datetime import datetime
import shutil

from models import Study, ResearchExperience, OperationResult, LogEntry
from normalizer import (
    normalize_for_matching, match_study_to_master, normalize_phase,
    collapse_x_runs, contains_protocol_token, is_already_masked,
    normalize_heading_key, normalize_subcat_key, is_uncategorized_key,
)
from docx_handler import CVDocxHandler, validate_cv_docx
from excel_parser import parse_master_xlsx, validate_master_xlsx, studies_to_research_experience
from database import DatabaseManager
from logger import OperationLogger
from config import get_config, AppConfig
from error_handler import FilePermissionError

# Regex to strip date-stamped result suffixes from CV filenames
_RESULT_SUFFIX_RE = re.compile(
    r'\s*\((?:Updated|Redacted)\s+\d{4}-\d{2}-\d{2}\)\s*$'
)

class CVProcessor:
    """Main processor for CV modification operations."""
    
    def __init__(self, config: Optional[AppConfig] = None):
        self.config = config or get_config()
    
    def _get_master_studies(
        self,
        master_path: Optional[Path] = None,
        site_id: Optional[int] = None
    ) -> Tuple[List[Study], str]:
        """
        Get master studies from either a file or a database site.
        
        Returns: (studies, error_message)
        """
        if master_path:
            is_valid, error = validate_master_xlsx(master_path)
            if not is_valid:
                return [], error
            
            studies = parse_master_xlsx(master_path)
            return studies, ""
        
        elif site_id is not None:
            with DatabaseManager(config=self.config) as db:
                site = db.get_site(site_id)
                if not site:
                    return [], f"Site with ID {site_id} not found or access denied"
                
                studies = db.get_studies(site_id)
                return studies, ""
        
        return [], "No master source specified (provide master_path or site_id)"
    
    def _create_temp_copy(self, source_path: Path) -> Path:
        """Create a temporary copy of a file for processing."""
        temp_dir = self.config.get_temp_path()
        temp_dir.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        temp_path = temp_dir / f"temp_{timestamp}_{source_path.name}"
        shutil.copy2(source_path, temp_path)
        
        return temp_path
    
    def _cleanup_temp(self, temp_path: Path) -> None:
        """Remove temporary file."""
        try:
            if temp_path.exists():
                temp_path.unlink()
        except Exception:
            pass
    
    @staticmethod
    def _derive_original_cv_name(cv_path: Path) -> str:
        """Derive the original CV base name.

        Strategy:
          1. Read the custom doc property ``_original_cv_name`` if present.
          2. Otherwise strip date-stamped suffixes from the filename.

        Handles filenames like:
          "Jane Doe CV (Updated 2025-03-16).docx" -> "Jane Doe CV"
          "Jane Doe CV (Redacted 2025-03-16).docx" -> "Jane Doe CV"
          "Jane Doe CV (Updated 2025-03-16) (Redacted 2025-03-16).docx" -> "Jane Doe CV"
          "Jane Doe CV.docx" -> "Jane Doe CV"
        """
        from_doc = CVProcessor._get_original_cv_name_from_doc(cv_path)
        if from_doc:
            logging.info(
                "[Processor] Original CV name from doc property: '%s'",
                from_doc,
            )
            return from_doc

        stem = cv_path.stem
        prev = None
        while prev != stem:
            prev = stem
            stem = _RESULT_SUFFIX_RE.sub('', stem)
        result = stem.strip()
        logging.info(
            "[Processor] Original CV name from suffix strip: '%s'",
            result,
        )
        return result

    _ORIGINAL_CV_MARKER = "_original_cv_name:"

    def _resolve_output_path(
        self,
        cv_path: Path,
        suffix_label: str = "Updated",
        output_path: Optional[Path] = None,
    ) -> Path:
        """Resolve the output file path inside a per-CV result folder.

        Structure:  <project_root>/result/<original_cv_name>/<file>.docx

        If *output_path* is already set (user-specified), return it as-is.
        """
        if output_path is not None:
            return output_path

        original_name = self._derive_original_cv_name(cv_path)
        date_str = datetime.now().strftime("%Y-%m-%d")
        result_dir = self.config.get_result_root() / original_name
        result_dir.mkdir(parents=True, exist_ok=True)

        filename = f"{original_name} ({suffix_label} {date_str}){cv_path.suffix}"
        resolved = result_dir / filename

        logging.info(
            "[Processor] Output routing: cv='%s' -> folder='%s', file='%s'",
            cv_path.name,
            result_dir,
            filename,
        )
        return resolved

    @staticmethod
    def _set_original_cv_name(handler, original_name: str) -> None:
        """Store the original CV base name in the document properties."""
        try:
            marker = CVProcessor._ORIGINAL_CV_MARKER
            handler.document.core_properties.keywords = (
                f"{marker}{original_name}"
            )
            logging.info(
                "[Processor] Stored _original_cv_name='%s' in doc properties",
                original_name,
            )
        except Exception as exc:
            logging.warning(
                "[Processor] Could not set _original_cv_name: %s", exc
            )

    @staticmethod
    def _get_original_cv_name_from_doc(doc_path: Path) -> Optional[str]:
        """Read the original CV base name from document properties."""
        try:
            from docx import Document as _Doc
            doc = _Doc(doc_path)
            kw = doc.core_properties.keywords or ""
            marker = CVProcessor._ORIGINAL_CV_MARKER
            if kw.startswith(marker):
                return kw[len(marker):]
        except Exception:
            pass
        return None

    def _build_identity_set(self, studies: List[Study]) -> Set[tuple]:
        """Build a set of identity tuples for deduplication."""
        identities = set()
        for study in studies:
            normalized_masked = normalize_for_matching(study.description_masked)
            identities.add(study.get_identity_tuple(normalized_masked))
        return identities
    
    def mode_a_update_inject(
        self,
        cv_path: Path,
        master_path: Optional[Path] = None,
        site_id: Optional[int] = None,
        manual_benchmark_year: Optional[int] = None,
        output_path: Optional[Path] = None,
        enable_sort_existing: Optional[bool] = None,
    ) -> OperationResult:
        """
        Mode A: Update/Inject
        
        - Parse all CV studies (even without hierarchy)
        - Match each CV study to master to get Phase/Subcategory and unmask
        - Inject new studies from master that aren't in CV
        - Sort and format everything properly
        
        Returns: OperationResult with details
        """
        logger = OperationLogger(config=self.config)
        logger.start_operation("Mode A - Update/Inject")
        
        # Resolve enable_sort_existing: parameter > config > default (True)
        if enable_sort_existing is None:
            enable_sort_existing = self.config.enable_sort_existing
        logging.info(
            "[Processor] enable_sort_existing=%s (config default=%s)",
            enable_sort_existing,
            self.config.enable_sort_existing,
        )
        logger.log(
            "config",
            details=f"enable_sort_existing={enable_sort_existing}",
        )
        
        # Validate CV
        is_valid, error = validate_cv_docx(cv_path)
        if not is_valid:
            return logger.to_result(False, error_message=error)
        
        # Get master studies
        master_studies, error = self._get_master_studies(master_path, site_id)
        if error:
            return logger.to_result(False, error_message=error)
        
        if not master_studies:
            return logger.to_result(False, error_message="No studies found in master source")
        
        try:
            # Load and parse CV
            handler = CVDocxHandler(
                cv_path,
                font_name=self.config.font_name,
                font_size=self.config.font_size,
                hanging_indent_inches=self.config.hanging_indent_inches,
            )
            handler.load()
            
            start_idx, end_idx = handler.find_research_experience_section()
            if start_idx is None:
                return logger.to_result(
                    False,
                    error_message="Research Experience section not found in CV"
                )
            
            cv_research = handler.parse_research_experience()
            cv_studies = cv_research.get_all_studies()
            
            # Calculate year benchmark
            if manual_benchmark_year is not None:
                # Manual benchmark takes priority - inject from this year onward
                # year_bound is the year BEFORE the first year to inject
                year_bound = manual_benchmark_year - 1
                logging.info(f"[Processor] Using manual benchmark: inject from {manual_benchmark_year} onward (year_bound={year_bound})")
            else:
                # Auto-find benchmark: highest year among ALL CV studies
                cv_years = [s.year for s in cv_studies if s.year > 0]
                if cv_years:
                    year_bound = max(cv_years)
                    logging.info(f"[Processor] Year benchmark from CV studies: {year_bound} (max of {len(cv_years)} studies)")
                else:
                    # Fall back to standalone year benchmark from parser
                    year_bound = handler.year_bound
                    logging.info(f"[Processor] Year benchmark from standalone line: {year_bound}")
            
            # Build new research experience structure from scratch
            new_research = ResearchExperience()
            matched_master_ids = set()  # Track which master studies have been matched
            # Track object ids of studies that originate from the CV
            # (vs newly injected from master) for preserve-order mode.
            _existing_study_ids = []  # ordered list of (id(study_obj), study_obj)
            
            # Step 1: Match each CV study to master and reorganize
            for cv_study in cv_studies:
                # If year is missing (0 or unknown), try to infer it
                if cv_study.year == 0:
                    from normalizer import infer_year_from_master
                    inferred_year, inferred_study, reason = infer_year_from_master(
                        cv_study.sponsor,
                        cv_study.description_full,
                        master_studies,
                        heading_year_bound=None,  # TODO: Pass from parser if available
                        full_match_min_score=88,
                        masked_match_min_score=85,
                    )
                    
                    if inferred_year:
                        cv_study.year = inferred_year
                        logger.log_matched_existing(
                            cv_study.phase,
                            cv_study.subcategory,
                            cv_study.year,
                            cv_study.sponsor,
                            cv_study.protocol,
                            f"Year inferred: {reason}"
                        )
                    else:
                        logger.log_skipped_no_match(
                            cv_study.phase,
                            cv_study.subcategory,
                            0,
                            cv_study.sponsor,
                            cv_study.protocol,
                            f"Could not infer year: {reason}"
                        )
                
                # Try to match CV study to master using masked descriptions
                # (CV studies may already be masked with XXXX)
                match_result = match_study_to_master(
                    cv_study.year,
                    cv_study.sponsor,
                    cv_study.protocol,
                    cv_study.description_full,  # This is the CV text (possibly masked)
                    master_studies,
                    self.config.fuzzy_threshold_full,
                    self.config.fuzzy_threshold_masked,
                )
                
                if match_result:
                    matched_study, match_type, score = match_result
                    
                    # Use master's phase/subcategory for categorization
                    # But preserve original CV content if it was matched via masked description
                    # (CV study was already masked, so keep all original CV fields except category)
                    if 'masked' in match_type:
                        # CV was already masked - keep original CV sponsor, protocol, and description
                        updated_study = Study(
                            phase=matched_study.phase,
                            subcategory=matched_study.subcategory,
                            year=cv_study.year if cv_study.year > 0 else matched_study.year,
                            sponsor=cv_study.sponsor,
                            protocol=cv_study.protocol,
                            description_full=cv_study.description_full,
                            description_masked=cv_study.description_full,  # CV is already masked
                        )
                    else:
                        # CV had full description - use master's full data
                        updated_study = Study(
                            phase=matched_study.phase,
                            subcategory=matched_study.subcategory,
                            year=matched_study.year,
                            sponsor=matched_study.sponsor,
                            protocol=matched_study.protocol,
                            description_full=matched_study.description_full,
                            description_masked=matched_study.description_masked,
                        )
                    
                    # Add to new structure
                    from normalizer import normalize_heading_key, normalize_subcat_key
                    phase_key = normalize_heading_key(updated_study.phase)
                    subcat_key = normalize_subcat_key(updated_study.subcategory)
                    phase = new_research.get_or_create_phase(updated_study.phase)
                    subcategory = phase.get_or_create_subcategory(updated_study.subcategory)
                    subcategory.studies.append(updated_study)
                    _existing_study_ids.append((id(updated_study), updated_study))
                    
                    logging.info(
                        "[Processor] CV study matched -> phase='%s' "
                        "(key='%s', node='%s'), subcat='%s' "
                        "(key='%s', node='%s')",
                        updated_study.phase,
                        phase_key,
                        phase.name,
                        updated_study.subcategory,
                        subcat_key,
                        subcategory.name,
                    )
                    
                    # Track matched master study
                    master_id = (matched_study.year, matched_study.sponsor, matched_study.protocol)
                    matched_master_ids.add(master_id)
                    
                    logger.log_matched_existing(
                        updated_study.phase,
                        updated_study.subcategory,
                        updated_study.year,
                        updated_study.sponsor,
                        updated_study.protocol,
                        f"Matched and categorized ({match_type}, score={score}), "
                        f"phase_key='{phase_key}', subcat_key='{subcat_key}'"
                    )
                else:
                    # No match - keep original but in Uncategorized
                    uncat_label = self.config.uncategorized_label
                    phase = new_research.get_or_create_phase(uncat_label)
                    subcategory = phase.get_or_create_subcategory("General")
                    subcategory.studies.append(cv_study)
                    _existing_study_ids.append((id(cv_study), cv_study))
                    
                    logger.log_skipped_no_match(
                        cv_study.phase,
                        cv_study.subcategory,
                        cv_study.year,
                        cv_study.sponsor,
                        cv_study.protocol,
                        f"No match in master - kept in {self.config.uncategorized_label}"
                    )
            
            # Step 2: Inject master studies not in CV
            # Only inject studies AFTER the year bound (calculated earlier from CV studies)
            logging.info(f"[Processor] Year bound for injection: {year_bound}")
            studies_injected = 0
            for master_study in master_studies:
                master_id = (master_study.year, master_study.sponsor, master_study.protocol)
                
                if master_id not in matched_master_ids:
                    # Check year bound - only inject if study year > year_bound
                    if year_bound is not None and master_study.year > 0:
                        if master_study.year <= year_bound:
                            logger.log_skipped_no_match(
                                master_study.phase,
                                master_study.subcategory,
                                master_study.year,
                                master_study.sponsor,
                                master_study.protocol,
                                f"Skipped - year {master_study.year} <= benchmark {year_bound}"
                            )
                            continue
                    
                    # This master study is not in CV - inject it
                    from normalizer import normalize_heading_key, normalize_subcat_key
                    inj_phase_key = normalize_heading_key(master_study.phase)
                    inj_subcat_key = normalize_subcat_key(master_study.subcategory)
                    phase = new_research.get_or_create_phase(master_study.phase)
                    subcategory = phase.get_or_create_subcategory(master_study.subcategory)
                    subcategory.studies.append(master_study)
                    
                    logging.info(
                        "[Processor] Injecting '%s %s' -> phase='%s' "
                        "(key='%s', node='%s'), subcat='%s' "
                        "(key='%s', node='%s')",
                        master_study.sponsor,
                        master_study.protocol,
                        master_study.phase,
                        inj_phase_key,
                        phase.name,
                        master_study.subcategory,
                        inj_subcat_key,
                        subcategory.name,
                    )
                    
                    logger.log_inserted(
                        master_study.phase,
                        master_study.subcategory,
                        master_study.year,
                        master_study.sponsor,
                        master_study.protocol,
                        f"Injected from master (not in CV)"
                        f"{f', year > {year_bound}' if year_bound else ''}, "
                        f"phase_key='{inj_phase_key}', "
                        f"subcat_key='{inj_subcat_key}', "
                        f"container_phase='{phase.name}', "
                        f"container_subcat='{subcategory.name}'"
                    )
                    studies_injected += 1
            
            # Check if any changes were made
            total_studies = len(new_research.get_all_studies())
            if total_studies == 0:
                logger.log_no_changes("No studies found after processing")
                log_json = logger.save_json()
                log_csv = logger.save_csv()
                return logger.to_result(True, error_message="No studies to process")
            
            # Resolve highlight_inserted from config
            highlight_inserted = self.config.highlight_inserted
            logging.info(
                "[Processor] highlight_inserted=%s",
                highlight_inserted,
            )

            # Build existing object ID set (used by both paths)
            existing_obj_ids = set(
                obj_id for obj_id, _ in _existing_study_ids
            )

            # Ensure category order entries for all phase/subcat combos in this run
            custom_order = None
            if site_id:
                try:
                    with DatabaseManager(config=self.config) as db:
                        order_entries = []
                        for phase in new_research.phases:
                            for subcat in phase.subcategories:
                                order_entries.append((phase.name, subcat.name))
                        if order_entries:
                            db._ensure_category_order_entries(site_id, order_entries)
                            logging.info(
                                "[Processor] Ensured %d phase/subcat entries in category order",
                                len(order_entries),
                            )
                        custom_order = db.get_category_order(site_id)
                except Exception:
                    pass
            
            if enable_sort_existing:
                # --- Path A: Full sort (current behavior) ---
                logging.info("[Processor] Sorting all studies (enable_sort_existing=True)")
                if custom_order:
                    new_research.sort_all_custom(custom_order)
                else:
                    new_research.sort_all()

                # Write back to document (full rewrite)
                handler.write_research_experience(
                    new_research,
                    include_protocol=True,
                    protocol_red=True,
                    highlight_new=highlight_inserted,
                    new_study_ids=existing_obj_ids if highlight_inserted else None,
                )
            else:
                # --- Path B: Preserve existing paragraphs completely ---
                # Only inject new study paragraphs; do NOT rewrite existing
                # content.  This preserves all original formatting, runs,
                # indentation, tabs, spacing, and ordering.
                logging.info(
                    "[Processor] Preserving existing paragraphs "
                    "(enable_sort_existing=False)"
                )
                existing_obj_ids = set(
                    obj_id for obj_id, _ in _existing_study_ids
                )

                # Collect only new studies with their target phase/subcat
                new_studies_for_injection = []
                for phase in new_research.phases:
                    for subcat in phase.subcategories:
                        for study in subcat.studies:
                            if id(study) not in existing_obj_ids:
                                new_studies_for_injection.append(
                                    (study, phase.name, subcat.name)
                                )

                logging.info(
                    "[Processor] Preserve-existing path: %d new studies "
                    "to inject into document",
                    len(new_studies_for_injection),
                )

                inserted_count = handler.inject_new_studies_only(
                    new_studies_for_injection,
                    include_protocol=True,
                    protocol_red=True,
                    highlight_inserted=highlight_inserted,
                )

                logger.log(
                    "splice-info",
                    details=(
                        f"Preserve-existing: injected {inserted_count} "
                        f"paragraphs without rewriting existing content"
                    ),
                )

            if output_path is None:
                output_path = self._resolve_output_path(
                    cv_path, suffix_label="Updated"
                )

            original_name = self._derive_original_cv_name(cv_path)
            self._set_original_cv_name(handler, original_name)
            
            try:
                handler.save(output_path)
            except PermissionError:
                raise FilePermissionError(output_path, "save")
            
            log_json = logger.save_json()
            log_csv = logger.save_csv()
            
            return logger.to_result(True, str(output_path))
            
        except Exception as e:
            return logger.to_result(False, error_message=f"Processing error: {str(e)}")
    
    def _splice_new_studies_preserving_order(
        self,
        new_research: ResearchExperience,
        existing_study_ids: list,
        studies_injected: int,
        year_bound,
        custom_order,
        logger: OperationLogger,
    ) -> None:
        """
        Re-order studies within *new_research* so that pre-existing CV studies
        keep their original relative order while newly injected studies are
        sorted among themselves and spliced in as a contiguous block at the
        correct insertion point per subcategory.

        Args:
            existing_study_ids: ordered list of (id(study_obj), study_obj)
                collected during Step 1 (CV matching).  The list order
                reflects the original CV paragraph order.

        Insertion point logic (per subcategory):
          1. Find the first existing study whose year >= year_bound.
             Insert the new block immediately before that study.
          2. If no existing study is at or above the year_bound, insert
             the new block at the top of the subcategory.
          3. If the subcategory has no existing studies at all, the new
             block becomes the entire study list (already sorted).

        Phase ordering still respects configured phase_order / custom_order.
        Subcategories are sorted alphabetically (same as normal mode).
        """
        # Build a set of python object ids for studies from the original CV
        existing_obj_ids = set(obj_id for obj_id, _ in existing_study_ids)

        # Build an order-index map keyed by python object id
        original_order = {}
        for idx, (obj_id, _) in enumerate(existing_study_ids):
            original_order[obj_id] = idx

        # Sort phases (always sorted by configured order)
        new_research.phases.sort(
            key=lambda p: new_research.get_phase_order_key(p.name)
        )

        for phase in new_research.phases:
            # Sort subcategories alphabetically
            phase.subcategories.sort(key=lambda sc: sc.name.lower())

            for subcat in phase.subcategories:
                existing = []
                newly_injected = []

                for study in subcat.studies:
                    if id(study) in existing_obj_ids:
                        existing.append(study)
                    else:
                        newly_injected.append(study)

                # Preserve original relative order for existing studies
                existing.sort(
                    key=lambda s: original_order.get(id(s), 999999)
                )

                # Sort new studies among themselves: year desc, sponsor, protocol
                newly_injected.sort(
                    key=lambda s: (
                        -s.year,
                        s.sponsor.lower(),
                        s.protocol.lower(),
                    )
                )

                existing_count = len(existing)
                new_count = len(newly_injected)

                if new_count == 0:
                    # Nothing to splice — keep existing order as-is
                    subcat.studies = existing
                    logging.info(
                        "[Processor] %s > %s: %d existing, 0 new — no splice needed",
                        phase.name,
                        subcat.name,
                        existing_count,
                    )
                    continue

                if existing_count == 0:
                    # No existing studies — new block is the whole list
                    subcat.studies = newly_injected
                    logger.log(
                        "splice-info",
                        phase=phase.name,
                        subcategory=subcat.name,
                        details=(
                            f"Subcategory empty; inserted {new_count} new "
                            f"studies as entire list"
                        ),
                    )
                    logging.info(
                        "[Processor] %s > %s: 0 existing, %d new — "
                        "inserted as entire list",
                        phase.name,
                        subcat.name,
                        new_count,
                    )
                    continue

                # Find insertion anchor: first existing study at/above year_bound
                anchor_idx = None
                anchor_reason = ""
                for i, s in enumerate(existing):
                    if year_bound is not None and s.year >= year_bound:
                        anchor_idx = i
                        anchor_reason = (
                            f"first existing study at/above benchmark "
                            f"year {year_bound} (year={s.year}, "
                            f"sponsor={s.sponsor})"
                        )
                        break

                if anchor_idx is None:
                    # No existing study at/above benchmark — insert at top
                    anchor_idx = 0
                    anchor_reason = (
                        f"no existing study at/above benchmark "
                        f"year {year_bound}; inserting at top"
                    )

                # Splice: existing[:anchor] + new_block + existing[anchor:]
                merged = (
                    existing[:anchor_idx]
                    + newly_injected
                    + existing[anchor_idx:]
                )
                subcat.studies = merged

                logger.log(
                    "splice-info",
                    phase=phase.name,
                    subcategory=subcat.name,
                    details=(
                        f"existing={existing_count}, new={new_count}, "
                        f"anchor_idx={anchor_idx}, reason={anchor_reason}"
                    ),
                )
                logging.info(
                    "[Processor] %s > %s: %d existing, %d new, "
                    "anchor_idx=%d (%s)",
                    phase.name,
                    subcat.name,
                    existing_count,
                    new_count,
                    anchor_idx,
                    anchor_reason,
                )

    def mode_b_redact_protocols(
        self,
        cv_path: Path,
        master_path: Optional[Path] = None,
        site_id: Optional[int] = None,
        output_path: Optional[Path] = None,
        sort_and_format: bool = False,
    ) -> OperationResult:
        """Mode B: Redact Protocols (in-place replacement).

        Only studies that contain a protocol token are redacted.
        Non-protocol studies are never touched. Already-masked lines
        are detected and skipped for idempotency.

        Args:
            sort_and_format: When True, subcategories that received at
                least one replacement are re-sorted (year desc, sponsor,
                protocol) after redaction.  Subcategories with zero
                replacements are untouched.  Default False.
        """
        logger = OperationLogger(config=self.config)
        logger.start_operation("Mode B - Redact Protocols")
        logger.log(
            "config",
            details=f"sort_and_format={sort_and_format}",
        )
        logging.info(
            "[Processor] mode_b_redact_protocols: sort_and_format=%s",
            sort_and_format,
        )

        is_valid, error = validate_cv_docx(cv_path)
        if not is_valid:
            return logger.to_result(False, error_message=error)

        master_studies, error = self._get_master_studies(master_path, site_id)
        if error:
            return logger.to_result(False, error_message=error)

        if not master_studies:
            return logger.to_result(
                False, error_message="No studies found in master source"
            )

        try:
            handler = CVDocxHandler(
                cv_path,
                font_name=self.config.font_name,
                font_size=self.config.font_size,
                hanging_indent_inches=self.config.hanging_indent_inches,
            )
            handler.load()

            start_idx, end_idx = handler.find_research_experience_section()
            if start_idx is None:
                return logger.to_result(
                    False,
                    error_message="Research Experience section not found in CV",
                )

            cv_research = handler.parse_research_experience()

            replacements = []
            affected_subcats = set()

            for phase in cv_research.phases:
                p_key = normalize_heading_key(phase.name)
                for subcategory in phase.subcategories:
                    s_key = normalize_subcat_key(subcategory.name)
                    subcat_tuple = (p_key, s_key)
                    study_para_list = handler._subcat_study_para_list.get(
                        subcat_tuple, []
                    )

                    for idx_in_list, study in enumerate(subcategory.studies):
                        raw_text = study.description_full
                        full_line = (
                            f"{study.sponsor} {study.protocol}: {raw_text}"
                            if study.protocol
                            else f"{study.sponsor}: {raw_text}"
                        )

                        if is_already_masked(full_line):
                            logger.log(
                                "skipped-already-masked",
                                phase=phase.name,
                                subcategory=subcategory.name,
                                year=study.year,
                                sponsor=study.sponsor,
                                protocol="",
                                details="Line already masked — skipped",
                            )
                            logging.debug(
                                "[Processor] Skipped already-masked: %s",
                                full_line[:80],
                            )
                            continue

                        if not contains_protocol_token(full_line):
                            logger.log(
                                "skipped-no-protocol",
                                phase=phase.name,
                                subcategory=subcategory.name,
                                year=study.year,
                                sponsor=study.sponsor,
                                protocol="",
                                details="No protocol token — not redacted",
                            )
                            continue

                        match_result = match_study_to_master(
                            study.year,
                            study.sponsor,
                            study.protocol,
                            study.description_full,
                            master_studies,
                            self.config.fuzzy_threshold_full,
                            self.config.fuzzy_threshold_masked,
                        )

                        if not match_result:
                            logger.log(
                                "skipped-no-protocol",
                                phase=phase.name,
                                subcategory=subcategory.name,
                                year=study.year,
                                sponsor=study.sponsor,
                                protocol=study.protocol,
                                details=(
                                    "Protocol present but no master match — "
                                    "kept original"
                                ),
                            )
                            continue

                        matched_study, match_type, score = match_result

                        if idx_in_list < len(study_para_list):
                            para_idx = study_para_list[idx_in_list]
                        else:
                            para_idx = handler._subcat_last_study_para.get(
                                subcat_tuple, -1
                            )

                        replacements.append({
                            "para_idx": para_idx,
                            "year": study.year,
                            "masked_sponsor": matched_study.sponsor,
                            "masked_description": matched_study.description_masked,
                            "phase": phase.name,
                            "subcategory": subcategory.name,
                            "protocol": study.protocol,
                            "match_type": match_type,
                            "score": score,
                        })
                        affected_subcats.add(subcat_tuple)

                        logger.log_replaced(
                            phase.name,
                            subcategory.name,
                            study.year,
                            matched_study.sponsor,
                            study.protocol,
                            (
                                f"Matched ({match_type}, score={score}), "
                                f"redacted in place at para {para_idx}"
                            ),
                        )

            replaced_count = handler.redact_studies_in_place(replacements)
            logging.info(
                "[Processor] Redacted %d paragraphs in place", replaced_count,
            )

            if sort_and_format and affected_subcats:
                logging.info(
                    "[Processor] Re-sorting %d affected subcategories",
                    len(affected_subcats),
                )
                for (pk, sk) in affected_subcats:
                    handler.sort_subcategory_in_place(pk, sk)
                    logger.log(
                        "sort-category",
                        details=(
                            f"Re-sorted subcategory ({pk}, {sk}) "
                            f"after redaction"
                        ),
                    )

            if output_path is None:
                output_path = self._resolve_output_path(
                    cv_path, suffix_label="Redacted",
                )

            original_name = self._derive_original_cv_name(cv_path)
            self._set_original_cv_name(handler, original_name)

            try:
                handler.save(output_path)
            except PermissionError:
                raise FilePermissionError(output_path, "save")

            log_json = logger.save_json()
            log_csv = logger.save_csv()

            return logger.to_result(True, str(output_path))

        except Exception as e:
            return logger.to_result(
                False, error_message=f"Processing error: {str(e)}"
            )
    
    def preview_changes(
        self,
        cv_path: Path,
        master_path: Optional[Path] = None,
        site_id: Optional[int] = None,
        mode: str = "update_inject",
        enable_sort_existing: Optional[bool] = None,
        sort_and_format: bool = False,
    ) -> Tuple[List[dict], str]:
        """Preview changes that would be made without applying them.

        Returns: (list of changes, error_message)
        """
        if enable_sort_existing is None:
            enable_sort_existing = self.config.enable_sort_existing
        logging.info(
            "[Processor] preview_changes: enable_sort_existing=%s, "
            "sort_and_format=%s",
            enable_sort_existing,
            sort_and_format,
        )

        is_valid, error = validate_cv_docx(cv_path)
        if not is_valid:
            return [], error

        master_studies, error = self._get_master_studies(master_path, site_id)
        if error:
            return [], error

        if not master_studies:
            return [], "No studies found in master source"

        try:
            handler = CVDocxHandler(
                cv_path,
                font_name=self.config.font_name,
                font_size=self.config.font_size,
                hanging_indent_inches=self.config.hanging_indent_inches,
            )
            handler.load()

            start_idx, end_idx = handler.find_research_experience_section()
            if start_idx is None:
                return [], "Research Experience section not found in CV"

            cv_research = handler.parse_research_experience()
            changes = []

            if mode == "update_inject":
                benchmark_year = cv_research.calculate_benchmark_year(
                    self.config.benchmark_min_count
                )
                existing_identities = self._build_identity_set(
                    cv_research.get_all_studies()
                )

                for master_study in master_studies:
                    if master_study.year <= benchmark_year:
                        continue

                    master_normalized = normalize_for_matching(
                        master_study.description_masked
                    )
                    master_identity = master_study.get_identity_tuple(
                        master_normalized
                    )

                    if master_identity not in existing_identities:
                        prev_phase_key = normalize_heading_key(
                            master_study.phase
                        )
                        prev_subcat_key = normalize_subcat_key(
                            master_study.subcategory
                        )
                        matched_phase_node = None
                        matched_subcat_node = None
                        for p in cv_research.phases:
                            if normalize_heading_key(p.name) == prev_phase_key:
                                matched_phase_node = p.name
                                for sc in p.subcategories:
                                    if (
                                        normalize_subcat_key(sc.name)
                                        == prev_subcat_key
                                    ):
                                        matched_subcat_node = sc.name
                                        break
                                break
                        desc_text = master_study.description_full
                        change_entry = {
                            "action": "inject",
                            "phase": master_study.phase,
                            "subcategory": master_study.subcategory,
                            "phase_key": prev_phase_key,
                            "subcat_key": prev_subcat_key,
                            "matched_phase_container": matched_phase_node,
                            "matched_subcat_container": matched_subcat_node,
                            "year": master_study.year,
                            "sponsor": master_study.sponsor,
                            "protocol": master_study.protocol,
                            "description": (
                                desc_text[:100] + "..."
                                if len(desc_text) > 100
                                else desc_text
                            ),
                            "enable_sort_existing": enable_sort_existing,
                        }
                        changes.append(change_entry)
                        existing_identities.add(master_identity)

            elif mode == "redact_protocols":
                affected_subcats_preview = set()
                for phase in cv_research.phases:
                    p_key = normalize_heading_key(phase.name)
                    for subcategory in phase.subcategories:
                        s_key = normalize_subcat_key(subcategory.name)
                        subcat_tuple = (p_key, s_key)
                        study_para_list = (
                            handler._subcat_study_para_list.get(
                                subcat_tuple, []
                            )
                        )

                        for idx_in_list, study in enumerate(
                            subcategory.studies
                        ):
                            full_line = (
                                f"{study.sponsor} {study.protocol}: "
                                f"{study.description_full}"
                                if study.protocol
                                else f"{study.sponsor}: "
                                     f"{study.description_full}"
                            )

                            if is_already_masked(full_line):
                                changes.append({
                                    "action": "skipped-already-masked",
                                    "phase": phase.name,
                                    "subcategory": subcategory.name,
                                    "year": study.year,
                                    "sponsor": study.sponsor,
                                    "protocol": "",
                                    "reason": "Already masked",
                                })
                                continue

                            if not contains_protocol_token(full_line):
                                changes.append({
                                    "action": "skipped-no-protocol",
                                    "phase": phase.name,
                                    "subcategory": subcategory.name,
                                    "year": study.year,
                                    "sponsor": study.sponsor,
                                    "protocol": "",
                                    "reason": "No protocol token",
                                })
                                continue

                            match_result = match_study_to_master(
                                study.year,
                                study.sponsor,
                                study.protocol,
                                study.description_full,
                                master_studies,
                                self.config.fuzzy_threshold_full,
                                self.config.fuzzy_threshold_masked,
                            )

                            if match_result:
                                matched_study, match_type, score = (
                                    match_result
                                )
                                anchor_idx = -1
                                if idx_in_list < len(study_para_list):
                                    anchor_idx = study_para_list[idx_in_list]
                                masked_desc = (
                                    matched_study.description_masked
                                )
                                affected_subcats_preview.add(subcat_tuple)
                                changes.append({
                                    "action": "redact",
                                    "phase": phase.name,
                                    "subcategory": subcategory.name,
                                    "year": study.year,
                                    "sponsor": study.sponsor,
                                    "protocol": study.protocol,
                                    "match_type": match_type,
                                    "match_score": score,
                                    "anchor_para_idx": anchor_idx,
                                    "new_sponsor": matched_study.sponsor,
                                    "new_description": (
                                        masked_desc[:100] + "..."
                                        if len(masked_desc) > 100
                                        else masked_desc
                                    ),
                                    "sort_and_format": sort_and_format,
                                    "would_resort_category": (
                                        sort_and_format
                                    ),
                                })
                            else:
                                changes.append({
                                    "action": "skipped-no-match",
                                    "phase": phase.name,
                                    "subcategory": subcategory.name,
                                    "year": study.year,
                                    "sponsor": study.sponsor,
                                    "protocol": study.protocol,
                                    "reason": (
                                        "Protocol present but no master match"
                                    ),
                                })

                if sort_and_format and affected_subcats_preview:
                    for c in changes:
                        if c.get("action") == "redact":
                            pk = normalize_heading_key(c["phase"])
                            sk = normalize_subcat_key(c["subcategory"])
                            c["would_resort_category"] = (
                                (pk, sk) in affected_subcats_preview
                            )

            return changes, ""

        except Exception as e:
            return [], f"Preview error: {str(e)}"
