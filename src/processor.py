"""
Core processing logic for CV modification modes.
Mode A: Update/Inject
Mode B: Redact Protocols
"""

from pathlib import Path
from typing import List, Optional, Tuple, Set
from datetime import datetime
import shutil

from models import Study, ResearchExperience, OperationResult, LogEntry
from normalizer import (
    normalize_for_matching, match_study_to_master, normalize_phase,
    collapse_x_runs
)
from docx_handler import CVDocxHandler, validate_cv_docx
from excel_parser import parse_master_xlsx, validate_master_xlsx, studies_to_research_experience
from database import DatabaseManager
from logger import OperationLogger
from config import get_config, AppConfig
from error_handler import FilePermissionError

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
            import logging
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
                    phase = new_research.get_or_create_phase(updated_study.phase)
                    subcategory = phase.get_or_create_subcategory(updated_study.subcategory)
                    subcategory.studies.append(updated_study)
                    
                    # Track matched master study
                    master_id = (matched_study.year, matched_study.sponsor, matched_study.protocol)
                    matched_master_ids.add(master_id)
                    
                    logger.log_matched_existing(
                        updated_study.phase,
                        updated_study.subcategory,
                        updated_study.year,
                        updated_study.sponsor,
                        updated_study.protocol,
                        f"Matched and categorized ({match_type}, score={score})"
                    )
                else:
                    # No match - keep original but in Uncategorized
                    phase = new_research.get_or_create_phase("Uncategorized")
                    subcategory = phase.get_or_create_subcategory("General")
                    subcategory.studies.append(cv_study)
                    
                    logger.log_skipped_no_match(
                        cv_study.phase,
                        cv_study.subcategory,
                        cv_study.year,
                        cv_study.sponsor,
                        cv_study.protocol,
                        "No match in master - kept in Uncategorized"
                    )
            
            # Step 2: Inject master studies not in CV
            # Only inject studies AFTER the year bound (calculated earlier from CV studies)
            import logging
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
                    phase = new_research.get_or_create_phase(master_study.phase)
                    subcategory = phase.get_or_create_subcategory(master_study.subcategory)
                    subcategory.studies.append(master_study)
                    
                    logger.log_inserted(
                        master_study.phase,
                        master_study.subcategory,
                        master_study.year,
                        master_study.sponsor,
                        master_study.protocol,
                        f"Injected from master (not in CV){f', year > {year_bound}' if year_bound else ''}"
                    )
                    studies_injected += 1
            
            # Check if any changes were made
            total_studies = len(new_research.get_all_studies())
            if total_studies == 0:
                logger.log_no_changes("No studies found after processing")
                log_json = logger.save_json()
                log_csv = logger.save_csv()
                return logger.to_result(True, error_message="No studies to process")
            
            # Sort the structure - apply custom order if available
            custom_order = None
            if site_id:
                try:
                    with DatabaseManager(config=self.config) as db:
                        custom_order = db.get_category_order(site_id)
                except:
                    pass
            
            if custom_order:
                new_research.sort_all_custom(custom_order)
            else:
                new_research.sort_all()
            
            # Write back to document
            handler.write_research_experience(
                new_research,
                include_protocol=True,
                protocol_red=True
            )
            
            # Save
            if output_path is None:
                date_str = datetime.now().strftime("%Y-%m-%d")
                result_dir = Path(__file__).parent.parent / "result"
                result_dir.mkdir(exist_ok=True)
                output_path = result_dir / f"{cv_path.stem} (Updated {date_str}){cv_path.suffix}"
            
            try:
                handler.save(output_path)
            except PermissionError:
                raise FilePermissionError(output_path, "save")
            
            # Save logs
            log_json = logger.save_json()
            log_csv = logger.save_csv()
            
            return logger.to_result(True, str(output_path))
            
        except Exception as e:
            return logger.to_result(False, error_message=f"Processing error: {str(e)}")
    
    def mode_b_redact_protocols(
        self,
        cv_path: Path,
        master_path: Optional[Path] = None,
        site_id: Optional[int] = None,
        output_path: Optional[Path] = None,
    ) -> OperationResult:
        """
        Mode B: Redact Protocols
        
        - For each CV study, match to master by Year and Column B
        - Replace with Column C (masked - no protocol, treatments as XXX)
        - Keep hierarchy and sorting
        
        Returns: OperationResult with details
        """
        logger = OperationLogger(config=self.config)
        logger.start_operation("Mode B - Redact Protocols")
        
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
            )
            handler.load()
            
            start_idx, end_idx = handler.find_research_experience_section()
            if start_idx is None:
                return logger.to_result(
                    False,
                    error_message="Research Experience section not found in CV"
                )
            
            cv_research = handler.parse_research_experience()
            
            # Process each study: match and redact
            # IMPORTANT: All studies are kept - we only modify matched ones
            changes_made = False
            
            for phase in cv_research.phases:
                for subcategory in phase.subcategories:
                    for study in subcategory.studies:
                        # Try to match to master
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
                            matched_study, match_type, score = match_result
                            
                            # Update study with masked version from master
                            old_protocol = study.protocol
                            study.protocol = ""  # Remove protocol
                            
                            # Use master's masked description (already excludes sponsor prefix)
                            study.description_masked = matched_study.description_masked
                            # Also use master's sponsor to ensure consistency
                            study.sponsor = matched_study.sponsor
                            
                            logger.log_replaced(
                                phase.name,
                                subcategory.name,
                                study.year,
                                study.sponsor,
                                old_protocol,
                                f"Matched ({match_type}, score={score}), redacted protocol"
                            )
                            changes_made = True
                        else:
                            # No match found - study might already be masked or not in master
                            # KEEP THE STUDY AS-IS (don't lose it!)
                            # Remove protocol if present, but keep description as-is
                            if study.protocol:
                                study.protocol = ""
                            # Log that we kept it without matching
                            logger.log_skipped_no_match(
                                phase.name,
                                subcategory.name,
                                study.year,
                                study.sponsor,
                                "",  # Protocol already removed
                                "No match in master - kept original, removed protocol"
                            )
            
            # Sort the structure - apply custom order if available
            custom_order = None
            if site_id:
                try:
                    with DatabaseManager(config=self.config) as db:
                        custom_order = db.get_category_order(site_id)
                except:
                    pass
            
            if custom_order:
                cv_research.sort_all_custom(custom_order)
            else:
                cv_research.sort_all()
            
            # Write back to document (no protocol, not red)
            handler.write_research_experience(
                cv_research,
                include_protocol=False,
                protocol_red=False
            )
            
            # Save
            if output_path is None:
                date_str = datetime.now().strftime("%Y-%m-%d")
                result_dir = Path(__file__).parent.parent / "result"
                result_dir.mkdir(exist_ok=True)
                output_path = result_dir / f"{cv_path.stem} (Redacted {date_str}){cv_path.suffix}"  
            
            try:
                handler.save(output_path)
            except PermissionError:
                raise FilePermissionError(output_path, "save")
            
            # Save logs
            log_json = logger.save_json()
            log_csv = logger.save_csv()
            
            return logger.to_result(True, str(output_path))
            
        except Exception as e:
            return logger.to_result(False, error_message=f"Processing error: {str(e)}")
    
    def preview_changes(
        self,
        cv_path: Path,
        master_path: Optional[Path] = None,
        site_id: Optional[int] = None,
        mode: str = "update_inject"
    ) -> Tuple[List[dict], str]:
        """
        Preview changes that would be made without applying them.
        
        Returns: (list of changes, error_message)
        """
        # Validate CV
        is_valid, error = validate_cv_docx(cv_path)
        if not is_valid:
            return [], error
        
        # Get master studies
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
                existing_identities = self._build_identity_set(cv_research.get_all_studies())
                
                for master_study in master_studies:
                    if master_study.year <= benchmark_year:
                        continue
                    
                    master_normalized = normalize_for_matching(master_study.description_masked)
                    master_identity = master_study.get_identity_tuple(master_normalized)
                    
                    if master_identity not in existing_identities:
                        changes.append({
                            "action": "inject",
                            "phase": master_study.phase,
                            "subcategory": master_study.subcategory,
                            "year": master_study.year,
                            "sponsor": master_study.sponsor,
                            "protocol": master_study.protocol,
                            "description": master_study.description_full[:100] + "..." if len(master_study.description_full) > 100 else master_study.description_full,
                        })
                        existing_identities.add(master_identity)
            
            elif mode == "redact_protocols":
                for phase in cv_research.phases:
                    for subcategory in phase.subcategories:
                        for study in subcategory.studies:
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
                                matched_study, match_type, score = match_result
                                changes.append({
                                    "action": "redact",
                                    "phase": phase.name,
                                    "subcategory": subcategory.name,
                                    "year": study.year,
                                    "sponsor": study.sponsor,
                                    "protocol": study.protocol,
                                    "match_score": score,
                                    "new_description": matched_study.description_masked[:100] + "..." if len(matched_study.description_masked) > 100 else matched_study.description_masked,
                                })
            
            return changes, ""
            
        except Exception as e:
            return [], f"Preview error: {str(e)}"
