"""
Import/Export functionality for site databases.
Handles .xlsx file import and export operations.
"""

from pathlib import Path
from datetime import datetime
from typing import List, Tuple, Optional
import shutil

from models import Study, Site
from database import DatabaseManager
from excel_parser import parse_master_xlsx, export_studies_to_xlsx, validate_master_xlsx
from logger import OperationLogger
from config import get_config, AppConfig


class ImportExportManager:
    """Manages import and export operations for site databases."""
    
    def __init__(self, config: Optional[AppConfig] = None):
        self.config = config or get_config()
    
    def import_xlsx_to_site(
        self,
        xlsx_path: Path,
        site_name: str,
        user_id: Optional[str] = None,
        replace_existing: bool = False,
    ) -> Tuple[bool, str, Optional[int]]:
        """
        Import an .xlsx file into a new or existing site.
        
        Args:
            xlsx_path: Path to the .xlsx file
            site_name: Name for the site
            user_id: User ID (defaults to current OS user)
            replace_existing: If True and site exists, replace its studies
        
        Returns: (success, message, site_id)
        """
        user_id = user_id or self.config.get_user_id()
        
        # Validate the xlsx file
        is_valid, error = validate_master_xlsx(xlsx_path)
        if not is_valid:
            return False, f"Invalid file: {error}", None
        
        try:
            # Parse the xlsx file
            studies = parse_master_xlsx(xlsx_path)
            
            if not studies:
                return False, "No studies found in the file", None
            
            with DatabaseManager(user_id=user_id, config=self.config) as db:
                # Check if site with this name already exists
                existing_sites = db.get_sites()
                existing_site = next(
                    (s for s in existing_sites if s.name.lower() == site_name.lower()),
                    None
                )
                
                if existing_site:
                    if replace_existing:
                        # Backup existing data
                        db.create_site_backup(existing_site.id, "Pre-import backup")
                        # Clear existing studies
                        db.clear_studies(existing_site.id)
                        site_id = existing_site.id
                    else:
                        return False, f"Site '{site_name}' already exists. Use replace_existing=True to replace.", None
                else:
                    # Create new site
                    site = db.create_site(site_name)
                    site_id = site.id
                
                # Add studies to site
                count = db.bulk_add_studies(site_id, studies)
                
                # Copy original file to imports folder
                imports_path = self.config.get_user_imports_path(user_id)
                imports_path.mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                import_copy = imports_path / f"{site_name}_{timestamp}{xlsx_path.suffix}"
                shutil.copy2(xlsx_path, import_copy)
                
                return True, f"Successfully imported {count} studies into site '{site_name}'", site_id
                
        except Exception as e:
            return False, f"Import error: {str(e)}", None
    
    def export_site_to_xlsx(
        self,
        site_id: int,
        output_path: Optional[Path] = None,
        user_id: Optional[str] = None,
    ) -> Tuple[bool, str, Optional[Path]]:
        """
        Export a site's studies to an .xlsx file.
        
        Args:
            site_id: ID of the site to export
            output_path: Output file path (defaults to exports folder)
            user_id: User ID (defaults to current OS user)
        
        Returns: (success, message, output_path)
        """
        user_id = user_id or self.config.get_user_id()
        
        try:
            with DatabaseManager(user_id=user_id, config=self.config) as db:
                # Verify site ownership and get site
                site = db.get_site(site_id)
                if not site:
                    return False, f"Site with ID {site_id} not found or access denied", None
                
                # Get studies
                studies = db.get_studies(site_id)
                
                if not studies:
                    return False, "No studies found in site", None
                
                if output_path is None:
                    result_dir = Path(__file__).parent.parent / "result"
                    safe_name = "".join(
                        c if c.isalnum() or c in '-_ ' else '_'
                        for c in site.name
                    )
                    site_folder = result_dir / safe_name
                    site_folder.mkdir(parents=True, exist_ok=True)
                    output_path = site_folder / f"{safe_name} - Master Study List.xlsx"
                
                # Get custom category order for this site
                custom_order = db.get_category_order(site_id)
                
                # Export
                try:
                    export_studies_to_xlsx(studies, output_path, custom_order=custom_order)
                except PermissionError:
                    from error_handler import FilePermissionError
                    raise FilePermissionError(output_path, "save")
                
                return True, f"Successfully exported {len(studies)} studies to {output_path}", output_path
                
        except PermissionError as e:
            from error_handler import FilePermissionError
            if isinstance(e, FilePermissionError):
                return False, str(e), None
            return False, f"Export error: File is open in another program. Please close it and try again.", None
        except Exception as e:
            return False, f"Export error: {str(e)}", None
    
    def duplicate_site(
        self,
        site_id: int,
        new_name: str,
        user_id: Optional[str] = None,
    ) -> Tuple[bool, str, Optional[int]]:
        """
        Duplicate a site with all its studies.
        
        Returns: (success, message, new_site_id)
        """
        user_id = user_id or self.config.get_user_id()
        
        try:
            with DatabaseManager(user_id=user_id, config=self.config) as db:
                # Verify source site
                source_site = db.get_site(site_id)
                if not source_site:
                    return False, f"Site with ID {site_id} not found or access denied", None
                
                # Check if new name already exists
                existing_sites = db.get_sites()
                if any(s.name.lower() == new_name.lower() for s in existing_sites):
                    return False, f"Site '{new_name}' already exists", None
                
                # Get source studies
                studies = db.get_studies(site_id)
                
                # Create new site
                new_site = db.create_site(new_name)
                
                # Copy studies (clear IDs)
                studies_to_add = []
                for study in studies:
                    new_study = Study(
                        phase=study.phase,
                        subcategory=study.subcategory,
                        year=study.year,
                        sponsor=study.sponsor,
                        protocol=study.protocol,
                        description_full=study.description_full,
                        description_masked=study.description_masked,
                    )
                    studies_to_add.append(new_study)
                
                count = db.bulk_add_studies(new_site.id, studies_to_add)
                
                return True, f"Successfully duplicated site with {count} studies", new_site.id
                
        except Exception as e:
            return False, f"Duplication error: {str(e)}", None
    
    def merge_sites(
        self,
        source_site_ids: List[int],
        target_name: str,
        user_id: Optional[str] = None,
    ) -> Tuple[bool, str, Optional[int]]:
        """
        Merge multiple sites into a new site.
        
        Returns: (success, message, new_site_id)
        """
        user_id = user_id or self.config.get_user_id()
        
        if not source_site_ids:
            return False, "No source sites specified", None
        
        try:
            with DatabaseManager(user_id=user_id, config=self.config) as db:
                # Verify all source sites
                all_studies = []
                for sid in source_site_ids:
                    site = db.get_site(sid)
                    if not site:
                        return False, f"Site with ID {sid} not found or access denied", None
                    studies = db.get_studies(sid)
                    all_studies.extend(studies)
                
                if not all_studies:
                    return False, "No studies found in source sites", None
                
                # Check if target name exists
                existing_sites = db.get_sites()
                if any(s.name.lower() == target_name.lower() for s in existing_sites):
                    return False, f"Site '{target_name}' already exists", None
                
                # Create new site
                new_site = db.create_site(target_name)
                
                # Deduplicate studies based on identity
                seen_identities = set()
                studies_to_add = []
                
                for study in all_studies:
                    identity = study.get_identity_tuple()
                    if identity not in seen_identities:
                        new_study = Study(
                            phase=study.phase,
                            subcategory=study.subcategory,
                            year=study.year,
                            sponsor=study.sponsor,
                            protocol=study.protocol,
                            description_full=study.description_full,
                            description_masked=study.description_masked,
                        )
                        studies_to_add.append(new_study)
                        seen_identities.add(identity)
                
                count = db.bulk_add_studies(new_site.id, studies_to_add)
                
                return True, f"Successfully merged {len(source_site_ids)} sites into '{target_name}' with {count} unique studies", new_site.id
                
        except Exception as e:
            return False, f"Merge error: {str(e)}", None
