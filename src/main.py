"""
CV Research Experience Manager - Main Entry Point

An offline application for managing the Research Experience section of CV documents.

Modes:
- Mode A: Update/Inject - Add new studies from master list
- Mode B: Redact Protocols - Remove protocols and mask treatments
- Mode C: Database - Manage multi-site study databases

All operations are local-only with no network access.
"""

import sys
import os
import json
import logging
import atexit
from pathlib import Path

if getattr(sys, "frozen", False):
    _app_dir = Path(sys.executable).parent.resolve()
    os.chdir(str(_app_dir))
    _src_in_bundle = str(Path(sys._MEIPASS))
    if _src_in_bundle not in sys.path:
        sys.path.insert(0, _src_in_bundle)
else:
    _app_dir = Path(__file__).parent.resolve()
    if str(_app_dir) not in sys.path:
        sys.path.insert(0, str(_app_dir))

from config import get_config, AppConfig, APP_VERSION, APP_NAME


def check_dependencies():
    """Check that all required dependencies are installed."""
    missing = []
    
    try:
        import docx
    except ImportError:
        missing.append("python-docx")
    
    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")
    
    try:
        import rapidfuzz
    except ImportError:
        missing.append("rapidfuzz")
    
    if missing:
        print("Missing required dependencies:")
        for dep in missing:
            print(f"  - {dep}")
        print("\nPlease install them with:")
        print(f"  pip install {' '.join(missing)}")
        return False
    
    return True


def check_writable():
    """Check that the data directory is writable."""
    config = get_config()
    data_path = config.data_path
    
    try:
        data_path.mkdir(parents=True, exist_ok=True)
        test_file = data_path / ".write_test"
        test_file.write_text("test")
        test_file.unlink()
        return True
    except (OSError, PermissionError) as e:
        print(f"Error: Cannot write to data directory: {data_path}")
        print(f"  {e}")
        print("\nPlease ensure the application has write permissions,")
        print("or configure a different data_root in ./data/config.json")
        return False


def run_gui():
    """Run the GUI application."""
    from gui import main
    main()


def run_cli():
    """Run CLI mode for scripting/automation."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="CV Research Experience Manager - CLI Mode",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Update/Inject mode with xlsx file
  python main.py --mode update --cv "my_cv.docx" --master "studies.xlsx"
  
  # Redact mode with saved site
  python main.py --mode redact --cv "my_cv.docx" --site 1
  
  # Import 7-column xlsx to database
  # Columns: Phase, Subcategory, Year, Sponsor, Protocol, Masked Description, Full Description
  python main.py --mode import --master "studies.xlsx" --site-name "My Site"
  
  # Export site to 7-column xlsx
  python main.py --mode export --site 1 --output "exported.xlsx"
  
  # List all sites
  python main.py --mode list-sites
  
  # Validate master list
  python main.py --mode validate-master --master "studies.xlsx"
  
  # Validate CV
  python main.py --mode validate-cv --cv "my_cv.docx"
  
  # Run database migration
  python main.py --mode migrate
"""
    )
    
    parser.add_argument(
        '--mode', '-m',
        choices=['update', 'redact', 'import', 'export', 'list-sites',
                 'validate-master', 'validate-cv', 'migrate', 'gui'],
        default='gui',
        help='Operation mode (default: gui)'
    )
    
    parser.add_argument(
        '--cv', '-c',
        help='Path to CV .docx file'
    )
    
    parser.add_argument(
        '--master', '-x',
        help='Path to master .xlsx file'
    )
    
    parser.add_argument(
        '--site', '-s',
        type=int,
        help='Site ID to use as master source'
    )
    
    parser.add_argument(
        '--site-name',
        help='Name for new site (import mode)'
    )
    
    parser.add_argument(
        '--output', '-o',
        help='Output file path'
    )
    
    parser.add_argument(
        '--preview',
        action='store_true',
        help='Preview changes without applying'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output structured JSON instead of plain text'
    )
    
    parser.add_argument(
        '--check-updates',
        action='store_true',
        help='One-time check for a newer release on GitHub (requires network)'
    )

    sort_group = parser.add_mutually_exclusive_group()
    sort_group.add_argument(
        '--sort-existing',
        action='store_true',
        default=None,
        dest='sort_existing',
        help='Sort all studies including existing ones (default behavior)'
    )
    sort_group.add_argument(
        '--no-sort-existing',
        action='store_false',
        dest='sort_existing',
        help='Only sort newly inserted studies; preserve existing CV order'
    )
    
    args = parser.parse_args()
    
    if args.check_updates:
        try:
            from update_checker import check_for_update, APP_VERSION
            print(f"Current version: {APP_VERSION}")
            print("Checking for updates...")
            info = check_for_update()
            if info:
                print(f"Newer version available: {info['tag_name']}")
                print(f"  Release page: {info['html_url']}")
            else:
                print("You are running the latest version.")
        except Exception as exc:
            print(f"Update check failed: {exc}")
        return

    if args.mode == 'gui':
        run_gui()
        return
    
    # CLI operations
    from processor import CVProcessor
    from import_export import ImportExportManager
    from database import DatabaseManager
    from config import get_config
    
    config = get_config()
    
    # --- Validate subcommands ---
    if args.mode == 'validate-master':
        if not args.master:
            print("Error: --master is required for validate-master mode")
            return
        from validators import validate_master_xlsx_strict
        report = validate_master_xlsx_strict(Path(args.master))
        if args.json:
            print(json.dumps(report, indent=2))
        else:
            status = "PASS" if report["valid"] else "FAIL"
            print(f"Master list validation: {status}")
            print(f"  Phases: {report['stats']['phases']}")
            print(f"  Subcategories: {report['stats']['subcategories']}")
            print(f"  Studies: {report['stats']['studies']}")
            if report["issues"]:
                print(f"  Issues ({len(report['issues'])}):\n")
                for issue in report["issues"]:
                    print(f"    [{issue['severity'].upper()}] Row {issue['row']}: {issue['message']}")
        sys.exit(0 if report["valid"] else 1)
    
    if args.mode == 'validate-cv':
        if not args.cv:
            print("Error: --cv is required for validate-cv mode")
            return
        from validators import validate_cv_docx_strict
        report = validate_cv_docx_strict(Path(args.cv))
        if args.json:
            print(json.dumps(report, indent=2))
        else:
            status = "PASS" if report["valid"] else "FAIL"
            print(f"CV validation: {status}")
            print(f"  Total paragraphs: {report['stats']['total_paragraphs']}")
            print(f"  Research Exp paragraphs: {report['stats']['research_exp_paragraphs']}")
            print(f"  Study lines: {report['stats']['study_lines']}")
            print(f"  Phase headings: {report['stats']['phase_headings']}")
            if report["issues"]:
                print(f"  Issues ({len(report['issues'])}):\n")
                for issue in report["issues"]:
                    print(f"    [{issue['severity'].upper()}] Line {issue['row']}: {issue['message']}")
        sys.exit(0 if report["valid"] else 1)
    
    if args.mode == 'migrate':
        from migrations import auto_migrate, get_schema_version, LATEST_VERSION, ensure_schema_info_table
        import sqlite3
        db_path = config.get_user_db_path()
        db_path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(db_path))
        conn.row_factory = sqlite3.Row
        ensure_schema_info_table(conn)
        current = get_schema_version(conn)
        print(f"Current schema version: {current}")
        print(f"Latest schema version: {LATEST_VERSION}")
        if current >= LATEST_VERSION:
            print("Database is up to date.")
        else:
            applied = auto_migrate(conn, db_path)
            for desc in applied:
                print(f"  Applied: {desc}")
            print(f"Migration complete. Now at version {get_schema_version(conn)}.")
        conn.close()
        return
    
    if args.mode == 'list-sites':
        with DatabaseManager(config=config) as db:
            sites = db.get_sites()
            if not sites:
                print("No sites found.")
            else:
                print(f"Found {len(sites)} site(s):\n")
                for site in sites:
                    count = db.get_study_count(site.id)
                    print(f"  [{site.id}] {site.name} - {count} studies")
                    print(f"      Created: {site.created_at}")
        return
    
    if args.mode == 'import':
        if not args.master:
            print("Error: --master is required for import mode")
            return
        if not args.site_name:
            print("Error: --site-name is required for import mode")
            return
        
        manager = ImportExportManager(config)
        success, message, site_id = manager.import_xlsx_to_site(
            Path(args.master),
            args.site_name,
            replace_existing=True
        )
        
        if success:
            print(f"Success: {message}")
            print(f"Site ID: {site_id}")
        else:
            print(f"Error: {message}")
        return
    
    if args.mode == 'export':
        if not args.site:
            print("Error: --site is required for export mode")
            return
        
        output_path = Path(args.output) if args.output else None
        
        manager = ImportExportManager(config)
        success, message, path = manager.export_site_to_xlsx(args.site, output_path)
        
        if success:
            print(f"Success: {message}")
        else:
            print(f"Error: {message}")
        return
    
    # Update/Redact modes
    if not args.cv:
        print("Error: --cv is required for update/redact modes")
        return
    
    if not args.master and not args.site:
        print("Error: --master or --site is required for update/redact modes")
        return
    
    processor = CVProcessor(config)
    cv_path = Path(args.cv)
    master_path = Path(args.master) if args.master else None
    output_path = Path(args.output) if args.output else None
    
    # Resolve enable_sort_existing: CLI flag > config default
    enable_sort_existing = args.sort_existing
    if enable_sort_existing is None:
        enable_sort_existing = config.enable_sort_existing
    logging.getLogger(__name__).info(
        "CLI: enable_sort_existing=%s (cli_flag=%s, config=%s)",
        enable_sort_existing,
        args.sort_existing,
        config.enable_sort_existing,
    )
    
    if args.preview:
        mode = "update_inject" if args.mode == "update" else "redact_protocols"
        changes, error = processor.preview_changes(
            cv_path, master_path, args.site, mode,
            enable_sort_existing=enable_sort_existing,
        )
        
        if error:
            if args.json:
                print(json.dumps({"error": error, "changes": []}, indent=2))
            else:
                print(f"Error: {error}")
        elif not changes:
            if args.json:
                print(json.dumps({"error": None, "changes": []}, indent=2))
            else:
                print("No changes to make.")
        else:
            if args.json:
                print(json.dumps({"error": None, "changes": changes}, indent=2))
            else:
                print(f"Found {len(changes)} changes:\n")
                for change in changes:
                    print(f"  \u2022 {change}")
        return
    
    if args.mode == 'update':
        result = processor.mode_a_update_inject(
            cv_path, master_path, args.site,
            output_path=output_path,
            enable_sort_existing=enable_sort_existing,
        )
    else:  # redact
        result = processor.mode_b_redact_protocols(cv_path, master_path, args.site, output_path)
    
    if result.success:
        print(f"Success! Output: {result.output_path}")
        counts = result.get_counts()
        print("\nSummary:")
        for op, count in counts.items():
            print(f"  {op}: {count}")
    else:
        print(f"Error: {result.error_message}")


def main():
    """Main entry point."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )

    logger = logging.getLogger(__name__)
    logger.info("%s v%s starting (frozen=%s)", APP_NAME, APP_VERSION, getattr(sys, "frozen", False))

    if not check_dependencies():
        sys.exit(1)

    if not check_writable():
        sys.exit(1)

    config = get_config()
    config.ensure_user_directories()

    if config.network_enabled:
        print("Warning: network_enabled was set to True. Forcing to False for offline operation.")
        config.network_enabled = False
        config.save()

    if config.offline_guard_enabled:
        try:
            from offline_guard import enforce_offline
            enforce_offline(fail_fast=False, block_sockets=True)
        except Exception as exc:
            logger.warning("Offline guard warning: %s", exc)

    try:
        from permissions import secure_user_directory, prune_user_backups, prune_user_logs
        user_path = config.get_user_data_path()
        secure_user_directory(user_path)
        prune_user_backups(user_path, config.backup_retention_days)
        prune_user_logs(user_path, config.log_retention_days)
    except Exception as exc:
        logger.warning("Permissions/pruning warning: %s", exc)

    is_gui = len(sys.argv) <= 1

    if is_gui:
        from instance_lock import acquire_instance_lock, release_instance_lock
        lock_dir = str(config.get_user_data_path())
        if not acquire_instance_lock(lock_dir):
            print("Another instance of CV Manager is already running.")
            try:
                import tkinter as _tk
                _r = _tk.Tk()
                _r.withdraw()
                from tkinter import messagebox as _mb
                _mb.showwarning(
                    APP_NAME,
                    "Another instance is already running.\n"
                    "Please close it before starting a new one.",
                )
                _r.destroy()
            except Exception:
                pass
            sys.exit(1)
        atexit.register(release_instance_lock)
        run_gui()
    else:
        run_cli()


if __name__ == "__main__":
    main()
