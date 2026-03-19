# CV Research Experience Manager
An **offline-only** desktop application for managing the Research Experience section of your CV document using master study lists. Written in Python 3.8+, it runs on Windows, macOS, and Linux with both a GUI and CLI interface.

---

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [How It Works](#how-it-works)
- [Architecture](#architecture)
- [File Formats](#file-formats)
- [Data Storage](#data-storage)
- [Configuration](#configuration)
- [Output Files](#output-files)
- [Formatting Rules](#formatting-rules)
- [Normalization](#normalization)
- [Testing](#testing)
- [Packaging](#packaging)
- [Error Handling](#error-handling)
- [Support](#support)

---

## Features

### Mode A: Update/Inject
- Parse existing Research Experience studies from your CV (.docx)
- Inject new studies from a master list (.xlsx or database site) above a benchmark year
- Protocols displayed in **bold red**
- Prevents duplicates (idempotent operation)
- Creates timestamped backup before modifications

### Mode B: Redact Protocols
- Match CV studies against master list using fuzzy matching
- Replace with masked versions (no protocols, treatments as XXX)
- Maintains document hierarchy and sorting
- Logs all matched/unmatched entries

### Mode C: Database Management
- **Per-user private** site databases (SQLite with WAL mode)
- Import `.xlsx` master lists into named sites
- Export sites back to `.xlsx`
- Full CRUD for phases, subcategories, and studies
- Automatic versioned backups with configurable retention

---

## Installation

### Requirements
- Python 3.8 or higher
- Windows, macOS, or Linux

### Setup

```bash
# Clone or download this repository
cd "Curriculum Vitae Manager"

# Install dependencies
pip install -r requirements.txt
```

### Dependencies
| Package | Purpose |
|---------|---------|
| `python-docx` | Read/write Word .docx documents |
| `openpyxl` | Read/write Excel .xlsx spreadsheets |
| `rapidfuzz` | Fuzzy text matching for study comparison |
| `pytest` | Test suite (development only) |

---

## Usage

### GUI Mode (Default)
```bash
py src/main.py
```
Opens the tkinter-based GUI with three tabs: **Update/Inject**, **Redact Protocols**, and **Database Management**.

### CLI Mode
```bash
# Validate master list (text output)
py src/main.py --mode validate-master --master "data/Master study list.xlsx"

# Validate master list (JSON output)
py src/main.py --mode validate-master --master "data/Master study list.xlsx" --json

# Validate CV
py src/main.py --mode validate-cv --cv "my_cv.docx"

# Preview update (JSON output for automation)
py src/main.py --mode update --cv "my_cv.docx" --master "studies.xlsx" --preview --json

# Preview redact
py src/main.py --mode redact --cv "my_cv.docx" --site 1 --preview

# Update/Inject
py src/main.py --mode update --cv "my_cv.docx" --master "studies.xlsx"

# Update/Inject without re-sorting existing studies
py src/main.py --mode update --cv "my_cv.docx" --master "studies.xlsx" --no-sort-existing

# Redact protocols
py src/main.py --mode redact --cv "my_cv.docx" --master "studies.xlsx"

# Import 7-column xlsx to DB
# Columns: Phase, Subcategory, Year, Sponsor, Protocol, Masked Description, Full Description
py src/main.py --mode import --master "studies.xlsx" --site-name "My Site"

# Export site to 7-column xlsx
py src/main.py --mode export --site 1 --output "exported.xlsx"

# List sites
py src/main.py --mode list-sites

# Run database migration
py src/main.py --mode migrate
```

The `--json` flag is available on `validate-master`, `validate-cv`, and `--preview` to produce structured JSON output suitable for automation.

The `--sort-existing` / `--no-sort-existing` flags control whether existing CV studies are re-sorted during Update/Inject. By default (`--sort-existing`), all studies are sorted. With `--no-sort-existing`, only newly inserted studies are sorted among themselves; the relative order of pre-existing CV studies is preserved.

---

## How It Works

### Startup Sequence

When you run `main.py`, the following happens in order:

1. **Dependency check** — verifies `python-docx`, `openpyxl`, and `rapidfuzz` are installed.
2. **Writability check** — confirms the `./data/` directory is writable.
3. **Config load** — loads `./data/config.json` (or creates defaults). Config is validated with type checks; invalid values fail fast with actionable messages.
5. **Permissions enforcement** — sets owner-only permissions on the user's data directory (`./data/users/{username}/`).
6. **Backup pruning** — removes backup files older than `backup_retention_days` (default: 90).
7. **Dispatch** — if command-line arguments are present, runs the CLI handler; otherwise, launches the tkinter GUI.

### Mode A: Update/Inject Flow

```
User selects CV (.docx) + Master source (.xlsx or DB site)
        │
        ▼
  ┌─────────────┐     ┌──────────────┐
  │ docx_handler│     │ excel_parser │  ← or database.py if using a site
  │  parse CV   │     │ parse master │
  └──────┬──────┘     └──────┬───────┘
         │                   │
         ▼                   ▼
  ┌─────────────────────────────────┐
  │         processor.py            │
  │  1. Parse CV → ResearchExp      │
  │  2. Parse master → Study list   │
  │  3. Calculate benchmark year    │
  │  4. For each master study:      │
  │     - normalizer normalizes     │
  │     - exact_match / fuzzy_match │
  │     - Skip if duplicate         │
  │     - Mark as "insert" if new   │
  │  5. Inject new studies          │
  │  6. Sort all (phases, subcats)  │
  └──────┬──────────────────────────┘
         │
         ▼
  ┌─────────────┐     ┌──────────┐
  │ docx_handler│     │ logger   │
  │ write output│     │ JSON/CSV │
  └─────────────┘     └──────────┘
```

**Key decisions**:
- The **benchmark year** determines which studies to inject. If the CV already has ≥4 studies in the latest year, only studies from that year and newer are injected. If fewer, it steps back one year.
- **Duplicate detection** uses a canonical identity tuple: `(phase, subcategory, year, sponsor, protocol, description_masked)`. Both exact and fuzzy matching are applied.
- The output document preserves all content outside the Research Experience section untouched.

### Mode B: Redact Protocols Flow

```
User selects CV (.docx) + Master source
        │
        ▼
  ┌─────────────┐     ┌──────────────┐
  │ docx_handler│     │ excel_parser │
  │  parse CV   │     │ parse master │
  └──────┬──────┘     └──────┬───────┘
         │                   │
         ▼                   ▼
  ┌─────────────────────────────────┐
  │         processor.py            │
  │  1. Parse CV → ResearchExp      │
  │  2. Parse master → Study list   │
  │  3. For each CV study:          │
  │     a. Skip if already masked   │
  │     b. Skip if no protocol token│
  │     c. match_study_to_master    │
  │     d. If matched: replace para │
  │        runs with masked content │
  │  4. (Optional) Re-sort affected │
  │     subcategories only          │
  └──────┬──────────────────────────┘
         │
         ▼
  ┌─────────────┐     ┌──────────┐
  │ docx_handler│     │ logger   │
  │ save in     │     │ JSON/CSV │
  │ place       │     │ (logs/)  │
  └─────────────┘     └──────────┘
```

**Key decisions**:
- **Replace-only by default**: Mode B redacts only studies that contain a protocol token (detected via normalized text, not font colour). Non-protocol studies and already-masked studies are never touched. Paragraph positions, headings, and surrounding structure are preserved.
- **Protocol detection**: Uses `contains_protocol_token()` — NFC normalization, casefold, whitespace collapse, dash canonicalization, then `extract_protocol()` with a sponsor-prefix cross-check. Does not rely on red font or bold styling.
- **Idempotency**: Already-masked lines (containing XXX and no protocol token) are detected by `is_already_masked()` and skipped. A second run produces no changes.
- **GUI option — "Sort and format studies"**: A checkbox in the Mode B panel (default unchecked). When **unchecked**, studies are replaced in place without changing order or formatting. When **checked**, subcategories that received at least one replacement are re-sorted (year desc → sponsor → protocol) after redaction. Subcategories with no replacements are untouched. No categories are moved between phases.
- **Preview**: `--preview` mode lists anchor paragraph indices and the resolved masked text for each planned replacement, indicates which categories would be re-sorted, and does not modify documents.
- **Logging**: Operations logged as `replaced`, `skipped-no-protocol`, `skipped-already-masked`, and `sort-category`. The `sort_and_format` flag state is recorded in the `config` log entry. Logs are written to the canonical logs directory only — no JSON/CSV in the result folder.

### Mode C: Database Management Flow

```
  ┌──────────────┐      ┌────────────────┐
  │ excel_parser │ ───▶ │ import_export  │
  │ parse .xlsx  │      │ import_xlsx_    │
  └──────────────┘      │ to_site()       │
                         └───────┬────────┘
                                 │
                                 ▼
                        ┌────────────────┐
                        │   database.py  │
                        │  SQLite CRUD   │
                        │  sites.db      │
                        └────────────────┘
                                 │
                                 ▼
                        ┌────────────────┐
                        │ import_export  │
                        │ export_site_   │
                        │ to_xlsx()      │
                        └───────┬────────┘
                                │
                                ▼
                        ┌────────────────┐
                        │  excel_parser  │
                        │  write .xlsx   │
                        └────────────────┘
```

The database stores sites as named collections of studies. Each site can be used as a master source for Mode A or Mode B instead of an .xlsx file.

---

## Architecture

### File Map

All source modules live in the `src/` subdirectory. Root-level files are limited to the launcher, config, and documentation.

```
Curriculum Vitae/
│
├── CV_Manager.bat             Windows launcher (installs deps + runs app)
├── README.md                  This file
├── requirements.txt           Python dependencies
├── cv_manager.spec            PyInstaller build spec for single-file executable
│
├── src/                       All application source code
│   ├── __init__.py            Package marker with version
│   ├── main.py                Entry point — startup checks, CLI dispatch, GUI launch
│   ├── gui.py                 Tkinter GUI (3 tabs: Update, Redact, Database)
│   ├── processor.py           Core logic for Mode A and Mode B operations
│   ├── docx_handler.py        Read/write CV .docx (find section, parse, format, save)
│   ├── excel_parser.py        Read/write master .xlsx (parse hierarchy, export)
│   ├── database.py            SQLite layer — CRUD for sites and studies
│   ├── import_export.py       Import .xlsx → DB site, export DB site → .xlsx
│   ├── normalizer.py          Text normalization and fuzzy/exact matching
│   ├── models.py              Dataclasses: Study, Phase, Subcategory, ResearchExperience, etc.
│   ├── config.py              AppConfig dataclass, load/save/validate config.json
│   ├── logger.py              Structured logging (JSON + CSV) for operations
│   ├── error_handler.py       Custom FilePermissionError and decorator
│   ├── progress_dialog.py     Tkinter spinner dialog for long-running GUI tasks
│   │
│   ├── offline_guard.py       Offline enforcement: proxy check, module scan, socket block
│   ├── validators.py          Deep validators for master .xlsx and CV .docx
│   ├── migrations.py          SQLite schema versioning, auto-migrate, rollback
│   ├── permissions.py         Directory permissions, log sanitization, backup pruning
│   │
│   ├── benchmark.py           Performance micro-benchmark script
│   ├── create_samples.py      Generate sample CV and master files for testing
│   ├── launcher.pyw           Windows GUI launcher (no console window)
│   │
│   └── tests/                 Test suite (200 tests)
│       ├── conftest.py        Shared fixtures (synthetic .docx, .xlsx, configs)
│       ├── test_normalizer.py Normalization and matching tests
│       ├── test_models.py     Dataclass behavior, sorting, dedup tests
│       ├── test_excel_parser.py  Parse/export/validate .xlsx tests
│       ├── test_database.py   SQLite CRUD, ownership, backup tests
│       ├── test_docx_handler.py  Section finding, parsing, writing, edge cases
│       ├── test_validators.py Deep validation tests for master and CV
│       ├── test_offline_guard.py  Proxy, module scan, socket guard tests
│       ├── test_migrations.py Schema version, migrate, rollback tests
│       ├── test_permissions.py  Sanitization, pruning, permission tests
│       ├── test_config.py     Config defaults, save/load, validation tests
│       └── test_integration.py  End-to-end Mode A, B, preview, import/export
│
└── data/                      Local data directory (created at runtime)
```

### How the Modules Connect

```
                        ┌──────────────┐
                        │   main.py    │
                        │  (entry pt)  │
                        └──────┬───────┘
                               │
              ┌────────────────┼─────────────────┐
              │                │                 │
              ▼                ▼                 ▼
        ┌──────────┐   ┌─────────────┐   ┌──────────────┐
        │  gui.py  │   │ CLI handler │   │ config.py    │
        │ (tkinter)│   │ (in main)   │   │ (settings)   │
        └────┬─────┘   └──────┬──────┘   └──────────────┘
             │                │
             └───────┬────────┘
                     │
        ┌────────────┼────────────────┐
        │            │                │
        ▼            ▼                ▼
  ┌───────────┐ ┌──────────┐  ┌──────────────┐
  │processor  │ │import_   │  │ validators   │
  │.py        │ │export.py │  │ .py          │
  │(Mode A/B) │ │(Mode C)  │  │(validate-*)  │
  └─────┬─────┘ └────┬─────┘  └──────────────┘
        │             │
   ┌────┼────┐   ┌────┼────┐
   │    │    │   │    │    │
   ▼    ▼    ▼   ▼    ▼    ▼
┌─────┐┌────┐┌────┐┌────┐┌──────┐
│docx_││excl││norm││data││logger│
│handl││_par││aliz││base││.py   │
│er   ││ser ││er  ││.py ││      │
└──┬──┘└──┬─┘└──┬─┘└──┬─┘└──────┘
   │      │     │     │
   │      │     │     ▼
   │      │     │  ┌──────────┐
   │      │     │  │migrations│
   │      │     │  │.py       │
   │      │     │  └──────────┘
   ▼      ▼     ▼
┌────────────────────────────────┐
│           models.py            │
│  Study, Phase, Subcategory,    │
│  ResearchExperience, Site, etc │
└────────────────────────────────┘
```

### Module Responsibilities

| Module | Depends On | Used By | What It Does |
|--------|-----------|---------|--------------|
| **main.py** | config, offline_guard, permissions, processor, import_export, database, validators, gui | — | Entry point. Runs startup checks, enforces offline mode, dispatches to GUI or CLI |
| **gui.py** | processor, database, import_export, config, models, normalizer, progress_dialog, error_handler | main.py | Tkinter GUI with three tabbed modes. Runs operations in background threads with a spinner dialog |
| **processor.py** | docx_handler, excel_parser, database, normalizer, models, logger, config | main.py, gui.py | Orchestrates Mode A (update/inject) and Mode B (redact). Parses CV and master data, performs matching, injects or redacts, produces result logs |
| **docx_handler.py** | models, normalizer, config | processor.py, validators.py | Opens .docx files via `python-docx`. Finds the "Research Experience" section, parses it into `ResearchExperience` objects, and writes structured studies back with correct formatting (font, bold, red, hanging indent) |
| **excel_parser.py** | models, normalizer | processor.py, import_export.py, validators.py | Parses master .xlsx files into `Study` lists by reading the Column A hierarchy (Phase → Subcategory → Year) and Columns B/C for descriptions. Also exports studies back to .xlsx |
| **database.py** | models, config, migrations | processor.py, import_export.py, gui.py | SQLite database manager. Handles all CRUD for sites and studies, with per-user isolation, WAL journaling, foreign key enforcement, and automatic schema migration |
| **import_export.py** | database, excel_parser, models, config | main.py, gui.py | Bridges .xlsx files and the database. Imports master lists into named sites, exports sites back to .xlsx, handles duplicates and merge logic |
| **normalizer.py** | rapidfuzz | processor.py, docx_handler.py, excel_parser.py, validators.py, gui.py | All text normalization (Unicode NFC, whitespace, dashes, quotes, colon spacing, phase names, X-run collapse) and matching (exact, fuzzy with configurable thresholds, protocol extraction, study line parsing) |
| **models.py** | — | Nearly all modules | Core dataclasses: `Study`, `Subcategory`, `Phase`, `ResearchExperience`, `Site`, `SiteVersion`, `LogEntry`, `OperationResult`. Defines identity tuples for dedup, sorting logic, and benchmark year calculation |
| **config.py** | — | Nearly all modules | `AppConfig` dataclass with all settings. Loads/saves `config.json`, validates types and ranges on load, provides user-specific paths, enforces `network_enabled = false` |
| **logger.py** | models | processor.py | `OperationLogger` that records every decision (inserted, matched, skipped, replaced) to JSON and CSV files with timestamps, hierarchy context, and match scores |
| **error_handler.py** | — | gui.py, docx_handler.py | Custom `FilePermissionError` exception and a `@handle_file_operation` decorator that converts raw `PermissionError` into user-friendly messages (e.g., "file is open in Word") |
| **progress_dialog.py** | tkinter | gui.py | Modal spinner dialog for long-running GUI operations. Runs the actual work in a background thread |
| **offline_guard.py** | — | main.py | Startup self-check: scans for proxy env vars, checks `sys.modules` for disallowed network packages, monkeypatches `socket.connect` to block all connections. Controlled by `config.offline_guard_enabled` |
| **validators.py** | normalizer, openpyxl, python-docx | main.py | Deep structural validators returning JSON reports. Master validator checks hierarchy, duplicates, empty cells, formulas. CV validator checks section presence, font, bold/red protocol styling, hanging indent |
| **migrations.py** | — | database.py, main.py | Schema version table (`schema_info`), sequential migration definitions, auto-migrate with pre-migration backups, rollback support |
| **permissions.py** | — | main.py | Sets owner-only permissions on user directories (chmod 700 on Unix, logs icacls guidance on Windows). Sanitizes log text in Redact mode. Prunes old backups by retention policy |
| **benchmark.py** | models, normalizer, excel_parser, database, config | — (standalone) | Micro-benchmark measuring parse, normalize, fuzzy match, and DB insert/read throughput at configurable scale |
| **create_samples.py** | python-docx, openpyxl | — (standalone) | Generates sample `sample_cv.docx` and `sample_master.xlsx` files in `./samples/` for manual testing |

---

## File Formats

### CV Document (.docx)
- Must contain a **"Research Experience"** heading (Heading 1 style)
- Studies formatted as: `{Year}<TAB>{Sponsor} {Protocol}: {Description}`
- Hierarchy: Phase heading → Subcategory heading → Study entries

### Master List (.xlsx) — 7-Column Format

The current import/export schema uses **7 explicit columns** with a mandatory header row:

| Phase | Subcategory | Year | Sponsor | Protocol | Masked Description | Full Description |
|-------|-------------|------|---------|----------|--------------------|------------------|
| Phase I | Oncology | 2024 | Pfizer | PF-123 | Pfizer: A study of XXX in lung cancer | Pfizer PF-123: A study of PF-123 (pembro) in lung cancer |
| Phase I | Cardiology | 2023 | Novartis | NVS-456 | Novartis: Another study with XXX | Novartis NVS-456: Another study with NVS-456 |
| Phase II–IV | Oncology | 2024 | Roche | RO-777 | Roche: Phase 3 XXX vs placebo | Roche RO-777: Phase 3 atezolizumab vs placebo |

- **Row 1 must be the header**: `Phase, Subcategory, Year, Sponsor, Protocol, Masked Description, Full Description`
- **Year**: Integer (1900–2100) or 0 for unknown
- **Legacy 3-column format**: No longer accepted for import. Export to 7-column first

> **Migration note**: If you have legacy 3-column `.xlsx` files, open them in the GUI (Mode C → Export), which will re-export in the 7-column format. Then re-import the new file.

---

## Data Storage

All data is stored locally under `./data/`:

```
./data/
  config.json              # Global configuration
  tmp/                     # Temporary files (auto-cleaned)
  users/
    {username}/
      sites.db             # SQLite database (WAL mode)
      exports/             # Exported .xlsx files
      imports/             # Imported .xlsx copies
      backups/             # Timestamped backups (auto-pruned)
      logs/                # Operation logs (JSON/CSV)
      config.json          # User-specific config
```

---

## Configuration

Edit `./data/config.json` to customize:

```json
{
  "fuzzy_threshold_full": 92,
  "fuzzy_threshold_masked": 90,
  "benchmark_min_count": 4,
  "highlight_inserted": false,
  "backup_retention_days": 90,
  "log_retention_days": 90,
  "uncategorized_label": "Uncategorized",
  "data_root": "./data",
  "font_name": "Calibri",
  "font_size": 11
}
```

### Key Settings
| Setting | Default | Description |
|---------|---------|-------------|
| `fuzzy_threshold_full` | 92 | Minimum match score for full descriptions |
| `fuzzy_threshold_masked` | 90 | Minimum match score for masked descriptions |
| `benchmark_min_count` | 4 | If ≤3 studies in latest year, benchmark = latest - 1 |
| `backup_retention_days` | 90 | Auto-delete backups older than this many days |
| `log_retention_days` | 90 | Auto-delete log files older than this many days |
| `highlight_inserted` | false | When true, newly injected studies are highlighted in yellow in the output .docx |
| `uncategorized_label` | Uncategorized | Display label for studies that don't match any master category. Must be non-empty |
| `font_name` | Calibri | Font family for output .docx. Allowed: Calibri, Times New Roman, Garamond, Helvetica, Roboto, Open Sans, Lato, Didot |
| `font_size` | 11 | Font size in points (6–72) |

Config is **validated on load** — invalid types or out-of-range values cause a fast failure with an actionable error message.

All settings are also accessible via the **Configuration → Settings** menu in the GUI, which provides a professional settings panel with Save and Reset to Defaults buttons. Each setting has a **tooltip icon** (ⓘ) that explains what the setting does and how changing it affects the program. Tooltips appear on hover, focus, or click, and are keyboard-accessible.

### Automatic Category Order Maintenance

When a new Phase or Subcategory is added to the database — via Mode C CRUD, Import, or Mode A processing — the **Category Order** is automatically updated to include the new entry at the end. This ensures that:

- Newly created categories appear immediately in Mode A and Mode B sorting without requiring a restart or manual save.
- Existing order is preserved; only genuinely new entries are appended.
- The operation is idempotent: adding the same category again does not create duplicates or shift existing indexes.
- If `phase_order` is configured in `config.json`, it remains authoritative for Phase-level sorting; the database order governs subcategory ordering within each phase.
- Deletion of a study does not remove its category from the order table, preserving historical ordering.

The **"Enable sorting for existing studies"** option is also available as a checkbox in the Update/Inject tab's Options panel. When unchecked, only newly inserted studies are sorted; the relative order of pre-existing CV studies is preserved.

### Robust Phase/Subcategory Matching

Phase and Subcategory headings in the CV are matched to master data using **normalized keys only** — the original heading text in the document is never altered.

Normalization for matching applies:
- **Unicode NFC** normalization
- **Casefold** (case-insensitive comparison)
- **Whitespace collapse** (multiple spaces/tabs → single space, trimmed)
- **Dash canonicalization** (en-dash, em-dash, and other Unicode dashes → hyphen)
- **Quote canonicalization** (smart/curly quotes → straight quotes)
- **Roman numeral equivalence** for phases via `PHASE_SYNONYMS` mapping (e.g. `"Phase 1"` ↔ `"Phase I"`, `"PHASE I"` ↔ `"Phase I"`, `"Phase II-IV"` ↔ `"Phase II–IV"`)

This means a CV with `"PHASE I"` and a master with `"Phase I"` will correctly match to the **same** container — no duplicate Phase or Subcategory blocks are created. Existing headings and their formatting are preserved as-is; only the matching key is normalized.

The `PHASE_SYNONYMS` constant in `normalizer.py` is authoritative for all recognized phase name variants.

### Style-Agnostic Heading Detection

Phase and Subcategory headings are detected **by text content**, regardless of the Word style applied (Normal, Heading 1, Heading 2, Heading 3, etc.). A heading-styled paragraph is treated as a section boundary **only** when its text matches a known major CV section name (e.g. "Education", "Publications"). Phase headings like "Phase I" and subcategory headings like "Healthy Adults" styled with Word heading styles are kept inside the Research Experience section.

### Preserve-Existing Mode (Sort Disabled)

When **"Enable sorting for existing studies"** is unchecked (or `--no-sort-existing` on the CLI), the program:
- **Preserves all original XML formatting**: indentation, tabs, spacing, run boundaries, font styles, bold/color properties — existing paragraph elements are moved, never recreated
- **Subcategories receiving new studies**: the combined list (existing + new) is sorted by year descending, sponsor, and protocol. Existing paragraph elements are reordered but their formatting is preserved
- **Subcategories with no new studies**: completely untouched — no reordering, no reformatting
- **Empty subcategories**: new studies are inserted immediately after the subcategory heading
- Creates new Phase/Subcategory heading paragraphs only when the target container does not yet exist in the document
- Maintains idempotency: subsequent runs detect already-inserted studies and do not add duplicates

### Subcategory Detection

Subcategory headings are distinguished from sponsor headings using a **look-ahead heuristic**: if the next non-empty paragraph after a short text line is a year-line (starts with a 4-digit year), the line is treated as a subcategory heading. This prevents subcategory names like "Healthy Adults" or "Vaccine" from being misclassified as sponsor company names. The look-ahead is combined with:
- **Force-subcategory rule**: immediately after a phase heading, the first non-year, non-phase line is always a subcategory
- **Sponsor keyword detection**: lines containing INC, LLC, CORP, PHARMA, etc. are treated as sponsor headings only when the next line is NOT a year-line

---

## Output Files

All output files are written into a **per-CV result folder** at the project root:

```
result/
  <Original CV Name>/
    <Original CV Name> (Updated YYYY-MM-DD).docx
    <Original CV Name> (Redacted YYYY-MM-DD).docx
    <Site Name> - Master Study List.xlsx     (Mode C export)
```

The result folder contains **only document files** (.docx, .xlsx). No JSON or CSV log files are placed in the result folder.

The original CV name is determined by:
1. Reading the custom document property `_original_cv_name` (set automatically by Mode A)
2. Falling back to stripping date-stamped suffixes like `(Updated YYYY-MM-DD)` or `(Redacted YYYY-MM-DD)` from the input filename

This means Mode B processing the output of Mode A writes into the **same** folder.

If an explicit `--output` path is provided, that path is used as-is (no subfolder routing).

### Logs
- JSON format: `{operation}_{timestamp}.json`
- CSV format: `{operation}_{timestamp}.csv`
- Logs are saved to the canonical logs directory only: `./data/users/{username}/logs/`

Contains: operation type (inserted, replaced, skipped-duplicate, etc.), phase, subcategory, year, sponsor, protocol, match scores, and details.

---

## Formatting Rules

### Study Display Format
```
{Year}<TAB>{Sponsor}{[ SPACE ]{Protocol}}: {Description}
```

### Typography
- **Year**: Not bold
- **Sponsor**: Bold
- **Protocol**: Bold + Red (Mode A only; removed in Mode B)
- **Font**: Calibri 11pt
- **Paragraph**: Left indent 0", hanging indent 0.5"

### Sorting Order
1. **Phases**: Phase I first, then Phase II–IV, then Uncategorized
2. **Subcategories**: Alphabetical within each phase
3. **Studies**: Year descending → Sponsor ascending → Protocol ascending

---

## Normalization

Text is normalized for matching (via `normalizer.py`):
- **Unicode**: NFC normalization applied first
- **Case**: Lowercased
- **Whitespace**: Tabs and multiple spaces collapsed to single space
- **Dashes**: `–`, `—`, `−` unified to `-`
- **Quotes**: Curly quotes (`'`, `'`, `"`, `"`) straightened
- **Colons**: Spacing canonicalized to `{word}: {word}`
- **Phases**: `Phase 1` → `Phase I`, `Phase 2-4` → `Phase II–IV`
- **X runs**: `XXXXXX` collapsed to `XXX` (for matching only, never in saved output)

---

## Testing

The project includes a comprehensive test suite with 460 tests:

```bash
# Run all tests
py -m pytest src/tests/ -v

# Run a specific module
py -m pytest src/tests/test_normalizer.py -v

# Run integration tests only
py -m pytest src/tests/test_integration.py -v

# Run with coverage (requires pytest-cov)
py -m pytest src/tests/ --cov=src --cov-report=term-missing
```

All tests are **hermetic** — they use synthetic data generated on the fly (no external files needed) and run in isolated temp directories.

### Benchmarking

```bash
py src/benchmark.py --count 1000
py src/benchmark.py --count 5000
py src/benchmark.py --count 10000
```

---

## Packaging

### Prerequisites

- Python 3.8+ (64-bit) on Windows
- Project dependencies: `pip install -r requirements.txt`
- PyInstaller: `pip install pyinstaller`

### Build Commands

**One-file build** (default — single `CV_Manager.exe` at project root):

```batch
build\build_win.bat
```
```powershell
.\build\build_win.ps1
```

**One-folder build** (exe + supporting files in `CV_Manager/`):

```batch
build\build_win.bat onedir
```
```powershell
.\build\build_win.ps1 -BuildMode onedir
```

**Console build** (visible console for debugging):

```powershell
.\build\build_win.ps1 -Console
```

The build scripts handle everything automatically: dependency install, build-number bump, icon generation, artifact cleanup, PyInstaller `--clean`, and shortcut creation.

The backward-compatible entry points still work:
```powershell
.\build\build.ps1                # delegates to build_win.ps1
.\CV_Manager.bat                 # builds then launches
```

### Application Icon

The build embeds `build/assets/app.ico` into the `.exe`.

**To update the icon:**
1. Place your custom `.ico` file at `build/assets/feather.ico`
2. Rebuild — the build script copies `feather.ico` as `app.ico` automatically
3. If no `feather.ico` is present, a feather-pen icon is generated via Pillow

**If Windows still shows an old icon after rebuild:**
1. Right-click the pinned taskbar icon → "Unpin from taskbar"
2. Delete the old `CV_Manager.exe`
3. Rebuild with `build\build_win.bat`
4. Re-pin the new `CV_Manager.exe`

Every build bumps a version resource (`build/build_number.txt`) embedded in the `.exe`, which signals Windows to invalidate its icon cache. In rare cases, run `.\build\refresh_shell_cache.ps1` for manual cache-clearing instructions.

### Taskbar Pinning

One-file builds extract to a temp directory on each run, which can cause Windows to show a second taskbar icon when the app is pinned. This is addressed by:

- **AppUserModelID** — set via `src/appid.py` at process startup, ensuring pinned shortcuts and running windows share the same identity
- **Stable runtime tmpdir** — the one-file build uses `_cv_manager_runtime` as a fixed extraction directory so Windows sees a consistent child-process path
- **Shortcuts with embedded AppID** — `build/create_shortcut.py` creates Desktop and Start Menu shortcuts that carry the same `AppUserModelID`

**Recommendation:** One-folder builds (`-BuildMode onedir`) give the most stable pinning behavior because there is no temp-directory extraction involved.

### Data Storage

Both `CV_Manager.bat` and `CV_Manager.exe` share the same `data/` folder:

```
Curriculum Vitae/          ← project root
  CV_Manager.bat           ← source launcher
  CV_Manager.exe           ← built executable
  data/
    config.json            ← shared config
    users/{username}/
      sites.db             ← shared database
      logs/
      backups/
      exports/
      results/
  result/
    {CV Name}/
```

### Console Toggle (Debugging)

To build with a visible console window for debugging:
```powershell
.\build\build_win.ps1 -Console
```
Or:
```batch
set CONSOLE_MODE=1
build\build_win.bat
```

### SmartScreen Warning

The built `.exe` is unsigned. On first launch, Windows SmartScreen may show "Windows protected your PC." Click **More info → Run anyway**. To avoid this, code-sign the executable with a certificate:
```powershell
signtool sign /f cert.pfx /p PASSWORD /t http://timestamp.digicert.com CV_Manager.exe
```

### Smoke Test

After building, verify the build artifacts:
```powershell
py scripts/smoke_build_check.py
py scripts/smoke_build_check.py --launch
```

This checks: exe exists, icon exists and is non-empty, version resource matches the current build number, and (with `--launch`) the exe runs and exits cleanly.

### Single-Instance Lock

The application prevents multiple GUI instances from writing to the same database simultaneously. If a second instance is launched, it shows a warning dialog and exits. The lock file is stored at `./data/users/{username}/.cv_manager.lock`.

### Font Note

Calibri is bundled with Windows. On macOS/Linux, the app writes Calibri as the font name in `.docx` output — Word on the target machine handles font substitution if Calibri is unavailable.

---

## Error Handling

- **Missing "Research Experience"**: Fails with clear error (section must exist in the .docx)
- **Read-only location**: Error with guidance to select a writable folder
- **File locked by Word**: Detected and reported with a user-friendly message
- **No injectable studies**: Reports "No changes" (not an error)
- **No redaction matches**: Reports "No changes" (not an error)
- **Access denied**: Blocked operations logged to `access_denied.log`
- **Invalid config**: Fails fast on startup with specific messages about which settings are wrong

---

## Support

This is an offline, local application. For issues:
1. Check the logs in `./data/users/{username}/logs/`
2. Verify file formats match the specifications above
3. Ensure write permissions to the data directory
4. Run `py src/main.py --mode validate-master --master "file.xlsx"` to check your master list
5. Run `py src/main.py --mode validate-cv --cv "file.docx"` to check your CV port

Please let me know if there is any bug/issue or any feature request.