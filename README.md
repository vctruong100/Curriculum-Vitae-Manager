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
- [Security & Privacy](#security--privacy)
- [Output Files](#output-files)
- [Formatting Rules](#formatting-rules)
- [Normalization](#normalization)
- [Testing](#testing)
- [Packaging](#packaging)
- [Error Handling](#error-handling)
- [License](#license)

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
cd "Curriculum Vitae"

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

# Redact protocols
py src/main.py --mode redact --cv "my_cv.docx" --master "studies.xlsx"

# Import master to DB
py src/main.py --mode import --master "studies.xlsx" --site-name "My Site"

# Export site
py src/main.py --mode export --site 1 --output "exported.xlsx"

# List sites
py src/main.py --mode list-sites

# Run database migration
py src/main.py --mode migrate
```

The `--json` flag is available on `validate-master`, `validate-cv`, and `--preview` to produce structured JSON output suitable for automation.

---

## How It Works

### Startup Sequence

When you run `main.py`, the following happens in order:

1. **Dependency check** — verifies `python-docx`, `openpyxl`, and `rapidfuzz` are installed.
2. **Writability check** — confirms the `./data/` directory is writable.
3. **Config load** — loads `./data/config.json` (or creates defaults). Config is validated with type checks; invalid values fail fast with actionable messages.
4. **Offline guard** — if `offline_guard_enabled` is `true` (the default), scans for proxy environment variables, checks for disallowed network modules, and monkeypatches `socket.connect` to block all outbound connections.
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
  │     - match_study_to_master     │
  │     - If matched: replace with  │
  │       masked description        │
  │     - Protocol removed          │
  │     - Treatments → XXX          │
  │  4. Sort and rewrite            │
  └──────┬──────────────────────────┘
         │
         ▼
  ┌─────────────┐     ┌──────────┐
  │ docx_handler│     │ logger   │
  │ write output│     │ JSON/CSV │
  └─────────────┘     └──────────┘
```

**Key decisions**:
- Every CV study is compared against the master list using `normalizer.match_study_to_master()`, which tries exact matching first, then falls back to fuzzy matching with configurable thresholds.
- Matched studies have their full description replaced with the masked version from Column C of the master list.
- Unmatched studies are preserved as-is but logged for review.

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

### Master List (.xlsx)
| Column A | Column B | Column C |
|----------|----------|----------|
| Phase I | | |
| Oncology | | |
| 2024 | Pfizer PF-123: Full description with treatment | Pfizer: Masked description with XXX |
| 2023 | Novartis NVS-456: Another study | Novartis: Masked version |
| Phase II–IV | | |
| ... | ... | ... |

- **Column A**: Hierarchy stream (Phase row → Subcategory row → Year for studies)
- **Column B**: Full description with protocol and treatment names
- **Column C**: Masked description (no protocol, treatments replaced with XXX)

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
  "use_track_changes": false,
  "phase_order": ["Phase I", "Phase II–IV"],
  "network_enabled": false,
  "offline_guard_enabled": true,
  "backup_retention_days": 90,
  "log_retention_days": 90,
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
| `network_enabled` | false | **Always false** — app is offline-only |
| `offline_guard_enabled` | true | Block network sockets and scan for proxy env vars at startup |
| `backup_retention_days` | 90 | Auto-delete backups older than this many days |
| `log_retention_days` | 90 | Auto-delete log files older than this many days |
| `font_name` | Calibri | Font family for output .docx. Allowed: Calibri, Times New Roman, Garamond, Helvetica, Roboto, Open Sans, Lato, Didot |
| `font_size` | 11 | Font size in points (6–72) |

Config is **validated on load** — invalid types or out-of-range values cause a fast failure with an actionable error message.

All settings are also accessible via the **Configuration → Settings** menu in the GUI, which provides a professional settings panel with Save and Reset to Defaults buttons.

---

## Security & Privacy

- **Offline-only**: Zero network requests. `offline_guard.py` monkeypatches `socket.connect` at startup to guarantee no connections can be made
- **Per-user isolation**: Each OS user has their own private database and directories
- **Local storage**: All files stored in local `./data/` directory
- **Restrictive permissions**: User folders set to owner-read/write only (chmod 700 on Unix; icacls guidance logged on Windows)
- **No telemetry**: No analytics, update checks, or external communication
- **Log sanitization**: In Redact mode, protocol-like tokens are replaced with `[REDACTED]` in all log output
- **Proxy detection**: Startup warns if `HTTP_PROXY`, `HTTPS_PROXY`, or similar env vars are set

---

## Output Files

### Updated CV
`{Original Name} (Updated YYYY-MM-DD).docx`

### Redacted CV
`{Original Name} (Redacted YYYY-MM-DD).docx`

### Logs
- JSON format: `{operation}_{timestamp}.json`
- CSV format: `{operation}_{timestamp}.csv`

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

The project includes a comprehensive test suite with 200 tests:

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

Build a single-file executable with PyInstaller:

```bash
pip install pyinstaller
pyinstaller cv_manager.spec
```

Output: `dist/CV_Manager.exe` (Windows) or `dist/CV_Manager` (macOS/Linux).

**Font note**: Calibri is bundled with Windows. On macOS/Linux, the app writes Calibri as the font name in .docx output — Word on the target machine handles font substitution if Calibri is unavailable.

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
5. Run `py src/main.py --mode validate-cv --cv "file.docx"` to check your CV
port

This is an offline, local application. For issues:
1. Check the logs in `./data/users/{username}/logs/`
2. Verify file formats match specifications above
3. Ensure write permissions to data directory
