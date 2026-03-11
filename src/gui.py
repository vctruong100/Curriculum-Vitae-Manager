"""
Offline GUI for the CV Research Experience Manager.
Built with tkinter - no network dependencies.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from pathlib import Path
from typing import Optional, List
import threading
from datetime import datetime

import re

from config import get_config, set_config, AppConfig, get_os_username, get_app_root, ALLOWED_FONTS
from models import Study, Site
from database import DatabaseManager
from processor import CVProcessor
from import_export import ImportExportManager
from normalizer import normalize_phase, validate_year
from progress_dialog import run_with_progress
from error_handler import FilePermissionError


class CVManagerApp:
    """Main application class for CV Research Experience Manager."""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("CV Research Experience Manager (Offline)")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        
        # Configuration
        self.config = get_config()
        self.config.ensure_user_directories()
        
        # Current user
        self.user_id = self.config.get_user_id()
        
        # Selected files/site
        self.cv_path: Optional[Path] = None
        self.master_path: Optional[Path] = None
        self.selected_site_id: Optional[int] = None
        
        # Setup UI
        self._setup_styles()
        self._create_menu()
        self._create_main_ui()
        
        # Status bar
        self._create_status_bar()
        
        # Refresh sites list
        self._refresh_sites()
    
    def _setup_styles(self):
        """Setup ttk styles."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure colors
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabel', background='#f0f0f0', font=('Segoe UI', 10))
        style.configure('TLabelframe', background='#f0f0f0')
        style.configure('TLabelframe.Label', background='#f0f0f0', font=('Segoe UI', 10))
        style.configure('TCheckbutton', background='#f0f0f0', font=('Segoe UI', 10))
        style.configure('TButton', font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'), background='#f0f0f0')
        style.configure('Title.TLabel', font=('Segoe UI', 14, 'bold'), background='#f0f0f0')
        style.configure('Status.TLabel', font=('Segoe UI', 9), background='#f0f0f0')
    
    def _create_menu(self):
        """Create application menu."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open CV (.docx)...", command=self._browse_cv)
        file_menu.add_command(label="Open Master (.xlsx)...", command=self._browse_master)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Database menu
        db_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Database", menu=db_menu)
        db_menu.add_command(label="Import .xlsx to Site...", command=self._import_xlsx)
        db_menu.add_command(label="Export Site to .xlsx...", command=self._export_site)
        db_menu.add_separator()
        db_menu.add_command(label="Open Data Folder", command=self._open_data_folder)
        
        # Configuration menu
        config_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Configuration", menu=config_menu)
        config_menu.add_command(label="Settings...", command=self._show_configuration)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self._show_about)
    
    def _create_main_ui(self):
        """Create main UI with notebook tabs."""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame,
            text="CV Research Experience Manager",
            style='Title.TLabel'
        )
        title_label.pack(pady=(0, 5))
        
        user_label = ttk.Label(
            main_frame,
            text=f"User: {self.user_id} | Mode: Offline Only",
            style='Status.TLabel'
        )
        user_label.pack(pady=(0, 10))
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Tab A: Update/Inject
        self.tab_update = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.tab_update, text="Mode A: Update/Inject")
        self._create_update_tab()
        
        # Tab B: Redact Protocols
        self.tab_redact = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.tab_redact, text="Mode B: Redact Protocols")
        self._create_redact_tab()
        
        # Tab C: Database Management
        self.tab_database = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.tab_database, text="Mode C: Database")
        self._create_database_tab()
    
    def _create_file_selection_frame(self, parent, include_site: bool = True) -> dict:
        """Create a reusable file selection frame."""
        frame = ttk.LabelFrame(parent, text="File Selection", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10))
        
        widgets = {}
        
        # CV File
        cv_frame = ttk.Frame(frame)
        cv_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(cv_frame, text="CV File (.docx):").pack(side=tk.LEFT)
        widgets['cv_entry'] = ttk.Entry(cv_frame, width=50)
        widgets['cv_entry'].pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(cv_frame, text="Browse...", command=lambda: self._browse_cv(widgets['cv_entry'])).pack(side=tk.LEFT)
        
        # Master source selection
        source_frame = ttk.LabelFrame(frame, text="Master Source", padding="5")
        source_frame.pack(fill=tk.X, pady=5)
        
        # Radio buttons for source type
        widgets['source_var'] = tk.StringVar(value="file")
        
        file_radio = ttk.Radiobutton(
            source_frame, text="Upload .xlsx file",
            variable=widgets['source_var'], value="file",
            command=lambda: self._toggle_source(widgets)
        )
        file_radio.pack(anchor=tk.W)
        
        # File entry
        file_frame = ttk.Frame(source_frame)
        file_frame.pack(fill=tk.X, pady=2)
        widgets['master_entry'] = ttk.Entry(file_frame, width=50)
        widgets['master_entry'].pack(side=tk.LEFT, padx=(20, 5), fill=tk.X, expand=True)
        widgets['master_browse'] = ttk.Button(file_frame, text="Browse...", command=lambda: self._browse_master(widgets['master_entry']))
        widgets['master_browse'].pack(side=tk.LEFT)
        
        if include_site:
            site_radio = ttk.Radiobutton(
                source_frame, text="Use saved Site database",
                variable=widgets['source_var'], value="site",
                command=lambda: self._toggle_source(widgets)
            )
            site_radio.pack(anchor=tk.W, pady=(5, 0))
            
            # Site dropdown
            site_frame = ttk.Frame(source_frame)
            site_frame.pack(fill=tk.X, pady=2)
            widgets['site_combo'] = ttk.Combobox(site_frame, width=47, state='disabled')
            widgets['site_combo'].pack(side=tk.LEFT, padx=(20, 5))
            widgets['site_refresh'] = ttk.Button(site_frame, text="Refresh", command=lambda: self._refresh_site_combo(widgets['site_combo']), state='disabled')
            widgets['site_refresh'].pack(side=tk.LEFT)
        
        return widgets
    
    def _toggle_source(self, widgets: dict):
        """Toggle between file and site source."""
        if widgets['source_var'].get() == "file":
            widgets['master_entry'].config(state='normal')
            widgets['master_browse'].config(state='normal')
            if 'site_combo' in widgets:
                widgets['site_combo'].config(state='disabled')
                widgets['site_refresh'].config(state='disabled')
        else:
            widgets['master_entry'].config(state='disabled')
            widgets['master_browse'].config(state='disabled')
            if 'site_combo' in widgets:
                widgets['site_combo'].config(state='readonly')
                widgets['site_refresh'].config(state='normal')
                self._refresh_site_combo(widgets['site_combo'])
    
    def _create_update_tab(self):
        """Create Mode A: Update/Inject tab."""
        # Description
        desc = ttk.Label(
            self.tab_update,
            text="Inject new studies from master list into CV. Studies above the benchmark year will be added.\nProtocols are shown in bold red.",
            wraplength=900
        )
        desc.pack(anchor=tk.W, pady=(0, 10))
        
        # File selection
        self.update_widgets = self._create_file_selection_frame(self.tab_update)
        
        # Options
        options_frame = ttk.LabelFrame(self.tab_update, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.update_preview_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Preview changes before applying",
            variable=self.update_preview_var
        ).pack(anchor=tk.W)
        
        # Benchmark Configuration
        benchmark_outer = ttk.LabelFrame(options_frame, text="Year Benchmark", padding="8")
        benchmark_outer.pack(fill=tk.X, pady=(8, 0))
        
        benchmark_frame = ttk.Frame(benchmark_outer)
        benchmark_frame.pack(fill=tk.X)
        
        self.auto_benchmark_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            benchmark_frame,
            text="Auto-detect from CV (use highest year found)",
            variable=self.auto_benchmark_var,
            command=self._toggle_benchmark_input
        ).pack(side=tk.LEFT)
        
        manual_label = ttk.Label(benchmark_frame, text="   or Manual year:")
        manual_label.pack(side=tk.LEFT, padx=(15, 0))
        
        self.benchmark_year_var = tk.StringVar()
        self.benchmark_year_entry = ttk.Entry(
            benchmark_frame, 
            width=10, 
            textvariable=self.benchmark_year_var, 
            state='disabled',
            font=('Segoe UI', 10)
        )
        self.benchmark_year_entry.pack(side=tk.LEFT, padx=(5, 8))
        
        help_label = ttk.Label(
            benchmark_frame, 
            text="(inject from this year onward)",
            foreground='#666666',
            font=('Segoe UI', 9, 'italic')
        )
        help_label.pack(side=tk.LEFT)
        
        # Action buttons
        btn_frame = ttk.Frame(self.tab_update)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(
            btn_frame,
            text="Preview Changes",
            command=self._preview_update
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="Update CV",
            command=self._run_update
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="Open Results Folder",
            command=self._open_results_folder
        ).pack(side=tk.LEFT, padx=5)
        
        # Results area
        results_frame = ttk.LabelFrame(self.tab_update, text="Results", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        self.update_results = tk.Text(results_frame, height=20, wrap=tk.WORD)
        self.update_results.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.update_results.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.update_results.config(yscrollcommand=scrollbar.set)
    
    def _create_redact_tab(self):
        """Create Mode B: Redact Protocols tab."""
        # Description
        desc = ttk.Label(
            self.tab_redact,
            text="Remove protocols and mask treatment names from CV studies.\nMatches CV entries against master list and replaces with masked versions.",
            wraplength=900
        )
        desc.pack(anchor=tk.W, pady=(0, 10))
        
        # File selection
        self.redact_widgets = self._create_file_selection_frame(self.tab_redact)
        
        # Action buttons
        btn_frame = ttk.Frame(self.tab_redact)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(
            btn_frame,
            text="Preview Redactions",
            command=self._preview_redact
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="Redact CV",
            command=self._run_redact
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="Open Results Folder",
            command=self._open_results_folder
        ).pack(side=tk.LEFT, padx=5)
        
        # Results area
        results_frame = ttk.LabelFrame(self.tab_redact, text="Results", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        self.redact_results = tk.Text(results_frame, height=20, wrap=tk.WORD)
        self.redact_results.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.redact_results.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.redact_results.config(yscrollcommand=scrollbar.set)
    
    def _create_database_tab(self):
        """Create Mode C: Database Management tab."""
        # Main container with paned window
        paned = ttk.PanedWindow(self.tab_database, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)
        
        # Left panel: Sites list + Categories/Subcategories ordering
        left_frame = ttk.Frame(paned, padding="5")
        paned.add(left_frame, weight=1)
        
        # Split left panel with vertical paned window
        left_paned = ttk.PanedWindow(left_frame, orient=tk.VERTICAL)
        left_paned.pack(fill=tk.BOTH, expand=True)
        
        # Upper half: My Sites
        sites_frame = ttk.Frame(left_paned)
        left_paned.add(sites_frame, weight=1)
        
        ttk.Label(sites_frame, text="My Sites", style='Header.TLabel').pack(anchor=tk.W)
        
        # Sites listbox
        list_frame = ttk.Frame(sites_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.sites_listbox = tk.Listbox(list_frame, width=30, height=8)
        self.sites_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sites_listbox.bind('<<ListboxSelect>>', self._on_site_select)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.sites_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sites_listbox.config(yscrollcommand=scrollbar.set)
        
        # Site buttons
        site_btn_frame = ttk.Frame(sites_frame)
        site_btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(site_btn_frame, text="New Site", command=self._create_site).pack(side=tk.LEFT, padx=2)
        ttk.Button(site_btn_frame, text="Rename", command=self._rename_site).pack(side=tk.LEFT, padx=2)
        ttk.Button(site_btn_frame, text="Delete", command=self._delete_site).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(site_btn_frame, text="Import", command=self._import_xlsx).pack(side=tk.RIGHT, padx=2)
        ttk.Button(site_btn_frame, text="Export", command=self._export_site).pack(side=tk.RIGHT, padx=2)
        
        # Lower half: Categories/Subcategories ordering
        ordering_frame = ttk.Frame(left_paned)
        left_paned.add(ordering_frame, weight=1)
        
        ttk.Label(ordering_frame, text="Category Order", style='Header.TLabel').pack(anchor=tk.W)
        
        # Categories listbox with drag support
        cat_list_frame = ttk.Frame(ordering_frame)
        cat_list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.categories_listbox = tk.Listbox(cat_list_frame, width=30, height=8, selectmode=tk.SINGLE)
        self.categories_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        cat_scrollbar = ttk.Scrollbar(cat_list_frame, orient=tk.VERTICAL, command=self.categories_listbox.yview)
        cat_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.categories_listbox.config(yscrollcommand=cat_scrollbar.set)
        
        # Bind drag events for reordering
        self.categories_listbox.bind('<Button-1>', self._on_cat_click)
        self.categories_listbox.bind('<B1-Motion>', self._on_cat_drag)
        self.categories_listbox.bind('<ButtonRelease-1>', self._on_cat_drop)
        self._cat_drag_data = {'index': None}
        
        # Category order buttons
        cat_btn_frame = ttk.Frame(ordering_frame)
        cat_btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(cat_btn_frame, text="Move Up", command=self._move_category_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(cat_btn_frame, text="Move Down", command=self._move_category_down).pack(side=tk.LEFT, padx=2)
        ttk.Button(cat_btn_frame, text="Save Order", command=self._save_category_order).pack(side=tk.RIGHT, padx=2)
        
        # Right panel: Studies
        right_frame = ttk.Frame(paned, padding="5")
        paned.add(right_frame, weight=3)
        
        self.studies_header = ttk.Label(right_frame, text="Studies", style='Header.TLabel')
        self.studies_header.pack(anchor=tk.W)
        
        # Search bar
        search_frame = ttk.Frame(right_frame)
        search_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT)
        self.study_search_var = tk.StringVar()
        self.study_search_var.trace('w', self._on_study_search)
        self.study_search_entry = ttk.Entry(search_frame, textvariable=self.study_search_var, width=40)
        self.study_search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(search_frame, text="Clear", command=lambda: self.study_search_var.set("")).pack(side=tk.LEFT)
        
        # Studies treeview with both scrollbars
        tree_frame = ttk.Frame(right_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Configure grid for proper scrollbar placement
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        columns = ('phase', 'subcategory', 'year', 'sponsor', 'protocol', 'desc_masked', 'desc_full')
        self.studies_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Track sort state
        self._study_sort_column = None
        self._study_sort_reverse = False
        
        # Set up column headings with sort functionality
        self.studies_tree.heading('phase', text='Phase', command=lambda: self._sort_studies_column('phase'))
        self.studies_tree.heading('subcategory', text='Subcategory', command=lambda: self._sort_studies_column('subcategory'))
        self.studies_tree.heading('year', text='Year', command=lambda: self._sort_studies_column('year'))
        self.studies_tree.heading('sponsor', text='Sponsor', command=lambda: self._sort_studies_column('sponsor'))
        self.studies_tree.heading('protocol', text='Protocol', command=lambda: self._sort_studies_column('protocol'))
        self.studies_tree.heading('desc_masked', text='Masked Description', command=lambda: self._sort_studies_column('desc_masked'))
        self.studies_tree.heading('desc_full', text='Full Description', command=lambda: self._sort_studies_column('desc_full'))
        
        self.studies_tree.column('phase', width=80, minwidth=60)
        self.studies_tree.column('subcategory', width=100, minwidth=80)
        self.studies_tree.column('year', width=50, minwidth=40)
        self.studies_tree.column('sponsor', width=120, minwidth=80)
        self.studies_tree.column('protocol', width=100, minwidth=60)
        self.studies_tree.column('desc_masked', width=250, minwidth=150)
        self.studies_tree.column('desc_full', width=250, minwidth=150)
        
        # Place treeview and scrollbars using grid
        self.studies_tree.grid(row=0, column=0, sticky='nsew')
        
        # Vertical scrollbar
        tree_scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.studies_tree.yview)
        tree_scroll_y.grid(row=0, column=1, sticky='ns')
        
        # Horizontal scrollbar
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.studies_tree.xview)
        tree_scroll_x.grid(row=1, column=0, sticky='ew')
        
        self.studies_tree.config(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        # Store all studies for filtering
        self._all_studies_data = []
        
        # Study buttons
        study_btn_frame = ttk.Frame(right_frame)
        study_btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(study_btn_frame, text="Add Study", command=self._add_study).pack(side=tk.LEFT, padx=2)
        ttk.Button(study_btn_frame, text="Edit Study", command=self._edit_study).pack(side=tk.LEFT, padx=2)
        ttk.Button(study_btn_frame, text="Delete Study", command=self._delete_study).pack(side=tk.LEFT, padx=2)
        ttk.Button(study_btn_frame, text="Export to .xlsx", command=self._export_site).pack(side=tk.RIGHT, padx=2)
        ttk.Button(study_btn_frame, text="Refresh", command=self._refresh_studies).pack(side=tk.RIGHT, padx=2)
    
    def _create_status_bar(self):
        """Create status bar at bottom of window."""
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            style='Status.TLabel',
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding=(5, 2)
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def _set_status(self, message: str):
        """Update status bar message."""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    # ==================== File Operations ====================
    
    def _browse_cv(self, entry: Optional[ttk.Entry] = None):
        """Browse for CV .docx file."""
        path = filedialog.askopenfilename(
            title="Select CV Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if path:
            self.cv_path = Path(path)
            if entry:
                entry.delete(0, tk.END)
                entry.insert(0, str(self.cv_path))
    
    def _browse_master(self, entry: Optional[ttk.Entry] = None):
        """Browse for master .xlsx file."""
        path = filedialog.askopenfilename(
            title="Select Master Study List",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if path:
            self.master_path = Path(path)
            if entry:
                entry.delete(0, tk.END)
                entry.insert(0, str(self.master_path))
    
    def _open_data_folder(self):
        """Open the user's data folder."""
        import os
        data_path = self.config.get_user_data_path()
        data_path.mkdir(parents=True, exist_ok=True)
        os.startfile(str(data_path))
    
    def _open_results_folder(self):
        """Open the results folder in the project root."""
        import os
        from config import get_app_root
        results_path = get_app_root() / "result"
        results_path.mkdir(exist_ok=True)
        os.startfile(str(results_path))
    
    def _toggle_benchmark_input(self):
        """Toggle the benchmark year input based on auto-find checkbox."""
        if self.auto_benchmark_var.get():
            self.benchmark_year_entry.config(state='disabled')
            self.benchmark_year_var.set('')
        else:
            self.benchmark_year_entry.config(state='normal')
    
    # ==================== Mode A: Update/Inject ====================
    
    def _get_source_from_widgets(self, widgets: dict) -> tuple:
        """Get master_path and site_id from widget state."""
        master_path = None
        site_id = None
        
        if widgets['source_var'].get() == "file":
            path_str = widgets['master_entry'].get().strip()
            if path_str:
                master_path = Path(path_str)
        else:
            if 'site_combo' in widgets:
                selection = widgets['site_combo'].get()
                if selection:
                    # Extract site ID from selection (format: "Name (ID: X)")
                    try:
                        site_id = int(selection.split("(ID: ")[-1].rstrip(")"))
                    except (ValueError, IndexError):
                        pass
        
        return master_path, site_id
    
    def _preview_update(self):
        """Preview Mode A changes."""
        cv_path_str = self.update_widgets['cv_entry'].get().strip()
        if not cv_path_str:
            messagebox.showerror("Error", "Please select a CV file")
            return
        
        master_path, site_id = self._get_source_from_widgets(self.update_widgets)
        if not master_path and site_id is None:
            messagebox.showerror("Error", "Please select a master source (file or site)")
            return
        
        self._set_status("Previewing changes...")
        self.update_results.delete(1.0, tk.END)
        
        try:
            processor = CVProcessor(self.config)
            changes, error = processor.preview_changes(
                Path(cv_path_str),
                master_path,
                site_id,
                mode="update_inject"
            )
            
            if error:
                self.update_results.insert(tk.END, f"Error: {error}\n")
            elif not changes:
                self.update_results.insert(tk.END, "No changes to make. CV is up to date.\n")
            else:
                self.update_results.insert(tk.END, f"Found {len(changes)} studies to inject:\n\n")
                for change in changes:
                    self.update_results.insert(
                        tk.END,
                        f"• {change['year']} | {change['phase']} > {change['subcategory']}\n"
                        f"  {change['sponsor']} {change['protocol']}: {change['description']}\n\n"
                    )
            
            self._set_status("Preview complete")
            
        except Exception as e:
            self.update_results.insert(tk.END, f"Error: {str(e)}\n")
            self._set_status("Preview failed")
    
    def _run_update(self):
        """Run Mode A: Update/Inject."""
        cv_path_str = self.update_widgets['cv_entry'].get().strip()
        if not cv_path_str:
            messagebox.showerror("Error", "Please select a CV file")
            return
        
        master_path, site_id = self._get_source_from_widgets(self.update_widgets)
        if not master_path and site_id is None:
            messagebox.showerror("Error", "Please select a master source (file or site)")
            return
        
        # Get manual benchmark year if specified
        manual_benchmark = None
        if not self.auto_benchmark_var.get():
            year_str = self.benchmark_year_var.get().strip()
            if year_str:
                try:
                    manual_benchmark = int(year_str)
                    if manual_benchmark < 1900 or manual_benchmark > 2100:
                        messagebox.showerror("Error", "Please enter a valid year (1900-2100)")
                        return
                except ValueError:
                    messagebox.showerror("Error", "Please enter a valid year number")
                    return
            else:
                messagebox.showerror("Error", "Please enter a benchmark year or enable Auto-find Benchmark")
                return
        
        self.update_results.delete(1.0, tk.END)
        
        try:
            # Run with progress dialog
            processor = CVProcessor(self.config)
            result = run_with_progress(
                self.root,
                "Updating CV",
                "Processing CV update, please wait...",
                processor.mode_a_update_inject,
                Path(cv_path_str),
                master_path,
                site_id,
                manual_benchmark
            )
            
            if result.success:
                self.update_results.insert(tk.END, "Update completed successfully!\n\n")
                if result.output_path:
                    self.update_results.insert(tk.END, f"Output file: {result.output_path}\n\n")
                
                counts = result.get_counts()
                self.update_results.insert(tk.END, "Summary:\n")
                for op, count in counts.items():
                    self.update_results.insert(tk.END, f"  • {op}: {count}\n")
                
                self._set_status("Update complete")
                messagebox.showinfo("Success", f"CV updated successfully!\n\nSaved to: {result.output_path}")
            else:
                self.update_results.insert(tk.END, f"Update failed: {result.error_message}\n")
                self._set_status("Update failed")
                
        except FilePermissionError as e:
            messagebox.showerror("File Access Error", str(e))
            self.update_results.insert(tk.END, f"Error: {str(e)}\n")
            self._set_status("Update failed")
        except Exception as e:
            self.update_results.insert(tk.END, f"Error: {str(e)}\n")
            self._set_status("Update failed")
    
    # ==================== Mode B: Redact ====================
    
    def _preview_redact(self):
        """Preview Mode B changes."""
        cv_path_str = self.redact_widgets['cv_entry'].get().strip()
        if not cv_path_str:
            messagebox.showerror("Error", "Please select a CV file")
            return
        
        master_path, site_id = self._get_source_from_widgets(self.redact_widgets)
        if not master_path and site_id is None:
            messagebox.showerror("Error", "Please select a master source (file or site)")
            return
        
        self._set_status("Previewing redactions...")
        self.redact_results.delete(1.0, tk.END)
        
        try:
            processor = CVProcessor(self.config)
            changes, error = processor.preview_changes(
                Path(cv_path_str),
                master_path,
                site_id,
                mode="redact_protocols"
            )
            
            if error:
                self.redact_results.insert(tk.END, f"Error: {error}\n")
            elif not changes:
                self.redact_results.insert(tk.END, "No studies matched for redaction.\n")
            else:
                self.redact_results.insert(tk.END, f"Found {len(changes)} studies to redact:\n\n")
                for change in changes:
                    self.redact_results.insert(
                        tk.END,
                        f"• {change['year']} | {change['sponsor']} {change['protocol']}\n"
                        f"  Match score: {change['match_score']}%\n"
                        f"  New text: {change['new_description']}\n\n"
                    )
            
            self._set_status("Preview complete")
            
        except Exception as e:
            self.redact_results.insert(tk.END, f"Error: {str(e)}\n")
            self._set_status("Preview failed")
    
    def _run_redact(self):
        """Run Mode B: Redact Protocols."""
        cv_path_str = self.redact_widgets['cv_entry'].get().strip()
        if not cv_path_str:
            messagebox.showerror("Error", "Please select a CV file")
            return
        
        master_path, site_id = self._get_source_from_widgets(self.redact_widgets)
        if not master_path and site_id is None:
            messagebox.showerror("Error", "Please select a master source (file or site)")
            return
        
        self.redact_results.delete(1.0, tk.END)
        
        try:
            # Run with progress dialog
            processor = CVProcessor(self.config)
            result = run_with_progress(
                self.root,
                "Redacting CV",
                "Processing CV redaction, please wait...",
                processor.mode_b_redact_protocols,
                Path(cv_path_str),
                master_path,
                site_id
            )
            
            if result.success:
                self.redact_results.insert(tk.END, "Redaction completed successfully!\n\n")
                if result.output_path:
                    self.redact_results.insert(tk.END, f"Output file: {result.output_path}\n\n")
                
                counts = result.get_counts()
                self.redact_results.insert(tk.END, "Summary:\n")
                for op, count in counts.items():
                    self.redact_results.insert(tk.END, f"  • {op}: {count}\n")
                
                self._set_status("Redaction complete")
                messagebox.showinfo("Success", f"CV redacted successfully!\n\nSaved to: {result.output_path}")
            else:
                self.redact_results.insert(tk.END, f"Redaction failed: {result.error_message}\n")
                self._set_status("Redaction failed")
                
        except FilePermissionError as e:
            messagebox.showerror("File Access Error", str(e))
            self.redact_results.insert(tk.END, f"Error: {str(e)}\n")
            self._set_status("Redaction failed")
        except Exception as e:
            self.redact_results.insert(tk.END, f"Error: {str(e)}\n")
            self._set_status("Redaction failed")
    
    # ==================== Mode C: Database ====================
    
    def _refresh_sites(self):
        """Refresh the sites listbox."""
        self.sites_listbox.delete(0, tk.END)
        
        try:
            with DatabaseManager(config=self.config) as db:
                sites = db.get_sites()
                for site in sites:
                    count = db.get_study_count(site.id)
                    self.sites_listbox.insert(tk.END, f"{site.name} ({count} studies)")
                    # Store site ID as data
                    self.sites_listbox.itemconfig(tk.END, {'foreground': 'black'})
            
            # Store sites for reference
            self._sites = sites
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sites: {str(e)}")
    
    def _refresh_site_combo(self, combo: ttk.Combobox):
        """Refresh a site combobox."""
        try:
            with DatabaseManager(config=self.config) as db:
                sites = db.get_sites()
                values = [f"{s.name} (ID: {s.id})" for s in sites]
                combo['values'] = values
                if values:
                    combo.current(0)
        except Exception as e:
            pass
    
    def _on_site_select(self, event):
        """Handle site selection."""
        selection = self.sites_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        if hasattr(self, '_sites') and idx < len(self._sites):
            site = self._sites[idx]
            self.selected_site_id = site.id
            self.studies_header.config(text=f"Studies - {site.name}")
            self._refresh_studies()
            self._refresh_categories()
    
    def _refresh_studies(self):
        """Refresh the studies treeview."""
        # Clear existing
        for item in self.studies_tree.get_children():
            self.studies_tree.delete(item)
        
        if not self.selected_site_id:
            self._all_studies_data = []
            return
        
        try:
            with DatabaseManager(config=self.config) as db:
                studies = db.get_studies(self.selected_site_id)
                self._studies = studies
                
                # Store all studies data for search filtering
                self._all_studies_data = []
                for study in studies:
                    self._all_studies_data.append({
                        'id': study.id,
                        'phase': study.phase,
                        'subcategory': study.subcategory,
                        'year': study.year,
                        'sponsor': study.sponsor,
                        'protocol': study.protocol,
                        'desc_masked': study.description_masked,
                        'desc_full': study.description_full
                    })
                
                # Apply search filter if active
                self._apply_study_filter()
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load studies: {str(e)}")
    
    def _apply_study_filter(self):
        """Apply search filter to studies treeview."""
        # Clear existing
        for item in self.studies_tree.get_children():
            self.studies_tree.delete(item)
        
        search_term = self.study_search_var.get().lower().strip() if hasattr(self, 'study_search_var') else ""
        
        for study in self._all_studies_data:
            # Check if study matches search term (search in all fields including descriptions)
            if search_term:
                searchable = f"{study['phase']} {study['subcategory']} {study['year']} {study['sponsor']} {study['protocol']} {study.get('desc_masked', '')} {study.get('desc_full', '')}".lower()
                if search_term not in searchable:
                    continue
            
            self.studies_tree.insert(
                '',
                tk.END,
                values=(
                    study['phase'],
                    study['subcategory'],
                    study['year'],
                    study['sponsor'],
                    study['protocol'],
                    study.get('desc_masked', ''),
                    study.get('desc_full', '')
                ),
                iid=str(study['id'])
            )
    
    def _on_study_search(self, *args):
        """Handle search input changes."""
        self._apply_study_filter()
    
    def _sort_studies_column(self, column):
        """Sort studies treeview by the specified column."""
        # Toggle sort direction if clicking same column
        if self._study_sort_column == column:
            self._study_sort_reverse = not self._study_sort_reverse
        else:
            self._study_sort_column = column
            self._study_sort_reverse = False
        
        # Get column index
        columns = ('phase', 'subcategory', 'year', 'sponsor', 'protocol', 'desc_masked', 'desc_full')
        col_idx = columns.index(column)
        
        # Sort the underlying data
        def sort_key(study):
            value = None
            if column == 'phase':
                value = study['phase']
            elif column == 'subcategory':
                value = study['subcategory']
            elif column == 'year':
                value = study['year']
            elif column == 'sponsor':
                value = study['sponsor']
            elif column == 'protocol':
                value = study['protocol']
            elif column == 'desc_masked':
                value = study.get('desc_masked', '')
            elif column == 'desc_full':
                value = study.get('desc_full', '')
            
            # Handle numeric sorting for year
            if column == 'year':
                return (value if value else 0)
            # String sorting for others
            return str(value).lower() if value else ''
        
        self._all_studies_data.sort(key=sort_key, reverse=self._study_sort_reverse)
        
        # Reapply filter to refresh display with new sort order
        self._apply_study_filter()
    
    def _refresh_categories(self):
        """Refresh the categories listbox for the selected site."""
        self.categories_listbox.delete(0, tk.END)
        
        if not self.selected_site_id:
            return
        
        try:
            with DatabaseManager(config=self.config) as db:
                studies = db.get_studies(self.selected_site_id)
                
                # Extract unique phase/subcategory combinations
                categories = {}
                for study in studies:
                    key = f"{study.phase} > {study.subcategory}"
                    if key not in categories:
                        categories[key] = {'phase': study.phase, 'subcategory': study.subcategory}
                
                # Try to load saved order from database
                saved_order = db.get_category_order(self.selected_site_id)
                
                if saved_order:
                    # Use saved order, but add any new categories at the end
                    final_order = []
                    for key in saved_order:
                        if key in categories:
                            final_order.append(key)
                    # Add any categories not in saved order
                    for key in sorted(categories.keys()):
                        if key not in final_order:
                            final_order.append(key)
                else:
                    # Sort by phase then subcategory (default)
                    final_order = sorted(categories.keys())
                
                self._category_order = final_order
                for key in final_order:
                    self.categories_listbox.insert(tk.END, key)
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load categories: {str(e)}")
    
    def _on_cat_click(self, event):
        """Handle mouse click on categories listbox."""
        self._cat_drag_data['index'] = self.categories_listbox.nearest(event.y)
    
    def _on_cat_drag(self, event):
        """Handle drag motion on categories listbox."""
        if self._cat_drag_data['index'] is None:
            return
        
        # Get the index under the current mouse position
        new_index = self.categories_listbox.nearest(event.y)
        
        if new_index != self._cat_drag_data['index']:
            # Swap items
            item = self.categories_listbox.get(self._cat_drag_data['index'])
            self.categories_listbox.delete(self._cat_drag_data['index'])
            self.categories_listbox.insert(new_index, item)
            self._cat_drag_data['index'] = new_index
            self.categories_listbox.selection_clear(0, tk.END)
            self.categories_listbox.selection_set(new_index)
    
    def _on_cat_drop(self, event):
        """Handle mouse release on categories listbox."""
        self._cat_drag_data['index'] = None
    
    def _move_category_up(self):
        """Move selected category up in the list."""
        selection = self.categories_listbox.curselection()
        if not selection or selection[0] == 0:
            return
        
        idx = selection[0]
        item = self.categories_listbox.get(idx)
        self.categories_listbox.delete(idx)
        self.categories_listbox.insert(idx - 1, item)
        self.categories_listbox.selection_set(idx - 1)
    
    def _move_category_down(self):
        """Move selected category down in the list."""
        selection = self.categories_listbox.curselection()
        if not selection or selection[0] >= self.categories_listbox.size() - 1:
            return
        
        idx = selection[0]
        item = self.categories_listbox.get(idx)
        self.categories_listbox.delete(idx)
        self.categories_listbox.insert(idx + 1, item)
        self.categories_listbox.selection_set(idx + 1)
    
    def _save_category_order(self):
        """Save the current category order to the database."""
        if not self.selected_site_id:
            messagebox.showwarning("Warning", "Please select a site first")
            return
        
        # Get current order from listbox
        order = []
        for i in range(self.categories_listbox.size()):
            order.append(self.categories_listbox.get(i))
        
        try:
            with DatabaseManager(config=self.config) as db:
                db.save_category_order(self.selected_site_id, order)
            messagebox.showinfo("Success", "Category order saved!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save category order: {str(e)}")
    
    def _create_site(self):
        """Create a new site."""
        name = simpledialog.askstring("New Site", "Enter site name:")
        if not name:
            return
        
        try:
            with DatabaseManager(config=self.config) as db:
                site = db.create_site(name.strip())
                messagebox.showinfo("Success", f"Site '{site.name}' created successfully!")
                self._refresh_sites()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create site: {str(e)}")
    
    def _rename_site(self):
        """Rename selected site."""
        if not self.selected_site_id:
            messagebox.showwarning("Warning", "Please select a site first")
            return
        
        # Get current name
        current_name = ""
        if hasattr(self, '_sites'):
            for site in self._sites:
                if site.id == self.selected_site_id:
                    current_name = site.name
                    break
        
        new_name = simpledialog.askstring("Rename Site", "Enter new name:", initialvalue=current_name)
        if not new_name or new_name == current_name:
            return
        
        try:
            with DatabaseManager(config=self.config) as db:
                if db.rename_site(self.selected_site_id, new_name.strip()):
                    messagebox.showinfo("Success", "Site renamed successfully!")
                    self._refresh_sites()
                else:
                    messagebox.showerror("Error", "Failed to rename site")
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to rename site: {str(e)}")
    
    def _delete_site(self):
        """Delete selected site."""
        if not self.selected_site_id:
            messagebox.showwarning("Warning", "Please select a site first")
            return
        
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this site?\nThis action cannot be undone."):
            return
        
        try:
            with DatabaseManager(config=self.config) as db:
                if db.delete_site(self.selected_site_id):
                    messagebox.showinfo("Success", "Site deleted successfully!")
                    self.selected_site_id = None
                    self._refresh_sites()
                    self._refresh_studies()
                else:
                    messagebox.showerror("Error", "Failed to delete site")
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete site: {str(e)}")
    
    def _import_xlsx(self):
        """Import .xlsx file to a site."""
        path = filedialog.askopenfilename(
            title="Select Master Study List to Import",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if not path:
            return
        
        name = simpledialog.askstring("Import", "Enter name for this site:")
        if not name:
            return
        
        self._set_status("Importing...")
        
        try:
            manager = ImportExportManager(self.config)
            success, message, site_id = manager.import_xlsx_to_site(
                Path(path),
                name.strip(),
                replace_existing=True
            )
            
            if success:
                messagebox.showinfo("Success", message)
                self._refresh_sites()
            else:
                messagebox.showerror("Error", message)
            
            self._set_status("Ready")
            
        except Exception as e:
            messagebox.showerror("Error", f"Import failed: {str(e)}")
            self._set_status("Import failed")
    
    def _export_site(self):
        """Export selected site to .xlsx."""
        if not self.selected_site_id:
            messagebox.showwarning("Warning", "Please select a site first")
            return
        
        self._set_status("Exporting...")
        
        try:
            manager = ImportExportManager(self.config)
            success, message, output_path = manager.export_site_to_xlsx(self.selected_site_id)
            
            if success:
                messagebox.showinfo("Success", f"{message}\n\nFile saved to:\n{output_path}")
            else:
                messagebox.showerror("Error", message)
            
            self._set_status("Ready")
            
        except FilePermissionError as e:
            messagebox.showerror("File Access Error", str(e))
            self._set_status("Export failed")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")
            self._set_status("Export failed")
    
    def _add_study(self):
        """Add a new study to selected site."""
        if not self.selected_site_id:
            messagebox.showwarning("Warning", "Please select a site first")
            return
        
        dialog = StudyDialog(self.root, "Add Study")
        if dialog.result:
            try:
                with DatabaseManager(config=self.config) as db:
                    study = Study(**dialog.result)
                    db.add_study(self.selected_site_id, study)
                    self._refresh_studies()
                    messagebox.showinfo("Success", "Study added successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add study: {str(e)}")
    
    def _edit_study(self):
        """Edit selected study."""
        selection = self.studies_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a study first")
            return
        
        study_id = int(selection[0])
        
        # Find the study
        study = None
        if hasattr(self, '_studies'):
            for s in self._studies:
                if s.id == study_id:
                    study = s
                    break
        
        if not study:
            return
        
        dialog = StudyDialog(self.root, "Edit Study", study)
        if dialog.result:
            try:
                with DatabaseManager(config=self.config) as db:
                    study.phase = dialog.result['phase']
                    study.subcategory = dialog.result['subcategory']
                    study.year = dialog.result['year']
                    study.sponsor = dialog.result['sponsor']
                    study.protocol = dialog.result['protocol']
                    study.description_full = dialog.result['description_full']
                    study.description_masked = dialog.result['description_masked']
                    
                    db.update_study(study)
                    self._refresh_studies()
                    messagebox.showinfo("Success", "Study updated successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update study: {str(e)}")
    
    def _delete_study(self):
        """Delete selected study."""
        selection = self.studies_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a study first")
            return
        
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this study?"):
            return
        
        study_id = int(selection[0])
        
        try:
            with DatabaseManager(config=self.config) as db:
                if db.delete_study(study_id, self.selected_site_id):
                    self._refresh_studies()
                    messagebox.showinfo("Success", "Study deleted successfully!")
                else:
                    messagebox.showerror("Error", "Failed to delete study")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete study: {str(e)}")
    
    def _show_configuration(self):
        """Show configuration dialog."""
        dialog = ConfigurationDialog(self.root, self.config)
        if dialog.result is not None:
            self.config = dialog.result
            set_config(self.config)
            self._set_status("Configuration saved")

    def _show_about(self):
        """Show about dialog with rendered README.md content."""
        ReadmeViewer(self.root)


class StudyDialog(simpledialog.Dialog):
    """Dialog for adding/editing a study."""
    
    def __init__(self, parent, title: str, study: Optional[Study] = None):
        self.study = study
        self.result = None
        super().__init__(parent, title)
    
    def body(self, master):
        """Create dialog body."""
        # Phase
        ttk.Label(master, text="Phase:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.phase_var = tk.StringVar(value=self.study.phase if self.study else "Phase I")
        self.phase_combo = ttk.Combobox(master, textvariable=self.phase_var, values=["Phase I", "Phase II–IV"])
        self.phase_combo.grid(row=0, column=1, sticky=tk.EW, pady=2)
        
        # Subcategory
        ttk.Label(master, text="Subcategory:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.subcat_entry = ttk.Entry(master, width=40)
        self.subcat_entry.grid(row=1, column=1, sticky=tk.EW, pady=2)
        if self.study:
            self.subcat_entry.insert(0, self.study.subcategory)
        
        # Year
        ttk.Label(master, text="Year:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.year_entry = ttk.Entry(master, width=10)
        self.year_entry.grid(row=2, column=1, sticky=tk.W, pady=2)
        if self.study:
            self.year_entry.insert(0, str(self.study.year))
        
        # Sponsor
        ttk.Label(master, text="Sponsor:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.sponsor_entry = ttk.Entry(master, width=40)
        self.sponsor_entry.grid(row=3, column=1, sticky=tk.EW, pady=2)
        if self.study:
            self.sponsor_entry.insert(0, self.study.sponsor)
        
        # Protocol
        ttk.Label(master, text="Protocol:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.protocol_entry = ttk.Entry(master, width=40)
        self.protocol_entry.grid(row=4, column=1, sticky=tk.EW, pady=2)
        if self.study and self.study.protocol:
            self.protocol_entry.insert(0, self.study.protocol)
        
        # Description Full
        ttk.Label(master, text="Description (Full):").grid(row=5, column=0, sticky=tk.NW, pady=2)
        self.desc_full_text = tk.Text(master, width=40, height=3)
        self.desc_full_text.grid(row=5, column=1, sticky=tk.EW, pady=2)
        if self.study:
            self.desc_full_text.insert(tk.END, self.study.description_full)
        
        # Description Masked
        ttk.Label(master, text="Description (Masked):").grid(row=6, column=0, sticky=tk.NW, pady=2)
        self.desc_masked_text = tk.Text(master, width=40, height=3)
        self.desc_masked_text.grid(row=6, column=1, sticky=tk.EW, pady=2)
        if self.study:
            self.desc_masked_text.insert(tk.END, self.study.description_masked)
        
        master.columnconfigure(1, weight=1)
        
        return self.phase_combo
    
    def validate(self):
        """Validate input."""
        year_str = self.year_entry.get().strip()
        year = validate_year(year_str)
        if year is None:
            messagebox.showerror("Error", "Please enter a valid 4-digit year")
            return False
        
        if not self.subcat_entry.get().strip():
            messagebox.showerror("Error", "Please enter a subcategory")
            return False
        
        if not self.sponsor_entry.get().strip():
            messagebox.showerror("Error", "Please enter a sponsor")
            return False
        
        return True
    
    def apply(self):
        """Apply changes."""
        self.result = {
            'phase': normalize_phase(self.phase_var.get()),
            'subcategory': self.subcat_entry.get().strip(),
            'year': int(self.year_entry.get().strip()),
            'sponsor': self.sponsor_entry.get().strip(),
            'protocol': self.protocol_entry.get().strip(),
            'description_full': self.desc_full_text.get(1.0, tk.END).strip(),
            'description_masked': self.desc_masked_text.get(1.0, tk.END).strip(),
        }


class ConfigurationDialog(tk.Toplevel):
    """Professional configuration dialog with all application settings."""

    _SECTION_BG = "#e8e8e8"
    _CARD_BG = "#ffffff"
    _ACCENT = "#2563eb"
    _BORDER = "#d0d0d0"

    def __init__(self, parent: tk.Tk, config: AppConfig):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title("Configuration — CV Research Experience Manager")
        self.geometry("720x680")
        self.minsize(680, 580)
        self.resizable(True, True)
        self.configure(bg=self._SECTION_BG)

        self._config = config
        self._defaults = AppConfig(data_root=config.data_root)
        self.result = None  # set to AppConfig on Save

        self._vars: dict = {}
        self._build_ui()
        self._load_values(config)

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.wait_window(self)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=self._ACCENT, height=48)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(
            hdr, text="  Application Settings", font=("Segoe UI", 13, "bold"),
            fg="white", bg=self._ACCENT, anchor="w"
        ).pack(fill=tk.X, padx=10, pady=10)

        # Scrollable body
        container = tk.Frame(self, bg=self._SECTION_BG)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container, bg=self._SECTION_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        self._body = tk.Frame(canvas, bg=self._SECTION_BG)

        self._body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._body, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Enable mouse-wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # --- Sections ---
        self._section_fuzzy()
        self._section_benchmark()
        self._section_year_inference()
        self._section_font()
        self._section_formatting()
        self._section_retention()
        self._section_phase_order()
        self._section_security()

        # Footer buttons
        footer = tk.Frame(self, bg=self._SECTION_BG)
        footer.pack(fill=tk.X, pady=(4, 10), padx=16)

        btn_reset = tk.Button(
            footer, text="Reset to Defaults", font=("Segoe UI", 10),
            bg="#f5f5f5", relief="groove", padx=14, pady=4,
            command=self._on_reset
        )
        btn_reset.pack(side=tk.LEFT)

        btn_cancel = tk.Button(
            footer, text="Cancel", font=("Segoe UI", 10),
            bg="#f5f5f5", relief="groove", padx=14, pady=4,
            command=self._on_cancel
        )
        btn_cancel.pack(side=tk.RIGHT, padx=(6, 0))

        btn_save = tk.Button(
            footer, text="Save", font=("Segoe UI", 10, "bold"),
            bg=self._ACCENT, fg="white", activebackground="#1d4ed8",
            activeforeground="white", relief="flat", padx=20, pady=4,
            command=self._on_save
        )
        btn_save.pack(side=tk.RIGHT)

    # ---- card helper ----
    def _card(self, title: str) -> tk.Frame:
        wrapper = tk.Frame(self._body, bg=self._SECTION_BG)
        wrapper.pack(fill=tk.X, padx=16, pady=(10, 0))

        tk.Label(
            wrapper, text=title, font=("Segoe UI", 11, "bold"),
            bg=self._SECTION_BG, fg="#333333", anchor="w"
        ).pack(fill=tk.X, pady=(0, 4))

        card = tk.Frame(wrapper, bg=self._CARD_BG, bd=1, relief="solid",
                        highlightbackground=self._BORDER, highlightthickness=1)
        card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=self._CARD_BG, padx=14, pady=10)
        inner.pack(fill=tk.X)
        return inner

    def _labeled_spinbox(self, parent, label: str, key: str,
                          from_: int, to: int, width: int = 8,
                          description: str = "") -> None:
        row = tk.Frame(parent, bg=self._CARD_BG)
        row.pack(fill=tk.X, pady=3)
        tk.Label(row, text=label, font=("Segoe UI", 10), bg=self._CARD_BG,
                 width=30, anchor="w").pack(side=tk.LEFT)
        var = tk.IntVar()
        sb = tk.Spinbox(row, from_=from_, to=to, textvariable=var,
                        width=width, font=("Segoe UI", 10), relief="solid", bd=1)
        sb.pack(side=tk.LEFT, padx=(4, 0))
        if description:
            tk.Label(row, text=description, font=("Segoe UI", 9, "italic"),
                     fg="#888888", bg=self._CARD_BG).pack(side=tk.LEFT, padx=(8, 0))
        self._vars[key] = var

    def _labeled_entry(self, parent, label: str, key: str,
                        width: int = 20, description: str = "") -> None:
        row = tk.Frame(parent, bg=self._CARD_BG)
        row.pack(fill=tk.X, pady=3)
        tk.Label(row, text=label, font=("Segoe UI", 10), bg=self._CARD_BG,
                 width=30, anchor="w").pack(side=tk.LEFT)
        var = tk.StringVar()
        ent = tk.Entry(row, textvariable=var, width=width,
                       font=("Segoe UI", 10), relief="solid", bd=1)
        ent.pack(side=tk.LEFT, padx=(4, 0))
        if description:
            tk.Label(row, text=description, font=("Segoe UI", 9, "italic"),
                     fg="#888888", bg=self._CARD_BG).pack(side=tk.LEFT, padx=(8, 0))
        self._vars[key] = var

    def _labeled_check(self, parent, label: str, key: str,
                        description: str = "") -> None:
        row = tk.Frame(parent, bg=self._CARD_BG)
        row.pack(fill=tk.X, pady=3)
        var = tk.BooleanVar()
        cb = tk.Checkbutton(row, text=label, variable=var,
                            font=("Segoe UI", 10), bg=self._CARD_BG,
                            activebackground=self._CARD_BG, anchor="w")
        cb.pack(side=tk.LEFT)
        if description:
            tk.Label(row, text=description, font=("Segoe UI", 9, "italic"),
                     fg="#888888", bg=self._CARD_BG).pack(side=tk.LEFT, padx=(8, 0))
        self._vars[key] = var

    def _labeled_combo(self, parent, label: str, key: str,
                        values: list, width: int = 20,
                        description: str = "") -> None:
        row = tk.Frame(parent, bg=self._CARD_BG)
        row.pack(fill=tk.X, pady=3)
        tk.Label(row, text=label, font=("Segoe UI", 10), bg=self._CARD_BG,
                 width=30, anchor="w").pack(side=tk.LEFT)
        var = tk.StringVar()
        cb = ttk.Combobox(row, textvariable=var, values=values,
                          width=width, state="readonly", font=("Segoe UI", 10))
        cb.pack(side=tk.LEFT, padx=(4, 0))
        if description:
            tk.Label(row, text=description, font=("Segoe UI", 9, "italic"),
                     fg="#888888", bg=self._CARD_BG).pack(side=tk.LEFT, padx=(8, 0))
        self._vars[key] = var

    # ---- sections ----
    def _section_fuzzy(self):
        c = self._card("Fuzzy Matching")
        self._labeled_spinbox(c, "Full description threshold:", "fuzzy_threshold_full",
                              0, 100, description="(0–100)")
        self._labeled_spinbox(c, "Masked description threshold:", "fuzzy_threshold_masked",
                              0, 100, description="(0–100)")

    def _section_benchmark(self):
        c = self._card("Benchmark Calculation")
        self._labeled_spinbox(c, "Minimum study count (step-back):", "benchmark_min_count",
                              1, 100, description="≤ this → step back 1 year")
        self._labeled_check(c, "Auto-find benchmark year", "auto_find_benchmark")
        self._labeled_entry(c, "Manual benchmark year:", "manual_benchmark_year",
                            width=8, description="Leave blank for auto")

    def _section_year_inference(self):
        c = self._card("Year Inference Thresholds")
        self._labeled_spinbox(c, "Full threshold:", "year_inference_full_threshold",
                              0, 100, description="(0–100)")
        self._labeled_spinbox(c, "Masked threshold:", "year_inference_masked_threshold",
                              0, 100, description="(0–100)")

    def _section_font(self):
        c = self._card("Font Settings")
        self._labeled_combo(c, "Font family:", "font_name",
                            values=ALLOWED_FONTS, width=22,
                            description="Applied to output .docx")
        self._labeled_spinbox(c, "Font size (pt):", "font_size", 6, 72)

    def _section_formatting(self):
        c = self._card("Formatting Options")
        self._labeled_check(c, "Highlight inserted studies", "highlight_inserted")
        self._labeled_check(c, "Use track changes", "use_track_changes")
        self._labeled_check(c, "Allow redaction without full match",
                            "allow_redaction_without_full_match")

    def _section_retention(self):
        c = self._card("Retention Policy")
        self._labeled_spinbox(c, "Backup retention (days):", "backup_retention_days",
                              1, 9999, description="Auto-delete older backups")
        self._labeled_spinbox(c, "Log retention (days):", "log_retention_days",
                              1, 9999, description="Auto-delete older logs")

    def _section_phase_order(self):
        c = self._card("Phase Order")
        tk.Label(c, text="Comma-separated list of phases in display order:",
                 font=("Segoe UI", 10), bg=self._CARD_BG, anchor="w").pack(fill=tk.X, pady=(0, 4))
        var = tk.StringVar()
        ent = tk.Entry(c, textvariable=var, font=("Segoe UI", 10),
                       relief="solid", bd=1)
        ent.pack(fill=tk.X, pady=(0, 2))
        tk.Label(c, text='e.g.  Phase I, Phase II–IV',
                 font=("Segoe UI", 9, "italic"), fg="#888888",
                 bg=self._CARD_BG, anchor="w").pack(fill=tk.X)
        self._vars["phase_order"] = var

    def _section_security(self):
        c = self._card("Security & Privacy")
        self._labeled_check(c, "Enable offline guard (block network at startup)",
                            "offline_guard_enabled")

    # ---- load / collect ----
    def _load_values(self, cfg: AppConfig):
        self._vars["fuzzy_threshold_full"].set(cfg.fuzzy_threshold_full)
        self._vars["fuzzy_threshold_masked"].set(cfg.fuzzy_threshold_masked)
        self._vars["benchmark_min_count"].set(cfg.benchmark_min_count)
        self._vars["auto_find_benchmark"].set(cfg.auto_find_benchmark)
        self._vars["manual_benchmark_year"].set(
            str(cfg.manual_benchmark_year) if cfg.manual_benchmark_year is not None else ""
        )
        self._vars["year_inference_full_threshold"].set(cfg.year_inference_full_threshold)
        self._vars["year_inference_masked_threshold"].set(cfg.year_inference_masked_threshold)
        self._vars["font_name"].set(cfg.font_name)
        self._vars["font_size"].set(cfg.font_size)
        self._vars["highlight_inserted"].set(cfg.highlight_inserted)
        self._vars["use_track_changes"].set(cfg.use_track_changes)
        self._vars["allow_redaction_without_full_match"].set(cfg.allow_redaction_without_full_match)
        self._vars["backup_retention_days"].set(cfg.backup_retention_days)
        self._vars["log_retention_days"].set(cfg.log_retention_days)
        self._vars["phase_order"].set(", ".join(cfg.phase_order))
        self._vars["offline_guard_enabled"].set(cfg.offline_guard_enabled)

    def _collect(self) -> dict:
        """Collect values from widgets into a config dict. Raises ValueError on bad input."""
        d = {}
        d["fuzzy_threshold_full"] = self._vars["fuzzy_threshold_full"].get()
        d["fuzzy_threshold_masked"] = self._vars["fuzzy_threshold_masked"].get()
        d["benchmark_min_count"] = self._vars["benchmark_min_count"].get()
        d["auto_find_benchmark"] = self._vars["auto_find_benchmark"].get()

        mbr = self._vars["manual_benchmark_year"].get().strip()
        d["manual_benchmark_year"] = int(mbr) if mbr else None

        d["year_inference_full_threshold"] = self._vars["year_inference_full_threshold"].get()
        d["year_inference_masked_threshold"] = self._vars["year_inference_masked_threshold"].get()
        d["font_name"] = self._vars["font_name"].get().strip()
        d["font_size"] = self._vars["font_size"].get()
        d["highlight_inserted"] = self._vars["highlight_inserted"].get()
        d["use_track_changes"] = self._vars["use_track_changes"].get()
        d["allow_redaction_without_full_match"] = self._vars["allow_redaction_without_full_match"].get()
        d["backup_retention_days"] = self._vars["backup_retention_days"].get()
        d["log_retention_days"] = self._vars["log_retention_days"].get()
        d["offline_guard_enabled"] = self._vars["offline_guard_enabled"].get()

        raw_phases = self._vars["phase_order"].get()
        d["phase_order"] = [p.strip() for p in raw_phases.split(",") if p.strip()]

        d["data_root"] = self._config.data_root
        d["network_enabled"] = False
        d["user_id_strategy"] = self._config.user_id_strategy
        return d

    # ---- actions ----
    def _on_save(self):
        try:
            d = self._collect()
            new_cfg = AppConfig(**d)  # triggers validation
            new_cfg.save()
            self.result = new_cfg
            self.destroy()
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid Configuration", str(exc), parent=self)

    def _on_reset(self):
        if messagebox.askyesno("Reset", "Reset all settings to defaults?", parent=self):
            self._load_values(self._defaults)

    def _on_cancel(self):
        self.result = None
        self.destroy()


class ReadmeViewer(tk.Toplevel):
    """Scrollable README.md viewer with basic Markdown rendering."""

    _BG = "#ffffff"
    _FG = "#24292f"
    _ACCENT = "#2563eb"
    _CODE_BG = "#f6f8fa"
    _BORDER = "#d0d7de"
    _H_COLORS = ["#24292f", "#24292f", "#24292f", "#57606a"]
    _FONT = "Segoe UI"
    _MONO = "Consolas"

    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.transient(parent)
        self.title("About — CV Research Experience Manager")
        self.geometry("900x700")
        self.minsize(700, 500)
        self.configure(bg=self._BG)

        self._build_ui()
        self._render_readme()

    def _build_ui(self):
        # Header bar
        hdr = tk.Frame(self, bg=self._ACCENT, height=48)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(
            hdr, text="  README.md", font=(self._FONT, 13, "bold"),
            fg="white", bg=self._ACCENT, anchor="w"
        ).pack(fill=tk.X, padx=10, pady=10)

        # Text widget
        container = tk.Frame(self, bg=self._BG)
        container.pack(fill=tk.BOTH, expand=True)

        self._text = tk.Text(
            container, wrap=tk.WORD, bg=self._BG, fg=self._FG,
            font=(self._FONT, 11), padx=28, pady=20,
            relief="flat", cursor="arrow", state="disabled",
            spacing1=2, spacing3=2, selectbackground="#c8d8f8"
        )
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self._text.yview)
        self._text.configure(yscrollcommand=scrollbar.set)

        self._text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Configure tags
        self._text.tag_configure("h1", font=(self._FONT, 22, "bold"), foreground=self._H_COLORS[0],
                                 spacing1=16, spacing3=8)
        self._text.tag_configure("h2", font=(self._FONT, 17, "bold"), foreground=self._H_COLORS[1],
                                 spacing1=14, spacing3=6)
        self._text.tag_configure("h3", font=(self._FONT, 14, "bold"), foreground=self._H_COLORS[2],
                                 spacing1=10, spacing3=4)
        self._text.tag_configure("h4", font=(self._FONT, 12, "bold"), foreground=self._H_COLORS[3],
                                 spacing1=8, spacing3=3)
        self._text.tag_configure("bold", font=(self._FONT, 11, "bold"))
        self._text.tag_configure("italic", font=(self._FONT, 11, "italic"))
        self._text.tag_configure("bold_italic", font=(self._FONT, 11, "bold italic"))
        self._text.tag_configure("code_inline", font=(self._MONO, 10), background="#eff1f3",
                                 foreground="#cf222e")
        self._text.tag_configure("code_block", font=(self._MONO, 10), background=self._CODE_BG,
                                 foreground="#24292f", lmargin1=30, lmargin2=30,
                                 rmargin=30, spacing1=1, spacing3=1)
        self._text.tag_configure("hr", font=(self._FONT, 4), foreground=self._BORDER,
                                 spacing1=10, spacing3=10)
        self._text.tag_configure("bullet", lmargin1=24, lmargin2=44,
                                 font=(self._FONT, 11))
        self._text.tag_configure("table_header", font=(self._FONT, 10, "bold"),
                                 background="#f0f3f6")
        self._text.tag_configure("table_row", font=(self._FONT, 10),
                                 background=self._BG)
        self._text.tag_configure("table_border", font=(self._FONT, 10),
                                 foreground=self._BORDER)
        self._text.tag_configure("link", foreground=self._ACCENT,
                                 font=(self._FONT, 11, "underline"))
        self._text.tag_configure("blockquote", foreground="#57606a",
                                 font=(self._FONT, 11, "italic"),
                                 lmargin1=30, lmargin2=30)

        # Close button
        footer = tk.Frame(self, bg=self._BG)
        footer.pack(fill=tk.X, pady=(0, 10))
        tk.Button(
            footer, text="Close", font=(self._FONT, 10),
            bg="#f5f5f5", relief="groove", padx=18, pady=4,
            command=self.destroy
        ).pack()

    def _render_readme(self):
        """Load and render README.md."""
        readme_path = get_app_root() / "README.md"
        if not readme_path.exists():
            self._text.configure(state="normal")
            self._text.insert("end", "README.md not found.", "italic")
            self._text.configure(state="disabled")
            return

        with open(readme_path, "r", encoding="utf-8") as f:
            content = f.read()

        self._text.configure(state="normal")
        self._text.delete("1.0", "end")

        in_code_block = False
        code_lines: list = []
        lines = content.split("\n")
        i = 0
        while i < len(lines):
            line = lines[i]

            # Code block fence
            if line.strip().startswith("```"):
                if in_code_block:
                    # End code block
                    block = "\n".join(code_lines)
                    if block:
                        self._text.insert("end", block + "\n", "code_block")
                    self._text.insert("end", "\n")
                    code_lines = []
                    in_code_block = False
                else:
                    in_code_block = True
                    code_lines = []
                i += 1
                continue

            if in_code_block:
                code_lines.append(line)
                i += 1
                continue

            stripped = line.strip()

            # Horizontal rule
            if stripped in ("---", "***", "___"):
                self._text.insert("end", "─" * 80 + "\n", "hr")
                i += 1
                continue

            # Headings
            heading_match = re.match(r"^(#{1,4})\s+(.*)", line)
            if heading_match:
                level = len(heading_match.group(1))
                text = heading_match.group(2).strip()
                tag = f"h{level}"
                self._text.insert("end", text + "\n", tag)
                i += 1
                continue

            # Table detection
            if "|" in line and i + 1 < len(lines) and re.match(r"^\s*\|[-\s|:]+\|\s*$", lines[i + 1]):
                # Render table
                self._render_table(lines, i)
                # Skip past the table
                while i < len(lines) and "|" in lines[i]:
                    i += 1
                self._text.insert("end", "\n")
                continue

            # Bullet list
            bullet_match = re.match(r"^(\s*)[-*]\s+(.*)", line)
            if bullet_match:
                indent_level = len(bullet_match.group(1)) // 2
                text = bullet_match.group(2)
                prefix = "  " * indent_level + "•  "
                self._insert_inline(prefix, text, "bullet")
                i += 1
                continue

            # Numbered list
            num_match = re.match(r"^(\s*)\d+\.\s+(.*)", line)
            if num_match:
                text = num_match.group(2)
                indent = num_match.group(1)
                num_prefix = line[:line.index(num_match.group(2))]
                self._insert_inline(indent + num_prefix.strip() + " ", text, "bullet")
                i += 1
                continue

            # Blank line
            if not stripped:
                self._text.insert("end", "\n")
                i += 1
                continue

            # Normal paragraph
            self._insert_inline("", stripped, None)
            i += 1

        self._text.configure(state="disabled")

    def _render_table(self, lines: list, start: int):
        """Render a markdown table starting at line index *start*.

        Uses a card-style layout: each data row is rendered as a
        block with header labels, which handles wide tables gracefully.
        For narrow tables (max cell <= 40 chars) a traditional grid is used.
        """
        # Parse header row
        header_line = lines[start].strip().strip("|")
        headers = [c.strip() for c in header_line.split("|")]
        # Separator (skip)
        # Data rows
        data_rows = []
        j = start + 2
        while j < len(lines) and "|" in lines[j]:
            row_line = lines[j].strip().strip("|")
            cells = [c.strip() for c in row_line.split("|")]
            data_rows.append(cells)
            j += 1

        col_count = len(headers)

        # Determine max cell width to decide rendering mode
        max_cell = max(
            (len(c) for row in data_rows for c in row),
            default=0
        )
        max_cell = max(max_cell, max((len(h) for h in headers), default=0))

        if max_cell <= 50:
            # --- Compact grid mode for narrow tables ---
            widths = [len(h) for h in headers]
            for row in data_rows:
                for ci in range(min(col_count, len(row))):
                    widths[ci] = max(widths[ci], len(row[ci]))
            widths = [min(w + 2, 52) for w in widths]

            header_text = " | ".join(h.ljust(widths[ci]) for ci, h in enumerate(headers))
            self._text.insert("end", " " + header_text + "\n", "table_header")

            sep_text = " | ".join("\u2500" * widths[ci] for ci in range(col_count))
            self._text.insert("end", " " + sep_text + "\n", "table_border")

            for row in data_rows:
                cells = []
                for ci in range(col_count):
                    val = row[ci] if ci < len(row) else ""
                    cells.append(val.ljust(widths[ci]))
                row_text = " | ".join(cells)
                self._text.insert("end", " " + row_text + "\n", "table_row")
        else:
            # --- Card mode for wide tables ---
            # Render header row as bold labels
            self._text.insert("end", "\n")
            for ri, row in enumerate(data_rows):
                for ci in range(col_count):
                    val = row[ci] if ci < len(row) else ""
                    if not val:
                        continue
                    hdr = headers[ci] if ci < len(headers) else ""
                    self._text.insert("end", f"  {hdr}: ", "bold")
                    self._insert_inline("", val, None)
                # Separator between rows
                if ri < len(data_rows) - 1:
                    self._text.insert("end", "  " + "\u2500" * 70 + "\n", "table_border")

    def _insert_inline(self, prefix: str, text: str, base_tag):
        """Insert text with inline markdown: **bold**, *italic*, `code`, [links]."""
        if prefix:
            self._text.insert("end", prefix, base_tag)

        # Regex to find inline patterns
        pattern = re.compile(
            r"(\*\*\*(.+?)\*\*\*)"   # ***bold italic***
            r"|(\*\*(.+?)\*\*)"       # **bold**
            r"|(\*(.+?)\*)"           # *italic*
            r"|(`([^`]+)`)"           # `code`
            r"|(\[([^\]]+)\]\([^\)]+\))"  # [link](url)
        )

        pos = 0
        for m in pattern.finditer(text):
            # Insert text before match
            if m.start() > pos:
                self._text.insert("end", text[pos:m.start()], base_tag)

            if m.group(2):  # bold italic
                tag = "bold_italic" if not base_tag else (base_tag, "bold_italic")
                self._text.insert("end", m.group(2), tag)
            elif m.group(4):  # bold
                tag = "bold" if not base_tag else (base_tag, "bold")
                self._text.insert("end", m.group(4), tag)
            elif m.group(6):  # italic
                tag = "italic" if not base_tag else (base_tag, "italic")
                self._text.insert("end", m.group(6), tag)
            elif m.group(8):  # code
                self._text.insert("end", m.group(8), "code_inline")
            elif m.group(10):  # link text
                self._text.insert("end", m.group(10), "link")

            pos = m.end()

        # Remainder
        if pos < len(text):
            self._text.insert("end", text[pos:], base_tag)

        self._text.insert("end", "\n", base_tag)


def main():
    """Main entry point."""
    root = tk.Tk()
    app = CVManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
