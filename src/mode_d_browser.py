"""
Mode D: Study Browser - Separate module for the study browser tab.
Provides filtering, formatting, and copying functionality for studies.
"""

import tkinter as tk
from tkinter import ttk, colorchooser
from typing import List, Dict, Optional, Set
import logging

from database import DatabaseManager
from config import ALLOWED_FONTS, HANGING_INDENT_MIN, HANGING_INDENT_MAX
from user_preferences import UserPreferences

browser_logger = logging.getLogger(__name__)

# Debounce delay for search
SEARCH_DEBOUNCE_MS = 300


class StudyBrowserTab:
    """Handles Mode D: Study Browser functionality."""
    
    def __init__(self, parent_frame, config, prefs_manager, prefs):
        self.parent = parent_frame
        self.config = config
        self.prefs_manager = prefs_manager
        self.prefs = prefs

        # State
        self.selected_site_id: Optional[int] = prefs.mode_d_selected_site_id
        self.all_studies: List[Dict] = []
        self.filtered_studies: List[Dict] = []
        self.selected_indices: Set[int] = set()

        # Filter state
        self.filter_years: Set[int] = set(prefs.mode_d_filter_years)
        self.filter_phases: Set[str] = set(prefs.mode_d_filter_phases)
        self.filter_subcategories: Set[str] = set(prefs.mode_d_filter_subcategories)

        # Debounce timers
        self._search_timer = None
        self._display_timer = None

        # Virtual scrolling state
        self._visible_rows: List[int] = []  # indices of currently visible rows
        self._ROW_BATCH_SIZE = 50  # Render rows in batches for smoothness
        self._SCROLL_DELAY_MS = 10  # Delay between batch renders

        self._create_ui()
        self._load_sites()
        if self.selected_site_id:
            self._load_studies()
    
    def _create_ui(self):
        """Create the Mode D UI layout."""
        # Main horizontal paned window
        main_paned = ttk.PanedWindow(self.parent, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True)
        
        # Left panel (sites + controls) - auto-size to content
        left_panel = ttk.Frame(main_paned)
        main_paned.add(left_panel, weight=0)
        
        # Right panel (study list)
        right_panel = ttk.Frame(main_paned)
        main_paned.add(right_panel, weight=1)
        
        self._create_left_panel(left_panel)
        self._create_right_panel(right_panel)
    
    def _create_left_panel(self, parent):
        """Create left panel with sites list and controls."""
        # Sites list
        sites_label = ttk.Label(parent, text="Sites", font=('Segoe UI', 12, 'bold'))
        sites_label.pack(anchor=tk.W, padx=5, pady=(5, 2))
        
        sites_frame = ttk.Frame(parent)
        sites_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.sites_listbox = tk.Listbox(sites_frame, height=6, width=25, exportselection=False)
        self.sites_listbox.pack(side=tk.LEFT, fill=tk.BOTH)
        self.sites_listbox.bind('<<ListboxSelect>>', self._on_site_select)
        
        sites_scroll = ttk.Scrollbar(sites_frame, orient=tk.VERTICAL, command=self.sites_listbox.yview)
        sites_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.sites_listbox.config(yscrollcommand=sites_scroll.set)
        
        # Multi-Search
        search_label = ttk.Label(parent, text="Multi-Search (one per line)", font=('Segoe UI', 10, 'bold'))
        search_label.pack(anchor=tk.W, padx=5, pady=(10, 2))
        
        multi_search_frame = ttk.Frame(parent)
        multi_search_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.multi_search_text = tk.Text(multi_search_frame, height=4, width=25, wrap=tk.WORD)
        self.multi_search_text.pack(side=tk.LEFT, fill=tk.BOTH)
        self.multi_search_text.insert('1.0', self.prefs.mode_d_multi_search)
        self.multi_search_text.bind('<KeyRelease>', lambda e: self._on_search_change_debounced())
        
        multi_scroll = ttk.Scrollbar(multi_search_frame, orient=tk.VERTICAL, command=self.multi_search_text.yview)
        multi_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.multi_search_text.config(yscrollcommand=multi_scroll.set)
        
        # Clear button
        ttk.Button(parent, text="Clear Multi-Search", command=self._clear_multi_search).pack(fill=tk.X, padx=5, pady=(0, 5))
        
        # Formatting controls
        format_frame = ttk.LabelFrame(parent, text="Formatting", padding="5")
        format_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Hanging indent
        indent_frame = ttk.Frame(format_frame)
        indent_frame.pack(fill=tk.X, pady=2)
        ttk.Label(indent_frame, text="Hanging Indent:").pack(side=tk.LEFT)
        self.indent_var = tk.DoubleVar(value=self.prefs.mode_d_hanging_indent)
        indent_spin = ttk.Spinbox(indent_frame, from_=HANGING_INDENT_MIN, to=HANGING_INDENT_MAX,
                                   increment=0.1, textvariable=self.indent_var, width=8)
        indent_spin.pack(side=tk.LEFT, padx=5)
        indent_spin.bind('<FocusOut>', lambda e: self._mark_format_dirty())
        
        # Font name
        font_frame = ttk.Frame(format_frame)
        font_frame.pack(fill=tk.X, pady=2)
        ttk.Label(font_frame, text="Font:").pack(side=tk.LEFT)
        self.font_var = tk.StringVar(value=self.prefs.mode_d_font_name)
        font_combo = ttk.Combobox(font_frame, textvariable=self.font_var, values=ALLOWED_FONTS,
                                   state='readonly', width=15)
        font_combo.pack(side=tk.LEFT, padx=5)
        font_combo.bind('<<ComboboxSelected>>', lambda e: self._mark_format_dirty())
        
        # Font size
        size_frame = ttk.Frame(format_frame)
        size_frame.pack(fill=tk.X, pady=2)
        ttk.Label(size_frame, text="Font Size:").pack(side=tk.LEFT)
        self.font_size_var = tk.IntVar(value=self.prefs.mode_d_font_size)
        size_spin = ttk.Spinbox(size_frame, from_=8, to=20, textvariable=self.font_size_var, width=8)
        size_spin.pack(side=tk.LEFT, padx=5)
        size_spin.bind('<FocusOut>', lambda e: self._mark_format_dirty())
        
        # Sponsor bold
        self.sponsor_bold_var = tk.BooleanVar(value=self.prefs.mode_d_sponsor_bold)
        sponsor_check = ttk.Checkbutton(format_frame, text="Bold Sponsor",
                                        variable=self.sponsor_bold_var,
                                        command=self._mark_format_dirty)
        sponsor_check.pack(anchor=tk.W, pady=2)
        
        # Sponsor color
        sponsor_color_frame = ttk.Frame(format_frame)
        sponsor_color_frame.pack(fill=tk.X, pady=2)
        ttk.Label(sponsor_color_frame, text="Sponsor Color:").pack(side=tk.LEFT)
        self.sponsor_color_var = tk.StringVar(value=self.prefs.mode_d_sponsor_color)
        ttk.Button(sponsor_color_frame, text="Choose", 
                   command=lambda: self._choose_color(self.sponsor_color_var)).pack(side=tk.LEFT, padx=5)
        self.sponsor_color_preview = tk.Label(sponsor_color_frame, text="  ", 
                                               bg=self.prefs.mode_d_sponsor_color, width=3, relief=tk.SUNKEN)
        self.sponsor_color_preview.pack(side=tk.LEFT)
        
        # Protocol color
        protocol_color_frame = ttk.Frame(format_frame)
        protocol_color_frame.pack(fill=tk.X, pady=2)
        ttk.Label(protocol_color_frame, text="Protocol Color:").pack(side=tk.LEFT)
        self.protocol_color_var = tk.StringVar(value=self.prefs.mode_d_protocol_color)
        ttk.Button(protocol_color_frame, text="Choose",
                   command=lambda: self._choose_color(self.protocol_color_var)).pack(side=tk.LEFT, padx=5)
        self.protocol_color_preview = tk.Label(protocol_color_frame, text="  ",
                                                bg=self.prefs.mode_d_protocol_color, width=3, relief=tk.SUNKEN)
        self.protocol_color_preview.pack(side=tk.LEFT)

        # Save button
        ttk.Button(format_frame, text="Save Format Settings", command=self._save_format_prefs).pack(fill=tk.X, pady=(10, 2))

        # Track pending changes
        self._format_pending = False

    def _mark_format_dirty(self):
        """Mark formatting as having pending changes."""
        self._format_pending = True

    def _create_right_panel(self, parent):
        """Create right panel with study list and controls."""
        # Top controls
        controls_frame = ttk.Frame(parent)
        controls_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Single search
        ttk.Label(controls_frame, text="Quick Search:").pack(side=tk.LEFT)
        self.single_search_var = tk.StringVar(value=self.prefs.mode_d_single_search)
        self._last_single_search = self.prefs.mode_d_single_search
        single_search_entry = ttk.Entry(controls_frame, textvariable=self.single_search_var, width=30)
        single_search_entry.pack(side=tk.LEFT, padx=5)
        # Use key release with debounce instead of trace for better performance
        single_search_entry.bind('<KeyRelease>', lambda e: self._on_single_search_debounced())

        # Show/Hide Protocol toggle
        self.show_protocol_var = tk.BooleanVar(value=self.prefs.mode_d_show_protocol)
        protocol_toggle = ttk.Checkbutton(controls_frame, text="Show Protocol",
                                          variable=self.show_protocol_var,
                                          command=self._refresh_display_lazy)
        protocol_toggle.pack(side=tk.LEFT, padx=10)

        # Filter button
        ttk.Button(controls_frame, text="Filter", command=self._show_filter_dialog).pack(side=tk.LEFT, padx=5)

        # Reset button - clears filters and search
        ttk.Button(controls_frame, text="Reset", command=self._reset_filters_and_search).pack(side=tk.LEFT, padx=5)
        
        # Select All button
        ttk.Button(controls_frame, text="Select All", command=self._select_all).pack(side=tk.LEFT, padx=5)
        
        # Copy All button
        ttk.Button(controls_frame, text="Copy All", command=self._copy_selected).pack(side=tk.LEFT, padx=5)
        
        # Study list frame
        list_frame = ttk.Frame(parent)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Canvas for scrolling
        self.canvas = tk.Canvas(list_frame, bg='white')
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.canvas.config(yscrollcommand=scrollbar.set)
        
        # Frame inside canvas
        self.studies_container = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.studies_container, anchor=tk.NW)
        
        # Bind resize
        self.studies_container.bind('<Configure>', lambda e: self.canvas.config(scrollregion=self.canvas.bbox('all')))
        self.canvas.bind('<Configure>', self._on_canvas_resize)
    
    def _on_canvas_resize(self, event):
        """Handle canvas resize to adjust inner frame width."""
        self.canvas.itemconfig(self.canvas_window, width=event.width)
    
    def _load_sites(self):
        """Load sites into listbox."""
        self.sites_listbox.delete(0, tk.END)
        try:
            with DatabaseManager(config=self.config) as db:
                sites = db.get_sites()
                self._sites = sites
                for i, site in enumerate(sites):
                    count = db.get_study_count(site.id)
                    self.sites_listbox.insert(tk.END, f"{site.name} ({count} studies)")
                    if site.id == self.selected_site_id:
                        self.sites_listbox.selection_set(i)
        except Exception as e:
            browser_logger.error(f"Failed to load sites: {e}")
    
    def _on_site_select(self, event):
        """Handle site selection."""
        selection = self.sites_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        if hasattr(self, '_sites') and idx < len(self._sites):
            site = self._sites[idx]
            self.selected_site_id = site.id
            self.prefs.mode_d_selected_site_id = site.id
            self.prefs_manager.save(self.prefs)
            self._load_studies()
    
    def _load_studies(self):
        """Load studies from selected site."""
        if not self.selected_site_id:
            return
        
        try:
            with DatabaseManager(config=self.config) as db:
                studies = db.get_studies(self.selected_site_id)
                category_order = db.get_category_order(self.selected_site_id) or []
                
                # Convert to dict format
                self.all_studies = []
                for study in studies:
                    self.all_studies.append({
                        'id': study.id,
                        'phase': study.phase,
                        'subcategory': study.subcategory,
                        'year': study.year,
                        'sponsor': study.sponsor,
                        'protocol': study.protocol or '',
                        'description_full': study.description_full,
                        'description_masked': study.description_masked,
                        'category_key': f"{study.phase} > {study.subcategory}"
                    })
                
                # Sort by category order, then year (desc)
                def sort_key(s):
                    cat_key = s['category_key']
                    cat_index = category_order.index(cat_key) if cat_key in category_order else 9999
                    return (cat_index, -s['year'], s['sponsor'], s['protocol'])
                
                self.all_studies.sort(key=sort_key)
                
                self._apply_filters()
        except Exception as e:
            browser_logger.error(f"Failed to load studies: {e}")
    
    def _on_search_change_debounced(self):
        """Debounced search to prevent lag while typing."""
        if self._search_timer:
            self.parent.after_cancel(self._search_timer)
        self._search_timer = self.parent.after(SEARCH_DEBOUNCE_MS, self._on_search_change)

    def _on_single_search_debounced(self):
        """Debounced single search to prevent lag while typing."""
        current = self.single_search_var.get()
        if current == self._last_single_search:
            return  # No change, skip
        self._last_single_search = current

        if self._search_timer:
            self.parent.after_cancel(self._search_timer)
        self._search_timer = self.parent.after(SEARCH_DEBOUNCE_MS, self._on_search_change)
    
    def _on_search_change(self):
        """Handle search input changes."""
        # Save search terms
        self.prefs.mode_d_single_search = self.single_search_var.get()
        self.prefs.mode_d_multi_search = self.multi_search_text.get('1.0', tk.END).strip()
        self.prefs_manager.save(self.prefs)
        
        self._apply_filters()
    
    def _clear_multi_search(self):
        """Clear multi-search text."""
        self.multi_search_text.delete('1.0', tk.END)
        self._on_search_change()
    
    def _reset_filters_and_search(self):
        """Reset all filters and search terms."""
        # Clear search
        self.single_search_var.set('')
        self._last_single_search = ''
        self.multi_search_text.delete('1.0', tk.END)

        # Clear filters
        self.filter_years.clear()
        self.filter_phases.clear()
        self.filter_subcategories.clear()

        # Save to preferences
        self.prefs.mode_d_single_search = ''
        self.prefs.mode_d_multi_search = ''
        self.prefs.mode_d_filter_years = []
        self.prefs.mode_d_filter_phases = []
        self.prefs.mode_d_filter_subcategories = []
        self.prefs_manager.save(self.prefs)

        # Refresh display
        self._apply_filters()
        browser_logger.info("Reset all filters and search terms")

    def _apply_filters(self):
        """Apply all filters and refresh display."""
        # Filter in memory first (fast)
        filtered = []

        # Get search terms
        single_search = self.single_search_var.get().strip().lower()
        multi_search_text = self.multi_search_text.get('1.0', tk.END).strip()
        multi_keywords = [kw.strip().lower() for kw in multi_search_text.split('\n') if kw.strip()]

        for study in self.all_studies:
            # Apply year filter
            if self.filter_years and study['year'] not in self.filter_years:
                continue

            # Apply phase filter
            if self.filter_phases and study['phase'] not in self.filter_phases:
                continue

            # Apply subcategory filter
            if self.filter_subcategories and study['subcategory'] not in self.filter_subcategories:
                continue

            # Apply search filters (single search takes priority)
            if single_search:
                # Fast early exit: check only key fields first
                if single_search in study['sponsor'].lower() or \
                   single_search in study['protocol'].lower() or \
                   single_search in str(study['year']).lower():
                    filtered.append(study)
                    continue
                # Full search if early exit fails
                searchable = f"{study['phase']} {study['subcategory']} {study['year']} {study['sponsor']} {study['protocol']} {study['description_full']} {study['description_masked']}".lower()
                if single_search in searchable:
                    filtered.append(study)
            elif multi_keywords:
                # Study must match ANY keyword (OR logic)
                searchable = f"{study['phase']} {study['subcategory']} {study['year']} {study['sponsor']} {study['protocol']} {study['description_full']} {study['description_masked']}".lower()
                if any(kw in searchable for kw in multi_keywords):
                    filtered.append(study)
            else:
                filtered.append(study)

        self.filtered_studies = filtered
        self._refresh_display_lazy()
    
    def _refresh_display(self):
        """Refresh the study display immediately (for small lists)."""
        self._refresh_display_lazy()

    def _refresh_display_lazy(self):
        """Refresh display with lazy loading for smooth performance."""
        # Cancel any pending display update
        if self._display_timer:
            self.parent.after_cancel(self._display_timer)

        # Clear container efficiently
        for widget in self.studies_container.winfo_children():
            widget.destroy()

        self.selected_indices.clear()
        self._visible_rows = []

        total = len(self.filtered_studies)
        if total == 0:
            self.studies_container.update_idletasks()
            self.canvas.config(scrollregion=self.canvas.bbox('all'))
            return

        # For small lists, render all at once
        if total <= self._ROW_BATCH_SIZE:
            for i, study in enumerate(self.filtered_studies):
                self._create_study_row(i, study)
            self.studies_container.update_idletasks()
            self.canvas.config(scrollregion=self.canvas.bbox('all'))
        else:
            # For large lists, render in batches for responsiveness
            self._render_batch(0, self._ROW_BATCH_SIZE)

    def _render_batch(self, start: int, end: int):
        """Render a batch of study rows."""
        batch_end = min(end, len(self.filtered_studies))

        for i in range(start, batch_end):
            self._create_study_row(i, self.filtered_studies[i])

        # Update scroll region after each batch
        self.studies_container.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox('all'))

        # Schedule next batch if there are more rows
        if batch_end < len(self.filtered_studies):
            self._display_timer = self.parent.after(
                self._SCROLL_DELAY_MS,
                lambda: self._render_batch(batch_end, batch_end + self._ROW_BATCH_SIZE)
            )
    
    def _create_study_row(self, index: int, study: Dict):
        """Create a single study row with formatting colors."""
        row_frame = ttk.Frame(self.studies_container)
        row_frame.pack(fill=tk.X, pady=1, padx=2)
        
        # Checkbox
        var = tk.BooleanVar(value=index in self.selected_indices)
        check = ttk.Checkbutton(row_frame, variable=var, 
                                command=lambda: self._toggle_selection(index, var.get()))
        check.pack(side=tk.LEFT, padx=(5, 2))
        
        # Copy button (right next to checkbox) - small like checkbox
        copy_label = tk.Label(row_frame, text="📋", cursor="hand2", font=('Segoe UI', 9))
        copy_label.pack(side=tk.LEFT, padx=(0, 5))
        copy_label.bind('<Button-1>', lambda e, s=study: self._copy_single(s))
        
        # Study text with colors using Text widget
        text_widget = tk.Text(row_frame, height=1, wrap=tk.NONE, relief=tk.FLAT,
                             borderwidth=0, highlightthickness=0,
                             font=(self.font_var.get(), self.font_size_var.get()))
        text_widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Configure color tags
        sponsor_color = self.sponsor_color_var.get()
        protocol_color = self.protocol_color_var.get()
        sponsor_bold = self.sponsor_bold_var.get()
        font_name = self.font_var.get()
        font_size = self.font_size_var.get()
        
        text_widget.tag_configure('sponsor', foreground=sponsor_color, 
                                 font=(font_name, font_size, 'bold' if sponsor_bold else 'normal'))
        text_widget.tag_configure('protocol', foreground=protocol_color, 
                                 font=(font_name, font_size))
        text_widget.tag_configure('normal', foreground='#000000', 
                                 font=(font_name, font_size))
        
        # Build text with tags
        year = str(study['year'])
        sponsor = study['sponsor']
        protocol = study['protocol']
        indent_inches = self.indent_var.get()
        show_protocol = self.show_protocol_var.get()
        
        # Add year
        text_widget.insert('end', year, 'normal')
        
        # Add indent
        if indent_inches > 0:
            text_widget.insert('end', '\t', 'normal')
        else:
            text_widget.insert('end', ' ', 'normal')
        
        # Add sponsor with color
        text_widget.insert('end', sponsor, 'sponsor')
        text_widget.insert('end', ' ', 'normal')
        
        # Add protocol
        if show_protocol and protocol:
            text_widget.insert('end', protocol, 'protocol')
            text_widget.insert('end', ': ', 'normal')
            desc = study['description_full']
        else:
            text_widget.insert('end', ': ', 'normal')
            desc = study['description_masked']
        
        # Add description
        text_widget.insert('end', desc, 'normal')
        
        # Make read-only
        text_widget.config(state=tk.DISABLED)
    
    def _format_study(self, study: Dict) -> str:
        """Format a study for display."""
        indent = "    " * int(self.indent_var.get() * 2)  # Approximate indent
        year = study['year']
        sponsor = study['sponsor']
        protocol = study['protocol']
        show_protocol = self.show_protocol_var.get()

        if show_protocol and protocol:
            desc = study['description_full']
            return f"{year}{indent}{sponsor} {protocol}: {desc}"
        else:
            desc = study['description_masked']
            return f"{year}{indent}{sponsor}: {desc}"
    
    def _toggle_selection(self, index: int, selected: bool):
        """Toggle study selection."""
        if selected:
            self.selected_indices.add(index)
        else:
            self.selected_indices.discard(index)
    
    def _select_all(self):
        """Select all visible studies."""
        self.selected_indices = set(range(len(self.filtered_studies)))
        self._refresh_display()
        
        # Re-check all checkboxes
        for widget in self.studies_container.winfo_children():
            for child in widget.winfo_children():
                if isinstance(child, ttk.Checkbutton):
                    child.invoke()
                    break
    
    def _copy_single(self, study: Dict):
        """Copy a single study to clipboard with formatting."""
        try:
            import win32clipboard
            
            # Get formatting settings
            font_name = self.font_var.get()
            font_size = self.font_size_var.get()
            sponsor_color = self.sponsor_color_var.get()
            protocol_color = self.protocol_color_var.get()
            sponsor_bold = self.sponsor_bold_var.get()
            indent_inches = self.indent_var.get()
            show_protocol = self.show_protocol_var.get()
            
            # Build study data
            year = str(study['year'])
            sponsor = study['sponsor']
            protocol = study['protocol']
            
            # Calculate indent in pixels (96 DPI standard)
            indent_px = int(indent_inches * 96)
            
            # Build HTML body with proper formatting
            html_body = '<html><head><meta charset="utf-8"></head><body>'
            
            # Paragraph with hanging indent
            html_body += f'<p style="font-family:{font_name};font-size:{font_size}pt;'
            if indent_px > 0:
                html_body += f'margin-left:{indent_px}px;text-indent:-{indent_px}px;'
            html_body += 'margin-top:0;margin-bottom:0;">'
            
            # Year followed by tab (use &#9; for HTML tab that Word recognizes)
            html_body += f'{year}'

            # Tab character as HTML entity - Word recognizes &#9; as a tab
            if indent_inches > 0:
                html_body += '&#9;'  # HTML tab entity
            else:
                html_body += ' '

            # Sponsor with color and bold
            sponsor_style = f'color:{sponsor_color};'
            if sponsor_bold:
                sponsor_style += 'font-weight:bold;'
            html_body += f'<span style="{sponsor_style}">{sponsor}</span> '

            # Protocol
            if show_protocol and protocol:
                html_body += f'<span style="color:{protocol_color};">{protocol}</span>: '
                desc = study['description_full']
            else:
                html_body += ': '
                desc = study['description_masked']

            # Description
            html_body += desc
            html_body += '</p></body></html>'
            
            # Create CF_HTML format with proper header
            html_utf8 = html_body.encode('utf-8')
            
            # CF_HTML requires specific header format
            html_prefix = "Version:0.9\r\nStartHTML:00000000\r\nEndHTML:00000000\r\nStartFragment:00000000\r\nEndFragment:00000000\r\n"
            html_with_markers = html_prefix + html_body
            
            # Calculate byte positions
            start_html = len(html_prefix)
            end_html = len(html_with_markers)
            start_fragment = html_with_markers.find('<body>') + 6
            end_fragment = html_with_markers.find('</body>')
            
            # Format header with positions
            cf_html_data = f"Version:0.9\r\nStartHTML:{start_html:08d}\r\nEndHTML:{end_html:08d}\r\nStartFragment:{start_fragment:08d}\r\nEndFragment:{end_fragment:08d}\r\n{html_body}"
            
            # Plain text fallback with tab
            if indent_inches > 0:
                text = f"{year}\t{sponsor} "
            else:
                text = f"{year} {sponsor} "
            
            if show_protocol and protocol:
                text += f"{protocol}: {study['description_full']}"
            else:
                text += f": {study['description_masked']}"
            
            # Copy to clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            
            # Set plain text
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
            
            # Set HTML format
            cf_html = win32clipboard.RegisterClipboardFormat('HTML Format')
            win32clipboard.SetClipboardData(cf_html, cf_html_data.encode('utf-8'))
            
            win32clipboard.CloseClipboard()
            browser_logger.info(f"Copied study with formatting: {study['sponsor']} {study['protocol']}")
            
        except ImportError:
            # Fallback if win32clipboard not available
            text = self._format_study(study)
            self.canvas.clipboard_clear()
            self.canvas.clipboard_append(text)
            browser_logger.warning("win32clipboard not available, copied as plain text")
        except Exception as e:
            # Fallback on any error
            text = self._format_study(study)
            self.canvas.clipboard_clear()
            self.canvas.clipboard_append(text)
            browser_logger.error(f"Error copying with formatting: {e}")
    
    def _copy_selected(self):
        """Copy all selected studies to clipboard with formatting."""
        if not self.selected_indices:
            return

        try:
            import win32clipboard

            # Get formatting settings
            font_name = self.font_var.get()
            font_size = self.font_size_var.get()
            sponsor_color = self.sponsor_color_var.get()
            protocol_color = self.protocol_color_var.get()
            sponsor_bold = self.sponsor_bold_var.get()
            indent_inches = self.indent_var.get()
            show_protocol = self.show_protocol_var.get()

            # Build HTML for all studies
            html_body = '<html><head><meta charset="utf-8"></head><body>'
            text_lines = []

            for i in sorted(self.selected_indices):
                if i >= len(self.filtered_studies):
                    continue
                study = self.filtered_studies[i]

                year = str(study['year'])
                sponsor = study['sponsor']
                protocol = study['protocol']
                indent_px = int(indent_inches * 96)

                # Paragraph with hanging indent
                html_body += f'<p style="font-family:{font_name};font-size:{font_size}pt;'
                if indent_px > 0:
                    html_body += f'margin-left:{indent_px}px;text-indent:-{indent_px}px;'
                html_body += 'margin-top:0;margin-bottom:0;">'

                # Year followed by tab (use &#9; for HTML tab that Word recognizes)
                html_body += f'{year}'

                # Tab character as HTML entity - Word recognizes &#9; as a tab
                text_line = year
                if indent_inches > 0:
                    html_body += '&#9;'  # HTML tab entity
                    text_line += '\t'
                else:
                    html_body += ' '
                    text_line += ' '

                # Sponsor with color and bold
                sponsor_style = f'color:{sponsor_color};'
                if sponsor_bold:
                    sponsor_style += 'font-weight:bold;'
                html_body += f'<span style="{sponsor_style}">{sponsor}</span> '
                text_line += f"{sponsor} "

                # Protocol
                if show_protocol and protocol:
                    html_body += f'<span style="color:{protocol_color};">{protocol}</span>: '
                    desc = study['description_full']
                    text_line += f"{protocol}: "
                else:
                    html_body += ': '
                    desc = study['description_masked']
                    text_line += ": "

                # Description
                html_body += desc
                html_body += '</p>'
                text_line += desc
                text_lines.append(text_line)

            html_body += '</body></html>'

            # Create CF_HTML format with proper header
            html_prefix = "Version:0.9\r\nStartHTML:00000000\r\nEndHTML:00000000\r\nStartFragment:00000000\r\nEndFragment:00000000\r\n"
            html_with_markers = html_prefix + html_body

            # Calculate byte positions
            start_html = len(html_prefix)
            end_html = len(html_with_markers)
            start_fragment = html_with_markers.find('<body>') + 6
            end_fragment = html_with_markers.find('</body>')

            # Format header with positions
            cf_html_data = f"Version:0.9\r\nStartHTML:{start_html:08d}\r\nEndHTML:{end_html:08d}\r\nStartFragment:{start_fragment:08d}\r\nEndFragment:{end_fragment:08d}\r\n{html_body}"

            # Plain text fallback
            text = '\n'.join(text_lines)

            # Copy to clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()

            # Set plain text
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)

            # Set HTML format
            cf_html = win32clipboard.RegisterClipboardFormat('HTML Format')
            win32clipboard.SetClipboardData(cf_html, cf_html_data.encode('utf-8'))

            win32clipboard.CloseClipboard()
            browser_logger.info(f"Copied {len(self.selected_indices)} studies with formatting")

        except ImportError:
            # Fallback if win32clipboard not available
            lines = []
            for i in sorted(self.selected_indices):
                if i < len(self.filtered_studies):
                    lines.append(self._format_study(self.filtered_studies[i]))
            text = '\n'.join(lines)
            self.canvas.clipboard_clear()
            self.canvas.clipboard_append(text)
            browser_logger.warning("win32clipboard not available, copied as plain text")
        except Exception as e:
            # Fallback on any error
            lines = []
            for i in sorted(self.selected_indices):
                if i < len(self.filtered_studies):
                    lines.append(self._format_study(self.filtered_studies[i]))
            text = '\n'.join(lines)
            self.canvas.clipboard_clear()
            self.canvas.clipboard_append(text)
            browser_logger.error(f"Error copying with formatting: {e}")
    
    def _show_filter_dialog(self):
        """Show filter dialog."""
        dialog = FilterDialog(self.canvas, self.all_studies, 
                             self.filter_years, self.filter_phases, self.filter_subcategories)
        if dialog.result:
            self.filter_years = dialog.result['years']
            self.filter_phases = dialog.result['phases']
            self.filter_subcategories = dialog.result['subcategories']
            
            # Save to preferences
            self.prefs.mode_d_filter_years = list(self.filter_years)
            self.prefs.mode_d_filter_phases = list(self.filter_phases)
            self.prefs.mode_d_filter_subcategories = list(self.filter_subcategories)
            self.prefs_manager.save(self.prefs)
            
            self._apply_filters()
    
    def _choose_color(self, color_var):
        """Choose a color."""
        color = colorchooser.askcolor(initialcolor=color_var.get())
        if color[1]:
            color_var.set(color[1])
            self._mark_format_dirty()

            # Update preview
            if color_var == self.sponsor_color_var:
                self.sponsor_color_preview.config(bg=color[1])
            else:
                self.protocol_color_preview.config(bg=color[1])
    
    def _save_format_prefs(self):
        """Save formatting preferences."""
        self.prefs.mode_d_hanging_indent = self.indent_var.get()
        self.prefs.mode_d_font_name = self.font_var.get()
        self.prefs.mode_d_font_size = self.font_size_var.get()
        self.prefs.mode_d_sponsor_bold = self.sponsor_bold_var.get()
        self.prefs.mode_d_sponsor_color = self.sponsor_color_var.get()
        self.prefs.mode_d_protocol_color = self.protocol_color_var.get()
        self.prefs_manager.save(self.prefs)
        self._format_pending = False
        self._refresh_display()

        # Show success message
        from tkinter import messagebox
        messagebox.showinfo("Settings Saved", "Formatting settings have been saved successfully!")
        browser_logger.info("Mode D formatting preferences saved")


class FilterDialog:
    """Professional filter dialog for Mode D."""
    
    def __init__(self, parent, all_studies: List[Dict], 
                 current_years: Set[int], current_phases: Set[str], current_subcats: Set[str]):
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Filter Studies")
        self.dialog.geometry("600x700")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(False, False)
        
        # Center dialog on parent
        self.dialog.update_idletasks()
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        dialog_w = 600
        dialog_h = 700
        x = parent_x + (parent_w // 2) - (dialog_w // 2)
        y = parent_y + (parent_h // 2) - (dialog_h // 2)
        self.dialog.geometry(f"{dialog_w}x{dialog_h}+{x}+{y}")
        
        # Extract unique values
        years = sorted(set(s['year'] for s in all_studies), reverse=True)
        phases = sorted(set(s['phase'] for s in all_studies))
        subcats = sorted(set(s['subcategory'] for s in all_studies))
        
        # Main container with padding
        main_frame = ttk.Frame(self.dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Filter Studies", font=('Segoe UI', 14, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # Create filter sections
        notebook = ttk.Notebook(self.dialog)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Years tab
        years_frame = ttk.Frame(notebook)
        notebook.add(years_frame, text="Years")
        self.year_vars = {}
        self._create_checkboxes(years_frame, years, current_years, self.year_vars)
        
        # Phases tab
        phases_frame = ttk.Frame(notebook)
        notebook.add(phases_frame, text="Phases")
        self.phase_vars = {}
        self._create_checkboxes(phases_frame, phases, current_phases, self.phase_vars)
        
        # Subcategories tab
        subcats_frame = ttk.Frame(notebook)
        notebook.add(subcats_frame, text="Subcategories")
        self.subcat_vars = {}
        self._create_checkboxes(subcats_frame, subcats, current_subcats, self.subcat_vars)
        
        # Buttons
        btn_frame = ttk.Frame(self.dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="Reset", command=self._reset).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Apply", command=self._apply).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
        self.dialog.wait_window()
    
    def _create_checkboxes(self, parent, items, selected_items, var_dict):
        """Create checkboxes for filter items."""
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind('<Configure>', lambda e: canvas.config(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.config(yscrollcommand=scrollbar.set)
        
        for item in items:
            var = tk.BooleanVar(value=item in selected_items)
            var_dict[item] = var
            ttk.Checkbutton(scrollable_frame, text=str(item), variable=var).pack(anchor=tk.W, padx=5, pady=2)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def _reset(self):
        """Reset all filters."""
        for var in self.year_vars.values():
            var.set(False)
        for var in self.phase_vars.values():
            var.set(False)
        for var in self.subcat_vars.values():
            var.set(False)
    
    def _apply(self):
        """Apply filters."""
        self.result = {
            'years': {k for k, v in self.year_vars.items() if v.get()},
            'phases': {k for k, v in self.phase_vars.items() if v.get()},
            'subcategories': {k for k, v in self.subcat_vars.items() if v.get()}
        }
        self.dialog.destroy()
