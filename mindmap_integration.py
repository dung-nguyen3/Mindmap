#!/usr/bin/env python3
"""
Mindmap Integration Module - Integrates mindmap functionality into Excel Master Chart App
Provides the Mindmap View tab with split-pane interface
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, List, Optional, Callable
import json

# Import mindmap canvas
from mindmap_canvas import (
    MindmapCanvas, MindmapNode, NodeStyle, NodeShape,
    LineStyle, LayoutType, BranchDirection,
    COLOR_SETS, get_color_set
)

# Try to import export libraries
try:
    from PIL import Image, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import svgwrite
    SVG_AVAILABLE = True
except ImportError:
    SVG_AVAILABLE = False


# ============================================================================
# COLUMN MAPPING DIALOG
# ============================================================================

class ColumnMappingDialog(tk.Toplevel):
    """Dialog for mapping Excel columns to mindmap hierarchy"""

    def __init__(self, parent, columns: List[str], callback: Callable):
        super().__init__(parent)

        self.title("Column Mapping for Mindmap")
        self.geometry("500x450")
        self.transient(parent)
        self.grab_set()

        self.columns = columns
        self.callback = callback
        self.result = None

        self._create_widgets()
        self._center_window()

    def _create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Instructions
        ttk.Label(main_frame, text="Map Excel columns to mindmap hierarchy levels:",
                 font=("Calibri", 11, "bold")).pack(anchor=tk.W, pady=(0, 15))

        # Mapping frame
        mapping_frame = ttk.LabelFrame(main_frame, text="Column Mapping", padding="10")
        mapping_frame.pack(fill=tk.X, pady=(0, 15))

        # Level mappings
        self.level_vars = []
        levels = [
            ("Root Nodes (Level 1):", "Groups items by this column"),
            ("Child Nodes (Level 2):", "Individual items under each group"),
            ("Details (Level 3):", "First detail level (optional)"),
            ("Details (Level 4):", "Second detail level (optional)"),
        ]

        column_options = ["(None)"] + self.columns

        for i, (label, hint) in enumerate(levels):
            frame = ttk.Frame(mapping_frame)
            frame.pack(fill=tk.X, pady=5)

            ttk.Label(frame, text=label, width=22).pack(side=tk.LEFT)

            var = tk.StringVar()
            combo = ttk.Combobox(frame, textvariable=var, values=column_options,
                                state="readonly", width=25)
            combo.pack(side=tk.LEFT, padx=(0, 10))

            # Set defaults
            if i < len(self.columns):
                var.set(self.columns[i])
            else:
                var.set("(None)")

            self.level_vars.append(var)

            ttk.Label(frame, text=hint, foreground="gray").pack(side=tk.LEFT)

        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 15))

        self.include_all_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Include all remaining columns as node details",
                       variable=self.include_all_var).pack(anchor=tk.W)

        self.color_by_group_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Auto-color nodes by group (Level 1)",
                       variable=self.color_by_group_var).pack(anchor=tk.W)

        # Preview frame
        preview_frame = ttk.LabelFrame(main_frame, text="Hierarchy Preview", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        self.preview_text = tk.Text(preview_frame, height=6, width=50, state='disabled',
                                   font=("Courier", 10))
        self.preview_text.pack(fill=tk.BOTH, expand=True)

        # Update preview when selection changes
        for var in self.level_vars:
            var.trace_add('write', lambda *args: self._update_preview())

        self._update_preview()

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(button_frame, text="Apply Mapping", command=self._apply).pack(side=tk.RIGHT)

    def _update_preview(self):
        """Update the hierarchy preview"""
        self.preview_text.config(state='normal')
        self.preview_text.delete('1.0', tk.END)

        lines = []
        indent = 0
        for var in self.level_vars:
            col = var.get()
            if col and col != "(None)":
                prefix = "  " * indent + ("└─ " if indent > 0 else "")
                lines.append(f"{prefix}{col}")
                indent += 1

        if self.include_all_var.get():
            # Add remaining columns
            selected = {var.get() for var in self.level_vars if var.get() != "(None)"}
            remaining = [c for c in self.columns if c not in selected]
            if remaining:
                prefix = "  " * indent + "└─ "
                lines.append(f"{prefix}({', '.join(remaining[:3])}{'...' if len(remaining) > 3 else ''})")

        self.preview_text.insert('1.0', "\n".join(lines))
        self.preview_text.config(state='disabled')

    def _apply(self):
        """Apply the mapping"""
        mapping = []
        for var in self.level_vars:
            col = var.get()
            if col and col != "(None)":
                mapping.append(col)

        if not mapping:
            messagebox.showwarning("No Mapping", "Please select at least one column for mapping.")
            return

        self.result = {
            'levels': mapping,
            'include_all': self.include_all_var.get(),
            'color_by_group': self.color_by_group_var.get()
        }

        self.callback(self.result)
        self.destroy()

    def _center_window(self):
        """Center the dialog on screen"""
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")


# ============================================================================
# STYLE EDITOR DIALOG
# ============================================================================

class StyleEditorDialog(tk.Toplevel):
    """Dialog for editing node style"""

    def __init__(self, parent, node: MindmapNode, callback: Callable):
        super().__init__(parent)

        self.title("Edit Node Style")
        self.geometry("400x500")
        self.transient(parent)
        self.grab_set()

        self.node = node
        self.callback = callback
        self.style = NodeStyle(
            shape=node.style.shape,
            fill_color=node.style.fill_color,
            border_color=node.style.border_color,
            border_width=node.style.border_width,
            text_color=node.style.text_color,
            font_family=node.style.font_family,
            font_size=node.style.font_size,
            font_bold=node.style.font_bold,
            font_italic=node.style.font_italic,
            line_color=node.style.line_color,
            line_width=node.style.line_width,
            line_style=node.style.line_style,
        )

        self._create_widgets()
        self._center_window()

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Shape selection
        shape_frame = ttk.LabelFrame(main_frame, text="Shape", padding="10")
        shape_frame.pack(fill=tk.X, pady=(0, 10))

        self.shape_var = tk.StringVar(value=self.style.shape.value)
        shapes = [("Rectangle", "rectangle"), ("Rounded", "rounded_rectangle"),
                 ("Ellipse", "ellipse"), ("Pill", "pill"),
                 ("Diamond", "diamond"), ("Hexagon", "hexagon")]

        shape_row = ttk.Frame(shape_frame)
        shape_row.pack(fill=tk.X)
        for text, value in shapes:
            ttk.Radiobutton(shape_row, text=text, variable=self.shape_var,
                          value=value).pack(side=tk.LEFT, padx=5)

        # Colors
        colors_frame = ttk.LabelFrame(main_frame, text="Colors", padding="10")
        colors_frame.pack(fill=tk.X, pady=(0, 10))

        # Fill color
        fill_row = ttk.Frame(colors_frame)
        fill_row.pack(fill=tk.X, pady=2)
        ttk.Label(fill_row, text="Fill:", width=10).pack(side=tk.LEFT)
        self.fill_btn = tk.Button(fill_row, width=10, bg=self._ensure_hash(self.style.fill_color),
                                 command=lambda: self._pick_color('fill'))
        self.fill_btn.pack(side=tk.LEFT)

        # Border color
        border_row = ttk.Frame(colors_frame)
        border_row.pack(fill=tk.X, pady=2)
        ttk.Label(border_row, text="Border:", width=10).pack(side=tk.LEFT)
        self.border_btn = tk.Button(border_row, width=10, bg=self._ensure_hash(self.style.border_color),
                                   command=lambda: self._pick_color('border'))
        self.border_btn.pack(side=tk.LEFT)

        # Text color
        text_row = ttk.Frame(colors_frame)
        text_row.pack(fill=tk.X, pady=2)
        ttk.Label(text_row, text="Text:", width=10).pack(side=tk.LEFT)
        self.text_btn = tk.Button(text_row, width=10, bg=self._ensure_hash(self.style.text_color),
                                 command=lambda: self._pick_color('text'))
        self.text_btn.pack(side=tk.LEFT)

        # Line color
        line_row = ttk.Frame(colors_frame)
        line_row.pack(fill=tk.X, pady=2)
        ttk.Label(line_row, text="Line:", width=10).pack(side=tk.LEFT)
        self.line_btn = tk.Button(line_row, width=10, bg=self._ensure_hash(self.style.line_color),
                                 command=lambda: self._pick_color('line'))
        self.line_btn.pack(side=tk.LEFT)

        # Preset colors
        preset_frame = ttk.LabelFrame(main_frame, text="Preset Colors", padding="10")
        preset_frame.pack(fill=tk.X, pady=(0, 10))

        preset_row = ttk.Frame(preset_frame)
        preset_row.pack(fill=tk.X)
        for i, color_set in enumerate(COLOR_SETS):
            btn = tk.Button(preset_row, width=2, bg=f"#{color_set['main']}",
                           command=lambda cs=color_set: self._apply_preset(cs))
            btn.pack(side=tk.LEFT, padx=2)

        # Font
        font_frame = ttk.LabelFrame(main_frame, text="Font", padding="10")
        font_frame.pack(fill=tk.X, pady=(0, 10))

        font_row = ttk.Frame(font_frame)
        font_row.pack(fill=tk.X)

        ttk.Label(font_row, text="Size:").pack(side=tk.LEFT)
        self.font_size_var = tk.StringVar(value=str(self.style.font_size))
        ttk.Spinbox(font_row, from_=8, to=48, width=5, textvariable=self.font_size_var).pack(side=tk.LEFT, padx=5)

        self.bold_var = tk.BooleanVar(value=self.style.font_bold)
        ttk.Checkbutton(font_row, text="Bold", variable=self.bold_var).pack(side=tk.LEFT, padx=10)

        self.italic_var = tk.BooleanVar(value=self.style.font_italic)
        ttk.Checkbutton(font_row, text="Italic", variable=self.italic_var).pack(side=tk.LEFT)

        # Line style
        line_frame = ttk.LabelFrame(main_frame, text="Connection Line", padding="10")
        line_frame.pack(fill=tk.X, pady=(0, 10))

        self.line_style_var = tk.StringVar(value=self.style.line_style.value)
        line_styles = [("Straight", "straight"), ("Curved", "curved"),
                      ("Orthogonal", "orthogonal"), ("Tapered", "tapered")]

        line_style_row = ttk.Frame(line_frame)
        line_style_row.pack(fill=tk.X)
        for text, value in line_styles:
            ttk.Radiobutton(line_style_row, text=text, variable=self.line_style_var,
                          value=value).pack(side=tk.LEFT, padx=5)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(button_frame, text="Apply", command=self._apply).pack(side=tk.RIGHT)

    def _ensure_hash(self, color: str) -> str:
        """Ensure color has # prefix"""
        return color if color.startswith('#') else f'#{color}'

    def _pick_color(self, color_type: str):
        """Open color picker"""
        from tkinter import colorchooser

        current = getattr(self.style, f'{color_type}_color')
        color = colorchooser.askcolor(color=self._ensure_hash(current), title=f"Choose {color_type} color")

        if color[1]:
            setattr(self.style, f'{color_type}_color', color[1])
            btn = getattr(self, f'{color_type}_btn')
            btn.config(bg=color[1])

    def _apply_preset(self, color_set: dict):
        """Apply a preset color set"""
        self.style.fill_color = f"#{color_set['main']}"
        self.style.border_color = f"#{color_set['header']}"

        self.fill_btn.config(bg=self.style.fill_color)
        self.border_btn.config(bg=self.style.border_color)

    def _apply(self):
        """Apply the style changes"""
        self.style.shape = NodeShape(self.shape_var.get())
        self.style.font_size = int(self.font_size_var.get())
        self.style.font_bold = self.bold_var.get()
        self.style.font_italic = self.italic_var.get()
        self.style.line_style = LineStyle(self.line_style_var.get())

        self.callback(self.style)
        self.destroy()

    def _center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")


# ============================================================================
# MINDMAP VIEW PANEL
# ============================================================================

class MindmapViewPanel(ttk.Frame):
    """The Mindmap View panel with split-pane interface"""

    def __init__(self, parent, get_excel_data: Callable, get_columns: Callable,
                 update_excel_data: Callable):
        super().__init__(parent)

        self.get_excel_data = get_excel_data
        self.get_columns = get_columns
        self.update_excel_data = update_excel_data

        self.column_mapping = None
        self.syncing = False  # Prevent recursive sync

        self._create_widgets()

    def _create_widgets(self):
        # Main container
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # Toolbar
        self._create_toolbar()

        # Split pane
        self.paned = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashwidth=8,
                                    sashrelief=tk.RAISED, opaqueresize=False)
        self.paned.grid(row=1, column=0, sticky='nsew', padx=5, pady=5)

        # Left panel: Excel data view
        self._create_excel_panel()

        # Right panel: Mindmap canvas
        self._create_mindmap_panel()

        # Status bar
        self._create_status_bar()

    def _create_toolbar(self):
        """Create the mindmap toolbar"""
        toolbar = ttk.Frame(self)
        toolbar.grid(row=0, column=0, sticky='ew', padx=5, pady=5)

        # Layout dropdown
        ttk.Label(toolbar, text="Layout:").pack(side=tk.LEFT, padx=(0, 5))
        self.layout_var = tk.StringVar(value="tree_right")
        layout_combo = ttk.Combobox(toolbar, textvariable=self.layout_var, state="readonly", width=18)
        layout_combo['values'] = [
            "Tree Right", "Tree Left", "Tree Down", "Tree Up",
            "Vertical Balanced", "Horizontal Balanced",
            "Vertical Custom", "Horizontal Custom",
            "Radial", "Free Form"
        ]
        layout_combo.current(0)
        layout_combo.pack(side=tk.LEFT, padx=(0, 15))
        layout_combo.bind('<<ComboboxSelected>>', self._on_layout_change)

        # Line style dropdown
        ttk.Label(toolbar, text="Lines:").pack(side=tk.LEFT, padx=(0, 5))
        self.line_var = tk.StringVar(value="curved")
        line_combo = ttk.Combobox(toolbar, textvariable=self.line_var, state="readonly", width=12)
        line_combo['values'] = ["Straight", "Curved", "Orthogonal", "Tapered"]
        line_combo.current(1)
        line_combo.pack(side=tk.LEFT, padx=(0, 15))
        line_combo.bind('<<ComboboxSelected>>', self._on_line_change)

        # Shape dropdown
        ttk.Label(toolbar, text="Shape:").pack(side=tk.LEFT, padx=(0, 5))
        self.shape_var = tk.StringVar(value="rounded_rectangle")
        shape_combo = ttk.Combobox(toolbar, textvariable=self.shape_var, state="readonly", width=14)
        shape_combo['values'] = ["Rectangle", "Rounded", "Ellipse", "Pill", "Diamond", "Hexagon"]
        shape_combo.current(1)
        shape_combo.pack(side=tk.LEFT, padx=(0, 15))
        shape_combo.bind('<<ComboboxSelected>>', self._on_shape_change)

        # Separator
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        # View buttons
        ttk.Button(toolbar, text="Fit All", command=self._fit_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Center", command=self._center_root).pack(side=tk.LEFT, padx=2)

        # Zoom controls
        ttk.Label(toolbar, text="Zoom:").pack(side=tk.LEFT, padx=(15, 5))
        ttk.Button(toolbar, text="-", width=3, command=self._zoom_out).pack(side=tk.LEFT)
        self.zoom_label = ttk.Label(toolbar, text="100%", width=6)
        self.zoom_label.pack(side=tk.LEFT)
        ttk.Button(toolbar, text="+", width=3, command=self._zoom_in).pack(side=tk.LEFT)

        # Separator
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        # Mapping button
        ttk.Button(toolbar, text="Column Mapping...", command=self._show_mapping_dialog).pack(side=tk.LEFT, padx=2)

        # Refresh button
        ttk.Button(toolbar, text="Refresh", command=self._refresh_mindmap).pack(side=tk.LEFT, padx=2)

        # Export dropdown
        export_menu = tk.Menu(self, tearoff=0)
        export_menu.add_command(label="Export as PNG...", command=self._export_png)
        export_menu.add_command(label="Export as SVG...", command=self._export_svg)
        export_menu.add_command(label="Export as Mermaid...", command=self._export_mermaid)

        export_btn = ttk.Menubutton(toolbar, text="Export")
        export_btn['menu'] = export_menu
        export_btn.pack(side=tk.RIGHT, padx=2)

    def _create_excel_panel(self):
        """Create the left panel with Excel data"""
        left_frame = ttk.Frame(self.paned)

        # Header
        header = ttk.Frame(left_frame)
        header.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header, text="Excel Data", font=("Calibri", 11, "bold")).pack(side=tk.LEFT)

        # Treeview for data display (editable)
        tree_frame = ttk.Frame(left_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        self.data_tree = ttk.Treeview(tree_frame, yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.config(command=self.data_tree.yview)
        hsb.config(command=self.data_tree.xview)

        self.data_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        # Bind double-click for editing
        self.data_tree.bind('<Double-1>', self._on_tree_double_click)

        self.paned.add(left_frame, minsize=200, width=350)

    def _create_mindmap_panel(self):
        """Create the right panel with mindmap canvas"""
        right_frame = ttk.Frame(self.paned)

        # Header
        header = ttk.Frame(right_frame)
        header.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header, text="Mindmap", font=("Calibri", 11, "bold")).pack(side=tk.LEFT)

        # Canvas with scrollbars
        canvas_frame = ttk.Frame(right_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(canvas_frame, orient="vertical")
        hsb = ttk.Scrollbar(canvas_frame, orient="horizontal")

        self.mindmap = MindmapCanvas(canvas_frame, yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.config(command=self.mindmap.yview)
        hsb.config(command=self.mindmap.xview)

        self.mindmap.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        canvas_frame.columnconfigure(0, weight=1)
        canvas_frame.rowconfigure(0, weight=1)

        # Set callbacks
        self.mindmap.on_node_selected = self._on_node_selected
        self.mindmap.on_node_edited = self._on_node_edited
        self.mindmap.on_structure_changed = self._on_structure_changed
        self.mindmap.on_style_edit = self._on_style_edit  # Connect style editor

        self.paned.add(right_frame, minsize=300)

    def _on_style_edit(self, node_id: str):
        """Show style editor dialog for a node"""
        node = self.mindmap.get_node(node_id)
        if not node:
            return

        def apply_style(new_style):
            self.mindmap.update_node_style(node_id, new_style)
            self.mindmap.redraw()

        StyleEditorDialog(self, node, apply_style)

    def _create_status_bar(self):
        """Create status bar"""
        status_frame = ttk.Frame(self)
        status_frame.grid(row=2, column=0, sticky='ew', padx=5, pady=(0, 5))

        self.node_count_label = ttk.Label(status_frame, text="Nodes: 0")
        self.node_count_label.pack(side=tk.LEFT, padx=(0, 20))

        self.sync_status_label = ttk.Label(status_frame, text="Synced", foreground="green")
        self.sync_status_label.pack(side=tk.LEFT)

        ttk.Label(status_frame, text="|").pack(side=tk.LEFT, padx=10)

        self.layout_label = ttk.Label(status_frame, text="Layout: Tree Right")
        self.layout_label.pack(side=tk.LEFT)

    # ========================================================================
    # DATA CONVERSION
    # ========================================================================

    def load_from_excel_data(self, mapping: dict = None):
        """Convert Excel data to mindmap based on column mapping"""
        if mapping:
            self.column_mapping = mapping

        if not self.column_mapping:
            return

        columns = self.get_columns()
        data = self.get_excel_data()

        if not data or not columns:
            return

        # Clear existing mindmap
        self.mindmap.clear()

        # Get level columns
        level_cols = self.column_mapping.get('levels', [])
        if not level_cols:
            return

        # Get column indices
        level_indices = []
        for col in level_cols:
            if col in columns:
                level_indices.append(columns.index(col))
            else:
                level_indices.append(-1)

        # Build tree structure
        # Level 0: Root (chart name or first level column name)
        root_col = level_cols[0] if level_cols else "Mindmap"
        root_id = self.mindmap.add_node(root_col)

        # Assign root style
        root_style = NodeStyle(
            shape=NodeShape.ELLIPSE,
            fill_color="#4472C4",
            border_color="#2E5090",
            text_color="#FFFFFF",
            font_size=14,
            font_bold=True
        )
        self.mindmap.nodes[root_id].style = root_style

        # Track created nodes for hierarchy
        node_cache = {}  # (level, value) -> node_id
        color_index = 0

        # Process each row
        for row_idx, row in enumerate(data):
            if not row or not any(cell.strip() if isinstance(cell, str) else cell for cell in row):
                continue  # Skip empty rows

            parent_id = root_id
            current_path = []

            for level, col_idx in enumerate(level_indices):
                if col_idx < 0 or col_idx >= len(row):
                    continue

                cell_value = row[col_idx]
                if isinstance(cell_value, str):
                    cell_value = cell_value.strip()
                if not cell_value:
                    continue

                current_path.append(str(cell_value))
                cache_key = tuple(current_path)

                if cache_key in node_cache:
                    parent_id = node_cache[cache_key]
                else:
                    # Create new node
                    # Assign color based on level 1 groups
                    if level == 0:
                        color_set = get_color_set(color_index)
                        color_index += 1
                        style = NodeStyle(
                            fill_color=f"#{color_set['header']}",
                            border_color=f"#{color_set['header']}",
                            font_bold=True,
                            font_size=12
                        )
                    elif level == 1:
                        # Get parent's color
                        parent_node = self.mindmap.get_node(parent_id)
                        if parent_node:
                            parent_fill = parent_node.style.fill_color.lstrip('#')
                            # Find matching color set
                            for cs in COLOR_SETS:
                                if cs['header'] == parent_fill:
                                    style = NodeStyle(
                                        fill_color=f"#{cs['main']}",
                                        border_color=f"#{cs['header']}"
                                    )
                                    break
                            else:
                                style = NodeStyle()
                        else:
                            style = NodeStyle()
                    else:
                        style = NodeStyle(font_size=10)

                    # Store extra data (other columns)
                    extra_data = {}
                    if self.column_mapping.get('include_all', True):
                        for i, col in enumerate(columns):
                            if col not in level_cols and i < len(row):
                                val = row[i]
                                if val and (isinstance(val, str) and val.strip()):
                                    extra_data[col] = str(val).strip()

                    node_id = self.mindmap.add_node(str(cell_value), parent_id=parent_id,
                                                   style=style, extra_data=extra_data)
                    node_cache[cache_key] = node_id
                    parent_id = node_id

        # Update tree view
        self._update_tree_view()

        # Redraw mindmap
        self.mindmap.redraw()
        self.mindmap.fit_all()

        # Update status
        self._update_status()

    def _update_tree_view(self):
        """Update the treeview with current Excel data"""
        # Clear existing
        for item in self.data_tree.get_children():
            self.data_tree.delete(item)

        columns = self.get_columns()
        data = self.get_excel_data()

        if not columns:
            return

        # Configure columns
        self.data_tree['columns'] = columns
        self.data_tree['show'] = 'headings'

        for col in columns:
            self.data_tree.heading(col, text=col)
            self.data_tree.column(col, width=100, minwidth=50)

        # Add data
        for row in data:
            if row and any(cell for cell in row):
                values = [str(cell) if cell else "" for cell in row]
                self.data_tree.insert('', 'end', values=values)

    # ========================================================================
    # EVENT HANDLERS
    # ========================================================================

    def _on_layout_change(self, event=None):
        """Handle layout dropdown change"""
        layout_map = {
            "Tree Right": LayoutType.TREE_RIGHT,
            "Tree Left": LayoutType.TREE_LEFT,
            "Tree Down": LayoutType.TREE_DOWN,
            "Tree Up": LayoutType.TREE_UP,
            "Vertical Balanced": LayoutType.VERTICAL_BALANCED,
            "Horizontal Balanced": LayoutType.HORIZONTAL_BALANCED,
            "Vertical Custom": LayoutType.VERTICAL_CUSTOM,
            "Horizontal Custom": LayoutType.HORIZONTAL_CUSTOM,
            "Radial": LayoutType.RADIAL,
            "Free Form": LayoutType.FREE_FORM,
        }

        selected = self.layout_var.get()
        if selected in layout_map:
            self.mindmap.set_layout(layout_map[selected])
            self.layout_label.config(text=f"Layout: {selected}")

    def _on_line_change(self, event=None):
        """Handle line style dropdown change"""
        line_map = {
            "Straight": LineStyle.STRAIGHT,
            "Curved": LineStyle.CURVED,
            "Orthogonal": LineStyle.ORTHOGONAL,
            "Tapered": LineStyle.TAPERED,
        }

        selected = self.line_var.get()
        if selected in line_map:
            self.mindmap.set_line_style(line_map[selected])

    def _on_shape_change(self, event=None):
        """Handle shape dropdown change"""
        shape_map = {
            "Rectangle": NodeShape.RECTANGLE,
            "Rounded": NodeShape.ROUNDED_RECTANGLE,
            "Ellipse": NodeShape.ELLIPSE,
            "Pill": NodeShape.PILL,
            "Diamond": NodeShape.DIAMOND,
            "Hexagon": NodeShape.HEXAGON,
        }

        selected = self.shape_var.get()
        if selected in shape_map:
            self.mindmap.set_default_shape(shape_map[selected])

    def _on_node_selected(self, node_id: str):
        """Handle node selection in mindmap - highlight corresponding row in tree view"""
        if not node_id or node_id not in self.mindmap.nodes:
            return

        node = self.mindmap.nodes[node_id]
        node_text = node.text

        # Find and select matching row in tree view
        for item in self.data_tree.get_children():
            values = self.data_tree.item(item, 'values')
            # Check if any column contains the node text
            if node_text in values:
                self.data_tree.selection_set(item)
                self.data_tree.focus(item)
                self.data_tree.see(item)  # Scroll to make visible
                break

    def _on_node_edited(self, node_id: str, new_text: str):
        """Handle node text edit in mindmap - sync back to Excel data"""
        if self.syncing:
            return

        self.syncing = True

        node = self.mindmap.nodes.get(node_id)
        if node:
            old_text = node.text if hasattr(node, '_old_text') else ""

            # Find and update matching row in tree view
            for item in self.data_tree.get_children():
                values = list(self.data_tree.item(item, 'values'))
                # Find column with old text and update it
                for i, val in enumerate(values):
                    if val == old_text or val == new_text:
                        values[i] = new_text
                        self.data_tree.item(item, values=values)

                        # Also update the main Excel sheet via callback
                        row_idx = self.data_tree.index(item)
                        self._sync_row_to_excel(row_idx, values)
                        break

        self.sync_status_label.config(text="Modified", foreground="orange")
        self.syncing = False

    def _sync_row_to_excel(self, row_idx: int, values: list):
        """Sync a row from tree view back to main Excel data"""
        try:
            # Get current Excel data
            data = self.get_excel_data()
            if row_idx < len(data):
                # Update the row
                for i, val in enumerate(values):
                    if i < len(data[row_idx]):
                        data[row_idx][i] = val

                # Push back to Excel
                self.update_excel_data(data)
        except Exception as e:
            print(f"Error syncing to Excel: {e}")

    def _on_structure_changed(self):
        """Handle structure change in mindmap"""
        self._update_status()
        self.sync_status_label.config(text="Modified", foreground="orange")

    def _on_tree_double_click(self, event):
        """Handle double-click in tree view for editing"""
        item = self.data_tree.selection()
        if not item:
            return

        # Get column clicked
        column = self.data_tree.identify_column(event.x)
        col_idx = int(column.replace('#', '')) - 1

        if col_idx < 0:
            return

        # Create entry for editing
        columns = self.get_columns()
        if col_idx >= len(columns):
            return

        # Get current value
        values = self.data_tree.item(item[0], 'values')
        if col_idx >= len(values):
            return

        current_value = values[col_idx]

        # Create popup entry
        x, y, width, height = self.data_tree.bbox(item[0], column)

        entry = ttk.Entry(self.data_tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus_set()

        def save_edit(event=None):
            new_value = entry.get()
            new_values = list(values)
            new_values[col_idx] = new_value
            self.data_tree.item(item[0], values=new_values)
            entry.destroy()

            # Update Excel data - sync back to the main sheet
            row_idx = self.data_tree.index(item[0])
            self._sync_row_to_excel(row_idx, new_values)

            # Also refresh mindmap to reflect changes
            if self.column_mapping:
                self.load_from_excel_data()

            self.sync_status_label.config(text="Modified", foreground="orange")

        def cancel_edit(event=None):
            entry.destroy()

        entry.bind('<Return>', save_edit)
        entry.bind('<Escape>', cancel_edit)
        entry.bind('<FocusOut>', save_edit)

    # ========================================================================
    # VIEW CONTROLS
    # ========================================================================

    def _fit_all(self):
        self.mindmap.fit_all()
        self._update_zoom_label()

    def _center_root(self):
        self.mindmap.center_on_root()

    def _zoom_in(self):
        self.mindmap._zoom(1.2, 0, 0)
        self._update_zoom_label()

    def _zoom_out(self):
        self.mindmap._zoom(0.8, 0, 0)
        self._update_zoom_label()

    def _update_zoom_label(self):
        zoom_pct = int(self.mindmap.zoom_level * 100)
        self.zoom_label.config(text=f"{zoom_pct}%")

    def _update_status(self):
        node_count = len(self.mindmap.nodes)
        self.node_count_label.config(text=f"Nodes: {node_count}")

    # ========================================================================
    # DIALOGS
    # ========================================================================

    def _show_mapping_dialog(self):
        """Show column mapping dialog"""
        columns = self.get_columns()
        if not columns:
            messagebox.showwarning("No Data", "Please load data in the Excel View first.")
            return

        ColumnMappingDialog(self, columns, self._apply_mapping)

    def _apply_mapping(self, mapping: dict):
        """Apply column mapping and refresh mindmap"""
        self.column_mapping = mapping
        self._refresh_mindmap()

    def _refresh_mindmap(self):
        """Refresh the mindmap from Excel data"""
        if not self.column_mapping:
            self._show_mapping_dialog()
        else:
            self.load_from_excel_data()
            self.sync_status_label.config(text="Synced", foreground="green")

    # ========================================================================
    # EXPORT
    # ========================================================================

    def _export_png(self):
        """Export mindmap as PNG using Pillow"""
        if not PIL_AVAILABLE:
            messagebox.showerror("Export Error", "PNG export requires Pillow library.\nInstall with: pip install Pillow")
            return

        if not self.mindmap.nodes:
            messagebox.showwarning("Export Warning", "No mindmap to export. Please generate a mindmap first.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("All files", "*.*")],
            title="Export Mindmap as PNG"
        )

        if filepath:
            try:
                # Calculate bounds of all nodes
                min_x = min(n.x - n.width/2 for n in self.mindmap.nodes.values()) - 50
                max_x = max(n.x + n.width/2 for n in self.mindmap.nodes.values()) + 50
                min_y = min(n.y - n.height/2 for n in self.mindmap.nodes.values()) - 50
                max_y = max(n.y + n.height/2 for n in self.mindmap.nodes.values()) + 50

                width = int(max_x - min_x)
                height = int(max_y - min_y)

                # Create image with white background
                img = Image.new('RGB', (width, height), 'white')
                draw = ImageDraw.Draw(img)

                # Offset to center content
                offset_x = -min_x
                offset_y = -min_y

                # Draw connections first
                for node in self.mindmap.nodes.values():
                    if node.parent_id and node.parent_id in self.mindmap.nodes:
                        parent = self.mindmap.nodes[node.parent_id]
                        px, py = parent.x + offset_x, parent.y + offset_y
                        cx, cy = node.x + offset_x, node.y + offset_y
                        line_color = node.style.line_color.lstrip('#')
                        if len(line_color) == 6:
                            line_color = tuple(int(line_color[i:i+2], 16) for i in (0, 2, 4))
                        else:
                            line_color = (100, 100, 100)
                        draw.line([(px, py), (cx, cy)], fill=line_color, width=2)

                # Draw nodes
                for node in self.mindmap.nodes.values():
                    x, y = node.x + offset_x, node.y + offset_y
                    w, h = max(node.width, 80), max(node.height, 30)

                    # Parse colors
                    fill_color = node.style.fill_color.lstrip('#')
                    if len(fill_color) == 6:
                        fill_color = tuple(int(fill_color[i:i+2], 16) for i in (0, 2, 4))
                    else:
                        fill_color = (227, 242, 253)

                    border_color = node.style.border_color.lstrip('#')
                    if len(border_color) == 6:
                        border_color = tuple(int(border_color[i:i+2], 16) for i in (0, 2, 4))
                    else:
                        border_color = (179, 212, 237)

                    # Draw rounded rectangle (simplified as rectangle)
                    x1, y1 = x - w/2, y - h/2
                    x2, y2 = x + w/2, y + h/2
                    draw.rectangle([x1, y1, x2, y2], fill=fill_color, outline=border_color, width=2)

                    # Draw text
                    text_color = node.style.text_color.lstrip('#')
                    if len(text_color) == 6:
                        text_color = tuple(int(text_color[i:i+2], 16) for i in (0, 2, 4))
                    else:
                        text_color = (0, 0, 0)

                    # Center text in node
                    try:
                        font = ImageFont.truetype("arial.ttf", node.style.font_size)
                    except:
                        font = ImageFont.load_default()

                    bbox = draw.textbbox((0, 0), node.text, font=font)
                    text_w = bbox[2] - bbox[0]
                    text_h = bbox[3] - bbox[1]
                    draw.text((x - text_w/2, y - text_h/2), node.text, fill=text_color, font=font)

                img.save(filepath)
                messagebox.showinfo("Export", f"PNG exported to {filepath}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export PNG: {e}")

    def _export_svg(self):
        """Export mindmap as SVG with proper node sizes and connections"""
        if not SVG_AVAILABLE:
            messagebox.showerror("Export Error", "SVG export requires svgwrite library.\nInstall with: pip install svgwrite")
            return

        if not self.mindmap.nodes:
            messagebox.showwarning("Export Warning", "No mindmap to export. Please generate a mindmap first.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".svg",
            filetypes=[("SVG files", "*.svg"), ("All files", "*.*")],
            title="Export Mindmap as SVG"
        )

        if filepath:
            try:
                # Calculate bounds
                min_x = min(n.x - n.width/2 for n in self.mindmap.nodes.values()) - 50
                max_x = max(n.x + n.width/2 for n in self.mindmap.nodes.values()) + 50
                min_y = min(n.y - n.height/2 for n in self.mindmap.nodes.values()) - 50
                max_y = max(n.y + n.height/2 for n in self.mindmap.nodes.values()) + 50

                width = int(max_x - min_x)
                height = int(max_y - min_y)
                offset_x = -min_x
                offset_y = -min_y

                # Create SVG with proper size
                dwg = svgwrite.Drawing(filepath, size=(f'{width}px', f'{height}px'))
                dwg.add(dwg.rect(insert=(0, 0), size=(width, height), fill='white'))

                # Draw connections first
                for node in self.mindmap.nodes.values():
                    if node.parent_id and node.parent_id in self.mindmap.nodes:
                        parent = self.mindmap.nodes[node.parent_id]
                        px, py = parent.x + offset_x, parent.y + offset_y
                        cx, cy = node.x + offset_x, node.y + offset_y

                        line_color = node.style.line_color
                        if not line_color.startswith('#'):
                            line_color = f'#{line_color}'

                        # Draw curved line
                        mid_x = (px + cx) / 2
                        path_data = f'M {px},{py} Q {mid_x},{py} {mid_x},{(py+cy)/2} T {cx},{cy}'
                        dwg.add(dwg.path(d=path_data, stroke=line_color, fill='none',
                                        stroke_width=node.style.line_width))

                # Draw nodes
                for node in self.mindmap.nodes.values():
                    x, y = node.x + offset_x, node.y + offset_y
                    w, h = max(node.width, 80), max(node.height, 30)

                    fill_color = node.style.fill_color
                    if not fill_color.startswith('#'):
                        fill_color = f'#{fill_color}'

                    border_color = node.style.border_color
                    if not border_color.startswith('#'):
                        border_color = f'#{border_color}'

                    text_color = node.style.text_color
                    if not text_color.startswith('#'):
                        text_color = f'#{text_color}'

                    # Draw rounded rectangle
                    dwg.add(dwg.rect(
                        insert=(x - w/2, y - h/2),
                        size=(w, h),
                        rx=8, ry=8,  # Rounded corners
                        fill=fill_color,
                        stroke=border_color,
                        stroke_width=node.style.border_width
                    ))

                    # Draw text
                    dwg.add(dwg.text(
                        node.text,
                        insert=(x, y + 5),  # Slight offset for vertical centering
                        text_anchor='middle',
                        font_family=node.style.font_family,
                        font_size=f'{node.style.font_size}px',
                        fill=text_color
                    ))

                dwg.save()
                messagebox.showinfo("Export", f"SVG exported to {filepath}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export SVG: {e}")

    def _export_mermaid(self):
        """Export mindmap as Mermaid markdown"""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".md",
            filetypes=[("Markdown files", "*.md"), ("All files", "*.*")],
            title="Export Mindmap as Mermaid"
        )

        if filepath:
            try:
                mermaid_code = self.mindmap.to_mermaid()
                with open(filepath, 'w') as f:
                    f.write(mermaid_code)
                messagebox.showinfo("Export", f"Mermaid markdown exported to {filepath}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export Mermaid: {e}")
