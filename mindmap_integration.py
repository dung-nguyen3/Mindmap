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
        self.geometry("550x520")
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

        # Mode selector at top
        mode_frame = ttk.LabelFrame(main_frame, text="Mapping Mode", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 15))

        self.mode_var = tk.StringVar(value="by_columns")

        modes = [
            ("by_columns", "By Columns", "Each column = hierarchy level (Drug Class → Drug Name → Route)"),
            ("by_parent", "By Parent", "Two columns: Node name + Parent name (flexible branching)"),
            ("by_level", "By Level", "Two columns: Level number + Node name (outline style)")
        ]

        for mode_id, mode_name, mode_desc in modes:
            frame = ttk.Frame(mode_frame)
            frame.pack(fill=tk.X, pady=2)
            ttk.Radiobutton(frame, text=mode_name, variable=self.mode_var,
                           value=mode_id, command=self._on_mode_change).pack(side=tk.LEFT)
            ttk.Label(frame, text=f"- {mode_desc}", foreground="gray").pack(side=tk.LEFT, padx=(10, 0))

        # Dynamic mapping frame (changes based on mode)
        self.mapping_container = ttk.Frame(main_frame)
        self.mapping_container.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Create all mode frames
        self._create_by_columns_frame()
        self._create_by_parent_frame()
        self._create_by_level_frame()

        # Show initial mode
        self._on_mode_change()

        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 15))

        self.include_all_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Include all remaining columns as node details",
                       variable=self.include_all_var).pack(anchor=tk.W)

        self.color_by_group_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Auto-color nodes by group",
                       variable=self.color_by_group_var).pack(anchor=tk.W)

        # Preview frame
        preview_frame = ttk.LabelFrame(main_frame, text="Hierarchy Preview", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        self.preview_text = tk.Text(preview_frame, height=5, width=50, state='disabled',
                                   font=("Courier", 10))
        self.preview_text.pack(fill=tk.BOTH, expand=True)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(button_frame, text="Apply Mapping", command=self._apply).pack(side=tk.RIGHT)

    def _create_by_columns_frame(self):
        """Create the By Columns mapping UI"""
        self.by_columns_frame = ttk.LabelFrame(self.mapping_container, text="Column Hierarchy", padding="10")

        self.level_vars = []
        levels = [
            ("Root Nodes (Level 1):", "Groups items by this column"),
            ("Child Nodes (Level 2):", "Individual items under each group"),
            ("Details (Level 3):", "First detail level (optional)"),
            ("Details (Level 4):", "Second detail level (optional)"),
        ]

        column_options = ["(None)"] + self.columns

        for i, (label, hint) in enumerate(levels):
            frame = ttk.Frame(self.by_columns_frame)
            frame.pack(fill=tk.X, pady=3)

            ttk.Label(frame, text=label, width=20).pack(side=tk.LEFT)

            var = tk.StringVar()
            combo = ttk.Combobox(frame, textvariable=var, values=column_options,
                                state="readonly", width=20)
            combo.pack(side=tk.LEFT, padx=(0, 10))

            if i < len(self.columns):
                var.set(self.columns[i])
            else:
                var.set("(None)")

            self.level_vars.append(var)
            var.trace_add('write', lambda *args: self._update_preview())

            ttk.Label(frame, text=hint, foreground="gray", font=("Calibri", 9)).pack(side=tk.LEFT)

    def _create_by_parent_frame(self):
        """Create the By Parent mapping UI"""
        self.by_parent_frame = ttk.LabelFrame(self.mapping_container, text="Parent-Child Columns", padding="10")

        column_options = ["(None)"] + self.columns

        # Node name column
        frame1 = ttk.Frame(self.by_parent_frame)
        frame1.pack(fill=tk.X, pady=5)
        ttk.Label(frame1, text="Node Name Column:", width=20).pack(side=tk.LEFT)
        self.node_name_var = tk.StringVar(value=self.columns[0] if self.columns else "(None)")
        combo1 = ttk.Combobox(frame1, textvariable=self.node_name_var, values=column_options,
                             state="readonly", width=20)
        combo1.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(frame1, text="The name/label of each node", foreground="gray", font=("Calibri", 9)).pack(side=tk.LEFT)
        self.node_name_var.trace_add('write', lambda *args: self._update_preview())

        # Parent column
        frame2 = ttk.Frame(self.by_parent_frame)
        frame2.pack(fill=tk.X, pady=5)
        ttk.Label(frame2, text="Parent Column:", width=20).pack(side=tk.LEFT)
        self.parent_col_var = tk.StringVar(value=self.columns[1] if len(self.columns) > 1 else "(None)")
        combo2 = ttk.Combobox(frame2, textvariable=self.parent_col_var, values=column_options,
                             state="readonly", width=20)
        combo2.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(frame2, text="The parent node (blank = root)", foreground="gray", font=("Calibri", 9)).pack(side=tk.LEFT)
        self.parent_col_var.trace_add('write', lambda *args: self._update_preview())

        # Help text
        help_text = ttk.Label(self.by_parent_frame,
                             text="Example: Node='Apple', Parent='Fruits' → Fruits → Apple",
                             foreground="blue", font=("Calibri", 9))
        help_text.pack(anchor=tk.W, pady=(10, 0))

    def _create_by_level_frame(self):
        """Create the By Level mapping UI"""
        self.by_level_frame = ttk.LabelFrame(self.mapping_container, text="Level-Based Columns", padding="10")

        column_options = ["(None)"] + self.columns

        # Level column
        frame1 = ttk.Frame(self.by_level_frame)
        frame1.pack(fill=tk.X, pady=5)
        ttk.Label(frame1, text="Level Column:", width=20).pack(side=tk.LEFT)
        self.level_col_var = tk.StringVar(value=self.columns[0] if self.columns else "(None)")
        combo1 = ttk.Combobox(frame1, textvariable=self.level_col_var, values=column_options,
                             state="readonly", width=20)
        combo1.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(frame1, text="Hierarchy depth (1, 2, 3...)", foreground="gray", font=("Calibri", 9)).pack(side=tk.LEFT)
        self.level_col_var.trace_add('write', lambda *args: self._update_preview())

        # Node name column
        frame2 = ttk.Frame(self.by_level_frame)
        frame2.pack(fill=tk.X, pady=5)
        ttk.Label(frame2, text="Node Name Column:", width=20).pack(side=tk.LEFT)
        self.level_name_var = tk.StringVar(value=self.columns[1] if len(self.columns) > 1 else "(None)")
        combo2 = ttk.Combobox(frame2, textvariable=self.level_name_var, values=column_options,
                             state="readonly", width=20)
        combo2.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(frame2, text="The name/label of each node", foreground="gray", font=("Calibri", 9)).pack(side=tk.LEFT)
        self.level_name_var.trace_add('write', lambda *args: self._update_preview())

        # Help text
        help_text = ttk.Label(self.by_level_frame,
                             text="Example: Level=2, Name='Apple' → Child node under most recent Level 1",
                             foreground="blue", font=("Calibri", 9))
        help_text.pack(anchor=tk.W, pady=(10, 0))

    def _on_mode_change(self):
        """Handle mode selection change"""
        # Hide all frames
        self.by_columns_frame.pack_forget()
        self.by_parent_frame.pack_forget()
        self.by_level_frame.pack_forget()

        # Show selected frame
        mode = self.mode_var.get()
        if mode == "by_columns":
            self.by_columns_frame.pack(fill=tk.BOTH, expand=True)
        elif mode == "by_parent":
            self.by_parent_frame.pack(fill=tk.BOTH, expand=True)
        elif mode == "by_level":
            self.by_level_frame.pack(fill=tk.BOTH, expand=True)

        self._update_preview()

    def _update_preview(self):
        """Update the hierarchy preview"""
        self.preview_text.config(state='normal')
        self.preview_text.delete('1.0', tk.END)

        mode = self.mode_var.get()
        lines = []

        if mode == "by_columns":
            indent = 0
            for var in self.level_vars:
                col = var.get()
                if col and col != "(None)":
                    prefix = "  " * indent + ("└─ " if indent > 0 else "")
                    lines.append(f"{prefix}{col}")
                    indent += 1

            if self.include_all_var.get():
                selected = {var.get() for var in self.level_vars if var.get() != "(None)"}
                remaining = [c for c in self.columns if c not in selected]
                if remaining:
                    prefix = "  " * indent + "└─ "
                    lines.append(f"{prefix}({', '.join(remaining[:3])}{'...' if len(remaining) > 3 else ''})")

        elif mode == "by_parent":
            node_col = self.node_name_var.get()
            parent_col = self.parent_col_var.get()
            if node_col != "(None)" and parent_col != "(None)":
                lines = [
                    "Root (parent is blank)",
                    "└─ Child A (parent = Root)",
                    "  └─ Grandchild (parent = Child A)",
                    "└─ Child B (parent = Root)"
                ]
            else:
                lines = ["Select Node Name and Parent columns"]

        elif mode == "by_level":
            level_col = self.level_col_var.get()
            name_col = self.level_name_var.get()
            if level_col != "(None)" and name_col != "(None)":
                lines = [
                    "Level 1: Root",
                    "└─ Level 2: Child A",
                    "  └─ Level 3: Grandchild",
                    "└─ Level 2: Child B"
                ]
            else:
                lines = ["Select Level and Node Name columns"]

        self.preview_text.insert('1.0', "\n".join(lines))
        self.preview_text.config(state='disabled')

    def _apply(self):
        """Apply the mapping"""
        mode = self.mode_var.get()

        if mode == "by_columns":
            mapping = []
            for var in self.level_vars:
                col = var.get()
                if col and col != "(None)":
                    mapping.append(col)

            if not mapping:
                messagebox.showwarning("No Mapping", "Please select at least one column for mapping.")
                return

            self.result = {
                'mode': 'by_columns',
                'levels': mapping,
                'include_all': self.include_all_var.get(),
                'color_by_group': self.color_by_group_var.get()
            }

        elif mode == "by_parent":
            node_col = self.node_name_var.get()
            parent_col = self.parent_col_var.get()

            if node_col == "(None)" or parent_col == "(None)":
                messagebox.showwarning("Missing Columns", "Please select both Node Name and Parent columns.")
                return

            self.result = {
                'mode': 'by_parent',
                'node_column': node_col,
                'parent_column': parent_col,
                'include_all': self.include_all_var.get(),
                'color_by_group': self.color_by_group_var.get()
            }

        elif mode == "by_level":
            level_col = self.level_col_var.get()
            name_col = self.level_name_var.get()

            if level_col == "(None)" or name_col == "(None)":
                messagebox.showwarning("Missing Columns", "Please select both Level and Node Name columns.")
                return

            self.result = {
                'mode': 'by_level',
                'level_column': level_col,
                'name_column': name_col,
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
        """Create the mindmap toolbar - compact design with dropdowns"""
        toolbar = ttk.Frame(self)
        toolbar.grid(row=0, column=0, sticky='ew', padx=5, pady=5)

        # Apply button (most important - first)
        ttk.Button(toolbar, text="Apply", command=self._apply_mindmap).pack(side=tk.LEFT, padx=(0, 10))

        # Layout dropdown
        ttk.Label(toolbar, text="Layout:").pack(side=tk.LEFT, padx=(0, 3))
        self.layout_var = tk.StringVar(value="tree_right")
        layout_combo = ttk.Combobox(toolbar, textvariable=self.layout_var, state="readonly", width=15)
        layout_combo['values'] = [
            "Tree Right", "Tree Left", "Tree Down", "Tree Up",
            "Vertical Balanced", "Horizontal Balanced", "Org Chart",
            "Radial", "Free Form"
        ]
        layout_combo.current(0)
        layout_combo.pack(side=tk.LEFT, padx=(0, 10))
        layout_combo.bind('<<ComboboxSelected>>', self._on_layout_change)

        # Style dropdown (combines Lines and Shape)
        style_menu = tk.Menu(self, tearoff=0)

        # Lines submenu
        lines_menu = tk.Menu(style_menu, tearoff=0)
        self.line_var = tk.StringVar(value="curved")
        for line in ["Straight", "Curved", "Orthogonal", "Tapered"]:
            lines_menu.add_radiobutton(label=line, variable=self.line_var, value=line.lower(),
                                       command=self._on_line_change)
        style_menu.add_cascade(label="Line Style", menu=lines_menu)

        # Shapes submenu
        shapes_menu = tk.Menu(style_menu, tearoff=0)
        self.shape_var = tk.StringVar(value="rounded")
        for shape in ["Rectangle", "Rounded", "Ellipse", "Pill", "Diamond", "Hexagon"]:
            shapes_menu.add_radiobutton(label=shape, variable=self.shape_var, value=shape.lower(),
                                        command=self._on_shape_change)
        style_menu.add_cascade(label="Node Shape", menu=shapes_menu)

        style_btn = ttk.Menubutton(toolbar, text="Style")
        style_btn['menu'] = style_menu
        style_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Zoom controls (compact)
        ttk.Label(toolbar, text="Zoom:").pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(toolbar, text="-", width=2, command=self._zoom_out).pack(side=tk.LEFT)
        self.zoom_var = tk.IntVar(value=100)
        self.zoom_slider = ttk.Scale(toolbar, from_=25, to=200, orient=tk.HORIZONTAL,
                                      variable=self.zoom_var, length=60,
                                      command=self._on_zoom_slider)
        self.zoom_slider.pack(side=tk.LEFT)
        self.zoom_label = ttk.Label(toolbar, text="100%", width=4)
        self.zoom_label.pack(side=tk.LEFT)
        ttk.Button(toolbar, text="+", width=2, command=self._zoom_in).pack(side=tk.LEFT, padx=(0, 10))

        # View dropdown
        view_menu = tk.Menu(self, tearoff=0)
        view_menu.add_command(label="Fit All", command=self._fit_all)
        view_menu.add_command(label="Center on Root", command=self._center_root)
        view_menu.add_command(label="Reset Zoom", command=self._zoom_reset)

        view_btn = ttk.Menubutton(toolbar, text="View")
        view_btn['menu'] = view_menu
        view_btn.pack(side=tk.LEFT, padx=(0, 5))

        # More dropdown (less common actions)
        more_menu = tk.Menu(self, tearoff=0)
        more_menu.add_command(label="Column Mapping...", command=self._show_mapping_dialog)
        more_menu.add_command(label="Refresh from Excel", command=self._refresh_mindmap)
        more_menu.add_separator()
        more_menu.add_command(label="Save Mindmap...", command=self._save_mindmap_json)
        more_menu.add_command(label="Load Mindmap...", command=self._load_mindmap_json)

        # Export submenu inside More
        export_submenu = tk.Menu(more_menu, tearoff=0)
        export_submenu.add_command(label="PNG Image...", command=self._export_png)
        export_submenu.add_command(label="SVG Vector...", command=self._export_svg)
        export_submenu.add_separator()
        export_submenu.add_command(label="OPML Outline...", command=self._export_opml)
        export_submenu.add_command(label="FreeMind (.mm)...", command=self._export_freemind)
        export_submenu.add_command(label="Markdown...", command=self._export_markdown)
        export_submenu.add_command(label="Mermaid...", command=self._export_mermaid)
        export_submenu.add_separator()
        export_submenu.add_command(label="JSON Data...", command=self._save_mindmap_json)
        more_menu.add_cascade(label="Export", menu=export_submenu)

        more_btn = ttk.Menubutton(toolbar, text="More")
        more_btn['menu'] = more_menu
        more_btn.pack(side=tk.LEFT)

    def _create_excel_panel(self):
        """Create the left panel with tabbed interface: Excel Data + Outline

        Excel Tab: Column selector + hierarchical treeview from Excel data
        Outline Tab: Free-form text editor with Tab indentation
        """
        left_frame = ttk.Frame(self.paned)

        # Create notebook for tabs
        self.left_notebook = ttk.Notebook(left_frame)
        self.left_notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: Excel Data
        excel_tab = ttk.Frame(self.left_notebook)
        self.left_notebook.add(excel_tab, text="Excel")
        self._create_excel_tab(excel_tab)

        # Tab 2: Outline
        outline_tab = ttk.Frame(self.left_notebook)
        self.left_notebook.add(outline_tab, text="Outline")
        self._create_outline_tab(outline_tab)

        # Auto-sync timer
        self._outline_sync_timer = None
        self._tree_sync_timer = None

        self.paned.add(left_frame, minsize=200, width=350)

    def _create_excel_tab(self, parent):
        """Create Excel data tab with column selector and hierarchy tree"""
        # Column selector frame
        selector_frame = ttk.LabelFrame(parent, text="Select Columns for Hierarchy", padding=5)
        selector_frame.pack(fill=tk.X, padx=5, pady=5)

        # Column checkboxes (will be populated when data loads)
        self.column_vars = {}  # column_name -> BooleanVar
        self.column_frame = ttk.Frame(selector_frame)
        self.column_frame.pack(fill=tk.X)

        # Generate button
        btn_frame = ttk.Frame(selector_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(btn_frame, text="Generate Hierarchy", command=self._generate_from_excel).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Select All", command=self._select_all_columns).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Clear All", command=self._clear_all_columns).pack(side=tk.LEFT)

        # Instructions
        ttk.Label(parent, text="Tab=indent | Backspace=outdent | Enter=new sibling | F2=edit",
                 font=("Calibri", 9), foreground="gray").pack(anchor=tk.W, padx=5)

        # Hierarchy treeview
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        self.hierarchy_tree = ttk.Treeview(tree_frame, selectmode='browse',
                                           yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.hierarchy_tree.heading('#0', text='Hierarchy', anchor='w')
        self.hierarchy_tree.column('#0', width=300, stretch=True)

        vsb.config(command=self.hierarchy_tree.yview)
        hsb.config(command=self.hierarchy_tree.xview)

        self.hierarchy_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        # Bind keyboard events
        self.hierarchy_tree.bind('<Tab>', self._tree_indent)
        self.hierarchy_tree.bind('<Shift-Tab>', self._tree_outdent)
        self.hierarchy_tree.bind('<BackSpace>', self._tree_backspace)
        self.hierarchy_tree.bind('<Return>', self._tree_new_sibling)
        self.hierarchy_tree.bind('<F2>', self._tree_edit_node)
        self.hierarchy_tree.bind('<Double-1>', self._tree_edit_node)
        self.hierarchy_tree.bind('<Delete>', self._tree_delete_node)

        # Right-click context menu
        self.tree_context_menu = tk.Menu(self, tearoff=0)
        self.tree_context_menu.add_command(label="Add Child", command=self._tree_add_child)
        self.tree_context_menu.add_command(label="Add Sibling", command=self._tree_add_sibling)
        self.tree_context_menu.add_separator()
        self.tree_context_menu.add_command(label="Indent (Tab)", command=lambda: self._tree_indent(None))
        self.tree_context_menu.add_command(label="Outdent (Shift+Tab)", command=lambda: self._tree_outdent(None))
        self.tree_context_menu.add_separator()
        self.tree_context_menu.add_command(label="Edit (F2)", command=lambda: self._tree_edit_node(None))
        self.tree_context_menu.add_command(label="Delete", command=lambda: self._tree_delete_node(None))

        self.hierarchy_tree.bind('<Button-3>', self._show_tree_context_menu)
        # Mac right-click
        self.hierarchy_tree.bind('<Button-2>', self._show_tree_context_menu)

        # Keep reference to data_tree for backward compatibility
        self.data_tree = self.hierarchy_tree

    def _create_outline_tab(self, parent):
        """Create outline text editor tab"""
        # Instructions
        ttk.Label(parent, text="Tab=indent | Shift+Tab=outdent | Backspace at start=outdent",
                 font=("Calibri", 9), foreground="gray").pack(anchor=tk.W, padx=5, pady=(5, 0))

        # Text widget
        text_frame = ttk.Frame(parent)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        vsb = ttk.Scrollbar(text_frame, orient="vertical")
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.outline_text = tk.Text(text_frame, wrap=tk.NONE, font=("Consolas", 11),
                                    yscrollcommand=vsb.set, undo=True, tabs=('2c',))
        self.outline_text.pack(fill=tk.BOTH, expand=True)
        vsb.config(command=self.outline_text.yview)

        # Default placeholder
        self.outline_text.insert('1.0', "Root\n\tChild 1\n\t\tGrandchild\n\tChild 2\n")

        # Bind events
        self.outline_text.bind('<KeyRelease>', self._on_outline_change)
        self.outline_text.bind('<Tab>', self._on_tab_press)
        self.outline_text.bind('<Shift-Tab>', self._on_shift_tab_press)
        self.outline_text.bind('<BackSpace>', self._on_backspace_press)

    def _update_column_checkboxes(self):
        """Update column checkboxes when Excel data changes"""
        # Clear existing
        for widget in self.column_frame.winfo_children():
            widget.destroy()
        self.column_vars.clear()

        columns = self.get_columns()
        if not columns:
            ttk.Label(self.column_frame, text="No data - load Excel first",
                     foreground="gray").pack(anchor=tk.W)
            return

        # Create checkbox for each column
        for i, col in enumerate(columns):
            var = tk.BooleanVar(value=True)  # Default checked
            self.column_vars[col] = var
            cb = ttk.Checkbutton(self.column_frame, text=col, variable=var)
            cb.pack(anchor=tk.W)

    def _select_all_columns(self):
        """Select all column checkboxes"""
        for var in self.column_vars.values():
            var.set(True)

    def _clear_all_columns(self):
        """Clear all column checkboxes"""
        for var in self.column_vars.values():
            var.set(False)

    def _generate_from_excel(self):
        """Generate hierarchy tree from Excel data based on selected columns"""
        columns = self.get_columns()
        data = self.get_excel_data()

        if not columns or not data:
            messagebox.showwarning("No Data", "Please load data in Excel View first.")
            return

        # Ensure column checkboxes are updated
        if not self.column_vars:
            self._update_column_checkboxes()

        # Get selected columns in order (only from columns that exist in column_vars)
        selected_cols = []
        for col in columns:
            if col in self.column_vars and self.column_vars[col].get():
                selected_cols.append(col)

        if not selected_cols:
            messagebox.showwarning("No Columns", "Please select at least one column.")
            return

        # Get column indices once
        col_indices = {}
        for col in selected_cols:
            try:
                col_indices[col] = columns.index(col)
            except ValueError:
                pass

        if not col_indices:
            messagebox.showwarning("Error", "Could not find selected columns in data.")
            return

        # Clear existing tree
        for item in self.hierarchy_tree.get_children():
            self.hierarchy_tree.delete(item)

        # Build hierarchy from Excel rows
        # Each row creates a path: col1_value -> col2_value -> col3_value -> ...
        node_cache = {}  # path_tuple -> tree_item_id
        rows_processed = 0
        nodes_created = 0

        for row_idx, row in enumerate(data):
            if not row:
                continue

            # Skip rows that are completely empty
            row_has_data = False
            for cell in row:
                if cell is not None:
                    cell_str = str(cell).strip()
                    if cell_str:
                        row_has_data = True
                        break

            if not row_has_data:
                continue

            rows_processed += 1
            parent_id = ''  # Start at root for each row
            current_path = []

            for col in selected_cols:
                col_idx = col_indices.get(col, -1)
                if col_idx < 0 or col_idx >= len(row):
                    continue

                cell_value = row[col_idx]
                if cell_value is None:
                    continue

                # Convert to string and strip whitespace
                cell_str = str(cell_value).strip()
                if not cell_str:
                    continue

                # Build the path for this node
                current_path.append(cell_str)
                path_key = tuple(current_path)

                # Create node if it doesn't exist
                if path_key not in node_cache:
                    item_id = self.hierarchy_tree.insert(parent_id, 'end', text=cell_str, open=True)
                    node_cache[path_key] = item_id
                    nodes_created += 1

                # Move to this node as parent for next level
                parent_id = node_cache[path_key]

        # Show summary
        if nodes_created == 0:
            messagebox.showinfo("Result", f"No hierarchy created.\nRows checked: {len(data)}\nRows with data: {rows_processed}\nSelected columns: {selected_cols}")
        else:
            # Make sure Excel tab is selected so tree is visible
            if hasattr(self, 'left_notebook'):
                self.left_notebook.select(0)  # Select Excel tab

            # Select first item and give focus to tree for keyboard navigation
            root_items = self.hierarchy_tree.get_children()
            if root_items:
                self.hierarchy_tree.selection_set(root_items[0])
                self.hierarchy_tree.focus(root_items[0])
                self.hierarchy_tree.see(root_items[0])  # Scroll to make visible
                self.hierarchy_tree.focus_set()  # Give keyboard focus to tree

            # Sync to mindmap
            self._sync_tree_to_mindmap()

    # ========================================================================
    # TREE KEYBOARD CONTROLS
    # ========================================================================

    def _tree_indent(self, event):
        """Indent selected item (make it child of previous sibling)"""
        selected = self.hierarchy_tree.selection()
        if not selected:
            return 'break'

        item = selected[0]
        parent = self.hierarchy_tree.parent(item)
        siblings = self.hierarchy_tree.get_children(parent)
        idx = list(siblings).index(item)

        if idx > 0:
            # Move under previous sibling
            new_parent = siblings[idx - 1]
            self.hierarchy_tree.move(item, new_parent, 'end')
            self.hierarchy_tree.item(new_parent, open=True)
            self._schedule_tree_sync()

        return 'break'

    def _tree_outdent(self, event):
        """Outdent selected item (make it sibling of parent)"""
        selected = self.hierarchy_tree.selection()
        if not selected:
            return 'break'

        item = selected[0]
        parent = self.hierarchy_tree.parent(item)

        if parent:  # Not at root level
            grandparent = self.hierarchy_tree.parent(parent)
            parent_idx = list(self.hierarchy_tree.get_children(grandparent)).index(parent)
            self.hierarchy_tree.move(item, grandparent, parent_idx + 1)
            self._schedule_tree_sync()

        return 'break'

    def _tree_backspace(self, event):
        """Handle backspace - outdent if not editing"""
        # Outdent the selected item
        return self._tree_outdent(event)

    def _tree_new_sibling(self, event):
        """Add new sibling node after selected item"""
        selected = self.hierarchy_tree.selection()

        if selected:
            item = selected[0]
            parent = self.hierarchy_tree.parent(item)
            idx = list(self.hierarchy_tree.get_children(parent)).index(item)
            new_item = self.hierarchy_tree.insert(parent, idx + 1, text="New Item")
        else:
            new_item = self.hierarchy_tree.insert('', 'end', text="New Item")

        self.hierarchy_tree.selection_set(new_item)
        self.hierarchy_tree.focus(new_item)
        self._schedule_tree_sync()

        # Start editing the new item
        self.after(100, lambda: self._tree_edit_node(None))
        return 'break'

    def _tree_add_child(self):
        """Add child node under selected item"""
        selected = self.hierarchy_tree.selection()
        parent = selected[0] if selected else ''

        new_item = self.hierarchy_tree.insert(parent, 'end', text="New Item")
        if parent:
            self.hierarchy_tree.item(parent, open=True)

        self.hierarchy_tree.selection_set(new_item)
        self.hierarchy_tree.focus(new_item)
        self._schedule_tree_sync()

        # Start editing
        self.after(100, lambda: self._tree_edit_node(None))

    def _tree_add_sibling(self):
        """Add sibling node (context menu version)"""
        self._tree_new_sibling(None)

    def _tree_delete_node(self, event):
        """Delete selected node"""
        selected = self.hierarchy_tree.selection()
        if not selected:
            return 'break'

        item = selected[0]
        # Select next or previous item before deleting
        next_item = self.hierarchy_tree.next(item) or self.hierarchy_tree.prev(item) or self.hierarchy_tree.parent(item)

        self.hierarchy_tree.delete(item)

        if next_item:
            self.hierarchy_tree.selection_set(next_item)
            self.hierarchy_tree.focus(next_item)

        self._schedule_tree_sync()
        return 'break'

    def _tree_edit_node(self, event):
        """Edit selected node text inline"""
        selected = self.hierarchy_tree.selection()
        if not selected:
            return 'break'

        item = selected[0]
        text = self.hierarchy_tree.item(item, 'text')

        # Get item bbox
        try:
            bbox = self.hierarchy_tree.bbox(item, '#0')
            if not bbox:
                return 'break'
        except:
            return 'break'

        x, y, width, height = bbox

        # Create entry widget for editing
        self.edit_entry = ttk.Entry(self.hierarchy_tree)
        self.edit_entry.insert(0, text)
        self.edit_entry.select_range(0, tk.END)

        self.edit_entry.place(x=x, y=y, width=max(width, 150), height=height)
        self.edit_entry.focus_set()

        def save_edit(e=None):
            new_text = self.edit_entry.get().strip()
            if new_text:
                self.hierarchy_tree.item(item, text=new_text)
            self.edit_entry.destroy()
            self._schedule_tree_sync()

        def cancel_edit(e=None):
            self.edit_entry.destroy()

        self.edit_entry.bind('<Return>', save_edit)
        self.edit_entry.bind('<Escape>', cancel_edit)
        self.edit_entry.bind('<FocusOut>', save_edit)

        return 'break'

    def _show_tree_context_menu(self, event):
        """Show right-click context menu"""
        # Select item under cursor
        item = self.hierarchy_tree.identify_row(event.y)
        if item:
            self.hierarchy_tree.selection_set(item)
            self.hierarchy_tree.focus(item)

        self.tree_context_menu.tk_popup(event.x_root, event.y_root)

    def _schedule_tree_sync(self):
        """Schedule tree sync with debouncing"""
        if self._tree_sync_timer:
            self.after_cancel(self._tree_sync_timer)
        self._tree_sync_timer = self.after(300, self._sync_tree_to_mindmap)

    def _sync_tree_to_mindmap(self):
        """Sync hierarchy tree to mindmap"""
        self._tree_sync_timer = None

        # Clear mindmap
        self.mindmap.clear()

        # Get all root items from hierarchy tree
        root_items = self.hierarchy_tree.get_children()
        if not root_items:
            self._update_status()
            self.sync_status_label.config(text="Empty", foreground="gray")
            return

        # Build mindmap from tree
        color_index = 0

        def add_node_recursive(tree_item, parent_id, depth):
            nonlocal color_index

            text = self.hierarchy_tree.item(tree_item, 'text')
            if not text:
                return None

            # Assign style based on depth
            if depth == 0:
                style = NodeStyle(
                    shape=NodeShape.ELLIPSE,
                    fill_color="#4472C4",
                    border_color="#2E5090",
                    text_color="#FFFFFF",
                    font_size=14,
                    font_bold=True
                )
            elif depth == 1:
                color_set = get_color_set(color_index)
                color_index += 1
                style = NodeStyle(
                    fill_color=f"#{color_set['header']}",
                    border_color=f"#{color_set['header']}",
                    font_bold=True,
                    font_size=12
                )
            else:
                style = NodeStyle(font_size=10)

            node_id = self.mindmap.add_node(text, parent_id=parent_id, style=style)

            # Process children
            for child in self.hierarchy_tree.get_children(tree_item):
                add_node_recursive(child, node_id, depth + 1)

            return node_id

        # If multiple root items, create a virtual root to hold them all
        if len(root_items) > 1:
            # Create virtual root node
            virtual_root_style = NodeStyle(
                shape=NodeShape.ELLIPSE,
                fill_color="#2E5090",
                border_color="#1a3a5c",
                text_color="#FFFFFF",
                font_size=16,
                font_bold=True
            )
            virtual_root_id = self.mindmap.add_node("Mindmap", parent_id=None, style=virtual_root_style)

            # Add all root items as children of virtual root
            for root_item in root_items:
                add_node_recursive(root_item, virtual_root_id, 1)
        else:
            # Single root item - add directly
            add_node_recursive(root_items[0], None, 0)

        # Redraw
        self.mindmap.redraw()
        self.mindmap.fit_all()
        self._update_status()
        self.sync_status_label.config(text="Synced", foreground="green")

    # ========================================================================
    # OUTLINE TAB CONTROLS
    # ========================================================================

    def _on_tab_press(self, event):
        """Handle Tab key - insert tab character"""
        self.outline_text.insert(tk.INSERT, '\t')
        self._schedule_outline_sync()
        return 'break'

    def _on_shift_tab_press(self, event):
        """Handle Shift+Tab - remove one level of indentation"""
        line_start = self.outline_text.index(f"{tk.INSERT} linestart")
        line_text = self.outline_text.get(line_start, f"{line_start} lineend")

        if line_text.startswith('\t'):
            self.outline_text.delete(line_start, f"{line_start}+1c")

        self._schedule_outline_sync()
        return 'break'

    def _on_backspace_press(self, event):
        """Handle Backspace - outdent if at start of line"""
        # Check if cursor is at start of line (after any indentation)
        cursor_pos = self.outline_text.index(tk.INSERT)
        line_start = self.outline_text.index(f"{cursor_pos} linestart")
        text_before = self.outline_text.get(line_start, cursor_pos)

        # If only tabs before cursor, remove one tab (outdent)
        if text_before and text_before.replace('\t', '') == '':
            if text_before.startswith('\t'):
                self.outline_text.delete(line_start, f"{line_start}+1c")
                self._schedule_outline_sync()
                return 'break'

        # Otherwise, let default backspace behavior happen
        return None

    def _on_outline_change(self, event=None):
        """Handle text changes in outline"""
        self._schedule_outline_sync()

    def _schedule_outline_sync(self):
        """Schedule outline sync with debouncing"""
        if self._outline_sync_timer:
            self.after_cancel(self._outline_sync_timer)
        self._outline_sync_timer = self.after(500, self._sync_outline_to_mindmap)

    def _sync_outline_to_mindmap(self):
        """Parse outline text and create mindmap"""
        self._outline_sync_timer = None

        text = self.outline_text.get('1.0', tk.END).strip()
        if not text:
            return

        # Parse outline into hierarchy
        lines = text.split('\n')
        if not lines:
            return

        # Clear existing mindmap
        self.mindmap.clear()

        # First pass: count how many root-level (level 0) items we have
        root_count = 0
        for line in lines:
            if not line.strip():
                continue
            level = len(line) - len(line.lstrip('\t'))
            if level == 0:
                root_count += 1

        # Track parents at each level
        level_nodes = {}  # level -> node_id
        color_index = 0
        virtual_root_id = None

        # If multiple root items, create a virtual root
        if root_count > 1:
            virtual_root_style = NodeStyle(
                shape=NodeShape.ELLIPSE,
                fill_color="#2E5090",
                border_color="#1a3a5c",
                text_color="#FFFFFF",
                font_size=16,
                font_bold=True
            )
            virtual_root_id = self.mindmap.add_node("Mindmap", parent_id=None, style=virtual_root_style)

        for line in lines:
            if not line.strip():
                continue

            # Count leading tabs to determine level
            level = 0
            for char in line:
                if char == '\t':
                    level += 1
                else:
                    break

            node_text = line.strip()
            if not node_text:
                continue

            # Find parent (most recent node at level - 1)
            if virtual_root_id and level == 0:
                # Multiple roots - attach to virtual root
                parent_id = virtual_root_id
                effective_level = 1  # Treat as level 1 for styling
            elif level > 0:
                parent_id = level_nodes.get(level - 1)
                effective_level = level + (1 if virtual_root_id else 0)
            else:
                parent_id = None
                effective_level = 0

            # Assign style based on effective level
            if effective_level == 0:
                style = NodeStyle(
                    shape=NodeShape.ELLIPSE,
                    fill_color="#4472C4",
                    border_color="#2E5090",
                    text_color="#FFFFFF",
                    font_size=14,
                    font_bold=True
                )
            elif effective_level == 1:
                color_set = get_color_set(color_index)
                color_index += 1
                style = NodeStyle(
                    fill_color=f"#{color_set['header']}",
                    border_color=f"#{color_set['header']}",
                    font_bold=True,
                    font_size=12
                )
            else:
                style = NodeStyle(font_size=10)

            # Create node
            node_id = self.mindmap.add_node(node_text, parent_id=parent_id, style=style)
            level_nodes[level] = node_id

            # Clear deeper levels
            for deeper in list(level_nodes.keys()):
                if deeper > level:
                    del level_nodes[deeper]

        # Redraw
        self.mindmap.redraw()
        self.mindmap.fit_all()
        self._update_status()
        self.sync_status_label.config(text="Synced", foreground="green")

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
        """Convert Excel data to mindmap based on column mapping

        Supports three mapping modes:
        - by_columns: Each column represents a hierarchy level
        - by_parent: Two columns (Node Name + Parent Name) for flexible branching
        - by_level: Two columns (Level Number + Node Name) for outline style
        """
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

        # Determine mode (default to by_columns for backward compatibility)
        mode = self.column_mapping.get('mode', 'by_columns')

        if mode == 'by_columns':
            self._load_by_columns(columns, data)
        elif mode == 'by_parent':
            self._load_by_parent(columns, data)
        elif mode == 'by_level':
            self._load_by_level(columns, data)

        # Update tree view
        self._update_tree_view()

        # Redraw mindmap
        self.mindmap.redraw()
        self.mindmap.fit_all()

        # Update status
        self._update_status()

    def _load_by_columns(self, columns: List[str], data: List):
        """Load mindmap using column-based hierarchy (original mode)"""
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
        node_cache = {}
        color_index = 0

        # Process each row
        for row_idx, row in enumerate(data):
            if not row or not any(cell.strip() if isinstance(cell, str) else cell for cell in row):
                continue

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

                # Support multiple branches: split by ; or | delimiters
                import re
                if isinstance(cell_value, str) and (';' in cell_value or '|' in cell_value):
                    values = [v.strip() for v in re.split(r'[;|]', cell_value) if v.strip()]
                else:
                    values = [str(cell_value)]

                created_node_id = None
                for value in values:
                    branch_path = current_path + [value]
                    cache_key = tuple(branch_path)

                    if cache_key in node_cache:
                        created_node_id = node_cache[cache_key]
                    else:
                        # Assign color based on level
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
                            parent_node = self.mindmap.get_node(parent_id)
                            if parent_node:
                                parent_fill = parent_node.style.fill_color.lstrip('#')
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

                        # Store extra data
                        extra_data = {}
                        if self.column_mapping.get('include_all', True):
                            for i, col in enumerate(columns):
                                if col not in level_cols and i < len(row):
                                    val = row[i]
                                    if val and (isinstance(val, str) and val.strip()):
                                        extra_data[col] = str(val).strip()

                        node_id = self.mindmap.add_node(value, parent_id=parent_id,
                                                       style=style, extra_data=extra_data)
                        node_cache[cache_key] = node_id
                        created_node_id = node_id

                if values:
                    current_path.append(values[0])
                    parent_id = created_node_id if created_node_id else parent_id

    def _load_by_parent(self, columns: List[str], data: List):
        """Load mindmap using parent-child column mapping (flexible branching)"""
        node_col = self.column_mapping.get('node_column')
        parent_col = self.column_mapping.get('parent_column')

        if not node_col or not parent_col:
            return

        node_idx = columns.index(node_col) if node_col in columns else -1
        parent_idx = columns.index(parent_col) if parent_col in columns else -1

        if node_idx < 0 or parent_idx < 0:
            return

        # First pass: collect all nodes and their parents
        node_parents = {}  # node_name -> parent_name
        node_rows = {}  # node_name -> row data (for extra columns)

        for row in data:
            if not row or len(row) <= max(node_idx, parent_idx):
                continue

            node_name = row[node_idx]
            parent_name = row[parent_idx] if parent_idx < len(row) else ""

            if isinstance(node_name, str):
                node_name = node_name.strip()
            if isinstance(parent_name, str):
                parent_name = parent_name.strip()

            if node_name:
                node_parents[node_name] = parent_name if parent_name else None
                node_rows[node_name] = row

        # Find root nodes (those with no parent or parent not in our list)
        root_nodes = [n for n, p in node_parents.items() if not p or p not in node_parents]

        if not root_nodes:
            return

        # Create nodes in order
        node_cache = {}  # node_name -> node_id
        color_index = 0

        def create_node_recursive(node_name, parent_id, depth):
            nonlocal color_index

            if node_name in node_cache:
                return node_cache[node_name]

            # Assign style based on depth
            if depth == 0:
                style = NodeStyle(
                    shape=NodeShape.ELLIPSE,
                    fill_color="#4472C4",
                    border_color="#2E5090",
                    text_color="#FFFFFF",
                    font_size=14,
                    font_bold=True
                )
            elif depth == 1:
                color_set = get_color_set(color_index)
                color_index += 1
                style = NodeStyle(
                    fill_color=f"#{color_set['header']}",
                    border_color=f"#{color_set['header']}",
                    font_bold=True,
                    font_size=12
                )
            else:
                style = NodeStyle(font_size=10)

            # Get extra data from row
            extra_data = {}
            if node_name in node_rows and self.column_mapping.get('include_all', True):
                row = node_rows[node_name]
                for i, col in enumerate(columns):
                    if col not in [node_col, parent_col] and i < len(row):
                        val = row[i]
                        if val and (isinstance(val, str) and val.strip()):
                            extra_data[col] = str(val).strip()

            node_id = self.mindmap.add_node(node_name, parent_id=parent_id,
                                           style=style, extra_data=extra_data)
            node_cache[node_name] = node_id

            # Find and create children
            children = [n for n, p in node_parents.items() if p == node_name]
            for child in children:
                create_node_recursive(child, node_id, depth + 1)

            return node_id

        # Create tree starting from root nodes
        for root_name in root_nodes:
            create_node_recursive(root_name, None, 0)

    def _load_by_level(self, columns: List[str], data: List):
        """Load mindmap using level-based column mapping (outline style)"""
        level_col = self.column_mapping.get('level_column')
        name_col = self.column_mapping.get('name_column')

        if not level_col or not name_col:
            return

        level_idx = columns.index(level_col) if level_col in columns else -1
        name_idx = columns.index(name_col) if name_col in columns else -1

        if level_idx < 0 or name_idx < 0:
            return

        # Track parent at each level
        level_parents = {}  # level -> most recent node_id at that level
        color_index = 0

        for row in data:
            if not row or len(row) <= max(level_idx, name_idx):
                continue

            level_val = row[level_idx]
            node_name = row[name_idx]

            if isinstance(node_name, str):
                node_name = node_name.strip()
            if not node_name:
                continue

            # Parse level (handle string or number)
            try:
                level = int(float(str(level_val).strip())) if level_val else 1
            except (ValueError, TypeError):
                level = 1

            # Find parent (most recent node at level - 1)
            parent_id = None
            if level > 1:
                parent_id = level_parents.get(level - 1)

            # Assign style based on level
            if level == 1:
                style = NodeStyle(
                    shape=NodeShape.ELLIPSE,
                    fill_color="#4472C4",
                    border_color="#2E5090",
                    text_color="#FFFFFF",
                    font_size=14,
                    font_bold=True
                )
            elif level == 2:
                color_set = get_color_set(color_index)
                color_index += 1
                style = NodeStyle(
                    fill_color=f"#{color_set['header']}",
                    border_color=f"#{color_set['header']}",
                    font_bold=True,
                    font_size=12
                )
            else:
                style = NodeStyle(font_size=10)

            # Get extra data
            extra_data = {}
            if self.column_mapping.get('include_all', True):
                for i, col in enumerate(columns):
                    if col not in [level_col, name_col] and i < len(row):
                        val = row[i]
                        if val and (isinstance(val, str) and val.strip()):
                            extra_data[col] = str(val).strip()

            node_id = self.mindmap.add_node(node_name, parent_id=parent_id,
                                           style=style, extra_data=extra_data)

            # Update level_parents for this and all deeper levels
            level_parents[level] = node_id
            # Clear deeper levels (they'll get new parents)
            for deeper_level in list(level_parents.keys()):
                if deeper_level > level:
                    del level_parents[deeper_level]

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
            "Org Chart": LayoutType.FOUR_WAY_BALANCED,
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
            "straight": LineStyle.STRAIGHT,
            "curved": LineStyle.CURVED,
            "orthogonal": LineStyle.ORTHOGONAL,
            "tapered": LineStyle.TAPERED,
        }

        selected = self.line_var.get().lower()
        if selected in line_map:
            self.mindmap.set_line_style(line_map[selected])
            self.mindmap.redraw()  # Force redraw to show line change

    def _on_shape_change(self, event=None):
        """Handle shape dropdown change"""
        shape_map = {
            "rectangle": NodeShape.RECTANGLE,
            "rounded": NodeShape.ROUNDED_RECTANGLE,
            "ellipse": NodeShape.ELLIPSE,
            "pill": NodeShape.PILL,
            "diamond": NodeShape.DIAMOND,
            "hexagon": NodeShape.HEXAGON,
        }

        selected = self.shape_var.get().lower()
        if selected in shape_map:
            self.mindmap.set_default_shape(shape_map[selected])
            self.mindmap.redraw()  # Force redraw

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
        """Handle node text edit in mindmap - do NOT sync back to Excel

        One-way sync only: Excel → Mindmap
        Mindmap edits are independent and won't affect Excel data.
        """
        # Just update the status label - no syncing to Excel
        self.sync_status_label.config(text="Mindmap edited (not synced)", foreground="blue")

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
        # Update slider if it exists
        if hasattr(self, 'zoom_var'):
            self.zoom_var.set(zoom_pct)

    def _on_zoom_slider(self, value):
        """Handle zoom slider change"""
        try:
            zoom_pct = int(float(value))
            new_zoom = zoom_pct / 100.0
            if abs(new_zoom - self.mindmap.zoom_level) > 0.01:
                self.mindmap.zoom_level = new_zoom
                self.mindmap.redraw()
                self.zoom_label.config(text=f"{zoom_pct}%")
        except ValueError:
            pass

    def _zoom_reset(self):
        """Reset zoom to 100%"""
        self.mindmap.zoom_level = 1.0
        self.mindmap.redraw()
        self._update_zoom_label()

    def _save_mindmap_json(self):
        """Save mindmap to JSON file"""
        if not self.mindmap.root_id:
            messagebox.showwarning("Save Warning", "No mindmap to save. Please generate a mindmap first.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Save Mindmap",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if filepath:
            try:
                data = self.mindmap.to_dict()
                with open(filepath, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                messagebox.showinfo("Save", f"Mindmap saved to {filepath}")
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save mindmap: {e}")

    def _load_mindmap_json(self):
        """Load mindmap from JSON file"""
        filepath = filedialog.askopenfilename(
            title="Load Mindmap",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if filepath:
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.mindmap.from_dict(data)
                self._update_tree_view()
                self._update_status()
                messagebox.showinfo("Load", f"Mindmap loaded from {filepath}")
            except Exception as e:
                messagebox.showerror("Load Error", f"Failed to load mindmap: {e}")

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
        """Refresh the mindmap from outline or Excel data"""
        # Primary: use outline text if available
        if hasattr(self, 'outline_text') and self.outline_text:
            self._sync_outline_to_mindmap()
        # Fallback: use Excel data with column mapping
        elif self.column_mapping:
            self.load_from_excel_data()
            self.sync_status_label.config(text="Synced", foreground="green")
        else:
            self._show_mapping_dialog()

    def _apply_mindmap(self):
        """Apply mindmap from outline text

        This is the main way to generate mindmap - from the outline panel.
        """
        if hasattr(self, 'outline_text') and self.outline_text:
            self._sync_outline_to_mindmap()
        else:
            # Fallback to Excel data
            columns = self.get_columns()
            data = self.get_excel_data()

            if not columns or not data:
                messagebox.showwarning("No Data", "Please type an outline in the left panel.")
                return

            if not self.column_mapping:
                num_levels = min(4, len(columns))
                self.column_mapping = {
                    'mode': 'by_columns',
                    'levels': columns[:num_levels],
                    'include_all': True,
                    'color_by_group': True
                }

            self.load_from_excel_data()
            self.sync_status_label.config(text="Synced", foreground="green")

    def auto_apply_if_data(self):
        """Auto-apply mindmap when switching to view"""
        # Update column checkboxes when switching to view
        if hasattr(self, 'column_frame'):
            self._update_column_checkboxes()

        # Check if hierarchy tree has content
        if hasattr(self, 'hierarchy_tree') and self.hierarchy_tree.get_children():
            self._sync_tree_to_mindmap()
            return

        # Use outline text if available and non-empty
        if hasattr(self, 'outline_text') and self.outline_text:
            text = self.outline_text.get('1.0', tk.END).strip()
            if text and text != "Root\n\tChild 1\n\t\tGrandchild\n\tChild 2":  # Skip placeholder
                self._sync_outline_to_mindmap()
                return

        # Fallback to Excel data
        columns = self.get_columns()
        data = self.get_excel_data()

        if columns and data:
            # If we have data but no mapping, create default
            if not self.column_mapping:
                num_levels = min(4, len(columns))
                self.column_mapping = {
                    'mode': 'by_columns',
                    'levels': columns[:num_levels],
                    'include_all': True,
                    'color_by_group': True
                }

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

    def _export_opml(self):
        """Export mindmap as OPML (Outline Processor Markup Language)

        OPML is a standard XML format for outlines, widely supported by
        mindmap applications like OmniOutliner, Workflowy, and others.
        """
        if not self.mindmap.nodes:
            messagebox.showwarning("Export Warning", "No mindmap to export. Please generate a mindmap first.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".opml",
            filetypes=[("OPML files", "*.opml"), ("XML files", "*.xml"), ("All files", "*.*")],
            title="Export Mindmap as OPML"
        )

        if filepath:
            try:
                import xml.etree.ElementTree as ET
                from xml.dom import minidom

                # Create OPML structure
                opml = ET.Element('opml', version='2.0')
                head = ET.SubElement(opml, 'head')
                title = ET.SubElement(head, 'title')
                title.text = 'Mindmap Export'

                body = ET.SubElement(opml, 'body')

                def add_outline(parent_elem, node_id):
                    """Recursively add outline elements"""
                    if node_id not in self.mindmap.nodes:
                        return
                    node = self.mindmap.nodes[node_id]
                    outline = ET.SubElement(parent_elem, 'outline', text=node.text)

                    # Add children
                    for child_id in node.children_ids:
                        add_outline(outline, child_id)

                # Start from root
                if self.mindmap.root_id:
                    add_outline(body, self.mindmap.root_id)

                # Pretty print XML
                rough_string = ET.tostring(opml, encoding='unicode')
                reparsed = minidom.parseString(rough_string)
                pretty_xml = reparsed.toprettyxml(indent='  ')

                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(pretty_xml)

                messagebox.showinfo("Export", f"OPML exported to {filepath}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export OPML: {e}")

    def _export_freemind(self):
        """Export mindmap as FreeMind (.mm) format

        FreeMind is a popular open-source mind mapping application.
        The .mm format is XML-based and supported by many other apps.
        """
        if not self.mindmap.nodes:
            messagebox.showwarning("Export Warning", "No mindmap to export. Please generate a mindmap first.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".mm",
            filetypes=[("FreeMind files", "*.mm"), ("XML files", "*.xml"), ("All files", "*.*")],
            title="Export Mindmap as FreeMind"
        )

        if filepath:
            try:
                import xml.etree.ElementTree as ET
                from xml.dom import minidom

                # Create FreeMind map structure
                map_elem = ET.Element('map', version='1.0.1')

                def add_node(parent_elem, node_id, position=None):
                    """Recursively add node elements"""
                    if node_id not in self.mindmap.nodes:
                        return
                    node = self.mindmap.nodes[node_id]

                    # Create node element
                    attrs = {'TEXT': node.text}
                    if position:
                        attrs['POSITION'] = position

                    # Add color if not default
                    fill = node.style.fill_color.lstrip('#')
                    if fill and fill != 'E3F2FD':
                        attrs['BACKGROUND_COLOR'] = f'#{fill}'

                    node_elem = ET.SubElement(parent_elem, 'node', **attrs)

                    # Add children with alternating positions for root children
                    for i, child_id in enumerate(node.children_ids):
                        if node_id == self.mindmap.root_id:
                            # Alternate left/right for root children
                            child_position = 'right' if i % 2 == 0 else 'left'
                            add_node(node_elem, child_id, child_position)
                        else:
                            add_node(node_elem, child_id)

                # Start from root
                if self.mindmap.root_id:
                    add_node(map_elem, self.mindmap.root_id)

                # Pretty print XML
                rough_string = ET.tostring(map_elem, encoding='unicode')
                reparsed = minidom.parseString(rough_string)
                pretty_xml = reparsed.toprettyxml(indent='  ')

                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(pretty_xml)

                messagebox.showinfo("Export", f"FreeMind file exported to {filepath}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export FreeMind: {e}")

    def _export_markdown(self):
        """Export mindmap as Markdown outline

        Creates a hierarchical bullet-point outline that can be
        viewed in any Markdown editor or converted to other formats.
        """
        if not self.mindmap.nodes:
            messagebox.showwarning("Export Warning", "No mindmap to export. Please generate a mindmap first.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".md",
            filetypes=[("Markdown files", "*.md"), ("Text files", "*.txt"), ("All files", "*.*")],
            title="Export Mindmap as Markdown Outline"
        )

        if filepath:
            try:
                lines = []

                def add_node(node_id, depth=0):
                    """Recursively add markdown lines"""
                    if node_id not in self.mindmap.nodes:
                        return
                    node = self.mindmap.nodes[node_id]

                    # Use heading for root, bullets for others
                    if depth == 0:
                        lines.append(f'# {node.text}\n')
                    else:
                        indent = '  ' * (depth - 1)
                        lines.append(f'{indent}- {node.text}')

                    # Add children
                    for child_id in node.children_ids:
                        add_node(child_id, depth + 1)

                # Start from root
                if self.mindmap.root_id:
                    add_node(self.mindmap.root_id)

                content = '\n'.join(lines)
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(content)

                messagebox.showinfo("Export", f"Markdown outline exported to {filepath}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export Markdown: {e}")
