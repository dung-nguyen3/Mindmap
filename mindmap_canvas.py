#!/usr/bin/env python3
"""
Mindmap Canvas Module - Handles mindmap rendering and interaction
Part of Excel Master Chart Creator with Mindmap Integration
"""

import tkinter as tk
from tkinter import ttk, colorchooser
import math
import uuid
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Any
from enum import Enum

# Import shared colors
try:
    from color_constants import COLOR_SETS, get_color_set, hex_to_rgb, get_contrast_text_color
except ImportError:
    # Fallback if color_constants not available
    COLOR_SETS = [
        {'header': 'B3D4ED', 'main': 'E3F2FD', 'row_label': 'CBE7FA', 'name': 'Ice Blue'},
        {'header': 'A8CCA8', 'main': 'C8E6C9', 'row_label': 'B8D9B9', 'name': 'Seafoam'},
        {'header': 'B8A4D0', 'main': 'D1C4E9', 'row_label': 'C4B4DC', 'name': 'Light Orchid'},
        {'header': 'E0D0B0', 'main': 'F7E7CE', 'row_label': 'EBDBBF', 'name': 'Champagne'},
        {'header': '9DC3E6', 'main': 'BDD7EE', 'row_label': 'AECDEA', 'name': 'Sky Blue'},
        {'header': 'D0E8FF', 'main': 'F0F8FF', 'row_label': 'E0F0FF', 'name': 'Pale Azure'},
        {'header': 'E8C4CC', 'main': 'FCE4EC', 'row_label': 'F2D4DC', 'name': 'Blush Pink'},
        {'header': 'D0C8DC', 'main': 'EDE7F6', 'row_label': 'DED7E9', 'name': 'Soft Lilac'},
        {'header': 'E0C8B0', 'main': 'FFE8D6', 'row_label': 'EFD8C3', 'name': 'Soft Tangerine'},
        {'header': 'A0C4E8', 'main': 'BBDEFB', 'row_label': 'ADD1F1', 'name': 'Powder Blue'},
    ]
    def get_color_set(index):
        return COLOR_SETS[index % len(COLOR_SETS)]
    def hex_to_rgb(hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    def get_contrast_text_color(bg_hex):
        r, g, b = hex_to_rgb(bg_hex)
        brightness = (r * 299 + g * 587 + b * 114) / 1000
        return "#000000" if brightness > 128 else "#FFFFFF"


# ============================================================================
# ENUMS
# ============================================================================

class NodeShape(Enum):
    RECTANGLE = "rectangle"
    ROUNDED_RECTANGLE = "rounded_rectangle"
    ELLIPSE = "ellipse"
    PILL = "pill"
    DIAMOND = "diamond"
    HEXAGON = "hexagon"


class LineStyle(Enum):
    STRAIGHT = "straight"
    CURVED = "curved"
    ORTHOGONAL = "orthogonal"
    TAPERED = "tapered"


class LayoutType(Enum):
    TREE_DOWN = "tree_down"
    TREE_UP = "tree_up"
    TREE_RIGHT = "tree_right"
    TREE_LEFT = "tree_left"
    VERTICAL_BALANCED = "vertical_balanced"
    HORIZONTAL_BALANCED = "horizontal_balanced"
    FOUR_WAY_BALANCED = "four_way_balanced"  # Combined vertical + horizontal
    VERTICAL_CUSTOM = "vertical_custom"
    HORIZONTAL_CUSTOM = "horizontal_custom"
    RADIAL = "radial"
    FREE_FORM = "free_form"


class BranchDirection(Enum):
    UP = "up"
    DOWN = "down"
    LEFT = "left"
    RIGHT = "right"
    AUTO = "auto"


# ============================================================================
# NODE DATA MODEL
# ============================================================================

@dataclass
class NodeStyle:
    """Style properties for a node"""
    shape: NodeShape = NodeShape.ROUNDED_RECTANGLE
    fill_color: str = "#E3F2FD"
    border_color: str = "#B3D4ED"
    border_width: int = 2
    text_color: str = "#000000"
    font_family: str = "Calibri"
    font_size: int = 11
    font_bold: bool = False
    font_italic: bool = False
    padding: int = 10

    # Line style for connection to parent
    line_color: str = "#666666"
    line_width: int = 2
    line_style: LineStyle = LineStyle.CURVED


@dataclass
class MindmapNode:
    """Represents a single node in the mindmap"""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    text: str = ""
    parent_id: Optional[str] = None
    children_ids: List[str] = field(default_factory=list)
    style: NodeStyle = field(default_factory=NodeStyle)

    # Position (set by layout engine or manually in free form)
    x: float = 0
    y: float = 0
    width: float = 100
    height: float = 40

    # For custom layouts
    direction: BranchDirection = BranchDirection.AUTO

    # State
    collapsed: bool = False
    selected: bool = False

    # Canvas item IDs (for tracking drawn elements)
    canvas_ids: Dict[str, int] = field(default_factory=dict)

    # Additional data from Excel columns
    extra_data: Dict[str, str] = field(default_factory=dict)


# ============================================================================
# MINDMAP CANVAS WIDGET
# ============================================================================

class MindmapCanvas(tk.Canvas):
    """Custom canvas widget for rendering and interacting with mindmaps"""

    # Large virtual canvas size for big mindmaps
    VIRTUAL_WIDTH = 10000
    VIRTUAL_HEIGHT = 10000

    def __init__(self, parent, **kwargs):
        # Default canvas settings
        kwargs.setdefault('bg', '#FAFAFA')
        kwargs.setdefault('highlightthickness', 0)

        super().__init__(parent, **kwargs)

        # Data
        self.nodes: Dict[str, MindmapNode] = {}
        self.root_id: Optional[str] = None

        # View state
        self.zoom_level: float = 1.0
        self.pan_offset_x: float = 0
        self.pan_offset_y: float = 0

        # Layout settings
        self.layout_type: LayoutType = LayoutType.TREE_RIGHT
        self.line_style: LineStyle = LineStyle.CURVED
        self.default_shape: NodeShape = NodeShape.ROUNDED_RECTANGLE

        # Spacing settings
        self.h_spacing = 180  # Horizontal spacing between levels
        self.v_spacing = 60   # Vertical spacing between siblings

        # Interaction state
        self.selected_nodes: List[str] = []
        self.dragging: bool = False
        self.drag_start_x: float = 0
        self.drag_start_y: float = 0
        self.drag_node_id: Optional[str] = None
        self.panning: bool = False
        self.pan_start_x: float = 0
        self.pan_start_y: float = 0

        # Editing state
        self.editing_node_id: Optional[str] = None
        self.edit_entry: Optional[tk.Entry] = None
        self.edit_window_id: Optional[int] = None

        # Undo/Redo stacks
        self.undo_stack: List[Dict] = []
        self.redo_stack: List[Dict] = []

        # Callbacks
        self.on_node_selected: Optional[callable] = None
        self.on_node_edited: Optional[callable] = None
        self.on_structure_changed: Optional[callable] = None
        self.on_style_edit: Optional[callable] = None  # For style editor dialog

        # Configure scrolling
        self.configure(scrollregion=(-self.VIRTUAL_WIDTH//2, -self.VIRTUAL_HEIGHT//2,
                                      self.VIRTUAL_WIDTH//2, self.VIRTUAL_HEIGHT//2))

        # Bind events
        self._bind_events()

    def _bind_events(self):
        """Bind all mouse and keyboard events"""
        # Mouse events
        self.bind('<Button-1>', self._on_click)
        self.bind('<Double-Button-1>', self._on_double_click)
        self.bind('<B1-Motion>', self._on_drag)
        self.bind('<ButtonRelease-1>', self._on_release)
        self.bind('<Button-2>', self._on_middle_click)  # Middle button for pan
        self.bind('<B2-Motion>', self._on_pan)
        self.bind('<ButtonRelease-2>', self._on_pan_release)
        self.bind('<Button-3>', self._on_right_click)

        # Zoom with scroll wheel - Ctrl+scroll for zoom, regular scroll for pan
        self.bind('<MouseWheel>', self._on_scroll_pan)  # Default: pan on Mac/Windows
        self.bind('<Shift-MouseWheel>', self._on_scroll_horizontal)  # Horizontal scroll
        self.bind('<Control-MouseWheel>', self._on_scroll_zoom)  # Ctrl+scroll for zoom
        self.bind('<Button-4>', self._on_scroll_pan_up)  # Linux scroll up
        self.bind('<Button-5>', self._on_scroll_pan_down)  # Linux scroll down
        self.bind('<Control-Button-4>', self._on_scroll_up)  # Linux Ctrl+scroll zoom
        self.bind('<Control-Button-5>', self._on_scroll_down)  # Linux Ctrl+scroll zoom

        # Keyboard
        self.bind('<Delete>', self._on_delete)
        self.bind('<BackSpace>', self._on_delete)
        self.bind('<Escape>', self._on_escape)
        self.bind('<Tab>', self._on_tab)
        self.bind('<Return>', self._on_enter)
        self.bind('<F2>', self._on_f2)

        # Enable keyboard focus
        self.bind('<Enter>', lambda e: self.focus_set())

        # Space + drag for panning
        self.bind('<KeyPress-space>', self._on_space_press)
        self.bind('<KeyRelease-space>', self._on_space_release)
        self.space_pressed = False

        # Keyboard shortcuts
        self.bind('<Control-z>', self._on_undo)
        self.bind('<Control-y>', self._on_redo)
        self.bind('<Control-Z>', self._on_undo)
        self.bind('<Control-Y>', self._on_redo)
        self.bind('<Control-c>', lambda e: self._copy_selected())
        self.bind('<Control-x>', lambda e: self._cut_selected())
        self.bind('<Control-v>', self._on_paste)
        self.bind('<Control-C>', lambda e: self._copy_selected())
        self.bind('<Control-X>', lambda e: self._cut_selected())
        self.bind('<Control-V>', self._on_paste)

        # Arrow key navigation
        self.bind('<Left>', self._on_arrow_left)
        self.bind('<Right>', self._on_arrow_right)
        self.bind('<Up>', self._on_arrow_up)
        self.bind('<Down>', self._on_arrow_down)

    # ========================================================================
    # NODE MANAGEMENT
    # ========================================================================

    def add_node(self, text: str, parent_id: Optional[str] = None,
                 style: Optional[NodeStyle] = None, extra_data: Dict = None,
                 save_undo: bool = True) -> str:
        """Add a new node to the mindmap"""
        if save_undo and (self.root_id is not None):  # Save undo unless it's the first node
            self._save_undo_state()

        node = MindmapNode(
            text=text,
            parent_id=parent_id,
            style=style or NodeStyle(shape=self.default_shape),
            extra_data=extra_data or {}
        )

        self.nodes[node.id] = node

        # Update parent's children list
        if parent_id and parent_id in self.nodes:
            self.nodes[parent_id].children_ids.append(node.id)
        elif parent_id is None:
            self.root_id = node.id

        return node.id

    def remove_node(self, node_id: str, remove_children: bool = True, save_undo: bool = True):
        """Remove a node and optionally its children"""
        if node_id not in self.nodes:
            return

        if save_undo:
            self._save_undo_state()

        node = self.nodes[node_id]

        # Remove children first (don't save undo for each child - already saved above)
        if remove_children:
            for child_id in node.children_ids.copy():
                self.remove_node(child_id, remove_children=True, save_undo=False)

        # Remove from parent's children list
        if node.parent_id and node.parent_id in self.nodes:
            parent = self.nodes[node.parent_id]
            if node_id in parent.children_ids:
                parent.children_ids.remove(node_id)

        # Clear canvas items
        for item_id in node.canvas_ids.values():
            self.delete(item_id)

        # Remove from selection
        if node_id in self.selected_nodes:
            self.selected_nodes.remove(node_id)

        # Remove from nodes dict
        del self.nodes[node_id]

        # Update root if needed
        if node_id == self.root_id:
            self.root_id = None

    def update_node_text(self, node_id: str, text: str, save_undo: bool = True):
        """Update the text of a node"""
        if node_id in self.nodes:
            if save_undo:
                self._save_undo_state()
            self.nodes[node_id].text = text
            self._redraw_node(node_id)

    def update_node_style(self, node_id: str, style: NodeStyle, save_undo: bool = True):
        """Update the style of a node"""
        if node_id in self.nodes:
            if save_undo:
                self._save_undo_state()
            self.nodes[node_id].style = style
            self._redraw_node(node_id)

    def get_node(self, node_id: str) -> Optional[MindmapNode]:
        """Get a node by ID"""
        return self.nodes.get(node_id)

    def clear(self):
        """Clear all nodes"""
        self.delete('all')
        self.nodes.clear()
        self.root_id = None
        self.selected_nodes.clear()

    # ========================================================================
    # DRAWING
    # ========================================================================

    def redraw(self):
        """Redraw the entire mindmap"""
        self.delete('all')

        if not self.root_id or self.root_id not in self.nodes:
            return

        # Apply layout
        self._apply_layout()

        # Draw connections first (so they're behind nodes)
        self._draw_all_connections()

        # Draw all nodes
        for node_id in self.nodes:
            self._draw_node(node_id)

        # Update scroll region based on content
        self._update_scroll_region()

    def _draw_node(self, node_id: str):
        """Draw a single node"""
        if node_id not in self.nodes:
            return

        node = self.nodes[node_id]
        style = node.style

        # Apply zoom
        x = node.x * self.zoom_level
        y = node.y * self.zoom_level

        # Calculate text size to determine node size
        font_style = ""
        if style.font_bold:
            font_style += "bold "
        if style.font_italic:
            font_style += "italic"
        font = (style.font_family, int(style.font_size * self.zoom_level), font_style.strip() or "normal")

        # Create temporary text to measure
        temp_text = self.create_text(0, 0, text=node.text, font=font)
        bbox = self.bbox(temp_text)
        self.delete(temp_text)

        if bbox:
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
        else:
            text_width = 80
            text_height = 20

        padding = style.padding * self.zoom_level
        node.width = text_width + padding * 2
        node.height = text_height + padding * 2

        # Node bounds
        x1 = x - node.width / 2
        y1 = y - node.height / 2
        x2 = x + node.width / 2
        y2 = y + node.height / 2

        # Draw shape based on type
        fill = style.fill_color if not style.fill_color.startswith('#') else style.fill_color
        if not fill.startswith('#'):
            fill = f'#{fill}'
        outline = style.border_color if style.border_color.startswith('#') else f'#{style.border_color}'

        # Selection highlight
        if node.selected or node_id in self.selected_nodes:
            outline = '#2196F3'
            border_width = style.border_width + 2
        else:
            border_width = style.border_width

        shape_id = self._draw_shape(style.shape, x1, y1, x2, y2, fill, outline, border_width)
        node.canvas_ids['shape'] = shape_id

        # Draw text
        text_color = style.text_color if style.text_color.startswith('#') else f'#{style.text_color}'
        text_id = self.create_text(x, y, text=node.text, font=font, fill=text_color,
                                   tags=('node_text', f'node_{node_id}'))
        node.canvas_ids['text'] = text_id

        # Bind click events to this node's items
        self.tag_bind(f'node_{node_id}', '<Button-1>', lambda e, nid=node_id: self._on_node_click(e, nid))
        self.tag_bind(f'node_{node_id}', '<Double-Button-1>', lambda e, nid=node_id: self._on_node_double_click(e, nid))
        self.tag_bind(f'node_{node_id}', '<Button-3>', lambda e, nid=node_id: self._on_node_right_click(e, nid))

    def _draw_shape(self, shape: NodeShape, x1, y1, x2, y2, fill, outline, width) -> int:
        """Draw a shape and return canvas item ID"""
        tags = ('node_shape',)

        if shape == NodeShape.RECTANGLE:
            return self.create_rectangle(x1, y1, x2, y2, fill=fill, outline=outline, width=width, tags=tags)

        elif shape == NodeShape.ROUNDED_RECTANGLE:
            return self._create_rounded_rect(x1, y1, x2, y2, radius=10*self.zoom_level,
                                            fill=fill, outline=outline, width=width, tags=tags)

        elif shape == NodeShape.ELLIPSE:
            return self.create_oval(x1, y1, x2, y2, fill=fill, outline=outline, width=width, tags=tags)

        elif shape == NodeShape.PILL:
            radius = (y2 - y1) / 2
            return self._create_rounded_rect(x1, y1, x2, y2, radius=radius,
                                            fill=fill, outline=outline, width=width, tags=tags)

        elif shape == NodeShape.DIAMOND:
            cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
            hw, hh = (x2 - x1) / 2, (y2 - y1) / 2
            points = [cx, y1, x2, cy, cx, y2, x1, cy]
            return self.create_polygon(points, fill=fill, outline=outline, width=width, tags=tags)

        elif shape == NodeShape.HEXAGON:
            cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
            hw, hh = (x2 - x1) / 2, (y2 - y1) / 2
            inset = hw * 0.25
            points = [x1 + inset, y1, x2 - inset, y1, x2, cy, x2 - inset, y2, x1 + inset, y2, x1, cy]
            return self.create_polygon(points, fill=fill, outline=outline, width=width, tags=tags)

        return self.create_rectangle(x1, y1, x2, y2, fill=fill, outline=outline, width=width, tags=tags)

    def _create_rounded_rect(self, x1, y1, x2, y2, radius=10, **kwargs) -> int:
        """Create a rounded rectangle"""
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1,
        ]
        return self.create_polygon(points, smooth=True, **kwargs)

    def _draw_all_connections(self):
        """Draw connections between all nodes"""
        for node_id, node in self.nodes.items():
            if node.parent_id and node.parent_id in self.nodes:
                self._draw_connection(node.parent_id, node_id)

    def _draw_connection(self, parent_id: str, child_id: str):
        """Draw a connection line between parent and child"""
        if parent_id not in self.nodes or child_id not in self.nodes:
            return

        parent = self.nodes[parent_id]
        child = self.nodes[child_id]
        style = child.style

        # Get connection points based on layout
        px, py = parent.x * self.zoom_level, parent.y * self.zoom_level
        cx, cy = child.x * self.zoom_level, child.y * self.zoom_level

        # Determine edge points based on relative positions
        # Parent edge
        if abs(cx - px) > abs(cy - py):
            # Horizontal connection
            if cx > px:
                px1 = px + parent.width / 2
                cx1 = cx - child.width / 2
            else:
                px1 = px - parent.width / 2
                cx1 = cx + child.width / 2
            py1, cy1 = py, cy
        else:
            # Vertical connection
            if cy > py:
                py1 = py + parent.height / 2
                cy1 = cy - child.height / 2
            else:
                py1 = py - parent.height / 2
                cy1 = cy + child.height / 2
            px1, cx1 = px, cx

        line_color = style.line_color if style.line_color.startswith('#') else f'#{style.line_color}'
        line_width = style.line_width * self.zoom_level

        # Draw based on line style
        if style.line_style == LineStyle.STRAIGHT:
            self.create_line(px1, py1, cx1, cy1, fill=line_color, width=line_width,
                           tags=('connection',), smooth=False)

        elif style.line_style == LineStyle.CURVED:
            # Bezier curve
            mid_x = (px1 + cx1) / 2
            mid_y = (py1 + cy1) / 2

            if abs(cx1 - px1) > abs(cy1 - py1):
                # Horizontal curve
                ctrl1_x, ctrl1_y = mid_x, py1
                ctrl2_x, ctrl2_y = mid_x, cy1
            else:
                # Vertical curve
                ctrl1_x, ctrl1_y = px1, mid_y
                ctrl2_x, ctrl2_y = cx1, mid_y

            self.create_line(px1, py1, ctrl1_x, ctrl1_y, ctrl2_x, ctrl2_y, cx1, cy1,
                           fill=line_color, width=line_width, smooth=True, tags=('connection',))

        elif style.line_style == LineStyle.ORTHOGONAL:
            # Right-angle connections
            if abs(cx1 - px1) > abs(cy1 - py1):
                mid_x = (px1 + cx1) / 2
                self.create_line(px1, py1, mid_x, py1, mid_x, cy1, cx1, cy1,
                               fill=line_color, width=line_width, tags=('connection',))
            else:
                mid_y = (py1 + cy1) / 2
                self.create_line(px1, py1, px1, mid_y, cx1, mid_y, cx1, cy1,
                               fill=line_color, width=line_width, tags=('connection',))

        elif style.line_style == LineStyle.TAPERED:
            # Tapered line (thick at parent, thin at child)
            self.create_line(px1, py1, cx1, cy1, fill=line_color,
                           width=line_width * 2, tags=('connection',))
            # Overlay thinner line for taper effect
            self.create_line(px1, py1, cx1, cy1, fill=line_color,
                           width=line_width * 0.5, tags=('connection',))

    def _redraw_node(self, node_id: str):
        """Redraw a single node (after text/style change)"""
        if node_id not in self.nodes:
            return

        node = self.nodes[node_id]

        # Delete old canvas items
        for item_id in node.canvas_ids.values():
            self.delete(item_id)
        node.canvas_ids.clear()

        # Redraw
        self._draw_node(node_id)

    def _update_scroll_region(self):
        """Update scroll region based on content"""
        if not self.nodes:
            return

        # Find bounds of all nodes
        min_x = min(n.x - n.width/2 for n in self.nodes.values()) * self.zoom_level - 100
        max_x = max(n.x + n.width/2 for n in self.nodes.values()) * self.zoom_level + 100
        min_y = min(n.y - n.height/2 for n in self.nodes.values()) * self.zoom_level - 100
        max_y = max(n.y + n.height/2 for n in self.nodes.values()) * self.zoom_level + 100

        # Ensure minimum size
        min_x = min(min_x, -self.VIRTUAL_WIDTH//2)
        max_x = max(max_x, self.VIRTUAL_WIDTH//2)
        min_y = min(min_y, -self.VIRTUAL_HEIGHT//2)
        max_y = max(max_y, self.VIRTUAL_HEIGHT//2)

        self.configure(scrollregion=(min_x, min_y, max_x, max_y))

    # ========================================================================
    # LAYOUT ENGINES
    # ========================================================================

    def _apply_layout(self):
        """Apply the current layout to all nodes"""
        if not self.root_id or self.root_id not in self.nodes:
            return

        layout_functions = {
            LayoutType.TREE_DOWN: self._layout_tree_down,
            LayoutType.TREE_UP: self._layout_tree_up,
            LayoutType.TREE_RIGHT: self._layout_tree_right,
            LayoutType.TREE_LEFT: self._layout_tree_left,
            LayoutType.VERTICAL_BALANCED: self._layout_vertical_balanced,
            LayoutType.HORIZONTAL_BALANCED: self._layout_horizontal_balanced,
            LayoutType.FOUR_WAY_BALANCED: self._layout_four_way_balanced,
            LayoutType.VERTICAL_CUSTOM: self._layout_vertical_custom,
            LayoutType.HORIZONTAL_CUSTOM: self._layout_horizontal_custom,
            LayoutType.RADIAL: self._layout_radial,
            LayoutType.FREE_FORM: lambda: None,  # No auto-layout
        }

        layout_func = layout_functions.get(self.layout_type, self._layout_tree_right)
        layout_func()

    def _get_subtree_height(self, node_id: str) -> float:
        """Calculate the total height needed for a subtree"""
        if node_id not in self.nodes:
            return 0

        node = self.nodes[node_id]

        if not node.children_ids or node.collapsed:
            return self.v_spacing

        total = sum(self._get_subtree_height(cid) for cid in node.children_ids)
        return max(total, self.v_spacing)

    def _get_subtree_width(self, node_id: str) -> float:
        """Calculate the total width needed for a subtree"""
        if node_id not in self.nodes:
            return 0

        node = self.nodes[node_id]

        if not node.children_ids or node.collapsed:
            return self.h_spacing

        total = sum(self._get_subtree_width(cid) for cid in node.children_ids)
        return max(total, self.h_spacing)

    def _layout_tree_right(self):
        """Layout with root on left, branches extending right"""
        def layout_subtree(node_id: str, x: float, y_start: float, y_end: float):
            if node_id not in self.nodes:
                return

            node = self.nodes[node_id]
            node.x = x
            node.y = (y_start + y_end) / 2

            if not node.children_ids or node.collapsed:
                return

            child_x = x + self.h_spacing
            total_height = sum(self._get_subtree_height(cid) for cid in node.children_ids)

            current_y = node.y - total_height / 2
            for child_id in node.children_ids:
                child_height = self._get_subtree_height(child_id)
                layout_subtree(child_id, child_x, current_y, current_y + child_height)
                current_y += child_height

        layout_subtree(self.root_id, 0, -self._get_subtree_height(self.root_id)/2,
                      self._get_subtree_height(self.root_id)/2)

    def _layout_tree_left(self):
        """Layout with root on right, branches extending left"""
        def layout_subtree(node_id: str, x: float, y_start: float, y_end: float):
            if node_id not in self.nodes:
                return

            node = self.nodes[node_id]
            node.x = x
            node.y = (y_start + y_end) / 2

            if not node.children_ids or node.collapsed:
                return

            child_x = x - self.h_spacing
            total_height = sum(self._get_subtree_height(cid) for cid in node.children_ids)

            current_y = node.y - total_height / 2
            for child_id in node.children_ids:
                child_height = self._get_subtree_height(child_id)
                layout_subtree(child_id, child_x, current_y, current_y + child_height)
                current_y += child_height

        layout_subtree(self.root_id, 0, -self._get_subtree_height(self.root_id)/2,
                      self._get_subtree_height(self.root_id)/2)

    def _layout_tree_down(self):
        """Layout with root on top, branches extending down"""
        def layout_subtree(node_id: str, y: float, x_start: float, x_end: float):
            if node_id not in self.nodes:
                return

            node = self.nodes[node_id]
            node.x = (x_start + x_end) / 2
            node.y = y

            if not node.children_ids or node.collapsed:
                return

            child_y = y + self.v_spacing * 1.5
            total_width = sum(self._get_subtree_width(cid) for cid in node.children_ids)

            current_x = node.x - total_width / 2
            for child_id in node.children_ids:
                child_width = self._get_subtree_width(child_id)
                layout_subtree(child_id, child_y, current_x, current_x + child_width)
                current_x += child_width

        layout_subtree(self.root_id, 0, -self._get_subtree_width(self.root_id)/2,
                      self._get_subtree_width(self.root_id)/2)

    def _layout_tree_up(self):
        """Layout with root on bottom, branches extending up"""
        def layout_subtree(node_id: str, y: float, x_start: float, x_end: float):
            if node_id not in self.nodes:
                return

            node = self.nodes[node_id]
            node.x = (x_start + x_end) / 2
            node.y = y

            if not node.children_ids or node.collapsed:
                return

            child_y = y - self.v_spacing * 1.5
            total_width = sum(self._get_subtree_width(cid) for cid in node.children_ids)

            current_x = node.x - total_width / 2
            for child_id in node.children_ids:
                child_width = self._get_subtree_width(child_id)
                layout_subtree(child_id, child_y, current_x, current_x + child_width)
                current_x += child_width

        layout_subtree(self.root_id, 0, -self._get_subtree_width(self.root_id)/2,
                      self._get_subtree_width(self.root_id)/2)

    def _layout_vertical_balanced(self):
        """Layout with root in center, children split up and down"""
        if self.root_id not in self.nodes:
            return

        root = self.nodes[self.root_id]
        root.x = 0
        root.y = 0

        if not root.children_ids:
            return

        # Split children: first half up, second half down
        mid = len(root.children_ids) // 2
        up_children = root.children_ids[:mid]
        down_children = root.children_ids[mid:]

        # Layout up children
        if up_children:
            total_width = sum(self._get_subtree_width(cid) for cid in up_children)
            current_x = -total_width / 2
            for child_id in up_children:
                child_width = self._get_subtree_width(child_id)
                self._layout_subtree_direction(child_id,
                                               (current_x + current_x + child_width) / 2,
                                               -self.v_spacing * 1.5, 'up')
                current_x += child_width

        # Layout down children
        if down_children:
            total_width = sum(self._get_subtree_width(cid) for cid in down_children)
            current_x = -total_width / 2
            for child_id in down_children:
                child_width = self._get_subtree_width(child_id)
                self._layout_subtree_direction(child_id,
                                               (current_x + current_x + child_width) / 2,
                                               self.v_spacing * 1.5, 'down')
                current_x += child_width

    def _layout_horizontal_balanced(self):
        """Layout with root in center, children split left and right"""
        if self.root_id not in self.nodes:
            return

        root = self.nodes[self.root_id]
        root.x = 0
        root.y = 0

        if not root.children_ids:
            return

        # Split children: first half left, second half right
        mid = len(root.children_ids) // 2
        left_children = root.children_ids[:mid]
        right_children = root.children_ids[mid:]

        # Layout left children
        if left_children:
            total_height = sum(self._get_subtree_height(cid) for cid in left_children)
            current_y = -total_height / 2
            for child_id in left_children:
                child_height = self._get_subtree_height(child_id)
                self._layout_subtree_direction(child_id,
                                               -self.h_spacing,
                                               (current_y + current_y + child_height) / 2, 'left')
                current_y += child_height

        # Layout right children
        if right_children:
            total_height = sum(self._get_subtree_height(cid) for cid in right_children)
            current_y = -total_height / 2
            for child_id in right_children:
                child_height = self._get_subtree_height(child_id)
                self._layout_subtree_direction(child_id,
                                               self.h_spacing,
                                               (current_y + current_y + child_height) / 2, 'right')
                current_y += child_height

    def _layout_four_way_balanced(self):
        """Layout with root in center, children split into 4 quadrants (org chart style)"""
        if self.root_id not in self.nodes:
            return

        root = self.nodes[self.root_id]
        root.x = 0
        root.y = 0

        if not root.children_ids:
            return

        # Split children into 4 groups: up-left, up-right, down-left, down-right
        n = len(root.children_ids)
        quarter = max(1, n // 4)

        up_left = root.children_ids[:quarter]
        up_right = root.children_ids[quarter:quarter*2]
        down_left = root.children_ids[quarter*2:quarter*3]
        down_right = root.children_ids[quarter*3:]

        # If odd distribution, balance it out
        if n <= 2:
            up_right = root.children_ids[:n//2] if n > 0 else []
            down_right = root.children_ids[n//2:] if n > 1 else []
            up_left = down_left = []
        elif n <= 4:
            up_left = [root.children_ids[0]] if n > 0 else []
            up_right = [root.children_ids[1]] if n > 1 else []
            down_left = [root.children_ids[2]] if n > 2 else []
            down_right = [root.children_ids[3]] if n > 3 else []

        # Layout up-left quadrant
        for i, child_id in enumerate(up_left):
            self._layout_subtree_direction(child_id,
                                           -self.h_spacing,
                                           -self.v_spacing * (i + 1), 'left')

        # Layout up-right quadrant
        for i, child_id in enumerate(up_right):
            self._layout_subtree_direction(child_id,
                                           self.h_spacing,
                                           -self.v_spacing * (i + 1), 'right')

        # Layout down-left quadrant
        for i, child_id in enumerate(down_left):
            self._layout_subtree_direction(child_id,
                                           -self.h_spacing,
                                           self.v_spacing * (i + 1), 'left')

        # Layout down-right quadrant
        for i, child_id in enumerate(down_right):
            self._layout_subtree_direction(child_id,
                                           self.h_spacing,
                                           self.v_spacing * (i + 1), 'right')

    def _layout_subtree_direction(self, node_id: str, x: float, y: float, direction: str):
        """Layout a subtree in a specific direction"""
        if node_id not in self.nodes:
            return

        node = self.nodes[node_id]
        node.x = x
        node.y = y

        if not node.children_ids or node.collapsed:
            return

        if direction in ('up', 'down'):
            dy = -self.v_spacing * 1.5 if direction == 'up' else self.v_spacing * 1.5
            total_width = sum(self._get_subtree_width(cid) for cid in node.children_ids)
            current_x = x - total_width / 2
            for child_id in node.children_ids:
                child_width = self._get_subtree_width(child_id)
                self._layout_subtree_direction(child_id,
                                              (current_x + current_x + child_width) / 2,
                                              y + dy, direction)
                current_x += child_width
        else:  # left or right
            dx = -self.h_spacing if direction == 'left' else self.h_spacing
            total_height = sum(self._get_subtree_height(cid) for cid in node.children_ids)
            current_y = y - total_height / 2
            for child_id in node.children_ids:
                child_height = self._get_subtree_height(child_id)
                self._layout_subtree_direction(child_id,
                                              x + dx,
                                              (current_y + current_y + child_height) / 2, direction)
                current_y += child_height

    def _layout_vertical_custom(self):
        """Layout with user-assigned up/down directions per branch"""
        if self.root_id not in self.nodes:
            return

        root = self.nodes[self.root_id]
        root.x = 0
        root.y = 0

        if not root.children_ids:
            return

        up_children = [cid for cid in root.children_ids
                      if self.nodes[cid].direction == BranchDirection.UP]
        down_children = [cid for cid in root.children_ids
                        if self.nodes[cid].direction in (BranchDirection.DOWN, BranchDirection.AUTO)]

        # Layout up children
        if up_children:
            total_width = sum(self._get_subtree_width(cid) for cid in up_children)
            current_x = -total_width / 2
            for child_id in up_children:
                child_width = self._get_subtree_width(child_id)
                self._layout_subtree_direction(child_id,
                                               (current_x + current_x + child_width) / 2,
                                               -self.v_spacing * 1.5, 'up')
                current_x += child_width

        # Layout down children
        if down_children:
            total_width = sum(self._get_subtree_width(cid) for cid in down_children)
            current_x = -total_width / 2
            for child_id in down_children:
                child_width = self._get_subtree_width(child_id)
                self._layout_subtree_direction(child_id,
                                               (current_x + current_x + child_width) / 2,
                                               self.v_spacing * 1.5, 'down')
                current_x += child_width

    def _layout_horizontal_custom(self):
        """Layout with user-assigned left/right directions per branch"""
        if self.root_id not in self.nodes:
            return

        root = self.nodes[self.root_id]
        root.x = 0
        root.y = 0

        if not root.children_ids:
            return

        left_children = [cid for cid in root.children_ids
                        if self.nodes[cid].direction == BranchDirection.LEFT]
        right_children = [cid for cid in root.children_ids
                         if self.nodes[cid].direction in (BranchDirection.RIGHT, BranchDirection.AUTO)]

        # Layout left children
        if left_children:
            total_height = sum(self._get_subtree_height(cid) for cid in left_children)
            current_y = -total_height / 2
            for child_id in left_children:
                child_height = self._get_subtree_height(child_id)
                self._layout_subtree_direction(child_id,
                                               -self.h_spacing,
                                               (current_y + current_y + child_height) / 2, 'left')
                current_y += child_height

        # Layout right children
        if right_children:
            total_height = sum(self._get_subtree_height(cid) for cid in right_children)
            current_y = -total_height / 2
            for child_id in right_children:
                child_height = self._get_subtree_height(child_id)
                self._layout_subtree_direction(child_id,
                                               self.h_spacing,
                                               (current_y + current_y + child_height) / 2, 'right')
                current_y += child_height

    def _layout_radial(self):
        """Layout with root in center, branches radiating outward"""
        if self.root_id not in self.nodes:
            return

        root = self.nodes[self.root_id]
        root.x = 0
        root.y = 0

        def layout_radial_subtree(node_id: str, angle_start: float, angle_end: float,
                                  radius: float, level: int):
            if node_id not in self.nodes:
                return

            node = self.nodes[node_id]

            if node_id != self.root_id:
                angle = (angle_start + angle_end) / 2
                node.x = radius * math.cos(angle)
                node.y = radius * math.sin(angle)

            if not node.children_ids or node.collapsed:
                return

            child_count = len(node.children_ids)
            angle_span = angle_end - angle_start
            angle_per_child = angle_span / child_count

            for i, child_id in enumerate(node.children_ids):
                child_angle_start = angle_start + i * angle_per_child
                child_angle_end = child_angle_start + angle_per_child
                layout_radial_subtree(child_id, child_angle_start, child_angle_end,
                                     radius + self.h_spacing, level + 1)

        layout_radial_subtree(self.root_id, 0, 2 * math.pi, 0, 0)

    # ========================================================================
    # EVENT HANDLERS
    # ========================================================================

    def _on_click(self, event):
        """Handle click on canvas"""
        if self.space_pressed:
            # Start panning
            self.panning = True
            self.pan_start_x = event.x
            self.pan_start_y = event.y
            return

        # Convert screen coordinates to canvas coordinates (handles pan offset)
        canvas_x = self.canvasx(event.x)
        canvas_y = self.canvasy(event.y)

        # Check if clicked on empty space
        clicked_items = self.find_overlapping(canvas_x - 2, canvas_y - 2, canvas_x + 2, canvas_y + 2)

        if not clicked_items or all('connection' in self.gettags(item) for item in clicked_items):
            # Clicked on empty space - clear selection
            self._clear_selection()

    def _on_node_click(self, event, node_id: str):
        """Handle click on a node"""
        # Check for Ctrl/Cmd for multi-select
        if event.state & 0x4:  # Ctrl key
            if node_id in self.selected_nodes:
                self.selected_nodes.remove(node_id)
            else:
                self.selected_nodes.append(node_id)
        else:
            self._clear_selection()
            self.selected_nodes = [node_id]

        # Start drag
        self.dragging = True
        self.drag_node_id = node_id
        self.drag_start_x = event.x
        self.drag_start_y = event.y

        self.redraw()

        if self.on_node_selected:
            self.on_node_selected(node_id)

    def _on_double_click(self, event):
        """Handle double-click on canvas"""
        pass

    def _on_node_double_click(self, event, node_id: str):
        """Handle double-click on node - start inline editing"""
        self._start_editing(node_id)

    def _on_right_click(self, event):
        """Handle right-click on canvas"""
        self._show_canvas_context_menu(event)

    def _on_node_right_click(self, event, node_id: str):
        """Handle right-click on node"""
        if node_id not in self.selected_nodes:
            self._clear_selection()
            self.selected_nodes = [node_id]
            self.redraw()

        self._show_node_context_menu(event, node_id)

    def _on_drag(self, event):
        """Handle drag motion"""
        if self.panning:
            # Pan the canvas
            dx = event.x - self.pan_start_x
            dy = event.y - self.pan_start_y
            self.scan_mark(event.x, event.y)
            self.scan_dragto(event.x + dx, event.y + dy, gain=1)
            return

        if self.dragging and self.drag_node_id and self.layout_type == LayoutType.FREE_FORM:
            # Move the node
            dx = (event.x - self.drag_start_x) / self.zoom_level
            dy = (event.y - self.drag_start_y) / self.zoom_level

            node = self.nodes.get(self.drag_node_id)
            if node:
                node.x += dx
                node.y += dy

            self.drag_start_x = event.x
            self.drag_start_y = event.y

            self.redraw()

    def _on_release(self, event):
        """Handle mouse release"""
        self.dragging = False
        self.drag_node_id = None
        self.panning = False

    def _on_middle_click(self, event):
        """Handle middle click for pan start"""
        self.panning = True
        self.scan_mark(event.x, event.y)

    def _on_pan(self, event):
        """Handle pan motion"""
        if self.panning:
            self.scan_dragto(event.x, event.y, gain=1)

    def _on_pan_release(self, event):
        """Handle pan release"""
        self.panning = False

    def _on_scroll_pan(self, event):
        """Handle scroll wheel for panning (Mac trackpad two-finger scroll)"""
        # On Mac, delta is small (1-2), on Windows it's larger (120)
        import platform
        if platform.system() == 'Darwin':  # Mac
            # Mac trackpad: scroll to pan vertically
            self.yview_scroll(-event.delta, "units")
        else:
            # Windows: larger delta values
            self.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def _on_scroll_horizontal(self, event):
        """Handle horizontal scroll (Shift+scroll)"""
        import platform
        if platform.system() == 'Darwin':
            self.xview_scroll(-event.delta, "units")
        else:
            self.xview_scroll(-1 if event.delta > 0 else 1, "units")

    def _on_scroll_zoom(self, event):
        """Handle Ctrl+scroll wheel for zoom"""
        x = self.canvasx(event.x)
        y = self.canvasy(event.y)

        # Calculate zoom factor
        if event.delta > 0:
            factor = 1.1
        else:
            factor = 0.9

        self._zoom(factor, x, y)

    def _on_scroll_pan_up(self, event):
        """Handle Linux scroll up for pan"""
        self.yview_scroll(-3, "units")

    def _on_scroll_pan_down(self, event):
        """Handle Linux scroll down for pan"""
        self.yview_scroll(3, "units")

    def _on_scroll_up(self, event):
        """Handle Linux Ctrl+scroll up for zoom"""
        x = self.canvasx(event.x)
        y = self.canvasy(event.y)
        self._zoom(1.1, x, y)

    def _on_scroll_down(self, event):
        """Handle Linux Ctrl+scroll down for zoom"""
        x = self.canvasx(event.x)
        y = self.canvasy(event.y)
        self._zoom(0.9, x, y)

    def _zoom(self, factor: float, center_x: float, center_y: float):
        """Apply zoom around a center point"""
        new_zoom = self.zoom_level * factor

        # Limit zoom range
        if new_zoom < 0.1:
            new_zoom = 0.1
        elif new_zoom > 3.0:
            new_zoom = 3.0

        if new_zoom == self.zoom_level:
            return

        self.zoom_level = new_zoom
        self.redraw()

    def _on_delete(self, event):
        """Handle delete key"""
        if self.selected_nodes:
            for node_id in self.selected_nodes.copy():
                if node_id != self.root_id:  # Don't delete root
                    self.remove_node(node_id)
            self.selected_nodes.clear()
            self.redraw()

            if self.on_structure_changed:
                self.on_structure_changed()

    def _on_escape(self, event):
        """Handle escape key"""
        if self.editing_node_id:
            self._cancel_editing()
        else:
            self._clear_selection()
            self.redraw()

    def _on_tab(self, event):
        """Handle tab key - add child to selected node"""
        if self.selected_nodes:
            parent_id = self.selected_nodes[0]
            new_id = self.add_node("New Node", parent_id=parent_id)
            self._clear_selection()
            self.selected_nodes = [new_id]
            self.redraw()
            self._start_editing(new_id)

            if self.on_structure_changed:
                self.on_structure_changed()
        return "break"

    def _on_enter(self, event):
        """Handle enter key - add sibling to selected node"""
        if self.editing_node_id:
            self._finish_editing()
            return "break"

        if self.selected_nodes:
            node_id = self.selected_nodes[0]
            node = self.nodes.get(node_id)
            if node and node.parent_id:
                new_id = self.add_node("New Node", parent_id=node.parent_id)
                self._clear_selection()
                self.selected_nodes = [new_id]
                self.redraw()
                self._start_editing(new_id)

                if self.on_structure_changed:
                    self.on_structure_changed()
        return "break"

    def _on_f2(self, event):
        """Handle F2 key - edit selected node"""
        if self.selected_nodes:
            self._start_editing(self.selected_nodes[0])

    def _on_space_press(self, event):
        """Handle space press for pan mode"""
        self.space_pressed = True
        self.config(cursor="fleur")

    def _on_space_release(self, event):
        """Handle space release"""
        self.space_pressed = False
        self.config(cursor="")
        self.panning = False

    def _on_undo(self, event=None):
        """Handle undo (Ctrl+Z)"""
        if self.undo_stack:
            # Save current state to redo stack
            self.redo_stack.append(self.to_dict())
            # Restore from undo stack
            state = self.undo_stack.pop()
            self.from_dict(state)
        return "break"

    def _on_redo(self, event=None):
        """Handle redo (Ctrl+Y)"""
        if self.redo_stack:
            # Save current state to undo stack
            self.undo_stack.append(self.to_dict())
            # Restore from redo stack
            state = self.redo_stack.pop()
            self.from_dict(state)
        return "break"

    def _save_undo_state(self):
        """Save current state to undo stack"""
        self.undo_stack.append(self.to_dict())
        # Clear redo stack on new action
        self.redo_stack.clear()
        # Limit stack size
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)

    def _on_paste(self, event=None):
        """Handle paste (Ctrl+V)"""
        if self.selected_nodes:
            self._paste_to(self.selected_nodes[0])
        elif self.root_id:
            self._paste_to(self.root_id)
        return "break"

    def _on_arrow_left(self, event):
        """Handle left arrow - navigate to parent"""
        if self.editing_node_id:
            return  # Don't navigate while editing

        if self.selected_nodes:
            node = self.nodes.get(self.selected_nodes[0])
            if node and node.parent_id:
                self._clear_selection()
                self.selected_nodes = [node.parent_id]
                self.redraw()
                if self.on_node_selected:
                    self.on_node_selected(node.parent_id)
        return "break"

    def _on_arrow_right(self, event):
        """Handle right arrow - navigate to first child"""
        if self.editing_node_id:
            return  # Don't navigate while editing

        if self.selected_nodes:
            node = self.nodes.get(self.selected_nodes[0])
            if node and node.children_ids:
                first_child = node.children_ids[0]
                self._clear_selection()
                self.selected_nodes = [first_child]
                self.redraw()
                if self.on_node_selected:
                    self.on_node_selected(first_child)
        return "break"

    def _on_arrow_up(self, event):
        """Handle up arrow - navigate to previous sibling"""
        if self.editing_node_id:
            return  # Don't navigate while editing

        if self.selected_nodes:
            node = self.nodes.get(self.selected_nodes[0])
            if node and node.parent_id:
                parent = self.nodes.get(node.parent_id)
                if parent:
                    siblings = parent.children_ids
                    idx = siblings.index(node.id) if node.id in siblings else -1
                    if idx > 0:
                        prev_sibling = siblings[idx - 1]
                        self._clear_selection()
                        self.selected_nodes = [prev_sibling]
                        self.redraw()
                        if self.on_node_selected:
                            self.on_node_selected(prev_sibling)
        return "break"

    def _on_arrow_down(self, event):
        """Handle down arrow - navigate to next sibling"""
        if self.editing_node_id:
            return  # Don't navigate while editing

        if self.selected_nodes:
            node = self.nodes.get(self.selected_nodes[0])
            if node and node.parent_id:
                parent = self.nodes.get(node.parent_id)
                if parent:
                    siblings = parent.children_ids
                    idx = siblings.index(node.id) if node.id in siblings else -1
                    if idx >= 0 and idx < len(siblings) - 1:
                        next_sibling = siblings[idx + 1]
                        self._clear_selection()
                        self.selected_nodes = [next_sibling]
                        self.redraw()
                        if self.on_node_selected:
                            self.on_node_selected(next_sibling)
        return "break"

    # ========================================================================
    # SELECTION
    # ========================================================================

    def _clear_selection(self):
        """Clear all selected nodes"""
        self.selected_nodes.clear()

    def select_node(self, node_id: str):
        """Select a single node"""
        self._clear_selection()
        if node_id in self.nodes:
            self.selected_nodes = [node_id]
            self.redraw()

    # ========================================================================
    # INLINE EDITING
    # ========================================================================

    def _start_editing(self, node_id: str):
        """Start inline editing of a node"""
        if node_id not in self.nodes:
            return

        # Cancel any existing edit first
        if self.editing_node_id:
            self._cleanup_editing()

        self.editing_node_id = node_id
        node = self.nodes[node_id]

        # Create entry widget
        font_size = max(int(node.style.font_size * self.zoom_level), 10)
        self.edit_entry = tk.Entry(self, font=(node.style.font_family, font_size),
                                   justify='center', bd=2, relief='solid')
        self.edit_entry.insert(0, node.text)
        self.edit_entry.select_range(0, tk.END)

        # Use canvas create_window to position at node's canvas coordinates
        # This ensures the entry stays with the node even when scrolled
        x = node.x * self.zoom_level
        y = node.y * self.zoom_level
        entry_width = max(node.width + 20, 120)

        self.edit_window_id = self.create_window(x, y, window=self.edit_entry,
                                                  width=entry_width, height=node.height)

        self.edit_entry.focus_set()
        self.edit_entry.bind('<Return>', lambda e: self._finish_editing())
        self.edit_entry.bind('<Escape>', lambda e: self._cancel_editing())
        # Don't bind FocusOut - it causes issues when clicking elsewhere
        self.edit_entry.bind('<Tab>', lambda e: self._finish_editing())

    def _finish_editing(self):
        """Finish inline editing and save changes"""
        if self.editing_node_id and self.edit_entry:
            new_text = self.edit_entry.get().strip()
            if new_text:
                self.update_node_text(self.editing_node_id, new_text)

                if self.on_node_edited:
                    self.on_node_edited(self.editing_node_id, new_text)

        self._cleanup_editing()

    def _cancel_editing(self):
        """Cancel inline editing without saving"""
        self._cleanup_editing()

    def _cleanup_editing(self):
        """Clean up editing widgets"""
        if hasattr(self, 'edit_window_id') and self.edit_window_id:
            self.delete(self.edit_window_id)
            self.edit_window_id = None
        if self.edit_entry:
            self.edit_entry.destroy()
            self.edit_entry = None
        self.editing_node_id = None
        self.redraw()

    # ========================================================================
    # CONTEXT MENUS
    # ========================================================================

    def _show_canvas_context_menu(self, event):
        """Show context menu for canvas"""
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Add Root Node", command=lambda: self._add_root_node(event))
        menu.add_separator()
        menu.add_command(label="Fit All", command=self.fit_all)
        menu.add_command(label="Center on Root", command=self.center_on_root)
        menu.add_separator()
        menu.add_command(label="Reset Zoom", command=lambda: self._set_zoom(1.0))

        menu.tk_popup(event.x_root, event.y_root)

    def _show_node_context_menu(self, event, node_id: str):
        """Show context menu for a node"""
        menu = tk.Menu(self, tearoff=0)

        menu.add_command(label="Add Child", command=lambda: self._add_child_to(node_id))
        menu.add_command(label="Add Sibling", command=lambda: self._add_sibling_to(node_id))
        menu.add_separator()
        menu.add_command(label="Edit Text", command=lambda: self._start_editing(node_id))
        menu.add_command(label="Edit Style...", command=lambda: self._show_style_dialog(node_id))
        menu.add_separator()
        menu.add_command(label="Cut", command=self._cut_selected)
        menu.add_command(label="Copy", command=self._copy_selected)
        menu.add_command(label="Paste", command=lambda: self._paste_to(node_id))
        menu.add_command(label="Delete", command=lambda: self._delete_node(node_id))
        menu.add_separator()

        # Branch direction submenu
        direction_menu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="Branch Direction", menu=direction_menu)
        direction_menu.add_command(label="Up", command=lambda: self._set_direction(node_id, BranchDirection.UP))
        direction_menu.add_command(label="Down", command=lambda: self._set_direction(node_id, BranchDirection.DOWN))
        direction_menu.add_command(label="Left", command=lambda: self._set_direction(node_id, BranchDirection.LEFT))
        direction_menu.add_command(label="Right", command=lambda: self._set_direction(node_id, BranchDirection.RIGHT))
        direction_menu.add_command(label="Auto", command=lambda: self._set_direction(node_id, BranchDirection.AUTO))

        menu.add_separator()

        node = self.nodes.get(node_id)
        if node and node.children_ids:
            if node.collapsed:
                menu.add_command(label="Expand", command=lambda: self._toggle_collapse(node_id))
            else:
                menu.add_command(label="Collapse", command=lambda: self._toggle_collapse(node_id))

        menu.tk_popup(event.x_root, event.y_root)

    def _add_root_node(self, event):
        """Add a root node at click position"""
        if self.root_id:
            return  # Already have a root

        x = self.canvasx(event.x) / self.zoom_level
        y = self.canvasy(event.y) / self.zoom_level

        self.root_id = self.add_node("Central Topic")
        self.nodes[self.root_id].x = x
        self.nodes[self.root_id].y = y

        self.redraw()

    def _add_child_to(self, node_id: str):
        """Add a child to a node"""
        new_id = self.add_node("New Node", parent_id=node_id)
        self._clear_selection()
        self.selected_nodes = [new_id]
        self.redraw()
        self._start_editing(new_id)

        if self.on_structure_changed:
            self.on_structure_changed()

    def _add_sibling_to(self, node_id: str):
        """Add a sibling to a node"""
        node = self.nodes.get(node_id)
        if node and node.parent_id:
            new_id = self.add_node("New Node", parent_id=node.parent_id)
            self._clear_selection()
            self.selected_nodes = [new_id]
            self.redraw()
            self._start_editing(new_id)

            if self.on_structure_changed:
                self.on_structure_changed()

    def _delete_node(self, node_id: str):
        """Delete a node"""
        if node_id != self.root_id:
            self.remove_node(node_id)
            self.redraw()

            if self.on_structure_changed:
                self.on_structure_changed()

    def _set_direction(self, node_id: str, direction: BranchDirection):
        """Set branch direction for a node"""
        if node_id in self.nodes:
            self.nodes[node_id].direction = direction
            self.redraw()

    def _toggle_collapse(self, node_id: str):
        """Toggle collapse state of a node"""
        if node_id in self.nodes:
            self.nodes[node_id].collapsed = not self.nodes[node_id].collapsed
            self.redraw()

    def _show_style_dialog(self, node_id: str):
        """Show style editing dialog - calls external callback if set"""
        if hasattr(self, 'on_style_edit') and self.on_style_edit:
            self.on_style_edit(node_id)
        else:
            # Fallback: no dialog available
            pass

    def _cut_selected(self):
        """Cut selected nodes"""
        self._copy_selected()
        for node_id in self.selected_nodes.copy():
            if node_id != self.root_id:
                self.remove_node(node_id)
        self.selected_nodes.clear()
        self.redraw()

    def _copy_selected(self):
        """Copy selected nodes to clipboard"""
        # Store in instance variable for now (could use system clipboard)
        self._clipboard = []
        for node_id in self.selected_nodes:
            if node_id in self.nodes:
                self._clipboard.append(self._serialize_node(node_id))

    def _paste_to(self, parent_id: str):
        """Paste clipboard nodes as children of parent"""
        if not hasattr(self, '_clipboard') or not self._clipboard:
            return

        for node_data in self._clipboard:
            self._deserialize_node(node_data, parent_id)

        self.redraw()

        if self.on_structure_changed:
            self.on_structure_changed()

    def _serialize_node(self, node_id: str) -> dict:
        """Serialize a node and its children"""
        if node_id not in self.nodes:
            return {}

        node = self.nodes[node_id]
        return {
            'text': node.text,
            'style': {
                'shape': node.style.shape.value,
                'fill_color': node.style.fill_color,
                'border_color': node.style.border_color,
                'text_color': node.style.text_color,
                'font_size': node.style.font_size,
            },
            'extra_data': node.extra_data,
            'children': [self._serialize_node(cid) for cid in node.children_ids]
        }

    def _deserialize_node(self, data: dict, parent_id: str):
        """Deserialize a node and its children"""
        style = NodeStyle(
            shape=NodeShape(data.get('style', {}).get('shape', 'rounded_rectangle')),
            fill_color=data.get('style', {}).get('fill_color', '#E3F2FD'),
            border_color=data.get('style', {}).get('border_color', '#B3D4ED'),
            text_color=data.get('style', {}).get('text_color', '#000000'),
            font_size=data.get('style', {}).get('font_size', 11),
        )

        node_id = self.add_node(data.get('text', 'Node'), parent_id=parent_id,
                               style=style, extra_data=data.get('extra_data', {}))

        for child_data in data.get('children', []):
            self._deserialize_node(child_data, node_id)

    # ========================================================================
    # VIEW CONTROLS
    # ========================================================================

    def fit_all(self):
        """Zoom to fit all nodes in view"""
        if not self.nodes:
            return

        # Find bounds
        min_x = min(n.x for n in self.nodes.values())
        max_x = max(n.x for n in self.nodes.values())
        min_y = min(n.y for n in self.nodes.values())
        max_y = max(n.y for n in self.nodes.values())

        # Add padding
        padding = 100
        content_width = max_x - min_x + padding * 2
        content_height = max_y - min_y + padding * 2

        # Calculate zoom to fit
        view_width = self.winfo_width()
        view_height = self.winfo_height()

        if view_width <= 0 or view_height <= 0:
            return

        zoom_x = view_width / content_width
        zoom_y = view_height / content_height

        new_zoom = min(zoom_x, zoom_y, 1.0)  # Don't zoom in more than 100%

        self.zoom_level = max(0.1, new_zoom)
        self.redraw()

        # Center the content
        center_x = (min_x + max_x) / 2
        center_y = (min_y + max_y) / 2

        self.xview_moveto(0.5 - center_x * self.zoom_level / self.VIRTUAL_WIDTH)
        self.yview_moveto(0.5 - center_y * self.zoom_level / self.VIRTUAL_HEIGHT)

    def center_on_root(self):
        """Center view on root node"""
        if self.root_id and self.root_id in self.nodes:
            root = self.nodes[self.root_id]
            self.xview_moveto(0.5 - root.x * self.zoom_level / self.VIRTUAL_WIDTH)
            self.yview_moveto(0.5 - root.y * self.zoom_level / self.VIRTUAL_HEIGHT)

    def _set_zoom(self, level: float):
        """Set zoom level"""
        self.zoom_level = max(0.1, min(3.0, level))
        self.redraw()

    def set_layout(self, layout_type: LayoutType):
        """Set the layout type and redraw"""
        self.layout_type = layout_type
        self.redraw()

    def set_line_style(self, style: LineStyle):
        """Set the line style for all connections"""
        self.line_style = style
        for node in self.nodes.values():
            node.style.line_style = style
        self.redraw()

    def set_default_shape(self, shape: NodeShape):
        """Set the default shape for new nodes"""
        self.default_shape = shape

    # ========================================================================
    # EXPORT
    # ========================================================================

    def to_mermaid(self) -> str:
        """Export mindmap to Mermaid markdown format"""
        if not self.root_id:
            return ""

        lines = ["```mermaid", "mindmap"]

        def add_node(node_id: str, indent: int):
            if node_id not in self.nodes:
                return

            node = self.nodes[node_id]
            prefix = "  " * indent

            # Root node uses special syntax
            if node_id == self.root_id:
                lines.append(f"{prefix}root(({node.text}))")
            else:
                lines.append(f"{prefix}{node.text}")

            for child_id in node.children_ids:
                add_node(child_id, indent + 1)

        add_node(self.root_id, 1)
        lines.append("```")

        return "\n".join(lines)

    def to_dict(self) -> dict:
        """Export mindmap to dictionary for JSON serialization"""
        return {
            'root_id': self.root_id,
            'layout': self.layout_type.value,
            'line_style': self.line_style.value,
            'nodes': {
                node_id: {
                    'text': node.text,
                    'parent_id': node.parent_id,
                    'x': node.x,
                    'y': node.y,
                    'direction': node.direction.value,
                    'collapsed': node.collapsed,
                    'style': {
                        'shape': node.style.shape.value,
                        'fill_color': node.style.fill_color,
                        'border_color': node.style.border_color,
                        'border_width': node.style.border_width,
                        'text_color': node.style.text_color,
                        'font_family': node.style.font_family,
                        'font_size': node.style.font_size,
                        'font_bold': node.style.font_bold,
                        'line_color': node.style.line_color,
                        'line_width': node.style.line_width,
                        'line_style': node.style.line_style.value,
                    },
                    'extra_data': node.extra_data,
                }
                for node_id, node in self.nodes.items()
            }
        }

    def from_dict(self, data: dict):
        """Import mindmap from dictionary"""
        self.clear()

        self.layout_type = LayoutType(data.get('layout', 'tree_right'))
        self.line_style = LineStyle(data.get('line_style', 'curved'))

        # First pass: create all nodes
        for node_id, node_data in data.get('nodes', {}).items():
            style_data = node_data.get('style', {})
            style = NodeStyle(
                shape=NodeShape(style_data.get('shape', 'rounded_rectangle')),
                fill_color=style_data.get('fill_color', '#E3F2FD'),
                border_color=style_data.get('border_color', '#B3D4ED'),
                border_width=style_data.get('border_width', 2),
                text_color=style_data.get('text_color', '#000000'),
                font_family=style_data.get('font_family', 'Calibri'),
                font_size=style_data.get('font_size', 11),
                font_bold=style_data.get('font_bold', False),
                line_color=style_data.get('line_color', '#666666'),
                line_width=style_data.get('line_width', 2),
                line_style=LineStyle(style_data.get('line_style', 'curved')),
            )

            node = MindmapNode(
                id=node_id,
                text=node_data.get('text', ''),
                parent_id=node_data.get('parent_id'),
                style=style,
                x=node_data.get('x', 0),
                y=node_data.get('y', 0),
                direction=BranchDirection(node_data.get('direction', 'auto')),
                collapsed=node_data.get('collapsed', False),
                extra_data=node_data.get('extra_data', {}),
            )

            self.nodes[node_id] = node

        # Second pass: rebuild children lists
        for node_id, node in self.nodes.items():
            if node.parent_id and node.parent_id in self.nodes:
                self.nodes[node.parent_id].children_ids.append(node_id)

        self.root_id = data.get('root_id')
        self.redraw()
