"""
Shared Color Constants for Mindmap and Excel Master Chart Apps
This module contains the color palette used across both applications.
"""

# ============================================================================
# COLOR SETS - 10 Professional Color Palettes with 3 Shades Each
# ============================================================================
# Each set has: header (darkest), main (medium), row_label (lightest)

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

# ============================================================================
# SPECIAL PURPOSE COLORS
# ============================================================================

MNEMONIC_BG = 'E6F3FF'        # Light blue for mnemonics
CLINICAL_PEARL_BG = 'E8F5E9'   # Light green for clinical pearls
ANALOGY_BOX_BG = 'FFF9E6'      # Light yellow for analogies
MAIN_TITLE_COLOR = '4472C4'    # Dark blue for titles

# ============================================================================
# UI COLORS
# ============================================================================

HEADER_BG_COLOR = "#4472C4"
HEADER_FONT_COLOR = "#FFFFFF"
DATA_FONT_COLOR = "#000000"
BORDER_COLOR = "#FFFFFF"
CANVAS_BG_COLOR = "#FAFAFA"
SELECTION_COLOR = "#2196F3"
HOVER_COLOR = "#E3F2FD"

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_color_set(index):
    """Get a color set by index (wraps around if index > 9)"""
    return COLOR_SETS[index % len(COLOR_SETS)]

def hex_to_rgb(hex_color):
    """Convert hex color (with or without #) to RGB tuple"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def rgb_to_hex(rgb):
    """Convert RGB tuple to hex color string (without #)"""
    return '{:02X}{:02X}{:02X}'.format(*rgb)

def lighten_color(hex_color, factor=0.2):
    """Lighten a hex color by a factor (0-1)"""
    r, g, b = hex_to_rgb(hex_color)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return rgb_to_hex((r, g, b))

def darken_color(hex_color, factor=0.2):
    """Darken a hex color by a factor (0-1)"""
    r, g, b = hex_to_rgb(hex_color)
    r = int(r * (1 - factor))
    g = int(g * (1 - factor))
    b = int(b * (1 - factor))
    return rgb_to_hex((r, g, b))

def get_contrast_text_color(bg_hex):
    """Return black or white text color based on background brightness"""
    r, g, b = hex_to_rgb(bg_hex)
    # Calculate perceived brightness
    brightness = (r * 299 + g * 587 + b * 114) / 1000
    return "#000000" if brightness > 128 else "#FFFFFF"
