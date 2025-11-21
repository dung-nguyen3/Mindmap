#!/usr/bin/env python3
"""
EXCEL MASTER CHART GENERATOR - AUTO-COLOR VERSION
==================================================

Enhanced template with AUTOMATIC COLOR ASSIGNMENT.
Just paste your data array and the script handles all coloring!

Features:
- ✅ Auto-detects drug class changes
- ✅ Auto-assigns colors (10-color rotation)
- ✅ Paste data arrays and run
- ✅ Zero manual color coding

Usage:
    1. Edit this file and paste your data in the DATA section below
    2. Run: python3 excel_master_chart_AUTO_COLOR.py
    3. Done! Perfect colored Excel chart created.

OR use command line:
    python3 excel_master_chart_AUTO_COLOR.py --output "Chart.xlsx"
"""

import argparse
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# COLOR SCHEME (from Excel Master Chart Only.txt)
# =============================================================================

MAIN_TITLE_COLOR = '4472C4'  # Dark blue for header

# 10-color rotation for drug classes/groups
COLOR_SETS = [
    'D9E2F3',  # 0: Ice Blue
    'C8E6C9',  # 1: Seafoam
    'D1C4E9',  # 2: Light Orchid
    'F7E7CE',  # 3: Champagne
    'BDD7EE',  # 4: Sky Blue
    'F0F8FF',  # 5: Pale Azure
    'FCE4EC',  # 6: Blush Pink
    'EDE7F6',  # 7: Soft Lilac
    'FFE8D6',  # 8: Soft Tangerine
    'BBDEFB',  # 9: Powder Blue
]

# White borders (invisible against pastel backgrounds)
thin_border = Border(
    left=Side(style='thin', color='FFFFFF'),
    right=Side(style='thin', color='FFFFFF'),
    top=Side(style='thin', color='FFFFFF'),
    bottom=Side(style='thin', color='FFFFFF')
)

# =============================================================================
# 📋 PASTE YOUR DATA HERE
# =============================================================================

# Column headers - EDIT THESE to match your chart type
HEADERS = [
    'Drug Class',
    'Drug Name (Brand)',
    'Route',
    'Mechanism',
    'Uses',
    'Adverse Effects',
    'Contraindications',
    'Resistance',
    'Drug Interactions',
    'Drug Combinations',
    'Special Considerations'
]

# Data rows - PASTE YOUR DATA ARRAY HERE
# Each row is a list: [column1, column2, column3, ...]
# IMPORTANT: Group by drug class! All drugs in same class should be together.
DATA = [
    # Example format (DELETE THESE and paste your real data):
    # ['NRTI', 'Tenofovir (Viread)', 'Oral', 'Adenosine analogue', 'HIV', 'GI', 'None', 'Rapid', 'None', 'Truvada', 'First line'],
    # ['NRTI', 'Lamivudine (Epivir)', 'Oral', 'Cytosine analogue', 'HIV', 'Headache', 'None', 'Rapid', 'None', 'With abacavir', 'Safe in pregnancy'],
    # ['NNRTI', 'Rilpivirine (Edurant)', 'Oral', 'Binds RT', 'HIV combo', 'Depression', 'None', 'Rapid', 'None', 'Cabenuva', 'With NRTIs'],
]

# Output filename - EDIT THIS
OUTPUT_FILENAME = 'Master_Chart.xlsx'

# =============================================================================
# WORKBOOK CREATION
# =============================================================================

def get_column_width(column_name):
    """Auto-determine column width based on content type."""
    if column_name.lower() in ['route', 'status', 'type']:
        return 12
    if any(term in column_name.lower() for term in ['class', 'category', 'condition']):
        return 22
    if 'name' in column_name.lower():
        return 28
    return 35

def create_master_chart(headers):
    """Create Master Chart workbook with specified column headers."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Master Chart"

    # Set column widths
    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        width = get_column_width(header)
        ws.column_dimensions[col_letter].width = width

    # Create header row
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(1, col_idx, header)
        cell.font = Font(bold=True, size=12, color='FFFFFF')
        cell.fill = PatternFill(start_color=MAIN_TITLE_COLOR,
                               end_color=MAIN_TITLE_COLOR,
                               fill_type='solid')
        cell.alignment = Alignment(horizontal='center',
                                  vertical='center',
                                  wrap_text=True)
        cell.border = thin_border

    ws.row_dimensions[1].height = 25
    ws.freeze_panes = 'A2'

    return wb, ws

def add_data_row(ws, row_num, data_values, color):
    """Add a single data row to the Master Chart."""
    for col_idx, value in enumerate(data_values, start=1):
        cell = ws.cell(row_num, col_idx, value)
        cell.alignment = Alignment(wrap_text=True, vertical='top')
        cell.border = thin_border
        cell.font = Font(size=10, color='000000')

        # First column is bold (grouping/class column)
        if col_idx == 1:
            cell.font = Font(bold=True, size=10, color='000000')

        # Apply background color
        cell.fill = PatternFill(start_color=color,
                               end_color=color,
                               fill_type='solid')

# =============================================================================
# ⭐ AUTO-COLOR ASSIGNMENT (THE MAGIC!)
# =============================================================================

def add_data_with_auto_colors(ws, data):
    """
    Add all data rows with AUTOMATIC color assignment.

    How it works:
    1. Tracks the drug class (first column) for each row
    2. When class changes, assigns next color in rotation
    3. All drugs in same class get same color

    Example:
        Row 1: NRTI → Assigns Ice Blue (color 0)
        Row 2: NRTI → Keeps Ice Blue (same class)
        Row 3: NRTI → Keeps Ice Blue (same class)
        Row 4: NNRTI → Assigns Seafoam (color 1, class changed!)
        Row 5: NNRTI → Keeps Seafoam (same class)
    """
    if not data:
        print("⚠️  No data provided! Please add data to the DATA array.")
        return 0

    previous_class = None
    current_color = None
    color_index = 0
    row_num = 2  # Start after header

    for row_data in data:
        if not row_data:  # Skip empty rows
            continue

        drug_class = row_data[0]  # First column = Drug Class/Group

        # Check if class changed
        if drug_class != previous_class:
            # New class detected - assign next color
            current_color = COLOR_SETS[color_index % len(COLOR_SETS)]
            print(f"  ✓ New group detected: '{drug_class}' → Color {color_index} ({current_color})")
            color_index += 1
            previous_class = drug_class

        # Add row with current color
        add_data_row(ws, row_num, row_data, current_color)
        row_num += 1

    return row_num - 2  # Return total rows added

# =============================================================================
# VALIDATION
# =============================================================================

def validate_data(headers, data):
    """Validate that data matches header count."""
    if not data:
        return True  # Empty data is okay, just creates template

    expected_cols = len(headers)
    issues = []

    for i, row in enumerate(data, start=1):
        if not row:  # Skip empty rows
            continue
        actual_cols = len(row)
        if actual_cols != expected_cols:
            issues.append(f"  Row {i}: Has {actual_cols} columns, expected {expected_cols}")

    if issues:
        print("\n❌ DATA VALIDATION ERRORS:")
        print(f"Expected {expected_cols} columns based on headers.")
        for issue in issues:
            print(issue)
        print("\nFix: Ensure each data row has exactly", expected_cols, "values.\n")
        return False

    return True

# =============================================================================
# MAIN FUNCTION
# =============================================================================

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Generate Excel Master Chart with AUTO-COLORING',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Use data from script:
    python3 excel_master_chart_AUTO_COLOR.py

  Custom output name:
    python3 excel_master_chart_AUTO_COLOR.py --output "HIV_Drugs.xlsx"
        """
    )

    parser.add_argument('--output', help='Output Excel filename', default=OUTPUT_FILENAME)
    args = parser.parse_args()

    print("=" * 70)
    print("EXCEL MASTER CHART GENERATOR - AUTO-COLOR VERSION")
    print("=" * 70)
    print()

    # Validate data
    if not validate_data(HEADERS, DATA):
        return 1

    # Show configuration
    print(f"📊 Configuration:")
    print(f"   Columns: {len(HEADERS)}")
    print(f"   Data rows: {len([d for d in DATA if d])}")
    print(f"   Output: {args.output}")
    print()

    # Create workbook
    print("🔨 Creating Master Chart...")
    wb, ws = create_master_chart(HEADERS)
    print(f"   ✓ Workbook created with {len(HEADERS)} columns")
    print(f"   ✓ Header row formatted (dark blue, white text)")
    print(f"   ✓ Frozen panes at A2")
    print()

    # Add data with auto-coloring
    if DATA and any(DATA):
        print("🎨 Adding data with AUTO-COLOR assignment...")
        rows_added = add_data_with_auto_colors(ws, DATA)
        print(f"   ✓ {rows_added} data rows added")
    else:
        print("⚠️  No data provided - creating empty template")
        print()
        print("To add data:")
        print("  1. Edit this script")
        print("  2. Paste your data array in the DATA section (line 65)")
        print("  3. Run the script again")

    print()

    # Save
    print(f"💾 Saving to: {args.output}")
    wb.save(args.output)
    print()
    print("=" * 70)
    print("✅ SUCCESS!")
    print("=" * 70)

    if DATA and any(DATA):
        print()
        print("📋 Summary:")
        print(f"   • Total rows: {rows_added}")
        print(f"   • Drug classes/groups: {len(set(row[0] for row in DATA if row))}")
        print(f"   • Auto-colored: YES ✅")
        print(f"   • File: {args.output}")
    else:
        print()
        print("Next steps:")
        print("  1. Open this script: excel_master_chart_AUTO_COLOR.py")
        print("  2. Find the DATA section (around line 65)")
        print("  3. Paste your data array (ask Claude to format it)")
        print("  4. Run: python3 excel_master_chart_AUTO_COLOR.py")

    print(2)
    return 0

if __name__ == '__main__':
    sys.exit(main())
