#!/usr/bin/env python3
"""
EXCEL MASTER CHART TEMPLATE GENERATOR
======================================

Generic template for creating Excel Master Charts from any source file.
Follows Excel Master Chart Only.txt instructions exactly.

Usage:
    Interactive mode (asks questions):
        python excel_master_chart_template.py

    Command line mode (provide all arguments):
        python excel_master_chart_template.py \
            --source "path/to/source.txt" \
            --output "Output_Chart.xlsx" \
            --columns "Column1,Column2,Column3"

Features:
    - Supports drug charts, condition charts, lab value charts
    - Auto-assigns colors to groups/classes (10 color rotation)
    - Proper formatting (fonts, alignment, borders, frozen header)
    - Emoji support (🟢 ⚠️ ❗️ ✅ 🚫)
    - Column width auto-adjustment
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
# COLUMN WIDTH GUIDELINES
# =============================================================================

def get_column_width(column_name):
    """
    Auto-determine column width based on content type.
    Based on Excel Master Chart Only.txt guidelines.
    """
    # Short text columns
    if column_name.lower() in ['route', 'status', 'type']:
        return 12

    # Medium text columns
    if any(term in column_name.lower() for term in ['class', 'category', 'condition']):
        return 22

    # Name columns
    if 'name' in column_name.lower():
        return 28

    # Long text columns (default)
    return 35

# =============================================================================
# WORKBOOK CREATION
# =============================================================================

def create_master_chart(headers):
    """
    Create Master Chart workbook with specified column headers.

    Args:
        headers: List of column header names

    Returns:
        (workbook, worksheet) tuple
    """
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

    # Format header row
    ws.row_dimensions[1].height = 25

    # Freeze header row
    ws.freeze_panes = 'A2'

    return wb, ws

# =============================================================================
# DATA ROW ADDITION
# =============================================================================

def add_data_row(ws, row_num, data_values, color):
    """
    Add a single data row to the Master Chart.

    Args:
        ws: Worksheet object
        row_num: Row number to add data to
        data_values: List of values (must match header count)
        color: Background color hex code (without #)
    """
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
# COLOR ASSIGNMENT
# =============================================================================

def get_color_for_group(group_name, group_index):
    """
    Get color for a drug class/group based on rotation index.

    Args:
        group_name: Name of the group/class
        group_index: Index in the list of groups (0-based)

    Returns:
        Color hex code (without #)
    """
    # Rotate through the 10 colors
    color_index = group_index % len(COLOR_SETS)
    return COLOR_SETS[color_index]

# =============================================================================
# INTERACTIVE MODE
# =============================================================================

def interactive_mode():
    """Run in interactive mode - ask user for all parameters."""
    print("=" * 70)
    print("EXCEL MASTER CHART GENERATOR - Interactive Mode")
    print("=" * 70)
    print()

    # Get source file
    print("Step 1: Source File")
    print("-" * 70)
    source_file = input("Enter path to source file (or drag & drop): ").strip().strip("'\"")
    print()

    # Get output file
    print("Step 2: Output File")
    print("-" * 70)
    output_file = input("Enter output filename (e.g., Drug_Chart.xlsx): ").strip()
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'
    print()

    # Get columns
    print("Step 3: Column Headers")
    print("-" * 70)
    print("Common presets:")
    print("  [1] Drug Chart (11 columns)")
    print("      Drug Class, Drug Name (Brand), Route, Mechanism, Uses,")
    print("      Adverse Effects, Contraindications, Resistance,")
    print("      Drug Interactions, Drug Combinations, Special Considerations")
    print()
    print("  [2] Condition Chart (7 columns)")
    print("      Condition, Epidemiology, Risk Factors, Clinical Presentation,")
    print("      Diagnostics, Labs, Treatment")
    print()
    print("  [3] Lab Value Chart (5 columns)")
    print("      Test Name, Normal Range, Increased In, Decreased In,")
    print("      Clinical Significance")
    print()
    print("  [4] Custom (you specify)")
    print()

    choice = input("Select preset [1-4]: ").strip()

    if choice == '1':
        columns = [
            'Drug Class', 'Drug Name (Brand)', 'Route', 'Mechanism', 'Uses',
            'Adverse Effects', 'Contraindications', 'Resistance',
            'Drug Interactions', 'Drug Combinations', 'Special Considerations'
        ]
    elif choice == '2':
        columns = [
            'Condition', 'Epidemiology', 'Risk Factors', 'Clinical Presentation',
            'Diagnostics', 'Labs', 'Treatment'
        ]
    elif choice == '3':
        columns = [
            'Test Name', 'Normal Range', 'Increased In', 'Decreased In',
            'Clinical Significance'
        ]
    else:
        print("\nEnter column headers separated by commas:")
        columns_input = input("Columns: ").strip()
        columns = [col.strip() for col in columns_input.split(',')]

    print()
    print("=" * 70)
    print("CONFIGURATION SUMMARY")
    print("=" * 70)
    print(f"Source file: {source_file}")
    print(f"Output file: {output_file}")
    print(f"Columns ({len(columns)}):")
    for i, col in enumerate(columns, 1):
        print(f"  {i}. {col}")
    print()

    proceed = input("Create chart with this configuration? [y/N]: ").strip().lower()
    if proceed != 'y':
        print("\nCancelled.")
        return None, None, None

    return source_file, output_file, columns

# =============================================================================
# MAIN FUNCTION
# =============================================================================

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Generate Excel Master Chart from source file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Interactive mode:
    python excel_master_chart_template.py

  Command line mode:
    python excel_master_chart_template.py \\
        --source "lecture.txt" \\
        --output "Drug_Chart.xlsx" \\
        --columns "Drug Class,Drug Name,Route,Mechanism,Uses,Adverse Effects"
        """
    )

    parser.add_argument('--source', help='Path to source file')
    parser.add_argument('--output', help='Output Excel filename')
    parser.add_argument('--columns', help='Comma-separated column headers')

    args = parser.parse_args()

    # Determine mode
    if args.source and args.output and args.columns:
        # Command line mode
        source_file = args.source
        output_file = args.output
        columns = [col.strip() for col in args.columns.split(',')]
    else:
        # Interactive mode
        result = interactive_mode()
        if result[0] is None:
            return
        source_file, output_file, columns = result

    # Create the workbook
    print("\n📊 Creating Master Chart...")
    wb, ws = create_master_chart(columns)

    print(f"✅ Workbook created with {len(columns)} columns")
    print(f"✅ Header row formatted (dark blue, white text)")
    print(f"✅ Frozen panes at A2")
    print()

    # Instructions for user
    print("=" * 70)
    print("⚠️  TEMPLATE CREATED - MANUAL DATA ENTRY REQUIRED")
    print("=" * 70)
    print()
    print("The Excel file has been created with the proper structure.")
    print("You now need to ADD YOUR DATA using one of these methods:")
    print()
    print("METHOD 1: Edit this script")
    print("-" * 70)
    print("Add your data by editing this script around line 350.")
    print("Example code to add rows:")
    print()
    print("    row = 2  # Start after header")
    print("    ")
    print("    # Add first item (using first color)")
    print("    add_data_row(ws, row, [")
    print("        'NRTI',  # Drug Class")
    print("        'Tenofovir (Viread)',  # Drug Name")
    print("        'Oral',  # Route")
    print("        '...',  # etc.")
    print("    ], COLOR_SETS[0])  # Ice Blue")
    print("    row += 1")
    print()
    print("METHOD 2: Manual entry in Excel")
    print("-" * 70)
    print("1. Open the generated Excel file")
    print("2. Add your data rows manually")
    print("3. Apply colors using the color codes from this script")
    print()
    print("Color codes for groups (rotate through these):")
    for i, color in enumerate(COLOR_SETS):
        print(f"  Group {i}: #{color}")
    print()

    # Save the workbook
    print(f"💾 Saving to: {output_file}")
    wb.save(output_file)
    print()
    print("✅ Template saved successfully!")
    print()
    print("=" * 70)
    print("NEXT STEPS")
    print("=" * 70)
    print("1. Open the Excel file to verify structure")
    print("2. Add your data using one of the methods above")
    print("3. Verify all groups have consistent colors")
    print("4. Use emojis: 🟢 ⚠️ ❗️ ✅ 🚫")
    print()

if __name__ == '__main__':
    main()
