# Excel Master Chart Creator v2.5

A powerful desktop application for macOS that generates beautifully formatted Excel master charts with automatic color-coding and professional styling. Features an Excel-like interface with auto-save, crash recovery, and support for both single-sheet Master Charts and comprehensive 4-tab Drug Charts.

![Application Type: Desktop GUI](https://img.shields.io/badge/Type-Desktop%20GUI-blue)
![Platform: macOS](https://img.shields.io/badge/Platform-macOS-lightgrey)
![Python: 3.7+](https://img.shields.io/badge/Python-3.7%2B-green)
![Version: 2.5](https://img.shields.io/badge/Version-2.5-brightgreen)

## Features

### ✨ Core Features (v2.5)
- **Excel-Like Interface**: Full spreadsheet widget with visible borders, 50 pre-populated rows
- **Auto-Save System**: Automatic background saving every 2 minutes with crash recovery
- **Menu Bar**: File/Edit/View/Help menus with keyboard shortcuts (Cmd+S, Cmd+O, etc.)
- **Recent Files**: Quick access to your last 10 saved files
- **Unsaved Changes Indicator**: Visual warning when you have unsaved work
- **Right-Click Context Menu**: Insert/delete rows, copy/paste, fill down, sort, and more
- **Row Management**: Quick add buttons (+10, +50, +100 rows) and custom quantity input
- **Live Color Preview**: Optional real-time row coloring in the grid (toggle on/off)
- **Delete Empty Rows**: One-click cleanup of blank rows
- **Two Export Formats**:
  - **Master Chart**: Single-sheet format with gradient colors
  - **Comprehensive**: 4-tab format with transposed drug tables
- **Pre-set Templates**: Drug Chart (11 columns), Condition Chart (7 columns), Lab Values (5 columns)
- **Custom Columns**: Define your own column layouts
- **Auto-Open**: Exported Excel files open automatically

### 🎨 3-Shade Color System
The application uses a professional **3-shade gradient system** for each of 10 color sets. Each color set includes:
- **Header shade**: For drug class headers (darker)
- **Main shade**: For primary data cells (medium)
- **Row label shade**: For row labels in transposed tables (lighter)

**10 Color Sets**:
1. **Ice Blue**: Header #B4C6E7, Main #D9E2F3, Row Label #C5D3ED
2. **Seafoam**: Header #A8CCA8, Main #C8E6C9, Row Label #B8D9B9
3. **Light Orchid**: Header #B8A4D0, Main #D1C4E9, Row Label #C4B4DC
4. **Champagne**: Header #E9D4A8, Main #F7E7CE, Row Label #F0DDBB
5. **Sky Blue**: Header #9DC3E6, Main #BDD7EE, Row Label #ADCDEA
6. **Pale Azure**: Header #D0E8F8, Main #F0F8FF, Row Label #E0F0FB
7. **Blush Pink**: Header #F4B8D4, Main #FCE4EC, Row Label #F8CEE0
8. **Soft Lilac**: Header #D0C4E0, Main #EDE7F6, Row Label #DED5EB
9. **Soft Tangerine**: Header #FFCFAA, Main #FFE8D6, Row Label #FFDBC0
10. **Powder Blue**: Header #90C8E8, Main #BBDEFB, Row Label #A5D3F1

Colors automatically rotate based on the first column (Drug Class, Condition, etc.) value changes!

**Special Purpose Colors** (Comprehensive format only):
- **Mnemonic Background**: #FFF9C4 (light yellow) - for memory aids
- **Clinical Pearl Background**: #E1F5DC (light green) - for key clinical insights
- **Analogy Box Background**: #FFE0B2 (light orange) - for conceptual comparisons

### 📊 Excel Export Formats

#### Format 1: Master Chart (Single Sheet)
Perfect for quick reference and printing:
- **Header Row**: Dark blue background (#4472C4), white bold text, centered, frozen
- **Data Rows**: 3-shade gradient backgrounds by group
- **First Column**: Bold formatting
- **Borders**: Thin borders for professional appearance
- **Column Widths**: Automatically sized based on content type
- **Freeze Panes**: Header row stays visible when scrolling
- **Word Wrap**: Enabled on all cells for long content

#### Format 2: Comprehensive (4-Tab Workbook)
Professional medical study guide format:

**Tab 1: Drug Details (Transposed Tables)**
- Drug classes as section headers with header shade
- Drug names as column headers with main shade
- Properties (Mechanism, Uses, etc.) as rows with row label shade
- Compact side-by-side comparison

**Tab 2: Master List (Full Table)**
- Complete drug information in traditional row format
- All columns visible
- Identical to Master Chart format

**Tab 3: Quick Reference (Essential Info)**
- Condensed view with key columns only
- Drug Class, Drug Name, Mechanism, Uses, Adverse Effects
- Optimized for quick lookup

**Tab 4: Index (Alphabetical)**
- All drugs sorted A-Z by drug name
- Quick navigation
- Same formatting as Master List

## Installation

### Prerequisites
- **Python 3.7 or higher** (macOS usually comes with Python 3)
- **pip** (Python package installer)

### Quick Setup

1. **Clone or download this repository**
   ```bash
   cd ~/Desktop
   git clone <repository-url>
   cd Tkinter-Study-Guides
   ```

2. **Install required packages**
   ```bash
   pip3 install -r requirements.txt
   ```

   Or install manually:
   ```bash
   pip3 install openpyxl tksheet
   ```

3. **Run the application**
   ```bash
   python3 excel_master_chart_app.py
   ```

That's it! The application should open with an Excel-like interface.

## Usage Guide

### 1. Select Your Column Layout (Top Panel)

**Choose a preset from the dropdown:**
- **Drug Chart (11 columns)**: Drug Class, Drug Name (Brand), Route, Mechanism, Uses, Adverse Effects, Contraindications, Resistance, Drug Interactions, Drug Combinations, Special Considerations
- **Condition Chart (7 columns)**: Condition, Epidemiology, Risk Factors, Clinical Presentation, Diagnostics, Labs, Treatment
- **Lab Values (5 columns)**: Test Name, Normal Range, Increased In, Decreased In, Clinical Significance
- **Custom**: Define your own columns (one per line in dialog box)

The grid updates immediately with 50 pre-populated empty rows ready for data entry.

### 2. Enter Your Data (Excel-Like Grid)

**Editing cells:**
- **Click once** to select a cell
- **Click again or press Enter** to start editing
- **Type your content** directly in the cell
- **Press Enter or Tab** to move to the next cell
- **Double-click** to edit an existing cell

**Row management:**
- **Add Row**: Adds 1 empty row at the bottom
- **+10 / +50 / +100**: Quick-add multiple rows
- **Custom quantity**: Enter number and click "Add Rows"
- **Delete Empty Rows**: Removes all blank rows automatically
- **Right-click** for context menu:
  - Insert row above/below
  - Delete selected row(s)
  - Copy/Paste
  - Fill down (copy cell to rows below)
  - Sort by column
  - Clear selection

**Menu bar shortcuts:**
- **File**: New (Cmd+N), Open (Cmd+O), Save (Cmd+S), Recent Files, Export
- **Edit**: Undo (Cmd+Z), Redo (Cmd+Shift+Z), Cut/Copy/Paste, Select All
- **View**: Toggle Live Color Preview, Zoom In/Out, Full Screen
- **Help**: About, Documentation

**Tips:**
- The first column is the "grouping" column - colors change when this value changes
- Auto-save runs every 2 minutes in the background
- Unsaved changes indicator (*) appears in title bar
- Live color preview (View menu) shows real-time row coloring

### 3. Save Your Work

**Auto-save:**
- Saves automatically every 2 minutes to temporary file
- Crash recovery offered if app closes unexpectedly

**Manual save:**
- **File → Save** (Cmd+S): Save to JSON file
- **File → Save As**: Save to new location
- **File → Recent Files**: Quick access to last 10 files
- JSON files are portable and can be shared

**Load data:**
- **File → Open** (Cmd+O): Load JSON file
- Example files in `example_data/` folder

### 4. Preview Colors (Optional)

**View → Preview Colors** to see how your data will be color-coded in Excel before exporting.

### 5. Export to Excel

1. **Select format**:
   - **Master Chart**: Single-sheet format (default)
   - **Comprehensive**: 4-tab workbook with transposed drug tables
2. **Enter filename** (default: `Master_Chart.xlsx` or `Comprehensive_Drug_Chart.xlsx`)
3. **Choose save location** (default: Desktop)
4. **Click "🚀 Export to Excel"**
5. The Excel file will be created and automatically opened!

### Example Workflow

```
1. Launch app: python3 excel_master_chart_app.py
2. Select "Drug Chart (11 columns)" from preset dropdown
3. Click in first cell, start typing (or click "+50" for quick rows)
4. Fill in drug information:
   - Click cell → Type → Press Enter to move down
   - Or Tab to move right
5. Use right-click menu for quick operations:
   - Fill down to copy cell to multiple rows
   - Insert rows in the middle as needed
6. Toggle "Live Color Preview" (View menu) to see real-time coloring
7. File → Save (Cmd+S) to save your work as JSON
8. Select export format: "Comprehensive" for 4-tab workbook
9. Click "🚀 Export to Excel"
10. Excel opens with beautifully formatted chart!
```

## Keyboard Shortcuts

**File Operations:**
- **Cmd+N**: New document (clear all data)
- **Cmd+O**: Open JSON file
- **Cmd+S**: Save to JSON
- **Cmd+Shift+S**: Save As

**Editing:**
- **Cmd+Z**: Undo
- **Cmd+Shift+Z**: Redo
- **Cmd+C**: Copy
- **Cmd+V**: Paste
- **Cmd+X**: Cut
- **Cmd+A**: Select All
- **Delete/Backspace**: Clear selected cells

**Grid Navigation:**
- **Enter/Return**: Move down to next cell
- **Tab**: Move right to next cell
- **Shift+Tab**: Move left to previous cell
- **Arrow keys**: Navigate cells
- **Escape**: Cancel cell edit

**View:**
- **Cmd+Plus**: Zoom in
- **Cmd+Minus**: Zoom out
- **Cmd+0**: Reset zoom

## Creating a macOS .app Bundle (Optional)

You can package this application as a native macOS .app that you can double-click to launch!

### Method 1: Using py2app (Recommended)

1. **Install py2app**
   ```bash
   pip3 install py2app
   ```

2. **Create setup file**
   ```bash
   py2applet --make-setup excel_master_chart_app.py
   ```

3. **Build the .app**
   ```bash
   python3 setup.py py2app
   ```

4. **Find your app**
   The application will be in `dist/excel_master_chart_app.app`

5. **Move to Applications folder**
   ```bash
   cp -r dist/excel_master_chart_app.app /Applications/
   ```

6. **Launch from Applications!**
   Now you can double-click the app from your Applications folder or Launchpad.

### Method 2: Using PyInstaller

1. **Install PyInstaller**
   ```bash
   pip3 install pyinstaller
   ```

2. **Build the app**
   ```bash
   pyinstaller --onefile --windowed --name="Excel Master Chart Creator" excel_master_chart_app.py
   ```

3. **Find your app**
   The application will be in `dist/Excel Master Chart Creator.app`

### Troubleshooting .app Creation

**Issue**: "App can't be opened because it is from an unidentified developer"

**Solution**:
```bash
# Right-click the app → Open (instead of double-clicking)
# Or remove quarantine attribute:
xattr -cr /Applications/excel_master_chart_app.app
```

**Issue**: Module not found errors

**Solution**: Make sure openpyxl is installed in the same Python environment:
```bash
pip3 install --upgrade openpyxl
```

## Advanced Configuration

### Customizing Color Sets

Edit the `COLOR_SETS` list in `excel_master_chart_app.py` (around line 50):

```python
COLOR_SETS = [
    {'header': 'B4C6E7', 'main': 'D9E2F3', 'row_label': 'C5D3ED'},
    # Add your own 3-shade sets
    # header = darker shade for headers
    # main = medium shade for data
    # row_label = lighter shade for labels
]
```

**Note**: Use hex colors WITHOUT the `#` symbol (e.g., 'B4C6E7' not '#B4C6E7')

### Customizing Column Widths

Edit the `COLUMN_WIDTH_MAP` in `excel_master_chart_app.py`:

```python
COLUMN_WIDTH_MAP = {
    "short": 12,    # Adjust as needed
    "medium": 22,
    "name": 28,
    "long": 35
}
```

## File Structure

```
Tkinter-Study-Guides/
├── excel_master_chart_app.py    # Main application
├── README.md                     # This file
└── example_data/                 # (Optional) Sample JSON files
```

## Requirements

- **Python**: 3.7+
- **openpyxl**: >=3.0.0 (Excel file generation)
- **tksheet**: >=7.0.0 (Excel-like spreadsheet widget)
- **tkinter**: GUI framework (included with Python)

## Platform Compatibility

| Platform | Status | Notes |
|----------|--------|-------|
| macOS | ✅ Fully Supported | Primary platform |
| Windows | ⚠️ Should work | File opening may differ |
| Linux | ⚠️ Should work | File opening may differ |

## Troubleshooting

### Application won't start
```bash
# Check Python version (should be 3.7+)
python3 --version

# Check if tkinter is available
python3 -m tkinter

# Reinstall dependencies
pip3 install --upgrade openpyxl tksheet
```

**Error: `ModuleNotFoundError: No module named 'tksheet'`**
```bash
pip3 install tksheet
```

### Auto-save not working
- Check the title bar for the unsaved changes indicator (*)
- Auto-save runs every 2 minutes automatically
- Manual save: File → Save (Cmd+S)
- Auto-save file location: `.autosave_excel_master_chart.json` in app directory

### Crash recovery dialog won't go away
- The auto-save file exists from a previous session
- Click "Yes" to restore, or "No" to start fresh
- To manually delete: remove `.autosave_excel_master_chart.json`

### Excel file won't open automatically
- The file is still created successfully
- Navigate to the save location and open manually
- On macOS: Check if Microsoft Excel or Numbers is installed

### Colors not appearing correctly in Comprehensive format
- Make sure you have data in the first column (grouping column)
- Use "View → Preview Colors" to verify assignments
- Different groups should have different values in the first column
- Comprehensive format uses 3-shade gradient system

### Grid is empty after selecting preset
- This is normal - 50 empty rows are pre-populated
- Just click a cell and start typing
- Or load example data: File → Open → `example_data/sample_drug_chart.json`

### Right-click menu doesn't appear
- Try clicking directly on a cell (not in the margin)
- Make sure a cell is selected first
- macOS permissions: System Preferences → Security & Privacy → Accessibility

### Can't edit cells
- Click once to select, then click again or press Enter to edit
- Or double-click directly on the cell
- Press Return/Enter to save, Escape to cancel

## Tips & Best Practices

1. **Let auto-save handle it**: Auto-save runs every 2 minutes, but use Cmd+S for peace of mind
2. **Use Quick Add buttons**: Click "+50" to add 50 rows at once instead of clicking "Add Row" 50 times
3. **Right-click is your friend**: Access powerful features like Fill Down, Insert Row, Sort
4. **Preview before export**: Use "View → Preview Colors" to check color assignments
5. **Group logically**: Ensure items in the same group have identical first-column values (case-sensitive!)
6. **Use descriptive names**: Clear column names help with readability
7. **Try both export formats**: Master Chart for printing, Comprehensive for comprehensive study
8. **Delete empty rows before export**: Use "Delete Empty Rows" button to clean up
9. **Use Recent Files**: File → Recent Files for quick access to your work
10. **Enable Live Color Preview**: View → Toggle Live Color Preview to see colors in real-time
11. **Learn keyboard shortcuts**: Cmd+C/V for copy/paste, Cmd+Z for undo work in the grid
12. **Test with small data first**: Try exporting 5-10 rows before adding 100+

## Examples

### Example 1: Drug Chart
```
Drug Class: Beta Blockers
Drug Name: Metoprolol
Route: Oral
Mechanism: Blocks beta-1 receptors
Uses: Hypertension, heart failure
...
```

### Example 2: Condition Chart
```
Condition: Diabetes Mellitus Type 2
Epidemiology: 10% of US adults
Risk Factors: Obesity, sedentary lifestyle
...
```

## Contributing

Feel free to submit issues or pull requests!

## License

MIT License - feel free to use and modify as needed.

## Support

If you encounter issues:
1. Check the Troubleshooting section above
2. Verify all prerequisites are installed
3. Try running with `python3 -v excel_master_chart_app.py` for verbose output

## Acknowledgments

- Built with Python's tkinter framework
- Excel generation powered by openpyxl
- Designed for medical/pharmacy education workflows

---

**Enjoy creating beautiful Excel master charts without touching code!** 🚀📊
