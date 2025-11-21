# Excel Master Chart Template - Usage Guide

## Overview

The `excel_master_chart_template.py` script creates properly formatted Excel Master Charts following the Excel Master Chart Only.txt template instructions.

**Key Features:**
- ✅ Automatic color rotation (10 colors for drug classes/groups)
- ✅ Proper formatting (fonts, borders, alignment, frozen header)
- ✅ Auto-adjusting column widths
- ✅ Supports drug charts, condition charts, lab value charts
- ✅ Interactive or command-line modes

---

## Quick Start

### Option 1: Interactive Mode (Easiest)

```bash
cd "/Users/kimnguyen/Documents/Study Guide Claude Q3/Claude Templates"
python3 excel_master_chart_template.py
```

The script will ask you:
1. Source file path
2. Output filename
3. Which columns to include (with presets)

### Option 2: Command Line Mode (Fastest)

```bash
python3 excel_master_chart_template.py \
  --source "path/to/source.txt" \
  --output "Drug_Chart.xlsx" \
  --columns "Drug Class,Drug Name,Route,Mechanism,Uses,Adverse Effects"
```

---

## Column Presets

### Preset 1: Drug Chart (11 columns)
```
Drug Class, Drug Name (Brand), Route, Mechanism, Uses,
Adverse Effects, Contraindications, Resistance,
Drug Interactions, Drug Combinations, Special Considerations
```

### Preset 2: Condition Chart (7 columns)
```
Condition, Epidemiology, Risk Factors, Clinical Presentation,
Diagnostics, Labs, Treatment
```

### Preset 3: Lab Value Chart (5 columns)
```
Test Name, Normal Range, Increased In, Decreased In,
Clinical Significance
```

### Preset 4: Custom
Specify your own columns separated by commas.

---

## How to Use with Claude

### Step 1: Generate the Template Structure

Run the script to create the Excel file with proper formatting:

```bash
python3 excel_master_chart_template.py \
  --source "24 HIV.txt" \
  --output "HIV_Master_Chart.xlsx" \
  --columns "Drug Class,Drug Name,Route,Mechanism,Uses,Adverse Effects,Contraindications"
```

This creates an Excel file with:
- ✅ Header row (dark blue, white text)
- ✅ Frozen panes
- ✅ Proper column widths
- ✅ Ready for data entry

### Step 2: Ask Claude to Fill in the Data

**Option A: Ask Claude to edit the Python script**

```
Claude, read the source file "24 HIV.txt" and edit the
excel_master_chart_template.py script to add all the HIV drugs
with their information. Use the add_data_row() function and
assign colors based on drug class.
```

**Option B: Ask Claude to create data for manual entry**

```
Claude, extract all HIV drugs from "24 HIV.txt" and format them
for the Excel Master Chart with these columns: Drug Class, Drug Name,
Route, Mechanism, Uses, Adverse Effects, Contraindications.
Give me the data in a format I can copy-paste into Excel.
```

**Option C: Ask Claude to create a custom script**

```
Claude, create a Python script that reads "24 HIV.txt" and
generates a complete Excel Master Chart using the
excel_master_chart_template.py functions.
```

---

## Color Assignment Rules

The script rotates through 10 colors for drug classes/groups:

| Index | Color Name | Hex Code | Use For |
|-------|------------|----------|---------|
| 0 | Ice Blue | #D9E2F3 | First drug class |
| 1 | Seafoam | #C8E6C9 | Second drug class |
| 2 | Light Orchid | #D1C4E9 | Third drug class |
| 3 | Champagne | #F7E7CE | Fourth drug class |
| 4 | Sky Blue | #BDD7EE | Fifth drug class |
| 5 | Pale Azure | #F0F8FF | Sixth drug class |
| 6 | Blush Pink | #FCE4EC | Seventh drug class |
| 7 | Soft Lilac | #EDE7F6 | Eighth drug class |
| 8 | Soft Tangerine | #FFE8D6 | Ninth drug class |
| 9 | Powder Blue | #BBDEFB | Tenth drug class |

**Important:** All drugs in the SAME class get the SAME color.

---

## Example Workflows

### Example 1: Drug Chart with Claude's Help

```bash
# Step 1: Create template structure
python3 excel_master_chart_template.py \
  --source "24 HIV.txt" \
  --output "HIV_Master_Chart.xlsx" \
  --columns "Drug Class,Drug Name (Brand),Route,Mechanism,Uses,Adverse Effects,Contraindications,Resistance,Drug Interactions,Drug Combinations,Special Considerations"

# Step 2: Ask Claude
"Claude, read 24 HIV.txt and create a Python script that uses the
excel_master_chart_template.py to fill in all HIV drugs with complete information."
```

### Example 2: Condition Chart

```bash
# Create template
python3 excel_master_chart_template.py \
  --source "diabetes_lecture.txt" \
  --output "Diabetes_Conditions.xlsx" \
  --columns "Condition,Epidemiology,Risk Factors,Clinical Presentation,Diagnostics,Labs,Treatment"

# Ask Claude to fill it
"Claude, extract all diabetes-related conditions from diabetes_lecture.txt
and populate the Excel Master Chart."
```

### Example 3: Quick Manual Entry

```bash
# Create template with minimal columns
python3 excel_master_chart_template.py \
  --source "notes.txt" \
  --output "Quick_Reference.xlsx" \
  --columns "Topic,Key Points,Clinical Pearls"

# Open and manually fill in the Excel file
open Quick_Reference.xlsx
```

---

## Emoji Usage

Use these standardized emojis in your charts:

| Emoji | Meaning | Example Usage |
|-------|---------|---------------|
| 🟢 | First-line/Preferred | "🟢 First line for naïve patients" |
| ⚠️ | Warning/Serious adverse effect | "⚠️ Lactic acidosis (rare)" |
| ❗️ | Critical information | "❗️ Screen HLA-B*5701 BEFORE use" |
| ✅ | Safe/Positive | "✅ Safe in pregnancy" |
| 🚫 | Absolute contraindication | "🚫 Sulfonamide allergy" |

---

## Formatting Specifications

Following Excel Master Chart Only.txt template:

### Header Row (Row 1)
- Background: #4472C4 (dark blue)
- Font: Calibri, Bold, Size 12, White
- Alignment: Center horizontal, center vertical
- Word wrap: Enabled
- Row height: 25
- **Frozen at row 2**

### Data Rows (Row 2+)
- Font: Calibri, Size 10, Black (#000000)
- Column A (first column): Bold
- Other columns: Regular weight
- Alignment: Left horizontal, top vertical
- Word wrap: Enabled
- Borders: White (#FFFFFF) - invisible
- Background: Pastel colors (rotated by drug class)

### Column Widths (auto-assigned)
- Short columns (Route, Status): 12 width
- Medium columns (Drug Class, Condition): 22 width
- Name columns: 28 width
- Long text columns (default): 35 width

---

## Troubleshooting

### "Module 'openpyxl' not found"
```bash
pip3 install openpyxl
```

### "File not found"
Use absolute paths or drag-and-drop the file when prompted in interactive mode.

### "Wrong number of columns"
Ensure the number of values in `add_data_row()` matches the number of column headers exactly.

### "Colors not showing"
Check that you're using the correct hex codes from `COLOR_SETS` array without the # symbol.

---

## Advanced: Editing the Script to Add Data

If you want to add data directly in the Python script:

1. Open `excel_master_chart_template.py`
2. Find the `main()` function (around line 350)
3. Before `wb.save(output_file)`, add your data:

```python
# Initialize row counter
row = 2

# Add first drug class (use COLOR_SETS[0])
add_data_row(ws, row, [
    'NRTI',
    'Tenofovir (Viread)',
    'Oral',
    'Nucleoside analogue → inhibits RT',
    '🟢 HIV - First line',
    '⚠️ Lactic acidosis',
    'Severe renal impairment',
    'Rapid resistance if alone',
    'None major',
    'Truvada (with emtricitabine)',
    'Very effective for PrEP'
], COLOR_SETS[0])  # Ice Blue
row += 1

# Add more drugs from same class (same color)
add_data_row(ws, row, [
    'NRTI',
    'Lamivudine (Epivir)',
    'Oral',
    'Cytosine analogue → inhibits RT',
    # ... etc
], COLOR_SETS[0])  # Ice Blue (same class)
row += 1

# Start new drug class (use COLOR_SETS[1])
add_data_row(ws, row, [
    'NNRTI',
    'Rilpivirine (Edurant)',
    # ... etc
], COLOR_SETS[1])  # Seafoam (new class)
row += 1
```

---

## Tips for Efficiency

1. **Use the template for structure only**
   - Let the script create the formatting
   - Have Claude generate the data-filling code

2. **Save commonly used column sets**
   - Create preset shortcuts for your frequent chart types

3. **Batch create multiple charts**
   - Run the script multiple times with different outputs
   - Have Claude fill them all in one session

4. **Token saving strategy**
   - Template creation: 0 tokens (runs independently)
   - Data extraction from source: Ask Claude (minimal tokens)
   - Best of both worlds!

---

## File Locations

**Template Script:**
```
/Users/kimnguyen/Documents/Study Guide Claude Q3/Claude Templates/excel_master_chart_template.py
```

**Usage Guide (this file):**
```
/Users/kimnguyen/Documents/Study Guide Claude Q3/Claude Templates/EXCEL_MASTER_CHART_USAGE.md
```

**Template Instructions:**
```
/Users/kimnguyen/Documents/Study Guide Claude Q3/Claude Templates/Excel Master Chart Only.txt
```

---

## Quick Reference Commands

```bash
# Navigate to templates folder
cd "/Users/kimnguyen/Documents/Study Guide Claude Q3/Claude Templates"

# Interactive mode
python3 excel_master_chart_template.py

# Drug chart (all 11 columns)
python3 excel_master_chart_template.py \
  --source "source.txt" \
  --output "Drug_Chart.xlsx" \
  --columns "Drug Class,Drug Name (Brand),Route,Mechanism,Uses,Adverse Effects,Contraindications,Resistance,Drug Interactions,Drug Combinations,Special Considerations"

# Condition chart (7 columns)
python3 excel_master_chart_template.py \
  --source "source.txt" \
  --output "Condition_Chart.xlsx" \
  --columns "Condition,Epidemiology,Risk Factors,Clinical Presentation,Diagnostics,Labs,Treatment"

# Lab value chart (5 columns)
python3 excel_master_chart_template.py \
  --source "source.txt" \
  --output "Lab_Chart.xlsx" \
  --columns "Test Name,Normal Range,Increased In,Decreased In,Clinical Significance"
```

---

## Support

For issues or questions:
1. Check this usage guide first
2. Review the Excel Master Chart Only.txt template instructions
3. Ask Claude for help with data extraction or script customization

---

**Version:** 1.0
**Last Updated:** 2025-01-18
**Template Compliance:** Excel Master Chart Only.txt
