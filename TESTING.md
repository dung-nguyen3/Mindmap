# Testing Guide - Excel Master Chart Creator

## Pre-Testing Checklist

Before testing, ensure:
- [ ] Python 3.7+ installed: `python3 --version`
- [ ] openpyxl installed: `pip3 install openpyxl`
- [ ] tkinter available: `python3 -m tkinter` (should open a window)

## Test Suite

### Test 1: Application Launch ✅
**Objective**: Verify the application starts without errors

**Steps**:
1. Run: `python3 excel_master_chart_app.py`
2. Application window should open
3. Title should read "Excel Master Chart Creator"
4. All three sections should be visible

**Expected Result**: Clean launch with no errors

---

### Test 2: Preset Selection ✅
**Objective**: Verify all presets load correctly

**Steps**:
1. Launch app
2. Select "Drug Chart (11 columns)" from dropdown
3. Verify "Current Columns" displays 11 column names
4. Select "Condition Chart (7 columns)"
5. Verify "Current Columns" displays 7 column names
6. Select "Lab Values (5 columns)"
7. Verify "Current Columns" displays 5 column names

**Expected Result**: Column display updates correctly for each preset

---

### Test 3: Custom Columns ✅
**Objective**: Verify custom column definition works

**Steps**:
1. Launch app
2. Select "Custom" from dropdown
3. Dialog should appear
4. Enter test columns (one per line):
   ```
   Test Column 1
   Test Column 2
   Test Column 3
   ```
5. Click "Save"
6. Verify columns appear in "Current Columns"
7. Verify treeview headers match

**Expected Result**: Custom columns are created and displayed

---

### Test 4: Add/Delete Rows ✅
**Objective**: Verify row management works

**Steps**:
1. Launch app with any preset
2. Click "Add Row" 5 times
3. Verify row count shows "Rows: 5"
4. Verify 5 rows appear in grid (numbered 1-5)
5. Select row 3
6. Click "Delete Selected Row"
7. Verify row count shows "Rows: 4"
8. Verify remaining rows are renumbered 1-4

**Expected Result**: Rows add and delete correctly with proper numbering

---

### Test 5: Cell Editing ✅
**Objective**: Verify data entry works

**Steps**:
1. Launch app, add a row
2. Double-click first cell
3. Entry box should appear
4. Type "Test Data"
5. Press Return
6. Verify cell shows "Test Data"
7. Double-click again
8. Type new data
9. Press Escape
10. Verify cell reverts to previous value

**Expected Result**: Edit/save/cancel all work correctly

---

### Test 6: Load Example Data ✅
**Objective**: Verify JSON loading works

**Steps**:
1. Launch app
2. Click "Load Data (JSON)"
3. Navigate to `example_data/sample_drug_chart.json`
4. Click "Open"
5. Verify preset changes to "Drug Chart (11 columns)"
6. Verify row count shows "Rows: 8"
7. Verify data appears in grid

**Expected Result**: JSON data loads completely and correctly

---

### Test 7: Color Preview ✅
**Objective**: Verify color assignment logic

**Steps**:
1. Load `example_data/sample_drug_chart.json`
2. Click "Preview Colors"
3. Preview window should open
4. Verify color assignments:
   - "Beta Blockers" → one color
   - "ACE Inhibitors" → different color
   - "Statins" → different color
   - "Diuretics (Loop)" → different color

**Expected Result**: Each drug class gets unique color

---

### Test 8: Save Data ✅
**Objective**: Verify JSON saving works

**Steps**:
1. Launch app
2. Add some test data (3-5 rows with content)
3. Click "Save Data (JSON)"
4. Save as "test_save.json"
5. Success message appears
6. Clear all data
7. Load "test_save.json"
8. Verify data matches what was saved

**Expected Result**: Data saves and loads identically

---

### Test 9: Excel Export (Core Feature) ✅
**Objective**: Verify Excel generation with all formatting

**Steps**:
1. Load `example_data/sample_drug_chart.json`
2. In "Output Filename", enter: "Test_Drug_Chart"
3. Select Desktop as save location
4. Click "🚀 Export to Excel"
5. Success message appears
6. Excel file opens automatically
7. Manually verify Excel file:
   - [ ] Header row: Dark blue background (#4472C4)
   - [ ] Header row: White bold text, centered
   - [ ] Header row: Frozen (scrolls with data staying visible)
   - [ ] Data rows: Different colors for different drug classes
   - [ ] First column: Bold text
   - [ ] Other columns: Regular text
   - [ ] All cells: Word wrap enabled
   - [ ] Column widths: Appropriate sizes
   - [ ] Colors match preview from Test 7

**Expected Result**: Excel file perfectly formatted per specifications

---

### Test 10: Excel Export - Empty Data ✅
**Objective**: Verify error handling for empty export

**Steps**:
1. Launch fresh app
2. Don't add any data
3. Click "Export to Excel"
4. Warning message should appear: "Please add some data before exporting"

**Expected Result**: User-friendly error message, no crash

---

### Test 11: Excel Export - No Filename ✅
**Objective**: Verify filename validation

**Steps**:
1. Launch app, add a row
2. Clear the "Output Filename" field
3. Click "Export to Excel"
4. Warning message should appear: "Please enter an output filename"

**Expected Result**: Validation prevents empty filename

---

### Test 12: Clear All Data ✅
**Objective**: Verify clear function works

**Steps**:
1. Load example data
2. Click "Clear All Data"
3. Confirmation dialog appears
4. Click "Yes"
5. Verify all rows deleted
6. Verify row count shows "Rows: 0"

**Expected Result**: All data cleared after confirmation

---

### Test 13: Large Dataset ✅
**Objective**: Verify app handles 50+ rows

**Steps**:
1. Launch app
2. Click "Add Row" 60 times
3. Verify scrolling works
4. Add data to several cells
5. Scroll through data
6. Export to Excel
7. Open Excel, verify all rows present

**Expected Result**: Smooth performance with large datasets

---

### Test 14: Auto-Color Algorithm ✅
**Objective**: Verify color rotation and assignment

**Steps**:
1. Create data with 12 different groups (more than 10 colors available)
2. Preview colors
3. Verify colors rotate (groups 11 and 12 reuse colors 1 and 2)
4. Export to Excel
5. Verify Excel matches preview

**Expected Result**: Colors rotate through 10-color palette

---

### Test 15: macOS Integration ✅
**Objective**: Verify macOS-specific features

**Steps**:
1. Export a file
2. Verify file opens with default app (Excel or Numbers)
3. Verify file is saved to correct location
4. Verify Desktop path default works

**Expected Result**: Seamless macOS integration

---

## Manual Excel Validation Checklist

After exporting, manually check Excel file:

### Header Formatting
- [ ] Background color: #4472C4 (dark blue)
- [ ] Font: Calibri, Bold, Size 12
- [ ] Font color: White
- [ ] Alignment: Center horizontal, center vertical
- [ ] Word wrap: Enabled
- [ ] Row height: ~25 pixels
- [ ] Frozen panes: Header stays visible when scrolling down

### Data Row Formatting
- [ ] Font: Calibri, Size 10, Black
- [ ] First column: Bold
- [ ] Other columns: Regular weight
- [ ] Alignment: Left horizontal, top vertical
- [ ] Word wrap: Enabled
- [ ] Background: Pastel color (changes by group)
- [ ] Borders: White (invisible)

### Color Assignment
- [ ] Same group = same color
- [ ] Different groups = different colors
- [ ] Colors from correct palette (10 pastel colors)
- [ ] Color changes occur at correct boundaries

### Column Widths
- [ ] Short columns (Route): ~12 width
- [ ] Medium columns (Drug Class): ~22 width
- [ ] Name columns: ~28 width
- [ ] Long columns (Mechanism, Uses): ~35 width

---

## Performance Benchmarks

| Test | Target | Actual |
|------|--------|--------|
| App launch time | < 2 seconds | _____ |
| Add 50 rows | < 5 seconds | _____ |
| Export 50 rows to Excel | < 3 seconds | _____ |
| Load JSON (50 rows) | < 2 seconds | _____ |
| Color preview generation | < 1 second | _____ |

---

## Known Limitations

1. **Platform**: Optimized for macOS; Windows/Linux may have minor differences
2. **Excel required**: Auto-open requires Microsoft Excel or Numbers installed
3. **Memory**: Very large datasets (1000+ rows) not tested
4. **Tkinter**: UI appearance varies by macOS version and theme

---

## Regression Testing

After any code changes, re-run:
1. Test 9 (Excel Export - Core Feature)
2. Test 7 (Color Preview)
3. Test 14 (Auto-Color Algorithm)
4. Test 5 (Cell Editing)

These are the most critical features.

---

## Bug Reporting Template

If you find a bug, report with:
```
**Bug Description**: [Brief description]
**Steps to Reproduce**:
1. [Step 1]
2. [Step 2]
3. [Step 3]
**Expected Result**: [What should happen]
**Actual Result**: [What actually happened]
**Environment**:
- macOS version: [e.g., macOS 14.0]
- Python version: [e.g., Python 3.11.0]
- openpyxl version: [e.g., 3.1.2]
**Screenshots**: [If applicable]
```

---

## Test Status

| Test # | Test Name | Status | Notes |
|--------|-----------|--------|-------|
| 1 | Application Launch | ⬜ | |
| 2 | Preset Selection | ⬜ | |
| 3 | Custom Columns | ⬜ | |
| 4 | Add/Delete Rows | ⬜ | |
| 5 | Cell Editing | ⬜ | |
| 6 | Load Example Data | ⬜ | |
| 7 | Color Preview | ⬜ | |
| 8 | Save Data | ⬜ | |
| 9 | Excel Export | ⬜ | |
| 10 | Empty Data Handling | ⬜ | |
| 11 | Filename Validation | ⬜ | |
| 12 | Clear All Data | ⬜ | |
| 13 | Large Dataset | ⬜ | |
| 14 | Auto-Color Algorithm | ⬜ | |
| 15 | macOS Integration | ⬜ | |

**Legend**: ⬜ Not Started | 🟡 In Progress | ✅ Passed | ❌ Failed

---

**Happy Testing!** 🧪
