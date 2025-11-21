# Troubleshooting Guide

## Common Issues & Solutions

---

### 🚨 Application Won't Start

#### Error: `ModuleNotFoundError: No module named 'tkinter'`

**Cause**: tkinter not installed or not found

**Solution for macOS**:
```bash
# Install Python from python.org (includes tkinter)
# Or use Homebrew:
brew install python-tk
```

**Solution for Linux**:
```bash
sudo apt-get install python3-tk
```

---

#### Error: `ModuleNotFoundError: No module named 'openpyxl'`

**Cause**: openpyxl not installed

**Solution**:
```bash
pip3 install openpyxl

# If that doesn't work, try:
python3 -m pip install openpyxl

# Or with specific user install:
pip3 install --user openpyxl
```

**Verification**:
```bash
python3 -c "import openpyxl; print('✅ openpyxl installed!')"
```

---

#### Error: `command not found: python3`

**Cause**: Python not installed or not in PATH

**Solution**:
```bash
# Check if python (not python3) works:
python --version

# If it shows Python 3.x, use 'python' instead of 'python3'
python excel_master_chart_app.py

# Otherwise, install Python from python.org
```

---

### 🚨 GUI Issues

#### Application window is blank or doesn't display properly

**Cause**: tkinter display issues, theme conflicts, or macOS version compatibility

**Solutions**:
1. Update Python to latest version
2. Try different Python installation (python.org vs Homebrew)
3. Run with verbose output:
   ```bash
   python3 -v excel_master_chart_app.py
   ```

---

#### Can't resize window or window is too small

**Cause**: Hard-coded window size might not fit your screen

**Solution**: Edit `excel_master_chart_app.py`, line ~100:
```python
# Change from:
self.root.geometry("1200x700")
# To:
self.root.geometry("1400x800")  # Or your preferred size
```

---

#### Buttons don't respond to clicks

**Cause**: macOS accessibility permissions or tkinter event loop issue

**Solution**:
1. Grant Terminal/iTerm accessibility permissions:
   - System Preferences → Security & Privacy → Privacy → Accessibility
   - Add Terminal.app
2. Restart the application

---

### 🚨 Data Entry Issues

#### Can't edit cells (nothing happens on double-click)

**Causes**:
1. Single-clicking instead of double-clicking
2. Clicking row number column (#)
3. Event binding issue

**Solutions**:
1. Make sure to **double-click** directly on the cell (not single-click)
2. Don't click the # column (row numbers) - click data cells
3. Try clicking another cell first, then double-click target cell
4. Restart the application

---

#### Cell edit box appears in wrong position

**Cause**: Window resizing or high-DPI display scaling

**Solution**: Scroll the cell into view before editing, or restart app

---

#### Data disappears after editing

**Cause**: Pressing Escape or clicking outside cell before pressing Return

**Solution**:
- Press **Return/Enter** to save edits
- Don't click outside or press Escape unless you want to cancel

---

### 🚨 Excel Export Issues

#### Error: `Permission denied` when exporting

**Cause**: File is currently open in Excel

**Solution**:
1. Close the Excel file
2. Export again
3. Or use different filename

---

#### Excel file is created but appears blank

**Causes**:
1. No data was added to grid
2. Data rows were added but cells are empty

**Solutions**:
1. Verify row count shows "Rows: X" where X > 0
2. Double-click cells and add content
3. Try loading example data first: `example_data/sample_drug_chart.json`

---

#### Colors are all the same

**Cause**: First column (grouping column) has identical values for all rows

**Solution**:
1. Make sure first column has different values for different groups
   - ✅ Correct: "Beta Blockers", "ACE Inhibitors", "Statins"
   - ❌ Wrong: "Beta Blockers", "Beta Blockers", "Beta Blockers"
2. Use "Preview Colors" to verify color assignments before exporting

---

#### Excel file won't open automatically

**Causes**:
1. Microsoft Excel or Numbers not installed
2. File association not set
3. macOS permissions issue

**Solutions**:
1. File is still created - navigate to save location and open manually
2. Install Microsoft Excel or Apple Numbers
3. Check default app for .xlsx files:
   - Right-click any .xlsx file → Get Info → Open with → Excel or Numbers

---

#### Formatting is wrong (no colors, no frozen panes, etc.)

**Cause**: openpyxl version incompatibility or Excel version issue

**Solutions**:
```bash
# Update openpyxl:
pip3 install --upgrade openpyxl

# Check version (should be 3.0+):
python3 -c "import openpyxl; print(openpyxl.__version__)"
```

---

### 🚨 JSON Save/Load Issues

#### Error: `Invalid JSON` when loading

**Cause**: JSON file is corrupted or manually edited incorrectly

**Solution**:
1. Don't manually edit JSON files outside the app
2. Use a JSON validator: https://jsonlint.com
3. Re-save from app if possible

---

#### Loaded data doesn't match what was saved

**Cause**: Preset changed between save and load

**Solution**: JSON files include preset info - make sure to use compatible presets

---

#### Can't find saved JSON file

**Cause**: Saved to unexpected location

**Solution**:
```bash
# Search for your JSON files:
find ~/Desktop -name "*.json" -type f

# Or use Spotlight:
# Press Cmd+Space, search for filename
```

---

### 🚨 py2app Packaging Issues

#### Error: `No module named 'py2app'`

**Solution**:
```bash
pip3 install py2app
```

---

#### Error during build: `ModuleNotFoundError`

**Cause**: Dependencies not detected

**Solution**: Edit `setup.py`, add missing modules to packages list:
```python
OPTIONS = {
    'packages': ['tkinter', 'openpyxl', 'json', 'pathlib'],
    # ...
}
```

---

#### .app won't open: "damaged and can't be opened"

**Cause**: macOS Gatekeeper security

**Solution**:
```bash
# Remove quarantine attribute:
xattr -cr /Applications/excel_master_chart_app.app

# Or:
# Right-click app → Open (instead of double-click)
# Click "Open" in warning dialog
```

---

#### .app crashes immediately after opening

**Causes**:
1. Missing dependencies in .app bundle
2. Python version mismatch

**Solutions**:
1. Run from Terminal to see error:
   ```bash
   /Applications/excel_master_chart_app.app/Contents/MacOS/excel_master_chart_app
   ```
2. Rebuild with all dependencies:
   ```bash
   python3 setup.py py2app --packages=tkinter,openpyxl
   ```

---

### 🚨 Performance Issues

#### App is slow or unresponsive

**Causes**:
1. Too many rows (100+)
2. macOS resource constraints

**Solutions**:
1. Export in batches (50 rows max per file)
2. Clear unused data
3. Close other applications
4. Upgrade RAM if possible

---

#### Excel export takes a long time

**Cause**: Large dataset with complex formatting

**Solution**: Normal for 50+ rows with full formatting. Wait for completion (up to 30 seconds for 100 rows).

---

### 🚨 Platform-Specific Issues

#### macOS: Window has black background

**Cause**: Dark mode theme conflict

**Solution**: Edit code to set light theme:
```python
# Add after imports:
import tkinter as tk
from tkinter import ttk

# Add in __init__:
self.root.tk.call("tk", "scaling", 1.5)  # Adjust as needed
```

---

#### macOS Big Sur+: Buttons look different

**Cause**: macOS native rendering changes

**Solution**: This is normal - tkinter adapts to macOS theme

---

#### Windows: File paths with backslashes cause errors

**Cause**: Path separator differences

**Solution**: App uses `pathlib.Path` which handles this - update to latest code

---

### 🚨 Data Validation Issues

#### Excel formula error: `#VALUE!`

**Cause**: No formulas should be in this app - if you see this, data was incorrectly entered

**Solution**: Re-enter data without `=` at start of cell

---

#### Special characters (emojis) don't appear in Excel

**Cause**: Excel font doesn't support character

**Solution**: Excel issue, not app issue - try changing Excel font to "Apple Color Emoji"

---

### 🚨 Installation Issues

#### `pip3: command not found`

**Solution**:
```bash
# Install pip:
python3 -m ensurepip --default-pip

# Or download get-pip.py:
curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
python3 get-pip.py
```

---

#### Permission denied when installing with pip

**Solution**:
```bash
# Install for current user only:
pip3 install --user openpyxl

# Or use sudo (not recommended):
sudo pip3 install openpyxl
```

---

## Debugging Techniques

### Enable Verbose Output
```bash
python3 -v excel_master_chart_app.py 2>&1 | tee app_log.txt
```

### Check Python Environment
```bash
python3 --version
pip3 list
which python3
```

### Verify tkinter
```bash
python3 -m tkinter
# Should open a small test window
```

### Test openpyxl Independently
```bash
python3 << EOF
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws['A1'] = 'Test'
wb.save('test.xlsx')
print('✅ openpyxl works!')
EOF
```

---

## Getting Help

If none of these solutions work:

1. **Check Python/Package Versions**:
   ```bash
   python3 --version
   pip3 show openpyxl
   ```

2. **Run with Error Logging**:
   ```bash
   python3 excel_master_chart_app.py 2> error_log.txt
   ```

3. **Provide Error Details**:
   - Full error message
   - Python version
   - macOS version
   - Steps to reproduce

4. **Try Example Data First**:
   - Load `example_data/sample_drug_chart.json`
   - Export to verify basic functionality

---

## Quick Diagnostic Commands

Run these to gather diagnostic info:

```bash
# System info
uname -a
sw_vers  # macOS only

# Python info
python3 --version
which python3
pip3 --version

# Package verification
pip3 list | grep openpyxl
python3 -c "import tkinter; print('✅ tkinter')"
python3 -c "import openpyxl; print('✅ openpyxl')"

# File permissions
ls -la excel_master_chart_app.py

# Disk space
df -h ~
```

---

**Still stuck? Create an issue with diagnostic output!**
