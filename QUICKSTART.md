# Quick Start Guide - Excel Master Chart Creator v2.5

## 5-Minute Setup

### 1. Install Dependencies (30 seconds)
```bash
pip3 install -r requirements.txt
```

Or manually:
```bash
pip3 install openpyxl tksheet
```

### 2. Run the App (10 seconds)
```bash
python3 excel_master_chart_app.py
```

You'll see an Excel-like interface with 50 pre-populated empty rows ready for data entry!

### 3. Try the Example Data (2 minutes)
1. **File → Open** (or Cmd+O)
2. Navigate to `example_data/sample_drug_chart.json`
3. Click **"Open"** - data loads into the grid automatically
4. **View → Preview Colors** to see the auto-color assignments
5. Select **"Comprehensive"** export format for 4-tab workbook
6. Click **"🚀 Export to Excel"**
7. Excel opens automatically with your beautifully formatted chart!

### 4. Create Your Own Chart (2 minutes)
1. Select a preset from the dropdown (e.g., "Condition Chart")
2. Click **"+50"** to add 50 rows instantly
3. Click a cell and start typing (or double-click to edit)
4. Press **Enter** to move down, **Tab** to move right
5. Use **right-click** for Fill Down, Insert Row, etc.
6. **File → Save** (Cmd+S) to save your work
7. Export!

---

## Common Questions

**Q: How do I edit a cell?**
A: Click once to select, then click again or press **Enter** to start editing. Or **double-click** directly. Press **Return** to save.

**Q: What's the difference between Master Chart and Comprehensive export?**
A: **Master Chart** is a single sheet (great for printing). **Comprehensive** is a 4-tab workbook with transposed drug tables, master list, quick reference, and alphabetical index (great for comprehensive study).

**Q: How does auto-coloring work?**
A: Colors use a **3-shade gradient system**. The app detects when the first column value changes and assigns a new color set. Each set has header (darker), main (medium), and row label (lighter) shades.

**Q: What's the right-click menu for?**
A: Quick access to powerful features like Insert Row, Delete Row, Fill Down (copy cell to rows below), Sort, Copy/Paste, and more!

**Q: Do I need to save manually?**
A: No! Auto-save runs every 2 minutes. But you can use **Cmd+S** anytime for peace of mind.

**Q: Can I add more than 11 columns?**
A: Yes! Choose "Custom" preset and define as many columns as you want (one per line in the dialog).

**Q: Where is my Excel file saved?**
A: By default, your Desktop. You can change this before exporting.

**Q: What if I accidentally close the app?**
A: No problem! The app will detect the auto-save file and offer to restore your work on next launch (crash recovery).

**Q: Can I use this on Windows/Linux?**
A: The app should work, but it's optimized for macOS. File auto-opening and some shortcuts may behave differently.

---

## Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| **File Operations** | |
| New | Cmd+N |
| Open | Cmd+O |
| Save | Cmd+S |
| Save As | Cmd+Shift+S |
| **Editing** | |
| Undo | Cmd+Z |
| Redo | Cmd+Shift+Z |
| Copy | Cmd+C |
| Paste | Cmd+V |
| Cut | Cmd+X |
| Select All | Cmd+A |
| **Grid Navigation** | |
| Edit cell | Click, then Enter (or double-click) |
| Move down | Return/Enter |
| Move right | Tab |
| Move left | Shift+Tab |
| Navigate | Arrow keys |
| Cancel edit | Escape |
| **View** | |
| Zoom in | Cmd+Plus |
| Zoom out | Cmd+Minus |
| Reset zoom | Cmd+0 |

---

## Tips for Best Results

1. **Keep first column consistent**: If using "Beta Blockers", don't also use "Beta-Blockers" or "beta blockers" - colors won't group correctly (case-sensitive!)
2. **Use Quick Add buttons**: Click "+50" instead of "Add Row" 50 times
3. **Master the right-click menu**: Fill Down, Insert Row, Sort - these save tons of time
4. **Let auto-save work**: It runs every 2 minutes automatically, but Cmd+S if you want
5. **Use Preview Colors**: View → Preview Colors to check assignments before exporting
6. **Try both export formats**: Master Chart for printing, Comprehensive for study
7. **Delete empty rows before export**: One-click cleanup with "Delete Empty Rows" button
8. **Start small**: Test with 5-10 rows before adding 100+
9. **Learn keyboard shortcuts**: Tab, Enter, Cmd+C/V work in the grid just like Excel
10. **Enable Live Color Preview**: See colors in real-time as you type (View menu)

---

## Example Workflow: Creating a Drug Chart

```
Step 1: Launch app
  → python3 excel_master_chart_app.py
  → See Excel-like interface with 50 empty rows

Step 2: Select preset
  → Choose "Drug Chart (11 columns)" from dropdown
  → Grid updates with 11 column headers

Step 3: Add more rows if needed
  → Click "+50" for 50 more rows (now have 100 total)
  → Or use custom quantity: Type "20" → "Add Rows"

Step 4: Enter data
  → Click first cell (Drug Class column, row 1)
  → Type "Beta Blockers" → Press Enter (moves down)
  → Or Tab to move right
  → Fill in: Metoprolol, Oral, Mechanism, Uses, etc.
  → Right-click → "Fill Down" to copy drug class to multiple drugs
  → Repeat for other classes: ACE Inhibitors, Statins, etc.

Step 5: Use power features
  → Right-click a column header → Sort alphabetically
  → Cmd+C / Cmd+V to copy/paste cells
  → Delete Empty Rows button to clean up

Step 6: Save work
  → File → Save (Cmd+S)
  → Save as "My_Cardiology_Drugs.json"

Step 7: Preview
  → View → Preview Colors
  → Verify each drug class has unique 3-shade color set

Step 8: Export
  → Select "Comprehensive" format (4-tab workbook)
  → Enter filename: "Cardiology_Drug_Guide.xlsx"
  → Click "🚀 Export to Excel"
  → Excel opens with 4 tabs: Drug Details, Master List, Quick Reference, Index!
```

---

## Troubleshooting Quick Fixes

**App won't start**
```bash
# Check Python version
python3 --version  # Should be 3.7+

# Reinstall dependencies
pip3 install --upgrade openpyxl tksheet
```

**Error: `ModuleNotFoundError: No module named 'tksheet'`**
```bash
pip3 install tksheet
```

**Grid looks empty**
- This is normal! 50 empty rows are pre-populated
- Just click a cell and start typing
- Or load example data: File → Open → `example_data/sample_drug_chart.json`

**Can't edit cells**
- Click once to select, then click again or press Enter to edit
- Or **double-click** directly on the cell
- Try clicking outside first if it's not responding

**Auto-save not working**
- Check title bar for unsaved changes indicator (*)
- Auto-save runs every 2 minutes automatically
- Force save: File → Save (Cmd+S)

**Crash recovery dialog appears**
- Previous auto-save file detected
- Click "Yes" to restore your work, or "No" to start fresh

**Excel file is blank**
- Make sure you added data and cells aren't empty
- Use "Delete Empty Rows" to clean up before export

**Colors are wrong in Comprehensive format**
- Check that first column values are EXACTLY the same for items in the same group (case-sensitive!)
- Use "View → Preview Colors" to verify assignments before exporting
- Comprehensive format uses 3-shade gradient system

**Right-click menu doesn't appear**
- Click directly on a cell, not in the margin
- Make sure a cell is selected first
- Try restarting the app

---

## Next Steps

1. ✅ Try the example data in `example_data/` (File → Open)
2. ✅ Master the Excel-like interface (click, type, Enter/Tab)
3. ✅ Explore the right-click menu (Insert Row, Fill Down, Sort)
4. ✅ Try both export formats (Master Chart vs. Comprehensive)
5. ✅ Learn keyboard shortcuts (Cmd+S, Cmd+Z, Cmd+C/V)
6. ✅ Enable Live Color Preview (View menu)
7. ✅ Use Recent Files for quick access (File → Recent Files)
8. ✅ Customize 3-shade color sets (advanced - edit .py file)
9. ✅ Package as macOS .app (see README.md)

---

## What's New in v2.5

🎉 **Major upgrade from v1.0!**

- **Excel-like interface** with tksheet (no more limited Treeview)
- **Auto-save & crash recovery** (every 2 minutes)
- **Menu bar** with keyboard shortcuts (Cmd+S, Cmd+O, Cmd+Z, etc.)
- **Right-click context menu** (Insert/Delete rows, Fill Down, Sort, etc.)
- **Quick Add buttons** (+10, +50, +100 rows)
- **Live Color Preview** toggle (real-time row coloring)
- **Recent Files** menu (last 10 files)
- **Unsaved changes indicator** (*)
- **3-shade color system** (header/main/row_label gradients)
- **Two export formats**: Master Chart (1 sheet) and Comprehensive (4 tabs with transposed drug tables)
- **Delete Empty Rows** one-click cleanup

---

**Need more help?** See the full README.md or TROUBLESHOOTING.md for detailed documentation.

**Enjoy creating beautiful charts!** 🚀
