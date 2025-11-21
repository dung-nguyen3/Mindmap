# HTML Learning Objects Workspace

This workspace is for creating interactive HTML-based learning objects.

## Purpose
Create interactive educational content including:
- Interactive lessons and tutorials
- Quizzes and assessments
- Flashcards and study tools
- Concept visualizations
- Interactive diagrams

## Structure

### Directories
- `lessons/` - Complete lesson modules
- `quizzes/` - Interactive quiz files
- `flashcards/` - Flashcard applications
- `visualizations/` - Interactive diagrams and charts
- `templates/` - Reusable HTML templates
- `assets/` - CSS, JavaScript, images

### Technology Stack
- HTML5 for structure
- CSS3 for styling
- Vanilla JavaScript for interactivity
- No build tools required - all self-contained

## Features

### Interactive Elements
- Multiple choice quizzes with instant feedback
- Drag-and-drop activities
- Fill-in-the-blank exercises
- Interactive timelines
- Concept maps
- Flashcard decks with spaced repetition

### Design Principles
- Mobile-responsive
- Accessible (WCAG compliant)
- Works offline (single-file HTML)
- Print-friendly study guides
- Clean, distraction-free design

## Usage Examples

### Create Interactive Quiz
```bash
# Claude will generate complete HTML file
Request: "Create a quiz on cell biology with 10 questions"
Output: quizzes/cell_biology_quiz.html
```

### Generate Flashcards
```bash
Request: "Create flashcards from my notes in notes.txt"
Output: flashcards/[topic]_flashcards.html
```

### Build Interactive Lesson
```bash
Request: "Create an interactive lesson on photosynthesis"
Output: lessons/photosynthesis_lesson.html
```

## Opening Learning Objects
All HTML files can be opened directly in any web browser:
```bash
open lessons/lesson_name.html
```

## Notes
- All files are self-contained (CSS/JS embedded)
- Can be shared as single files
- Work offline
- Compatible with LMS platforms (SCORM export possible)
- Can embed videos, images, diagrams

## General Working Guidelines
- **IMPORTANT:** Only use information from the source file that was directly requested
- **MANDATORY EXCEPTION - MUST USE WebSearch:** When creating study materials, you MUST actively research and include:
  - **Mnemonics/memory tricks** - Use WebSearch to find PROVEN, ESTABLISHED mnemonics (do NOT invent them)
  - **Analogies** - Use WebSearch to research GOOD analogies that explain mechanisms and complex concepts
  - Proven memory techniques for the specific topic
  - Clinical associations and "if X think Y" patterns
  - **CRITICAL: You cannot skip this step - WebSearch is REQUIRED for mnemonics and analogies**

## Study Guide Creation Protocol (MANDATORY)

**When given a task to create (Word, Excel, HTML, any template), Claude MUST FIRST complete and state the verification checklist BEFORE creating any files:**

```
VERIFICATION CHECKLIST:
☐ Source file: [exact filename and path]
☐ Instruction template: [template name]
☐ Source-only policy: I will ONLY use information from [source file]. Exception: Memory tricks/mnemonics will be researched.
☐ MANDATORY: I will use WebSearch to research mnemonics and analogies - I will NOT invent them.
☐ Save location: [Class]/[Exam]/Claude Study Tools/
```

**This verification is REQUIRED before beginning work on any study guide creation task. State this checklist FIRST, then begin work. Do not skip this step.**

**IMPORTANT: This verification process applies to EACH new study guide creation. If user says "do this for file 35" or "create study guide for lecture 38", Claude must complete the full verification checklist again for that new file before beginning work. Each file = new verification process.**

## Post-Creation Verification Protocol (MANDATORY)

**After creating ANY study guide (Word, Excel, HTML), Claude MUST automatically re-analyze the completed document for:**

1. **Source Accuracy:**
   - All information comes from source file only (except mnemonics)
   - No external medical facts added without asterisk (*)
   - All page references included where required 

2. **Template Compliance:**
   - Followed ALL instructions from the template file exactly
   - Correct structure (all required tabs/sections present)
   - Correct formatting (colors, fonts, styles match template)
   - All required elements included (tables, boxes, mnemonics, analogies, etc.)

3. **Completeness:**
   - All learning objectives answered completely (all parts)
   - All comparison tables created as specified
   - Master chart includes ALL Drugs topics from source
   - No missing content from source file

4. **Quality Checks:**
   - No incorrect groupings or incorrect merged cells
   - No spelling errors
   - Proper formatting throughout as instructed in template instructions
   - TodoWrite tool used to track progress 

**CRITICAL: After document creation, Claude must state: "Post-creation verification complete" and report any issues found. Fix issues immediately if found.**

## File Source Instructions
- Always verify information comes from the specified source file
- Do not add external knowledge unless explicitly instructed
- **EXCEPTION FOR CRITICAL INFORMATION:** If you believe information is VERY important to include that is not in the source file, you MAY include it BUT you MUST mark it with an asterisk (*) so the user knows it's external knowledge
- **Format for external critical info:** Add * next to the information (e.g., "No bone marrow suppression*" or "Contraindicated in pregnancy*")
- Try to minimize external additions - only include if truly critical/must-know
- Cite specific page numbers or sections when referencing source material
- If source file doesn't contain requested information, inform user rather than filling in gaps
- Always ask user which file you are going to use unless user already told you the exact file
- Before using a different file from what the user instructed, always ask if you are allowed to first
- Before researching web, always ask first UNLESS it is for memory tricks/mnemonics/analogies (which is MANDATORY - do not ask)
- If user did not give you the file, paste the file, or copy the file path to use, always ask user
- Do not search other files in workspace unless you ask or unless user instructed

## File Organization
- Always save output files in the "Claude Study Tools" folder within the class directory
- Example: If source is in Pharmacology/Exam 3/Extract/, save output in Pharmacology/Exam 3/Claude Study Tools/
- Create "Claude Study Tools" folder if it doesn't exist
- Keep all generated study materials (Excel, Word, etc.) in Claude Study Tools folder

## Learning Objectives & Data Accuracy
- When a learning objective has multiple parts, answer ALL parts of that learning objective
- Example: "Pelvic exam: describe exam components. Describe normal and some abnormal findings. Demonstrate proper technique" = Answer ALL 3 parts
- After completing answers for each learning objective, VERIFY you answered all parts
- Before grouping items together (drugs, conditions, findings), VERIFY they share the same information
- Do NOT group drug classes or drug names together that don't belong
- Do NOT group clinical findings together unless they have identical information
- Always double-check before merging cells or creating combined categories

## Key Comparison Charts Guidelines
- Create additional key comparison charts for information with similar categories that need differentiation
- Look for patterns across topics that students need to compare and contrast
- Examples:
  - For drugs: Compare mechanisms of action (all drugs have different MOAs - help differentiate)
  - For clinical exams: Compare techniques (e.g., pelvic exam technique vs rectal exam technique)
  - For conditions: Compare diagnostic criteria, symptoms, treatments when similar
- Create comparison charts for any category that appears across multiple topics/drugs/conditions
- Goal: Help students differentiate between similar concepts with side-by-side comparisons
- **REQUIRED for Drug Mechanism comparisons:** Add "Analogy" column/row with 1-2 sentence analogy explanation for EACH drug's mechanism to help students understand how it works (e.g., "Like cutting the assembly line belt"). Research good analogies using web search.

## File Accuracy Analysis Protocol (MANDATORY)

**When asked to analyze an existing study guide file for accuracy, follow these steps:**

### Step 1: Read Source File
- Read the ENTIRE source file thoroughly
- Identify all drugs/topics/conditions mentioned
- Note all learning objectives
- Document all specific information for each item

### Step 2: Load and Examine Existing File
- Load the Excel/Word/HTML file
- List all sheets/sections
- Print out drug/topic lists and their classifications

### Step 3: Systematic Verification Checks

**Check 1: Drug/Topic Names**
- Verify ALL names match source EXACTLY (including spelling, capitalization, brand names)
- Check for any names not in source (external additions)
- List: All items present ✓ or missing ✗

**Check 2: Drug/Topic Classification**
- Verify each item is assigned to correct class/category
- Check source for class assignments
- Confirm: All classifications accurate ✓

**Check 3: Merged Cells and Groupings**
- Identify all merged cells/grouped items
- For EACH merged cell, verify in source that items truly share IDENTICAL information
- Check: Do these drugs/items actually belong together? Are they grouped in source?
- Common issues:
  - Different drugs merged with different MOAs
  - Different adverse effects incorrectly grouped
  - Items with different information merged together

**Check 4: Information Accuracy (Drug-Specific)**
- For EACH drug/item, verify:
  - Does this MOA belong to THIS specific drug or is it class-wide?
  - Are these adverse effects for THIS drug specifically?
  - Is THIS drug marked first-line (or just its class)?
  - Are contraindications specific to THIS drug?
  - Is THIS drug in this combination (source says so)?
- Check for external information not in source
- **Verify external info is marked with asterisk (*):** If any critical information was added that's not in source, it MUST have * next to it

**Check 5: Format and Colors**
- Verify format follows instruction template specifications:
  - Correct color scheme (soft pastels with black text)
  - Correct hex codes used
  - Proper sheet structure (5 tabs for Complete, 4 tabs for Drug Chart Excel)
  - Column widths appropriate
  - Row heights fit content
  - All data cells have pastel backgrounds (not just first column)
  - Text wrapping enabled
  - Frozen headers where specified
  - **Memory tricks/mnemonics row present after EACH drug class table** (Drug Chart Excel)
    - Merged cells across all columns
    - Light blue background (#E6F3FF)
    - Label: "💡 MEMORY TRICKS & MNEMONICS"
    - Contains researched mnemonics (not invented)
    - 2 blank rows after mnemonic before next drug class

**Check 6: Emojis and Visual Indicators**
- Verify ONLY emojis from instruction template are used
- Check: 🟢 (first-line), ⚠️ (warnings), 🚫 (contraindications), ✅, ❌, 💊, 🧪
- No extra emojis added that aren't in instructions

### Step 4: Document All Issues Found
Create summary:
- Issue 1: [Description] - [Source says] vs [File shows] - ❌ INCORRECT
- Issue 2: [Description] - [Source says] vs [File shows] - ❌ INCORRECT
- ✅ CORRECT: [What's accurate]

### Step 5: Fix Issues
- Fix each issue identified
- Update cells with correct information from source only
- Correct any spelling to match source EXACTLY

### Step 6: Re-Verify After Fixes
- Re-analyze entire file
- Verify all drug groupings are correct
- Confirm all information matches source
- Check format/colors match instruction template
- Create final verification checklist:
  - ✅ All names match source exactly
  - ✅ All classes correct
  - ✅ No incorrect groupings
  - ✅ No external information
  - ✅ Format matches instruction template
  - ✅ Colors correct per instructions

### Step 7: Use TodoWrite Tool
- Track each verification step as a todo
- Mark completed as you go
- Keep user informed of progress

**CRITICAL:** Always perform Step 6 (Re-Verify After Fixes) - never skip the final verification check.
