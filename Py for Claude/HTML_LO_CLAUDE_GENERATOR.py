#!/usr/bin/env python3
"""
HTML LEARNING OBJECTIVES - CLAUDE GENERATOR TEMPLATE
=====================================================

This template is for CLAUDE to create complete ready-to-run HTML learning objectives.

INSTRUCTIONS FOR CLAUDE:
1. Copy this entire template
2. Fill in the SOURCE_INFO section
3. Fill in the DATA sections for all 4 tabs:
   - Tab 1: Learning Objectives (LO data)
   - Tab 2: Key Comparisons (comparison tables)
   - Tab 3: Master Comparison Tables (comprehensive tables)
   - Tab 4: Summary (mnemonics, pearls, if-then rules)
4. Update OUTPUT_FILENAME
5. Save as: [LectureNumber]_[Topic]_Learning_Objectives.py
6. User runs: python3 script.py (HTML opens automatically in browser)

VERIFICATION CHECKLIST (Claude must complete BEFORE creating):
☑ Source file: [exact path]
☑ Instruction template: HTML LO with Master Chart 10-30.txt
☑ Source-only policy: ONLY use information from source file
☑ MANDATORY: WebSearch used for mnemonics (not data extraction)
☑ Save location: [Class]/[Exam]/Claude Study Tools/

Features:
- ✅ 4-tab structure (LOs, Key Comparisons, Master Tables, Summary)
- ✅ Color-coded emergency levels (red/orange/white)
- ✅ Interactive navigation
- ✅ All data hard-coded (ready to run)
- ✅ Opens browser automatically
"""

import sys
import subprocess
import platform
from pathlib import Path

# =============================================================================
# SOURCE INFORMATION (Claude fills this in)
# =============================================================================

SOURCE_FILE = "path/to/source/file.txt"  # Claude: Update with actual source file path
LECTURE_TOPIC = "Topic Name"  # Claude: Update with topic (e.g., "Red Eye Differential")
CREATION_DATE = "2025-01-19"  # Claude: Update with creation date

# =============================================================================
# TAB 1: LEARNING OBJECTIVES DATA
# =============================================================================

# Format: List of dictionaries with 'number', 'objective', and 'answer'
LEARNING_OBJECTIVES = [
    # Example format - Claude replaces with real data:
    # {
    #     'number': 1,
    #     'objective': 'Describe the pathophysiology of conjunctivitis',
    #     'answer': '''
    #         <h4>Pathophysiology of Conjunctivitis</h4>
    #         <p>Inflammation of the conjunctiva caused by:</p>
    #         <ul>
    #             <li><strong>Viral:</strong> Adenovirus (most common), HSV, VZV</li>
    #             <li><strong>Bacterial:</strong> S. aureus, S. pneumoniae, H. influenzae</li>
    #             <li><strong>Allergic:</strong> IgE-mediated hypersensitivity</li>
    #         </ul>
    #         <p><span class="highlight-blue">Key mechanism:</span> Vasodilation and inflammatory cell infiltration</p>
    #     '''
    # },
]

# =============================================================================
# TAB 2: KEY COMPARISONS DATA
# =============================================================================

# Format: List of comparison tables
# Each table has: title, headers (list), rows (list of dicts), urgency_level
KEY_COMPARISONS = [
    # Example format - Claude replaces with real data:
    # {
    #     'title': 'Viral vs Bacterial vs Allergic Conjunctivitis',
    #     'headers': ['Feature', 'Viral', 'Bacterial', 'Allergic'],
    #     'rows': [
    #         {
    #             'urgency': 'routine',  # 'emergency', 'urgent', or 'routine'
    #             'cells': ['Discharge', 'Watery', '<strong>Purulent</strong>', 'Stringy mucoid']
    #         },
    #         {
    #             'urgency': 'routine',
    #             'cells': ['Itching', 'Minimal', 'Minimal', '<strong>Severe</strong>']
    #         },
    #         {
    #             'urgency': 'routine',
    #             'cells': ['Treatment', 'Supportive', 'Topical antibiotics', 'Antihistamines']
    #         }
    #     ]
    # },
]

# =============================================================================
# TAB 3: MASTER COMPARISON TABLES DATA
# =============================================================================

# Format: Comprehensive tables with ALL conditions
MASTER_TABLES = [
    # Example format - Claude replaces with real data:
    # {
    #     'title': 'Complete Red Eye Differential Diagnosis',
    #     'description': 'Comprehensive comparison of all red eye conditions from lecture',
    #     'headers': ['Condition', 'Vision', 'Pain', 'Discharge', 'Pupil', 'Key Features', 'Management'],
    #     'rows': [
    #         {
    #             'urgency': 'emergency',  # Red background
    #             'cells': [
    #                 '<strong>Acute Angle-Closure Glaucoma</strong>',
    #                 '<span class="highlight-red">Severely decreased</span>',
    #                 '<span class="highlight-red">Severe</span>',
    #                 'None',
    #                 'Mid-dilated, non-reactive',
    #                 'IOP >40, corneal edema, nausea/vomiting',
    #                 '<span class="highlight-red">EMERGENCY:</span> Immediate ophthalmology'
    #             ]
    #         },
    #         {
    #             'urgency': 'urgent',  # Orange background
    #             'cells': [
    #                 '<strong>Corneal Ulcer</strong>',
    #                 'Decreased',
    #                 'Severe',
    #                 'Purulent',
    #                 'Normal or miotic',
    #                 'White infiltrate, hypopyon possible',
    #                 '<span class="highlight-yellow">URGENT:</span> Ophthalmology within 24h'
    #             ]
    #         },
    #         {
    #             'urgency': 'routine',  # White background
    #             'cells': [
    #                 '<strong>Viral Conjunctivitis</strong>',
    #                 'Normal',
    #                 'Minimal',
    #                 'Watery',
    #                 'Normal',
    #                 'Preauricular LAD, follicles',
    #                 'Supportive care, highly contagious'
    #             ]
    #         }
    #     ]
    # },
]

# =============================================================================
# TAB 4: SUMMARY DATA
# =============================================================================

SUMMARY_DATA = {
    # One-sentence summaries
    'key_points': [
        # Example: 'Acute angle-closure glaucoma = mid-dilated fixed pupil + IOP >40 + severe pain → EMERGENCY',
        # Example: 'Viral conjunctivitis = watery discharge + preauricular LAD → supportive care',
    ],

    # Mnemonics (researched via WebSearch - NOT invented!)
    'mnemonics': [
        # Example:
        # {
        #     'topic': 'Corneal Ulcer Red Flags',
        #     'mnemonic': 'FEVER',
        #     'breakdown': {
        #         'F': 'Fluorescein uptake',
        #         'E': 'Eye pain severe',
        #         'V': 'Vision decreased',
        #         'E': 'Epithelial defect',
        #         'R': 'Referral urgent'
        #     }
        # },
    ],

    # If-Then clinical rules
    'if_then_rules': [
        # Example: {'if': 'Mid-dilated fixed pupil + severe eye pain', 'then': 'Think acute angle-closure glaucoma → EMERGENCY'},
        # Example: {'if': 'Dendritic lesion on fluorescein', 'then': 'Think HSV keratitis → start antiviral'},
    ],

    # Critical values
    'critical_values': [
        # Example: {'parameter': 'Normal IOP', 'value': '10-21 mmHg'},
        # Example: {'parameter': 'Emergency IOP', 'value': '>40 mmHg (acute angle-closure)'},
    ],

    # Key definitions
    'definitions': [
        # Example: {'term': 'Hyphema', 'definition': 'Blood in anterior chamber'},
        # Example: {'term': 'Hypopyon', 'definition': 'Pus in anterior chamber'},
    ],

    # High-yield pearls
    'pearls': [
        # Example: '⚠️ Never patch a contact lens wearer with corneal abrasion (risk of Pseudomonas)',
        # Example: '✅ Viral conjunctivitis is self-limited but highly contagious for 10-12 days',
    ]
}

# Output filename - Claude updates based on topic
OUTPUT_FILENAME = 'Learning_Objectives.html'  # e.g., 'Red_Eye_Learning_Objectives.html'

# =============================================================================
# HTML GENERATION FUNCTIONS
# =============================================================================

def generate_html():
    """Generate complete HTML file with all 4 tabs."""

    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{LECTURE_TOPIC} - Learning Objectives</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            background: #f5f5f5;
        }}

        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 2rem;
            text-align: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}

        .header h1 {{
            font-size: 2rem;
            margin-bottom: 0.5rem;
        }}

        .header .meta {{
            opacity: 0.9;
            font-size: 0.9rem;
        }}

        .tabs {{
            display: flex;
            background: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            position: sticky;
            top: 0;
            z-index: 100;
        }}

        .tab-button {{
            flex: 1;
            padding: 1rem;
            background: white;
            border: none;
            border-bottom: 3px solid transparent;
            cursor: pointer;
            font-size: 1rem;
            font-weight: 500;
            transition: all 0.3s;
        }}

        .tab-button:hover {{
            background: #f5f5f5;
        }}

        .tab-button.active {{
            border-bottom-color: #667eea;
            color: #667eea;
            background: #f9fafb;
        }}

        .container {{
            max-width: 1400px;
            margin: 0 auto;
            padding: 2rem;
        }}

        .tab-content {{
            display: none;
            background: white;
            border-radius: 8px;
            padding: 2rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}

        .tab-content.active {{
            display: block;
            animation: fadeIn 0.3s;
        }}

        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        .lo-item {{
            margin-bottom: 3rem;
            padding-bottom: 2rem;
            border-bottom: 2px solid #e5e7eb;
        }}

        .lo-item:last-child {{
            border-bottom: none;
        }}

        .lo-number {{
            display: inline-block;
            background: #667eea;
            color: white;
            padding: 0.25rem 0.75rem;
            border-radius: 4px;
            font-weight: bold;
            margin-right: 0.5rem;
        }}

        .lo-objective {{
            font-size: 1.25rem;
            font-weight: 600;
            color: #1f2937;
            margin-bottom: 1rem;
        }}

        .lo-answer {{
            color: #4b5563;
            padding-left: 1rem;
        }}

        .comparison-table-section {{
            margin-bottom: 3rem;
        }}

        .comparison-table-section:last-child {{
            margin-bottom: 0;
        }}

        .comparison-title {{
            font-size: 1.5rem;
            font-weight: 600;
            color: #1f2937;
            margin-bottom: 1rem;
            margin-top: 2.5rem;
        }}

        .comparison-title:first-child {{
            margin-top: 0;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 1rem 0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }}

        th {{
            background: #f3f4f6;
            padding: 1rem;
            text-align: left;
            font-weight: 600;
            border: 1px solid #e5e7eb;
        }}

        td {{
            padding: 1rem;
            border: 1px solid #e5e7eb;
            vertical-align: top;
        }}

        tr.emergency {{
            background: #ffebee;
        }}

        tr.urgent {{
            background: #fff3e0;
        }}

        tr.routine {{
            background: white;
        }}

        .highlight-red {{
            background: #fee;
            color: #c00;
            padding: 2px 6px;
            border-radius: 3px;
            font-weight: 600;
        }}

        .highlight-yellow {{
            background: #fff4e6;
            color: #d97706;
            padding: 2px 6px;
            border-radius: 3px;
            font-weight: 600;
        }}

        .highlight-purple {{
            background: #f3e8ff;
            color: #7c3aed;
            padding: 2px 6px;
            border-radius: 3px;
            font-weight: 600;
        }}

        .highlight-blue {{
            background: #dbeafe;
            color: #1e40af;
            padding: 2px 6px;
            border-radius: 3px;
            font-weight: 600;
        }}

        .highlight-green {{
            background: #dcfce7;
            color: #15803d;
            padding: 2px 6px;
            border-radius: 3px;
            font-weight: 600;
        }}

        .summary-section {{
            margin-bottom: 2.5rem;
        }}

        .summary-section h3 {{
            color: #1f2937;
            font-size: 1.5rem;
            margin-bottom: 1rem;
            padding-bottom: 0.5rem;
            border-bottom: 2px solid #667eea;
        }}

        .key-point {{
            background: #f9fafb;
            padding: 1rem;
            margin: 0.5rem 0;
            border-left: 4px solid #667eea;
            border-radius: 4px;
        }}

        .mnemonic {{
            background: #fff7ed;
            padding: 1.5rem;
            margin: 1rem 0;
            border-radius: 8px;
            border: 2px solid #fb923c;
        }}

        .mnemonic-title {{
            font-weight: 700;
            color: #ea580c;
            font-size: 1.2rem;
            margin-bottom: 0.5rem;
        }}

        .mnemonic-letters {{
            font-size: 1.5rem;
            font-weight: bold;
            color: #c2410c;
            margin: 0.5rem 0;
            letter-spacing: 2px;
        }}

        .mnemonic-breakdown {{
            margin-top: 0.75rem;
        }}

        .mnemonic-breakdown div {{
            margin: 0.25rem 0;
            padding-left: 1rem;
        }}

        .if-then {{
            background: #f0fdf4;
            padding: 1rem;
            margin: 0.5rem 0;
            border-radius: 6px;
            border-left: 4px solid #22c55e;
        }}

        .if-then-if {{
            font-weight: 600;
            color: #15803d;
        }}

        .if-then-then {{
            color: #166534;
            margin-top: 0.25rem;
        }}

        .definition {{
            padding: 0.75rem;
            margin: 0.5rem 0;
            background: #faf5ff;
            border-radius: 4px;
        }}

        .definition-term {{
            font-weight: 700;
            color: #7c3aed;
        }}

        .pearl {{
            background: #fef3c7;
            padding: 1rem;
            margin: 0.5rem 0;
            border-radius: 6px;
            border-left: 4px solid #f59e0b;
        }}

        .master-description {{
            background: #eff6ff;
            padding: 1rem;
            border-radius: 6px;
            margin-bottom: 1.5rem;
            border-left: 4px solid #3b82f6;
            color: #1e40af;
        }}

        ul {{
            margin-left: 1.5rem;
            margin-top: 0.5rem;
            margin-bottom: 0.5rem;
        }}

        li {{
            margin: 0.25rem 0;
        }}

        @media print {{
            .tabs {{ display: none; }}
            .tab-content {{ display: block !important; page-break-before: always; }}
            .header {{ break-after: avoid; }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>{LECTURE_TOPIC}</h1>
        <div class="meta">
            Learning Objectives | Source: {SOURCE_FILE} | Generated: {CREATION_DATE}
        </div>
    </div>

    <div class="tabs">
        <button class="tab-button active" onclick="switchTab(0)">Learning Objectives</button>
        <button class="tab-button" onclick="switchTab(1)">Key Comparisons</button>
        <button class="tab-button" onclick="switchTab(2)">Master Tables</button>
        <button class="tab-button" onclick="switchTab(3)">Summary</button>
    </div>

    <div class="container">
        <!-- Tab 1: Learning Objectives -->
        <div class="tab-content active" id="tab-0">
            {generate_tab1_content()}
        </div>

        <!-- Tab 2: Key Comparisons -->
        <div class="tab-content" id="tab-1">
            {generate_tab2_content()}
        </div>

        <!-- Tab 3: Master Comparison Tables -->
        <div class="tab-content" id="tab-2">
            {generate_tab3_content()}
        </div>

        <!-- Tab 4: Summary -->
        <div class="tab-content" id="tab-3">
            {generate_tab4_content()}
        </div>
    </div>

    <script>
        function switchTab(index) {{
            // Hide all tabs
            document.querySelectorAll('.tab-content').forEach(tab => {{
                tab.classList.remove('active');
            }});
            document.querySelectorAll('.tab-button').forEach(btn => {{
                btn.classList.remove('active');
            }});

            // Show selected tab
            document.getElementById('tab-' + index).classList.add('active');
            document.querySelectorAll('.tab-button')[index].classList.add('active');

            // Scroll to top
            window.scrollTo({{ top: 0, behavior: 'smooth' }});
        }}
    </script>
</body>
</html>'''

    return html


def generate_tab1_content():
    """Generate Tab 1: Learning Objectives content."""
    if not LEARNING_OBJECTIVES:
        return '<p style="color: #9ca3af;">No learning objectives provided.</p>'

    content = ''
    for lo in LEARNING_OBJECTIVES:
        content += f'''
        <div class="lo-item">
            <div class="lo-objective">
                <span class="lo-number">LO {lo['number']}</span>
                {lo['objective']}
            </div>
            <div class="lo-answer">
                {lo['answer']}
            </div>
        </div>
        '''

    return content


def generate_tab2_content():
    """Generate Tab 2: Key Comparisons content."""
    if not KEY_COMPARISONS:
        return '<p style="color: #9ca3af;">No key comparisons provided.</p>'

    content = ''
    for comparison in KEY_COMPARISONS:
        content += f'<h2 class="comparison-title">{comparison["title"]}</h2>\n'
        content += '<table>\n<thead><tr>\n'

        # Headers
        for header in comparison['headers']:
            content += f'<th>{header}</th>\n'
        content += '</tr></thead>\n<tbody>\n'

        # Rows
        for row in comparison['rows']:
            urgency_class = row.get('urgency', 'routine')
            content += f'<tr class="{urgency_class}">\n'
            for cell in row['cells']:
                content += f'<td>{cell}</td>\n'
            content += '</tr>\n'

        content += '</tbody>\n</table>\n'

    return content


def generate_tab3_content():
    """Generate Tab 3: Master Comparison Tables content."""
    if not MASTER_TABLES:
        return '<p style="color: #9ca3af;">No master tables provided.</p>'

    content = ''
    for table in MASTER_TABLES:
        content += f'<h2 class="comparison-title">{table["title"]}</h2>\n'

        if 'description' in table:
            content += f'<div class="master-description">{table["description"]}</div>\n'

        content += '<table>\n<thead><tr>\n'

        # Headers
        for header in table['headers']:
            content += f'<th>{header}</th>\n'
        content += '</tr></thead>\n<tbody>\n'

        # Rows
        for row in table['rows']:
            urgency_class = row.get('urgency', 'routine')
            content += f'<tr class="{urgency_class}">\n'
            for cell in row['cells']:
                content += f'<td>{cell}</td>\n'
            content += '</tr>\n'

        content += '</tbody>\n</table>\n'

    return content


def generate_tab4_content():
    """Generate Tab 4: Summary content."""
    content = ''

    # Key Points
    if SUMMARY_DATA.get('key_points'):
        content += '<div class="summary-section">\n<h3>🎯 Key Points</h3>\n'
        for point in SUMMARY_DATA['key_points']:
            content += f'<div class="key-point">{point}</div>\n'
        content += '</div>\n'

    # Mnemonics
    if SUMMARY_DATA.get('mnemonics'):
        content += '<div class="summary-section">\n<h3>🧠 Mnemonics</h3>\n'
        for mnem in SUMMARY_DATA['mnemonics']:
            content += '<div class="mnemonic">\n'
            content += f'<div class="mnemonic-title">{mnem["topic"]}</div>\n'
            content += f'<div class="mnemonic-letters">{mnem["mnemonic"]}</div>\n'
            content += '<div class="mnemonic-breakdown">\n'
            for letter, meaning in mnem['breakdown'].items():
                content += f'<div><strong>{letter}</strong> = {meaning}</div>\n'
            content += '</div>\n</div>\n'
        content += '</div>\n'

    # If-Then Rules
    if SUMMARY_DATA.get('if_then_rules'):
        content += '<div class="summary-section">\n<h3>🔍 If-Then Clinical Rules</h3>\n'
        for rule in SUMMARY_DATA['if_then_rules']:
            content += '<div class="if-then">\n'
            content += f'<div class="if-then-if">IF: {rule["if"]}</div>\n'
            content += f'<div class="if-then-then">THEN: {rule["then"]}</div>\n'
            content += '</div>\n'
        content += '</div>\n'

    # Critical Values
    if SUMMARY_DATA.get('critical_values'):
        content += '<div class="summary-section">\n<h3>📊 Critical Values</h3>\n'
        content += '<table>\n<thead><tr><th>Parameter</th><th>Value</th></tr></thead>\n<tbody>\n'
        for val in SUMMARY_DATA['critical_values']:
            content += f'<tr><td>{val["parameter"]}</td><td>{val["value"]}</td></tr>\n'
        content += '</tbody>\n</table>\n</div>\n'

    # Definitions
    if SUMMARY_DATA.get('definitions'):
        content += '<div class="summary-section">\n<h3>📖 Key Definitions</h3>\n'
        for defn in SUMMARY_DATA['definitions']:
            content += '<div class="definition">\n'
            content += f'<span class="definition-term">{defn["term"]}:</span> {defn["definition"]}\n'
            content += '</div>\n'
        content += '</div>\n'

    # Pearls
    if SUMMARY_DATA.get('pearls'):
        content += '<div class="summary-section">\n<h3>💎 High-Yield Pearls</h3>\n'
        for pearl in SUMMARY_DATA['pearls']:
            content += f'<div class="pearl">{pearl}</div>\n'
        content += '</div>\n'

    if not content:
        content = '<p style="color: #9ca3af;">No summary data provided.</p>'

    return content


# =============================================================================
# VALIDATION
# =============================================================================

def validate_data():
    """Validate that required data is provided."""
    issues = []

    if not LEARNING_OBJECTIVES:
        issues.append("⚠️  No learning objectives provided")

    if not KEY_COMPARISONS:
        issues.append("⚠️  No key comparisons provided")

    if not MASTER_TABLES:
        issues.append("⚠️  No master tables provided")

    if not any(SUMMARY_DATA.values()):
        issues.append("⚠️  No summary data provided")

    return issues


# =============================================================================
# MAIN FUNCTION
# =============================================================================

def main():
    """Main entry point."""
    print("=" * 70)
    print(f"HTML LEARNING OBJECTIVES GENERATOR - {LECTURE_TOPIC}")
    print("=" * 70)
    print()
    print(f"Source: {SOURCE_FILE}")
    print()

    # Validate data
    issues = validate_data()
    if issues:
        print("⚠️  Validation Warnings:")
        for issue in issues:
            print(f"  {issue}")
        print()
        print("Continuing anyway... (file will be created with available data)")
        print()

    # Show configuration
    print("📊 Configuration:")
    print(f"   Learning Objectives: {len(LEARNING_OBJECTIVES)}")
    print(f"   Key Comparisons: {len(KEY_COMPARISONS)}")
    print(f"   Master Tables: {len(MASTER_TABLES)}")
    print(f"   Summary sections: {sum(1 for v in SUMMARY_DATA.values() if v)}")
    print(f"   Output: {OUTPUT_FILENAME}")
    print()

    # Generate HTML
    print("🔨 Generating HTML...")
    html_content = generate_html()
    print("   ✓ HTML structure created")
    print("   ✓ All 4 tabs generated")
    print("   ✓ Interactive navigation added")
    print()

    # Save file
    print(f"💾 Saving to: {OUTPUT_FILENAME}")
    output_path = Path(OUTPUT_FILENAME)
    output_path.write_text(html_content, encoding='utf-8')
    print()
    print("=" * 70)
    print("✅ SUCCESS!")
    print("=" * 70)

    # Summary
    print()
    print("📋 Summary:")
    print(f"   • File: {OUTPUT_FILENAME}")
    print(f"   • Tabs: 4 (LOs, Key Comparisons, Master Tables, Summary)")
    print(f"   • Interactive: YES ✅")
    print(f"   • Mobile-responsive: YES ✅")
    print()

    # Auto-open in browser
    print(f"🌐 Opening {OUTPUT_FILENAME} in browser...")
    try:
        file_path = str(output_path.absolute())
        if platform.system() == 'Darwin':  # macOS
            subprocess.run(['open', file_path], check=True)
        elif platform.system() == 'Windows':
            subprocess.run(['start', file_path], shell=True, check=True)
        else:  # Linux
            subprocess.run(['xdg-open', file_path], check=True)
        print("   ✓ Opened in default browser")
    except Exception as e:
        print(f"   ⚠️  Could not auto-open file: {e}")
        print(f"   → Manually open: {OUTPUT_FILENAME}")

    print()
    return 0


if __name__ == '__main__':
    sys.exit(main())
