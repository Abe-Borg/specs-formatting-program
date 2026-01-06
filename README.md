# Spec Automation Tool

A GUI application for automatically formatting MEP construction specifications to match architect-provided templates. This tool extracts content from office master specifications, intelligently filters out editing instructions and specifier notes, and rebuilds the documents with the exact formatting DNA from your project's architectural template.

## 🎯 What It Does

The Spec Automation Tool solves a common problem in construction documentation: taking your office's master specifications and reformatting them to match the specific style requirements of each project's architect.

**The Process:**
1. **Extract** → Reads your office master `.docx` files and converts content to markdown, automatically filtering out:
   - Specifier notes and editing instructions
   - Copyright notices and metadata
   - Any content that shouldn't appear in the final specification
   
2. **Reformat** → Rebuilds the specification using the architect's template, applying:
   - Custom heading styles (CSI Level 0, Level 1, Level 2, etc.)
   - List formatting with proper indentation
   - Page layout, margins, headers, and footers from the template

3. **Output** → Produces clean, properly formatted specifications ready for coordination and printing

## 🚀 Quick Start

### Prerequisites

- **Python 3.7+** installed on your Windows machine
- **pip** (Python package installer)

### Installation

1. **Install Required Libraries:**
   ```bash
   pip install python-docx
   ```

2. **Download the Tool:**
   - Save `spec_tool_gui.py` to your preferred location
   - Save `project_config.json` to the same folder
   - Save `check_styles.py` (optional utility) to the same folder

### First Run

Double-click `spec_tool_gui.py` or run from command line:
```bash
python spec_tool_gui.py
```

## 📖 How to Use

### Step 1: Prepare Your Files

**You'll need:**
- **Office Masters Folder** → Directory containing your office's master specification `.docx` files
- **Architect Template** → The `.docx` file provided by the architect with their formatting
- **Output Folder** → Where you want the formatted specifications saved

### Step 2: Configure Style Mappings (First Time Only)

Before running the tool on a new project, you need to identify what style names the architect is using:

1. Run the **`check_styles.py`** utility:
   ```bash
   python check_styles.py
   ```
   This will print all paragraph styles found in the architect's template.

2. Open `project_config.json` and update the style names to match:
   ```json
   {
       "styles": {
           "Title": "CSILevel0",      ← Update these
           "Part": "CSILevel1",        ← with actual
           "Article": "CSILevel2"      ← style names
       },
       "list_levels": [
           "CSILevel3", 
           "CSILevel4", 
           "CSILevel5"
       ]
   }
   ```

### Step 3: Run the Tool

1. Launch `spec_tool_gui.py`
2. Click **Browse** buttons to select:
   - Office Masters Folder
   - Architect Template (.docx)
   - Project Config (.json)
   - Output Folder
3. Click **🚀 PROCESS ALL SPECS**
4. Watch the log window for progress
5. Find your formatted specifications in the Output Folder

## ⚙️ Configuration Guide

### Understanding `project_config.json`

```json
{
    "styles": {
        "Title": "CSILevel0",       // Section headers (SECTION 23 21 13)
        "Part": "CSILevel1",        // Part headers (PART 1 - GENERAL)
        "Article": "CSILevel2"      // Article headers (1.01 SUMMARY)
    },
    "list_levels": [
        "CSILevel3",                // First indent (A. Item)
        "CSILevel4",                // Second indent (1. Subitem)
        "CSILevel5",                // Third indent (a. Detail)
        "CSILevel6",                // Fourth indent (1) Further detail)
        "CSILevel7"                 // Fifth indent (a) Deepest level)
    ],
    "options": {
        "strip_heading_numbers": true,    // Remove "1.01" from "1.01 SUMMARY"
        "strip_list_labels": true         // Remove "A." from "A. List item"
    }
}
```

### Finding Style Names

Use `check_styles.py` to inspect any Word document:

```bash
python check_styles.py
```

Make sure `architect_style.docx` points to the correct template file in the script, or modify the script to accept command-line arguments.

### Customizing Filtering

In `spec_tool_gui.py`, you can customize what gets filtered out during extraction:

```python
# Line 47-54
IGNORED_STYLES = [
    "Specifier Note", 
    "Note", 
    "Instruction", 
    "Editing Instruction", 
    'CMT'
]

# Line 57-64
IGNORED_STARTS = [
    "See Editing Instruction",
    "Adjust list below",
    "Retain ",
    "Delete ",
    "Edit ",
    "Verify that Section titles"
]
```

Add or remove patterns based on what appears in your office masters.

## 🛠️ Advanced Usage

### Processing Individual Files

While the GUI processes entire folders, you can modify the script for single-file processing by calling the functions directly:

```python
from spec_tool_gui import extract_master_to_markdown, rebuild_from_markdown, load_config

config = load_config("project_config.json")
extract_master_to_markdown("master.docx", "temp.md", print)
rebuild_from_markdown("temp.md", "template.docx", "output.docx", config, print)
```

### Creating Project-Specific Configs

For projects with multiple architects or phases, create separate config files:

```
project_configs/
├── architect_a_config.json
├── architect_b_config.json
└── phase_2_config.json
```

Select the appropriate config when running the tool.

## 🐛 Troubleshooting

### "No .docx files found in Master folder"
- Make sure you're pointing to a folder containing `.docx` files (not `.doc`)
- Check that files aren't hidden or in subfolders
- Ensure file names don't start with `~$` (temp files)

### "Template not found"
- Verify the architect template path is correct
- Make sure it's a `.docx` file (not `.dotx` template)
- Try copying the template to the same folder as the script

### Formatting looks wrong
- Run `check_styles.py` to verify style names
- Compare `project_config.json` style names with template styles
- Check that `strip_heading_numbers` and `strip_list_labels` are set appropriately

### Content is missing
- Check `IGNORED_STYLES` and `IGNORED_STARTS` lists
- Make sure legitimate content isn't being filtered
- Examine the temporary `.md` files (uncomment line 237 to keep them)

### Program freezes
- Large files may take time to process
- Check the log window for errors
- Try processing one file manually to diagnose issues

## 📝 File Structure

```
spec-automation-tool/
├── spec_tool_gui.py          # Main application
├── project_config.json        # Style mapping configuration
├── check_styles.py           # Utility to inspect Word styles
└── README.md                 # This file
```

## 🔄 Workflow Integration

**Typical Project Workflow:**

1. **Project Kickoff** → Architect provides template with their style guide
2. **Style Analysis** → Run `check_styles.py` on architect template
3. **Config Setup** → Create/update `project_config.json` with correct style names
4. **Batch Processing** → Use GUI to process all relevant specifications
5. **QC Review** → Spot-check formatted specs for correctness
6. **Coordination** → Proceed with normal specification coordination workflow

## 🎓 Background

This tool implements the "extract → clean → rebuild" pattern for document formatting automation. It works by:

- Converting Word documents to an intermediate markdown representation
- Applying intelligent filtering rules during extraction
- Using the architect's template as a "formatting DNA" source
- Surgically rebuilding documents with the target formatting

The approach preserves content while completely replacing formatting, ensuring consistency across all project specifications.

## 💡 Tips

- **Test First:** Always test with one specification before batch processing
- **Keep Temps:** Uncomment line 237 to keep markdown files for debugging
- **Version Control:** Keep project configs in version control with project files
- **Backup:** Keep original masters untouched; always work with copies
- **Style Guide:** Create a reference document showing what each style looks like

## 📄 License

---

## Copyright Notice

**Copyright © 2025 Andrew Gossman. All Rights Reserved.**

This software and associated documentation files (the "Software") are the proprietary property of Andrew Gossman. 

**Unauthorized copying, modification, distribution, or use of this Software, via any medium, is strictly prohibited without express written permission from the copyright holder.**

This Software is provided for review and reference purposes only. No license or right to use, copy, modify, or distribute this Software for any purpose, commercial or non-commercial, is granted.

For licensing inquiries, please contact Andrew Gossman.

---


# Critical Architectural Limitation

Markdown as the intermediate format is fundamentally lossy. This single design choice cascades into most of the issues below.

When you extract to markdown and rebuild, you permanently lose:

| Lost Element | Impact |
|--------------|--------|
| Tables | Schedules, performance criteria, product tables—gone |
| Images | Diagrams, details, manufacturer logos—gone |
| Inline formatting | Bold, italic, underline within paragraphs—gone |
| Hyperlinks | URL references—gone |
| Track changes | Revision history—gone |
| Comments | Review notes—gone |
| Fields | Auto-numbering, cross-references, TOC entries—gone |
| Bookmarks | Internal document links—gone |

For MEP specs that include equipment schedules or piping diagrams, this is a dealbreaker.

---

## Regex-Based Structure Detection Failures

The current detection logic has several blind spots:

### Part Detection Issues

```python
# Current: Requires "PART" AND a dash
elif text.upper().startswith("PART") and "-" in text:
```

Will fail on:
- `PART 1 GENERAL` (no dash—common in older masters)
- `PART ONE - GENERAL` (spelled out)
- Multi-line part titles where the number is separate from the title

### Article Detection Issues

```python
# Current: Article detection
elif re.match(r'^\d+\.\d+\s', text):
```

Will fail on:
- `1.01` with no space after (tight formatting)
- `1.1 SUMMARY` (single digit after decimal)
- Sub-articles like `2.01.A` or `2.01.1` (nested numbering)

### List Label Detection Issues

```python
# Current: List label detection
if re.match(r'^[A-Z]\.\s', text):       # Level 1
elif re.match(r'^\d+\.\s', text):       # Level 2
elif re.match(r'^[a-z]\.\s', text):     # Level 3
elif re.match(r'^\d+\)\s', text):       # Level 4
elif re.match(r'^[a-z]\)\s', text):     # Level 5
```

Will fail on:
- CSI Level 6: `(1)`, `(2)`, etc.
- CSI Level 7: `(a)`, `(b)`, etc.
- Roman numerals: `i.`, `ii.`, `I.`, `II.`
- Double letters: `AA.`, `BB.` (overflow labels)
- Continuation paragraphs (text that belongs to the previous list item but has no label)

---

## Style Dependency Issues

### Hard-coded Ignored Styles

```python
IGNORED_STYLES = [
    "Specifier Note", 
    "Note", 
    "Instruction", 
    "Editing Instruction", 
    'CMT'
]
```

Different master file publishers use different style names:
- **BSD:** `SpecNote`, `EditNote`
- **ARCOM:** `Specifier Note`, `SpN`
- **Custom masters:** Anything goes

This requires per-publisher configuration, not hard-coding.

### Silent Style Fallback

```python
def add_safe_paragraph(text, style_name):
    try:
        doc.add_paragraph(text, style=style_name)
    except KeyError:
        doc.add_paragraph(text, style='Normal')  # Silent failure!
```

If your config references `CSILevel6` but the template only has up to `CSILevel5`, you'll get `Normal` style with no warning. The output will look wrong and you won't know why.

---

## Config Validation Gaps

Your `project_config.json` references styles like `CSILevel0` through `CSILevel7`, but:

- No validation that these styles exist in the template
- No schema validation for the JSON structure
- `KeyError` risk if required keys are missing:

```python
style_map = config.get("styles", DEFAULT_CONFIG["styles"])
# Later...
add_safe_paragraph(clean_text, style_map["Title"])  # KeyError if "Title" missing
```

---

## Threading & GUI Issues

```python
def log(self, message):
    self.log_area.config(state='normal')
    self.log_area.insert(tk.END, message + "\n")
    # Called from worker thread - not thread-safe!
```

Tkinter is not thread-safe. Calling widget methods from a worker thread can cause:
- Random crashes
- Frozen GUI
- Garbled text

You need to use `root.after()` or a queue-based approach.

---

## Indentation Detection is Fragile

```python
dash_index = raw_text.find("-")
indent_level = dash_index // 2
```

This assumes:
- Exactly 2 spaces per indent level
- No tabs
- The dash position in the intermediate markdown accurately reflects the intended CSI level

If anything shifts the dash position (mixed tabs/spaces, different editor settings), your list hierarchy breaks.

---

## File Handling Edge Cases

| Scenario | Current Behavior |
|----------|------------------|
| File locked by Word | Crash or silent failure |
| Non-UTF-8 encoding | Mojibake or crash |
| Path with special characters | Platform-dependent failure |
| Very large files (100+ pages) | Memory issues, no progress feedback |
| Corrupted .docx | Unhandled exception |

---

## What's Missing Entirely

1. **Table preservation** — Would require XML-level extraction and rebuild
2. **Image handling** — Need to extract from docx media folder, track positions
3. **Run-level formatting** — Bold/italic requires parsing `<w:r>` elements, not just paragraph text
4. **Continuation paragraphs** — Text that belongs to a list item but has no label
5. **Spec note filtering by bracket syntax** — Many masters use `[specifier note text]` inline
6. **Multi-section files** — Some masters combine multiple spec sections in one file
7. **Validation/preview** — No way to see what will change before committing
8. **Undo/rollback** — If something goes wrong, you've overwritten your work

---

## Recommendations for Improvement

### Short-term (Keep current architecture)

- Add style existence validation before processing
- Fix threading with a proper queue
- Add more regex patterns for edge cases
- Make `IGNORED_STYLES` configurable in JSON
- Add a "dry run" mode that shows what would happen

### Medium-term (Better architecture)

- Skip markdown entirely—work at the python-docx XML level
- Copy paragraph-by-paragraph while transforming styles
- Preserve tables and images by copying elements directly

### Long-term (Production-ready)

- Use your SpecCleanse approach—work at the ZIP/XML level
- Build a proper AST representation of CSI structure
- Implement surgical style remapping without content extraction
- Add a diffing/preview system
