import re
import os
from docx import Document

# ==========================================
# 🛠️ CONFIGURATION AREA (EDIT THIS PER PROJECT)
# ==========================================

# INPUTS
INPUT_MASTER = "office_master.docx"
ARCHITECT_TEMPLATE = "architect_style.docx"
OUTPUT_FILE = "Final_Project_Spec.docx"
TEMP_MARKDOWN = "temp_converted.md"

# ------------------------------------------
# 🎨 STYLE MAPPING (The "Translation Layer")
# ------------------------------------------
# INSTRUCTIONS: Uncomment the block that matches your current Architect.

# --- OPTION A: THE "CSI LEVEL" ARCHITECT (Your current one) ---
STYLE_MAP = {
    "Title":   "CSILevel0",  # Matches # (Section Title)
    "Part":    "CSILevel1",  # Matches ## (Part 1)
    "Article": "CSILevel2",  # Matches ### (1.01)
}
# Map indent levels to styles: [Level 0 (A.), Level 1 (1.), Level 2 (a.)...]
LIST_MAP = ["CSILevel3", "CSILevel4", "CSILevel5", "CSILevel6", "CSILevel7"]


# --- OPTION B: THE "STANDARD WORD" ARCHITECT (Most others) ---
# STYLE_MAP = {
#     "Title":   "Heading 1",
#     "Part":    "Heading 2",
#     "Article": "Heading 3",
# }
# # Standard Word usually uses "List Bullet", "List Bullet 2", etc.
# LIST_MAP = ["List Bullet", "List Bullet 2", "List Bullet 3", "List Bullet 4", "List Bullet 5"]


# ==========================================
# END OF CONFIGURATION
# ==========================================

def extract_master_to_markdown(docx_path, md_path):
    """
    Step 1: Extract content to Markdown (No changes needed here)
    """
    if not os.path.exists(docx_path):
        print(f"❌ Error: Master file {docx_path} not found.")
        return

    doc = Document(docx_path)
    md_lines = []

    print(f"1️⃣  Extracting content from {docx_path}...")

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text: continue

        # Regex detection for CSI Structure
        if text.upper().startswith("SECTION"):
            md_lines.append(f"# {text}")
        elif text.upper().startswith("PART") and "-" in text:
            md_lines.append(f"\n## {text}")
        elif re.match(r'^\d+\.\d+\s', text):
            md_lines.append(f"\n### {text}")
        else:
            # Indentation Detection
            if re.match(r'^[A-Z]\.\s', text):       md_lines.append(f"- {text}")     # Level 0
            elif re.match(r'^\d+\.\s', text):       md_lines.append(f"  - {text}")   # Level 1
            elif re.match(r'^[a-z]\.\s', text):     md_lines.append(f"    - {text}") # Level 2
            elif re.match(r'^\d+\)\s', text):       md_lines.append(f"      - {text}") # Level 3
            elif re.match(r'^[a-z]\)\s', text):     md_lines.append(f"        - {text}") # Level 4
            else:                                   md_lines.append(f"  {text}")     # Default text

    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))
    print(f"   Converted to Markdown: {md_path}")


def rebuild_from_markdown(md_path, template_path, output_path):
    """
    Step 2: Rebuild using the Configuration Map
    """
    if not os.path.exists(template_path):
        print(f"❌ Error: Template {template_path} not found.")
        return

    print(f"2️⃣  Applying styles from {template_path}...")
    
    doc = Document(template_path)
    
    # Clean Body
    for element in list(doc.element.body):
        if element.tag.endswith('sectPr'): continue
        doc.element.body.remove(element)

    with open(md_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    # Helper function to safely add paragraph with fallback
    def add_safe_paragraph(text, style_name):
        try:
            doc.add_paragraph(text, style=style_name)
        except KeyError:
            print(f"⚠️  WARNING: Style '{style_name}' not found in template. Using 'Normal'.")
            doc.add_paragraph(text, style='Normal')
        except Exception as e:
            print(f"❌ Error adding text '{text[:20]}...': {e}")

    # Process Lines
    for line in lines:
        raw_text = line.rstrip()
        stripped_text = raw_text.strip()
        if not stripped_text: continue

        # --- HEADERS (Using STYLE_MAP) ---
        if stripped_text.startswith("# "):
            add_safe_paragraph(stripped_text.replace("# ", ""), STYLE_MAP["Title"])
            
        elif stripped_text.startswith("## "):
            add_safe_paragraph(stripped_text.replace("## ", ""), STYLE_MAP["Part"])
            
        elif stripped_text.startswith("### "):
            add_safe_paragraph(stripped_text.replace("### ", ""), STYLE_MAP["Article"])

        # --- LISTS (Using LIST_MAP) ---
        elif stripped_text.startswith("- "):
            dash_index = raw_text.find("-")
            indent_level = dash_index // 2 
            clean_text = stripped_text[2:]

            # Pick the correct style from the list
            if indent_level < len(LIST_MAP):
                target_style = LIST_MAP[indent_level]
            else:
                # If nested deeper than our map, use the last available style
                target_style = LIST_MAP[-1]
            
            add_safe_paragraph(clean_text, target_style)

        # --- STANDARD TEXT ---
        else:
            add_safe_paragraph(stripped_text, 'Normal')

    doc.save(output_path)
    print(f"✅ Success! Generated: {output_path}")

if __name__ == "__main__":
    extract_master_to_markdown(INPUT_MASTER, TEMP_MARKDOWN)
    rebuild_from_markdown(TEMP_MARKDOWN, ARCHITECT_TEMPLATE, OUTPUT_FILE)