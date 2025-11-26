import re
import os
from docx import Document
from docx.shared import Inches, Pt

# ==========================================
# CONFIGURATION
# ==========================================
INPUT_MASTER = "office_master.docx"       # Your existing spec
ARCHITECT_TEMPLATE = "architect_style.docx" # The file with correct margins/headers
OUTPUT_FILE = "Final_Project_Spec.docx"
TEMP_MARKDOWN = "temp_converted.md"

# Indentation Settings (in Inches)
# These align with your 2-space, 4-space markdown rules
INDENT_PER_LEVEL = 0.5 

# ==========================================
# STEP 1: CONVERT MASTER DOCX TO MARKDOWN
# ==========================================
def extract_master_to_markdown(docx_path, md_path):
    """
    Reads the Master Docx and attempts to create the custom Markdown.
    NOTE: This assumes the Master Docx has the "A." "1." typed out 
    or detectable as text. If they are auto-numbered lists, 
    python-docx reads them as empty text.
    """
    if not os.path.exists(docx_path):
        print(f"❌ Error: Master file {docx_path} not found.")
        return

    doc = Document(docx_path)
    md_lines = []

    print(f"1️⃣  Extracting content from {docx_path}...")

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # --- LOGIC TO DETECT HIERARCHY ---
        
        # 1. Section Titles (Start with SECTION)
        if text.upper().startswith("SECTION"):
            md_lines.append(f"# {text}")
            
        # 2. Parts (PART 1 - GENERAL)
        elif text.upper().startswith("PART") and "-" in text:
            md_lines.append(f"\n## {text}")
            
        # 3. Articles (1.01 SUMMARY) -> Regex for digit.digit
        elif re.match(r'^\d+\.\d+\s', text):
            md_lines.append(f"\n### {text}")
            
        # 4. Paragraph Levels (The hard part - detecting A. 1. a.)
        # We check the start of the string to assign indentation depth
        else:
            # Level 1: A. B. C.
            if re.match(r'^[A-Z]\.\s', text):
                md_lines.append(f"- {text}")
            
            # Level 2: 1. 2. 3.
            elif re.match(r'^\d+\.\s', text):
                md_lines.append(f"  - {text}")
                
            # Level 3: a. b. c.
            elif re.match(r'^[a-z]\.\s', text):
                md_lines.append(f"    - {text}")
                
            # Level 4: 1) 2)
            elif re.match(r'^\d+\)\s', text):
                md_lines.append(f"      - {text}")
                
            # Level 5: a) b)
            elif re.match(r'^[a-z]\)\s', text):
                md_lines.append(f"        - {text}")
                
            # Level 6: 1) 2) (Nested deeper, unlikely but possible)
            # Since 1) matches Level 4 regex, context matters, but for raw extraction:
            else:
                # Fallback for standard text or unformatted lines
                # We assume it belongs to the previous level or is a note
                md_lines.append(f"  {text}")

    # Write to Markdown file
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))
    
    print(f"   Converted to Markdown: {md_path}")


# ==========================================
# STEP 2: REBUILD DOCX USING TEMPLATE
# ==========================================
def rebuild_from_markdown(md_path, template_path, output_path):
    """
    Takes the custom Markdown and injects it into the Architect's template.
    It strictly enforces indentation based on the dash hierarchy.
    """
    if not os.path.exists(template_path):
        print(f"❌ Error: Template {template_path} not found.")
        return

    print(f"2️⃣  Applying styles from {template_path}...")
    
    # Load Architect's Doc
    doc = Document(template_path)
    
    # CLEAR BODY CONTENT
    # This keeps margins, headers, footers, and styles intact
    for element in list(doc.element.body):
        # Don't delete section properties (sectPr) which hold margins/headers
        if element.tag.endswith('sectPr'):
            continue
        doc.element.body.remove(element)

    # Read the Markdown
    with open(md_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    # Process Line by Line
    for line in lines:
        raw_text = line.rstrip() # Keep leading spaces, remove trailing newline
        stripped_text = raw_text.strip()
        
        if not stripped_text:
            continue

        # --- HEADING LOGIC ---
        if stripped_text.startswith("# "):
            doc.add_heading(stripped_text.replace("# ", ""), level=1)
            
        elif stripped_text.startswith("## "):
            doc.add_heading(stripped_text.replace("## ", ""), level=2)
            
        elif stripped_text.startswith("### "):
            doc.add_heading(stripped_text.replace("### ", ""), level=3)

        # --- LIST LOGIC (The "Embedded Label" System) ---
        elif stripped_text.startswith("- "):
            # Calculate Indent Level based on leading spaces in raw_text
            # We look for the position of the dash
            dash_index = raw_text.find("-")
            
            # 2 spaces per level logic
            # 0 spaces = Level 0 (A.)
            # 2 spaces = Level 1 (1.)
            # 4 spaces = Level 2 (a.)
            indent_level = dash_index // 2 
            
            # Clean the text (Remove "- " but keep "A. Text")
            content = stripped_text[2:] 
            
            # Create a 'Normal' paragraph (inherits Font from Architect)
            p = doc.add_paragraph(content)
            
            # FORCE INDENTATION
            # We manually push the left indent based on the level.
            # We use a "hanging indent" so if text wraps, it looks nice.
            p_format = p.paragraph_format
            
            # Math: 
            # Level 0: 0.0" Indent
            # Level 1: 0.5" Indent
            # Level 2: 1.0" Indent
            indent_amount = Inches(indent_level * INDENT_PER_LEVEL)
            
            p_format.left_indent = indent_amount
            
            # Optional: Add a hanging indent (first line sticks out)
            # This makes "A." stick out and the text wrap aligned
            p_format.first_line_indent = Inches(-0.25)
            # We have to increase left_indent slightly to compensate for the hang
            p_format.left_indent = indent_amount + Inches(0.25)

        # --- STANDARD TEXT / UNKNOWN ---
        else:
            doc.add_paragraph(stripped_text, style='Normal')

    doc.save(output_path)
    print(f"✅ Success! Generated: {output_path}")

# ==========================================
# EXECUTION
# ==========================================
if __name__ == "__main__":
    # Ensure you have these files in the folder before running
    # 1. office_master.docx
    # 2. architect_style.docx
    
    extract_master_to_markdown(INPUT_MASTER, TEMP_MARKDOWN)
    rebuild_from_markdown(TEMP_MARKDOWN, ARCHITECT_TEMPLATE, OUTPUT_FILE)