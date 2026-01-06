from docx import Document

"""
Style Inspector Utility for Word Documents

This utility script extracts and displays all paragraph styles from a Word document.
Use this before configuring a new project to identify the exact style names used in
the architect's template document.

Usage:
    1. Update the filename 'architect_style.docx' to point to your template
    2. Run: python check_styles.py
    3. Copy the style names into your project_config.json

Output:
    Prints a bulleted list of all paragraph style names found in the document.

Example Output:
    --- STYLES FOUND IN TEMPLATE ---
    • Normal
    • Heading 1
    • CSILevel0
    • CSILevel1
    • List Bullet

Dependencies:
    - python-docx

Author: Abraham
Project: Spec Automation Tool
"""

# Load the architect's template
doc = Document('architect_style.docx')

print("--- STYLES FOUND IN TEMPLATE ---")
# Loop through all styles and print the Paragraph styles
for style in doc.styles:
    if style.type.name == 'PARAGRAPH':
        print(f"• {style.name}")