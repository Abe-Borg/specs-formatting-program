from docx import Document

# Load the architect's template
doc = Document('architect_style.docx')

print("--- STYLES FOUND IN TEMPLATE ---")
# Loop through all styles and print the Paragraph styles
for style in doc.styles:
    if style.type.name == 'PARAGRAPH':
        print(f"• {style.name}")