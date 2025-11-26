import os
from docx import Document

print(f"Current Working Directory: {os.getcwd()}")
print("\n--- FILES PYTHON SEES ---")
files = os.listdir()
for f in files:
    # We use repr() to reveal hidden spaces or weird characters
    print(f"Found: {repr(f)}") 

print("\n--- ATTEMPTING TO OPEN ---")
target = 'architect_style.docx'
if target in files:
    print(f"✅ Python sees '{target}'. Opening...")
    try:
        doc = Document(target)
        print("✅ SUCCESS: File opened.")
        print("Styles found inside:")
        for s in doc.styles:
            if s.type.name == 'PARAGRAPH':
                print(f" - {s.name}")
    except Exception as e:
        print(f"❌ ERROR opening file: {e}")
else:
    print(f"❌ Python DOES NOT see '{target}' in this list.")