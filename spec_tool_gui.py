import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import threading
import os
import re
import json
from docx import Document

# ==========================================
# 🧠 THE LOGIC ENGINE (Our Previous Script)
# ==========================================

DEFAULT_CONFIG = {
    "styles": {
        "Title": "Heading 1", 
        "Part": "Heading 2", 
        "Article": "Heading 3"
    },
    "list_levels": ["List Bullet", "List Bullet 2", "List Bullet 3"],
    "options": {
        "strip_heading_numbers": True, 
        "strip_list_labels": True
    }
}

def load_config(config_path):
    if config_path and os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except Exception:
            return DEFAULT_CONFIG
    return DEFAULT_CONFIG

def extract_master_to_markdown(docx_path, md_path, log_func):
    if not os.path.exists(docx_path):
        return False

    try:
        doc = Document(docx_path)
        md_lines = []

        # 🚫 STYLES TO IGNORE (Customize this list!)
        # Check your master file to see what the notes are called.
        IGNORED_STYLES = [
            "Specifier Note", 
            "Note", 
            "Instruction", 
            "Editing Instruction", 
            'CMT'
        ]

        # 🚫 KEYWORDS TO SKIP (Backup if styles are messy)
        # If a line starts with these, we skip it.
        IGNORED_STARTS = [
            "See Editing Instruction",
            "Adjust list below",
            "Retain ",
            "Delete ",
            "Edit ",
            "Verify that Section titles"
        ]

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue

            # CHECK 1: Is it an ignored style?
            if para.style.name in IGNORED_STYLES:
                continue

            # CHECK 2: Does it look like an instruction? (Keyword check)
            # (Only use this if styles aren't working perfectly)
            is_instruction = False
            for keyword in IGNORED_STARTS:
                if text.startswith(keyword):
                    is_instruction = True
                    break
            if is_instruction:
                continue

            # Regex detection for CSI Structure
            if text.upper().startswith("SECTION"):
                md_lines.append(f"# {text}")
            elif text.upper().startswith("PART") and "-" in text:
                md_lines.append(f"\n## {text}")
            elif re.match(r'^\d+\.\d+\s', text):
                md_lines.append(f"\n### {text}")
            else:
                # Indentation Detection
                if re.match(r'^[A-Z]\.\s', text):       md_lines.append(f"- {text}")
                elif re.match(r'^\d+\.\s', text):       md_lines.append(f"  - {text}")
                elif re.match(r'^[a-z]\.\s', text):     md_lines.append(f"    - {text}")
                elif re.match(r'^\d+\)\s', text):       md_lines.append(f"      - {text}")
                elif re.match(r'^[a-z]\)\s', text):     md_lines.append(f"        - {text}")
                else:                                   md_lines.append(f"  {text}")

        with open(md_path, "w", encoding="utf-8") as f:
            f.write("\n".join(md_lines))
        return True
    except Exception as e:
        log_func(f"❌ Error extracting {os.path.basename(docx_path)}: {str(e)}")
        return False

def rebuild_from_markdown(md_path, template_path, output_path, config, log_func):
    if not os.path.exists(template_path):
        log_func(f"❌ Template not found: {template_path}")
        return False

    try:
        # Load Config
        style_map = config.get("styles", DEFAULT_CONFIG["styles"])
        list_map = config.get("list_levels", DEFAULT_CONFIG["list_levels"])
        options = config.get("options", DEFAULT_CONFIG["options"])

        doc = Document(template_path)
        
        # Clear Body (Keep Headers/Footers/Margins)
        for element in list(doc.element.body):
            if element.tag.endswith('sectPr'): continue
            doc.element.body.remove(element)

        with open(md_path, "r", encoding="utf-8") as f:
            lines = f.readlines()

        # Helpers
        def add_safe_paragraph(text, style_name):
            try:
                doc.add_paragraph(text, style=style_name)
            except KeyError:
                doc.add_paragraph(text, style='Normal')

        def clean_heading_label(text):
            if options.get("strip_heading_numbers"):
                return re.sub(r'^\d+\.\d+\s+', '', text)
            return text

        def clean_list_label(text):
            if options.get("strip_list_labels"):
                return re.sub(r'^[A-Za-z0-9]+[\.\)]\s+', '', text)
            return text

        previous_line_was_title = False

        for line in lines:
            raw_text = line.rstrip()
            stripped_text = raw_text.strip()
            
            if not stripped_text: continue

            # SPECIAL: End of Section
            if stripped_text.upper().startswith("END OF SECTION"):
                add_safe_paragraph(stripped_text, style_map["Title"])
                previous_line_was_title = False
                continue

            # HEADERS
            if stripped_text.startswith("# "):
                clean_text = stripped_text.replace("# ", "")
                add_safe_paragraph(clean_text, style_map["Title"])
                previous_line_was_title = True 
                
            elif stripped_text.startswith("## "):
                clean_text = stripped_text.replace("## ", "")
                add_safe_paragraph(clean_text, style_map["Part"])
                previous_line_was_title = False
                
            elif stripped_text.startswith("### "):
                text_no_hash = stripped_text.replace("### ", "")
                clean_text = clean_heading_label(text_no_hash)
                add_safe_paragraph(clean_text, style_map["Article"])
                previous_line_was_title = False

            # LISTS
            elif stripped_text.startswith("- "):
                dash_index = raw_text.find("-")
                indent_level = dash_index // 2 
                
                text_without_dash = stripped_text[2:]
                clean_content = clean_list_label(text_without_dash)

                if indent_level < len(list_map):
                    target_style = list_map[indent_level]
                else:
                    target_style = list_map[-1]
                
                add_safe_paragraph(clean_content, target_style)
                previous_line_was_title = False

            # STANDARD TEXT
            else:
                if previous_line_was_title:
                    add_safe_paragraph(stripped_text, style_map["Title"])
                else:
                    add_safe_paragraph(stripped_text, 'Normal')
                previous_line_was_title = False

        doc.save(output_path)
        return True
    except Exception as e:
        log_func(f"❌ Error building {os.path.basename(output_path)}: {str(e)}")
        return False

# ==========================================
# 🖥️ THE GUI APPLICATION
# ==========================================

class SpecToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Spec Automation Tool")
        self.root.geometry("700x550")
        
        # Style
        style = ttk.Style()
        style.theme_use('clam')

        # Variables
        self.master_folder = tk.StringVar()
        self.template_file = tk.StringVar()
        self.config_file = tk.StringVar()
        self.output_folder = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # --- INPUTS FRAME ---
        input_frame = ttk.LabelFrame(self.root, text="Project Settings", padding="10")
        input_frame.pack(fill="x", padx=10, pady=5)

        # 1. Master Folder
        ttk.Label(input_frame, text="Office Masters Folder:").grid(row=0, column=0, sticky="w")
        ttk.Entry(input_frame, textvariable=self.master_folder, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_master).grid(row=0, column=2)

        # 2. Template File
        ttk.Label(input_frame, text="Architect Template (.docx):").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(input_frame, textvariable=self.template_file, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_template).grid(row=1, column=2)

        # 3. Config File
        ttk.Label(input_frame, text="Project Config (.json):").grid(row=2, column=0, sticky="w")
        ttk.Entry(input_frame, textvariable=self.config_file, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_config).grid(row=2, column=2)

        # 4. Output Folder
        ttk.Label(input_frame, text="Output Folder:").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(input_frame, textvariable=self.output_folder, width=50).grid(row=3, column=1, padx=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_output).grid(row=3, column=2)

        # --- ACTIONS FRAME ---
        action_frame = ttk.Frame(self.root, padding="10")
        action_frame.pack(fill="x", padx=10)

        self.run_btn = ttk.Button(action_frame, text="🚀 PROCESS ALL SPECS", command=self.start_processing)
        self.run_btn.pack(fill="x", pady=5)
        
        self.progress = ttk.Progressbar(action_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x")

        # --- LOG FRAME ---
        log_frame = ttk.LabelFrame(self.root, text="Processing Log", padding="10")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_area = scrolledtext.ScrolledText(log_frame, state='disabled', height=10)
        self.log_area.pack(fill="both", expand=True)

    # --- BROWSE FUNCTIONS ---
    def browse_master(self):
        path = filedialog.askdirectory()
        if path: self.master_folder.set(path)

    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path: self.template_file.set(path)

    def browse_config(self):
        path = filedialog.askopenfilename(filetypes=[("JSON Config", "*.json")])
        if path: self.config_file.set(path)

    def browse_output(self):
        path = filedialog.askdirectory()
        if path: self.output_folder.set(path)

    def log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    # --- PROCESSING LOGIC ---
    def start_processing(self):
        # Validate Inputs
        if not all([self.master_folder.get(), self.template_file.get(), self.output_folder.get()]):
            messagebox.showerror("Error", "Please select Master Folder, Template, and Output Folder.")
            return

        # Run in separate thread to keep GUI responsive
        threading.Thread(target=self.run_batch, daemon=True).start()

    def run_batch(self):
        self.run_btn.config(state="disabled")
        self.log("--- Starting Batch Process ---")
        
        masters_dir = self.master_folder.get()
        template_path = self.template_file.get()
        config_path = self.config_file.get()
        output_dir = self.output_folder.get()

        # Load Config
        config = load_config(config_path)
        if config == DEFAULT_CONFIG:
            self.log("⚠️ No config file loaded (or failed). Using default styles.")
        else:
            self.log("✅ Config loaded successfully.")

        # Find Docx Files
        files = [f for f in os.listdir(masters_dir) if f.endswith(".docx") and not f.startswith("~$")]
        total_files = len(files)
        
        if total_files == 0:
            self.log("❌ No .docx files found in Master folder.")
            self.run_btn.config(state="normal")
            return

        self.progress["maximum"] = total_files
        self.progress["value"] = 0

        for i, filename in enumerate(files):
            master_path = os.path.join(masters_dir, filename)
            # Create a temp markdown path
            temp_md_path = os.path.join(output_dir, f"temp_{filename}.md")
            # Create final output path
            final_docx_path = os.path.join(output_dir, f"Formatted_{filename}")

            self.log(f"Processing: {filename}...")

            # Step 1: Extract
            if extract_master_to_markdown(master_path, temp_md_path, self.log):
                # Step 2: Rebuild
                if rebuild_from_markdown(temp_md_path, template_path, final_docx_path, config, self.log):
                    self.log(f"   ✅ Done.")
                
                # Cleanup Temp
                if os.path.exists(temp_md_path):
                    os.remove(temp_md_path)
            
            self.progress["value"] = i + 1
            self.root.update_idletasks()

        self.log("--- Batch Complete ---")
        messagebox.showinfo("Done", f"Successfully processed {total_files} files.")
        self.run_btn.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    app = SpecToolApp(root)
    root.mainloop()


