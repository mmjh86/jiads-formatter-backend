# formatter.py
import os
import re
import shutil
import zipfile
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

def format_jiads_manuscript(input_path, output_path, anthropic_api_key=None):
    # Step 1: Load manuscript
    doc = Document(input_path)

    # Step 2: Page setup (US Letter, margins)
    for section in doc.sections:
        section.page_width = Cm(21.59)
        section.page_height = Cm(27.94)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)

    # Step 3: Fix paragraph styles (your logic from notebook)
    HEADING_RE = re.compile(r'^(\d+\.?[\d.]* .{2,60}|Abstract|References|Keywords|Acknowledgment|Introduction|Conclusion|Discussion)$', re.IGNORECASE)
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        # Title detection (first short paragraph)
        if i <= 3 and len(text) < 200 and len(text) > 10:
            para.style = doc.styles['Normal']
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = False
                run.font.name = 'Times New Roman'
                run.font.size = Pt(14)
            continue
        # Fix misclassified body paragraphs
        if para.style.name in ('Heading 1', 'Heading 2') and len(text) > 80 and not HEADING_RE.match(text):
            para.style = doc.styles['Normal']
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for run in para.runs:
                run.bold = False
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
        # Real headings
        elif para.style.name == 'Heading 1' and HEADING_RE.match(text):
            for run in para.runs:
                run.bold = True
                run.font.name = 'Times New Roman'
                run.font.size = Pt(13)
        # Normal body
        elif para.style.name == 'Normal':
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for run in para.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                run.bold = False

    # Save intermediate
    temp_file = "_temp_formatted.docx"
    doc.save(temp_file)

    # Step 4: Inject logo and journal title from template (using your working method)
    TEMPLATE_PATH = "templates/JIADS_Article_Template_FINAL.docx"
    shutil.copy2(TEMP_FILE, output_path)  # start with the temp file
    # ... (insert your header injection code – the one that finally worked)
    # For brevity, I'll assume you have a function `inject_header(temp_file, output_path, volume, issue, year)`
    # You can reuse the exact code from your notebook that produced a non‑corrupt file.

    # Step 5: Citation conversion (optional)
    if anthropic_api_key:
        import anthropic
        from citations import reformat_citations, detect_style
        client = anthropic.Anthropic(api_key=anthropic_api_key)
        final_doc = Document(output_path)
        style = detect_style(final_doc)
        if style.name != 'ieee':
            result = reformat_citations(final_doc, 'ieee', client)
            final_doc.save(output_path)

    # Cleanup
    if os.path.exists(temp_file):
        os.remove(temp_file)

    return output_path