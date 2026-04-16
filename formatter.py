# formatter.py
import os
import re
import shutil
import zipfile
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


def format_jiads_manuscript(
    input_path,
    output_path,
    anthropic_api_key=None,
    volume="X",
    issue="Y",
    year="2024",
):
    # ── Step 1: Load manuscript ───────────────────────────────────────
    doc = Document(input_path)

    # ── Step 2: Page setup ────────────────────────────────────────────
    for section in doc.sections:
        section.page_width    = Cm(21.59)
        section.page_height   = Cm(27.94)
        section.top_margin    = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin   = Cm(3.17)
        section.right_margin  = Cm(3.17)

    # ── Step 3: Fix paragraph styles ─────────────────────────────────
    HEADING_RE = re.compile(
        r"^(\d+\.?[\d.]* .{2,60}|Abstract|References|Keywords|"
        r"Acknowledgment|Introduction|Conclusion|Discussion)$",
        re.IGNORECASE,
    )

    for i, para in enumerate(doc.paragraphs):
        text  = para.text.strip()
        style = para.style.name
        if not text:
            continue

        # Title — first short non-reference paragraph
        if i <= 3 and 10 < len(text) < 200 and not text.startswith("["):
            para.style     = doc.styles["Normal"]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before      = Pt(0)
            para.paragraph_format.space_after       = Pt(12)
            para.paragraph_format.line_spacing      = 1.0
            para.paragraph_format.first_line_indent = Pt(0)
            for run in para.runs:
                run.bold       = False
                run.font.name  = "Times New Roman"
                run.font.size  = Pt(14)
            continue

        # Body text wrongly styled as a heading
        if style in ("Heading 1", "Heading 2") and len(text) > 80 and not HEADING_RE.match(text):
            para.style     = doc.styles["Normal"]
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            para.paragraph_format.space_before      = Pt(0)
            para.paragraph_format.space_after       = Pt(0)
            para.paragraph_format.line_spacing      = 1.0
            para.paragraph_format.first_line_indent = Pt(0)
            for run in para.runs:
                run.bold       = False
                run.font.name  = "Times New Roman"
                run.font.size  = Pt(11)
            continue

        # Real section headings (H1)
        if style == "Heading 1" and HEADING_RE.match(text):
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            para.paragraph_format.space_before = Pt(12)
            para.paragraph_format.space_after  = Pt(6)
            for run in para.runs:
                run.bold       = True
                run.font.name  = "Times New Roman"
                run.font.size  = Pt(13)
            continue

        # Real subsection headings (H2)
        if style == "Heading 2" and len(text) <= 100:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            para.paragraph_format.space_before = Pt(10)
            para.paragraph_format.space_after  = Pt(4)
            for run in para.runs:
                run.bold       = True
                run.font.name  = "Times New Roman"
                run.font.size  = Pt(12)
            continue

        # Normal body paragraphs
        if style == "Normal":
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            para.paragraph_format.space_before      = Pt(0)
            para.paragraph_format.space_after       = Pt(0)
            para.paragraph_format.line_spacing      = 1.0
            para.paragraph_format.first_line_indent = Pt(0)
            for run in para.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(11)
                if run.bold:
                    run.bold = False

    # ── Step 4: Save formatted doc to temp file ───────────────────────
    temp_file = output_path + ".tmp.docx"
    doc.save(temp_file)

    # ── Step 5: Inject JIADS logo + journal title ─────────────────────
    # Try to find the template file in common locations
    template_candidates = [
        "JIADS_Article_Template_FINAL (2).docx",
        "JIADS_Article_Template_FINAL.docx",
        "templates/JIADS_Article_Template_FINAL (2).docx",
        "templates/JIADS_Article_Template_FINAL.docx",
        os.path.join(os.path.dirname(__file__), "JIADS_Article_Template_FINAL (2).docx"),
        os.path.join(os.path.dirname(__file__), "templates", "JIADS_Article_Template_FINAL.docx"),
    ]
    template_path = next((p for p in template_candidates if os.path.exists(p)), None)

    if template_path:
        try:
            _inject_jiads_header(temp_file, output_path, template_path, volume, issue, year)
            os.remove(temp_file)
        except Exception as e:
            # Header injection failed — use formatted doc without logo
            print(f"Warning: logo injection failed ({e}), saving without logo")
            shutil.move(temp_file, output_path)
    else:
        # No template found — save without logo, add plain text header
        print("Warning: JIADS template not found, saving without logo")
        _inject_text_header(temp_file, output_path, volume, issue, year)
        if os.path.exists(temp_file):
            os.remove(temp_file)

    # ── Step 6: Citation conversion (optional) ────────────────────────
    if anthropic_api_key:
        try:
            import anthropic
            from citations import reformat_citations, detect_style

            client    = anthropic.Anthropic(api_key=anthropic_api_key)
            final_doc = Document(output_path)
            style     = detect_style(final_doc)
            if style.name != "ieee":
                result = reformat_citations(final_doc, "ieee", client)
                final_doc.save(output_path)
                print(f"Citations converted: {result.changed_count} references")
            else:
                print("Citations already IEEE — no conversion needed")
        except ImportError:
            print("citations.py not found — skipping citation conversion")
        except Exception as e:
            print(f"Citation conversion failed: {e} — continuing without it")

    return output_path


# ─────────────────────────────────────────────────────────────────────
# LOGO INJECTION  (tested and confirmed working)
# ─────────────────────────────────────────────────────────────────────

def _inject_jiads_header(temp_file, output_path, template_path, volume, issue, year):
    """
    Copies the JIADS logo image + journal title paragraph from the
    template into the manuscript via ZIP manipulation.
    Confirmed working — produces valid .docx that opens in Word.
    """
    # Read template assets
    with zipfile.ZipFile(template_path, "r") as tz:
        tmpl_doc_xml = tz.read("word/document.xml").decode("utf-8")
        logo_bytes   = tz.read("word/media/image1.png")

    # Extract logo paragraph (identified by its paraId in the template)
    logo_match = re.search(
        r"(<w:p [^>]*099BD9BF[^>]*>.*?</w:p>)",
        tmpl_doc_xml, re.DOTALL
    )
    if not logo_match:
        raise ValueError("Logo paragraph not found in template")
    logo_para_xml = logo_match.group(1)

    # Extract journal title paragraph
    journal_match = re.search(
        r"(<w:p\b[^>]*>(?:(?!</w:p>).)*Journal of Informatics(?:(?!</w:p>).)*</w:p>)",
        tmpl_doc_xml, re.DOTALL
    )
    if not journal_match:
        raise ValueError("Journal title paragraph not found in template")
    journal_para_xml = journal_match.group(1)

    # Update volume/issue/year in journal title
    journal_para_xml = re.sub(r"\d+\(\d+\)", f"{volume}({issue})", journal_para_xml)
    journal_para_xml = re.sub(r"\b20\d{2}\b", year, journal_para_xml)

    # Read manuscript zip
    with zipfile.ZipFile(temp_file, "r") as zin:
        ms_doc_xml    = zin.read("word/document.xml").decode("utf-8")
        ms_rels_xml   = zin.read("word/_rels/document.xml.rels").decode("utf-8")
        content_types = zin.read("[Content_Types].xml").decode("utf-8")
        all_files     = {name: zin.read(name) for name in zin.namelist()}

    # Add logo image
    logo_filename = "jiads_logo.png"
    new_rid       = "rId_logo1"
    all_files[f"word/media/{logo_filename}"] = logo_bytes

    # Register content type for image
    if logo_filename not in content_types:
        content_types = content_types.replace(
            "</Types>",
            f'  <Override PartName="/word/media/{logo_filename}" ContentType="image/png"/>\n</Types>',
        )

    # Add image relationship
    new_rel = (
        f'<Relationship Id="{new_rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="media/{logo_filename}"/>'
    )
    ms_rels_xml = ms_rels_xml.replace(
        "</Relationships>", f"  {new_rel}\n</Relationships>"
    )

    # Remap rId in logo paragraph to new relationship id
    logo_para_updated = logo_para_xml.replace('r:embed="rId8"', f'r:embed="{new_rid}"')

    # Prepend logo + journal title to document body
    header_block   = logo_para_updated + "\n" + journal_para_xml
    ms_doc_updated = ms_doc_xml.replace("<w:body>", f"<w:body>\n{header_block}\n", 1)

    # Write output zip
    all_files["word/document.xml"]            = ms_doc_updated.encode("utf-8")
    all_files["word/_rels/document.xml.rels"] = ms_rels_xml.encode("utf-8")
    all_files["[Content_Types].xml"]          = content_types.encode("utf-8")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)


def _inject_text_header(temp_file, output_path, volume, issue, year):
    """
    Fallback when template is not available.
    Adds a plain text journal title header without the logo.
    """
    with zipfile.ZipFile(temp_file, "r") as zin:
        ms_doc_xml = zin.read("word/document.xml").decode("utf-8")
        all_files  = {name: zin.read(name) for name in zin.namelist()}

    header_xml = (
        "<w:p>"
        "<w:pPr><w:jc w:val=\"center\"/></w:pPr>"
        "<w:r>"
        "<w:rPr><w:b/><w:sz w:val=\"20\"/></w:rPr>"
        f"<w:t>Journal of Informatics and Advanced Data Science | "
        f"Vol. {volume}, Issue {issue}, {year}</w:t>"
        "</w:r>"
        "</w:p>"
    )

    ms_doc_updated = ms_doc_xml.replace("<w:body>", f"<w:body>\n{header_xml}\n", 1)
    all_files["word/document.xml"] = ms_doc_updated.encode("utf-8")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)
