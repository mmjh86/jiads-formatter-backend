"""
layout_formatter.py
===================
Core layout formatter.

Supported input:  .docx, .doc (auto-converted), .txt
Supported output: .docx

Pipeline per document:
  1. Load / convert to python-docx Document
  2. Detect paragraph roles (body, abstract, h1/h2/h3, caption, reference)
  3. Apply StyleProfile: margins, columns, fonts, spacing, alignment
  4. Return FormattedResult with the output Document + change log

Install:
    pip install python-docx

For .doc support also install LibreOffice:
    sudo apt install libreoffice   # or brew install libreoffice
"""

import os
import re
import subprocess
import tempfile
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt, RGBColor

from style_profiles import StyleProfile, ParagraphSpec, get_profile


# ─────────────────────────────────────────────────────────────────────
# PARAGRAPH ROLE DETECTION
# ─────────────────────────────────────────────────────────────────────

class ParaRole(str, Enum):
    TITLE     = "title"
    ABSTRACT  = "abstract"
    H1        = "h1"
    H2        = "h2"
    H3        = "h3"
    BODY      = "body"
    CAPTION   = "caption"
    REFERENCE = "reference"
    SKIP      = "skip"    # empty / table / image host — leave untouched


# Heading style name substrings (case-insensitive)
_H1_STYLES = {"heading 1", "h1", "section", "title"}
_H2_STYLES = {"heading 2", "h2", "subsection"}
_H3_STYLES = {"heading 3", "h3", "subsubsection"}

# Caption patterns
_CAPTION_RE = re.compile(
    r"^\s*(fig(ure)?|table|exhibit|appendix)[.\s]?\s*\d*",
    re.IGNORECASE
)

# Reference section header patterns
_REF_HEADER_RE = re.compile(
    r"^\s*(references?|bibliography|works?\s+cited)\s*$",
    re.IGNORECASE
)

# Reference entry patterns — starts with [n] or author-year
_REF_ENTRY_IEEE = re.compile(r"^\s*\[\d+\]")
_REF_ENTRY_APA  = re.compile(r"^\s*[A-Z][a-zA-Zé\-\s]+,\s+[A-Z]")


def detect_role(para, in_refs_section: bool, prev_role: Optional[ParaRole]) -> ParaRole:
    """
    Heuristically assigns a ParaRole to a paragraph.
    Uses style name, text patterns, and position in document.
    """
    text      = para.text.strip()
    style_name= (para.style.name or "").lower()

    # Empty → skip
    if not text:
        return ParaRole.SKIP

    # Explicit style-name matching
    if any(s in style_name for s in _H1_STYLES):
        return ParaRole.H1
    if any(s in style_name for s in _H2_STYLES):
        return ParaRole.H2
    if any(s in style_name for s in _H3_STYLES):
        return ParaRole.H3
    if "abstract" in style_name:
        return ParaRole.ABSTRACT
    if "caption" in style_name:
        return ParaRole.CAPTION

    # Reference section detection
    if _REF_HEADER_RE.match(text):
        return ParaRole.H1   # treat as a heading, signal to caller

    if in_refs_section:
        if _REF_ENTRY_IEEE.match(text) or _REF_ENTRY_APA.match(text):
            return ParaRole.REFERENCE
        if len(text) > 20:   # long paragraph in ref section = ref entry
            return ParaRole.REFERENCE

    # Caption detection from text
    if _CAPTION_RE.match(text):
        return ParaRole.CAPTION

    # Title: first non-empty paragraph, short, no punctuation at end
    if prev_role is None and len(text) < 200 and not text.endswith("."):
        return ParaRole.TITLE

    # Abstract: explicitly labeled or immediately follows title
    if re.match(r"^\s*abstract\s*[:—]?", text, re.IGNORECASE):
        return ParaRole.ABSTRACT
    if prev_role == ParaRole.TITLE and len(text) > 100:
        return ParaRole.ABSTRACT

    return ParaRole.BODY


# ─────────────────────────────────────────────────────────────────────
# ALIGNMENT MAP
# ─────────────────────────────────────────────────────────────────────

_ALIGN_MAP = {
    "left":    WD_ALIGN_PARAGRAPH.LEFT,
    "center":  WD_ALIGN_PARAGRAPH.CENTER,
    "right":   WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


# ─────────────────────────────────────────────────────────────────────
# CHANGE LOG
# ─────────────────────────────────────────────────────────────────────

@dataclass
class Change:
    role:    str
    what:    str   # "font" | "spacing" | "alignment" | "margin" | "columns"
    before:  str
    after:   str
    para_preview: str = ""   # first 60 chars of paragraph text


@dataclass
class FormattedResult:
    document:  Document
    profile:   StyleProfile
    changes:   list[Change]    = field(default_factory=list)
    warnings:  list[str]       = field(default_factory=list)
    abstract_word_count: Optional[int] = None


# ─────────────────────────────────────────────────────────────────────
# DOCUMENT LOADER  (.doc / .docx / .txt)
# ─────────────────────────────────────────────────────────────────────

def load_document(path: str) -> Document:
    """
    Loads any supported file format and returns a python-docx Document.

    .docx → load directly
    .doc  → convert via LibreOffice, then load
    .txt  → create a new Document and populate with plain text paragraphs
    """
    p   = Path(path)
    ext = p.suffix.lower()

    if ext == ".docx":
        return Document(str(p))

    elif ext == ".doc":
        return _convert_doc_to_docx(str(p))

    elif ext == ".txt":
        return _txt_to_docx(str(p))

    else:
        raise ValueError(
            f"Unsupported file format: {ext}. "
            f"Supported: .docx, .doc, .txt"
        )


def _convert_doc_to_docx(doc_path: str) -> Document:
    """
    Converts a legacy .doc file to .docx using LibreOffice headless.
    Requires: sudo apt install libreoffice  OR  brew install libreoffice
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "docx",
             "--outdir", tmpdir, doc_path],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed for {doc_path}:\n{result.stderr}\n"
                "Make sure LibreOffice is installed: sudo apt install libreoffice"
            )
        stem     = Path(doc_path).stem
        out_path = os.path.join(tmpdir, f"{stem}.docx")
        if not os.path.exists(out_path):
            raise RuntimeError(f"Conversion succeeded but output not found at {out_path}")
        return Document(out_path)


def _txt_to_docx(txt_path: str) -> Document:
    """
    Creates a new Document from a plain text file.
    Blank lines → paragraph breaks.
    Lines that look like headings (ALL CAPS, short) → Heading 1 style.
    """
    doc = Document()
    with open(txt_path, encoding="utf-8", errors="replace") as f:
        lines = f.readlines()

    for line in lines:
        stripped = line.rstrip("\n")
        if not stripped.strip():
            doc.add_paragraph("")
            continue
        # Heuristic: short ALL-CAPS line = heading
        if stripped.isupper() and len(stripped) < 80:
            p = doc.add_paragraph(stripped)
            p.style = doc.styles["Heading 1"]
        else:
            doc.add_paragraph(stripped)

    return doc


# ─────────────────────────────────────────────────────────────────────
# PAGE SETUP  (margins + columns)
# ─────────────────────────────────────────────────────────────────────

def _apply_page_setup(doc: Document, profile: StyleProfile,
                      changes: list[Change]):
    """Sets page size, margins, and column layout."""
    for section in doc.sections:
        # Page size
        if profile.page_size == "letter":
            section.page_width  = Cm(21.59)
            section.page_height = Cm(27.94)
        else:  # A4
            section.page_width  = Cm(21.0)
            section.page_height = Cm(29.7)

        m = profile.margins
        old = (f"T{section.top_margin.cm:.2f} "
               f"B{section.bottom_margin.cm:.2f} "
               f"L{section.left_margin.cm:.2f} "
               f"R{section.right_margin.cm:.2f}")

        section.top_margin    = Cm(m.top)
        section.bottom_margin = Cm(m.bottom)
        section.left_margin   = Cm(m.left)
        section.right_margin  = Cm(m.right)
        if m.gutter:
            section.gutter = Cm(m.gutter)

        new = (f"T{m.top:.2f} B{m.bottom:.2f} "
               f"L{m.left:.2f} R{m.right:.2f}")
        changes.append(Change("page", "margin", old, new))

        # Column layout
        if profile.columns.count > 1:
            _set_columns(section, profile.columns.count,
                         int(Cm(profile.columns.space)))
            changes.append(Change("page", "columns", "1",
                                  str(profile.columns.count)))


def _set_columns(section, count: int, space_dxa: int):
    """
    Adds multi-column layout to a section via direct XML manipulation.
    python-docx does not expose this natively.
    """
    sectPr = section._sectPr
    # Remove existing cols element if present
    existing = sectPr.find(qn("w:cols"))
    if existing is not None:
        sectPr.remove(existing)

    cols_el = OxmlElement("w:cols")
    cols_el.set(qn("w:num"),    str(count))
    cols_el.set(qn("w:space"),  str(space_dxa))
    cols_el.set(qn("w:equalWidth"), "1")
    sectPr.append(cols_el)


# ─────────────────────────────────────────────────────────────────────
# PARAGRAPH FORMATTER
# ─────────────────────────────────────────────────────────────────────

def _apply_para_spec(para, spec: ParagraphSpec,
                     role: str, changes: list[Change]):
    """
    Applies a ParagraphSpec to a single paragraph:
    font, size, bold, italic, alignment, spacing, first-line indent.
    Records each change.
    """
    preview = para.text[:60].replace("\n", " ")

    # ── Alignment ─────────────────────────────────────────────────────
    old_align = str(para.alignment)
    new_align = _ALIGN_MAP.get(spec.alignment, WD_ALIGN_PARAGRAPH.LEFT)
    if para.alignment != new_align:
        para.alignment = new_align
        changes.append(Change(role, "alignment", old_align,
                              spec.alignment, preview))

    # ── Paragraph spacing ─────────────────────────────────────────────
    pf = para.paragraph_format
    old_sp = (f"before={_pt(pf.space_before)} "
              f"after={_pt(pf.space_after)} "
              f"line={pf.line_spacing}")
    pf.space_before = Pt(spec.spacing.before)
    pf.space_after  = Pt(spec.spacing.after)

    if spec.spacing.line_rule == "exact":
        pf.line_spacing       = Pt(spec.spacing.line)
    else:
        # Multiply — python-docx accepts float for proportional
        pf.line_spacing       = spec.spacing.line
        pf.line_spacing_rule  = None  # auto/proportional

    new_sp = (f"before={spec.spacing.before} "
              f"after={spec.spacing.after} "
              f"line={spec.spacing.line}")
    if old_sp != new_sp:
        changes.append(Change(role, "spacing", old_sp, new_sp, preview))

    # ── First-line indent ─────────────────────────────────────────────
    if spec.first_line_indent != 0.0:
        pf.first_line_indent = Cm(spec.first_line_indent)
    else:
        pf.first_line_indent = Pt(0)

    # ── Font on every run ─────────────────────────────────────────────
    runs_changed = False
    for run in para.runs:
        old_fn   = run.font.name or ""
        old_size = run.font.size

        run.font.name  = spec.font.name
        run.font.size  = Pt(spec.font.size)
        run.font.bold  = spec.font.bold
        run.font.italic= spec.font.italic

        if old_fn != spec.font.name or old_size != Pt(spec.font.size):
            runs_changed = True

    if runs_changed:
        changes.append(Change(
            role, "font",
            f"mixed",
            f"{spec.font.name} {spec.font.size}pt "
            f"{'bold ' if spec.font.bold else ''}"
            f"{'italic' if spec.font.italic else ''}".strip(),
            preview
        ))

    # ── Direct paragraph XML font (covers paragraphs with no runs) ────
    pPr = para._p.get_or_add_pPr()
    rPr = _get_or_create_pPr_rPr(pPr)
    _set_xml_font(rPr, spec)


def _get_or_create_pPr_rPr(pPr):
    """
    Finds or creates a w:rPr element inside a w:pPr element.
    python-docx CT_PPr does not expose get_or_add_rPr(), so we do it manually.
    """
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    return rPr


def _set_xml_font(rPr, spec: ParagraphSpec):
    """Sets font in pPr/rPr XML so empty paragraphs also get styled."""
    # Font name
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"),    spec.font.name)
    rFonts.set(qn("w:hAnsi"),    spec.font.name)
    rFonts.set(qn("w:cs"),       spec.font.name)

    # Size (sz = half-points)
    half_pts = str(int(spec.font.size * 2))
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), half_pts)

    # Bold
    b_el = rPr.find(qn("w:b"))
    if spec.font.bold:
        if b_el is None:
            rPr.append(OxmlElement("w:b"))
    else:
        if b_el is not None:
            rPr.remove(b_el)


def _pt(val) -> str:
    """Safe string conversion of a docx length value."""
    if val is None: return "none"
    try:    return f"{val.pt:.1f}"
    except: return str(val)


# ─────────────────────────────────────────────────────────────────────
# ROLE → SPEC MAP
# ─────────────────────────────────────────────────────────────────────

def _spec_for_role(role: ParaRole, profile: StyleProfile) -> Optional[ParagraphSpec]:
    return {
        ParaRole.TITLE:     profile.h1,       # title treated as H1 visually
        ParaRole.ABSTRACT:  profile.abstract,
        ParaRole.H1:        profile.h1,
        ParaRole.H2:        profile.h2,
        ParaRole.H3:        profile.h3,
        ParaRole.BODY:      profile.body,
        ParaRole.CAPTION:   profile.caption,
        ParaRole.REFERENCE: profile.references,
        ParaRole.SKIP:      None,
    }.get(role)


# ─────────────────────────────────────────────────────────────────────
# MAIN FORMATTER
# ─────────────────────────────────────────────────────────────────────

def format_document(
    input_path:  str,
    style:       str,
    output_path: Optional[str] = None,
) -> FormattedResult:
    """
    Loads a .doc/.docx/.txt file, applies the named style profile,
    saves to output_path (or auto-named), and returns a FormattedResult.

    Args:
        input_path:  path to source document
        style:       profile name — "ieee" | "apa7" | "harvard" | custom
        output_path: where to save result (default: <stem>_<style>.docx)

    Returns:
        FormattedResult with .document, .profile, .changes, .warnings
    """
    profile  = get_profile(style)
    doc      = load_document(input_path)
    changes: list[Change]  = []
    warnings: list[str]    = []

    # ── Page setup ────────────────────────────────────────────────────
    _apply_page_setup(doc, profile, changes)

    # ── Paragraph formatting ──────────────────────────────────────────
    in_refs   = False
    prev_role = None
    abstract_words = None

    for para in doc.paragraphs:
        role = detect_role(para, in_refs, prev_role)

        # Track when we enter the references section
        if _REF_HEADER_RE.match(para.text.strip()):
            in_refs = True

        spec = _spec_for_role(role, profile)
        if spec is None:
            prev_role = role
            continue

        _apply_para_spec(para, spec, role.value, changes)

        # Word count abstract
        if role == ParaRole.ABSTRACT:
            wc = len(para.text.split())
            abstract_words = (abstract_words or 0) + wc

        prev_role = role

    # ── Abstract word limit warning ───────────────────────────────────
    if abstract_words and profile.abstract_word_limit:
        if abstract_words > profile.abstract_word_limit:
            warnings.append(
                f"Abstract is {abstract_words} words "
                f"(limit: {profile.abstract_word_limit}). "
                "Please shorten before submission."
            )

    # ── Tables: apply body font to all table cells ────────────────────
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _apply_para_spec(para, profile.body,
                                     "table_cell", changes)

    # ── Save ──────────────────────────────────────────────────────────
    if output_path is None:
        stem        = Path(input_path).stem
        output_path = str(Path(input_path).parent / f"{stem}_{style}.docx")

    doc.save(output_path)

    return FormattedResult(
        document            = doc,
        profile             = profile,
        changes             = changes,
        warnings            = warnings,
        abstract_word_count = abstract_words,
    )


# ─────────────────────────────────────────────────────────────────────
# BATCH FORMATTER
# ─────────────────────────────────────────────────────────────────────

def format_batch(
    input_paths: list[str],
    style:       str,
    output_dir:  Optional[str] = None,
) -> list[FormattedResult]:
    """
    Format multiple documents with the same style profile.
    Errors on individual files are caught and added as warnings.
    """
    results = []
    for path in input_paths:
        try:
            out = None
            if output_dir:
                stem = Path(path).stem
                out  = os.path.join(output_dir, f"{stem}_{style}.docx")
            results.append(format_document(path, style, out))
        except Exception as e:
            # Return a partial result with the error as a warning
            results.append(FormattedResult(
                document = Document(),
                profile  = get_profile(style),
                warnings = [f"Failed to format {path}: {e}"],
            ))
    return results


# ─────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys, json

    if len(sys.argv) < 3:
        print("Usage: python layout_formatter.py <input_file> <style> [output_file]")
        print(f"Styles: ieee, apa7, harvard")
        sys.exit(1)

    inp   = sys.argv[1]
    style = sys.argv[2]
    out   = sys.argv[3] if len(sys.argv) > 3 else None

    print(f"Formatting {inp} → {style}…")
    result = format_document(inp, style, out)

    print(f"\n✅ Saved to: {out or inp}")
    print(f"   Changes applied: {len(result.changes)}")
    if result.warnings:
        print(f"   ⚠️  Warnings:")
        for w in result.warnings:
            print(f"      · {w}")

    print("\n── Change log ──")
    for c in result.changes[:20]:
        print(f"  [{c.role:12s}] {c.what:10s}  {c.before!r:30s} → {c.after!r}")
    if len(result.changes) > 20:
        print(f"  … and {len(result.changes)-20} more changes")
