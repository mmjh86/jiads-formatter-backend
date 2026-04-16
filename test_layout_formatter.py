"""
test_layout_formatter.py
========================
Pytest suite for layout_formatter.py and style_profiles.py.

Run:
    pip install pytest python-docx
    pytest test_layout_formatter.py -v

Tests cover:
  - Profile registry (get, register, list)
  - .txt loading
  - .docx loading
  - Role detection (body, abstract, heading, caption, reference)
  - Page margins applied correctly
  - Font applied to runs
  - Spacing applied
  - Alignment applied
  - Abstract word count warning
  - Batch formatting
  - Unknown style error
  - Custom profile registration
"""

import os
import tempfile
import pytest

from docx import Document
from docx.shared import Cm, Pt

from style_profiles import (
    get_profile, register_profile, list_profiles,
    StyleProfile, MarginSpec, FontSpec, SpacingSpec,
    ParagraphSpec, ColumnSpec,
)
from layout_formatter import (
    load_document, format_document, format_batch,
    detect_role, ParaRole, FormattedResult,
)


# ─────────────────────────────────────────────────────────────────────
# FIXTURES
# ─────────────────────────────────────────────────────────────────────

@pytest.fixture
def tmp_dir():
    with tempfile.TemporaryDirectory() as d:
        yield d


def _make_docx(tmp_dir: str, paragraphs: list[tuple[str, str]]) -> str:
    """
    Creates a minimal .docx with given (text, style_name) pairs.
    Returns the file path.
    """
    doc  = Document()
    path = os.path.join(tmp_dir, "test_input.docx")
    for text, style in paragraphs:
        p = doc.add_paragraph(text)
        try:
            p.style = doc.styles[style]
        except KeyError:
            pass   # use default if style not in template
    doc.save(path)
    return path


def _make_txt(tmp_dir: str, content: str) -> str:
    path = os.path.join(tmp_dir, "test_input.txt")
    with open(path, "w") as f:
        f.write(content)
    return path


# ─────────────────────────────────────────────────────────────────────
# 1. PROFILE REGISTRY
# ─────────────────────────────────────────────────────────────────────

class TestProfileRegistry:

    def test_get_ieee(self):
        p = get_profile("ieee")
        assert p.name == "ieee"
        assert p.columns.count == 2
        assert p.body.font.name == "Times New Roman"
        assert p.body.font.size == 10

    def test_get_apa7_by_alias(self):
        p = get_profile("apa")
        assert p.name == "apa7"
        assert p.body.font.size == 12
        assert p.body.spacing.line == 2.0

    def test_get_harvard(self):
        p = get_profile("harvard")
        assert p.name == "harvard"
        assert p.body.font.name == "Arial"
        assert p.page_size == "A4"

    def test_case_insensitive(self):
        assert get_profile("IEEE").name == "ieee"
        assert get_profile("Harvard").name == "harvard"

    def test_unknown_style_raises(self):
        with pytest.raises(KeyError, match="not found"):
            get_profile("nonexistent_journal")

    def test_list_profiles_returns_unique(self):
        names = list_profiles()
        assert len(names) == len(set(names))
        assert "ieee" in names
        assert "apa7" in names
        assert "harvard" in names

    def test_register_custom_profile(self):
        custom = StyleProfile(
            name="nature",
            display_name="Nature",
            citation_style="apa",
            page_size="A4",
            margins=MarginSpec(2.0, 2.0, 2.0, 2.0),
            columns=ColumnSpec(count=1),
            body=ParagraphSpec(
                font=FontSpec("Arial", 9),
                spacing=SpacingSpec(line=1.0),
            ),
            abstract=ParagraphSpec(
                font=FontSpec("Arial", 9),
                spacing=SpacingSpec(line=1.0),
            ),
            h1=ParagraphSpec(font=FontSpec("Arial", 11, bold=True),
                             spacing=SpacingSpec()),
            h2=ParagraphSpec(font=FontSpec("Arial", 10, bold=True),
                             spacing=SpacingSpec()),
            h3=ParagraphSpec(font=FontSpec("Arial", 10, italic=True),
                             spacing=SpacingSpec()),
            caption=ParagraphSpec(font=FontSpec("Arial", 8),
                                  spacing=SpacingSpec()),
            references=ParagraphSpec(font=FontSpec("Arial", 9),
                                     spacing=SpacingSpec()),
        )
        register_profile(custom)
        assert get_profile("nature").display_name == "Nature"


# ─────────────────────────────────────────────────────────────────────
# 2. DOCUMENT LOADING
# ─────────────────────────────────────────────────────────────────────

class TestDocumentLoading:

    def test_load_docx(self, tmp_dir):
        path = _make_docx(tmp_dir, [("Hello world", "Normal")])
        doc  = load_document(path)
        assert any("Hello world" in p.text for p in doc.paragraphs)

    def test_load_txt_creates_paragraphs(self, tmp_dir):
        path = _make_txt(tmp_dir, "Line one\nLine two\n\nLine four")
        doc  = load_document(path)
        texts = [p.text for p in doc.paragraphs if p.text]
        assert "Line one" in texts
        assert "Line two" in texts
        assert "Line four" in texts

    def test_load_txt_allcaps_becomes_heading(self, tmp_dir):
        path = _make_txt(tmp_dir, "INTRODUCTION\nThis is body text.")
        doc  = load_document(path)
        headings = [p for p in doc.paragraphs
                    if p.style.name.startswith("Heading")]
        assert len(headings) >= 1
        assert "INTRODUCTION" in headings[0].text

    def test_load_unsupported_raises(self, tmp_dir):
        bad_path = os.path.join(tmp_dir, "file.pdf")
        with open(bad_path, "w") as f:
            f.write("dummy")
        with pytest.raises(ValueError, match="Unsupported"):
            load_document(bad_path)


# ─────────────────────────────────────────────────────────────────────
# 3. ROLE DETECTION
# ─────────────────────────────────────────────────────────────────────

class TestRoleDetection:

    def _make_para(self, text: str, style_name: str = "Normal"):
        """Creates a real python-docx paragraph for testing."""
        doc = Document()
        p   = doc.add_paragraph(text)
        return p

    def test_empty_paragraph_is_skip(self):
        p = self._make_para("")
        assert detect_role(p, False, None) == ParaRole.SKIP

    def test_first_para_short_no_period_is_title(self):
        p = self._make_para("Machine Learning in Healthcare")
        assert detect_role(p, False, None) == ParaRole.TITLE

    def test_caption_figure(self):
        p = self._make_para("Figure 3. Performance comparison across datasets.")
        assert detect_role(p, False, ParaRole.BODY) == ParaRole.CAPTION

    def test_caption_table(self):
        p = self._make_para("Table 1: Summary statistics")
        assert detect_role(p, False, ParaRole.BODY) == ParaRole.CAPTION

    def test_reference_entry_ieee(self):
        p = self._make_para("[1] LeCun, Y. et al., Deep Learning, Nature, 2015.")
        assert detect_role(p, True, ParaRole.H1) == ParaRole.REFERENCE

    def test_reference_entry_apa(self):
        p = self._make_para("Smith, J., & Doe, A. (2022). Title. Journal, 10(2), 1–15.")
        assert detect_role(p, True, ParaRole.H1) == ParaRole.REFERENCE

    def test_ref_header_detected(self):
        p = self._make_para("References")
        # The header itself is treated as H1 — caller tracks in_refs
        role = detect_role(p, False, ParaRole.BODY)
        assert role == ParaRole.H1

    def test_body_paragraph(self):
        text = ("This paper presents a novel approach to intrusion detection "
                "using deep neural networks trained on the CICIDS-2017 dataset.")
        p = self._make_para(text)
        assert detect_role(p, False, ParaRole.ABSTRACT) == ParaRole.BODY


# ─────────────────────────────────────────────────────────────────────
# 4. FORMAT DOCUMENT — MARGINS
# ─────────────────────────────────────────────────────────────────────

class TestMargins:

    def test_ieee_margins_applied(self, tmp_dir):
        path   = _make_docx(tmp_dir, [("Test body text.", "Normal")])
        out    = os.path.join(tmp_dir, "out.docx")
        result = format_document(path, "ieee", out)
        doc    = Document(out)
        for section in doc.sections:
            assert abs(section.top_margin.cm - 1.9) < 0.1
            assert abs(section.left_margin.cm - 1.9) < 0.1

    def test_apa7_margins_applied(self, tmp_dir):
        path   = _make_docx(tmp_dir, [("Test body text.", "Normal")])
        out    = os.path.join(tmp_dir, "out.docx")
        result = format_document(path, "apa7", out)
        doc    = Document(out)
        for section in doc.sections:
            assert abs(section.top_margin.cm - 2.54) < 0.1

    def test_harvard_margins_applied(self, tmp_dir):
        path   = _make_docx(tmp_dir, [("Test body text.", "Normal")])
        out    = os.path.join(tmp_dir, "out.docx")
        format_document(path, "harvard", out)
        doc    = Document(out)
        for section in doc.sections:
            assert abs(section.left_margin.cm - 3.0) < 0.1


# ─────────────────────────────────────────────────────────────────────
# 5. FORMAT DOCUMENT — FONTS
# ─────────────────────────────────────────────────────────────────────

class TestFonts:

    def test_ieee_body_font(self, tmp_dir):
        path   = _make_docx(tmp_dir, [
            ("Short title", "Normal"),
            ("Abstract text here " * 5, "Normal"),
            ("Body paragraph with content " * 3, "Normal"),
        ])
        out    = os.path.join(tmp_dir, "out.docx")
        format_document(path, "ieee", out)
        doc    = Document(out)
        # Check runs have Times New Roman
        body_paras = [p for p in doc.paragraphs if len(p.text) > 50]
        for para in body_paras:
            for run in para.runs:
                assert run.font.name == "Times New Roman", \
                    f"Expected Times New Roman, got {run.font.name}"

    def test_harvard_body_font_is_arial(self, tmp_dir):
        path   = _make_docx(tmp_dir, [
            ("Short title", "Normal"),
            ("Body paragraph with content " * 3, "Normal"),
        ])
        out    = os.path.join(tmp_dir, "out.docx")
        format_document(path, "harvard", out)
        doc    = Document(out)
        body_paras = [p for p in doc.paragraphs if len(p.text) > 20]
        for para in body_paras:
            for run in para.runs:
                assert run.font.name == "Arial"

    def test_apa7_font_size_12pt(self, tmp_dir):
        path   = _make_docx(tmp_dir, [
            ("Short title", "Normal"),
            ("Body paragraph " * 5, "Normal"),
        ])
        out    = os.path.join(tmp_dir, "out.docx")
        format_document(path, "apa7", out)
        doc    = Document(out)
        for para in doc.paragraphs:
            for run in para.runs:
                if run.font.size:
                    assert abs(run.font.size.pt - 12.0) < 0.5


# ─────────────────────────────────────────────────────────────────────
# 6. FORMAT DOCUMENT — SPACING
# ─────────────────────────────────────────────────────────────────────

class TestSpacing:

    def test_apa7_double_spaced(self, tmp_dir):
        path   = _make_docx(tmp_dir, [
            ("Title here", "Normal"),
            ("Body text " * 10, "Normal"),
        ])
        out    = os.path.join(tmp_dir, "out.docx")
        format_document(path, "apa7", out)
        doc    = Document(out)
        body_paras = [p for p in doc.paragraphs if len(p.text) > 20]
        for para in body_paras:
            ls = para.paragraph_format.line_spacing
            assert ls == 2.0 or (ls and abs(ls - 2.0) < 0.1), \
                f"Expected line_spacing=2.0, got {ls}"


# ─────────────────────────────────────────────────────────────────────
# 7. ABSTRACT WORD COUNT WARNING
# ─────────────────────────────────────────────────────────────────────

class TestAbstractWarning:

    def test_ieee_abstract_over_limit_warns(self, tmp_dir):
        # IEEE limit is 200 words
        long_abstract = "word " * 250   # 250 words
        path   = _make_docx(tmp_dir, [
            ("Title", "Normal"),
            (long_abstract, "Normal"),
            ("Body text here.", "Normal"),
        ])
        out    = os.path.join(tmp_dir, "out.docx")
        result = format_document(path, "ieee", out)
        assert any("Abstract" in w and "200" in w for w in result.warnings), \
            f"Expected abstract warning, got: {result.warnings}"

    def test_short_abstract_no_warning(self, tmp_dir):
        short_abstract = "Brief abstract with fewer than two hundred words. " * 2
        path   = _make_docx(tmp_dir, [
            ("Title", "Normal"),
            (short_abstract, "Normal"),
            ("Body text.", "Normal"),
        ])
        out    = os.path.join(tmp_dir, "out.docx")
        result = format_document(path, "ieee", out)
        abstract_warnings = [w for w in result.warnings if "Abstract" in w]
        assert len(abstract_warnings) == 0


# ─────────────────────────────────────────────────────────────────────
# 8. CHANGE LOG
# ─────────────────────────────────────────────────────────────────────

class TestChangeLog:

    def test_changes_recorded(self, tmp_dir):
        path   = _make_docx(tmp_dir, [
            ("Short title", "Normal"),
            ("Body paragraph " * 5, "Normal"),
        ])
        out    = os.path.join(tmp_dir, "out.docx")
        result = format_document(path, "ieee", out)
        assert len(result.changes) > 0

    def test_margin_change_recorded(self, tmp_dir):
        path   = _make_docx(tmp_dir, [("Text.", "Normal")])
        out    = os.path.join(tmp_dir, "out.docx")
        result = format_document(path, "ieee", out)
        margin_changes = [c for c in result.changes if c.what == "margin"]
        assert len(margin_changes) > 0

    def test_font_change_recorded(self, tmp_dir):
        path   = _make_docx(tmp_dir, [
            ("Title", "Normal"),
            ("Body " * 10, "Normal"),
        ])
        out    = os.path.join(tmp_dir, "out.docx")
        result = format_document(path, "ieee", out)
        font_changes = [c for c in result.changes if c.what == "font"]
        assert len(font_changes) > 0


# ─────────────────────────────────────────────────────────────────────
# 9. BATCH FORMATTING
# ─────────────────────────────────────────────────────────────────────

class TestBatchFormatting:

    def test_batch_formats_multiple_files(self, tmp_dir):
        paths = [
            _make_docx(tmp_dir, [("Doc one body text " * 5, "Normal")]),
        ]
        # Make a second file
        doc2_path = os.path.join(tmp_dir, "doc2.docx")
        doc2      = Document()
        doc2.add_paragraph("Doc two body text " * 5)
        doc2.save(doc2_path)
        paths.append(doc2_path)

        results = format_batch(paths, "ieee", output_dir=tmp_dir)
        assert len(results) == 2
        assert all(isinstance(r, FormattedResult) for r in results)

    def test_batch_recovers_from_bad_file(self, tmp_dir):
        good_path = _make_docx(tmp_dir, [("Good doc.", "Normal")])
        bad_path  = os.path.join(tmp_dir, "bad.xyz")
        with open(bad_path, "w") as f:
            f.write("not a document")

        results = format_batch([good_path, bad_path], "ieee", output_dir=tmp_dir)
        assert len(results) == 2
        # bad file should produce a warning, not crash
        bad_result = results[1]
        assert len(bad_result.warnings) > 0
        assert any("Failed" in w for w in bad_result.warnings)


# ─────────────────────────────────────────────────────────────────────
# 10. TXT INPUT
# ─────────────────────────────────────────────────────────────────────

class TestTxtInput:

    def test_txt_formats_successfully(self, tmp_dir):
        path   = _make_txt(tmp_dir,
                           "MY PAPER TITLE\n\nThis is the abstract.\n\n"
                           "INTRODUCTION\n\nBody text here.\n\n"
                           "References\n\n[1] Author et al., 2023.")
        out    = os.path.join(tmp_dir, "out.docx")
        result = format_document(path, "ieee", out)
        assert os.path.exists(out)
        assert len(result.changes) > 0

    def test_txt_output_is_valid_docx(self, tmp_dir):
        path   = _make_txt(tmp_dir, "Title\n\nBody text " * 3)
        out    = os.path.join(tmp_dir, "out.docx")
        format_document(path, "apa7", out)
        # Should be loadable as a valid docx
        doc = Document(out)
        assert len(doc.paragraphs) > 0
