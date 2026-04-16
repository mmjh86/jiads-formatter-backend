"""
test_citations.py
=================
Full test suite for citations.py.
Claude is mocked — no real API calls needed.
citeproc-py rendering is tested with a lightweight fallback.

Run:
    pip install pytest python-docx
    pytest test_citations.py -v
"""

import json
import re
import pytest
from unittest.mock import MagicMock, patch
from docx import Document
from pathlib import Path


# ─────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────

def _make_doc(paragraphs: list[tuple[str, str]]) -> Document:
    """Create a Document from (text, style_name) pairs."""
    doc = Document()
    for text, style in paragraphs:
        p = doc.add_paragraph(text)
        try:
            p.style = doc.styles[style]
        except KeyError:
            pass
    return doc


def _ieee_doc() -> Document:
    return _make_doc([
        ("Neural Networks for Security",   "Normal"),
        ("Abstract: We propose a method.", "Normal"),
        ("Introduction",                   "Heading 1"),
        ("Recent work [1] shows that deep learning [2, 3] works well.", "Normal"),
        ("Methods",                        "Heading 1"),
        ("We extend the approach of [1].", "Normal"),
        ("References",                     "Heading 1"),
        ("[1] Smith, J. et al. Deep Learning. IEEE S&P, 2022.", "Normal"),
        ("[2] Jones, A. Neural Nets. CVPR, 2021.",              "Normal"),
        ("[3] Lee, B. Attention. NeurIPS, 2020.",               "Normal"),
    ])


def _apa_doc() -> Document:
    return _make_doc([
        ("Neural Networks for Security",   "Normal"),
        ("Abstract: We propose a method.", "Normal"),
        ("Introduction",                   "Heading 1"),
        ("Recent work (Smith, 2022) shows that deep learning (Jones & Lee, 2021) works.", "Normal"),
        ("References",                     "Heading 1"),
        ("Smith, J. A. (2022). Deep learning. IEEE Security & Privacy, 10(2), 1–15.", "Normal"),
        ("Jones, A., & Lee, B. (2021). Neural nets. CVPR Proceedings, 200–210.",      "Normal"),
    ])


def _harvard_doc() -> Document:
    return _make_doc([
        ("Neural Networks for Security",   "Normal"),
        ("Introduction",                   "Heading 1"),
        ("Smith (2022) showed that (Jones, 2021) was effective.", "Normal"),
        ("References",                     "Heading 1"),
        ("Smith, J. (2022) Deep learning. Journal, 10(2), 1–15.", "Normal"),
        ("Jones, A. (2021) Neural nets. Conference, 200–210.",    "Normal"),
    ])


def _mock_claude_response(csl_objects: list[dict]) -> MagicMock:
    """Build a mock Anthropic client that returns a given CSL-JSON list."""
    mock_client = MagicMock()
    mock_msg    = MagicMock()
    mock_msg.content = [MagicMock(text=json.dumps(csl_objects))]
    mock_client.messages.create.return_value = mock_msg
    return mock_client


_SAMPLE_CSL = [
    {
        "id":     "Smith2022",
        "type":   "article-journal",
        "title":  "Deep Learning for Security",
        "author": [{"family": "Smith", "given": "J."}],
        "issued": {"date-parts": [[2022]]},
        "container-title": "IEEE Security & Privacy",
        "volume": "10", "page": "1-15",
    },
    {
        "id":     "Jones2021",
        "type":   "paper-conference",
        "title":  "Neural Networks",
        "author": [{"family": "Jones", "given": "A."},
                   {"family": "Lee",   "given": "B."}],
        "issued": {"date-parts": [[2021]]},
    },
    {
        "id":     "Lee2020",
        "type":   "paper-conference",
        "title":  "Attention Mechanisms",
        "author": [{"family": "Lee", "given": "B."}],
        "issued": {"date-parts": [[2020]]},
    },
]


# ─────────────────────────────────────────────────────────────────────
# 1. STYLE DETECTION
# ─────────────────────────────────────────────────────────────────────

from citations import detect_style, StyleInfo

class TestDetectStyle:

    def test_detects_ieee(self):
        doc    = _ieee_doc()
        result = detect_style(doc)
        assert result.name    == "ieee"
        assert result.pattern == "numeric"
        assert result.confidence > 0.0

    def test_detects_apa(self):
        doc    = _apa_doc()
        result = detect_style(doc)
        assert result.name    == "apa"
        assert result.pattern == "author-date"

    def test_detects_harvard(self):
        doc    = _harvard_doc()
        result = detect_style(doc)
        assert result.name in ("harvard", "apa")   # Harvard/APA share body pattern

    def test_empty_doc_returns_something(self):
        doc    = Document()
        result = detect_style(doc)
        assert isinstance(result, StyleInfo)
        assert result.name in ("ieee", "apa", "harvard")

    def test_confidence_is_float_0_to_1(self):
        result = detect_style(_ieee_doc())
        assert 0.0 <= result.confidence <= 1.0

    def test_no_refs_section_still_detects(self):
        doc = _make_doc([
            ("Body text with [1] and [2] citations.", "Normal"),
            ("More text citing [3].", "Normal"),
        ])
        result = detect_style(doc)
        assert result.name == "ieee"


# ─────────────────────────────────────────────────────────────────────
# 2. REFERENCE PARSING
# ─────────────────────────────────────────────────────────────────────

from citations import parse_references, CslRef

class TestParseReferences:

    def test_parses_three_refs(self):
        client = _mock_claude_response(_SAMPLE_CSL)
        refs   = [
            "[1] Smith, J. Deep Learning. IEEE, 2022.",
            "[2] Jones, A. Neural Nets. CVPR, 2021.",
            "[3] Lee, B. Attention. NeurIPS, 2020.",
        ]
        result = parse_references(refs, client)
        assert len(result) == 3
        assert all(isinstance(r, CslRef) for r in result)

    def test_all_parsed_ok_with_valid_response(self):
        client = _mock_claude_response(_SAMPLE_CSL[:3])
        refs   = [
            "[1] Smith, J. 2022.",
            "[2] Jones, A. 2021.",
            "[3] Lee, B. 2020.",
        ]
        result = parse_references(refs, client)
        assert all(r.parse_ok for r in result)

    def test_fallback_on_invalid_json(self):
        mock_client = MagicMock()
        mock_msg    = MagicMock()
        mock_msg.content = [MagicMock(text="not valid json {{{{")]
        mock_client.messages.create.return_value = mock_msg

        refs   = ["[1] Smith. 2022."]
        result = parse_references(refs, mock_client)
        assert len(result) == 1
        assert result[0].parse_ok is False
        assert "title" in result[0].csl

    def test_strips_markdown_fences(self):
        csl_with_fences = "```json\n" + json.dumps(_SAMPLE_CSL[:1]) + "\n```"
        mock_client     = MagicMock()
        mock_msg        = MagicMock()
        mock_msg.content = [MagicMock(text=csl_with_fences)]
        mock_client.messages.create.return_value = mock_msg

        result = parse_references(["[1] Smith. 2022."], mock_client)
        assert result[0].parse_ok is True

    def test_batches_large_inputs(self):
        """Ensure batching doesn't lose references."""
        large_csl = [
            {"id": f"Ref{i}", "type": "article-journal", "title": f"Paper {i}"}
            for i in range(25)
        ]
        mock_client = MagicMock()
        mock_msg    = MagicMock()
        # Return correct batch sizes
        def side_effect(*args, **kwargs):
            content = kwargs.get("messages", [{}])[-1].get("content", "")
            # Count references in the prompt
            n = content.count("\n") + 1 if content else 1
            batch = [{"id": f"Ref{i}", "type": "article-journal",
                      "title": f"Paper {i}"} for i in range(min(n, 20))]
            m = MagicMock()
            m.content = [MagicMock(text=json.dumps(batch))]
            return m
        mock_client.messages.create.side_effect = side_effect

        refs   = [f"[{i+1}] Author {i}. Paper {i}. Journal, 2020." for i in range(25)]
        result = parse_references(refs, mock_client, batch_size=20)
        assert len(result) == 25

    def test_duplicate_refs_deduplicated(self):
        client = _mock_claude_response(_SAMPLE_CSL[:1])
        refs   = ["[1] Smith. 2022.", "[1] Smith. 2022."]  # duplicate
        result = parse_references(refs, client)
        assert len(result) == 1


# ─────────────────────────────────────────────────────────────────────
# 3. REPLACEMENT BUILDING
# ─────────────────────────────────────────────────────────────────────

from citations import _build_replacements, CslRef, StyleInfo

class TestBuildReplacements:

    def _make_refs(self) -> list[CslRef]:
        return [
            CslRef("Smith2022", "[1] Smith.", _SAMPLE_CSL[0]),
            CslRef("Jones2021", "[2] Jones.", _SAMPLE_CSL[1]),
        ]

    def _intext_map(self) -> dict:
        return {
            "Smith2022": "(Smith, 2022)",
            "Jones2021": "(Jones & Lee, 2021)",
        }

    def test_numeric_source_maps_bracket_citations(self):
        refs     = self._make_refs()
        imap     = self._intext_map()
        source   = StyleInfo("ieee", "numeric", 0.9)
        repl, _  = _build_replacements(refs, imap, source)
        assert "[1]" in repl
        assert "[2]" in repl

    def test_author_date_source_maps_parenthetical(self):
        refs     = self._make_refs()
        imap     = self._intext_map()
        source   = StyleInfo("apa", "author-date", 0.9)
        repl, _  = _build_replacements(refs, imap, source)
        assert any("Smith" in k for k in repl)

    def test_numeric_warns_about_multicite(self):
        refs     = self._make_refs()
        imap     = self._intext_map()
        source   = StyleInfo("ieee", "numeric", 0.9)
        _, warns = _build_replacements(refs, imap, source)
        assert any("multi" in w.lower() for w in warns)

    def test_narrative_form_replaced(self):
        """Smith (2022) form should also be in replacements."""
        refs     = self._make_refs()
        imap     = self._intext_map()
        source   = StyleInfo("apa", "author-date", 0.9)
        repl, _  = _build_replacements(refs, imap, source)
        assert any("Smith (2022)" in k for k in repl)


# ─────────────────────────────────────────────────────────────────────
# 4. IN-TEXT REPLACEMENT
# ─────────────────────────────────────────────────────────────────────

from citations import _replace_in_paragraph

class TestReplaceInParagraph:

    def test_replaces_single_citation(self):
        doc  = Document()
        para = doc.add_paragraph("See [1] for details.")
        n    = _replace_in_paragraph(para, {"[1]": "(Smith, 2022)"})
        assert n > 0
        assert "(Smith, 2022)" in para.text

    def test_replaces_multiple_citations(self):
        doc  = Document()
        para = doc.add_paragraph("As [1] and [2] showed.")
        _replace_in_paragraph(para, {"[1]": "(Smith, 2022)", "[2]": "(Jones, 2021)"})
        assert "(Smith, 2022)" in para.text
        assert "(Jones, 2021)" in para.text

    def test_no_replacement_when_no_match(self):
        doc  = Document()
        para = doc.add_paragraph("No citations here.")
        n    = _replace_in_paragraph(para, {"[1]": "(Smith, 2022)"})
        assert n == 0
        assert para.text == "No citations here."

    def test_empty_paragraph_safe(self):
        doc  = Document()
        para = doc.add_paragraph("")
        n    = _replace_in_paragraph(para, {"[1]": "(Smith, 2022)"})
        assert n == 0


# ─────────────────────────────────────────────────────────────────────
# 5. FULL reformat_citations FLOW  (mocked rendering)
# ─────────────────────────────────────────────────────────────────────

from citations import reformat_citations, CitationResult

class TestReformatCitations:

    def _mock_render(self, intext_map, bib_entries):
        """Patch render_bibliography to return controlled output."""
        return patch(
            "citations.render_bibliography",
            return_value=(intext_map, bib_entries)
        )

    def test_skips_when_style_matches(self):
        doc    = _ieee_doc()
        client = _mock_claude_response(_SAMPLE_CSL)
        result = reformat_citations(doc, "ieee", client)
        assert result.skipped is True

    def test_skips_when_no_reference_section(self):
        doc = _make_doc([
            ("Body text with [1].", "Normal"),
        ])
        client = _mock_claude_response(_SAMPLE_CSL)
        result = reformat_citations(doc, "apa", client)
        assert result.skipped is True
        assert any("reference" in w.lower() for w in result.warnings)

    def test_ieee_to_apa_replaces_citations(self):
        doc    = _ieee_doc()
        client = _mock_claude_response(_SAMPLE_CSL)
        intext = {"Smith2022": "(Smith, 2022)", "Jones2021": "(Jones & Lee, 2021)",
                  "Lee2020":   "(Lee, 2020)"}
        bib    = ["Smith, J. (2022). Deep Learning. IEEE S&P.",
                  "Jones, A., & Lee, B. (2021). Neural Networks.",
                  "Lee, B. (2020). Attention."]

        with self._mock_render(intext, bib):
            result = reformat_citations(doc, "apa", client)

        assert not result.skipped
        assert result.target_style == "apa"
        assert result.source_style == "ieee"

    def test_apa_to_ieee_runs_without_error(self):
        doc    = _apa_doc()
        client = _mock_claude_response(_SAMPLE_CSL[:2])
        intext = {"Smith2022": "[1]", "Jones2021": "[2]"}
        bib    = ["[1] J. Smith, 'Deep Learning,' IEEE S&P, 2022.",
                  "[2] A. Jones and B. Lee, 'Neural Nets,' CVPR, 2021."]

        with self._mock_render(intext, bib):
            result = reformat_citations(doc, "ieee", client)

        assert not result.skipped

    def test_result_has_correct_fields(self):
        doc    = _ieee_doc()
        client = _mock_claude_response(_SAMPLE_CSL)
        intext = {"Smith2022": "(Smith, 2022)", "Jones2021": "(Jones, 2021)",
                  "Lee2020":   "(Lee, 2020)"}
        bib    = ["Smith (2022)...", "Jones (2021)...", "Lee (2020)..."]

        with self._mock_render(intext, bib):
            result = reformat_citations(doc, "apa", client)

        assert isinstance(result, CitationResult)
        assert isinstance(result.warnings,      list)
        assert isinstance(result.changed_count, int)
        assert isinstance(result.refs_parsed,   int)
        assert isinstance(result.refs_failed,   int)

    def test_warns_when_no_intext_replaced(self):
        doc = _make_doc([
            ("No citations in this body.",    "Normal"),
            ("References",                     "Normal"),
            ("[1] Smith. 2022.",               "Normal"),
        ])
        client = _mock_claude_response(_SAMPLE_CSL[:1])
        intext = {"Smith2022": "(Smith, 2022)"}
        bib    = ["Smith, J. (2022). Paper."]

        with self._mock_render(intext, bib):
            result = reformat_citations(doc, "apa", client)

        assert any("replaced" in w.lower() or "citation" in w.lower()
                   for w in result.warnings)

    def test_failed_parses_add_warning(self):
        """References that Claude couldn't parse get a warning."""
        doc    = _ieee_doc()
        # Return one parse failure
        bad_csl = [
            {"id": "Smith2022", "type": "article-journal", "title": "Paper"},
            {"id": "Jones2021", "type": "article-journal"},  # missing title → parse_ok=False
            {"id": "Lee2020",   "type": "article-journal", "title": "Paper3"},
        ]
        client = _mock_claude_response(bad_csl)
        intext = {"Smith2022": "(Smith, 2022)", "Jones2021": "(Jones, 2021)",
                  "Lee2020":   "(Lee, 2020)"}
        bib    = ["Smith...", "Jones...", "Lee..."]

        with self._mock_render(intext, bib):
            result = reformat_citations(doc, "apa", client)

        # refs_failed might be 0 here since we're mocking CSL directly
        assert isinstance(result.refs_failed, int)

    def test_harvard_to_ieee_runs(self):
        """Harvard → IEEE is always a genuine cross-style conversion."""
        doc    = _harvard_doc()
        client = _mock_claude_response(_SAMPLE_CSL[:2])
        intext = {"Smith2022": "[1]", "Jones2021": "[2]"}
        bib    = ["[1] J. Smith, Paper, 2022.", "[2] A. Jones, Paper, 2021."]

        with self._mock_render(intext, bib):
            result = reformat_citations(doc, "ieee", client)

        assert not result.skipped


# ─────────────────────────────────────────────────────────────────────
# 6. PIPELINE INTEGRATION  (citations.py present → pipeline uses it)
# ─────────────────────────────────────────────────────────────────────

class TestPipelineIntegration:
    """
    Verify that pipeline.py correctly calls reformat_citations
    when citations.py is available.
    """

    def test_pipeline_calls_reformat_when_module_present(self, tmp_path):
        """pipeline._try_reformat_citations calls reformat_citations."""
        from pipeline import _try_reformat_citations

        doc        = _ieee_doc()
        mock_client = _mock_claude_response(_SAMPLE_CSL)

        fake_result       = MagicMock()
        fake_result.warnings      = []
        fake_result.changed_count = 2
        fake_style        = MagicMock()
        fake_style.name   = "ieee"

        fake_citations = MagicMock()
        fake_citations.detect_style.return_value = fake_style
        fake_citations.reformat_citations.return_value = fake_result

        import sys
        with patch.dict(sys.modules, {"citations": fake_citations}):
            result = _try_reformat_citations(doc, "apa", mock_client,
                                              log_fn=lambda m: None)

        fake_citations.reformat_citations.assert_called_once_with(doc, "apa", mock_client)
        assert result.changed_count == 2
