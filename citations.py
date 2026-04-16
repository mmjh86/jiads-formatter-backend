"""
citations.py
============
Citation style detection, parsing, rendering, and replacement.

Supports:
  · Detection:  IEEE numeric, APA author-date, Harvard author-date
  · Conversion: all directions (IEEE ↔ APA ↔ Harvard)
  · Parsing:    Claude parses every reference into CSL-JSON
  · Rendering:  citeproc-py with correct single-registration pattern

Install:
    pip install citeproc-py anthropic python-docx

CSL files (download once):
    mkdir -p csl_styles
    curl -L -o csl_styles/ieee.csl    https://www.zotero.org/styles/ieee
    curl -L -o csl_styles/apa-7.csl   https://www.zotero.org/styles/apa-7th-edition
    curl -L -o csl_styles/harvard.csl https://www.zotero.org/styles/harvard-university-of-warwick

Usage:
    from citations import detect_style, reformat_citations
    client  = anthropic.Anthropic()
    style   = detect_style(doc)
    result  = reformat_citations(doc, "apa", client)
"""

import re
import json
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from docx import Document
from docx.oxml.ns import qn

logger = logging.getLogger(__name__)

CSL_DIR = Path(__file__).parent / "csl_styles"

# ─────────────────────────────────────────────────────────────────────
# DATA TYPES
# ─────────────────────────────────────────────────────────────────────

@dataclass
class StyleInfo:
    name:       str          # "ieee" | "apa" | "harvard"
    pattern:    str          # "numeric" | "author-date"
    confidence: float        # 0.0–1.0


@dataclass
class CslRef:
    """A single reference in CSL-JSON form."""
    id:        str
    raw:       str           # original paragraph text
    csl:       dict          # CSL-JSON fields
    parse_ok:  bool = True   # False if Claude could not parse


@dataclass
class CitationResult:
    warnings:      list[str] = field(default_factory=list)
    changed_count: int = 0
    source_style:  Optional[str] = None
    target_style:  Optional[str] = None
    skipped:       bool = False
    refs_parsed:   int = 0
    refs_failed:   int = 0


# ─────────────────────────────────────────────────────────────────────
# 1. STYLE DETECTION
# ─────────────────────────────────────────────────────────────────────

# IEEE: [1], [2, 3], [1]–[3] in body text
_IEEE_INTEXT   = re.compile(r"\[(\d+(?:[,\s]*\d+)*)\]")
# APA/Harvard: (Smith, 2020) or (Smith & Jones, 2020) or Smith (2020)
_APA_INTEXT    = re.compile(
    r"\(([A-Z][a-zA-Zé\-\s]+(?:\s*[&,]\s*[A-Z][a-zA-Zé\-\s]+)*"
    r"(?:\s*et\s+al\.?)?,\s*\d{4}[a-z]?)\)"
)
# IEEE reference entry: [1] Author...
_IEEE_REF      = re.compile(r"^\s*\[\d+\]")
# APA reference entry: Author, A. A. (Year).
_APA_REF       = re.compile(
    r"^\s*[A-Z][a-zA-Zé\-]+,\s+[A-Z][\w.]*\s*[\(&]?\s*\d{4}"
)
# Harvard reference entry: Author (Year)  — similar to APA
_HARVARD_REF   = re.compile(
    r"^\s*[A-Z][a-zA-Zé\-]+,\s+[A-Z][\w.]*\s*\(\d{4}\)"
)

_REF_SECTION   = re.compile(
    r"^\s*(references?|bibliography|works?\s+cited)\s*$", re.IGNORECASE
)


def _reference_paragraphs(doc: Document) -> list[str]:
    in_refs = False
    refs    = []
    for para in doc.paragraphs:
        t = para.text.strip()
        if not in_refs and _REF_SECTION.match(t):
            in_refs = True
            continue
        if in_refs and t:
            refs.append(t)
    return refs


def detect_style(doc: Document) -> StyleInfo:
    """
    Analyses body text and reference list to identify citation style.
    Returns StyleInfo with name, pattern, and confidence score.
    """
    body_text = " ".join(
        p.text for p in doc.paragraphs
        if p.text.strip() and not _REF_SECTION.match(p.text.strip())
    )
    ref_paras = _reference_paragraphs(doc)

    ieee_body  = len(_IEEE_INTEXT.findall(body_text))
    apa_body   = len(_APA_INTEXT.findall(body_text))
    ieee_refs  = sum(1 for r in ref_paras if _IEEE_REF.match(r))
    apa_refs   = sum(1 for r in ref_paras if _APA_REF.match(r))
    harv_refs  = sum(1 for r in ref_paras if _HARVARD_REF.match(r))

    ieee_score = ieee_body * 2 + ieee_refs * 3
    apa_score  = apa_body  * 2 + apa_refs  * 3
    harv_score = apa_body  * 1 + harv_refs * 3   # Harvard and APA share body pattern

    total = max(ieee_score + apa_score + harv_score, 1)

    if ieee_score >= apa_score and ieee_score >= harv_score:
        return StyleInfo("ieee", "numeric",      min(ieee_score / total, 1.0))
    if harv_score >= apa_score:
        return StyleInfo("harvard", "author-date", min(harv_score / total, 1.0))
    return StyleInfo("apa", "author-date",       min(apa_score  / total, 1.0))


# ─────────────────────────────────────────────────────────────────────
# 2. REFERENCE PARSING  (Claude for every reference)
# ─────────────────────────────────────────────────────────────────────

_PARSE_SYSTEM = """\
You are a bibliographic data extraction engine.
Given raw reference strings, return a JSON array of CSL-JSON objects.

Rules:
- Output ONLY a valid JSON array — no preamble, no markdown fences.
- Each object must include at minimum: "id", "type", "title".
- Infer "type" from context: "article-journal", "paper-conference",
  "book", "chapter", "thesis", "webpage", "report".
- For "author" use array of {"family": ..., "given": ...}.
- For "issued" use {"date-parts": [[year]]}.
- If a field cannot be determined, omit it — do not guess.
- "id" must be unique: use AuthorYear format e.g. "Smith2020".
"""

_PARSE_USER = """\
Parse these references into CSL-JSON. Return a JSON array, one object per reference.

References:
{refs}
"""

_SCHEMA_REQUIRED = {"id", "type", "title"}


def _call_claude(client, prompt: str) -> str:
    msg = client.messages.create(
        model      = "claude-opus-4-5",
        max_tokens = 4096,
        system     = _PARSE_SYSTEM,
        messages   = [{"role": "user", "content": prompt}],
    )
    return msg.content[0].text.strip()


def _clean_json(text: str) -> str:
    """Strip markdown fences if Claude wrapped output despite instructions."""
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"\s*```$",          "", text, flags=re.MULTILINE)
    return text.strip()


def _validate_csl(obj: dict) -> bool:
    return _SCHEMA_REQUIRED.issubset(obj.keys())


def parse_references(
    ref_paragraphs: list[str],
    client,
    batch_size: int = 20,
) -> list[CslRef]:
    """
    Sends all references to Claude in batches.
    Returns a list of CslRef objects (one per reference paragraph).
    """
    results: list[CslRef] = []
    # Deduplicate while preserving order
    seen, unique = set(), []
    for r in ref_paragraphs:
        if r not in seen:
            seen.add(r); unique.append(r)

    for i in range(0, len(unique), batch_size):
        batch = unique[i : i + batch_size]
        numbered = "\n".join(f"{j+1}. {r}" for j, r in enumerate(batch))
        prompt   = _PARSE_USER.format(refs=numbered)

        try:
            raw     = _call_claude(client, prompt)
            cleaned = _clean_json(raw)
            parsed  = json.loads(cleaned)

            # Match parsed objects back to batch entries by position
            for j, ref_text in enumerate(batch):
                if j < len(parsed) and isinstance(parsed[j], dict):
                    csl = parsed[j]
                    ok  = _validate_csl(csl)
                    if ok and not csl.get("id"):
                        csl["id"] = f"ref{i + j}"
                    results.append(CslRef(
                        id       = csl.get("id", f"ref{i+j}"),
                        raw      = ref_text,
                        csl      = csl,
                        parse_ok = ok,
                    ))
                else:
                    logger.warning("No parsed object for reference %d: %s", i+j, ref_text[:60])
                    results.append(CslRef(
                        id       = f"ref{i+j}",
                        raw      = ref_text,
                        csl      = {"id": f"ref{i+j}", "type": "article-journal",
                                    "title": ref_text[:120]},
                        parse_ok = False,
                    ))
        except (json.JSONDecodeError, Exception) as e:
            logger.error("Claude parse failed for batch %d: %s", i, e)
            for j, ref_text in enumerate(batch):
                results.append(CslRef(
                    id       = f"ref{i+j}",
                    raw      = ref_text,
                    csl      = {"id": f"ref{i+j}", "type": "article-journal",
                                "title": ref_text[:120]},
                    parse_ok = False,
                ))
    return results


# ─────────────────────────────────────────────────────────────────────
# 3. CITATION RENDERING  (citeproc-py, single-registration pattern)
# ─────────────────────────────────────────────────────────────────────

def _csl_path(style_name: str) -> Path:
    mapping = {
        "ieee":    "ieee.csl",
        "apa":     "apa-7.csl",
        "apa7":    "apa-7.csl",
        "harvard": "harvard.csl",
    }
    fname = mapping.get(style_name.lower())
    if not fname:
        raise ValueError(f"No CSL file for style '{style_name}'. "
                         f"Available: {list(mapping)}")
    path = CSL_DIR / fname
    if not path.exists():
        raise FileNotFoundError(
            f"CSL file not found: {path}\n"
            f"Run: curl -L -o {path} https://www.zotero.org/styles/{fname.replace('.csl','')}"
        )
    return path


def render_bibliography(
    csl_refs:    list[CslRef],
    target_style: str,
) -> tuple[dict[str, str], list[str]]:
    """
    Renders all references in target_style.

    CRITICAL: registers ALL citations in a single pass before rendering.
    Rendering one at a time (the naive approach) produces incorrect output
    because citeproc-py uses global state for disambiguation.

    Returns:
        intext_map:  {ref_id → in-text citation string}
        bib_entries: [formatted bibliography string per ref, in order]
    """
    from citeproc import CitationStylesStyle, CitationStylesBibliography
    from citeproc import Citation, CitationItem
    from citeproc.source.json import CiteProcJSON
    from citeproc import formatter

    # Build source
    csl_data   = [ref.csl for ref in csl_refs]
    bib_source = CiteProcJSON(csl_data)
    style      = CitationStylesStyle(str(_csl_path(target_style)), validate=False)
    bibliography = CitationStylesBibliography(style, bib_source, formatter.plain)

    # ── Register ALL citations first ──────────────────────────────────
    citation_objects = []
    for ref in csl_refs:
        cit = Citation([CitationItem(ref.id)])
        bibliography.register(cit)
        citation_objects.append((ref.id, cit))

    # ── Render in-text strings ────────────────────────────────────────
    intext_map: dict[str, str] = {}
    for ref_id, cit in citation_objects:
        try:
            intext_map[ref_id] = str(bibliography.cite(cit, lambda x: None))
        except Exception as e:
            logger.warning("In-text render failed for %s: %s", ref_id, e)
            intext_map[ref_id] = f"[{ref_id}]"

    # ── Render bibliography entries ───────────────────────────────────
    bib_entries: list[str] = []
    try:
        for entry in bibliography.bibliography():
            bib_entries.append(str(entry))
    except Exception as e:
        logger.error("Bibliography render failed: %s", e)
        bib_entries = [ref.raw for ref in csl_refs]

    return intext_map, bib_entries


# ─────────────────────────────────────────────────────────────────────
# 4. IN-TEXT REPLACEMENT
# ─────────────────────────────────────────────────────────────────────

def _build_replacements(
    csl_refs:    list[CslRef],
    intext_map:  dict[str, str],
    source_style: StyleInfo,
) -> dict[str, str]:
    """
    Builds a mapping of old in-text citation text → new rendered string.
    Handles single citations and warns on multi-cite groups.
    """
    replacements: dict[str, str] = {}
    warnings:     list[str]      = []

    if source_style.pattern == "numeric":
        # Map [n] → intext_map[ref_id] using position
        for i, ref in enumerate(csl_refs, 1):
            old = f"[{i}]"
            new = intext_map.get(ref.id, old)
            replacements[old] = new
        # Warn about multi-cites: [1,2] or [1, 2]
        warnings.append(
            "Multi-citations (e.g. [1,2,3]) were not automatically combined. "
            "Please verify grouped citations manually."
        )
    else:
        # APA/Harvard: (Author, Year) — already rendered by citeproc
        # Map raw author-year pattern → rendered string
        for ref in csl_refs:
            authors = ref.csl.get("author", [])
            year    = (ref.csl.get("issued", {})
                           .get("date-parts", [[""]])[0][0])
            if authors:
                family = authors[0].get("family", "")
                if len(authors) == 1:
                    old_key = f"({family}, {year})"
                elif len(authors) == 2:
                    fam2    = authors[1].get("family", "")
                    old_key = f"({family} & {fam2}, {year})"
                else:
                    old_key = f"({family} et al., {year})"
                new = intext_map.get(ref.id, old_key)
                replacements[old_key] = new
                # Also handle narrative form: Smith (2020)
                narrative     = f"{family} ({year})"
                replacements[narrative] = new.strip("()")

    return replacements, warnings


def _replace_in_paragraph(para, replacements: dict[str, str]) -> int:
    """
    Replaces citation strings in a paragraph's runs.
    Handles citations that may span multiple runs by working on
    the full paragraph text and redistributing to runs.
    Returns number of replacements made.
    """
    full_text = para.text
    replaced  = 0

    for old, new in replacements.items():
        if old in full_text:
            full_text = full_text.replace(old, new)
            replaced += full_text.count(new)

    if replaced == 0:
        return 0

    # Redistribute text across runs (preserving run count and formatting)
    if not para.runs:
        return 0

    # Simplest safe approach: put all text in first run, clear the rest
    # This preserves the first run's formatting for the whole paragraph.
    # A more sophisticated approach would redistribute proportionally.
    para.runs[0].text = full_text
    for run in para.runs[1:]:
        run.text = ""

    return replaced


# ─────────────────────────────────────────────────────────────────────
# 5. REFERENCE LIST REPLACEMENT
# ─────────────────────────────────────────────────────────────────────

def _replace_reference_list(doc: Document, bib_entries: list[str]) -> int:
    """
    Finds the reference section and replaces its paragraphs
    with the newly rendered bibliography entries.
    Returns number of paragraphs replaced.
    """
    # Find reference section start index
    ref_start = None
    for i, para in enumerate(doc.paragraphs):
        if _REF_SECTION.match(para.text.strip()):
            ref_start = i
            break

    if ref_start is None:
        logger.warning("No reference section found — bibliography not replaced.")
        return 0

    # Collect existing reference paragraphs
    existing_ref_paras = []
    for para in doc.paragraphs[ref_start + 1:]:
        if para.text.strip():
            existing_ref_paras.append(para)

    # Replace text in existing paragraphs (reuse formatting)
    replaced = 0
    for i, entry in enumerate(bib_entries):
        if i < len(existing_ref_paras):
            existing_ref_paras[i].runs[0].text = entry if existing_ref_paras[i].runs else entry
            for run in (existing_ref_paras[i].runs[1:] if existing_ref_paras[i].runs else []):
                run.text = ""
            replaced += 1
        else:
            # Add new paragraph after the last reference paragraph
            last = existing_ref_paras[-1] if existing_ref_paras else doc.paragraphs[ref_start]
            new_para = last._element.addnext(last._element.__class__())
            # Simpler: use Document.add_paragraph — but we need insertion order
            # So we copy the last ref para's XML and set text
            import copy
            new_p = copy.deepcopy(last._p)
            last._p.addnext(new_p)
            from docx.text.paragraph import Paragraph
            Paragraph(new_p, doc).runs[0].text = entry if Paragraph(new_p, doc).runs else entry
            replaced += 1

    # Clear any leftover old reference paragraphs
    for i in range(len(bib_entries), len(existing_ref_paras)):
        existing_ref_paras[i].clear()

    return replaced


# ─────────────────────────────────────────────────────────────────────
# 6. MAIN ENTRY POINT
# ─────────────────────────────────────────────────────────────────────

def reformat_citations(
    doc:          Document,
    target_style: str,
    client,
) -> CitationResult:
    """
    Detects the source citation style, parses all references via Claude,
    renders them in target_style, and replaces in-text citations +
    reference list in-place in the Document.

    Args:
        doc:          python-docx Document (modified in-place)
        target_style: "ieee" | "apa" | "apa7" | "harvard"
        client:       anthropic.Anthropic() instance

    Returns:
        CitationResult with warnings, counts, and success flag
    """
    result = CitationResult(target_style=target_style)

    # ── Detect source style ───────────────────────────────────────────
    source = detect_style(doc)
    result.source_style = source.name
    logger.info("Detected source style: %s (confidence %.2f)", source.name, source.confidence)

    if source.name == target_style.lower().replace("apa7", "apa"):
        logger.info("Source and target styles match — skipping.")
        result.skipped = True
        return result

    # ── Extract reference paragraphs ──────────────────────────────────
    ref_paras = _reference_paragraphs(doc)
    if not ref_paras:
        result.warnings.append(
            "No reference section found. Add a 'References' heading before your reference list."
        )
        result.skipped = True
        return result

    logger.info("Parsing %d references via Claude…", len(ref_paras))

    # ── Parse to CSL-JSON ─────────────────────────────────────────────
    csl_refs = parse_references(ref_paras, client)
    result.refs_parsed = sum(1 for r in csl_refs if r.parse_ok)
    result.refs_failed = sum(1 for r in csl_refs if not r.parse_ok)

    if result.refs_failed:
        result.warnings.append(
            f"{result.refs_failed} reference(s) could not be fully parsed. "
            "They have been included with minimal metadata — please verify."
        )

    # ── Render in target style ────────────────────────────────────────
    logger.info("Rendering %d references in %s style…", len(csl_refs), target_style)
    try:
        intext_map, bib_entries = render_bibliography(csl_refs, target_style)
    except FileNotFoundError as e:
        result.warnings.append(str(e))
        result.skipped = True
        return result

    # ── Build replacement map ─────────────────────────────────────────
    replacements, repl_warnings = _build_replacements(csl_refs, intext_map, source)
    result.warnings.extend(repl_warnings)

    # ── Replace in-text citations ─────────────────────────────────────
    total_replaced = 0
    for para in doc.paragraphs:
        if _REF_SECTION.match(para.text.strip()):
            break
        total_replaced += _replace_in_paragraph(para, replacements)

    result.changed_count = total_replaced
    logger.info("Replaced %d in-text citations.", total_replaced)

    # ── Replace reference list ────────────────────────────────────────
    n_bib = _replace_reference_list(doc, bib_entries)
    logger.info("Replaced %d bibliography entries.", n_bib)

    if total_replaced == 0:
        result.warnings.append(
            "No in-text citations were replaced. "
            "Check that your citation format matches the detected source style "
            f"({source.name})."
        )

    return result
