"""
style_profiles.py
=================
Journal style definitions for IEEE, APA 7th, and Harvard.
Extensible: add a new style by defining a dict and registering it.

All measurements in:
  - Cm  → centimetres  (for margins)
  - Pt  → points       (for font sizes, spacing)
  - DXA → 1/20th pt   (internal python-docx unit, converted automatically)

Usage:
    from style_profiles import get_profile, register_profile, list_profiles
    profile = get_profile("ieee")
"""

from dataclasses import dataclass, field
from typing import Optional
from docx.shared import Cm, Pt


# ─────────────────────────────────────────────────────────────────────
# DATA STRUCTURES
# ─────────────────────────────────────────────────────────────────────

@dataclass
class MarginSpec:
    top:    float   # cm
    bottom: float   # cm
    left:   float   # cm
    right:  float   # cm
    gutter: float = 0.0   # cm


@dataclass
class FontSpec:
    name:  str
    size:  float     # pt
    bold:  bool = False
    italic: bool = False


@dataclass
class SpacingSpec:
    before:     float = 0.0    # pt — space before paragraph
    after:      float = 0.0    # pt — space after paragraph
    line:       float = 1.0    # multiplier (1.0 = single, 2.0 = double)
    line_rule:  str   = "auto" # "auto" | "exact" | "atLeast"


@dataclass
class ParagraphSpec:
    font:           FontSpec
    spacing:        SpacingSpec
    alignment:      str = "left"   # "left"|"center"|"right"|"justify"
    first_line_indent: float = 0.0  # cm


@dataclass
class ColumnSpec:
    count:  int   = 1
    space:  float = 1.25  # cm gap between columns


@dataclass
class StyleProfile:
    name:           str
    display_name:   str
    citation_style: str            # "ieee" | "apa" | "harvard"
    page_size:      str            # "A4" | "letter"
    margins:        MarginSpec
    columns:        ColumnSpec
    body:           ParagraphSpec
    abstract:       ParagraphSpec
    h1:             ParagraphSpec  # section heading
    h2:             ParagraphSpec  # subsection heading
    h3:             ParagraphSpec  # subsubsection heading
    caption:        ParagraphSpec
    references:     ParagraphSpec
    line_numbers:   bool = False
    abstract_word_limit: Optional[int] = None
    notes:          str = ""       # human-readable notes for users


# ─────────────────────────────────────────────────────────────────────
# BUILT-IN PROFILES
# ─────────────────────────────────────────────────────────────────────

_IEEE = StyleProfile(
    name          = "ieee",
    display_name  = "IEEE Transactions",
    citation_style= "ieee",
    page_size     = "letter",
    margins       = MarginSpec(top=1.9, bottom=2.54, left=1.9, right=1.9),
    columns       = ColumnSpec(count=2, space=0.51),

    body = ParagraphSpec(
        font     = FontSpec("Times New Roman", 10),
        spacing  = SpacingSpec(before=0, after=0, line=1.0),
        alignment= "justify",
        first_line_indent=0.0,
    ),
    abstract = ParagraphSpec(
        font     = FontSpec("Times New Roman", 9, italic=True),
        spacing  = SpacingSpec(before=0, after=6, line=1.0),
        alignment= "justify",
    ),
    h1 = ParagraphSpec(
        font     = FontSpec("Times New Roman", 10, bold=False),
        spacing  = SpacingSpec(before=12, after=3, line=1.0),
        alignment= "center",
    ),
    h2 = ParagraphSpec(
        font     = FontSpec("Times New Roman", 10, italic=True),
        spacing  = SpacingSpec(before=6, after=3, line=1.0),
        alignment= "left",
    ),
    h3 = ParagraphSpec(
        font     = FontSpec("Times New Roman", 10, italic=True),
        spacing  = SpacingSpec(before=6, after=0, line=1.0),
        alignment= "left",
    ),
    caption = ParagraphSpec(
        font     = FontSpec("Times New Roman", 8),
        spacing  = SpacingSpec(before=3, after=6, line=1.0),
        alignment= "center",
    ),
    references = ParagraphSpec(
        font     = FontSpec("Times New Roman", 8),
        spacing  = SpacingSpec(before=0, after=3, line=1.0),
        alignment= "left",
        first_line_indent=0.0,
    ),
    abstract_word_limit=200,
    notes="Two-column layout. Section headings centred in small caps. "
          "References as numbered list [1], [2], …",
)

_APA7 = StyleProfile(
    name          = "apa7",
    display_name  = "APA 7th Edition",
    citation_style= "apa",
    page_size     = "letter",
    margins       = MarginSpec(top=2.54, bottom=2.54, left=2.54, right=2.54),
    columns       = ColumnSpec(count=1),

    body = ParagraphSpec(
        font     = FontSpec("Times New Roman", 12),
        spacing  = SpacingSpec(before=0, after=0, line=2.0),
        alignment= "left",
        first_line_indent=1.27,
    ),
    abstract = ParagraphSpec(
        font     = FontSpec("Times New Roman", 12),
        spacing  = SpacingSpec(before=0, after=0, line=2.0),
        alignment= "left",
        first_line_indent=0.0,
    ),
    h1 = ParagraphSpec(
        font     = FontSpec("Times New Roman", 12, bold=True),
        spacing  = SpacingSpec(before=0, after=0, line=2.0),
        alignment= "center",
    ),
    h2 = ParagraphSpec(
        font     = FontSpec("Times New Roman", 12, bold=True),
        spacing  = SpacingSpec(before=0, after=0, line=2.0),
        alignment= "left",
    ),
    h3 = ParagraphSpec(
        font     = FontSpec("Times New Roman", 12, bold=True, italic=True),
        spacing  = SpacingSpec(before=0, after=0, line=2.0),
        alignment= "left",
        first_line_indent=1.27,
    ),
    caption = ParagraphSpec(
        font     = FontSpec("Times New Roman", 12),
        spacing  = SpacingSpec(before=0, after=0, line=2.0),
        alignment= "left",
    ),
    references = ParagraphSpec(
        font     = FontSpec("Times New Roman", 12),
        spacing  = SpacingSpec(before=0, after=0, line=2.0),
        alignment= "left",
        first_line_indent=-1.27,    # hanging indent
    ),
    abstract_word_limit=250,
    notes="Double-spaced throughout. Hanging indent on references. "
          "Running head on every page.",
)

_HARVARD = StyleProfile(
    name          = "harvard",
    display_name  = "Harvard (Warwick)",
    citation_style= "harvard",
    page_size     = "A4",
    margins       = MarginSpec(top=2.54, bottom=2.54, left=3.0, right=2.54),
    columns       = ColumnSpec(count=1),

    body = ParagraphSpec(
        font     = FontSpec("Arial", 12),
        spacing  = SpacingSpec(before=0, after=12, line=1.5),
        alignment= "left",
        first_line_indent=0.0,
    ),
    abstract = ParagraphSpec(
        font     = FontSpec("Arial", 11),
        spacing  = SpacingSpec(before=0, after=12, line=1.5),
        alignment= "justify",
    ),
    h1 = ParagraphSpec(
        font     = FontSpec("Arial", 14, bold=True),
        spacing  = SpacingSpec(before=18, after=6, line=1.0),
        alignment= "left",
    ),
    h2 = ParagraphSpec(
        font     = FontSpec("Arial", 13, bold=True),
        spacing  = SpacingSpec(before=12, after=6, line=1.0),
        alignment= "left",
    ),
    h3 = ParagraphSpec(
        font     = FontSpec("Arial", 12, bold=True, italic=True),
        spacing  = SpacingSpec(before=12, after=6, line=1.0),
        alignment= "left",
    ),
    caption = ParagraphSpec(
        font     = FontSpec("Arial", 10, italic=True),
        spacing  = SpacingSpec(before=3, after=12, line=1.0),
        alignment= "left",
    ),
    references = ParagraphSpec(
        font     = FontSpec("Arial", 12),
        spacing  = SpacingSpec(before=0, after=6, line=1.5),
        alignment= "left",
        first_line_indent=-1.27,    # hanging indent
    ),
    abstract_word_limit=300,
    notes="A4 page, left margin wider for binding. "
          "Author-date citations (Surname, Year). "
          "References section hanging indent.",
)


# ─────────────────────────────────────────────────────────────────────
# REGISTRY
# ─────────────────────────────────────────────────────────────────────

_REGISTRY: dict[str, StyleProfile] = {
    "ieee":    _IEEE,
    "apa7":    _APA7,
    "apa":     _APA7,     # alias
    "harvard": _HARVARD,
}


def get_profile(name: str) -> StyleProfile:
    """
    Returns a StyleProfile by name (case-insensitive).
    Raises KeyError with helpful message if not found.
    """
    key = name.lower().strip()
    if key not in _REGISTRY:
        available = list_profiles()
        raise KeyError(
            f"Style '{name}' not found. Available: {available}\n"
            f"To add a custom style: register_profile(your_profile)"
        )
    return _REGISTRY[key]


def register_profile(profile: StyleProfile):
    """
    Register a custom StyleProfile.
    Call this before get_profile() if you need custom journal styles.

    Example:
        from style_profiles import register_profile, StyleProfile, ...
        my_profile = StyleProfile(name="nature", ...)
        register_profile(my_profile)
    """
    _REGISTRY[profile.name.lower()] = profile
    # Also register common aliases
    if profile.name.lower() not in _REGISTRY:
        _REGISTRY[profile.citation_style.lower()] = profile


def list_profiles() -> list[str]:
    """Returns unique profile names (no aliases)."""
    seen, names = set(), []
    for k, v in _REGISTRY.items():
        if v.name not in seen:
            seen.add(v.name)
            names.append(v.name)
    return names
