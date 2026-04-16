"""
pipeline.py
===========
Integration glue: citations → layout → validation in one call.

Connects:
  · citations.py      → detect + reformat citation style
  · layout_formatter  → apply margins/fonts/spacing
  · (optional) validation report

Single entry point:
    result = run_pipeline("manuscript.docx", target_style="ieee",
                          anthropic_client=client)

Returns PipelineResult with output path, change log, warnings, and cost.

Install:
    pip install python-docx anthropic

.env:
    ANTHROPIC_API_KEY=sk-ant-...
"""

import os
import time
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from docx import Document

from layout_formatter import (
    load_document, format_document, FormattedResult, Change,
)
from style_profiles import get_profile, StyleProfile


# ─────────────────────────────────────────────────────────────────────
# CITATION STUB  (replace with real citations.py when available)
# ─────────────────────────────────────────────────────────────────────

@dataclass
class CitationResult:
    """Mirrors the return shape citations.py should produce."""
    warnings:      list[str] = field(default_factory=list)
    changed_count: int = 0
    source_style:  Optional[str] = None
    target_style:  Optional[str] = None
    skipped:       bool = False    # True if styles already match


def _try_reformat_citations(
    doc:             Document,
    target_style:    str,
    anthropic_client,
    log_fn = print,
) -> CitationResult:
    """
    Attempts to import and call citations.reformat_citations().
    Falls back gracefully if citations.py is not yet available.
    """
    try:
        from citations import reformat_citations, detect_style  # type: ignore

        source = detect_style(doc)
        if source and source.name.lower() == target_style.lower():
            log_fn(f"   Citations already in {target_style} style — skipping reformat.")
            return CitationResult(
                source_style=source.name,
                target_style=target_style,
                skipped=True,
            )

        log_fn(f"   Reformatting citations: {source.name if source else 'unknown'} → {target_style}")
        result = reformat_citations(doc, target_style, anthropic_client)
        return CitationResult(
            warnings      = result.warnings,
            changed_count = result.changed_count,
            source_style  = source.name if source else None,
            target_style  = target_style,
        )

    except ImportError:
        log_fn("   ⚠️  citations.py not found — skipping citation reformat.")
        log_fn("       Add citations.py alongside pipeline.py to enable this step.")
        return CitationResult(
            warnings=["citations.py not installed — citation style not changed."],
            skipped=True,
        )
    except Exception as e:
        log_fn(f"   ⚠️  Citation reformat failed: {e}")
        return CitationResult(
            warnings=[f"Citation reformat error: {e}"],
            skipped=True,
        )


# ─────────────────────────────────────────────────────────────────────
# PIPELINE RESULT
# ─────────────────────────────────────────────────────────────────────

@dataclass
class PipelineResult:
    # Outputs
    output_path:    str
    profile:        StyleProfile

    # Step results
    citation_result: CitationResult
    layout_result:   FormattedResult

    # Aggregated
    warnings:       list[str]     = field(default_factory=list)
    errors:         list[str]     = field(default_factory=list)
    elapsed_sec:    float         = 0.0
    success:        bool          = True

    @property
    def all_changes(self) -> list[Change]:
        return self.layout_result.changes

    @property
    def change_summary(self) -> str:
        n = len(self.all_changes)
        cit = self.citation_result
        parts = [f"{n} layout changes"]
        if not cit.skipped:
            parts.append(f"{cit.changed_count} citation(s) reformatted")
        if self.warnings:
            parts.append(f"{len(self.warnings)} warning(s)")
        return " · ".join(parts)

    def report(self) -> str:
        lines = [
            f"{'='*56}",
            f"  Pipeline Result — {self.profile.display_name}",
            f"{'='*56}",
            f"  Output:   {self.output_path}",
            f"  Status:   {'✅ success' if self.success else '❌ failed'}",
            f"  Time:     {self.elapsed_sec:.1f}s",
            f"  Changes:  {self.change_summary}",
        ]
        if self.warnings:
            lines.append("\n  ⚠️  Warnings:")
            for w in self.warnings:
                lines.append(f"     · {w}")
        if self.errors:
            lines.append("\n  ❌ Errors:")
            for e in self.errors:
                lines.append(f"     · {e}")
        lines.append("")
        return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────────────────────────────────

def run_pipeline(
    input_path:      str,
    target_style:    str,
    output_path:     Optional[str] = None,
    anthropic_client = None,
    skip_citations:  bool = False,
    log_fn           = print,
) -> PipelineResult:
    """
    Runs the full formatting pipeline on a document.

    Steps:
      1. Load document (.doc / .docx / .txt)
      2. Reformat citations to target style   [skippable]
      3. Apply layout profile (margins, fonts, spacing, columns)
      4. Save output .docx
      5. Return PipelineResult with full audit trail

    Args:
        input_path:      source file (.doc / .docx / .txt)
        target_style:    "ieee" | "apa7" | "harvard" | any registered style
        output_path:     where to save (default: <stem>_<style>.docx)
        anthropic_client: Anthropic client for AI citation parsing
                          (pass None to skip AI-assisted parsing)
        skip_citations:  set True to bypass citation step entirely
        log_fn:          callable for progress messages (default: print)

    Returns:
        PipelineResult
    """
    t_start  = time.time()
    warnings = []
    errors   = []
    profile  = get_profile(target_style)

    # ── Auto-name output ──────────────────────────────────────────────
    if output_path is None:
        stem        = Path(input_path).stem
        output_path = str(Path(input_path).parent / f"{stem}_{target_style}.docx")

    log_fn(f"\n📄 Pipeline start: {Path(input_path).name} → {profile.display_name}")
    log_fn(f"   Output: {output_path}")

    # ── Step 1: Load ──────────────────────────────────────────────────
    log_fn("\n[1/3] Loading document…")
    try:
        doc = load_document(input_path)
        log_fn(f"   ✅ Loaded ({len(doc.paragraphs)} paragraphs)")
    except Exception as e:
        errors.append(f"Load failed: {e}")
        log_fn(f"   ❌ {e}")
        # Return a minimal failed result
        return PipelineResult(
            output_path     = output_path,
            profile         = profile,
            citation_result = CitationResult(skipped=True),
            layout_result   = FormattedResult(
                document=Document(), profile=profile),
            warnings=warnings, errors=errors,
            elapsed_sec=time.time() - t_start,
            success=False,
        )

    # ── Step 2: Citations ─────────────────────────────────────────────
    log_fn("\n[2/3] Citation reformat…")
    if skip_citations or anthropic_client is None and not _citations_available():
        if anthropic_client is None:
            log_fn("   Skipping — no Anthropic client provided.")
        else:
            log_fn("   Skipping — skip_citations=True.")
        cit_result = CitationResult(skipped=True)
    else:
        cit_result = _try_reformat_citations(
            doc, profile.citation_style, anthropic_client, log_fn
        )

    warnings.extend(cit_result.warnings)

    # Save the citation-reformatted doc to a temp path, then reload
    # for the layout step so both steps work on the file system
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        doc.save(tmp_path)
        intermediate_path = tmp_path
    except Exception as e:
        log_fn(f"   ⚠️  Could not save intermediate doc: {e}. Using original.")
        intermediate_path = input_path

    # ── Step 3: Layout ────────────────────────────────────────────────
    log_fn("\n[3/3] Applying layout profile…")
    try:
        layout_result = format_document(
            intermediate_path,
            target_style,
            output_path,
        )
        log_fn(f"   ✅ {len(layout_result.changes)} layout changes applied.")
        warnings.extend(layout_result.warnings)
    except Exception as e:
        errors.append(f"Layout failed: {e}")
        log_fn(f"   ❌ Layout error: {e}")
        log_fn(traceback.format_exc())
        layout_result = FormattedResult(
            document=doc, profile=profile,
            warnings=[f"Layout step failed: {e}"]
        )
        success = False
    else:
        success = True
    finally:
        # Clean up temp file
        try:
            if intermediate_path != input_path:
                os.unlink(intermediate_path)
        except OSError:
            pass

    elapsed = time.time() - t_start
    log_fn(f"\n✅ Done in {elapsed:.1f}s → {output_path}")

    result = PipelineResult(
        output_path     = output_path,
        profile         = profile,
        citation_result = cit_result,
        layout_result   = layout_result,
        warnings        = warnings,
        errors          = errors,
        elapsed_sec     = elapsed,
        success         = success and not errors,
    )
    log_fn(result.report())
    return result


def _citations_available() -> bool:
    """Check whether citations.py is importable."""
    try:
        import citations  # type: ignore  # noqa: F401
        return True
    except ImportError:
        return False


# ─────────────────────────────────────────────────────────────────────
# BATCH PIPELINE
# ─────────────────────────────────────────────────────────────────────

def run_batch_pipeline(
    input_paths:     list[str],
    target_style:    str,
    output_dir:      Optional[str] = None,
    anthropic_client = None,
    skip_citations:  bool = False,
    log_fn           = print,
) -> list[PipelineResult]:
    """
    Runs the pipeline on multiple files.
    Errors on individual files are captured — other files still process.

    Args:
        input_paths:  list of source file paths
        target_style: formatting style to apply to all files
        output_dir:   directory for output files (default: same dir as input)
        anthropic_client: Anthropic client for citation parsing
        skip_citations: bypass citation step for all files
        log_fn:       logging callable

    Returns:
        list of PipelineResult, one per input file
    """
    results = []
    total   = len(input_paths)

    for i, path in enumerate(input_paths, 1):
        log_fn(f"\n── File {i}/{total}: {Path(path).name} ──")
        out = None
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            stem = Path(path).stem
            out  = os.path.join(output_dir, f"{stem}_{target_style}.docx")

        try:
            result = run_pipeline(
                input_path      = path,
                target_style    = target_style,
                output_path     = out,
                anthropic_client= anthropic_client,
                skip_citations  = skip_citations,
                log_fn          = log_fn,
            )
        except Exception as e:
            log_fn(f"   ❌ Unexpected error: {e}")
            profile = get_profile(target_style)
            result  = PipelineResult(
                output_path     = out or path,
                profile         = profile,
                citation_result = CitationResult(skipped=True),
                layout_result   = FormattedResult(
                    document=Document(), profile=profile),
                errors          = [str(e)],
                success         = False,
            )
        results.append(result)

    # Summary
    n_ok   = sum(1 for r in results if r.success)
    n_fail = total - n_ok
    log_fn(f"\n{'='*56}")
    log_fn(f"  Batch complete: {n_ok}/{total} succeeded, {n_fail} failed")
    log_fn(f"{'='*56}\n")
    return results


# ─────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    from dotenv import load_dotenv
    load_dotenv()

    if len(sys.argv) < 3:
        print("Usage: python pipeline.py <input> <style> [output]")
        print("       python pipeline.py *.docx ieee --batch --outdir ./out")
        print("Styles: ieee, apa7, harvard")
        sys.exit(1)

    inp   = sys.argv[1]
    style = sys.argv[2]
    out   = sys.argv[3] if len(sys.argv) > 3 and not sys.argv[3].startswith("--") else None

    # Optional Anthropic client
    client = None
    if os.getenv("ANTHROPIC_API_KEY"):
        try:
            import anthropic
            client = anthropic.Anthropic()
        except ImportError:
            print("⚠️  anthropic not installed — citation AI parsing disabled.")

    result = run_pipeline(inp, style, out, anthropic_client=client)
    sys.exit(0 if result.success else 1)
