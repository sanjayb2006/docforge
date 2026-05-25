"""
app/services/docx/transformer.py

Step 1 — High-Fidelity In-Place Document Transformer.

PHILOSOPHY
----------
The rebuilder clears the document body and reconstructs it from JSON.
This transformer never clears anything. It works on the original XML
directly, replacing only the text content of target sections while
leaving every surrounding XML node untouched.

WHAT IT DOES
------------
1. shutil.copy2(original → output)         — start with an exact binary copy
2. Open only output_doc (original untouched on disk)
3. Build a section index map from live paragraph objects
4. For each section targeted by AI:
   a. Locate the heading paragraph by index (live object, no recreation)
   b. Locate the body paragraphs that follow it
   c. Replace text in those paragraphs using _replace_text_in_section()
5. Save — styles.xml, theme, relationships, margins, headers/footers,
   images, tables, field codes, bookmarks all survive because we never
   touched them.

WHAT IT NEVER DOES
------------------
- Never calls _clear_body()
- Never calls doc.add_paragraph() for preserved content
- Never calls doc.add_heading()
- Never calls doc.add_run() for preserved runs
- Never touches paragraphs outside targeted sections
- Never reconstructs pPr (paragraph properties)
- Never reconstructs rPr (run properties) for preserved sections

TEXT REPLACEMENT STRATEGY
--------------------------
The AI generates N characters. The original section has M characters
spread across K paragraphs, each with their own run structure.

We use a "drain and fill" approach:

  Step A — drain: collect all body paragraphs for the section
  Step B — split AI text into logical paragraphs (double-newline split)
  Step C — for each AI paragraph:
    - If a corresponding original paragraph exists: reuse it
      * Clear only the w:t text nodes inside existing runs
      * Fill first run with text, clear the rest
      * Preserve ALL w:pPr and w:rPr untouched
    - If more AI paragraphs than original: clone the last original
      paragraph's XML structure (deep copy pPr + rPr), insert after it
  Step D — if fewer AI paragraphs than original: remove excess paragraphs
           from the body XML (they are genuinely gone, not just empty)

This means:
  - Formatting of AI paragraphs = formatting of the original paragraph
    they replaced (correct font, size, spacing, indent, alignment)
  - No style drift for text that fits in existing paragraphs
  - New overflow paragraphs inherit the last original paragraph's style

SECTION INDEX MAP
-----------------
Built once per run from live document paragraph objects:

  {
    "1. Aim": {
        "heading_index": 3,
        "body_indices": [4, 5],
    },
    "2. Theory": {
        "heading_index": 6,
        "body_indices": [7, 8, 9],
    }
  }

This is derived from the parsed structure JSON (already in DB) mapped
to live paragraph indices — no re-parsing of the DOCX needed.

FALLBACK BEHAVIOUR
------------------
If anything goes wrong for a specific section (XML corruption, unexpected
structure), that section is logged as a warning and the original content
is left intact. The transformer never fails the whole job due to one
bad section.

PARALLEL DEPLOYMENT
-------------------
This module is deployed alongside rebuilder.py. The rewrite pipeline
selects which to use based on a flag. Both write the same output format.
"""

from __future__ import annotations

import copy
import logging
import shutil
from pathlib import Path
from typing import Any

from docx import Document as DocxDocument
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

log = logging.getLogger(__name__)

_TAG  = "[TRANSFORM]"
_WARN = "[TRANSFORM-WARN]"


# ══════════════════════════════════════════════════════════════════════════════
# SECTION INDEX MAP
# ══════════════════════════════════════════════════════════════════════════════

def build_section_index(
    doc: DocxDocument,
    structure: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    """
    Build a map of {heading_text: {heading_index, body_indices}} from the
    parsed structure JSON, validated against live paragraph objects.

    The structure JSON has paragraph indices recorded at parse time.
    We trust those indices — they are stable for a given file.

    Returns only sections that are actually findable in the live document.
    Sections where the heading_index is out of range are logged and skipped.
    """
    paragraphs = doc.paragraphs
    total = len(paragraphs)
    index_map: dict[str, dict[str, Any]] = {}

    for section in structure.get("sections", []):
        heading_text  = section.get("heading_text", "")
        heading_index = section.get("heading_index", -1)

        if heading_index < 0 or heading_index >= total:
            log.warning(
                "%s Section '%s' heading_index=%d out of range (doc has %d paras) — skipped",
                _WARN, heading_text[:50], heading_index, total,
            )
            continue

        # Validate the heading text matches what we expect
        live_text = paragraphs[heading_index].text.strip()
        if live_text != heading_text.strip():
            log.warning(
                "%s Section '%s' index mismatch: found '%s' at index %d — skipped",
                _WARN, heading_text[:50], live_text[:50], heading_index,
            )
            continue

        # Collect body paragraph indices from the structure
        body_indices = [
            p["index"]
            for p in section.get("paragraphs", [])
            if p["index"] < total
        ]

        index_map[heading_text] = {
            "heading_index": heading_index,
            "body_indices":  body_indices,
        }
        log.debug(
            "%s Mapped '%s' → heading[%d] body[%s]",
            _TAG, heading_text[:50], heading_index,
            f"{body_indices[0]}..{body_indices[-1]}" if body_indices else "empty",
        )

    log.info(
        "%s Section index built: %d/%d sections mapped",
        _TAG, len(index_map), len(structure.get("sections", [])),
    )
    return index_map


# ══════════════════════════════════════════════════════════════════════════════
# TEXT REPLACEMENT — PARAGRAPH LEVEL
# ══════════════════════════════════════════════════════════════════════════════

def _get_para_text_length(para) -> int:
    """Total character count of a paragraph's text."""
    return sum(len(run.text) for run in para.runs)


def _clear_para_text(para) -> None:
    """
    Set all w:t text nodes in a paragraph to empty string.
    Preserves ALL w:rPr and w:pPr — only w:t content is cleared.
    """
    for t_el in para._p.findall(f".//{qn('w:t')}"):
        t_el.text = ""


def _set_first_run_text(para, text: str) -> bool:
    """
    Write text into the first run of a paragraph.
    If the paragraph has no runs, add one — but only in the emergencyfallback
    case (no existing run structure to reuse). Returns True on success.
    """
    runs = para.runs
    if not runs:
        # No runs at all — paragraph is structural only (e.g., image container)
        # Don't touch it
        log.debug("%s Para has no runs — skipping text injection", _WARN)
        return False

    # First run gets the full text
    # Preserves its w:rPr exactly — only w:t changes
    first_run = runs[0]
    first_run.text = text

    # All subsequent runs in this paragraph: clear their text
    # (their rPr is preserved, they just become empty)
    for run in runs[1:]:
        run.text = ""

    return True


def _clone_paragraph_structure(para) -> Any:
    """
    Deep-copy a paragraph's XML node.
    Used when we need more paragraphs than the original had.
    The clone gets the same pPr and rPr but its text is cleared.
    """
    p_clone = copy.deepcopy(para._p)
    # Clear all text in the clone
    for t_el in p_clone.findall(f".//{qn('w:t')}"):
        t_el.text = ""
    return p_clone


def _insert_para_after(doc: DocxDocument, ref_para, new_p_xml) -> Any:
    """
    Insert a new paragraph XML node immediately after ref_para in the body.
    Returns the inserted element.
    """
    body = doc.element.body
    ref_p = ref_para._p
    ref_idx = list(body).index(ref_p)
    body.insert(ref_idx + 1, new_p_xml)
    return new_p_xml


def _insert_para_before(doc: DocxDocument, ref_para, new_p_xml) -> Any:
    """
    Insert a new paragraph XML node immediately before ref_para in the body.
    Returns the inserted element.
    """
    body = doc.element.body
    ref_p = ref_para._p
    ref_idx = list(body).index(ref_p)
    body.insert(ref_idx, new_p_xml)
    return new_p_xml


def _paragraph_has_section_break(para) -> bool:
    pPr = para._p.find(qn("w:pPr"))
    return pPr is not None and pPr.find(qn("w:sectPr")) is not None


def _remove_paragraph(doc: DocxDocument, para) -> None:
    """Remove a paragraph's XML node from the body entirely."""
    body = doc.element.body
    p = para._p
    if p in body:
        body.remove(p)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION TEXT REPLACEMENT
# ══════════════════════════════════════════════════════════════════════════════

def _split_ai_text(text: str) -> list[str]:
    """
    Split AI-generated text into logical paragraphs.
    Double newline = paragraph boundary.
    Single newlines within a block are collapsed to space.
    Empty blocks are discarded.
    """
    return [
        block.strip().replace("\n", " ")
        for block in text.strip().split("\n\n")
        if block.strip()
    ]


def _replace_section_text(
    doc:          DocxDocument,
    body_indices: list[int],
    ai_text:      str,
    section_name: str,
    heading_index: int | None = None,
) -> None:
    """
    Replace text in a section's body paragraphs with AI-generated content.

    Preserves ALL formatting — pPr, rPr, spacing, alignment, fonts.
    Only w:t text content changes.

    Args:
        doc:          The output document (modified in place)
        body_indices: List of paragraph indices for this section's body
        ai_text:      AI-generated plain text for this section
        section_name: For logging only
        heading_index: Paragraph index of the section heading, if available
    """
    ai_paragraphs = _split_ai_text(ai_text)

    if not ai_paragraphs:
        log.warning(
            "%s Section '%s': AI text is empty — original content preserved",
            _WARN, section_name[:50],
        )
        return

    # Get live paragraph objects for this section
    # Important: we snapshot the list now because we may add/remove elements
    all_paras = doc.paragraphs
    body_paras = []
    for idx in body_indices:
        if idx < len(all_paras):
            para = all_paras[idx]
            # Skip image/drawing paragraphs — never overwrite them
            has_drawing = bool(
                para._p.findall(f".//{qn('wp:inline')}", para._p.nsmap) or
                para._p.findall(f".//{qn('wp:anchor')}", para._p.nsmap)
            )
            if has_drawing:
                log.debug(
                    "%s Para[%d] has image — preserved in-place", _TAG, idx
                )
                continue

            if _paragraph_has_section_break(para) or not para.runs:
                log.debug(
                    "%s Para[%d] is structural or section-break only — preserved in-place",
                    _TAG,
                    idx,
                )
                continue

            body_paras.append(para)

    if not body_paras:
        # Section has no writable body paragraphs
        # This happens with sections that only had images, or empty sections
        # For empty sections: inject one new paragraph using Normal style
        log.info(
            "%s Section '%s': no writable body paras — injecting after heading",
            _TAG, section_name[:50],
        )
        _inject_after_heading(doc, heading_index, body_indices, ai_paragraphs, section_name)
        return

    n_orig = len(body_paras)
    n_ai   = len(ai_paragraphs)

    log.debug(
        "%s '%s': %d original body paras → %d AI paragraphs",
        _TAG, section_name[:50], n_orig, n_ai,
    )

    last_inserted = None
    for i, ai_para_text in enumerate(ai_paragraphs):
        if i < n_orig:
            # Reuse original paragraph i — formatting preserved entirely
            ok = _set_first_run_text(body_paras[i], ai_para_text)
            if not ok:
                log.debug(
                    "%s Para[%d] has no runs — skipped",
                    _WARN, body_indices[i] if i < len(body_indices) else -1,
                )
            last_inserted = body_paras[i]
        else:
            # ── Case 2: More AI paras than original — clone and insert ────────
            last_orig = body_paras[n_orig - 1]
            p_clone = _clone_paragraph_structure(last_orig)

            # Set text in the clone's first run
            for t_el in p_clone.findall(f".//{qn('w:t')}"):
                t_el.text = ai_para_text
                break  # only the first w:t gets the text

            if last_inserted is None:
                last_inserted = last_orig

            # Insert clone in the same section as the last original paragraph.
            # If the last paragraph carries a section break, insert before it
            # so the cloned overflow paragraphs remain in the same section.
            if _paragraph_has_section_break(last_inserted):
                last_inserted = _insert_para_before(doc, last_inserted, p_clone)
            else:
                last_inserted = _insert_para_after(doc, last_inserted, p_clone)

    # ── Case 3: Fewer AI paras than original — remove excess ──────────────────
    if n_orig > n_ai:
        # Paragraphs body_paras[n_ai:] are excess — remove them from XML
        for excess_para in body_paras[n_ai:]:
            if _paragraph_has_section_break(excess_para):
                log.debug(
                    "%s Preserving section-break paragraph instead of removing excess",
                    _TAG,
                )
                _set_first_run_text(excess_para, "")
                continue

            try:
                _remove_paragraph(doc, excess_para)
                log.debug("%s Removed excess paragraph", _TAG)
            except Exception as e:
                log.debug("%s Could not remove excess para: %s", _WARN, e)


def _inject_after_heading(
    doc:            DocxDocument,
    heading_index:  int | None,
    body_indices:   list[int],
    ai_paragraphs:  list[str],
    section_name:   str,
) -> None:
    """
    Fallback for sections with no writable body paragraphs.
    Injects new Normal paragraphs after the heading or after the last known
    body index while preserving section boundaries.

    This is the only place the transformer creates new paragraph XML,
    and it only happens for genuinely empty sections.
    """
    body = doc.element.body
    all_paras = doc.paragraphs
    insert_after = None
    insert_before = None

    if heading_index is not None and 0 <= heading_index < len(all_paras):
        insert_after = all_paras[heading_index]
    elif body_indices:
        last_index = max(body_indices)
        if last_index < len(all_paras):
            insert_after = all_paras[last_index]

    if insert_after is not None and _paragraph_has_section_break(insert_after):
        insert_before = insert_after
        insert_after = None

    for text in ai_paragraphs:
        p_el = OxmlElement("w:p")
        r_el = OxmlElement("w:r")
        t_el = OxmlElement("w:t")
        t_el.text = text
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r_el.append(t_el)
        p_el.append(r_el)

        if insert_before is not None:
            insert_before = _insert_para_before(doc, insert_before, p_el)
        elif insert_after is not None:
            insert_after = _insert_para_after(doc, insert_after, p_el)
        else:
            sectPr = body.find(qn("w:sectPr"))
            if sectPr is not None:
                body.insert(list(body).index(sectPr), p_el)
            else:
                body.append(p_el)

    log.info(
        "%s Injected %d new paras for empty section '%s'",
        _TAG, len(ai_paragraphs), section_name[:50],
    )


def transform_docx(
    original_path: str | Path,
    output_path:   str | Path,
    structure:     dict[str, Any],
    ai_results:    dict[str, str],
    style_profile: dict[str, Any] | None = None,
) -> Path:
    """
    Transform the original DOCX in-place using AI-generated section text.

    This transformer copies the original file, opens the copy, and applies text
    replacements in-place so that formatting, images, headers, footers, and
    section properties are preserved.
    """
    original_path = Path(original_path)
    output_path = Path(output_path)

    if not ai_results:
        shutil.copy2(original_path, output_path)
        log.info("%s No AI results — original copied unchanged", _TAG)
        return output_path

    shutil.copy2(original_path, output_path)
    log.info("%s Copied original → %s", _TAG, output_path.name)

    doc = DocxDocument(str(output_path))
    orig_para_count = len(doc.paragraphs)

    index_map = build_section_index(doc, structure)
    sections_applied = 0
    sections_skipped = 0

    for heading_text, ai_text in ai_results.items():
        mapping = index_map.get(heading_text)
        if mapping is None:
            log.warning(
                "%s Section '%s' not mapped in live document — skipped",
                _WARN, heading_text[:50],
            )
            sections_skipped += 1
            continue

        body_indices = mapping["body_indices"]
        heading_index = mapping.get("heading_index")

        try:
            _replace_section_text(
                doc=doc,
                body_indices=body_indices,
                ai_text=ai_text,
                section_name=heading_text,
                heading_index=heading_index,
            )
            sections_applied += 1
            log.info(
                "%s Applied AI to '%s' (%d body paras, %d AI chars)",
                _TAG,
                heading_text[:50],
                len(body_indices),
                len(ai_text),
            )
        except Exception as exc:
            log.warning(
                "%s Failed to apply AI to '%s': %s — original preserved",
                _WARN, heading_text[:50], exc,
            )
            sections_skipped += 1

    log.info(
        "%s Transform complete: %d applied, %d skipped",
        _TAG, sections_applied, sections_skipped,
    )

    out_para_count = len(doc.paragraphs)
    doc.save(str(output_path))

    log.info(
        "%s Saved → %s  paras: %d→%d (delta=%+d)  size: %d bytes",
        _TAG,
        output_path.name,
        orig_para_count,
        out_para_count,
        out_para_count - orig_para_count,
        output_path.stat().st_size,
    )

    return output_path


def validate_output(output_path: str | Path) -> bool:
    """Re-open and sanity-check the transformed DOCX."""
    try:
        doc = DocxDocument(str(output_path))
        para_count  = len(doc.paragraphs)
        table_count = len(doc.tables)
        if para_count == 0:
            log.error("%s Validation FAILED: 0 paragraphs", _TAG)
            return False
        log.info(
            "%s Validation OK: %d paragraphs, %d tables",
            _TAG, para_count, table_count,
        )
        return True
    except Exception as e:
        log.error("%s Validation FAILED: %s", _TAG, e)
        return False
