"""
docx/rebuilder.py  (v2 — hardened)

Rebuilds a DOCX preserving the original's formatting with high fidelity.

v2 improvements over v1:
  - Full paragraph format application (spacing, indent, alignment)
  - Mixed-run reconstruction (bold/italic/font per run, not just plain text)
  - Table preservation: copies table XML nodes from original verbatim
  - Image preservation: copies drawing XML nodes from original verbatim
  - Page break insertion where original had them
  - Section break preservation (sectPr injection)
  - Header/footer cloning from original document relationships
  - TOC preservation (TOC paragraphs copied verbatim)
  - Detailed debug logging at every insertion step
  - Mismatch warnings when style names can't be resolved
  - Fallback chain documented in logs
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
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

log = logging.getLogger(__name__)

_WARN = "[REBUILD-MISMATCH]"
_FALLBACK = "[REBUILD-FALLBACK]"

# heading level → Word style name
HEADING_STYLE: dict[int, str] = {
    0: "Title",
    1: "Heading 1", 2: "Heading 2", 3: "Heading 3",
    4: "Heading 4", 5: "Heading 5", 6: "Heading 6",
}

ALIGN_MAP: dict[str, WD_ALIGN_PARAGRAPH] = {
    "left":    WD_ALIGN_PARAGRAPH.LEFT,
    "center":  WD_ALIGN_PARAGRAPH.CENTER,
    "right":   WD_ALIGN_PARAGRAPH.RIGHT,
    "both":    WD_ALIGN_PARAGRAPH.JUSTIFY,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "distribute": WD_ALIGN_PARAGRAPH.DISTRIBUTE,
}


# ── Body clearing ──────────────────────────────────────────────────────────────

def _clear_body(doc: DocxDocument) -> None:
    """Remove all content from body, preserving only sectPr."""
    body = doc.element.body
    keep_tags = {"sectPr"}
    to_remove = [
        child for child in body
        if (child.tag.split("}")[-1] if "}" in child.tag else child.tag) not in keep_tags
    ]
    for el in to_remove:
        body.remove(el)
    log.debug("Cleared %d body elements", len(to_remove))


# ── Paragraph formatting application ──────────────────────────────────────────

def _apply_paragraph_format(para, para_info: dict[str, Any], doc: DocxDocument) -> None:
    """
    Apply spacing, indentation, and alignment from the parsed paragraph dict.
    Uses python-docx API first; falls back to XML injection for unsupported properties.
    """
    pf = para.paragraph_format

    # Spacing
    spacing = para_info.get("spacing", {})
    if spacing:
        before = spacing.get("before")
        after  = spacing.get("after")
        line   = spacing.get("line")
        rule   = spacing.get("lineRule", "auto")
        try:
            if before is not None:
                pf.space_before = Pt(before)
            if after is not None:
                pf.space_after = Pt(after)
            if line is not None and rule == "auto":
                # line is in twip-twentieths; 240 = single spacing
                pf.line_spacing = line / 240
        except Exception as e:
            log.debug("%s spacing apply error: %s", _FALLBACK, e)

    # Indentation
    indent = para_info.get("indentation", {})
    if indent:
        try:
            if "left" in indent:
                pf.left_indent = Pt(indent["left"])
            if "right" in indent:
                pf.right_indent = Pt(indent["right"])
            if "firstLine" in indent:
                pf.first_line_indent = Pt(indent["firstLine"])
            if "hanging" in indent:
                pf.first_line_indent = Pt(-indent["hanging"])
        except Exception as e:
            log.debug("%s indent apply error: %s", _FALLBACK, e)

    # Alignment
    align_str = para_info.get("alignment")
    if align_str:
        align_val = ALIGN_MAP.get(align_str.lower())
        if align_val is not None:
            pf.alignment = align_val
        else:
            log.debug("%s Unknown alignment '%s'", _WARN, align_str)


def _apply_style(para, style_name: str, doc: DocxDocument) -> bool:
    """Apply named style; return True on success."""
    try:
        para.style = doc.styles[style_name]
        return True
    except KeyError:
        log.warning("%s Style '%s' not found in document", _WARN, style_name)
        return False


# ── Run reconstruction ─────────────────────────────────────────────────────────

def _clear_paragraph_runs(para) -> None:
    """Remove all existing run XML from a paragraph before rebuilding it."""
    for run in list(para._p.findall(qn("w:r"))):
        para._p.remove(run)


def _add_runs(para, runs: list[dict[str, Any]]) -> None:
    """
    Reconstruct inline runs with per-run formatting (bold/italic/font/size/color).
    Falls back to plain text if run data is missing.
    """
    if not runs:
        return

    for run_info in runs:
        text = run_info.get("text", "")
        if not text:
            continue
        run = para.add_run(text)

        try:
            if run_info.get("bold"):
                run.bold = True
            if run_info.get("italic"):
                run.italic = True
            if run_info.get("underline"):
                run.underline = True

            font_name = run_info.get("font")
            size_pt   = run_info.get("size_pt")
            if font_name:
                run.font.name = font_name
            if size_pt:
                run.font.size = Pt(size_pt)

            color_hex = run_info.get("color")
            if color_hex:
                # Inject color via XML (python-docx RGBColor is finicky)
                hex_val = color_hex.lstrip("#")
                rPr = run._r.get_or_add_rPr()
                color_el = OxmlElement("w:color")
                color_el.set(qn("w:val"), hex_val.upper())
                rPr.append(color_el)
        except Exception as e:
            log.debug("Run format error: %s", e)


# ── Page break insertion ───────────────────────────────────────────────────────

def _insert_page_break(doc: DocxDocument) -> None:
    """Insert an explicit page break paragraph."""
    para = doc.add_paragraph()
    run = para.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)
    log.debug("Inserted page break")


# ── Image paragraph cloning ───────────────────────────────────────────────────

def _clone_image_paragraphs(
    original_doc: DocxDocument,
    output_doc: DocxDocument,
) -> int:
    """
    Clone all paragraphs containing inline or anchored images from the original
    verbatim (deep copy of the paragraph XML node).

    Images are stored as w:drawing > wp:inline/wp:anchor inside paragraph XML.
    Their binary data lives in the document relationships — by copying from
    the original (which we shutil.copy2'd before opening), the rId references
    remain valid. python-docx cannot add images via rId, so XML-level clone
    is the only reliable approach.

    Returns count of image paragraphs cloned.
    """
    IMAGE_TAGS = (
        f"{{{qn('wp:inline').split('}')[0][1:]}}}inline",
        f"{{{qn('wp:anchor').split('}')[0][1:]}}}anchor",
    )

    body = output_doc.element.body
    sectPr = body.find(qn("w:sectPr"))
    count = 0

    for para in original_doc.paragraphs:
        has_image = (
            para._p.findall(f".//{qn('wp:inline')}", para._p.nsmap) or
            para._p.findall(f".//{qn('wp:anchor')}", para._p.nsmap)
        )
        if has_image:
            p_copy = copy.deepcopy(para._p)
            if sectPr is not None:
                body.insert(list(body).index(sectPr), p_copy)
            else:
                body.append(p_copy)
            count += 1
            log.debug("Cloned image paragraph (rIds: %s)", _image_rids(para))

    if count:
        log.info("Cloned %d image-bearing paragraphs from original", count)
    return count


def _image_rids(para) -> list[str]:
    """Extract relationship IDs from image blip elements (for debug logging)."""
    rids = []
    for blip in para._p.findall(f".//{qn('a:blip')}", para._p.nsmap):
        rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed", "")
        if rid:
            rids.append(rid)
    return rids


# ── Table XML cloning ──────────────────────────────────────────────────────────

def _clone_tables_from_original(
    original_doc: DocxDocument,
    output_doc: DocxDocument,
    structure: dict[str, Any],
) -> None:
    """
    Clone table XML nodes from the original document verbatim.
    Tables are appended at the end of the body (after all text sections).
    Future: positional table insertion using paragraph index markers.
    """
    if not original_doc.tables:
        return

    body = output_doc.element.body
    # Insert before sectPr if present
    sectPr = body.find(qn("w:sectPr"))

    for i, table in enumerate(original_doc.tables):
        try:
            table_copy = copy.deepcopy(table._tbl)
            if sectPr is not None:
                body.insert(list(body).index(sectPr), table_copy)
            else:
                body.append(table_copy)
            log.debug("Cloned table %d (%d rows)", i, len(table.rows))
        except Exception as e:
            log.warning("%s Failed to clone table %d: %s", _WARN, i, e)


# ── Heading insertion ──────────────────────────────────────────────────────────

def _add_heading(doc: DocxDocument, text: str, level: int) -> None:
    style_name = HEADING_STYLE.get(level, "Heading 1")
    # python-docx add_heading level=0 maps to Title
    effective_level = level if level > 0 else 0
    try:
        para = doc.add_heading(text, level=effective_level)
        log.debug("Added H%d: '%s'", level, text[:50])
    except Exception as e:
        log.warning("%s add_heading failed (level=%d): %s — using paragraph fallback", _WARN, level, e)
        para = doc.add_paragraph(text)
        _apply_style(para, style_name, doc)


# ── Body paragraph insertion ───────────────────────────────────────────────────

def _add_body_paragraph_from_info(
    doc: DocxDocument,
    para_info: dict[str, Any],
    style_profile: dict[str, Any],
) -> None:
    """
    Add a paragraph using its parsed para_info (original formatting preserved).
    Uses run data for mixed-format paragraphs; falls back to plain text.
    """
    text = para_info.get("text", "")
    runs = para_info.get("runs", [])
    style = para_info.get("style_name", "Normal")
    has_page_break = para_info.get("has_page_break")
    section_break = para_info.get("section_break")

    # Skip truly empty paragraphs unless they have break markers
    if not text and not runs and not has_page_break and not section_break:
        return

    para = doc.add_paragraph()
    _clear_paragraph_runs(para)

    # Apply named style
    if style and style != "Normal":
        _apply_style(para, style, doc)

    # Add text: prefer run-level reconstruction for fidelity
    if runs and any(r.get("text") for r in runs):
        _add_runs(para, runs)
    elif text:
        para.add_run(text)

    # Preserve page breaks within the paragraph
    if has_page_break:
        br_run = para.add_run()
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        br_run._r.append(br)
        log.debug("Added page break inside paragraph")

    # Preserve section breaks at paragraph end
    if section_break:
        pPr = para._p.get_or_add_pPr()
        sectPr = OxmlElement("w:sectPr")
        type_el = OxmlElement("w:type")
        type_el.set(qn("w:val"), section_break)
        sectPr.append(type_el)
        pPr.append(sectPr)
        log.debug("Added section break type=%s", section_break)

    # Apply paragraph formatting
    _apply_paragraph_format(para, para_info, doc)


def _add_ai_paragraph(
    doc: DocxDocument,
    text: str,
    style_profile: dict[str, Any],
) -> None:
    """
    Add a paragraph of AI-generated text with body style from profile.
    No run-level detail available — single run per paragraph.
    """
    para = doc.add_paragraph()
    _clear_paragraph_runs(para)
    para.add_run(text)
    _apply_style(para, "Normal", doc)

    body_style = style_profile.get("body_style", {})
    _apply_paragraph_format(para, body_style, doc)


def _split_ai_text(text: str) -> list[str]:
    """Split AI output into paragraphs on double newlines."""
    return [
        block.strip().replace("\n", " ")
        for block in text.strip().split("\n\n")
        if block.strip()
    ]


# ── Header/footer cloning ──────────────────────────────────────────────────────

def _clone_header_footer(
    original_doc: DocxDocument,
    output_doc: DocxDocument,
) -> None:
    """
    Clone header and footer XML from original document sections into output.
    Only processes sections that have non-linked headers/footers.
    """
    try:
        for i, (orig_sec, out_sec) in enumerate(
            zip(original_doc.sections, output_doc.sections)
        ):
            # Header
            try:
                if not orig_sec.header.is_linked_to_previous:
                    orig_hdr = orig_sec.header._element
                    out_hdr = out_sec.header._element
                    # Clear output header content
                    for child in list(out_hdr):
                        out_hdr.remove(child)
                    # Copy original header content
                    for child in orig_hdr:
                        out_hdr.append(copy.deepcopy(child))
                    log.debug("Cloned header for section %d", i)
            except Exception as e:
                log.debug("%s Header clone error (section %d): %s", _WARN, i, e)

            # Footer
            try:
                if not orig_sec.footer.is_linked_to_previous:
                    orig_ftr = orig_sec.footer._element
                    out_ftr = out_sec.footer._element
                    for child in list(out_ftr):
                        out_ftr.remove(child)
                    for child in orig_ftr:
                        out_ftr.append(copy.deepcopy(child))
                    log.debug("Cloned footer for section %d", i)
            except Exception as e:
                log.debug("%s Footer clone error (section %d): %s", _WARN, i, e)
    except Exception as e:
        log.warning("%s Header/footer clone failed: %s", _WARN, e)


# ── TOC paragraph cloning ──────────────────────────────────────────────────────

def _clone_toc_paragraphs(
    original_doc: DocxDocument,
    output_doc: DocxDocument,
) -> int:
    """
    Clone TOC paragraphs verbatim from original.
    Returns count of cloned TOC paragraphs.
    Inserts after preamble (at current body position).
    """
    TOC_PREFIXES = ("toc", "table of contents", "contents")
    body = output_doc.element.body
    sectPr = body.find(qn("w:sectPr"))
    count = 0

    for para in original_doc.paragraphs:
        style_name = (para.style.name or "").lower()
        if any(style_name.startswith(p) for p in TOC_PREFIXES):
            p_copy = copy.deepcopy(para._p)
            if sectPr is not None:
                body.insert(list(body).index(sectPr), p_copy)
            else:
                body.append(p_copy)
            count += 1

    if count:
        log.debug("Cloned %d TOC paragraphs", count)
    return count


# ── Main entry point ───────────────────────────────────────────────────────────

def rebuild_docx(
    original_path: str | Path,
    output_path: str | Path,
    structure: dict[str, Any],
    ai_results: dict[str, str],
    style_profile: dict[str, Any],
) -> Path:
    """
    Rebuild a DOCX with AI content injected, formatting fully preserved.

    Strategy:
      1. Copy original → output (clones styles.xml, theme, fonts, relationships)
      2. Clear body content (keep sectPr for margins/layout)
      3. Clone header/footer from original
      4. Re-insert preamble (original content verbatim)
      5. Clone TOC paragraphs if present
      6. For each section: insert heading + AI content or original paragraphs
      7. Clone tables from original (verbatim XML copy)
      8. Save and return path
    """
    original_path = Path(original_path)
    output_path   = Path(output_path)

    # ── 1. Copy ──────────────────────────────────────────────────────────────
    shutil.copy2(original_path, output_path)
    log.info("Copied original → %s", output_path.name)

    original_doc = DocxDocument(str(original_path))
    output_doc   = DocxDocument(str(output_path))

    # ── 2. Clear body ────────────────────────────────────────────────────────
    _clear_body(output_doc)

    # ── 3. Clone header/footer ───────────────────────────────────────────────
    _clone_header_footer(original_doc, output_doc)

    # ── 4. Preamble ──────────────────────────────────────────────────────────
    preamble = structure.get("preamble", [])
    for para_info in preamble:
        _add_body_paragraph_from_info(output_doc, para_info, style_profile)
    if preamble:
        log.debug("Inserted %d preamble paragraphs", len(preamble))

    # ── 5. TOC ───────────────────────────────────────────────────────────────
    if structure.get("has_toc"):
        toc_count = _clone_toc_paragraphs(original_doc, output_doc)
        log.info("TOC present: cloned %d paragraphs", toc_count)

    # ── 6. Sections ──────────────────────────────────────────────────────────
    sections = structure.get("sections", [])
    ai_sections = 0
    preserved_sections = 0

    for section in sections:
        heading_text  = section["heading_text"]
        heading_level = section["heading_level"]

        _add_heading(output_doc, heading_text, heading_level)

        if heading_text in ai_results and ai_results[heading_text]:
            # AI-generated content
            for para_text in _split_ai_text(ai_results[heading_text]):
                _add_ai_paragraph(output_doc, para_text, style_profile)
            ai_sections += 1
            log.debug("AI content inserted under '%s'", heading_text)
        else:
            # Original content preserved with full formatting
            original_paras = section.get("paragraphs", [])
            inserted = 0
            for para_info in original_paras:
                if (
                    not para_info.get("is_empty")
                    or para_info.get("has_page_break")
                    or para_info.get("section_break")
                ):
                    _add_body_paragraph_from_info(output_doc, para_info, style_profile)
                    inserted += 1
            preserved_sections += 1
            log.debug("Preserved %d paragraphs under '%s'", inserted, heading_text)

    log.info(
        "Sections: %d AI-generated, %d preserved from original",
        ai_sections, preserved_sections,
    )

    # ── 7. Tables ────────────────────────────────────────────────────────────
    if original_doc.tables:
        _clone_tables_from_original(original_doc, output_doc, structure)
        log.info("Cloned %d tables from original", len(original_doc.tables))

    # ── 7b. Images ───────────────────────────────────────────────────────────
    img_count = _clone_image_paragraphs(original_doc, output_doc)
    if img_count:
        log.info("Cloned %d image paragraphs from original", img_count)

    # ── 8. Save ──────────────────────────────────────────────────────────────
    output_doc.save(str(output_path))
    log.info("Saved rebuilt DOCX → %s (%d bytes)", output_path.name, output_path.stat().st_size)
    return output_path


# ── Validation ─────────────────────────────────────────────────────────────────

def validate_output(output_path: str | Path) -> bool:
    """Re-open and sanity-check the output DOCX."""
    try:
        doc = DocxDocument(str(output_path))
        para_count  = len(doc.paragraphs)
        table_count = len(doc.tables)
        if para_count == 0:
            log.error("Validation FAILED: 0 paragraphs in output")
            return False
        log.info("Validation OK: %d paragraphs, %d tables", para_count, table_count)
        return True
    except Exception as e:
        log.error("Validation FAILED: %s", e)
        return False
