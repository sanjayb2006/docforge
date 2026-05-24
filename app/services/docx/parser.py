"""
docx/parser.py  (v2 — hardened)

Extracts the full logical + visual structure of a DOCX:
  - headings (level, text, index, numbering)
  - body paragraphs with per-run detail (bold/italic/font/size)
  - indentation (left/right/firstLine/hanging)
  - tables (cell text, row/col counts, style)
  - images (relationship IDs, EMU dimensions, inline vs anchor)
  - page breaks and section breaks (explicit + implicit)
  - TOC paragraphs (detected by style name)
  - list items (numId + indent level)
  - paragraph-level spacing (before/after/line)

v2 changes vs v1:
  - per-paragraph spacing extracted (was missing)
  - indentation captured per paragraph
  - runs captured (inline bold/italic/font overrides)
  - page break detection (w:lastRenderedPageBreak, w:br type=page)
  - section break detection (sectPr inside pPr)
  - tables extracted to structure (was only counted)
  - images inventoried with size info
  - list detection via numPr
  - TOC style detection
  - debug logging at every extraction step
"""

from __future__ import annotations
import logging
from pathlib import Path
from typing import Any

from docx import Document as DocxDocument
from docx.oxml.ns import qn
from lxml import etree

log = logging.getLogger(__name__)

# ── Constants ──────────────────────────────────────────────────────────────────

HEADING_STYLE_MAP: dict[str, int] = {
    "heading 1": 1, "heading 2": 2, "heading 3": 3,
    "heading 4": 4, "heading 5": 5, "heading 6": 6,
    "title": 0,
}

TOC_STYLE_PREFIXES = ("toc", "table of contents", "contents")


# ── Low-level XML helpers ──────────────────────────────────────────────────────

def _pt(twips: str | None) -> float | None:
    """Half-twips or twips string → points. Returns None on any failure."""
    if twips is None:
        return None
    try:
        return round(int(twips) / 20, 2)
    except (ValueError, TypeError):
        return None


def _half_pt(half_pts: str | None) -> float | None:
    """Half-points string → points."""
    if half_pts is None:
        return None
    try:
        return round(int(half_pts) / 2, 2)
    except (ValueError, TypeError):
        return None


def _bool_elem(el, tag: str) -> bool:
    """True if <tag> exists AND its w:val is not '0' or 'false'."""
    child = el.find(qn(tag))
    if child is None:
        return False
    val = child.get(qn("w:val"), "true")
    return val.lower() not in ("0", "false")


# ── Run extraction ─────────────────────────────────────────────────────────────

def _extract_runs(para) -> list[dict[str, Any]]:
    """
    Extract inline run properties for a paragraph.
    Captures font, size, bold, italic, underline, color per run.
    Used to detect mixed formatting (a common source of fidelity loss).
    """
    runs = []
    for run in para.runs:
        rPr = run._r.find(qn("w:rPr"))
        run_info: dict[str, Any] = {"text": run.text}

        if rPr is not None:
            # Font
            fonts_el = rPr.find(qn("w:rFonts"))
            if fonts_el is not None:
                for attr in ("ascii", "hAnsi", "cs"):
                    val = fonts_el.get(qn(f"w:{attr}"))
                    if val:
                        run_info["font"] = val
                        break

            # Size
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                run_info["size_pt"] = _half_pt(sz.get(qn("w:val")))

            # Bold / italic
            if _bool_elem(rPr, "w:b"):
                run_info["bold"] = True
            if _bool_elem(rPr, "w:i"):
                run_info["italic"] = True

            # Color
            color_el = rPr.find(qn("w:color"))
            if color_el is not None:
                v = color_el.get(qn("w:val"), "")
                if v and v.upper() != "AUTO":
                    run_info["color"] = f"#{v.upper()}"

            # Underline
            u = rPr.find(qn("w:u"))
            if u is not None:
                run_info["underline"] = u.get(qn("w:val"), "single")

            # Highlight
            hl = rPr.find(qn("w:highlight"))
            if hl is not None:
                run_info["highlight"] = hl.get(qn("w:val"), "")

        if run_info["text"]:          # skip empty runs
            runs.append(run_info)

    return runs


# ── Paragraph-level extraction ─────────────────────────────────────────────────

def _resolve_heading_level(para) -> int | None:
    style_name = (para.style.name or "").lower().strip()
    if style_name in HEADING_STYLE_MAP:
        return HEADING_STYLE_MAP[style_name]
    # Outline level fallback (custom heading styles)
    pPr = para._p.find(qn("w:pPr"))
    if pPr is not None:
        ol = pPr.find(qn("w:outlineLvl"))
        if ol is not None:
            v = ol.get(qn("w:val"))
            if v is not None:
                try:
                    return int(v) + 1
                except ValueError:
                    pass
    return None


def _extract_indentation(pPr) -> dict[str, float]:
    """Extract paragraph indentation in points."""
    indent: dict[str, float] = {}
    if pPr is None:
        return indent
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return indent
    for attr in ("left", "right", "firstLine", "hanging"):
        v = ind.get(qn(f"w:{attr}"))
        pt = _pt(v)
        if pt is not None and pt != 0:
            indent[attr] = pt
    return indent


def _extract_para_spacing(pPr) -> dict[str, Any]:
    """Extract before/after/line spacing from paragraph properties."""
    spacing: dict[str, Any] = {}
    if pPr is None:
        return spacing
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        return spacing
    for attr in ("before", "after", "line"):
        v = sp.get(qn(f"w:{attr}"))
        pt = _pt(v)
        if pt is not None:
            spacing[attr] = pt
    lr = sp.get(qn("w:lineRule"))
    if lr:
        spacing["lineRule"] = lr
    return spacing


def _has_page_break(para) -> bool:
    """True if paragraph contains an explicit page break."""
    for br in para._p.findall(f".//{qn('w:br')}"):
        br_type = br.get(qn("w:type"), "")
        if br_type in ("page", "column"):
            return True
    return False


def _has_section_break(para) -> str | None:
    """
    Returns section break type string if paragraph ends with a section break,
    else None. Types: nextPage, oddPage, evenPage, continuous, nextColumn.
    """
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return None
    sectPr = pPr.find(qn("w:sectPr"))
    if sectPr is None:
        return None
    pgType = sectPr.find(qn("w:type"))
    if pgType is not None:
        return pgType.get(qn("w:val"), "nextPage")
    return "nextPage"


def _is_toc_paragraph(para) -> bool:
    style_name = (para.style.name or "").lower().strip()
    return any(style_name.startswith(p) for p in TOC_STYLE_PREFIXES)


def _extract_list_info(pPr) -> dict[str, int] | None:
    """Extract list numId + indent level, or None if not a list item."""
    if pPr is None:
        return None
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return None
    ilvl_el = numPr.find(qn("w:ilvl"))
    numId_el = numPr.find(qn("w:numId"))
    if ilvl_el is None or numId_el is None:
        return None
    try:
        return {
            "ilvl": int(ilvl_el.get(qn("w:val"), 0)),
            "numId": int(numId_el.get(qn("w:val"), 0)),
        }
    except (ValueError, TypeError):
        return None


def _alignment_name(para) -> str | None:
    try:
        align = para.alignment
        return align.name if align else None
    except Exception:
        return None


# ── Table extraction ───────────────────────────────────────────────────────────

def _extract_tables(doc: DocxDocument) -> list[dict[str, Any]]:
    """
    Extract all tables: row count, col count, style, and cell text matrix.
    For large tables (>10 rows), only first 10 rows are captured in full;
    a 'truncated' flag is set.
    """
    tables_out = []
    for t_idx, table in enumerate(doc.tables):
        row_count = len(table.rows)
        col_count = max((len(r.cells) for r in table.rows), default=0)
        style_name = ""
        try:
            style_name = table.style.name or ""
        except Exception:
            pass

        max_rows = min(row_count, 10)
        cells: list[list[str]] = []
        for row in table.rows[:max_rows]:
            cells.append([c.text.strip() for c in row.cells])

        tables_out.append({
            "table_index": t_idx,
            "row_count": row_count,
            "col_count": col_count,
            "style": style_name,
            "cells": cells,
            "truncated": row_count > 10,
        })
        log.debug("Table %d: %dx%d, style='%s'", t_idx, row_count, col_count, style_name)

    return tables_out


# ── Image extraction ───────────────────────────────────────────────────────────

def _extract_images(doc: DocxDocument) -> list[dict[str, Any]]:
    """
    Extract image metadata: relationship ID, dimensions (EMU → pt), paragraph index.
    Does NOT extract binary image data — that stays in the DOCX relationships.
    """
    images = []
    for p_idx, para in enumerate(doc.paragraphs):
        for drawing in para._p.findall(f".//{qn('a:blip')}", para._p.nsmap):
            rId = drawing.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed", "")
            # Walk up to find extent
            extent = drawing.getparent()
            while extent is not None and extent.tag != qn("wp:extent"):
                # Look for extent sibling in the inline/anchor container
                parent = extent.getparent()
                if parent is None:
                    break
                cx = parent.find(qn("wp:extent"))
                if cx is not None:
                    extent = cx
                    break
                extent = parent

            cx_val = cy_val = None
            if extent is not None and hasattr(extent, "get"):
                try:
                    cx_val = round(int(extent.get("cx", 0)) / 12700, 1)
                    cy_val = round(int(extent.get("cy", 0)) / 12700, 1)
                except (ValueError, TypeError):
                    pass

            images.append({
                "paragraph_index": p_idx,
                "relationship_id": rId,
                "width_pt": cx_val,
                "height_pt": cy_val,
            })
            log.debug("Image rId=%s at para %d (%s×%s pt)", rId, p_idx, cx_val, cy_val)

    return images


# ── Main ───────────────────────────────────────────────────────────────────────

def parse_docx(file_path: str | Path) -> dict[str, Any]:
    """
    Parse a DOCX and return a full DocumentStructure dict (JSON-safe).

    Structure keys:
      title, sections, preamble, tables, images,
      table_count, image_count, total_paragraphs,
      has_toc, page_break_count, section_break_count
    """
    file_path = Path(file_path)
    log.info("Parsing DOCX: %s", file_path.name)

    doc = DocxDocument(str(file_path))

    sections_out: list[dict] = []
    preamble_out: list[dict] = []
    current_section: dict | None = None
    in_preamble = True
    title = "Untitled"
    has_toc = False
    page_break_count = 0
    section_break_count = 0

    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        style_name = para.style.name or "Normal"
        level = _resolve_heading_level(para)
        pPr = para._p.find(qn("w:pPr"))

        # Structural flags
        has_pb = _has_page_break(para)
        sb_type = _has_section_break(para)
        is_toc = _is_toc_paragraph(para)
        list_info = _extract_list_info(pPr)

        if has_pb:
            page_break_count += 1
            log.debug("Page break at paragraph %d", idx)
        if sb_type:
            section_break_count += 1
            log.debug("Section break (%s) at paragraph %d", sb_type, idx)
        if is_toc:
            has_toc = True

        para_dict: dict[str, Any] = {
            "index": idx,
            "text": text,
            "style_name": style_name,
            "heading_level": level,
            "alignment": _alignment_name(para),
            "is_empty": (text == ""),
            "is_toc": is_toc,
            "has_page_break": has_pb,
            "section_break": sb_type,
            "indentation": _extract_indentation(pPr),
            "spacing": _extract_para_spacing(pPr),
            "list_info": list_info,
            "runs": _extract_runs(para),
        }

        if level is not None:
            in_preamble = False
            if level == 0 and title == "Untitled":
                title = text

            current_section = {
                "heading_index": idx,
                "heading_level": level,
                "heading_text": text,
                "paragraphs": [],
            }
            sections_out.append(current_section)
            log.debug("Section H%d: '%s'", level, text[:60])
        else:
            if in_preamble:
                preamble_out.append(para_dict)
            elif current_section is not None:
                current_section["paragraphs"].append(para_dict)
            elif sections_out:
                sections_out[-1]["paragraphs"].append(para_dict)
            else:
                preamble_out.append(para_dict)

    if title == "Untitled" and sections_out:
        title = sections_out[0]["heading_text"]

    tables = _extract_tables(doc)
    images = _extract_images(doc)

    result = {
        "title": title,
        "sections": sections_out,
        "preamble": preamble_out,
        "tables": tables,
        "images": images,
        "table_count": len(doc.tables),
        "image_count": len(images),
        "total_paragraphs": len(doc.paragraphs),
        "has_toc": has_toc,
        "page_break_count": page_break_count,
        "section_break_count": section_break_count,
    }

    log.info(
        "Parsed '%s': %d sections, %d tables, %d images, "
        "%d page-breaks, %d section-breaks, toc=%s",
        file_path.name, len(sections_out), len(tables), len(images),
        page_break_count, section_break_count, has_toc,
    )
    return result
