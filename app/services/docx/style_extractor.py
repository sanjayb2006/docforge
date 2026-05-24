"""
docx/style_extractor.py  (v2 — hardened)

Extracts a complete StyleProfile from a DOCX.

v2 improvements:
  - Resolved font inheritance chain (style → parent → docDefaults)
  - Header/footer text + detected fields (page numbers, dates)
  - Custom font inventory (all unique fonts in use, not just heading fonts)
  - Paragraph indent profile (first-line, left, hanging)
  - Tab stop extraction
  - TOC style detection
  - Per-section page layout (for multi-section docs)
  - Debug logging on every extraction with mismatch warnings
  - Fallback chain logged explicitly so rebuild can replicate it
"""

from __future__ import annotations
import logging
from pathlib import Path
from typing import Any

from docx import Document as DocxDocument
from docx.oxml.ns import qn

log = logging.getLogger(__name__)

_WARN = "[STYLE-MISMATCH]"
_FALLBACK = "[STYLE-FALLBACK]"


# ── Unit helpers ───────────────────────────────────────────────────────────────

def _twips_pt(v: str | None) -> float | None:
    if v is None:
        return None
    try:
        return round(int(v) / 20, 2)
    except (ValueError, TypeError):
        return None


def _half_pt(v: str | None) -> float | None:
    if v is None:
        return None
    try:
        return round(int(v) / 2, 2)
    except (ValueError, TypeError):
        return None


def _emu_pt(v: str | None) -> float | None:
    if v is None:
        return None
    try:
        return round(int(v) / 12700, 2)
    except (ValueError, TypeError):
        return None


# ── Font resolution with inheritance ──────────────────────────────────────────

def _fonts_from_rPr(rPr) -> dict[str, str]:
    """Extract all font variants from a w:rFonts element."""
    if rPr is None:
        return {}
    fonts_el = rPr.find(qn("w:rFonts"))
    if fonts_el is None:
        return {}
    out = {}
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        v = fonts_el.get(qn(f"w:{attr}"))
        if v:
            out[attr] = v
    return out


def _resolve_rPr(style_el, doc: DocxDocument) -> dict[str, Any]:
    """
    Resolve run properties following the Word inheritance chain:
      style rPr → basedOn style rPr → docDefaults rPr
    Logs each level so debugging is straightforward.
    """
    result: dict[str, Any] = {}
    chain: list[str] = []

    current_el = style_el
    seen_ids: set[str] = set()

    while current_el is not None:
        style_id = current_el.get(qn("w:styleId"), "<unknown>")
        if style_id in seen_ids:
            log.debug("%s Circular basedOn detected at styleId=%s", _WARN, style_id)
            break
        seen_ids.add(style_id)

        rPr = current_el.find(qn("w:rPr"))
        if rPr is not None:
            # Font
            if "font" not in result:
                fonts = _fonts_from_rPr(rPr)
                if fonts:
                    result["font"] = fonts.get("ascii") or fonts.get("hAnsi") or next(iter(fonts.values()))
                    result["font_variants"] = fonts
                    chain.append(f"font from {style_id}")

            # Size
            if "size_pt" not in result:
                sz = rPr.find(qn("w:sz"))
                if sz is not None:
                    v = _half_pt(sz.get(qn("w:val")))
                    if v:
                        result["size_pt"] = v
                        chain.append(f"size from {style_id}")

            # Bold
            if "bold" not in result:
                b = rPr.find(qn("w:b"))
                if b is not None:
                    val = b.get(qn("w:val"), "true").lower()
                    result["bold"] = val not in ("0", "false")
                    chain.append(f"bold from {style_id}")

            # Italic
            if "italic" not in result:
                i = rPr.find(qn("w:i"))
                if i is not None:
                    val = i.get(qn("w:val"), "true").lower()
                    result["italic"] = val not in ("0", "false")

            # Color
            if "color" not in result:
                color_el = rPr.find(qn("w:color"))
                if color_el is not None:
                    v = color_el.get(qn("w:val"), "")
                    if v and v.upper() not in ("AUTO", ""):
                        result["color"] = f"#{v.upper()}"
                        chain.append(f"color from {style_id}")

        # Walk up to basedOn
        basedOn = current_el.find(qn("w:basedOn"))
        if basedOn is None:
            break
        parent_id = basedOn.get(qn("w:val"))
        if not parent_id:
            break

        # Find parent style element
        parent_el = None
        for s in doc.styles.element.findall(qn("w:style")):
            if s.get(qn("w:styleId")) == parent_id:
                parent_el = s
                break
        if parent_el is None:
            log.debug("%s basedOn '%s' not found in styles", _WARN, parent_id)
            break
        current_el = parent_el

    if chain:
        log.debug("Style resolution chain: %s", " → ".join(chain))

    # Final fallback: docDefaults
    if "font" not in result or "size_pt" not in result:
        try:
            docDefaults = doc.styles.element.find(f".//{qn('w:docDefaults')}")
            if docDefaults is not None:
                rPrDef = docDefaults.find(f".//{qn('w:rPrDefault')}/{qn('w:rPr')}")
                if rPrDef is not None:
                    if "font" not in result:
                        fonts = _fonts_from_rPr(rPrDef)
                        if fonts:
                            result["font"] = fonts.get("ascii") or next(iter(fonts.values()))
                            log.debug("%s Font from docDefaults: %s", _FALLBACK, result["font"])
                    if "size_pt" not in result:
                        sz = rPrDef.find(qn("w:sz"))
                        if sz is not None:
                            v = _half_pt(sz.get(qn("w:val")))
                            if v:
                                result["size_pt"] = v
                                log.debug("%s Size from docDefaults: %s pt", _FALLBACK, v)
        except Exception as e:
            log.debug("docDefaults fallback error: %s", e)

    return result


# ── Paragraph properties ───────────────────────────────────────────────────────

def _extract_pPr(pPr) -> dict[str, Any]:
    """Extract full paragraph property block from a pPr element."""
    out: dict[str, Any] = {}
    if pPr is None:
        return out

    # Spacing
    sp = pPr.find(qn("w:spacing"))
    if sp is not None:
        spacing: dict[str, Any] = {}
        for attr in ("before", "after", "line"):
            v = _twips_pt(sp.get(qn(f"w:{attr}")))
            if v is not None:
                spacing[attr] = v
        lr = sp.get(qn("w:lineRule"))
        if lr:
            spacing["lineRule"] = lr
        if spacing:
            out["spacing"] = spacing

    # Indentation
    ind = pPr.find(qn("w:ind"))
    if ind is not None:
        indent: dict[str, float] = {}
        for attr in ("left", "right", "firstLine", "hanging"):
            v = _twips_pt(ind.get(qn(f"w:{attr}")))
            if v is not None and v != 0:
                indent[attr] = v
        if indent:
            out["indentation"] = indent

    # Alignment
    jc = pPr.find(qn("w:jc"))
    if jc is not None:
        out["alignment"] = jc.get(qn("w:val"), "left")

    # Tab stops
    tabs = pPr.find(qn("w:tabs"))
    if tabs is not None:
        tab_list = []
        for tab in tabs.findall(qn("w:tab")):
            pos = _twips_pt(tab.get(qn("w:pos")))
            val = tab.get(qn("w:val"), "")
            if pos is not None:
                tab_list.append({"pos_pt": pos, "type": val})
        if tab_list:
            out["tabs"] = tab_list

    return out


# ── Heading style extraction ───────────────────────────────────────────────────

def _extract_heading_styles(doc: DocxDocument) -> dict[str, dict[str, Any]]:
    level_names = {
        "Heading 1": "h1", "Heading 2": "h2", "Heading 3": "h3",
        "Heading 4": "h4", "Heading 5": "h5", "Heading 6": "h6",
        "Title": "title",
    }
    heading_map: dict[str, dict[str, Any]] = {}

    for style_name, key in level_names.items():
        try:
            style = doc.styles[style_name]
        except KeyError:
            log.debug("%s Heading style '%s' missing in document", _WARN, style_name)
            continue

        el = style.element
        info: dict[str, Any] = {"style_name": style_name}

        # Resolved run properties (font inheritance)
        info.update(_resolve_rPr(el, doc))

        # Paragraph properties
        pPr = el.find(qn("w:pPr"))
        info.update(_extract_pPr(pPr))

        # Sanity warnings
        if "font" not in info:
            log.warning("%s '%s' has no resolved font — will use document default", _WARN, style_name)
        if "size_pt" not in info:
            log.warning("%s '%s' has no resolved size — will use document default", _WARN, style_name)

        heading_map[key] = info
        log.debug(
            "Heading style '%s': font=%s size=%s bold=%s",
            style_name, info.get("font"), info.get("size_pt"), info.get("bold"),
        )

    return heading_map


# ── Body (Normal) style ────────────────────────────────────────────────────────

def _extract_body_style(doc: DocxDocument) -> dict[str, Any]:
    result: dict[str, Any] = {}
    try:
        style = doc.styles["Normal"]
        el = style.element
        result.update(_resolve_rPr(el, doc))
        pPr = el.find(qn("w:pPr"))
        result.update(_extract_pPr(pPr))
        log.debug(
            "Body style: font=%s size=%s spacing=%s",
            result.get("font"), result.get("size_pt"), result.get("spacing"),
        )
    except Exception as e:
        log.warning("%s Could not extract body style: %s", _WARN, e)
    return result


# ── Header/footer ──────────────────────────────────────────────────────────────

def _extract_header_footer(doc: DocxDocument) -> dict[str, Any]:
    """
    Extract header and footer text for each section.
    Detects page number fields (PAGE, NUMPAGES).
    """
    hf_info: dict[str, Any] = {
        "headers": [], "footers": [],
        "has_page_numbers": False,
    }

    PAGE_FIELDS = {"PAGE", "NUMPAGES", "SECTIONPAGES"}

    def _scan_part(part, kind: str, section_idx: int):
        if part is None:
            return
        texts = []
        has_pn = False
        try:
            for para in part.paragraphs:
                t = para.text.strip()
                if t:
                    texts.append(t)
                # Detect field codes
                for fld in para._p.findall(f".//{qn('w:fldChar')}"):
                    pass  # presence of fldChar indicates field
                for instrText in para._p.findall(f".//{qn('w:instrText')}"):
                    instr = (instrText.text or "").strip().upper()
                    if any(f in instr for f in PAGE_FIELDS):
                        has_pn = True
        except Exception as e:
            log.debug("Header/footer extraction error (section %d, %s): %s", section_idx, kind, e)
            return

        if texts or has_pn:
            entry = {"section": section_idx, "text": " | ".join(texts)}
            if has_pn:
                entry["has_page_numbers"] = True
                hf_info["has_page_numbers"] = True
            hf_info[f"{kind}s"].append(entry)

    for i, section in enumerate(doc.sections):
        try:
            _scan_part(section.header, "header", i)
            _scan_part(section.footer, "footer", i)
        except Exception as e:
            log.debug("Section %d header/footer error: %s", i, e)

    log.debug(
        "Header/footer: %d headers, %d footers, page_numbers=%s",
        len(hf_info["headers"]), len(hf_info["footers"]), hf_info["has_page_numbers"],
    )
    return hf_info


# ── Font inventory ─────────────────────────────────────────────────────────────

def _inventory_fonts(doc: DocxDocument) -> list[str]:
    """
    Scan all paragraphs and collect every unique font name in use.
    Helps detect custom/embedded fonts that need to be preserved.
    """
    fonts_seen: set[str] = set()
    for para in doc.paragraphs:
        for run in para.runs:
            rPr = run._r.find(qn("w:rPr"))
            if rPr is None:
                continue
            fonts_el = rPr.find(qn("w:rFonts"))
            if fonts_el is None:
                continue
            for attr in ("ascii", "hAnsi", "cs"):
                v = fonts_el.get(qn(f"w:{attr}"))
                if v:
                    fonts_seen.add(v)

    if fonts_seen:
        log.debug("Fonts in use: %s", sorted(fonts_seen))
    return sorted(fonts_seen)


# ── Page layout ────────────────────────────────────────────────────────────────

def _extract_page_layout(doc: DocxDocument) -> list[dict[str, Any]]:
    """Extract page size + margins for every section (multi-section docs)."""
    layouts = []
    for i, section in enumerate(doc.sections):
        layout: dict[str, Any] = {"section_index": i}
        try:
            layout["width_pt"] = round(section.page_width.pt, 2)
            layout["height_pt"] = round(section.page_height.pt, 2)
        except Exception:
            pass
        for side in ("top", "bottom", "left", "right", "header", "footer"):
            try:
                v = getattr(section, f"{side}_margin", None)
                if v is not None:
                    layout[f"margin_{side}_pt"] = round(v.pt, 2)
            except Exception:
                pass
        try:
            layout["orientation"] = section.orientation.name
        except Exception:
            pass
        layouts.append(layout)
        log.debug("Section %d layout: %s", i, layout)
    return layouts


# ── TOC styles ─────────────────────────────────────────────────────────────────

def _extract_toc_styles(doc: DocxDocument) -> list[dict[str, Any]]:
    """Extract TOC 1–9 styles if present."""
    toc_styles = []
    for level in range(1, 10):
        name = f"TOC {level}"
        try:
            style = doc.styles[name]
            el = style.element
            info: dict[str, Any] = {"level": level, "style_name": name}
            info.update(_resolve_rPr(el, doc))
            pPr = el.find(qn("w:pPr"))
            info.update(_extract_pPr(pPr))
            toc_styles.append(info)
        except KeyError:
            break  # TOC levels are contiguous
    if toc_styles:
        log.debug("TOC styles found: %d levels", len(toc_styles))
    return toc_styles


# ── Main ───────────────────────────────────────────────────────────────────────

def extract_style_profile(file_path: str | Path) -> dict[str, Any]:
    """
    Extract complete StyleProfile from a DOCX.
    Returns JSON-serialisable dict stored in Document.style_profile.
    """
    file_path = Path(file_path)
    log.info("Extracting style profile: %s", file_path.name)

    doc = DocxDocument(str(file_path))

    page_layouts = _extract_page_layout(doc)
    first_layout = page_layouts[0] if page_layouts else {}

    # Backwards-compatible page_margins key (first section only)
    page_margins = {
        k.replace("margin_", "").replace("_pt", ""): v
        for k, v in first_layout.items()
        if k.startswith("margin_")
    }

    profile: dict[str, Any] = {
        # Backwards-compat keys (used by rebuilder)
        "page_margins": page_margins,
        "page_size": {
            "width_pt": first_layout.get("width_pt"),
            "height_pt": first_layout.get("height_pt"),
        },
        # Full data
        "page_layouts": page_layouts,
        "document_defaults": _resolve_rPr(
            doc.styles.element.find(f".//{qn('w:docDefaults')}"), doc
        ),
        "body_style": _extract_body_style(doc),
        "heading_styles": _extract_heading_styles(doc),
        "toc_styles": _extract_toc_styles(doc),
        "header_footer": _extract_header_footer(doc),
        "fonts_in_use": _inventory_fonts(doc),
        "section_count": len(doc.sections),
    }

    log.info(
        "Style profile complete: %d heading styles, %d fonts, "
        "%d sections, page_numbers=%s",
        len(profile["heading_styles"]),
        len(profile["fonts_in_use"]),
        profile["section_count"],
        profile["header_footer"].get("has_page_numbers"),
    )
    return profile
