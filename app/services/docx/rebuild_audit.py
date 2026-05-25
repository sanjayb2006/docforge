"""
app/services/docx/rebuild_audit.py

Step 0 Instrumentation — Rebuild Fidelity Audit.

PURPOSE
-------
Capture a detailed before/after snapshot of every rebuild operation so we
have hard evidence of exactly what is being lost. This module is read-only
with respect to document content — it never modifies any DOCX file.

HOW IT WORKS
------------
1. snapshot_document()   — call BEFORE rebuild, on original_path
2. snapshot_document()   — call AFTER rebuild, on output_path
3. diff_snapshots()      — compare the two, produce a RebuildAudit
4. log_audit()           — emit structured log lines with [AUDIT] prefix
5. audit_to_dict()       — JSON-serialisable form for DB storage / dashboard

WHAT IT MEASURES
----------------
Paragraph layer:
  - total paragraph count delta
  - per-paragraph: style name, text fingerprint (first 40 chars + length),
    run count, spacing before/after, indent left, alignment

Run layer:
  - total run count delta
  - run-level font name, size, bold, italic presence per paragraph

Element layer:
  - table count delta
  - image count delta
  - page break count delta
  - section break count delta

Structure layer:
  - heading count and level distribution
  - heading text match (order-sensitive)
  - sections present in original but missing from output
  - sections present in output but absent from original (duplicates)

TEXT FINGERPRINT
---------------
We do NOT compare full text (AI content legitimately differs).
For preserved sections we compare:
  - first 40 chars of paragraph text (catches truncation / duplication)
  - character count (±20% tolerance flags as anomaly)

IMPORTANT: This module is purely observational. It has zero side effects.
It does not change, patch, or wrap the rebuilder. VS Code AI can safely
import and call it standalone or alongside the rebuilder.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docx import Document as DocxDocument
from docx.oxml.ns import qn

log = logging.getLogger(__name__)

_TAG = "[AUDIT]"       # prefix for all audit log lines
_WARN = "[AUDIT-WARN]" # prefix for anomaly lines
_LOSS = "[AUDIT-LOSS]" # prefix for confirmed fidelity loss


# ══════════════════════════════════════════════════════════════════════════════
# DATA STRUCTURES
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class ParagraphSnapshot:
    """Everything observable about one paragraph — no content modification."""
    index:       int
    style_name:  str
    text_head:   str          # first 60 chars
    text_len:    int          # total character count
    run_count:   int
    is_heading:  bool
    heading_level: int | None
    alignment:   str | None
    space_before_pt: float | None
    space_after_pt:  float | None
    line_spacing:    float | None
    indent_left_pt:  float | None
    has_page_break:  bool
    has_section_break: bool
    has_image:       bool
    has_field_code:  bool     # w:fldChar — page numbers, cross-refs, captions
    run_fonts:       list[str]   # unique font names across runs
    run_sizes:       list[float] # unique sizes across runs
    has_bold:        bool
    has_italic:      bool


@dataclass
class DocumentSnapshot:
    """Complete observable state of a DOCX before or after rebuild."""
    file_path:       str
    para_count:      int
    table_count:     int
    image_count:     int
    page_break_count: int
    section_break_count: int
    heading_count:   int
    heading_levels:  dict[int, int]   # {level: count}
    heading_texts:   list[str]         # ordered list
    total_run_count: int
    unique_fonts:    list[str]
    paragraphs:      list[ParagraphSnapshot]
    page_margins_pt: dict[str, float]


@dataclass
class ParagraphDiff:
    """Comparison result for a single paragraph index."""
    index:    int
    # What changed
    style_changed:    bool = False
    orig_style:       str = ""
    out_style:        str = ""
    text_truncated:   bool = False
    text_duplicated:  bool = False
    orig_text_len:    int = 0
    out_text_len:     int = 0
    run_count_changed: bool = False
    orig_runs:        int = 0
    out_runs:         int = 0
    spacing_changed:  bool = False
    indent_changed:   bool = False
    alignment_changed: bool = False
    image_lost:       bool = False
    field_code_lost:  bool = False
    font_changed:     bool = False
    orig_fonts:       list[str] = field(default_factory=list)
    out_fonts:        list[str] = field(default_factory=list)


@dataclass
class RebuildAudit:
    """
    Complete before/after comparison for one rebuild operation.
    Produced by diff_snapshots(). Immutable after creation.
    """
    original_path:  str
    output_path:    str

    # Counts
    orig_para_count:  int = 0
    out_para_count:   int = 0
    para_count_delta: int = 0     # positive = paragraphs added; negative = lost

    orig_table_count: int = 0
    out_table_count:  int = 0

    orig_image_count: int = 0
    out_image_count:  int = 0

    orig_heading_count: int = 0
    out_heading_count:  int = 0

    orig_run_count:   int = 0
    out_run_count:    int = 0

    # Heading structure
    headings_matched:  int = 0
    headings_missing:  list[str] = field(default_factory=list)   # in orig, not in output
    headings_added:    list[str] = field(default_factory=list)   # in output, not in orig

    # Margin drift
    margin_drifts:   dict[str, tuple[float, float]] = field(default_factory=dict)

    # Per-paragraph diffs (only paragraphs with changes)
    para_diffs:      list[ParagraphDiff] = field(default_factory=list)

    # Aggregate loss indicators
    tables_lost:     int = 0
    images_lost:     int = 0
    field_codes_lost: int = 0

    # Anomaly counts
    style_mismatches:    int = 0
    spacing_anomalies:   int = 0
    run_count_anomalies: int = 0
    text_truncations:    int = 0
    text_duplications:   int = 0
    font_changes:        int = 0

    # Estimated severity
    severity: str = "ok"   # ok | minor | major | critical


# ══════════════════════════════════════════════════════════════════════════════
# SNAPSHOT HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _count_images_in_para(para) -> bool:
    return bool(
        para._p.findall(f".//{qn('wp:inline')}", para._p.nsmap) or
        para._p.findall(f".//{qn('wp:anchor')}", para._p.nsmap)
    )


def _has_field_code(para) -> bool:
    return bool(para._p.findall(f".//{qn('w:fldChar')}"))


def _has_page_break(para) -> bool:
    for br in para._p.findall(f".//{qn('w:br')}"):
        if br.get(qn("w:type"), "") in ("page", "column"):
            return True
    return False


def _has_section_break(para) -> bool:
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return False
    return pPr.find(qn("w:sectPr")) is not None


def _spacing(para) -> tuple[float | None, float | None, float | None]:
    """Returns (before_pt, after_pt, line_pt)."""
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return None, None, None
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        return None, None, None

    def _v(attr: str) -> float | None:
        val = sp.get(qn(f"w:{attr}"))
        try:
            return round(int(val) / 20, 1) if val else None
        except (ValueError, TypeError):
            return None

    return _v("before"), _v("after"), _v("line")


def _indent_left(para) -> float | None:
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return None
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return None
    val = ind.get(qn("w:left"))
    try:
        return round(int(val) / 20, 1) if val else None
    except (ValueError, TypeError):
        return None


def _alignment(para) -> str | None:
    try:
        a = para.alignment
        return a.name if a else None
    except Exception:
        return None


def _run_fonts_and_sizes(para) -> tuple[list[str], list[float]]:
    fonts: set[str] = set()
    sizes: set[float] = set()
    for run in para.runs:
        rPr = run._r.find(qn("w:rPr"))
        if rPr is None:
            continue
        fonts_el = rPr.find(qn("w:rFonts"))
        if fonts_el is not None:
            for attr in ("ascii", "hAnsi"):
                v = fonts_el.get(qn(f"w:{attr}"))
                if v:
                    fonts.add(v)
                    break
        sz = rPr.find(qn("w:sz"))
        if sz is not None:
            try:
                sizes.add(round(int(sz.get(qn("w:val"), 0)) / 2, 1))
            except (ValueError, TypeError):
                pass
    return sorted(fonts), sorted(sizes)


def _is_heading(para) -> tuple[bool, int | None]:
    style_name = (para.style.name or "").lower().strip()
    heading_map = {
        "heading 1": 1, "heading 2": 2, "heading 3": 3,
        "heading 4": 4, "heading 5": 5, "heading 6": 6,
        "title": 0,
    }
    if style_name in heading_map:
        return True, heading_map[style_name]
    # Outline level fallback
    pPr = para._p.find(qn("w:pPr"))
    if pPr is not None:
        ol = pPr.find(qn("w:outlineLvl"))
        if ol is not None:
            v = ol.get(qn("w:val"))
            if v is not None:
                try:
                    return True, int(v) + 1
                except ValueError:
                    pass
    return False, None


def _snapshot_paragraph(para, idx: int) -> ParagraphSnapshot:
    """Build a complete ParagraphSnapshot from a live paragraph object."""
    text = para.text or ""
    is_h, hlevel = _is_heading(para)
    before_pt, after_pt, line_pt = _spacing(para)
    fonts, sizes = _run_fonts_and_sizes(para)

    has_b = any(run.bold for run in para.runs if run.bold is not None)
    has_i = any(run.italic for run in para.runs if run.italic is not None)

    return ParagraphSnapshot(
        index             = idx,
        style_name        = para.style.name or "Normal",
        text_head         = text[:60],
        text_len          = len(text),
        run_count         = len(para.runs),
        is_heading        = is_h,
        heading_level     = hlevel,
        alignment         = _alignment(para),
        space_before_pt   = before_pt,
        space_after_pt    = after_pt,
        line_spacing      = line_pt,
        indent_left_pt    = _indent_left(para),
        has_page_break    = _has_page_break(para),
        has_section_break = _has_section_break(para),
        has_image         = _count_images_in_para(para),
        has_field_code    = _has_field_code(para),
        run_fonts         = fonts,
        run_sizes         = sizes,
        has_bold          = has_b,
        has_italic        = has_i,
    )


def _page_margins(doc: DocxDocument) -> dict[str, float]:
    margins: dict[str, float] = {}
    try:
        section = doc.sections[0]
        for side in ("top", "bottom", "left", "right"):
            v = getattr(section, f"{side}_margin", None)
            if v is not None:
                margins[side] = round(v.pt, 1)
    except Exception:
        pass
    return margins


# ══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ══════════════════════════════════════════════════════════════════════════════

def snapshot_document(file_path: str | Path, label: str = "") -> DocumentSnapshot:
    """
    Capture a complete observable snapshot of a DOCX file.
    No modifications — purely reads the file.

    Args:
        file_path: Path to DOCX file
        label:     Optional label for log output ("original" | "output")

    Returns:
        DocumentSnapshot — call before and after rebuild, then diff_snapshots()
    """
    file_path = Path(file_path)
    label_str = f" [{label}]" if label else ""
    log.info("%s Snapshotting%s: %s", _TAG, label_str, file_path.name)

    doc = DocxDocument(str(file_path))
    paragraphs: list[ParagraphSnapshot] = []

    heading_levels: dict[int, int] = {}
    heading_texts: list[str] = []
    total_runs = 0
    fonts_seen: set[str] = set()
    page_breaks = 0
    section_breaks = 0
    images = 0

    for idx, para in enumerate(doc.paragraphs):
        snap = _snapshot_paragraph(para, idx)
        paragraphs.append(snap)

        total_runs += snap.run_count
        fonts_seen.update(snap.run_fonts)

        if snap.has_page_break:
            page_breaks += 1
        if snap.has_section_break:
            section_breaks += 1
        if snap.has_image:
            images += 1

        if snap.is_heading and snap.heading_level is not None:
            heading_levels[snap.heading_level] = heading_levels.get(snap.heading_level, 0) + 1
            heading_texts.append(snap.text_head.strip())

    snap_out = DocumentSnapshot(
        file_path         = str(file_path),
        para_count        = len(doc.paragraphs),
        table_count       = len(doc.tables),
        image_count       = images,
        page_break_count  = page_breaks,
        section_break_count = section_breaks,
        heading_count     = sum(heading_levels.values()),
        heading_levels    = heading_levels,
        heading_texts     = heading_texts,
        total_run_count   = total_runs,
        unique_fonts      = sorted(fonts_seen),
        paragraphs        = paragraphs,
        page_margins_pt   = _page_margins(doc),
    )

    log.info(
        "%s%s → %d paras, %d tables, %d images, %d headings, %d runs, fonts=%s",
        _TAG, label_str,
        snap_out.para_count,
        snap_out.table_count,
        snap_out.image_count,
        snap_out.heading_count,
        snap_out.total_run_count,
        snap_out.unique_fonts,
    )
    return snap_out


def diff_snapshots(
    original: DocumentSnapshot,
    output:   DocumentSnapshot,
    ai_sections: set[str] | None = None,
) -> RebuildAudit:
    """
    Compare two document snapshots and produce a RebuildAudit.

    Args:
        original:    Snapshot of the original file (before rebuild)
        output:      Snapshot of the rebuilt file (after rebuild)
        ai_sections: Set of heading texts that were AI-replaced.
                     Paragraphs under these headings are expected to differ —
                     we skip text-content comparison for them but still check
                     structure (style, spacing, count).

    Returns:
        RebuildAudit with all findings populated.
    """
    ai_sections = ai_sections or set()
    audit = RebuildAudit(
        original_path = original.file_path,
        output_path   = output.file_path,
    )

    # ── Counts ────────────────────────────────────────────────────────────────
    audit.orig_para_count    = original.para_count
    audit.out_para_count     = output.para_count
    audit.para_count_delta   = output.para_count - original.para_count
    audit.orig_table_count   = original.table_count
    audit.out_table_count    = output.table_count
    audit.orig_image_count   = original.image_count
    audit.out_image_count    = output.image_count
    audit.orig_heading_count = original.heading_count
    audit.out_heading_count  = output.heading_count
    audit.orig_run_count     = original.total_run_count
    audit.out_run_count      = output.total_run_count
    audit.tables_lost        = max(0, original.table_count - output.table_count)
    audit.images_lost        = max(0, original.image_count - output.image_count)

    # ── Heading structure ─────────────────────────────────────────────────────
    orig_headings_set = set(original.heading_texts)
    out_headings_set  = set(output.heading_texts)
    audit.headings_matched = len(orig_headings_set & out_headings_set)
    audit.headings_missing = sorted(orig_headings_set - out_headings_set)
    audit.headings_added   = sorted(out_headings_set - orig_headings_set)

    # ── Margin drift ──────────────────────────────────────────────────────────
    for side, orig_val in original.page_margins_pt.items():
        out_val = output.page_margins_pt.get(side)
        if out_val is not None and abs(orig_val - out_val) > 1.0:  # 1pt tolerance
            audit.margin_drifts[side] = (orig_val, out_val)

    # ── Per-paragraph comparison ──────────────────────────────────────────────
    # We compare by index up to min(orig, out) length
    compare_len = min(len(original.paragraphs), len(output.paragraphs))

    # Track which heading we're currently under (for AI-section skipping)
    current_heading: str | None = None

    for i in range(compare_len):
        op = original.paragraphs[i]
        rp = output.paragraphs[i]

        # Track heading context
        if op.is_heading:
            current_heading = op.text_head.strip()

        in_ai_section = current_heading in ai_sections

        diff = ParagraphDiff(index=i)
        changed = False

        # Style name
        if op.style_name != rp.style_name:
            diff.style_changed = True
            diff.orig_style    = op.style_name
            diff.out_style     = rp.style_name
            audit.style_mismatches += 1
            changed = True

        # Text length — only check preserved (non-AI) sections
        if not in_ai_section and op.text_len > 0:
            diff.orig_text_len = op.text_len
            diff.out_text_len  = rp.text_len
            ratio = rp.text_len / op.text_len if op.text_len else 1.0

            if ratio < 0.5:  # lost >50% of chars
                diff.text_truncated = True
                audit.text_truncations += 1
                changed = True
            elif ratio > 1.8:  # 80% more chars than original
                diff.text_duplicated = True
                audit.text_duplications += 1
                changed = True

        # Run count — significant if not AI section
        if not in_ai_section and op.run_count > 0:
            diff.orig_runs = op.run_count
            diff.out_runs  = rp.run_count
            if abs(op.run_count - rp.run_count) > max(1, op.run_count // 3):
                diff.run_count_changed = True
                audit.run_count_anomalies += 1
                changed = True

        # Spacing — check all paragraphs including AI sections (formatting should hold)
        sp_changed = False
        if op.space_before_pt is not None and rp.space_before_pt is not None:
            if abs(op.space_before_pt - rp.space_before_pt) > 2.0:
                sp_changed = True
        if op.space_after_pt is not None and rp.space_after_pt is not None:
            if abs(op.space_after_pt - rp.space_after_pt) > 2.0:
                sp_changed = True
        if sp_changed:
            diff.spacing_changed = True
            audit.spacing_anomalies += 1
            changed = True

        # Indentation
        if op.indent_left_pt is not None and rp.indent_left_pt is not None:
            if abs(op.indent_left_pt - rp.indent_left_pt) > 2.0:
                diff.indent_changed = True
                changed = True

        # Alignment
        if op.alignment and rp.alignment and op.alignment != rp.alignment:
            diff.alignment_changed = True
            changed = True

        # Image presence
        if op.has_image and not rp.has_image:
            diff.image_lost = True
            changed = True

        # Field code presence
        if op.has_field_code and not rp.has_field_code:
            diff.field_code_lost = True
            audit.field_codes_lost += 1
            changed = True

        # Font changes (preserved sections only)
        if not in_ai_section and op.run_fonts and rp.run_fonts:
            if set(op.run_fonts) != set(rp.run_fonts):
                diff.font_changed = True
                diff.orig_fonts   = op.run_fonts
                diff.out_fonts    = rp.run_fonts
                audit.font_changes += 1
                changed = True

        if changed:
            audit.para_diffs.append(diff)

    # ── Severity classification ───────────────────────────────────────────────
    audit.severity = _classify_severity(audit)

    return audit


def _classify_severity(audit: RebuildAudit) -> str:
    """
    Classify overall rebuild severity based on audit findings.

    ok       — paragraph count within ±2, no critical losses
    minor    — small count delta, some spacing/style anomalies
    major    — significant para count delta, tables/images lost, many anomalies
    critical — headings missing, severe count delta, complete element loss
    """
    if audit.tables_lost > 0 or audit.images_lost > 1:
        return "critical"
    if audit.headings_missing:
        return "critical"
    if abs(audit.para_count_delta) > 10:
        return "critical"
    if audit.text_duplications > 0:
        return "critical"

    if abs(audit.para_count_delta) > 4:
        return "major"
    if audit.style_mismatches > 3:
        return "major"
    if audit.spacing_anomalies > 5:
        return "major"
    if audit.text_truncations > 0:
        return "major"

    if abs(audit.para_count_delta) > 1:
        return "minor"
    if audit.style_mismatches > 0:
        return "minor"
    if audit.spacing_anomalies > 0:
        return "minor"
    if audit.font_changes > 0:
        return "minor"

    return "ok"


# ══════════════════════════════════════════════════════════════════════════════
# LOGGING OUTPUT
# ══════════════════════════════════════════════════════════════════════════════

def log_audit(audit: RebuildAudit) -> None:
    """
    Emit structured log lines from a RebuildAudit.
    All lines prefixed with [AUDIT], [AUDIT-WARN], or [AUDIT-LOSS].
    Designed to be grep-able from server logs.

    Call this after diff_snapshots() to get the full picture in logs.
    """
    sev = audit.severity.upper()
    log.info(
        "%s ══ REBUILD AUDIT [%s] ══════════════════════════",
        _TAG, sev,
    )

    # ── Paragraph counts ──────────────────────────────────────────────────────
    delta_sign = "+" if audit.para_count_delta >= 0 else ""
    log.info(
        "%s paragraphs: orig=%d  out=%d  delta=%s%d",
        _TAG,
        audit.orig_para_count,
        audit.out_para_count,
        delta_sign,
        audit.para_count_delta,
    )
    if abs(audit.para_count_delta) > 2:
        log.warning(
            "%s paragraph count delta=%d%d — significant drift",
            _WARN, delta_sign, audit.para_count_delta,
        )

    # ── Element counts ────────────────────────────────────────────────────────
    log.info(
        "%s tables:   orig=%d  out=%d",
        _TAG, audit.orig_table_count, audit.out_table_count,
    )
    log.info(
        "%s images:   orig=%d  out=%d",
        _TAG, audit.orig_image_count, audit.out_image_count,
    )
    log.info(
        "%s headings: orig=%d  out=%d  matched=%d",
        _TAG,
        audit.orig_heading_count,
        audit.out_heading_count,
        audit.headings_matched,
    )
    log.info(
        "%s runs:     orig=%d  out=%d",
        _TAG, audit.orig_run_count, audit.out_run_count,
    )

    # ── Losses ────────────────────────────────────────────────────────────────
    if audit.tables_lost:
        log.error(
            "%s TABLES LOST: %d table(s) present in original, missing from output",
            _LOSS, audit.tables_lost,
        )
    if audit.images_lost:
        log.error(
            "%s IMAGES LOST: %d image(s) present in original, missing from output",
            _LOSS, audit.images_lost,
        )
    if audit.field_codes_lost:
        log.warning(
            "%s FIELD CODES LOST: %d (page numbers, TOC refs, cross-references)",
            _WARN, audit.field_codes_lost,
        )

    # ── Heading structure ─────────────────────────────────────────────────────
    if audit.headings_missing:
        for h in audit.headings_missing:
            log.error(
                "%s HEADING MISSING from output: '%s'",
                _LOSS, h[:60],
            )
    if audit.headings_added:
        for h in audit.headings_added:
            log.warning(
                "%s HEADING ADDED (not in original): '%s'",
                _WARN, h[:60],
            )

    # ── Margin drift ──────────────────────────────────────────────────────────
    for side, (orig_v, out_v) in audit.margin_drifts.items():
        log.warning(
            "%s MARGIN DRIFT [%s]: orig=%.1fpt  out=%.1fpt  delta=%.1fpt",
            _WARN, side, orig_v, out_v, out_v - orig_v,
        )

    # ── Aggregate anomaly counts ───────────────────────────────────────────────
    log.info(
        "%s anomalies: style_mismatches=%d  spacing=%d  run_count=%d  "
        "text_truncations=%d  text_duplications=%d  font_changes=%d",
        _TAG,
        audit.style_mismatches,
        audit.spacing_anomalies,
        audit.run_count_anomalies,
        audit.text_truncations,
        audit.text_duplications,
        audit.font_changes,
    )

    # ── Per-paragraph diffs (most significant first) ───────────────────────────
    if audit.para_diffs:
        log.info("%s ── Per-paragraph changes (%d paragraphs affected) ──", _TAG, len(audit.para_diffs))
        # Show the worst ones first (prioritise losses and duplications)
        sorted_diffs = sorted(
            audit.para_diffs,
            key=lambda d: (
                d.image_lost * 10 +
                d.text_duplicated * 8 +
                d.text_truncated * 6 +
                d.field_code_lost * 4 +
                d.style_changed * 2 +
                d.spacing_changed
            ),
            reverse=True,
        )
        for diff in sorted_diffs[:20]:   # cap at 20 lines to avoid log spam
            changes: list[str] = []
            if diff.image_lost:      changes.append("IMAGE-LOST")
            if diff.text_duplicated: changes.append(f"DUPLICATED(orig={diff.orig_text_len} out={diff.out_text_len})")
            if diff.text_truncated:  changes.append(f"TRUNCATED(orig={diff.orig_text_len} out={diff.out_text_len})")
            if diff.field_code_lost: changes.append("FIELD-CODE-LOST")
            if diff.style_changed:   changes.append(f"STYLE({diff.orig_style!r}→{diff.out_style!r})")
            if diff.spacing_changed: changes.append("SPACING")
            if diff.indent_changed:  changes.append("INDENT")
            if diff.alignment_changed: changes.append("ALIGNMENT")
            if diff.run_count_changed: changes.append(f"RUNS({diff.orig_runs}→{diff.out_runs})")
            if diff.font_changed:    changes.append(f"FONT({diff.orig_fonts}→{diff.out_fonts})")

            level = log.error if (diff.image_lost or diff.text_duplicated) else log.warning
            level("%s para[%d]: %s", _WARN, diff.index, "  ".join(changes))

    log.info("%s severity=%s ══════════════════════════════════════════", _TAG, sev)


# ══════════════════════════════════════════════════════════════════════════════
# SERIALISATION
# ══════════════════════════════════════════════════════════════════════════════

def audit_to_dict(audit: RebuildAudit) -> dict[str, Any]:
    """
    Convert a RebuildAudit to a JSON-serialisable dict.
    Suitable for storing in GenerationJob.ai_results or returning via API.
    """
    return {
        "severity":            audit.severity,
        "original_path":       audit.original_path,
        "output_path":         audit.output_path,
        "para_count":          {"orig": audit.orig_para_count,   "out": audit.out_para_count,   "delta": audit.para_count_delta},
        "table_count":         {"orig": audit.orig_table_count,  "out": audit.out_table_count,  "lost": audit.tables_lost},
        "image_count":         {"orig": audit.orig_image_count,  "out": audit.out_image_count,  "lost": audit.images_lost},
        "heading_count":       {"orig": audit.orig_heading_count, "out": audit.out_heading_count, "matched": audit.headings_matched},
        "run_count":           {"orig": audit.orig_run_count,    "out": audit.out_run_count},
        "headings_missing":    audit.headings_missing,
        "headings_added":      audit.headings_added,
        "margin_drifts":       {k: {"orig": v[0], "out": v[1]} for k, v in audit.margin_drifts.items()},
        "field_codes_lost":    audit.field_codes_lost,
        "anomalies": {
            "style_mismatches":    audit.style_mismatches,
            "spacing_anomalies":   audit.spacing_anomalies,
            "run_count_anomalies": audit.run_count_anomalies,
            "text_truncations":    audit.text_truncations,
            "text_duplications":   audit.text_duplications,
            "font_changes":        audit.font_changes,
        },
        "para_diffs_count":    len(audit.para_diffs),
        "para_diffs_sample":   [
            {
                "index": d.index,
                "style_changed":     d.style_changed,
                "text_truncated":    d.text_truncated,
                "text_duplicated":   d.text_duplicated,
                "spacing_changed":   d.spacing_changed,
                "image_lost":        d.image_lost,
                "field_code_lost":   d.field_code_lost,
            }
            for d in audit.para_diffs[:10]
        ],
    }
