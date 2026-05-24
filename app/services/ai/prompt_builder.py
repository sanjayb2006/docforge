"""
app/services/ai/prompt_builder.py

Builds OpenAI API message lists for the content generation pipeline.

Each section gets a prompt that includes:
  1. System role   — academic writer persona + hard formatting rules
  2. Global context — document subject, student info, purpose
  3. Style hints   — derived from the StyleProfile (paragraph density cue)
  4. Section heading + user instruction
  5. Optional extra context (lab data, readings, measurements)
  6. Output constraints — plain text, no markdown, no heading repetition

Quality levers:
  - temperature=0.4  → consistent, non-hallucinating academic writing
  - No markdown in output — Word handles all formatting
  - Third-person passive voice instructions for technical sections
"""

from __future__ import annotations
from typing import Any

# ── System prompt ──────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are an expert technical writer specialising in Indian engineering university \
documentation, particularly VTU (Visvesvaraya Technological University) lab reports, \
project reports, and academic submissions.

Your writing rules (non-negotiable):
- Write in third person, passive voice for technical sections
- Use precise, academic language — no colloquial phrases
- Output PLAIN TEXT only — no markdown, no asterisks, no bullet symbols, no headers
- Separate logical paragraphs with a blank line (double newline)
- Do NOT repeat the section heading — start directly with the content
- Do NOT add sub-headings unless the instruction explicitly requests them
- Do NOT write "In conclusion" or "In summary" unless asked
- Be factually accurate; if you don't know a specific value, state it generally
- Aim for completeness: cover the topic as a well-informed engineering student would
"""


# ── Per-section prompt ─────────────────────────────────────────────────────────

def build_section_prompt(
    section_heading: str,
    instruction:     str,
    global_context:  str,
    style_profile:   dict[str, Any],
    extra_context:   str | None = None,
) -> list[dict[str, str]]:
    """
    Build the messages list for one section generation call.

    Returns:
        OpenAI-format messages: [{"role": ..., "content": ...}, ...]
    """
    parts: list[str] = []

    if global_context.strip():
        parts.append(f"DOCUMENT CONTEXT:\n{global_context.strip()}")

    parts.append(f"SECTION HEADING: {section_heading}")
    parts.append(f"INSTRUCTION: {instruction.strip()}")

    if extra_context and extra_context.strip():
        parts.append(f"ADDITIONAL DATA / CONTEXT:\n{extra_context.strip()}")

    style_hint = _style_hint(style_profile)
    if style_hint:
        parts.append(f"WRITING STYLE NOTE: {style_hint}")

    parts.append(
        "Now write the body content for this section. "
        "Plain text only. No markdown. Do not repeat the heading."
    )

    return [
        {"role": "system",  "content": SYSTEM_PROMPT},
        {"role": "user",    "content": "\n\n".join(parts)},
    ]


def _style_hint(style_profile: dict[str, Any]) -> str:
    """Derive a writing density hint from the body style font size."""
    body  = style_profile.get("body_style", {})
    size  = body.get("size_pt")
    if size and size >= 12:
        return "Use full, well-developed paragraphs (4–6 sentences each)."
    if size and size <= 10:
        return "Keep paragraphs concise (2–4 sentences each)."
    return ""


# ── Bulk regeneration prompt ───────────────────────────────────────────────────

def build_bulk_prompt(
    structure:      dict[str, Any],
    global_context: str,
    style_profile:  dict[str, Any],
) -> list[dict[str, str]]:
    """
    For replace_all mode on small documents (≤ 6 sections).
    Single API call returns all sections as a JSON object.
    """
    sections = structure.get("sections", [])
    headings = [s["heading_text"] for s in sections]

    system = (
        SYSTEM_PROMPT
        + "\n\nRespond ONLY with a valid JSON object. "
        "Keys are the exact section headings provided. "
        "Values are the plain text body content for each section. "
        "No markdown in the JSON values. No extra keys."
    )

    user_parts: list[str] = []
    if global_context.strip():
        user_parts.append(f"DOCUMENT CONTEXT:\n{global_context.strip()}")

    user_parts.append("Generate body content for ALL of these sections:\n" +
                       "\n".join(f"  - {h}" for h in headings))

    user_parts.append(
        'Respond ONLY with JSON, e.g.:\n'
        '{"Section A": "body text...", "Section B": "body text..."}'
    )

    return [
        {"role": "system", "content": system},
        {"role": "user",   "content": "\n\n".join(user_parts)},
    ]
