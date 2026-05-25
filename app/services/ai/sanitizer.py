"""
app/services/ai/sanitizer.py

Content sanitization layer — removes polluted/debug content before transformer injection.

This layer ensures that ONLY clean semantic content reaches the transformer.
It removes:
  - "Generated section for..."
  - "Instruction:" fragments
  - "Context:" metadata
  - "fallback generated draft" narration
  - Prompt echoes and debug text
  - Metadata wrapped in identifiable patterns

Philosophy:
  - Conservative: only remove patterns we are SURE are pollution
  - Preserve legitimate academic content
  - Log all removals for debugging
  - Never fail the whole job due to sanitization
"""

from __future__ import annotations

import logging
import re
from typing import Any

log = logging.getLogger(__name__)

_TAG = "[SANITIZE]"


# Pollution patterns that match ENTIRE LINE (line-level filtering)
# These are conservative: we only remove lines that are PURELY pollution.
FULL_LINE_PATTERNS = [
    # Fallback generator narration (full line)
    r"^Generated section for .+\.$",
    r"^This paragraph is a fallback generated draft\.$",

    # Instruction leakage (full line)
    r"^Instruction:\s+.+$",

    # Context leakage (full line)
    r"^Context:\s+.+$",
    r"^Extra context:\s+.+$",
    r"^DOCUMENT CONTEXT:\s*$",
    r"^SECTION HEADING:\s+.+$",
    r"^WRITING STYLE NOTE:\s+.+$",
    r"^ADDITIONAL DATA / CONTEXT:\s*$",

    # Common instruction patterns (full line)
    r"^Now write the body content.+$",
    r"^Plain text only\..*$",
    r"^Do not repeat the heading\.$",

    # Prompt metadata (full line)
    r"^Style profile applied\.$",
]

# Pollution fragments that may be embedded in a larger response.
FRAGMENT_PATTERNS = [
    r"Generated section for [^\.\n]+\.?",
    r"Instruction:\s*[^\.\n]+(?:\.|$)",
    r"Context:\s*[^\.\n]+(?:\.|$)",
    r"Extra context:\s*[^\.\n]+(?:\.|$)",
    r"DOCUMENT CONTEXT:\s*.*(?:\n|$)",
    r"SECTION HEADING:\s*[^\.\n]+(?:\.|$)",
    r"WRITING STYLE NOTE:\s*[^\.\n]+(?:\.|$)",
    r"ADDITIONAL DATA / CONTEXT:\s*.*(?:\n|$)",
    r"\[FALLBACK\]|\[DEBUG\]|\[GENERATED\]",
    r"This paragraph is a fallback generated draft\.?",
    r"Now write the body content.*",
    r"Plain text only\..*",
    r"Do not repeat the heading\..*",
    r"Style profile applied\.",
]

# Compile for efficiency
COMPILED_LINE_PATTERNS = [re.compile(pattern, re.IGNORECASE | re.MULTILINE) for pattern in FULL_LINE_PATTERNS]
COMPILED_FRAGMENT_PATTERNS = [re.compile(pattern, re.IGNORECASE) for pattern in FRAGMENT_PATTERNS]


# ══════════════════════════════════════════════════════════════════════════════
# SANITIZATION FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def sanitize_generated_content(text: str, section_heading: str = "") -> str:
    """
    Clean AI-generated content by removing known pollution patterns.

    Args:
        text:              Raw output from AI or fallback generator
        section_heading:   For logging/context only

    Returns:
        Cleaned text with pollution patterns removed.
    """
    if not text or not isinstance(text, str):
        return ""

    original_len = len(text)
    cleaned = text

    # Remove entire lines that match pollution patterns
    lines = cleaned.split("\n")
    filtered_lines = []
    removed_count = 0

    for line in lines:
        is_polluted = False
        for pattern in COMPILED_LINE_PATTERNS:
            if pattern.fullmatch(line):
                is_polluted = True
                removed_count += 1
                log.debug(
                    "%s Removed pollution line (heading='%s'): %s",
                    _TAG, section_heading[:50], line[:80],
                )
                break

        if not is_polluted:
            filtered_lines.append(line)

    cleaned = "\n".join(filtered_lines)

    # Remove inline pollution fragments without discarding surrounding text.
    for pattern in COMPILED_FRAGMENT_PATTERNS:
        cleaned = pattern.sub("", cleaned)

    # Normalise whitespace and blank lines after removals.
    cleaned = re.sub(r"[ \t]{2,}", " ", cleaned)
    cleaned = re.sub(r"\n\s*\n+", "\n\n", cleaned)
    cleaned = cleaned.strip()

    final_len = len(cleaned)
    if removed_count > 0 or final_len != original_len:
        log.info(
            "%s Sanitized '%s': %d lines removed, %d → %d chars",
            _TAG, section_heading[:50], removed_count, original_len, final_len,
        )

    return cleaned


def is_clean_content(text: str) -> bool:
    """
    Check if content appears to be clean (no obvious pollution).

    Returns True if content is safe to inject into the transformer.
    """
    if not text or not isinstance(text, str):
        return False

    # Check for obvious red flags
    for pattern in COMPILED_LINE_PATTERNS + COMPILED_FRAGMENT_PATTERNS:
        if pattern.search(text):
            return False

    # Must have at least some substance
    if len(text.strip()) < 10:
        return False

    return True


def sanitize_ai_results(
    ai_results: dict[str, str],
    structure: dict[str, Any] | None = None,
) -> dict[str, str]:
    """
    Sanitize all generated content in the results dict.

    Args:
        ai_results: {heading_text: generated_body_text}
        structure:  Optional structure dict for better logging

    Returns:
        {heading_text: sanitized_text} with dirty content removed.
    """
    if not ai_results:
        return {}

    sanitized: dict[str, str] = {}
    removed_count = 0

    for heading, text in ai_results.items():
        cleaned = sanitize_generated_content(text, section_heading=heading)

        if not cleaned:
            log.warning(
                "%s Section '%s': all content was pollution — section will be empty",
                _TAG, heading[:50],
            )
            removed_count += 1
            continue

        if cleaned != text:
            log.info(
                "%s Section '%s': content cleaned (pollution removed)",
                _TAG, heading[:50],
            )

        sanitized[heading] = cleaned

    log.info(
        "%s Sanitized %d sections, %d removed completely",
        _TAG, len(ai_results), removed_count,
    )

    return sanitized


# ══════════════════════════════════════════════════════════════════════════════
# VALIDATION
# ══════════════════════════════════════════════════════════════════════════════

def validate_sanitized_content(
    heading: str,
    original_text: str,
    sanitized_text: str,
) -> tuple[bool, str]:
    """
    Validate that sanitized content is usable and hasn't been corrupted.

    Returns:
        (is_valid, message)
    """
    if not sanitized_text or not sanitized_text.strip():
        return False, f"Sanitized text is empty for '{heading[:50]}'"

    # Ensure we didn't accidentally remove all content
    original_words = len(original_text.split())
    sanitized_words = len(sanitized_text.split())

    if sanitized_words == 0:
        return False, f"No words remain after sanitization for '{heading[:50]}'"

    # Warn if we removed >50% of content (likely too aggressive)
    if original_words > 10 and sanitized_words < original_words * 0.5:
        return True, f"Warning: removed >{50}% of content for '{heading[:50]}' (only pollution?)"

    return True, f"Sanitized content for '{heading[:50]}' is valid"
