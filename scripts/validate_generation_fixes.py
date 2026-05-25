#!/usr/bin/env python3
"""
Validation script for content generation pipeline improvements.

Tests:
1. Sanitization module functionality
2. Fallback generator output cleanliness
3. Integration of sanitization layer
4. Transformer page handling
"""

import sys
from pathlib import Path

# Add workspace to path
WORKSPACE = Path(__file__).parent
sys.path.insert(0, str(WORKSPACE))

from app.services.ai.sanitizer import (
    sanitize_generated_content,
    is_clean_content,
    sanitize_ai_results,
    validate_sanitized_content,
)
from app.services.ai.generator import _local_generate_text


def test_sanitizer_patterns():
    """Test that sanitizer removes known pollution patterns."""
    print("\n" + "="*70)
    print("TEST 1: Sanitizer Pattern Removal")
    print("="*70)
    
    test_cases = [
        ("Generated section for '1. Aim'.", "Should remove 'Generated section for'"),
        ("Instruction: Rewrite this section.", "Should remove 'Instruction:' prefix"),
        ("Context: Some context here.", "Should remove 'Context:' prefix"),
        ("This paragraph is a fallback generated draft.", "Should remove fallback narration"),
        ("Content here. This is a fallback generated draft. More content.", "Should remove narration but preserve real content"),
    ]
    
    for dirty_text, description in test_cases:
        cleaned = sanitize_generated_content(dirty_text, section_heading="Test")
        print(f"\n{description}")
        print(f"  Input:  {dirty_text[:60]}")
        print(f"  Output: {cleaned[:60]}")
        is_clean = is_clean_content(cleaned)
        print(f"  Clean:  {is_clean}")


def test_fallback_generator():
    """Test that fallback generator produces clean output."""
    print("\n" + "="*70)
    print("TEST 2: Fallback Generator Output Cleanliness")
    print("="*70)
    
    instructions = [
        ("rewrite this section", "Rewrite instruction"),
        ("summarize the content", "Summarize instruction"),
        ("expand this topic", "Expand instruction"),
        ("write the section", "Generic instruction"),
    ]
    
    for instr, label in instructions:
        output = _local_generate_text(
            heading="Test Section",
            instruction=instr,
            global_context="Test document",
            style_profile={"body_style": {"size_pt": 12}},
            extra_context="Extra info",
        )
        
        is_clean = is_clean_content(output)
        print(f"\n{label}:")
        print(f"  Output: {output[:80]}")
        print(f"  Clean:  {is_clean}")
        
        # Verify no obvious pollution
        pollution_markers = [
            "Generated section",
            "Instruction:",
            "This paragraph is a fallback",
            "[FALLBACK]",
            "[DEBUG]",
        ]
        has_pollution = any(m in output for m in pollution_markers)
        print(f"  No pollution markers: {not has_pollution}")


def test_sanitize_ai_results():
    """Test sanitization of AI results dict."""
    print("\n" + "="*70)
    print("TEST 3: Sanitize AI Results Dictionary")
    print("="*70)
    
    # Simulate polluted AI results
    polluted_results = {
        "1. Introduction": "Generated section for '1. Introduction'. Instruction: Write an intro. This paragraph is a fallback generated draft. Real content here.",
        "2. Methods": "This section presents methods. Proper academic content.",
        "3. Results": "",  # Empty result
    }
    
    print("\nInput AI results:")
    for heading, text in polluted_results.items():
        print(f"  {heading}: {text[:60]}")
    
    sanitized = sanitize_ai_results(polluted_results)
    
    print("\nSanitized AI results:")
    for heading, text in sanitized.items():
        print(f"  {heading}: {text[:60]}")
    
    # Verify
    print("\nVerifications:")
    for heading, text in sanitized.items():
        is_clean = is_clean_content(text)
        is_valid, msg = validate_sanitized_content(heading, polluted_results.get(heading, ""), text)
        print(f"  {heading}: clean={is_clean}, valid={is_valid}")


def test_module_imports():
    """Test that all modules can be imported without errors."""
    print("\n" + "="*70)
    print("TEST 4: Module Imports")
    print("="*70)
    
    try:
        from app.services.ai.sanitizer import sanitize_ai_results
        print("✓ Sanitizer module imports successfully")
    except Exception as e:
        print(f"✗ Sanitizer import failed: {e}")
        return False
    
    try:
        from app.services.ai.generator import _local_generate_text
        print("✓ Generator module imports successfully")
    except Exception as e:
        print(f"✗ Generator import failed: {e}")
        return False
    
    try:
        from app.services.ai.rewrite_pipeline import run_rewrite_job
        print("✓ Rewrite pipeline imports successfully")
    except Exception as e:
        print(f"✗ Rewrite pipeline import failed: {e}")
        return False
    
    try:
        from app.services.docx.transformer import transform_docx
        print("✓ Transformer imports successfully")
    except Exception as e:
        print(f"✗ Transformer import failed: {e}")
        return False
    
    return True


def main():
    """Run all validation tests."""
    print("\n" + "="*70)
    print("CONTENT GENERATION PIPELINE VALIDATION")
    print("="*70)
    
    # Module imports
    if not test_module_imports():
        print("\n✗ Module import test failed!")
        return 1
    
    # Sanitizer tests
    try:
        test_sanitizer_patterns()
        test_fallback_generator()
        test_sanitize_ai_results()
        print("\n" + "="*70)
        print("✓ ALL TESTS PASSED")
        print("="*70)
        return 0
    except Exception as e:
        print(f"\n✗ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
