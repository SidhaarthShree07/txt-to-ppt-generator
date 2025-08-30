"""
Helpers to detect content formats and placeholder capacities
"""
import re
from typing import List, Dict, Any

BULLET_PREFIXES = ("• ", "- ", "* ", "•", "-")


def detect_content_format(items: List[str]) -> str:
    """
    Detect likely content format from a list of strings
    Returns: 'numbered_list' | 'bullet_list' | 'paragraph' | 'multi_paragraph'
    """
    if not items:
        return 'bullet_list'

    # Normalize
    norm = [str(x).strip() for x in items if str(x).strip()]
    if not norm:
        return 'bullet_list'

    # Check separators are not considered
    norm_wo_markers = [x for x in norm if not _is_separator(x)]

    numbered_re = re.compile(r"^\d+[\.|\)]\s")
    numbered = sum(1 for x in norm_wo_markers if numbered_re.match(x))
    bullets = sum(1 for x in norm_wo_markers if x.startswith(BULLET_PREFIXES))

    # Paragraph detection: if only one long item
    if len(norm_wo_markers) == 1 and len(norm_wo_markers[0]) > 140:
        return 'paragraph'

    if numbered >= max(1, len(norm_wo_markers) // 2):
        return 'numbered_list'
    if bullets >= max(1, len(norm_wo_markers) // 2):
        return 'bullet_list'

    # If multiple medium-length lines, call it bullet list
    if len(norm_wo_markers) > 1:
        return 'bullet_list'

    return 'paragraph'


def has_separators(items: List[str]) -> bool:
    return any(_is_separator(x) for x in items or [])


def _is_separator(text: str) -> bool:
    t = str(text).upper()
    return any(sep in t for sep in ["[NEXT_PLACEHOLDER]", "[PLACEHOLDER]", "[TEXT_AREA]", "---", "###"]) or t == "|" or t == "||"


def count_groups_by_separators(items: List[str]) -> int:
    if not items:
        return 1
    groups = 1
    for x in items:
        if _is_separator(x):
            groups += 1
    return groups


def get_content_placeholders_from_template_slide(template_slide: Dict[str, Any]) -> List[Dict[str, Any]]:
    placeholders = template_slide.get('placeholders', [])
    result = []
    for ph in placeholders:
        ph_type = str(ph.get('type', '')).upper()
        if 'CONTENT' in ph_type or 'BODY' in ph_type or 'TEXT' in ph_type or 'OBJECT' in ph_type:
            result.append(ph)
    return result


def placeholder_capacity(ph: Dict[str, Any]) -> Dict[str, Any]:
    return {
        'suggested_lines': ph.get('suggested_lines', ph.get('lines_capacity', 5) or 5),
        'max_chars_per_line': ph.get('max_chars_per_line', ph.get('chars_per_line', 80) or 80),
        'text_format': ph.get('text_format', 'bullet_list'),
        'index': ph.get('index', ph.get('idx'))
    }

