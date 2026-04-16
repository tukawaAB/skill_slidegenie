"""JSON post-processing for OCR output normalization.

Ported from: ppt-addin/backend/services/make_pptx/json_gen/json_postprocess.py
"""
from collections import Counter

from slidegenie.utils.constants import PowerPointConfig


def normalize_font_sizes(json_data):
    """Normalize outlier font sizes to nearest group value."""

    def collect_items_with_font_size(data, results=None):
        if results is None:
            results = []
        if isinstance(data, dict):
            if "font_size" in data:
                results.append(data)
            for value in data.values():
                collect_items_with_font_size(value, results)
        elif isinstance(data, list):
            for item in data:
                collect_items_with_font_size(item, results)
        return results

    def find_nearest_group_value(value, group_values):
        nearest = None
        min_diff = float("inf")
        for gv in sorted(group_values):
            diff = abs(gv - value)
            if diff < min_diff or (diff == min_diff and (nearest is None or gv > nearest)):
                min_diff = diff
                nearest = gv
        return nearest

    items = collect_items_with_font_size(json_data)
    if not items:
        return

    font_sizes = [item["font_size"] for item in items]
    counter = Counter(font_sizes)
    group_values = {size for size, count in counter.items() if count > 1}
    if not group_values:
        return

    for item in items:
        if item["font_size"] not in group_values:
            item["font_size"] = find_nearest_group_value(item["font_size"], group_values)


def normalize_text_alignment(json_data):
    """Auto-set text_align based on content (list items → left, others → center)."""

    def collect_items_with_text(data, results=None):
        if results is None:
            results = []
        if isinstance(data, dict):
            if "text" in data:
                results.append(data)
            for value in data.values():
                collect_items_with_text(value, results)
        elif isinstance(data, list):
            for item in data:
                collect_items_with_text(item, results)
        return results

    def is_list_item_line(ln):
        s = ln.strip()
        if not s:
            return False
        if s[0] in PowerPointConfig.LIST_BULLET_CHARS:
            return True
        if s[0] in "-*" and len(s) > 1 and s[1] in " \t":
            return True
        if PowerPointConfig.LIST_NUMBERED_PATTERN.match(s):
            return True
        return False

    items = collect_items_with_text(json_data)
    for item in items:
        if "text_align" in item:
            continue
        text = item.get("text", "")
        lines = text.split("\n")
        item["text_align"] = "left" if any(is_list_item_line(ln) for ln in lines) else "center"


def json_postprocess(json_data):
    """Run all post-processing steps on OCR JSON data."""
    normalize_font_sizes(json_data)
    normalize_text_alignment(json_data)
    return json_data
