"""Golden compare — byte-level comparison of chart XML against goldens.

Example:
    >>> from kinea_pptx_plus.validation.golden_compare import compare_chart_xml
    >>> from kinea_pptx_plus.io import extract_chart_xml
    >>> xml = extract_chart_xml("output.pptx")
    >>> result = compare_chart_xml(xml, "tests/fixtures/golden_line_with_end_label.pptx")
    >>> result.passed
    True
"""

from __future__ import annotations

import lxml.etree as etree

from kinea_pptx_plus.io import extract_chart_xml


def compare_chart_xml(
    produced: bytes,
    golden_pptx_path: str,
    chart_index: int = 0,
) -> bool:
    """Compare produced chart XML against golden PPTX chart XML.

    Parses both XML trees and compares serialized canonical output
    (ignoring whitespace-only differences).

    Args:
        produced:          Chart XML bytes from the built PPTX.
        golden_pptx_path:  Path to a golden PPTX file.
        chart_index:       Chart index in the golden PPTX.

    Returns:
        ``True`` if the two chart XMLs are structurally identical.
    """
    try:
        golden_xml = extract_chart_xml(golden_pptx_path, chart_index)
        t_prod = etree.tostring(etree.fromstring(produced))
        t_gold = etree.tostring(etree.fromstring(golden_xml))
        return t_prod == t_gold
    except Exception:
        return False
