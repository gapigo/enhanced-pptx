"""Schema gate — validates chart XML parses cleanly."""

from __future__ import annotations

import lxml.etree as etree


def validate_chart_xml(chart_xml: bytes | str) -> bool:
    """Check that chart XML parses without error.

    Args:
        chart_xml: The ``<c:chartSpace>`` XML as bytes or string.

    Returns:
        ``True`` if parsing succeeds.
    """
    try:
        root = etree.fromstring(chart_xml) if isinstance(chart_xml, bytes) \
            else etree.fromstring(chart_xml.encode("utf-8"))
        return root is not None
    except etree.XMLSyntaxError:
        return False
