"""Tree operations for OOXML chart XML: ordered insert, replace, deep clone.

These helpers work on ``lxml.etree`` elements and assume the
chart-space / plot-area document structure.

Example:
    >>> from lxml import etree
    >>> parent = etree.Element("{c}chartSpace")
    >>> child = etree.Element("{c}chart")
    >>> insert_sorted(child, parent, key="c:plotArea")
"""

from __future__ import annotations

import copy
from typing import Callable

import lxml.etree as etree


def insert_sorted(
    new_elem: etree.Element,
    parent: etree.Element,
    key: str | Callable[[etree.Element], str] = "",
) -> None:
    """Insert *new_elem* into *parent* at the position defined by *key*.

    Args:
        new_elem: The element to insert.
        parent:   The parent container.
        key:      If a string, insert **after** the element whose
                  local-name (or full ``{ns}tag``) matches.
                  If a callable, ``key(child)`` returns the reference tag
                  to insert after.  Empty string appends.

    Example:
        >>> from lxml import etree
        >>> root = etree.Element("{c}chartSpace")
        >>> chart = etree.SubElement(root, "{c}chart")
        >>> plot = etree.Element("{c}plotArea")
        >>> insert_sorted(plot, root, "c:chart")  # after c:chart
    """
    if not key:
        parent.append(new_elem)
        return

    ref_tag = key if isinstance(key, str) else ""
    for idx, child in enumerate(parent):
        match = child.tag == ref_tag or (
            not ref_tag and key(child) == ref_tag  # type: ignore[arg-type]
        )
        if match:
            parent.insert(idx + 1, new_elem)
            return
    parent.append(new_elem)


def replace_child(
    parent: etree.Element,
    old: etree.Element,
    new: etree.Element,
) -> None:
    """Replace *old* child of *parent* with *new* (position-preserving)."""
    for idx, child in enumerate(parent):
        if child is old:
            parent.remove(old)
            parent.insert(idx, new)
            return
    raise ValueError("old element is not a child of parent")


def deep_clone(elem: etree.Element) -> etree.Element:
    """Return a deep copy of *elem* including all children and text."""
    return copy.deepcopy(elem)
