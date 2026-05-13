"""Win32 probe gate — opens PPTX in PowerPoint via COM to detect corruption.

Only runs on Windows (``os.name == 'nt'``).  Skips silently otherwise.

Example:
    >>> from kinea_pptx_plus.validation.win32_probe import probe_pptx
    >>> result = probe_pptx("output.pptx")
    >>> result.passed  # True if no corruption detected
    True
"""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass


@dataclass
class Win32ProbeResult:
    """Result of the win32 probe gate."""
    passed: bool
    detail: str = ""


def probe_pptx(pptx_path: str) -> Win32ProbeResult:
    """Open the PPTX in PowerPoint via COM to detect corruption.

    Only runs on Windows.  On other platforms returns ``passed=True``
    with a skip note.

    Args:
        pptx_path: Path to a ``.pptx`` file.

    Returns:
        ``Win32ProbeResult`` with ``passed`` flag.
    """
    if os.name != "nt":
        return Win32ProbeResult(
            passed=True,
            detail="SKIP: win32 probe only runs on Windows",
        )

    try:
        import win32com.client  # type: ignore[import-untyped]
    except ImportError:
        return Win32ProbeResult(
            passed=False,
            detail="FAIL: pywin32 not installed (pip install kinea-pptx-plus[windows])",
        )

    try:
        abs_path = os.path.abspath(pptx_path)
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.Visible = 0  # type: ignore[attr-defined]
        pres = app.Presentations.Open(abs_path)  # type: ignore[attr-defined]
        pres.Close()  # type: ignore[attr-defined]
        app.Quit()  # type: ignore[attr-defined]
        return Win32ProbeResult(passed=True, detail="OK")
    except Exception as exc:
        return Win32ProbeResult(
            passed=False,
            detail=f"FAIL: {exc}",
        )
