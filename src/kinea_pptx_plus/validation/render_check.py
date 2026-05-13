"""Render-OCR gate — converts PPTX to PNG via LibreOffice and runs OCR.

Example:
    >>> from kinea_pptx_plus.validation.render_check import render_check
    >>> result = render_check("output.pptx")
    >>> result.passed
    True
"""

from __future__ import annotations

import os
import subprocess
import tempfile
from dataclasses import dataclass


@dataclass
class RenderCheckResult:
    """Result of the render-OCR gate."""
    passed: bool
    texts_found: list[str] = None  # type: ignore
    detail: str = ""


def render_check(
    pptx_path: str,
    expected_texts: list[str] | None = None,
) -> RenderCheckResult:
    """Render PPTX to PNG via LibreOffice and check for expected text.

    Args:
        pptx_path:     Path to a ``.pptx`` file.
        expected_texts: Optional list of strings to search for in OCR.

    Returns:
        ``RenderCheckResult``.
    """
    # Check for LibreOffice
    soffice = _find_soffice()
    if soffice is None:
        return RenderCheckResult(
            passed=False,
            detail="SKIP: LibreOffice not found in PATH",
        )

    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "png",
                 "--outdir", tmpdir, os.path.abspath(pptx_path)],
                capture_output=True, timeout=60, check=True,
            )
        except (subprocess.TimeoutExpired, subprocess.CalledProcessError) as exc:
            return RenderCheckResult(
                passed=False,
                detail=f"FAIL: render error: {exc}",
            )

        # Find generated PNG
        pngs = [f for f in os.listdir(tmpdir) if f.endswith(".png")]
        if not pngs:
            return RenderCheckResult(
                passed=False,
                detail="FAIL: no PNG produced",
            )

        # Run OCR via pytesseract
        try:
            from PIL import Image
            import pytesseract
        except ImportError:
            return RenderCheckResult(
                passed=False,
                detail="FAIL: Pillow or pytesseract not installed",
            )

        try:
            img = Image.open(os.path.join(tmpdir, pngs[0]))
            text = pytesseract.image_to_string(img)
        except Exception as exc:
            return RenderCheckResult(
                passed=False,
                detail=f"FAIL: OCR error: {exc}",
            )

        texts_found = []

        if expected_texts:
            for et in expected_texts:
                if et.lower() in text.lower():
                    texts_found.append(et)

            missing = set(expected_texts) - set(texts_found)
            if missing:
                return RenderCheckResult(
                    passed=False,
                    texts_found=texts_found,
                    detail=f"FAIL: missing texts: {missing}",
                )

        return RenderCheckResult(
            passed=True,
            texts_found=texts_found or [],
            detail=f"OK: {len(pngs)} PNG(s) rendered, OCR complete",
        )


def _find_soffice() -> str | None:
    """Locate the LibreOffice binary."""
    candidates = ["soffice", "soffice.exe"]
    if os.name == "nt":
        candidates.extend([
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ])
    for c in candidates:
        try:
            subprocess.run([c, "--version"], capture_output=True, timeout=5)
            return c
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    return None
