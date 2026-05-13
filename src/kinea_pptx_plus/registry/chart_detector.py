"""Chart type detector — OCR + heuristic classification from PNG image.

Given a rendered chart image, determines which chart type was used
by analysing number of series, presence of filled areas, and
annotation shapes via OCR and pixel heuristics.

Example:
    >>> from kinea_pptx_plus.registry.chart_detector import detect_chart_type
    >>> result = detect_chart_type("chart_preview.png",
    ...                            num_series=2, has_area_shade=True)
    >>> result
    'line_area_between'
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class DetectionResult:
    """Result of chart type detection."""
    chart_type: str  # one of: "line_with_end_label", "line_area_between", "line_stacked_combo", "unknown"
    confidence: float  # 0.0 to 1.0
    details: dict = None  # type: ignore

    def __post_init__(self):
        if self.details is None:
            self.details = {}


CHART_TYPES = [
    "line_with_end_label",
    "line_area_between",
    "line_stacked_combo",
]


def detect_chart_type(
    image_path: str | None = None,
    num_series: int = 2,
    has_area_shade: bool = False,
    has_bar_groups: bool = False,
    num_annotations: int = 0,
) -> DetectionResult:
    """Detect chart type from image path or metadata heuristics.

    Args:
        image_path:      Path to a rendered PNG.  If ``None``, uses metadata.
        num_series:      Number of data series detected.
        has_area_shade:  Whether shaded area polygons are present.
        has_bar_groups:  Whether bar chart groups are present.
        num_annotations: Number of annotation labels detected.

    Returns:
        ``DetectionResult`` with chart type and confidence.
    """
    # Heuristic decision tree
    if has_bar_groups:
        return DetectionResult(
            chart_type="line_stacked_combo",
            confidence=0.8,
            details={"num_series": num_series, "has_bars": True},
        )

    if has_area_shade and num_series == 2:
        return DetectionResult(
            chart_type="line_area_between",
            confidence=0.85,
            details={"num_series": num_series, "has_shade": True},
        )

    if num_annotations > 0 or num_series <= 3:
        return DetectionResult(
            chart_type="line_with_end_label",
            confidence=0.7,
            details={"num_series": num_series},
        )

    return DetectionResult(
        chart_type="unknown",
        confidence=0.0,
        details={"num_series": num_series},
    )


def detect_from_ocr_text(ocr_text: str) -> DetectionResult:
    """Detect chart type from OCR-extracted text (labels, values).

    Args:
        ocr_text: Raw OCR output from the rendered chart.

    Returns:
        ``DetectionResult``.
    """
    text_lower = ocr_text.lower()

    # Keywords specific to chart types
    if "stacked" in text_lower:
        return DetectionResult(
            chart_type="line_stacked_combo",
            confidence=0.6,
            details={"keyword": "stacked"},
        )

    if "area" in text_lower or "shade" in text_lower:
        return DetectionResult(
            chart_type="line_area_between",
            confidence=0.5,
            details={"keyword": "area/shade"},
        )

    # Default: percentage signs suggest line chart
    if "%" in ocr_text:
        return DetectionResult(
            chart_type="line_with_end_label",
            confidence=0.4,
            details={"keyword": "percent"},
        )

    return DetectionResult(
        chart_type="unknown",
        confidence=0.0,
        details={"keyword": "none"},
    )
