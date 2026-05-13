"""Namespace constants for OOXML chart-related XML.

Every namespace used across the package lives here in one place.
"""

NSMAP: dict[str, str] = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "14": "http://schemas.microsoft.com/office/drawing/2010/main",
}

A_NS = NSMAP["a"]
C_NS = NSMAP["c"]
P_NS = NSMAP["p"]
R_NS = NSMAP["r"]
