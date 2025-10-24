"""Debug drawing helpers for visualising OCR overlays."""

from __future__ import annotations

import logging
from typing import Iterable, Tuple

try:
    import fitz  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    fitz = None  # type: ignore

logger = logging.getLogger(__name__)


def draw_debug_overlay(
    page: fitz.Page,
    rect: "fitz.Rect",
    color: Tuple[float, float, float] = (0, 1, 0),
) -> None:
    """Draw a translucent rectangle representing OCR bounding box."""
    if fitz is None:  # pragma: no cover - defensive check
        raise RuntimeError("PyMuPDF is required for debug overlay drawing.")
    corners = [
        fitz.Point(rect.x0, rect.y0),
        fitz.Point(rect.x1, rect.y0),
        fitz.Point(rect.x1, rect.y1),
        fitz.Point(rect.x0, rect.y1),
    ]
    shape = page.new_shape()
    shape.draw_polyline(corners + [corners[0]])
    shape.finish(color=color, fill=(color[0], color[1], color[2], 0.1))
    shape.commit()


def draw_visible_text(page: fitz.Page, rect: "fitz.Rect", rotation: float) -> None:
    """Overlay a faint gray rectangle to indicate QA text placement."""
    if fitz is None:  # pragma: no cover
        raise RuntimeError("PyMuPDF is required for debug overlay drawing.")
    shape = page.new_shape()
    corners = [
        fitz.Point(rect.x0, rect.y0),
        fitz.Point(rect.x1, rect.y0),
        fitz.Point(rect.x1, rect.y1),
        fitz.Point(rect.x0, rect.y1),
    ]
    shape.draw_polyline(corners + [corners[0]])
    shape.finish(color=(0.6, 0.6, 0.6), fill=(0.6, 0.6, 0.6, 0.05))
    shape.commit()
