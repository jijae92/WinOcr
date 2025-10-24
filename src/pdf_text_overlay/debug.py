"""Debug drawing helpers for visualising OCR overlays."""

from __future__ import annotations

import logging
from typing import Iterable, Tuple

try:
    import fitz  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    fitz = None  # type: ignore

logger = logging.getLogger(__name__)


def _transform_point(
    x_px: float,
    y_px: float,
    width_px: float,
    height_px: float,
    width_pt: float,
    height_pt: float,
    rotation: int,
) -> Tuple[float, float]:
    scale_x = width_pt / width_px
    scale_y = height_pt / height_px
    x_pt = x_px * scale_x
    y_pt = height_pt - (y_px * scale_y)

    rotation_norm = rotation % 360
    if rotation_norm == 0:
        return x_pt, y_pt
    if rotation_norm == 90:
        return height_pt - y_pt, x_pt
    if rotation_norm == 180:
        return width_pt - x_pt, height_pt - y_pt
    if rotation_norm == 270:
        return y_pt, width_pt - x_pt
    logger.debug("Unsupported rotation %s; returning unrotated coordinates.", rotation)
    return x_pt, y_pt


def draw_debug_overlay(
    page: fitz.Page,
    bbox: Tuple[float, float, float, float],
    width_px: float,
    height_px: float,
    width_pt: float,
    height_pt: float,
    rotation: int,
    color: Tuple[float, float, float] = (0, 1, 0),
) -> None:
    """Draw a translucent rectangle representing OCR bounding box."""
    if fitz is None:  # pragma: no cover - defensive check
        raise RuntimeError("PyMuPDF is required for debug overlay drawing.")
    x_px, y_px, w_px, h_px = bbox
    corners: Iterable[Tuple[float, float]] = [
        _transform_point(x_px, y_px, width_px, height_px, width_pt, height_pt, rotation),
        _transform_point(x_px + w_px, y_px, width_px, height_px, width_pt, height_pt, rotation),
        _transform_point(x_px + w_px, y_px + h_px, width_px, height_px, width_pt, height_pt, rotation),
        _transform_point(x_px, y_px + h_px, width_px, height_px, width_pt, height_pt, rotation),
    ]
    points = [fitz.Point(x, y) for x, y in corners]
    shape = page.new_shape()
    shape.draw_polyline(points + [points[0]])
    shape.finish(color=color, fill=(color[0], color[1], color[2], 0.1))
    shape.commit()
