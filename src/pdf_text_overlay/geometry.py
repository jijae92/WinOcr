"""Geometry helpers for mapping OCR coordinates into PDF space."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Tuple


@dataclass
class Rect:
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def height(self) -> float:
        return self.y1 - self.y0


@dataclass
class MappingConfig:
    image_rect: Rect
    page_rect: Rect
    width_px: float
    height_px: float
    offset_pt: Tuple[float, float] = (0.0, 0.0)
    scale_corr: Tuple[float, float] = (1.0, 1.0)
    rotation: int = 0
    deskew: float = 0.0


@dataclass
class Placement:
    anchor: Tuple[float, float]
    rect: Rect
    font_size: float
    rotate: float
    width_pt: float
    height_pt: float


def map_bbox_to_pdf(
    bbox: Tuple[float, float, float, float],
    baseline_ratio: float,
    font_scale: float,
    config: MappingConfig,
) -> Placement:
    """Map a pixel-space bounding box to PDF space using the provided configuration."""
    x_px, y_px, w_px, h_px = bbox
    if config.width_px <= 0 or config.height_px <= 0:
        raise ValueError("Mapping configuration requires positive width/height in pixels.")

    scale_x = (config.image_rect.width / config.width_px) * config.scale_corr[0]
    scale_y = (config.image_rect.height / config.height_px) * config.scale_corr[1]
    if scale_x == 0 or scale_y == 0:
        raise ValueError("Scaling factors must be non-zero.")

    offset_x, offset_y = config.offset_pt

    x_pt = config.image_rect.x0 + (x_px * scale_x) + offset_x
    y_bottom = config.image_rect.y1 - ((y_px + h_px) * scale_y) + offset_y

    width_pt = max(w_px * scale_x, 0.1)
    height_pt = max(h_px * scale_y, 0.1)

    font_size = max(height_pt * font_scale, 2.0)
    baseline = y_bottom + font_size * baseline_ratio

    rect = Rect(
        x0=x_pt,
        y0=baseline - height_pt * (1 - baseline_ratio),
        x1=x_pt + width_pt,
        y1=baseline + height_pt * baseline_ratio,
    )

    anchor = apply_rotation((x_pt, baseline), config.page_rect, config.rotation)
    rotated_rect = rotate_rect(rect, config.page_rect, config.rotation)

    return Placement(
        anchor=anchor,
        rect=rotated_rect,
        font_size=font_size,
        rotate=config.rotation + config.deskew,
        width_pt=width_pt,
        height_pt=height_pt,
    )


def apply_rotation(
    point: Tuple[float, float],
    page_rect: Rect,
    rotation: int,
) -> Tuple[float, float]:
    """Rotate a point according to the page rotation (multiples of 90)."""
    x, y = point
    rot = rotation % 360
    if rot == 0:
        return (x, y)
    if rot == 90:
        return (page_rect.width - y, x)
    if rot == 180:
        return (page_rect.width - x, page_rect.height - y)
    if rot == 270:
        return (y, page_rect.height - x)
    raise ValueError("Rotation must be 0/90/180/270 degrees.")


def rotate_rect(rect: Rect, page_rect: Rect, rotation: int) -> Rect:
    """Rotate a rectangle axis-aligned with rotation multiples of 90 degrees."""
    rot = rotation % 360
    if rot == 0:
        return rect
    if rot == 180:
        return Rect(
            x0=page_rect.width - rect.x1,
            y0=page_rect.height - rect.y1,
            x1=page_rect.width - rect.x0,
            y1=page_rect.height - rect.y0,
        )
    if rot == 90:
        return Rect(
            x0=page_rect.width - rect.y1,
            y0=rect.x0,
            x1=page_rect.width - rect.y0,
            y1=rect.x1,
        )
    if rot == 270:
        return Rect(
            x0=rect.y0,
            y0=page_rect.height - rect.x1,
            x1=rect.y1,
            y1=page_rect.height - rect.x0,
        )
    raise ValueError("Rotation must be 0/90/180/270 degrees.")
