"""Overlay OCR text as an invisible/searchable layer on PDF pages."""

from __future__ import annotations

import logging
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple

try:
    import fitz  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    fitz = None  # type: ignore

from .debug import draw_debug_overlay
from .fonts import register_font
from .ocr_io import OCRPage

logger = logging.getLogger(__name__)


@dataclass
class OverlaySettings:
    """Overlay configuration passed from CLI."""

    method: str = "invisible"
    granularity: str = "word"
    baseline_ratio: float = 0.15
    keep_spaces: bool = False
    dehyphen: bool = False
    debug_overlay: bool = False
    pdfa: bool = False
    font_path: Optional[Path] = None
    dpi: int = 300


def _normalize_text(text: str, keep_spaces: bool) -> str:
    """Normalize OCR text for overlay."""
    normalized = unicodedata.normalize("NFKC", text)
    normalized = normalized.replace("\u200b", " ").replace("\u200c", "").replace("\u200d", "").replace("\ufeff", "")
    if keep_spaces:
        return normalized
    collapsed = " ".join(normalized.split())
    return collapsed


def _sort_key(entry: Tuple[float, float]) -> Tuple[float, float]:
    y, x = entry
    return (round(y, 2), round(x, 2))


def _convert_bbox_to_point(
    bbox: Tuple[float, float, float, float],
    width_px: float,
    height_px: float,
    width_pt: float,
    height_pt: float,
    baseline_ratio: float,
) -> Tuple[float, float, float]:
    """Convert top-left pixel bbox to PDF point coordinates."""
    x_px, y_px, w_px, h_px = bbox
    if w_px <= 0 or h_px <= 0:
        raise ValueError("Bounding box width/height must be positive.")
    scale_x = width_pt / width_px
    scale_y = height_pt / height_px

    baseline_px = y_px + h_px * baseline_ratio
    x_pt = x_px * scale_x
    y_pt = height_pt - (baseline_px * scale_y)
    font_size = h_px * scale_y
    return (x_pt, y_pt, font_size)


def _apply_rotation(
    x_pt: float,
    y_pt: float,
    width_pt: float,
    height_pt: float,
    rotation: int,
) -> Tuple[float, float, int]:
    """Rotate anchor point according to page rotation."""
    rotation_norm = rotation % 360
    if rotation_norm == 0:
        return x_pt, y_pt, 0
    if rotation_norm == 90:
        return height_pt - y_pt, x_pt, 90
    if rotation_norm == 180:
        return width_pt - x_pt, height_pt - y_pt, 180
    if rotation_norm == 270:
        return y_pt, width_pt - x_pt, 270
    logger.warning("Unsupported rotation %s detected; proceeding without rotation.", rotation)
    return x_pt, y_pt, 0


def _iter_granularity(
    page: OCRPage,
    granularity: str,
) -> Iterable[Tuple[str, Tuple[float, float, float, float]]]:
    if granularity == "line":
        for line in page.lines:
            yield line.text, line.bbox
    else:
        for word in page.words:
            yield word.text, word.bbox


def apply_text_overlay(
    pdf_path: Path,
    pages: Sequence[OCRPage],
    output_path: Path,
    settings: OverlaySettings,
) -> None:
    """Insert invisible OCR text into a PDF."""
    if fitz is None:
        raise RuntimeError("PyMuPDF (fitz) is required for overlay generation.")
    doc = fitz.open(pdf_path)
    font_name = register_font(doc, settings.font_path)
    logger.info("Using font '%s' for OCR overlay.", font_name)

    page_map = {page.index: page for page in pages}

    missing_indices = sorted(set(range(len(doc))) - set(page_map.keys()))
    if missing_indices:
        logger.warning("OCR data missing for pages: %s", ", ".join(str(i) for i in missing_indices))

    for index in range(len(doc)):
        fitz_page = doc.load_page(index)
        ocr_page = page_map.get(index)
        if ocr_page is None:
            logger.debug("Skipping page %d (no OCR data).", index)
            continue

        logger.debug("Overlaying page %d (rotation=%d).", index, fitz_page.rotation)
        width_pt = float(fitz_page.rect.width)
        height_pt = float(fitz_page.rect.height)
        width_px = float(ocr_page.width_px)
        height_px = float(ocr_page.height_px)
        rotation = ocr_page.rotation if ocr_page.rotation is not None else fitz_page.rotation

        entries = list(_iter_granularity(ocr_page, settings.granularity))
        if not entries:
            continue
        sorted_entries = sorted(
            entries,
            key=lambda item: _sort_key((item[1][1], item[1][0])),
        )

        prepared_entries: List[Tuple[str, Tuple[float, float, float, float]]] = []
        for raw_text, bbox in sorted_entries:
            normalized = _normalize_text(raw_text, settings.keep_spaces)
            if (
                settings.dehyphen
                and settings.granularity == "line"
                and prepared_entries
                and prepared_entries[-1][0].endswith("-")
                and normalized
                and normalized[0].islower()
            ):
                previous_text, previous_bbox = prepared_entries[-1]
                prepared_entries[-1] = (previous_text[:-1] + normalized, previous_bbox)
                continue
            if (
                settings.dehyphen
                and settings.granularity == "line"
                and normalized.endswith("-")
            ):
                normalized = normalized[:-1]
            prepared_entries.append((normalized, bbox))

        for text, bbox in prepared_entries:
            try:
                x_pt, y_pt, font_size = _convert_bbox_to_point(
                    bbox,
                    width_px=width_px,
                    height_px=height_px,
                    width_pt=width_pt,
                    height_pt=height_pt,
                    baseline_ratio=settings.baseline_ratio,
                )
            except ValueError:
                logger.debug("Skipping invalid bbox on page %d: %s", index, bbox)
                continue

            x_rot, y_rot, rotate = _apply_rotation(x_pt, y_pt, width_pt, height_pt, rotation)
            if not text:
                continue

            font_size = max(font_size, 2.0)
            options = {
                "fontname": font_name,
                "fontsize": font_size,
                "overlay": True,
                "rotate": rotate,
            }

            if settings.method == "invisible":
                options["render_mode"] = 3
            else:
                options["render_mode"] = 0
                options["opacity"] = 0.02
                options["color"] = (0, 0, 0)

            fitz_point = fitz.Point(x_rot, y_rot)
            fitz_page.insert_text(fitz_point, text, **options)

            if settings.debug_overlay:
                draw_debug_overlay(
                    fitz_page,
                    bbox=bbox,
                    width_px=width_px,
                    height_px=height_px,
                    width_pt=width_pt,
                    height_pt=height_pt,
                    rotation=rotation,
                    color=(1, 0, 0),
                )

    output_path.parent.mkdir(parents=True, exist_ok=True)

    temp_path = output_path.with_suffix(".tmp.pdf")
    doc.save(temp_path, deflate=True, garbage=4)
    doc.close()

    if settings.pdfa:
        _attempt_pdfa_conversion(temp_path, output_path)
    else:
        temp_path.replace(output_path)
        logger.info("Saved searchable PDF to %s", output_path)


def _attempt_pdfa_conversion(temp_path: Path, output_path: Path) -> None:
    """Try to convert to PDF/A using pikepdf."""
    try:
        import pikepdf  # type: ignore
    except ImportError:
        logger.warning("pikepdf not available; exporting regular PDF instead of PDF/A-2b.")
        temp_path.replace(output_path)
        return

    success = False
    try:
        with pikepdf.open(temp_path) as pdf:
            try:
                kwargs = {}
                compliance = getattr(pikepdf.Pdf, "PDFA_2B", None)
                if compliance is not None:
                    kwargs["compliance"] = compliance
                sRGB_profile = getattr(pikepdf, "sRGB_PROFILE", None)
                if sRGB_profile is not None:
                    kwargs["icc_profile"] = sRGB_profile
                pdf.make_pdfa(output_path, **kwargs)
                success = True
                logger.info("Saved PDF/A-2b (best effort) file to %s", output_path)
            except Exception as exc:
                logger.warning("Failed to enforce PDF/A-2b (%s); emitting regular PDF.", exc)
                pdf.save(output_path)
                success = True
    finally:
        temp_path.unlink(missing_ok=True)

    if not success:
        logger.warning("PDF/A conversion was not applied; output may not be compliant.")
