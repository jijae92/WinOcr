"""Overlay OCR text onto PDFs with precise image alignment."""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import fitz  # type: ignore

from .debug import draw_debug_overlay, draw_visible_text
from .fonts import register_font
from .geometry import MappingConfig, Rect, map_bbox_to_pdf
from .ocr_io import OCRLine, OCRPage, OCRWord
from .text_utils import dehyphenize, normalize_token

logger = logging.getLogger(__name__)

ALIGN_AUTO = "auto"
ALIGN_PAGE = "page"


@dataclass
class OverlaySettings:
    method: str = "invisible"
    granularity: str = "word"
    baseline_ratio: float = 0.15
    font_scale: float = 1.0
    keep_spaces: bool = False
    dehyphen: bool = False
    cjk_join: bool = False
    debug_overlay: bool = False
    visible_qa: bool = False
    pdfa: bool = False
    font_path: Optional[Path] = None
    dpi: Optional[int] = None
    align: str = ALIGN_AUTO
    offset_pt: Tuple[float, float] = (0.0, 0.0)
    scale_corr: Tuple[float, float] = (1.0, 1.0)
    rotate_override: Optional[int] = None
    deskew: float = 0.0
    calibrate: int = 0
    dump_debug_json: Optional[Path] = None


@dataclass
class AlignmentInfo:
    rect: Rect
    width_px: float
    height_px: float
    rotation: int
    xref: Optional[int] = None
    source: str = ""


def apply_text_overlay(
    pdf_path: Path,
    pages: Sequence[OCRPage],
    output_path: Path,
    settings: OverlaySettings,
) -> None:
    """Insert normalized OCR text into a PDF using precise alignment."""
    doc = fitz.open(pdf_path)
    font_name = register_font(doc, settings.font_path)
    logger.info("Using font '%s' for OCR overlay.", font_name)

    page_map = {page.index: page for page in pages}
    debug_payload: List[Dict[str, object]] = []

    for index in range(len(doc)):
        fitz_page = doc.load_page(index)
        ocr_page = page_map.get(index)
        if ocr_page is None:
            logger.debug("Skipping page %d (no OCR data).", index)
            continue

        alignment = _determine_alignment(doc, fitz_page, ocr_page, settings)
        logger.info(
            "Page %d: rect=%s px_size=(%.1f, %.1f) rotation=%d source=%s",
            index,
            alignment.rect,
            alignment.width_px,
            alignment.height_px,
            alignment.rotation,
            alignment.source,
        )
        _overlay_single_page(
            fitz_page=fitz_page,
            ocr_page=ocr_page,
            font_name=font_name,
            alignment=alignment,
            settings=settings,
            debug_payload=debug_payload,
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

    if settings.dump_debug_json:
        settings.dump_debug_json.parent.mkdir(parents=True, exist_ok=True)
        settings.dump_debug_json.write_text(json.dumps(debug_payload, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info("Dumped debug mapping to %s", settings.dump_debug_json)


def _determine_alignment(
    doc: fitz.Document,
    page: fitz.Page,
    ocr_page: OCRPage,
    settings: OverlaySettings,
) -> AlignmentInfo:
    page_rect = Rect(*page.bound().xf)  # type: ignore[arg-type]
    rotation = page.rotation
    if settings.rotate_override is not None:
        rotation = settings.rotate_override

    align = settings.align or ALIGN_AUTO
    if align == ALIGN_PAGE:
        width_px, height_px = _page_dimensions_px(ocr_page, settings)
        return AlignmentInfo(rect=page_rect, width_px=width_px, height_px=height_px, rotation=rotation, source="page")

    if align.startswith("image-rect:"):
        rect = _parse_manual_rect(align.split(":", 1)[1])
        width_px = ocr_page.width_px or _page_dimensions_px(ocr_page, settings)[0]
        height_px = ocr_page.height_px or _page_dimensions_px(ocr_page, settings)[1]
        return AlignmentInfo(rect=rect, width_px=width_px, height_px=height_px, rotation=rotation, source="manual")

    xref_target = None
    if align.startswith("image:"):
        try:
            xref_target = int(align.split(":", 1)[1])
        except ValueError as exc:
            raise ValueError("--align image:<xref> expects numeric xref") from exc

    image_entries = page.get_images(full=True)
    candidates: List[Tuple[int, Rect]] = []
    for entry in image_entries:
        xref = entry[0]
        if xref_target is not None and xref_target != xref:
            continue
        rects = page.get_image_rects(xref)
        if not rects:
            continue
        largest_rect = max(rects, key=lambda r: r.width * r.height)
        candidates.append((xref, Rect(*largest_rect)))

    if not candidates:
        logger.warning("No image rectangles found on page %d; falling back to page alignment.", page.number)
        width_px, height_px = _page_dimensions_px(ocr_page, settings)
        return AlignmentInfo(rect=page_rect, width_px=width_px, height_px=height_px, rotation=rotation, source="page-fallback")

    if xref_target is not None:
        chosen = candidates[0]
    else:
        chosen = max(candidates, key=lambda item: item[1].width * item[1].height)
    xref, image_rect = chosen

    info = doc.extract_image(xref)
    width_px = float(info.get("width", ocr_page.width_px or 0))
    height_px = float(info.get("height", ocr_page.height_px or 0))
    if width_px == 0 or height_px == 0:
        logger.warning("Image xref=%d missing pixel size metadata; using OCR JSON dimensions.", xref)
        width_px = ocr_page.width_px or 0.0
        height_px = ocr_page.height_px or 0.0
    if width_px == 0 or height_px == 0:
        width_px, height_px = _page_dimensions_px(ocr_page, settings)

    return AlignmentInfo(
        rect=image_rect,
        width_px=width_px,
        height_px=height_px,
        rotation=rotation,
        xref=xref,
        source=f"image:xref={xref}",
    )


def _parse_manual_rect(data: str) -> Rect:
    parts = [float(item) for item in data.split(",")]
    if len(parts) != 4:
        raise ValueError("--align image-rect:x0,y0,x1,y1 expects four comma-separated numbers.")
    x0, y0, x1, y1 = parts
    return Rect(x0, y0, x1, y1)


def _page_dimensions_px(ocr_page: OCRPage, settings: OverlaySettings) -> Tuple[float, float]:
    if ocr_page.width_px and ocr_page.height_px:
        return float(ocr_page.width_px), float(ocr_page.height_px)
    if not settings.dpi:
        raise ValueError("OCR JSON missing width/height and --dpi not provided.")
    # width_pt = 72 * width_inch
    # width_px = dpi * width_inch -> width_px = width_pt * dpi / 72
    raise ValueError("OCR JSON must include width_px/height_px when align=page.")


def _overlay_single_page(
    fitz_page: fitz.Page,
    ocr_page: OCRPage,
    font_name: str,
    alignment: AlignmentInfo,
    settings: OverlaySettings,
    debug_payload: List[Dict[str, object]],
) -> None:
    mapping = MappingConfig(
        image_rect=alignment.rect,
        page_rect=Rect(*fitz_page.bound().xf),  # type: ignore[arg-type]
        width_px=alignment.width_px,
        height_px=alignment.height_px,
        offset_pt=settings.offset_pt,
        scale_corr=settings.scale_corr,
        rotation=alignment.rotation,
        deskew=settings.deskew,
    )

    entries = list(_iter_granularity(ocr_page, settings.granularity))
    if settings.dehyphen and settings.granularity == "line":
        normalized_lines = [
            normalize_token(text, settings.keep_spaces, settings.cjk_join) for text, _ in entries
        ]
        normalized_lines = dehyphenize(normalized_lines)
        entries = list(zip(normalized_lines, [bbox for _, bbox in entries]))

    samples: List[Dict[str, object]] = []
    color_cycle = _color_cycle()
    method = settings.method
    if settings.visible_qa:
        method = "visible"

    render_mode, color, opacity = _resolve_render_mode(method)

    for idx, (raw_text, bbox) in enumerate(entries):
        text = normalize_token(raw_text, settings.keep_spaces, settings.cjk_join)
        if not text:
            continue
        try:
            placement = map_bbox_to_pdf(
                bbox=bbox,
                baseline_ratio=settings.baseline_ratio,
                font_scale=settings.font_scale,
                config=mapping,
            )
        except ValueError as exc:
            logger.debug("Skipping bbox %s due to %s", bbox, exc)
            continue

        if idx < settings.calibrate:
            samples.append(
                {
                    "text": text,
                    "bbox_px": [float(value) for value in bbox],
                    "anchor_pt": placement.anchor,
                    "font_size": placement.font_size,
                }
            )

        rect = fitz.Rect(placement.rect.x0, placement.rect.y0, placement.rect.x1, placement.rect.y1)
        fitz_point = fitz.Point(*placement.anchor)
        options = {
            "fontname": font_name,
            "fontsize": placement.font_size,
            "rotate": placement.rotate,
            "render_mode": render_mode,
            "overlay": True,
        }
        if color is not None:
            options["color"] = color
        if opacity is not None:
            options["opacity"] = opacity

        fitz_page.insert_text(fitz_point, text, **options)

        if method == "visible":
            draw_visible_text(fitz_page, rect, placement.rotate)

        if settings.debug_overlay:
            debug_color = next(color_cycle)
            draw_debug_overlay(fitz_page, rect, debug_color)

        if settings.dump_debug_json is not None:
            debug_payload.append(
                {
                    "page": fitz_page.number,
                    "text": text,
                    "bbox_px": [float(value) for value in bbox],
                    "rect_pt": [placement.rect.x0, placement.rect.y0, placement.rect.x1, placement.rect.y1],
                    "anchor_pt": placement.anchor,
                    "font_size": placement.font_size,
                }
            )

    if samples:
        logger.info("Calibration samples (first %d words):", len(samples))
        for sample in samples:
            logger.info("  %s -> anchor %s font %.2f", sample["text"], sample["anchor_pt"], sample["font_size"])


def _iter_granularity(page: OCRPage, granularity: str) -> Iterable[Tuple[str, Tuple[float, float, float, float]]]:
    if granularity == "line":
        for line in page.lines:
            yield line.text, line.bbox
    else:
        for word in page.words:
            yield word.text, word.bbox


def _color_cycle() -> Iterable[Tuple[float, float, float]]:
    colors = [
        (1.0, 0.2, 0.2),
        (0.2, 1.0, 0.2),
        (0.2, 0.5, 1.0),
        (1.0, 0.7, 0.2),
        (0.8, 0.2, 1.0),
    ]
    idx = 0
    while True:
        yield colors[idx % len(colors)]
        idx += 1


def _resolve_render_mode(method: str) -> Tuple[int, Optional[Tuple[float, float, float]], Optional[float]]:
    if method == "opacity":
        return 0, (0.0, 0.0, 0.0), 0.02
    if method == "visible":
        return 0, (0.6, 0.6, 0.6), 1.0
    return 3, None, None


def _attempt_pdfa_conversion(temp_path: Path, output_path: Path) -> None:
    try:
        import pikepdf  # type: ignore
    except ImportError:
        logger.warning("pikepdf not available; exporting regular PDF instead of PDF/A-2b.")
        temp_path.replace(output_path)
        return

    try:
        with pikepdf.open(temp_path) as pdf:
            try:
                pdf.make_pdfa(output_path, compliance=pikepdf.Pdf.PDFA_2B, icc_profile=pikepdf.sRGB_PROFILE)  # type: ignore[attr-defined]
                logger.info("Saved PDF/A-2b (best effort) file to %s", output_path)
            except Exception as exc:
                logger.warning("Failed to enforce PDF/A-2b (%s); emitting regular PDF.", exc)
                pdf.save(output_path)
    finally:
        temp_path.unlink(missing_ok=True)
