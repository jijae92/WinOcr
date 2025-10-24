"""OCR data structures, JSON IO, and optional WinRT OCR integration."""

from __future__ import annotations

import asyncio
import json
import logging
import os
import platform
import struct
import sys
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import fitz  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    fitz = None  # type: ignore

try:
    from PIL import Image
except ImportError:  # pragma: no cover - optional dependency
    Image = None  # type: ignore

logger = logging.getLogger(__name__)


@dataclass
class OCRWord:
    text: str
    bbox: Tuple[float, float, float, float]


@dataclass
class OCRLine:
    text: str
    bbox: Tuple[float, float, float, float]
    words: List[OCRWord] = field(default_factory=list)


@dataclass
class OCRPage:
    index: int
    width_px: float
    height_px: float
    rotation: Optional[int]
    words: List[OCRWord] = field(default_factory=list)
    lines: List[OCRLine] = field(default_factory=list)


class WinRTUnavailableError(RuntimeError):
    """Raised when WinRT OCR cannot be used."""


class _RectProtocol:
    __slots__ = ()

    x: float
    y: float
    width: float
    height: float


def _extract_rect(entity: Any) -> Optional[_RectProtocol]:
    """Try to extract a rectangle-like object from WinRT OCR entities."""
    for attr in ("bounding_rect", "boundingRect", "BoundingRect", "rect"):
        rect = getattr(entity, attr, None)
        if rect is not None:
            return rect
    return None


def _decode_bbox(value: Sequence[float]) -> Tuple[float, float, float, float]:
    if len(value) != 4:
        raise ValueError("Bounding box must contain 4 numeric entries.")
    x, y, w, h = map(float, value)
    return (x, y, w, h)


def load_ocr_json(path: Path) -> List[OCRPage]:
    """Load OCR JSON file into structured objects."""
    if not path.is_file():
        raise FileNotFoundError(f"OCR JSON not found: {path}")
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, list):
        raise ValueError("OCR JSON must be a list of page objects.")

    pages: List[OCRPage] = []
    for entry in payload:
        index = int(entry.get("page", len(pages)))
        width_px = float(entry.get("width_px") or entry.get("width"))
        height_px = float(entry.get("height_px") or entry.get("height"))
        rotation = entry.get("rotation")
        words_raw = entry.get("words", [])
        lines_raw = entry.get("lines", [])
        words = [
            OCRWord(text=str(word["text"]), bbox=_decode_bbox(word["bbox"]))
            for word in words_raw
            if "text" in word and "bbox" in word
        ]
        lines = []
        for line in lines_raw:
            if "text" not in line or "bbox" not in line:
                continue
            line_words = [
                OCRWord(text=str(word["text"]), bbox=_decode_bbox(word["bbox"]))
                for word in line.get("words", [])
                if "text" in word and "bbox" in word
            ]
            lines.append(
                OCRLine(
                    text=str(line["text"]),
                    bbox=_decode_bbox(line["bbox"]),
                    words=line_words,
                )
            )

        if not lines and words:
            lines = _lines_from_words(words)
        pages.append(
            OCRPage(
                index=index,
                width_px=width_px,
                height_px=height_px,
                rotation=int(rotation) if rotation is not None else None,
                words=words,
                lines=lines,
            )
        )
    return pages


def save_ocr_json(pages: Sequence[OCRPage], path: Path) -> None:
    """Serialize OCR pages to JSON for reuse."""
    data: List[Dict[str, Any]] = []
    for page in pages:
        data.append(
            {
                "page": page.index,
                "width_px": page.width_px,
                "height_px": page.height_px,
                "rotation": page.rotation,
                "words": [
                    {"text": word.text, "bbox": list(map(float, word.bbox))}
                    for word in page.words
                ],
                "lines": [
                    {
                        "text": line.text,
                        "bbox": list(map(float, line.bbox)),
                        "words": [
                            {"text": word.text, "bbox": list(map(float, word.bbox))}
                            for word in line.words
                        ],
                    }
                    for line in page.lines
                ],
            }
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("Dumped OCR JSON to %s", path)


def _lines_from_words(words: Sequence[OCRWord], tolerance: float = 4.0) -> List[OCRLine]:
    """Group words into lines based on Y proximity."""
    sorted_words = sorted(words, key=lambda w: (w.bbox[1], w.bbox[0]))
    lines: List[OCRLine] = []
    current: Optional[OCRLine] = None
    for word in sorted_words:
        if current is None:
            current = OCRLine(text=word.text, bbox=word.bbox, words=[word])
            continue
        _, current_y, _, current_h = current.bbox
        x, y, w, h = word.bbox
        if abs(y - current_y) <= tolerance:
            current.text += " " + word.text
            current.words.append(word)
            x0 = min(current.bbox[0], x)
            y0 = min(current.bbox[1], y)
            x1 = max(current.bbox[0] + current.bbox[2], x + w)
            y1 = max(current.bbox[1] + current.bbox[3], y + h)
            current.bbox = (x0, y0, x1 - x0, y1 - y0)
        else:
            lines.append(current)
            current = OCRLine(text=word.text, bbox=word.bbox, words=[word])
    if current:
        lines.append(current)
    return lines


def winrt_available() -> bool:
    """Return whether WinRT OCR can be used on this system."""
    if os.name != "nt":
        return False
    if os.environ.get("WSL_DISTRO_NAME") or "microsoft" in platform.release().lower():
        return False
    if sys_version_tuple() != (3, 11):
        return False
    if struct.calcsize("P") * 8 != 64:
        return False
    try:
        import winrt.windows.media.ocr  # type: ignore  # noqa: F401
    except ImportError:
        return False
    return True


def sys_version_tuple() -> Tuple[int, int]:
    return (sys.version_info.major, sys.version_info.minor)


async def _pil_to_software_bitmap(image: Image.Image) -> Any:
    from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter  # type: ignore
    from winrt.windows.graphics.imaging import BitmapDecoder  # type: ignore

    buffer = BytesIO()
    image.save(buffer, format="PNG")
    data = buffer.getvalue()

    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(data)
    await writer.store_async()
    await writer.flush_async()
    writer.detach_stream()
    writer.close()
    stream.seek(0)

    decoder = await BitmapDecoder.create_async(stream)
    return await decoder.get_software_bitmap_async()


async def _recognize_bitmap(engine: Any, bitmap: Any) -> Any:
    try:
        result = await engine.recognize_async(bitmap)
        return result
    except Exception:
        from winrt.windows.graphics.imaging import BitmapPixelFormat, SoftwareBitmap  # type: ignore

        converted = SoftwareBitmap.convert(bitmap, BitmapPixelFormat.GRAY8)
        return await engine.recognize_async(converted)


def run_winrt_ocr(
    pdf_path: Path,
    dpi: int,
    language: str,
    max_pages: Optional[int] = None,
) -> List[OCRPage]:
    """Render PDF pages and run WinRT OCR."""
    if fitz is None:
        raise WinRTUnavailableError("PyMuPDF (fitz) is required to render PDF pages for OCR.")
    if Image is None:
        raise WinRTUnavailableError("Pillow is required to convert PDF renders for OCR.")
    if not winrt_available():
        raise WinRTUnavailableError(
            "WinRT OCR is unavailable. Use --ocr-json or install 64-bit Python 3.11 on Windows with winrt packages."
        )

    from winrt.windows.media.ocr import OcrEngine  # type: ignore
    from winrt.windows.globalization import Language  # type: ignore

    doc = fitz.open(pdf_path)
    scale = dpi / 72.0
    pages: List[OCRPage] = []

    try:
        engine = None
        try:
            engine = OcrEngine.try_create_from_language(Language(language))
        except Exception:
            engine = None
        if engine is None:
            engine = OcrEngine.try_create_from_user_profile_languages()
        if engine is None:
            raise WinRTUnavailableError(
                "Windows OCR language not available. Install the language pack via Windows Settings."
            )

        total_pages = len(doc)
        count = min(total_pages, max_pages) if max_pages else total_pages
        for index in range(count):
            page = doc.load_page(index)
            matrix = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            bitmap = asyncio.run(_pil_to_software_bitmap(image))
            result = asyncio.run(_recognize_bitmap(engine, bitmap))

            words: List[OCRWord] = []
            lines: List[OCRLine] = []
            for line in result.lines:
                line_words: List[OCRWord] = []
                for word in line.words:
                    rect = _extract_rect(word)
                    if rect is None:
                        continue
                    bbox = (rect.x, rect.y, rect.width, rect.height)
                    word_obj = OCRWord(text=word.text, bbox=bbox)
                    words.append(word_obj)
                    line_words.append(word_obj)

                line_rect_obj = _extract_rect(line)
                if line_rect_obj is None and line_words:
                    xs = [w.bbox[0] for w in line_words]
                    ys = [w.bbox[1] for w in line_words]
                    xe = [w.bbox[0] + w.bbox[2] for w in line_words]
                    ye = [w.bbox[1] + w.bbox[3] for w in line_words]
                    line_bbox = (
                        min(xs),
                        min(ys),
                        max(xe) - min(xs),
                        max(ye) - min(ys),
                    )
                elif line_rect_obj is not None:
                    line_bbox = (
                        line_rect_obj.x,
                        line_rect_obj.y,
                        line_rect_obj.width,
                        line_rect_obj.height,
                    )
                else:
                    line_bbox = (0.0, 0.0, 0.0, 0.0)

                lines.append(
                    OCRLine(
                        text=line.text,
                        bbox=line_bbox,
                        words=line_words,
                    )
                )

            pages.append(
                OCRPage(
                    index=index,
                    width_px=pix.width,
                    height_px=pix.height,
                    rotation=page.rotation,
                    words=words,
                    lines=lines,
                )
            )
    finally:
        doc.close()
    return pages
