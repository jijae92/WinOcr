# REQUIREMENTS.txt
# pip install -r requirements.txt
"""Windows PDF to text/markdown extractor using WinRT OCR."""

from __future__ import annotations

import argparse
import asyncio
import importlib
import json
import logging
import os
import platform
import struct
import sys
import time
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING, Any, Awaitable, Dict, Iterable, List, Optional, Sequence, Tuple, TypeVar

try:
    import fitz  # type: ignore
except ImportError as exc:  # pragma: no cover - import guard
    fitz = None  # type: ignore
    FITZ_IMPORT_ERROR = exc
else:
    FITZ_IMPORT_ERROR = None

try:
    from PIL import Image
except ImportError as exc:  # pragma: no cover - import guard
    Image = None  # type: ignore
    PIL_IMPORT_ERROR = exc
else:
    PIL_IMPORT_ERROR = None

if TYPE_CHECKING:  # pragma: no cover - typing only
    from winrt.windows.graphics.imaging import SoftwareBitmap  # type: ignore
    from winrt.windows.media.ocr import OcrResult as WinRTOcrResult  # type: ignore
else:
    SoftwareBitmap = Any
    WinRTOcrResult = Any


logger = logging.getLogger("winocr_pdf")


class OCRToolError(Exception):
    """Custom exception for predictable CLI exits."""

    def __init__(self, message: str, exit_code: int = 1) -> None:
        super().__init__(message)
        self.exit_code = exit_code


@dataclass
class PageOCR:
    """Structured OCR output for a single page."""

    index: int
    width: int
    height: int
    lines: List[Dict[str, Any]]
    plain_text: str


OcrEngine: Any = None
Language: Any = None
InMemoryRandomAccessStream: Any = None
DataWriter: Any = None
BitmapDecoder: Any = None
SoftwareBitmap_cls: Any = None
BitmapPixelFormat: Any = None
BitmapAlphaMode: Any = None

_WINRT_READY = False
_WINRT_MODULES: Tuple[Tuple[str, str], ...] = (
    ("winrt.windows.foundation", "winrt-Windows.Foundation"),
    ("winrt.windows.foundation.collections", "winrt-Windows.Foundation.Collections"),
    ("winrt.windows.media.ocr", "winrt-Windows.Media.Ocr"),
    ("winrt.windows.globalization", "winrt-Windows.Globalization"),
    ("winrt.windows.graphics.imaging", "winrt-Windows.Graphics.Imaging"),
    ("winrt.windows.storage", "winrt-Windows.Storage"),
    ("winrt.windows.storage.streams", "winrt-Windows.Storage.Streams"),
)
_WINRT_INSTALL_CMD = (
    "python -m pip install -U winrt-runtime "
    "winrt-Windows.Foundation winrt-Windows.Foundation.Collections "
    "winrt-Windows.Media.Ocr winrt-Windows.Globalization "
    "winrt-Windows.Graphics.Imaging winrt-Windows.Storage "
    "winrt-Windows.Storage.Streams"
)

_LANGUAGE_MAP = {
    "ko-kr": "ko",
    "ko": "ko",
    "en-us": "en",
    "en": "en",
    "ja-jp": "ja",
    "ja": "ja",
    "zh-cn": "zh-Hans",
    "zh-hans": "zh-Hans",
    "zh-tw": "zh-Hant",
    "zh-hant": "zh-Hant",
}

_FORMAT_ALIASES = {
    "text": "text",
    "txt": "text",
    "plain": "text",
    "md": "md",
    "markdown": "md",
    "both": "both",
}


T = TypeVar("T")


def _ensure_winrt() -> None:
    """Ensure required WinRT projections are importable and cached."""

    global _WINRT_READY
    global OcrEngine, Language
    global InMemoryRandomAccessStream, DataWriter
    global BitmapDecoder, SoftwareBitmap_cls, BitmapPixelFormat, BitmapAlphaMode
    global WinRTOcrResult

    if _WINRT_READY:
        return

    missing: List[str] = []
    for module_name, package_name in _WINRT_MODULES:
        try:
            importlib.import_module(module_name)
        except ModuleNotFoundError:
            missing.append(f"{module_name} (pip: {package_name})")

    if missing:
        missing_text = "\n".join(f"  - {item}" for item in missing)
        raise OCRToolError(
            (
                "Missing WinRT projection modules:\n"
                f"{missing_text}\n"
                "Install/update them with:\n"
                f"  {_WINRT_INSTALL_CMD}"
            ),
            exit_code=2,
        )

    media_ocr = importlib.import_module("winrt.windows.media.ocr")
    globalization = importlib.import_module("winrt.windows.globalization")
    streams = importlib.import_module("winrt.windows.storage.streams")
    imaging = importlib.import_module("winrt.windows.graphics.imaging")

    # The following imports validate availability of supporting namespaces.
    importlib.import_module("winrt.windows.foundation")
    importlib.import_module("winrt.windows.foundation.collections")
    importlib.import_module("winrt.windows.storage")

    OcrEngine = media_ocr.OcrEngine
    WinRTOcrResult = media_ocr.OcrResult
    Language = globalization.Language
    InMemoryRandomAccessStream = streams.InMemoryRandomAccessStream
    DataWriter = streams.DataWriter
    BitmapDecoder = imaging.BitmapDecoder
    SoftwareBitmap_cls = imaging.SoftwareBitmap
    BitmapPixelFormat = imaging.BitmapPixelFormat
    BitmapAlphaMode = imaging.BitmapAlphaMode

    _WINRT_READY = True


def configure_logging(verbose: bool) -> None:
    """Configure root logger once."""

    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def _parse_output_format(value: str) -> str:
    key = value.strip().lower()
    if key in _FORMAT_ALIASES:
        return _FORMAT_ALIASES[key]
    raise argparse.ArgumentTypeError("--fmt must be one of: text, md, both (aliases: txt, markdown).")


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    """Parse CLI arguments."""

    parser = argparse.ArgumentParser(
        description="Convert a PDF into images and extract text via Windows OCR.",
    )
    parser.add_argument(
        "-i",
        "--input",
        required=True,
        help="Path to input PDF file (absolute or relative).",
    )
    parser.add_argument(
        "-o",
        "--outdir",
        default="./out",
        help="Output directory (default: ./out).",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="Rendering DPI for PDF pages (default: 300).",
    )
    parser.add_argument(
        "--lang",
        default="ko-KR",
        help="OCR language tag (BCP-47, e.g. ko-KR, en-US).",
    )
    parser.add_argument(
        "--fmt",
        type=_parse_output_format,
        default="both",
        help="Output format: text, md, or both (aliases: txt, markdown).",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=None,
        help="Limit the number of pages to process (default: all).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Execute OCR without writing output files.",
    )
    parser.add_argument(
        "--dump-pages",
        action="store_true",
        help="Dump rendered page images (.png) and layout JSON alongside results.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose debug logging.",
    )

    args = parser.parse_args(argv)

    if args.dpi <= 0:
        parser.error("--dpi must be a positive integer.")
    if args.max_pages is not None and args.max_pages <= 0:
        parser.error("--max-pages must be a positive integer when provided.")
    return args


def ensure_environment() -> None:
    """Validate runtime environment and core dependencies."""

    if os.name != "nt":
        raise OCRToolError(
            "This tool runs on Windows 10/11 only.",
            exit_code=2,
        )

    if os.environ.get("WSL_DISTRO_NAME") or "microsoft" in platform.release().lower():
        raise OCRToolError(
            "WinRT OCR is not supported inside WSL. Please run from Windows PowerShell or CMD.",
            exit_code=2,
        )

    if sys.version_info[:2] != (3, 11):
        raise OCRToolError(
            "Python 3.11 (64-bit) is required. Install Python 3.11.x from python.org and retry.",
            exit_code=2,
        )

    if struct.calcsize("P") * 8 != 64:
        raise OCRToolError(
            "A 64-bit Python installation is required for WinRT.",
            exit_code=2,
        )

    if FITZ_IMPORT_ERROR is not None:
        raise OCRToolError(
            "Missing dependency: PyMuPDF (pymupdf). Install with 'pip install pymupdf'.",
            exit_code=2,
        )
    if PIL_IMPORT_ERROR is not None:
        raise OCRToolError(
            "Missing dependency: Pillow. Install with 'pip install pillow'.",
            exit_code=2,
        )

    _ensure_winrt()


def resolve_paths(input_path: str, outdir: str, create_output_dir: bool) -> Tuple[Path, Path]:
    """Resolve filesystem paths with Windows-friendly handling."""

    pdf_path = Path(input_path).expanduser()
    if not pdf_path.is_file():
        raise OCRToolError(
            f"Input PDF not found: {pdf_path}. Verify the path and ensure the file exists.",
            exit_code=2,
        )
    output_dir = Path(outdir).expanduser()
    if create_output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)
    elif not output_dir.exists():
        logger.warning("Output directory %s does not exist (dry-run mode).", output_dir)
    return pdf_path, output_dir


def _normalize_language_tag(language_code: str) -> str:
    canonical = language_code.strip().replace("_", "-")
    key = canonical.lower()
    if key in _LANGUAGE_MAP:
        return _LANGUAGE_MAP[key]
    if "-" in canonical:
        base = canonical.split("-")[0]
        return base.lower()
    return canonical.lower()


def create_ocr_engine(language_code: str) -> Tuple[Any, Optional[str], bool]:
    """Initialize the Windows OCR engine with language fallback."""

    _ensure_winrt()

    assert OcrEngine is not None
    assert Language is not None

    normalized_code = _normalize_language_tag(language_code)
    requested_language = None
    fallback_used = False

    try:
        requested_language = Language(normalized_code)
    except Exception as exc:  # pragma: no cover - defensive fallback
        logger.warning(
            "Invalid OCR language code '%s'; falling back to user profile languages (%s).",
            language_code,
            exc,
        )
    else:
        engine = OcrEngine.try_create_from_language(requested_language)
        if engine:
            resolved = (
                engine.recognizer_language.language_tag
                if engine.recognizer_language
                else requested_language.language_tag
            )
            return engine, resolved, False
        fallback_used = True
        logger.warning(
            "WinRT OCR language '%s' not available; falling back to user profile languages.",
            normalized_code,
        )

    engine = OcrEngine.try_create_from_user_profile_languages()
    if engine:
        resolved = (
            engine.recognizer_language.language_tag
            if engine.recognizer_language
            else None
        )
        return engine, resolved, fallback_used or requested_language is None

    raise OCRToolError(
        (
            "Windows OCR language not available. "
            "Windows 설정 → 시간 및 언어 → 언어 및 지역 → 해당 언어 추가 → 텍스트 인식(OCR) 설치"
        ),
        exit_code=3,
    )


async def _pil_image_to_software_bitmap(image: Image.Image) -> SoftwareBitmap:
    """Convert a PIL Image into a SoftwareBitmap suitable for WinRT OCR."""

    _ensure_winrt()

    buffer = BytesIO()
    image.save(buffer, format="PNG")
    data = buffer.getvalue()
    logger.debug("Saved PIL image to PNG buffer (%d bytes).", len(data))

    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(data)
    await writer.store_async()
    await writer.flush_async()
    writer.detach_stream()
    writer.close()
    stream.seek(0)

    decoder = await BitmapDecoder.create_async(stream)
    sbmp = await decoder.get_software_bitmap_async()
    logger.debug(
        "Decoded SoftwareBitmap: %dx%d px, format=%s",
        sbmp.pixel_width,
        sbmp.pixel_height,
        sbmp.bitmap_pixel_format,
    )
    if sbmp.bitmap_pixel_format != BitmapPixelFormat.GRAY8:
        logger.debug("Converting SoftwareBitmap to GRAY8 format for OCR compatibility.")
        try:
            sbmp = SoftwareBitmap_cls.convert(sbmp, BitmapPixelFormat.GRAY8, BitmapAlphaMode.IGNORE)
        except TypeError:
            sbmp = SoftwareBitmap_cls.convert(sbmp, BitmapPixelFormat.GRAY8)
    return sbmp


async def _recognize_page_async(engine: Any, image: Image.Image) -> WinRTOcrResult:
    """Run WinRT OCR on a PIL image."""

    sbmp = await _pil_image_to_software_bitmap(image)
    try:
        result = await engine.recognize_async(sbmp)
        return result
    except Exception as exc:
        logger.debug(
            "Initial OCR attempt failed (%s); retrying after GRAY8 conversion fallback.",
            exc,
        )
        sbmp_gray = SoftwareBitmap_cls.convert(sbmp, BitmapPixelFormat.GRAY8)
        result = await engine.recognize_async(sbmp_gray)
        return result


def run_async(coro: Awaitable[T]) -> T:
    """Run an async coroutine synchronously."""

    return asyncio.run(coro)


def render_page_to_image(page: "fitz.Page", dpi: int) -> Image.Image:
    """Render a PDF page to a PIL image."""

    assert Image is not None

    scale = dpi / 72.0
    matrix = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    logger.debug(
        "Rendered page %d to image: %dx%d px, DPI=%d.",
        page.number,
        image.width,
        image.height,
        dpi,
    )
    return image


def extract_line_data(ocr_result: WinRTOcrResult) -> Tuple[str, List[Dict[str, Any]]]:
    """Convert an OCR result into plain text and structured line data."""

    lines: List[Dict[str, Any]] = []
    for line in ocr_result.lines:
        words: List[Dict[str, Any]] = []
        for word in line.words:
            rect = word.bounding_rect
            words.append(
                {
                    "text": word.text,
                    "bbox": [rect.x, rect.y, rect.width, rect.height],
                }
            )
        lines.append({"text": line.text, "words": words})
    plain_text = ocr_result.text.replace("\r\n", "\n").rstrip()
    return plain_text, lines


def build_text_output(pages: Iterable[PageOCR]) -> str:
    """Build plain text output with page delimiters."""

    parts: List[str] = []
    for page in pages:
        parts.append(f"===== Page {page.index + 1} =====")
        parts.append(page.plain_text or "")
    return "\n".join(parts).rstrip() + "\n"


def build_markdown_output(pages: Iterable[PageOCR]) -> str:
    """Build Markdown output with page headings."""

    blocks: List[str] = []
    for page in pages:
        blocks.append(f"## Page {page.index + 1}")
        content = page.plain_text.strip() if page.plain_text else "_No text recognized._"
        blocks.append(content)
    return "\n\n".join(blocks).rstrip() + "\n"


def build_layout_payload(pages: Iterable[PageOCR], dpi: int, input_file: Path, resolved_language: Optional[str]) -> Dict[str, Any]:
    """Create JSON structure for OCR layout results."""

    return {
        "file": str(input_file),
        "dpi": dpi,
        "lang": resolved_language,
        "pages": [
            {
                "index": page.index,
                "width": page.width,
                "height": page.height,
                "lines": page.lines,
            }
            for page in pages
        ],
    }


def process_pdf(
    pdf_path: Path,
    output_dir: Path,
    dpi: int,
    max_pages: Optional[int],
    language_code: str,
    dump_pages: bool,
    dry_run: bool,
) -> Tuple[List[PageOCR], Optional[str]]:
    """Render each page of the PDF and perform OCR."""

    engine, resolved_lang, fallback_used = create_ocr_engine(language_code)
    logger.info(
        "Initialized OCR engine with language=%s (fallback used=%s).",
        resolved_lang or _normalize_language_tag(language_code),
        fallback_used,
    )

    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        raise OCRToolError(f"Failed to open PDF: {exc}", exit_code=2) from exc

    try:
        total_pages = doc.page_count
        if total_pages == 0:
            raise OCRToolError("PDF contains no pages to process.", exit_code=4)

        target_count = min(total_pages, max_pages) if max_pages else total_pages
        logger.info(
            "Processing %d page(s) (of %d total) at %d DPI.",
            target_count,
            total_pages,
            dpi,
        )
        start_time = time.perf_counter()
        pages: List[PageOCR] = []

        for index in range(target_count):
            page = doc.load_page(index)
            page_start = time.perf_counter()
            image = render_page_to_image(page, dpi)
            ocr_result = run_async(_recognize_page_async(engine, image))
            plain_text, line_data = extract_line_data(ocr_result)
            pages.append(
                PageOCR(
                    index=index,
                    width=image.width,
                    height=image.height,
                    lines=line_data,
                    plain_text=plain_text,
                )
            )
            page_elapsed = time.perf_counter() - page_start
            logger.info("Page %d processed in %.2f seconds.", index + 1, page_elapsed)

            if dump_pages and not dry_run:
                image_path = output_dir / f"{pdf_path.stem}_page{index + 1:04d}.png"
                image.save(image_path, format="PNG")
                logger.debug("Dumped rendered page to %s.", image_path)
            elif dump_pages and dry_run:
                logger.debug(
                    "Dry-run: skipping page image dump for page %d (path would be %s).",
                    index + 1,
                    output_dir / f"{pdf_path.stem}_page{index + 1:04d}.png",
                )

        total_elapsed = time.perf_counter() - start_time
        logger.info("Completed OCR for %d page(s) in %.2f seconds.", target_count, total_elapsed)
        return pages, resolved_lang
    finally:
        doc.close()


def write_outputs(
    pdf_path: Path,
    output_dir: Path,
    pages: List[PageOCR],
    resolved_language: Optional[str],
    dpi: int,
    output_format: str,
    dump_pages: bool,
) -> None:
    """Persist outputs in the requested formats."""

    base_name = pdf_path.stem
    if output_format in {"text", "both"}:
        txt_path = output_dir / f"{base_name}.txt"
        text_content = build_text_output(pages)
        txt_path.write_text(text_content, encoding="utf-8")
        logger.info("Wrote text output to %s.", txt_path)

    if output_format in {"md", "both"}:
        md_path = output_dir / f"{base_name}.md"
        md_content = build_markdown_output(pages)
        md_path.write_text(md_content, encoding="utf-8")
        logger.info("Wrote Markdown output to %s.", md_path)

    if dump_pages:
        layout_path = output_dir / f"{base_name}_layout.json"
        layout_payload = build_layout_payload(pages, dpi, pdf_path, resolved_language)
        layout_path.write_text(json.dumps(layout_payload, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info("Wrote layout JSON to %s.", layout_path)


def main(argv: Optional[Sequence[str]] = None) -> int:
    """CLI entry point."""

    args = parse_args(argv)
    configure_logging(args.verbose)
    try:
        ensure_environment()
        pdf_path, output_dir = resolve_paths(args.input, args.outdir, create_output_dir=not args.dry_run)
        pages, resolved_lang = process_pdf(
            pdf_path=pdf_path,
            output_dir=output_dir,
            dpi=args.dpi,
            max_pages=args.max_pages,
            language_code=args.lang,
            dump_pages=args.dump_pages,
            dry_run=args.dry_run,
        )
        if args.dry_run:
            logger.info("Dry-run mode enabled; skipping write of output files.")
        else:
            write_outputs(
                pdf_path=pdf_path,
                output_dir=output_dir,
                pages=pages,
                resolved_language=resolved_lang,
                dpi=args.dpi,
                output_format=args.fmt,
                dump_pages=args.dump_pages,
            )
    except OCRToolError as exc:
        logger.error("%s", exc)
        return exc.exit_code
    except Exception as exc:  # pragma: no cover - unexpected failure
        logger.exception("Unexpected error: %s", exc)
        return 1
    logger.info("OCR extraction finished successfully.")
    return 0


if __name__ == "__main__":
    if len(sys.argv) == 1:
        print("Sample usage:")
        print(r'  python winocr_pdf.py -i "C:\\path\\doc.pdf" --lang ko-KR --dpi 300 -o .\\out --fmt both')
        print(r'  python winocr_pdf.py -i "C:\\path\\doc.pdf" --lang en-US --fmt text --dump-pages')
    sys.exit(main())
