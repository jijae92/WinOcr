"""Command line interface for pdf_text_overlay."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

import click

from . import __version__
from .overlay import OverlaySettings, apply_text_overlay
from .ocr_io import (
    OCRPage,
    WinRTUnavailableError,
    load_ocr_json,
    run_winrt_ocr,
    save_ocr_json,
    winrt_available,
)

LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"


def _setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format=LOG_FORMAT)


def _version_callback(ctx: click.Context, param: click.Option, value: bool) -> None:
    if not value or ctx.resilient_parsing:
        return
    click.echo(f"pdf_text_overlay {__version__}")
    ctx.exit()


def _load_or_run_ocr(
    pdf_path: Path,
    ocr_json: Optional[Path],
    dpi: int,
    lang: str,
    max_pages: Optional[int],
    dump_ocr_json: Optional[Path],
) -> list[OCRPage]:
    if ocr_json:
        pages = load_ocr_json(ocr_json)
        logging.info("Loaded OCR JSON with %d page entries.", len(pages))
    else:
        if not winrt_available():
            raise click.UsageError(
                "WinRT OCR is unavailable on this system. Provide --ocr-json or run on Windows 10/11 "
                "with 64-bit Python 3.11 and WinRT packages installed."
            )
        logging.info("Running WinRT OCR (dpi=%d, lang=%s)...", dpi, lang)
        try:
            pages = run_winrt_ocr(pdf_path, dpi=dpi, language=lang, max_pages=max_pages)
        except WinRTUnavailableError as exc:
            raise click.UsageError(str(exc)) from exc
        if dump_ocr_json:
            save_ocr_json(pages, dump_ocr_json)
    return pages


@click.command(
    context_settings={"help_option_names": ["-h", "--help"]},
    help="Overlay invisible text onto image-only PDFs to make them searchable.",
)
@click.option(
    "--version",
    "show_version",
    is_flag=True,
    callback=_version_callback,
    expose_value=False,
    is_eager=True,
    help="Show the pdf_text_overlay version and exit.",
)
@click.option("--pdf", "pdf_path", type=click.Path(exists=True, dir_okay=False, path_type=Path), required=True, help="Input PDF file.")
@click.option("--ocr-json", "ocr_json", type=click.Path(exists=True, dir_okay=False, path_type=Path), default=None, help="Precomputed OCR JSON path.")
@click.option("--out", "out_path", type=click.Path(dir_okay=False, path_type=Path), required=True, help="Output searchable PDF path.")
@click.option("--dpi", default=300, show_default=True, type=click.IntRange(72, 1200), help="Render DPI when running WinRT OCR.")
@click.option("--lang", default="ko-KR", show_default=True, help="OCR language tag for WinRT OCR.")
@click.option("--granularity", default="word", type=click.Choice(["word", "line"]), show_default=True, help="Overlay granularity.")
@click.option("--method", default="invisible", type=click.Choice(["invisible", "opacity"]), show_default=True, help="Rendering method for hidden text.")
@click.option("--font", "font_path", type=click.Path(exists=True, dir_okay=False, path_type=Path), default=None, help="Font file to embed for overlay text.")
@click.option("--baseline-ratio", default=0.15, show_default=True, type=click.FloatRange(0.0, 1.0), help="Baseline ratio within bounding box.")
@click.option("--debug-overlay", is_flag=True, help="Draw translucent bounding boxes for inspection.")
@click.option("--keep-spaces", is_flag=True, help="Keep original whitespace instead of collapsing.")
@click.option("--dehyphen", is_flag=True, help="Recombine hyphenated line endings.")
@click.option("--pdfa", is_flag=True, help="Attempt to export as PDF/A-2b.")
@click.option("--dump-ocr-json", type=click.Path(dir_okay=False, path_type=Path), default=None, help="Write OCR JSON to this path after WinRT OCR.")
@click.option("--max-pages", type=click.IntRange(1, 5000), default=None, help="Limit number of pages processed.")
@click.option("--verbose", is_flag=True, help="Enable verbose logging.")
def main(
    pdf_path: Path,
    ocr_json: Optional[Path],
    out_path: Path,
    dpi: int,
    lang: str,
    granularity: str,
    method: str,
    font_path: Optional[Path],
    baseline_ratio: float,
    debug_overlay: bool,
    keep_spaces: bool,
    dehyphen: bool,
    pdfa: bool,
    dump_ocr_json: Optional[Path],
    max_pages: Optional[int],
    verbose: bool,
) -> None:
    """Create a searchable PDF by overlaying OCR text."""

    _setup_logging(verbose)
    logging.info("Starting pdf_text_overlay for %s", pdf_path)

    pages = _load_or_run_ocr(
        pdf_path=pdf_path,
        ocr_json=ocr_json,
        dpi=dpi,
        lang=lang,
        max_pages=max_pages,
        dump_ocr_json=dump_ocr_json,
    )

    if not pages:
        raise click.ClickException("No OCR data available to overlay.")

    if baseline_ratio <= 0.0 or baseline_ratio >= 1.0:
        logging.warning("Baseline ratio %.3f is extreme; consider staying within 0.1-0.3.", baseline_ratio)

    settings = OverlaySettings(
        method=method,
        granularity=granularity,
        baseline_ratio=baseline_ratio,
        keep_spaces=keep_spaces,
        dehyphen=dehyphen,
        debug_overlay=debug_overlay,
        pdfa=pdfa,
        font_path=font_path,
        dpi=dpi,
    )

    apply_text_overlay(pdf_path=pdf_path, pages=pages, output_path=out_path, settings=settings)
    logging.info("Completed overlay.")
