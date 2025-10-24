"""Command line interface for pdf_text_overlay."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional, Tuple

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


def _parse_pair(value: str, ctx: click.Context, param: click.Parameter) -> Tuple[float, float]:
    try:
        x_str, y_str = value.split(",")
        return float(x_str), float(y_str)
    except Exception as exc:  # pragma: no cover - click handles errors
        raise click.BadParameter("Expected two comma-separated numbers, e.g. 0.0,0.0") from exc


def _parse_scale(value: str, ctx: click.Context, param: click.Parameter) -> Tuple[float, float]:
    sx, sy = _parse_pair(value, ctx, param)
    if sx == 0 or sy == 0:
        raise click.BadParameter("Scale corrections must be non-zero.")
    return sx, sy


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
@click.option("--font-scale", default=1.0, show_default=True, type=click.FloatRange(0.1, 3.0), help="Scale factor applied to font size relative to bbox height.")
@click.option("--align", default="auto", show_default=True, help="Alignment mode: auto, page, image:<xref>, or image-rect:x0,y0,x1,y1.")
@click.option("--offset-pt", default="0,0", callback=_parse_pair, help="Global offset in points applied after mapping (dx,dy).")
@click.option("--scale-corr", default="1,1", callback=_parse_scale, help="Scale correction multipliers (sx,sy).")
@click.option("--rotate", type=click.Choice(["0", "90", "180", "270"]), default=None, help="Override page rotation (degrees).")
@click.option("--deskew", default=0.0, show_default=True, type=click.FloatRange(-5.0, 5.0), help="Additional rotation applied to text placement.")
@click.option("--debug-overlay", is_flag=True, help="Draw translucent bounding boxes for inspection.")
@click.option("--visible-qa", is_flag=True, help="Render text visible in light gray for QA (overrides --method).")
@click.option("--keep-spaces", is_flag=True, help="Keep original whitespace instead of collapsing.")
@click.option("--dehyphen", is_flag=True, help="Recombine hyphenated line endings.")
@click.option("--cjk-join", is_flag=True, help="Remove spaces between consecutive CJK characters.")
@click.option("--pdfa", is_flag=True, help="Attempt to export as PDF/A-2b.")
@click.option("--dump-ocr-json", type=click.Path(dir_okay=False, path_type=Path), default=None, help="Write OCR JSON to this path after WinRT OCR.")
@click.option("--dump-debug-json", type=click.Path(dir_okay=False, path_type=Path), default=None, help="Write detailed placement diagnostics to JSON.")
@click.option("--calibrate", default=0, show_default=True, type=click.IntRange(0, 1000), help="Log first N word placements for calibration.")
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
    font_scale: float,
    align: str,
    offset_pt: Tuple[float, float],
    scale_corr: Tuple[float, float],
    rotate: Optional[str],
    deskew: float,
    debug_overlay: bool,
    visible_qa: bool,
    keep_spaces: bool,
    dehyphen: bool,
    cjk_join: bool,
    pdfa: bool,
    dump_ocr_json: Optional[Path],
    dump_debug_json: Optional[Path],
    calibrate: int,
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

    rotate_override = int(rotate) if rotate is not None else None

    settings = OverlaySettings(
        method=method,
        granularity=granularity,
        baseline_ratio=baseline_ratio,
        font_scale=font_scale,
        keep_spaces=keep_spaces,
        dehyphen=dehyphen,
        cjk_join=cjk_join,
        debug_overlay=debug_overlay,
        visible_qa=visible_qa,
        pdfa=pdfa,
        font_path=font_path,
        dpi=dpi,
        align=align,
        offset_pt=offset_pt,
        scale_corr=scale_corr,
        rotate_override=rotate_override,
        deskew=deskew,
        calibrate=calibrate,
        dump_debug_json=dump_debug_json,
    )

    apply_text_overlay(pdf_path=pdf_path, pages=pages, output_path=out_path, settings=settings)
    logging.info("Completed overlay.")
