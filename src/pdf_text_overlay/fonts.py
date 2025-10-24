"""Font resolution and registration utilities."""

from __future__ import annotations

import logging
import os
from importlib import resources
from pathlib import Path
from typing import Iterable, Optional, TYPE_CHECKING

try:
    import fitz  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    fitz = None  # type: ignore

logger = logging.getLogger(__name__)

if TYPE_CHECKING:  # pragma: no cover - typing only
    import fitz  # type: ignore

_WINDOWS_FONT_DIRS = [
    Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
    Path(os.environ.get("LOCALAPPDATA", r"C:\Users\Public")) / "Microsoft\Windows\Fonts",
]


def _iter_packaged_fonts() -> Iterable[Path]:
    try:
        font_root = resources.files("pdf_text_overlay.resources") / "fonts"
    except (ModuleNotFoundError, AttributeError):
        return
    for entry in font_root.iterdir():
        if entry.name.lower().endswith((".ttf", ".otf", ".ttc")):
            with resources.as_file(entry) as path:
                yield path


def _iter_system_fonts() -> Iterable[Path]:
    for directory in _WINDOWS_FONT_DIRS:
        if not directory.exists():
            continue
        for candidate in directory.glob("*.tt*"):
            yield candidate
    linux_candidates = [
        Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
        Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.otf"),
    ]
    for candidate in linux_candidates:
        if candidate.exists():
            yield candidate


def resolve_font(font_path: Optional[Path]) -> Path:
    """Resolve font path from CLI or fallbacks."""
    if font_path:
        candidate = font_path.expanduser()
        if not candidate.is_file():
            raise FileNotFoundError(f"Font file not found: {candidate}")
        return candidate

    for candidate in _iter_packaged_fonts():
        logger.debug("Found packaged font: %s", candidate)
        return candidate

    for candidate in _iter_system_fonts():
        logger.debug("Found system font: %s", candidate)
        return candidate

    raise FileNotFoundError(
        "No usable font found. Provide a font path via --font pointing to a Unicode TrueType/OpenType font."
    )


def register_font(doc: fitz.Document, font_path: Optional[Path]) -> str:
    """Register font with fitz document and return its internal name."""
    if fitz is None:
        raise RuntimeError("PyMuPDF (fitz) is required to register fonts.")
    try:
        resolved = resolve_font(font_path)
    except FileNotFoundError as exc:
        logger.warning("%s Falling back to built-in Helvetica font.", exc)
        return "helv"

    try:
        font_name = doc.insert_font(fontfile=str(resolved))
        logger.debug("Registered font %s as %s", resolved, font_name)
        return font_name
    except Exception as exc:  # pragma: no cover - dependent on font availability
        logger.warning("Failed to load font %s (%s). Falling back to Helvetica.", resolved, exc)
        return "helv"
