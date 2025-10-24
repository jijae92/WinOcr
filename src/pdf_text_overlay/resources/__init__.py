"""Package data shipped with pdf_text_overlay."""

from importlib import resources as _resources

__all__ = ["open_text", "files"]

open_text = _resources.open_text
files = _resources.files
