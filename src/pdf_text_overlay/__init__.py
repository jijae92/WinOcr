"""Create searchable PDFs by overlaying invisible OCR text layers."""

__version__ = "0.1.0"

__all__ = ["__version__", "main"]


def main() -> None:
    """Console entry point proxy for `python -m pdf_text_overlay`."""
    from .cli import main as _main

    _main()
