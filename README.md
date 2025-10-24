# WinOCR PDF CLI

Windows-only PDF to text/markdown converter that renders pages with PyMuPDF and extracts text using the built-in Windows OCR engine (WinRT: `Windows.Media.Ocr`).

## Prerequisites
- Windows 10 or 11.
- 64-bit Python 3.11 (`py -3.11`).
- Run from PowerShell/CMD (not WSL).
- Matching language pack with OCR support installed via **Settings → Time & Language → Language & Region**.

## Setup
1. Clone or download this repository.
2. In PowerShell, run:
   ```powershell
   Set-ExecutionPolicy -Scope Process RemoteSigned
   .\install_winocr.ps1
   ```
   - Creates/updates the virtual environment (`.venv-winocr` by default).
   - Installs pinned dependencies from `requirements.txt`.
   - Verifies WinRT projections, PyMuPDF, and Pillow imports.

### Manual verification commands
Execute inside the activated virtual environment:
```powershell
where python
python -c "import sys; print(sys.executable)"
python -m pip list | findstr /I winrt
python -c "import winrt.windows.foundation, winrt.windows.foundation.collections"
python -c "from winrt.windows.media.ocr import OcrEngine; print(OcrEngine)"
```

## Usage
```powershell
python winocr_pdf.py -i "C:\path\sample.pdf" --lang ko-KR --dpi 300 -o .\out --fmt both
python winocr_pdf.py -i "C:\path\sample.pdf" --lang en-US --fmt text --dump-pages
```

### CLI options
- `-i/--input`: PDF file path (required).
- `-o/--outdir`: Output directory (default `.\out`).
- `--dpi`: Render DPI (default 300).
- `--lang`: BCP-47 language tag (e.g., `ko-KR`, `en-US`). Automatically normalized (e.g., `ko-KR → ko`).
- `--fmt`: Output type (`text`, `md`, `both`; aliases `txt`, `markdown`).
- `--max-pages`: Process only the first N pages.
- `--dry-run`: Perform OCR without writing any output files.
- `--dump-pages`: Save per-page PNG snapshots and `<name>_layout.json` with word bounding boxes.
- `--verbose`: Enable debug-level logging.

### Outputs
- `<outdir>\<name>.txt`: Plain-text OCR with page separators (when `--fmt` includes `text`).
- `<outdir>\<name>.md`: Markdown summary (when `--fmt` includes `md`).
- `<outdir>\<name>_layout.json`: Line/word geometry (only when `--dump-pages` is set).
- Optional PNGs: `<outdir>\<name>_page0001.png`, etc. (when `--dump-pages` is set).

## Troubleshooting
- **WSL detected**: The script exits with guidance to run from Windows PowerShell; WinRT APIs are unavailable in WSL.
- **Python version/architecture**: Ensure `where python` resolves to `3.11.x` in `C:\` (64-bit).
- **Missing WinRT modules**: Re-run `install_winocr.ps1` or execute  
  `python -m pip install -U winrt-runtime winrt-Windows.Foundation winrt-Windows.Foundation.Collections winrt-Windows.Media.Ocr winrt-Windows.Globalization winrt-Windows.Graphics.Imaging winrt-Windows.Storage winrt-Windows.Storage.Streams`.
- **Language pack errors**: Install the OCR language via Settings and restart the script.
- **PDF issues**: Confirm the file opens in a viewer; encrypted or empty PDFs will raise descriptive errors.
