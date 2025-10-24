param(
    [string]$VenvPath = ".\.venv-winocr"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "=== WinOCR environment bootstrap ==="
Write-Host "Target venv: $VenvPath"

if (Test-Path $VenvPath) {
    Write-Host "Virtual environment already exists. Reusing it."
} else {
    Write-Host "Creating virtual environment with Python 3.11..."
    py -3.11 -m venv $VenvPath
}

$pythonExe = Join-Path $VenvPath "Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    throw "Python executable not found at $pythonExe. Ensure Python 3.11 (64-bit) is installed."
}

Write-Host "Upgrading pip/setuptools/wheel..."
& $pythonExe -m pip install --upgrade pip setuptools wheel

$requirementsPath = Join-Path $PSScriptRoot "requirements.txt"
if (-not (Test-Path $requirementsPath)) {
    throw "requirements.txt not found at $requirementsPath."
}

Write-Host "Installing WinOCR dependencies..."
& $pythonExe -m pip install --requirement $requirementsPath

$importChecks = @(
    "winrt.windows.foundation",
    "winrt.windows.foundation.collections",
    "winrt.windows.media.ocr",
    "winrt.windows.globalization",
    "winrt.windows.graphics.imaging",
    "winrt.windows.storage",
    "winrt.windows.storage.streams"
)

Write-Host "Verifying WinRT imports..."
foreach ($module in $importChecks) {
    & $pythonExe -c "import importlib; importlib.import_module('$module')" | Out-Null
}

Write-Host "Verifying OcrEngine availability..."
& $pythonExe -c "from winrt.windows.media.ocr import OcrEngine; print('OcrEngine available:', OcrEngine)" | Out-Null

Write-Host "Verifying pymupdf and pillow..."
& $pythonExe -c "import fitz; from PIL import Image; print('PyMuPDF version:', getattr(fitz, 'VersionBind', 'unknown')); print('Pillow version:', Image.__version__)" | Out-Null

Write-Host "Environment ready. Activate with:`n  $VenvPath\Scripts\Activate.ps1"
