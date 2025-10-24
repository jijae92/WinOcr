param(
    [switch]$Editable = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "=== pdf_text_overlay installation verifier ==="

if (-not (Test-Path ".venv") -and -not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python executable not found on PATH." -ForegroundColor Red
    exit 1
}

$python = Get-Command python | Select-Object -First 1
Write-Host "Using python at $($python.Source)"

Write-Host "Python version:"
& $python.Source -c "import sys; print(sys.version)"

if ($Editable) {
    Write-Host "Installing package in editable mode (pip install -e .)..."
    & $python.Source -m pip install -e .
} else {
    Write-Host "Installing package (pip install .)..."
    & $python.Source -m pip install .
}

Write-Host "`nChecking entry points..."

try {
    Write-Host "where pdf_text_overlay"
    where pdf_text_overlay
} catch {
    Write-Warning "pdf_text_overlay script not found on PATH. Ensure the virtual environment Scripts directory is on PATH."
}

try {
    Write-Host "`npython -m pdf_text_overlay.cli --version"
    & $python.Source -m pdf_text_overlay.cli --version
} catch {
    Write-Warning "Module execution failed. Verify the package installed correctly in the current interpreter."
}

try {
    Write-Host "`npdf_text_overlay --version"
    pdf_text_overlay --version
} catch {
    Write-Warning "pdf_text_overlay command failed. Confirm PATH includes $((Split-Path $python.Source))"
}

try {
    Write-Host "`npython -m pdf_text_overlay.cli --help | Select-Object -First 5"
    & $python.Source -m pdf_text_overlay.cli --help | Select-Object -First 5
} catch {
    Write-Warning "Unable to display CLI help."
}

Write-Host "`nValidation complete."
