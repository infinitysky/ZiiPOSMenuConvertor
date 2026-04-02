#Requires -Version 5.1
<#
  ZiiPOS Menu Converter V3 -- Build Script
  Usage: Right-click -> Run with PowerShell
         or: powershell -ExecutionPolicy Bypass -File build.ps1
#>

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectDir

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  ZiiPOS Menu Converter V3 -- Build Script" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ── Step 1: Check Python ──────────────────────────────────────
Write-Host "[1/6] Checking Python..." -ForegroundColor Yellow
try {
    $pyver = python --version 2>&1
    Write-Host "       $pyver" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Python not found in PATH." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# ── Step 2: Install dependencies ──────────────────────────────
Write-Host ""
Write-Host "[2/6] Installing dependencies..." -ForegroundColor Yellow
pip install --upgrade pip 2>&1 | Out-Null
pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to install requirements.txt" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
pip install pyinstaller
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to install PyInstaller" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "       Dependencies OK" -ForegroundColor Green

# ── Step 3: Pre-download OCR models ──────────────────────────
Write-Host ""
Write-Host "[3/6] Pre-downloading OCR models (PP-OCRv5)..." -ForegroundColor Yellow

$ocrScript = @"
from rapidocr import RapidOCR, EngineType, LangDet, LangRec, ModelType, OCRVersion
RapidOCR(params={
    'Det.ocr_version': OCRVersion.PPOCRV5,
    'Det.lang_type': LangDet.CH,
    'Det.model_type': ModelType.MOBILE,
    'Cls.ocr_version': OCRVersion.PPOCRV5,
    'Cls.lang_type': LangDet.CH,
    'Cls.model_type': ModelType.MOBILE,
    'Rec.ocr_version': OCRVersion.PPOCRV5,
    'Rec.lang_type': LangRec.CH,
    'Rec.model_type': ModelType.MOBILE,
})
print('       PP-OCRv5 models ready')
"@
python -c $ocrScript
if ($LASTEXITCODE -ne 0) {
    Write-Host "WARNING: Could not pre-download OCR models." -ForegroundColor DarkYellow
}

python -c "from rapid_table import RapidTable; RapidTable(); print('       RapidTable model ready')"
if ($LASTEXITCODE -ne 0) {
    Write-Host "WARNING: Could not pre-download RapidTable model." -ForegroundColor DarkYellow
}

# ── Step 4: Clean previous build ──────────────────────────────
Write-Host ""
Write-Host "[4/6] Cleaning previous build..." -ForegroundColor Yellow
foreach ($dir in @("build", "dist", "__pycache__")) {
    if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
}
Get-ChildItem -Path . -Directory -Recurse -Filter "__pycache__" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "       Clean" -ForegroundColor Green

# ── Step 5: Build EXE ────────────────────────────────────────
Write-Host ""
Write-Host "[5/6] Building EXE with PyInstaller..." -ForegroundColor Yellow
Write-Host "       This may take several minutes..." -ForegroundColor DarkGray
Write-Host ""
pyinstaller Menu_Converter_V3.spec --noconfirm --clean
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "ERROR: PyInstaller build failed!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# ── Step 6: Verify output ────────────────────────────────────
Write-Host ""
Write-Host "[6/6] Verifying output..." -ForegroundColor Yellow
$exePath = Join-Path $ProjectDir "dist\Menu_Converter_V3.exe"
if (Test-Path $exePath) {
    $size = (Get-Item $exePath).Length
    $sizeMB = [math]::Round($size / 1MB, 1)
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "  BUILD SUCCESS" -ForegroundColor Green
    Write-Host "  Output: $exePath" -ForegroundColor Green
    Write-Host "  Size:   $sizeMB MB" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
} else {
    Write-Host "ERROR: EXE not found after build!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Read-Host "Press Enter to exit"
exit 0
