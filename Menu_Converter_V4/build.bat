@echo off
echo ============================================================
echo   ZiiPOS Menu Converter V4 -- Build Script
echo ============================================================
echo.

cd /d "%~dp0"

echo [1/5] Checking Python...
python --version
if errorlevel 1 (
    echo ERROR: Python not found in PATH.
    pause
    exit /b 1
)

echo.
echo [2/5] Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install requirements.
    pause
    exit /b 1
)
pip install pyinstaller

echo.
echo [3/5] Cleaning previous build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo [4/5] Building EXE with PyInstaller...
python -m PyInstaller Menu_Converter_V4.spec --noconfirm --clean
if errorlevel 1 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo [5/5] Verifying output...
if exist "dist\Menu_Converter_V4.exe" (
    echo ============================================================
    echo   BUILD SUCCESS
    echo   Output: dist\Menu_Converter_V4.exe
    echo ============================================================
) else (
    echo ERROR: EXE not found!
)

echo.
pause
