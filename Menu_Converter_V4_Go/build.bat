@echo off
echo ============================================================
echo   ZiiPOS Menu Converter V4 Go -- Build Script
echo ============================================================
echo.

cd /d "%~dp0"

echo [1/5] Checking Go...
go version
if errorlevel 1 (
    echo ERROR: Go not found in PATH.
    pause
    exit /b 1
)

echo.
echo [2/5] Checking C compiler (required for WebView)...
where gcc >nul 2>&1
if errorlevel 1 (
    for /f "delims=" %%G in ('dir /s /b "%LOCALAPPDATA%\Microsoft\WinGet\Packages\*\mingw64\bin\gcc.exe" 2^>nul') do (
        set "PATH=%%~dpG;%PATH%"
        goto :gcc_found
    )
    echo ERROR: gcc not found in PATH.
    echo Install MinGW-w64: winget install BrechtSanders.WinLibs.POSIX.UCRT
    pause
    exit /b 1
)
:gcc_found
gcc --version

echo.
echo [3/5] Downloading dependencies...
set CGO_ENABLED=1
go get github.com/xuri/excelize/v2@v2.9.0
if errorlevel 1 goto :mod_fail
go get github.com/webview/webview_go@v0.0.0-20240831120633-6173450d4dd6
if errorlevel 1 goto :mod_fail
go get github.com/sqweek/dialog@v0.0.0-20260123140253-64c163d53aac
if errorlevel 1 goto :mod_fail
go mod tidy
if errorlevel 1 goto :mod_fail
echo       Dependencies OK
goto :build

:mod_fail
echo ERROR: go mod tidy failed.
echo Tip: check network / GOPROXY, or run: go env GOPROXY
pause
exit /b 1

:build
echo.
echo [4/5] Building EXE...
if not exist dist mkdir dist
go build -ldflags="-H windowsgui -s -w" -o dist\Menu_Converter_V4_Go.exe .
if errorlevel 1 (
    echo ERROR: Build failed!
    echo Tip: WebView needs CGO. Install MinGW-w64 or MSVC build tools.
    pause
    exit /b 1
)

echo.
echo [5/5] Verifying output...
if exist "dist\Menu_Converter_V4_Go.exe" (
    echo ============================================================
    echo   BUILD SUCCESS
    echo   Output: dist\Menu_Converter_V4_Go.exe
    echo   Requires: WebView2 runtime ^(Win10/11^)
    echo ============================================================
) else (
    echo ERROR: EXE not found!
)

echo.
pause
