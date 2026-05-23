# ZiiPOS Menu Converter V4 (Go)

Web UI + embedded WebView2 browser, packaged as a single Windows exe.

## Features

- Excel Import (same logic as Python V4)
- PE Menu Import + image extraction
- UI languages: EN / CN / JP
- Closing the window shuts down the local HTTP server

## Build

Requirements:

- Go 1.22+
- **CGO enabled** (`CGO_ENABLED=1`) — WebView binding needs a C compiler
- MinGW-w64 or MSVC (Visual Studio Build Tools)
- WebView2 runtime (preinstalled on most Windows 10/11)

```bat
build.bat
```

Output: `dist\Menu_Converter_V4_Go.exe`

## Manual build

```bat
set CGO_ENABLED=1
go get github.com/xuri/excelize/v2@v2.9.0
go get github.com/webview/webview_go@v0.0.0-20240831120633-6173450d4dd6
go get github.com/sqweek/dialog@v0.0.0-20260123140253-64c163d53aac
go mod tidy
go build -ldflags="-H windowsgui -s -w" -o dist\Menu_Converter_V4_Go.exe .
```

## Troubleshooting

| Error | Fix |
|-------|-----|
| `unknown revision` in go mod tidy | Run `build.bat` again — pinned versions in script |
| `gcc not found` / CGO error | Install [MinGW-w64](https://www.mingw-w64.org/) or VS Build Tools |
| Blank window | Install [WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/) |

## Config

- `%APPDATA%\ZiiPOSMenuConverter\settings_v4_go.json`
- Template: `C:\Ziitech\ZiiPOS_MenuTemplate.xlsx` (auto-download)
- Default output: `C:\Ziitech\Menu`
