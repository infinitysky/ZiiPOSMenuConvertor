# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for Menu Converter V4.
Single exe, console=True, tkinter GUI.
"""

from PyInstaller.utils.hooks import collect_all

np_d, np_b, np_h = collect_all('numpy')
oxl_d, oxl_b, oxl_h = collect_all('openpyxl')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=np_b + oxl_b,
    datas=[
        ('lib', 'lib'),
    ] + np_d + oxl_d,
    hiddenimports=[
        'app',
        'i18n',
        'Menu_Converter',
        'pe_parser',
        'PIL',
        'PIL.Image',
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.cell._writer',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        'openpyxl.reader.excel',
        'openpyxl.writer.excel',
        'openpyxl.drawing.image',
        'xlsxwriter',
        'pandas',
        'pandas._libs',
        'pandas._libs.tslibs',
        'numpy',
        'numpy._core',
        'numpy._core._exceptions',
        'numpy._core._methods',
        'numpy._core.multiarray',
        'numpy._core.umath',
        'numpy._core._dtype',
        'numpy._core._internal',
        'numpy.core',
        'numpy.core._exceptions',
        'numpy.core._methods',
        'numpy.core.multiarray',
        'numpy.core.umath',
        'tqdm',
        'wget',
    ] + np_h + oxl_h,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Menu_Converter_V4',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
