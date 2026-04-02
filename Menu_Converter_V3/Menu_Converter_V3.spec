# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for Menu Converter V3.
Single exe, console=True, pywebview + Flask.
"""

from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

ort_d, ort_b, ort_h = collect_all('onnxruntime')
rocr_d, rocr_b, rocr_h = collect_all('rapidocr')
rtbl_d, rtbl_b, rtbl_h = collect_all('rapid_table')
np_d, np_b, np_h = collect_all('numpy')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=ort_b + rocr_b + rtbl_b + np_b,
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
        ('lib', 'lib'),
    ] + ort_d + rocr_d + rtbl_d + np_d,
    hiddenimports=[
        'webview',
        'flask',
        'jinja2',
        'markupsafe',
        'onnxruntime',
        'rapidocr',
        'rapidocr.main',
        'rapidocr.utils',
        'rapid_table',
        'rapid_table.main',
        'argostranslate',
        'argostranslate.package',
        'argostranslate.translate',
        'ctranslate2',
        'sentencepiece',
        'stanza',
        'fitz',
        'docx',
        'PIL',
        'PIL.Image',
        'openpyxl',
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
        'openai',
        'engineio',
        'socketio',
        'omegaconf',
        'colorlog',
        'pyclipper',
        'shapely',
    ] + ort_h + rocr_h + rtbl_h + np_h,
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
    name='Menu_Converter_V3',
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
