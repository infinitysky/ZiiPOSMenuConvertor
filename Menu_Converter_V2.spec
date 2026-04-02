# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

ort_datas, ort_binaries, ort_hidden = collect_all('onnxruntime')
rocr_datas, rocr_binaries, rocr_hidden = collect_all('rapidocr_onnxruntime')

a = Analysis(
    ['Menu_Converter_V2.py'],
    pathex=[],
    binaries=ort_binaries + rocr_binaries,
    datas=ort_datas + rocr_datas,
    hiddenimports=[
        'onnxruntime',
        'rapidocr_onnxruntime',
        'argostranslate',
        'argostranslate.translate',
        'argostranslate.package',
        'ctranslate2',
        'sentencepiece',
        'stanza',
        'fitz',
        'docx',
        'PIL',
        'openai',
        'Menu_Converter',
        'i18n',
        'ai_reader',
        'menu_parser',
        'translator',
    ] + ort_hidden + rocr_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Menu_Converter_V2',
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
