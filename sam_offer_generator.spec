# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

hiddenimports = [
    'gui.main_window',
    'gui.ui_style',
    'gui.path_helpers',

    'brands.stulz',
    'brands.riello',
    'brands.dc_eltek',
    'brands.generator',
    'brands.registry',

    'core.docx_renderer',
    'core.excel_reader',
    'core.project_scanner',

    'PySide6.QtCore',
    'PySide6.QtGui',
    'PySide6.QtWidgets',
]

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config.example.json', '.'),
        ('config', 'config'),
    ],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'PySide6.QtWebEngineWidgets',
        'PySide6.QtWebEngineCore',
        'PySide6.QtQml',
        'PySide6.QtQuick',
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas.tests',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SAM-Offer-Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)
