# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_data_files

hiddenimports = []

# GUI
hiddenimports += collect_submodules("gui")

# CORE
hiddenimports += collect_submodules("core")

# BRANDS
hiddenimports += collect_submodules("brands")

# PySide6
hiddenimports += collect_submodules("PySide6")

datas = []
datas += collect_data_files("PySide6")

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config.example.json', '.'),
        ('templates', 'templates'),
        ('assets', 'assets'),
        ('config', 'config'),
    ] + datas,
    hiddenimports=hiddenimports,
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
    name='SAM-Offer-Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
)