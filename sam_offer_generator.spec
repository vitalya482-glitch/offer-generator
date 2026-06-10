# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

hiddenimports = [
    'gui.main_window',
    'gui.ui_style',
    'gui.path_helpers',
    'gui.settings_dialog',
    'gui.reference_table_dialog',
    'gui.spec_preview_dialog',
    'gui.calc_builder_dialog',
    'gui.pages',
    'gui.pages.stulz_page',
    'gui.pages.riello_page',
    'gui.pages.battery_page',
    'gui.pages.genset_page',

    'brands.stulz',
    'brands.riello',
    'brands.dc_eltek',
    'brands.generator',
    'brands.registry',

    'core.docx_renderer',
    'core.excel_reader',
    'core.manager_profile',
    'core.models',
    'core.project_scanner',
    'core.runtime_paths',
    'core.stulz_reference',
    'core.stulz_specification',
    'core.update_client',
    'core.utils',
    'core.pdf_parsers',
    'core.pdf_parsers.stulz_calc_pdf',
    'core.pdf_parsers.stulz_winplan_pdf',

    'PySide6.QtCore',
    'PySide6.QtGui',
    'PySide6.QtWidgets',
]

# PyInstaller sometimes misses optional imports used by these libraries in the
# frozen EXE. Collect them explicitly so the GitHub Actions artifact runs on a
# clean Windows machine.
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('pypdf')
hiddenimports += collect_submodules('docx')
hiddenimports += collect_submodules('num2words')

datas = [
    ('config.example.json', '.'),
    ('config', 'config'),
]
datas += collect_data_files('openpyxl')

# Main GUI/CLI application.
a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
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
        'pandas.tests',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SAM-Offer-Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)

# Small updater process. It is collected into the same one-dir folder.
# It does not request administrator rights and only updates files inside the
# current portable application directory.
updater_a = Analysis(
    ['updater.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PySide6', 'tkinter', 'matplotlib', 'pandas'],
    noarchive=False,
)

updater_pyz = PYZ(updater_a.pure)

updater_exe = EXE(
    updater_pyz,
    updater_a.scripts,
    [],
    exclude_binaries=True,
    name='updater',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)

coll = COLLECT(
    exe,
    updater_exe,
    a.binaries,
    updater_a.binaries,
    a.zipfiles,
    updater_a.zipfiles,
    a.datas,
    updater_a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SAM-Offer-Generator',
)
