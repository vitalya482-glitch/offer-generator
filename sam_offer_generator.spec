# -*- mode: python ; coding: utf-8 -*-

import ast
from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

ROOT_DIR = Path(SPECPATH).resolve()
APP_ICON = ROOT_DIR / 'assets' / 'app_icon.ico'
if not APP_ICON.exists():
    raise FileNotFoundError(f'Application icon not found: {APP_ICON}')


def registered_brand_modules() -> list[str]:
    """Read BRANDS from brands/registry.py without importing application code.

    Brand modules are loaded dynamically through import_module(), so PyInstaller
    cannot discover them from normal static imports. Keep the registry as the
    single source of truth and force every registered module into the frozen EXE.
    """
    registry_path = ROOT_DIR / 'brands' / 'registry.py'
    tree = ast.parse(registry_path.read_text(encoding='utf-8'), filename=str(registry_path))
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        if not any(isinstance(target, ast.Name) and target.id == 'BRANDS' for target in node.targets):
            continue
        if not isinstance(node.value, ast.Dict):
            break
        modules: list[str] = []
        for value in node.value.values:
            if isinstance(value, ast.Constant) and isinstance(value.value, str):
                modules.append(value.value)
        if modules:
            return modules
        break
    raise RuntimeError(f'Could not read BRANDS registry from {registry_path}')


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
    'gui.pages.stulz_page_runtime',
    'gui.pages.riello_page',
    'gui.pages.battery_page',
    'gui.pages.genset_page',
    'gui.pages.hvac_page',

    'brands.stulz',
    'brands.stulz_runtime',
    'brands.stulz_ui_runtime',
    'brands.stulz_legend_runtime',
    'brands.riello',
    'brands.dc_eltek',
    'brands.generator',
    'brands.hvac',
    'brands.hvac.offer_builder',
    'brands.hvac.template_finder',
    'brands.registry',

    'core.docx_renderer',
    'core.excel_reader',
    'core.excel_calc_parser',
    'core.final_offer_word_maker',
    'core.riello_price',
    'core.riello_excel_exporter',
    'core.manager_profile',
    'core.models',
    'core.project_scanner',
    'core.runtime_paths',
    'core.stulz_reference',
    'core.stulz_spec_catalog',
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

# Dynamic modules named in brands.registry are mandatory. This is more reliable
# than package discovery for local source packages and prevents a release from
# missing a newly introduced runtime wrapper such as stulz_legend_runtime.
hiddenimports += registered_brand_modules()

# Dynamic brand loading is used by brands.registry and STULZ also patches its
# page at runtime. Collect these small application packages as an additional
# safety net for helper modules that are not themselves registry targets.
hiddenimports += collect_submodules('brands')
hiddenimports += collect_submodules('gui.pages')
hiddenimports += collect_submodules('core.pdf_parsers')
hiddenimports = list(dict.fromkeys(hiddenimports))

# PyInstaller sometimes misses optional imports used by these libraries in the
# frozen EXE. Collect them explicitly so the GitHub Actions artifact runs on a
# clean Windows machine.
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('pypdf')
hiddenimports += collect_submodules('docx')
hiddenimports += collect_submodules('num2words')
hiddenimports += collect_submodules('fitz')

datas = [
    ('config.example.json', '.'),
    ('config', 'config'),
    ('assets', 'assets'),
    ('prices', 'prices'),
    ('templates', 'templates'),
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
    icon=str(APP_ICON),
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
