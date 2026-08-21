# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

hiddenimports = [
    "openpyxl",
    "openpyxl.cell._writer",
    "openpyxl.chart",
    "openpyxl.chartsheet",
    "openpyxl.comments",
    "openpyxl.descriptors",
    "openpyxl.drawing",
    "openpyxl.formatting",
    "openpyxl.formula",
    "openpyxl.packaging",
    "openpyxl.pivot",
    "openpyxl.reader.excel",
    "openpyxl.styles",
    "openpyxl.utils",
    "openpyxl.workbook",
    "openpyxl.worksheet",
    "openpyxl.writer.excel",
    "openpyxl.xml.functions",
    "docx",
    "docx2pdf",
    "brands",
    "brands.registry",
    "brands.stulz",
    "brands.riello",
    "brands.dc_eltek",
    "brands.battery",
    "brands.hvac",
    "brands.genset",
    "config.settings",
    "core.final_offer_word_maker",
    "core.lvk_updater_launcher",
    "core.template_engine",
    "gui.main_window",
    "gui.pages.stulz_page",
    "gui.pages.riello_page",
    "gui.pages.dc_eltek_page",
    "gui.pages.battery_page",
    "gui.pages.hvac_page",
    "gui.pages.genset_page",
]

# Brand modules are selected dynamically through brands.registry. PyInstaller
# cannot discover those imports from the string module names, so include the
# complete brand package. The STULZ feature stack also installs GUI runtime
# layers dynamically; collect all page modules for the same reason. This keeps
# new runtime files from being present in GitHub but missing from the EXE.
hiddenimports += collect_submodules("brands")
hiddenimports += collect_submodules("gui.pages")
hiddenimports += collect_submodules("openpyxl")

datas = []
datas += collect_data_files("openpyxl")
datas += [
    ("assets", "assets"),
    ("templates", "templates"),
    ("prices", "prices"),
    ("config", "config"),
    ("config.example.json", "."),
    ("requirements.txt", "."),
    ("app.update.json", "."),
]

app_a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

app_pyz = PYZ(app_a.pure, app_a.zipped_data, cipher=block_cipher)

app_exe = EXE(
    app_pyz,
    app_a.scripts,
    [],
    exclude_binaries=True,
    name="SAM-Offer-Generator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    app_exe,
    app_a.binaries,
    app_a.zipfiles,
    app_a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="SAM-Offer-Generator",
)
