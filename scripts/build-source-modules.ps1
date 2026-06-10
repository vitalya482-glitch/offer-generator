$ErrorActionPreference = "Stop"
python tools/package_modules.py --output dist/source-modules
Write-Host "Source modules: dist/source-modules"
