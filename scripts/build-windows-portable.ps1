$ErrorActionPreference = "Stop"

python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

pyinstaller sam_offer_generator.spec --clean --noconfirm
python tools/prepare_portable_release.py --dist dist/SAM-Offer-Generator
python tools/package_modules.py --output dist/source-modules

$zipPath = "dist/SAM-Offer-Generator-windows-portable.zip"
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}
Compress-Archive -Path "dist/SAM-Offer-Generator/*" -DestinationPath $zipPath -Force

Write-Host "Portable release: $zipPath"
Write-Host "Source modules: dist/source-modules"
