$ErrorActionPreference = "Stop"

function Compress-WithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [int]$Attempts = 5
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            Compress-Archive -Path $Path -DestinationPath $DestinationPath -Force
            return
        } catch {
            if ($attempt -eq $Attempts) {
                throw
            }
            Start-Sleep -Seconds $attempt
        }
    }
}

python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

pyinstaller sam_offer_generator.spec --clean --noconfirm
python tools/prepare_portable_release.py --dist dist/SAM-Offer-Generator

Push-Location "dist/SAM-Offer-Generator"
try {
    & ".\SAM-Offer-Generator.exe" --self-check
    if ($LASTEXITCODE -ne 0) {
        throw "SAM-Offer-Generator.exe --self-check failed with exit code $LASTEXITCODE"
    }
} finally {
    Pop-Location
}

python tools/package_modules.py --output dist/source-modules

$zipPath = "dist/SAM-Offer-Generator-windows-portable.zip"
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}
Compress-WithRetry -Path "dist/SAM-Offer-Generator/*" -DestinationPath $zipPath

Write-Host "Portable release: $zipPath"
Write-Host "Source modules: dist/source-modules"
