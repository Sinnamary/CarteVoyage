# Synchronise l'Excel depuis Google Drive, genere le site, puis publie.
# Usage :  .\sync_excel.ps1
#          .\sync_excel.ps1 -SkipPublish
#          .\sync_excel.ps1 -SkipGenerate

param(
    [switch]$SkipGenerate,
    [switch]$SkipPublish,
    [string]$Message = "Mise a jour du site"
)

Set-Location $PSScriptRoot

Write-Host "=== Synchronisation Excel (Google Drive) ===" -ForegroundColor Cyan
python scripts/outils/sync_excel_from_drive.py
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

if (-not $SkipGenerate) {
    Write-Host "`n=== Generation du site ===" -ForegroundColor Cyan
    python generer_site.py
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if (-not $SkipPublish) {
    & "$PSScriptRoot\publier.ps1" $Message
}
