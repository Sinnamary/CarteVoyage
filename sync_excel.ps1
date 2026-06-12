# Orchestrateur : par defaut phases 1 + 2 (workflow, sans publication).
# Usage :  .\sync_excel.ps1                  # preparer + generer
#          .\sync_excel.ps1 -Publish          # + publication (apres controle local)
#
# Contournements ponctuels :
#          .\sync_excel.ps1 -PrepareOnly
#          .\sync_excel.ps1 -WebOnly

param(
    [switch]$Publish,
    [switch]$PrepareOnly,
    [switch]$WebOnly,
    [string]$Message = "Mise a jour du site"
)

Set-Location $PSScriptRoot

if (-not $WebOnly) {
    Write-Host "=== Phase 1 : preparation Excel (Google Drive) ===" -ForegroundColor Cyan
    python preparer_excel.py
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if (-not $PrepareOnly) {
    Write-Host "`n=== Phase 2 : generation du site web ===" -ForegroundColor Cyan
    python generer_site.py
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if ($Publish -and -not $PrepareOnly) {
    & "$PSScriptRoot\publier.ps1" $Message
}
