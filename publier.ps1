# Publie le site sur GitHub Pages.
# Usage :  .\publier.ps1            (message par defaut)
#          .\publier.ps1 "Ajout de la ville de Rome"

param([string]$Message = "Mise a jour du site")

Set-Location $PSScriptRoot

Write-Host "Synchronisation des listes deroulantes Excel..." -ForegroundColor Cyan
python scripts/sync_listes_validations.py
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "Regeneration de la carte depuis Excel..." -ForegroundColor Cyan
python scripts/build_map.py
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

git add -A

git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
    Write-Host "Aucune modification a publier." -ForegroundColor Yellow
    exit 0
}

git commit -m $Message
git push

Write-Host ""
Write-Host "Publication envoyee ! Le site sera a jour dans 1 a 2 minutes :" -ForegroundColor Green
Write-Host "https://sinnamary.github.io/CarteVoyage/web/" -ForegroundColor Cyan
