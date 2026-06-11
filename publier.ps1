# Publie le site sur GitHub Pages (apres verification locale).
# Usage :  .\publier.ps1            (message par defaut)
#          .\publier.ps1 "Ajout de la ville de Rome"
#
# Ne regenere pas le site : lancez generer_site.py avant de publier.

param([string]$Message = "Mise a jour du site")

Set-Location $PSScriptRoot

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
