# Publie le site sur GitHub Pages (fichiers web/ uniquement).
# Usage :  .\publier.ps1
#          .\publier.ps1 "Ajout de la ville de Rome"
#
# Ne regenere pas le site : lancez generer_site.ps1 avant de publier.
# Le programme (scripts, docs…) se sauvegarde separement via git (voir README).

param([string]$Message = "Mise a jour du site")

Set-Location $PSScriptRoot

git add web/

git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
    Write-Host "Aucune modification web a publier." -ForegroundColor Yellow
    exit 0
}

git commit -m $Message
git push

Write-Host ""
Write-Host "Publication envoyee (dossier web/ uniquement) !" -ForegroundColor Green
Write-Host "Le site sera a jour dans 1 a 2 minutes :" -ForegroundColor Green
Write-Host "https://sinnamary.github.io/CarteVoyage/web/" -ForegroundColor Cyan
