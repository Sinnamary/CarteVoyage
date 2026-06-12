# Phase 2 du workflow (sans argument = lit le fichier Drive de base).
# Contournements : python generer_site.py --help

param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Args)

Set-Location $PSScriptRoot
python generer_site.py @Args
exit $LASTEXITCODE
