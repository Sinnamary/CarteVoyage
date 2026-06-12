# Phase 1 du workflow (sans argument = comportement complet).
# Contournements : python preparer_excel.py --help

param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Args)

Set-Location $PSScriptRoot
python preparer_excel.py @Args
exit $LASTEXITCODE
