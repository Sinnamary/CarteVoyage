#!/usr/bin/env python3
"""Resynchronise les listes deroulantes Excel avec la feuille Listes."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl

from outils.cli_args import SyncListesArgs
from outils.excel_utils import (
    LISTES_COLUMN_LETTERS,
    backup_excel,
    build_listes_ranges,
    default_excel_path,
    ensure_listes_values,
    sync_listes_validations,
)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Met a jour les plages de validation des feuilles Jour depuis Listes."
    )
    _ = parser.add_argument("excel", nargs="?", default=str(default_excel_path()))
    _ = parser.add_argument("--dry-run", action="store_true", help="Afficher sans ecrire")
    args = parser.parse_args(namespace=SyncListesArgs())

    excel_path = Path(args.excel).resolve()
    if not excel_path.exists():
        raise SystemExit(f"Fichier introuvable: {excel_path}")

    wb = openpyxl.load_workbook(excel_path)
    listes_changes = ensure_listes_values(wb)
    if listes_changes:
        print("Entrees Listes ajoutees:")
        for change in listes_changes:
            print(f"  - {change}")

    ranges = build_listes_ranges(wb)

    print(f"Feuille Listes ({excel_path.name}):")
    for col_letter in LISTES_COLUMN_LETTERS:
        if col_letter in ranges:
            print(f"  {col_letter}: {ranges[col_letter]}")

    changes = sync_listes_validations(wb)
    if not changes:
        print("\nRien a mettre a jour: les combos pointent deja vers les bonnes plages.")
        return

    print(f"\n{len(changes)} validation(s) a corriger:")
    for change in changes[:20]:
        print(f"  - {change}")
    if len(changes) > 20:
        print(f"  ... et {len(changes) - 20} autre(s)")

    if args.dry_run:
        print("\n(dry-run: Excel non modifie)")
        return

    backup_path = backup_excel(excel_path)
    wb.save(excel_path)
    print(f"\nBackup: {backup_path}")
    print("Excel mis a jour.")


if __name__ == "__main__":
    main()
