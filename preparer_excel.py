#!/usr/bin/env python3
"""Prepare le classeur Excel : lecture Drive, verifications, enrichissements, ecriture Drive."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

from pipeline_common import resolve_workbook_paths, run_step

BYPASS_EPILOG = """
Contournements ponctuels (le comportement par defaut suit le workflow) :
  --skip-drive-pull   ne pas telecharger depuis Drive (reprise sur excel/ local)
  --skip-drive-push   ne pas reecrire sur Drive (test local uniquement)
  --skip-sync         ne pas resynchroniser les listes deroulantes
  --skip-overview     ne pas regenerer la vue d'ensemble
  --skip-verify       ne pas verifier la structure (deconseille)
  --skip-geocode      ne pas geocoder les nouveaux lieux
  --geocode-force     regeocoder tous les lieux
  excel               autre classeur que excel/ par defaut
"""


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Phase 1 du workflow : lit le fichier de base Google Drive, "
            "enrichit le classeur (listes, vue d'ensemble, geolocalisation), "
            "verifie la structure et reecrit le resultat sur Drive."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=BYPASS_EPILOG,
    )
    parser.add_argument("excel", nargs="?", help=argparse.SUPPRESS)
    parser.add_argument("--skip-drive-pull", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-drive-push", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-sync", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-overview", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-verify", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-geocode", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--geocode-force", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()

    _, drive_path, excel_args = resolve_workbook_paths(args.excel)

    if drive_path and not args.skip_drive_pull:
        run_step(
            "Lecture Google Drive (backup local)",
            "outils/sync_excel_from_drive.py",
        )

    if not args.skip_sync:
        run_step(
            "Synchronisation des listes deroulantes",
            "outils/sync_listes_validations.py",
            *excel_args,
        )

    if not args.skip_overview:
        run_step(
            "Vue d'ensemble Excel",
            "site_web/build_overview.py",
            *excel_args,
        )

    if not args.skip_verify:
        run_step(
            "Verification du classeur",
            "outils/verify_planning_workbook.py",
            *excel_args,
        )

    if not args.skip_geocode:
        geocode_args = list(excel_args)
        if args.geocode_force:
            geocode_args.append("--force")
        run_step(
            "Geocodage des lieux",
            "outils/geocode_excel.py",
            *geocode_args,
        )

    if drive_path and not args.skip_drive_push:
        run_step(
            "Ecriture sur Google Drive",
            "outils/sync_excel_to_drive.py",
            *excel_args,
        )

    print("\nClasseur prepare.", flush=True)
    if drive_path:
        print(f"  Fichier de base : {drive_path}", flush=True)
    else:
        print(f"  Fichier local   : {excel_args[0]}", flush=True)
    print(
        "\nCorrigez le planning sur Google Drive si besoin, puis relancez preparer_excel.py.",
        flush=True,
    )
    print("Quand le classeur est valide, generez le site : .\\generer_site.ps1", flush=True)


if __name__ == "__main__":
    main()
