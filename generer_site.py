#!/usr/bin/env python3
"""Genere le site web en local a partir du fichier Excel (phase 2, sans publication Git)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

from outils.cli_args import GenererSiteArgs
from pipeline_common import ROOT, resolve_workbook_paths, run_step, source_excel_args

BYPASS_EPILOG = """
Contournements ponctuels (le comportement par defaut suit le workflow) :
  --drive-pull   retélécharge d'abord Drive vers excel/ avant de lire le fichier
  --no-osrm      statistiques sans appels OSRM
  excel          autre classeur que le fichier de base configure
"""


def main() -> None:
    """Point d'entree CLI : enchaine geocodage, stats, carte et controle."""
    parser = argparse.ArgumentParser(
        description=(
            "Phase 2 du workflow : genere web/ (carte, stats, controle) "
            "a partir du fichier de base Google Drive. "
            "Sans argument, lit directement le fichier Drive configure "
            "(apres preparer_excel.py)."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=BYPASS_EPILOG,
    )
    _ = parser.add_argument(
        "excel",
        nargs="?",
        help=argparse.SUPPRESS,
    )
    _ = parser.add_argument(
        "--drive-pull",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    _ = parser.add_argument(
        "--no-osrm",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    args = parser.parse_args(namespace=GenererSiteArgs())

    local_excel, drive_path, _ = resolve_workbook_paths(args.excel)

    if drive_path and args.drive_pull:
        run_step(
            "Relecture Google Drive",
            "outils/sync_excel_from_drive.py",
        )

    web_args = source_excel_args(local_excel, drive_path, args.excel)

    run_step("Carte interactive", "site_web/build_map.py", *web_args)

    stats_args = list(web_args)
    if args.no_osrm:
        stats_args.append("--no-osrm")
    run_step("Statistiques", "site_web/build_stats.py", *stats_args)

    run_step("Controle de coherence", "site_web/build_inspect.py", *web_args)

    print("\nSite genere en local.", flush=True)
    if drive_path and not args.excel:
        print(f"  Excel   : {drive_path} (fichier de base Google Drive)", flush=True)
    else:
        print(f"  Excel   : {web_args[0]}", flush=True)
    print(f"  Carte    : {ROOT / 'web' / 'index.html'}", flush=True)
    print(f"  Stats    : {ROOT / 'web' / 'stats.html'}", flush=True)
    print(f"  Controle : {ROOT / 'web' / 'inspect.html'}", flush=True)
    print("\nVerifiez le site dans le navigateur, puis publiez avec : .\\publier.ps1", flush=True)


if __name__ == "__main__":
    main()
