#!/usr/bin/env python3
"""Genere le site web en local a partir du fichier Excel (sans publication Git)."""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SCRIPTS = ROOT / "scripts"


def run_step(label: str, script: str, *args: str) -> None:
    path = SCRIPTS / script
    if not path.exists():
        raise SystemExit(f"Script introuvable: {path}")

    print(f"\n=== {label} ===", flush=True)
    cmd = [sys.executable, str(path), *args]
    result = subprocess.run(cmd, cwd=ROOT)
    if result.returncode != 0:
        raise SystemExit(result.returncode)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Genere le site en local (Excel -> web/). La publication Git est separee (publier.ps1).",
    )
    parser.add_argument(
        "excel",
        nargs="?",
        help="Chemin du classeur Excel (defaut: excel/Voyage Aout 2026.xlsx).",
    )
    parser.add_argument(
        "--skip-sync",
        action="store_true",
        help="Ne pas synchroniser les listes deroulantes Excel.",
    )
    parser.add_argument(
        "--skip-overview",
        action="store_true",
        help="Ne pas regenerer la feuille Vue d'ensemble.",
    )
    parser.add_argument(
        "--skip-verify",
        action="store_true",
        help="Ne pas verifier la structure du classeur.",
    )
    parser.add_argument(
        "--skip-geocode",
        action="store_true",
        help="Ne pas lancer le geocodage (les lieux deja geolocalises dans Excel ne sont jamais re-demandes).",
    )
    parser.add_argument(
        "--no-osrm",
        action="store_true",
        help="Statistiques sans appels OSRM (distances a vol d'oiseau).",
    )
    parser.add_argument(
        "--geocode-force",
        action="store_true",
        help="Regeocoder tous les lieux (passe --force a geocode_excel.py).",
    )
    args = parser.parse_args()

    excel_args: list[str] = []
    if args.excel:
        excel_args = [str(Path(args.excel).resolve())]

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

    run_step("Carte interactive", "site_web/build_map.py", *excel_args)

    stats_args = list(excel_args)
    if args.no_osrm:
        stats_args.append("--no-osrm")
    run_step("Statistiques", "site_web/build_stats.py", *stats_args)

    if args.skip_overview:
        run_step(
            "Snapshot overview.json",
            "site_web/build_overview.py",
            "--snapshot-only",
            *excel_args,
        )

    inspect_args = list(excel_args)
    run_step("Controle de coherence", "site_web/build_inspect.py", *inspect_args)

    print("\nSite genere en local.", flush=True)
    print(f"  Carte    : {ROOT / 'web' / 'index.html'}", flush=True)
    print(f"  Stats    : {ROOT / 'web' / 'stats.html'}", flush=True)
    print(f"  Controle : {ROOT / 'web' / 'inspect.html'}", flush=True)
    print("\nVerifiez le site dans le navigateur, puis publiez avec : .\\publier.ps1", flush=True)


if __name__ == "__main__":
    main()
