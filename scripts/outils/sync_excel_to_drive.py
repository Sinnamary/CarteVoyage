#!/usr/bin/env python3
"""Copie le classeur Excel enrichi (excel/) vers le fichier de base sur Google Drive."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from outils.cli_args import SyncToDriveArgs
from outils.drive_config import default_drive_config_path, resolve_source_path
from outils.excel_utils import default_excel_path
from outils.excel_workbook_sync import backup_drive_source, run_workbook_copy


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Ecrit le classeur Excel local (vue d'ensemble, geolocalisation…) "
            "vers le fichier de base sur Google Drive (source_path)."
        ),
    )
    _ = parser.add_argument(
        "source",
        nargs="?",
        default=str(default_excel_path()),
        help="Classeur local a envoyer (defaut: excel/Voyage Aout 2026.xlsx).",
    )
    _ = parser.add_argument(
        "--dest",
        help="Chemin cible sur le disque Google Drive (sinon data/drive_config.json).",
    )
    _ = parser.add_argument(
        "--config",
        default=str(default_drive_config_path()),
        help="Fichier de configuration Drive (defaut: data/drive_config.json).",
    )
    _ = parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Afficher la copie prevue sans ecrire.",
    )
    _ = parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Ne pas sauvegarder l'ancien fichier Drive avant remplacement.",
    )
    args = parser.parse_args(namespace=SyncToDriveArgs())

    source_path = Path(args.source).resolve()
    config_path = Path(args.config).resolve()

    if args.dest:
        dest_path = Path(args.dest).expanduser().resolve()
    else:
        if not config_path.exists():
            raise SystemExit(
                f"Configuration introuvable: {config_path}\n"
                + "Copiez data/drive_config.example.json vers data/drive_config.json "
                + "et renseignez source_path (chemin sur le disque Google Drive)."
            )
        dest_path = resolve_source_path(config_path)

    _ = run_workbook_copy(
        source_path,
        dest_path,
        dry_run=args.dry_run,
        skip_backup=args.no_backup,
        backup_dest=backup_drive_source,
        action_label="Excel ecrit sur Google Drive",
    )


if __name__ == "__main__":
    main()
