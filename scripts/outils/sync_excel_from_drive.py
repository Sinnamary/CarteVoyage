#!/usr/bin/env python3
"""Copie le classeur Excel depuis le disque Google Drive local vers excel/."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from outils.cli_args import SyncFromDriveArgs
from outils.drive_config import default_drive_config_path, resolve_source_path
from outils.excel_utils import default_excel_path
from outils.excel_workbook_sync import run_workbook_copy


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Copie le classeur Excel depuis le disque Google Drive local "
            "vers excel/ (avec sauvegarde horodatee de l'ancienne version)."
        ),
    )
    _ = parser.add_argument(
        "dest",
        nargs="?",
        default=str(default_excel_path()),
        help="Chemin de destination (defaut: excel/Voyage Aout 2026.xlsx).",
    )
    _ = parser.add_argument(
        "--source",
        help="Chemin source sur le disque Google Drive (sinon data/drive_config.json).",
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
        help="Ne pas sauvegarder l'ancien fichier local avant remplacement.",
    )
    args = parser.parse_args(namespace=SyncFromDriveArgs())

    config_path = Path(args.config).resolve()
    if args.source:
        source_path = Path(args.source).expanduser().resolve()
    else:
        if not config_path.exists():
            raise SystemExit(
                f"Configuration introuvable: {config_path}\n"
                + "Copiez data/drive_config.example.json vers data/drive_config.json "
                + "et renseignez source_path (chemin sur le disque Google Drive)."
            )
        source_path = resolve_source_path(config_path)

    dest_path = Path(args.dest).resolve()
    _ = run_workbook_copy(
        source_path,
        dest_path,
        dry_run=args.dry_run,
        skip_backup=args.no_backup,
        action_label="Excel synchronise depuis Google Drive",
    )


if __name__ == "__main__":
    main()
