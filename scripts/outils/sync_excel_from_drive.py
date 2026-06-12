#!/usr/bin/env python3
"""Copie le classeur Excel depuis le disque Google Drive local vers excel/."""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl

from outils.drive_config import default_drive_config_path, resolve_source_path
from outils.excel_utils import backup_excel_timestamped, default_excel_path

GOOGLE_SHORTCUT_SUFFIXES = {".gsheet", ".gdoc", ".gslides", ".gdraw", ".gform"}
EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}


def validate_source(source_path: Path) -> None:
    if not source_path.exists():
        raise SystemExit(
            f"Fichier source introuvable: {source_path}\n"
            "Verifiez le chemin dans data/drive_config.json "
            "(Google Drive doit etre monte et le fichier synchronise)."
        )

    suffix = source_path.suffix.lower()
    if suffix in GOOGLE_SHORTCUT_SUFFIXES:
        raise SystemExit(
            f"Le fichier {source_path.name} est un raccourci Google Sheets, pas un .xlsx.\n"
            "Sur Drive, enregistrez le fichier au format Excel (.xlsx) "
            "ou placez une copie .xlsx dans un dossier synchronise."
        )
    if suffix not in EXCEL_SUFFIXES:
        raise SystemExit(
            f"Extension inattendue {suffix!r} pour {source_path.name}. "
            f"Attendu: {', '.join(sorted(EXCEL_SUFFIXES))}."
        )


def validate_workbook(path: Path) -> None:
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        wb.close()
    except Exception as exc:
        raise SystemExit(f"Fichier Excel illisible ({path.name}): {exc}") from exc


def copy_workbook(source_path: Path, dest_path: Path) -> None:
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=dest_path.suffix,
        dir=dest_path.parent,
    ) as tmp:
        tmp_path = Path(tmp.name)
    try:
        shutil.copy2(source_path, tmp_path)
        validate_workbook(tmp_path)
        tmp_path.replace(dest_path)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise


def run_sync(
    source_path: Path,
    dest_path: Path,
    *,
    dry_run: bool = False,
    skip_backup: bool = False,
) -> Path | None:
    validate_source(source_path)

    backup_path = None
    if dest_path.exists() and not skip_backup:
        backup_path = backup_excel_timestamped(dest_path)

    if dry_run:
        print("(dry-run) Copie prevue:")
        print(f"  Source : {source_path}")
        print(f"  Cible  : {dest_path}")
        if backup_path:
            print(f"  Backup : {backup_path}")
        return backup_path

    copy_workbook(source_path, dest_path)

    print(f"Excel synchronise depuis Google Drive:")
    print(f"  Source : {source_path}")
    print(f"  Cible  : {dest_path}")
    if backup_path:
        print(f"  Backup : {backup_path}")
    return backup_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Copie le classeur Excel depuis le disque Google Drive local "
            "vers excel/ (avec sauvegarde horodatee de l'ancienne version)."
        ),
    )
    parser.add_argument(
        "dest",
        nargs="?",
        default=str(default_excel_path()),
        help="Chemin de destination (defaut: excel/Voyage Aout 2026.xlsx).",
    )
    parser.add_argument(
        "--source",
        help="Chemin source sur le disque Google Drive (sinon data/drive_config.json).",
    )
    parser.add_argument(
        "--config",
        default=str(default_drive_config_path()),
        help="Fichier de configuration Drive (defaut: data/drive_config.json).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Afficher la copie prevue sans ecrire.",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Ne pas sauvegarder l'ancien fichier local avant remplacement.",
    )
    args = parser.parse_args()

    config_path = Path(args.config).resolve()
    if args.source:
        source_path = Path(args.source).expanduser().resolve()
    else:
        if not config_path.exists():
            raise SystemExit(
                f"Configuration introuvable: {config_path}\n"
                f"Copiez data/drive_config.example.json vers data/drive_config.json "
                "et renseignez source_path (chemin sur le disque Google Drive)."
            )
        source_path = resolve_source_path(config_path)

    dest_path = Path(args.dest).resolve()
    run_sync(
        source_path,
        dest_path,
        dry_run=args.dry_run,
        skip_backup=args.no_backup,
    )


if __name__ == "__main__":
    main()
