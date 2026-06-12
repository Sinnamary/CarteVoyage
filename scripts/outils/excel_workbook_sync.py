"""Copie securisee de classeurs Excel (validation, backup, ecriture atomique)."""

from __future__ import annotations

import shutil
import tempfile
from datetime import datetime
from collections.abc import Callable
from pathlib import Path

import openpyxl

from outils.excel_utils import backup_excel_timestamped, excel_dir

GOOGLE_SHORTCUT_SUFFIXES = {".gsheet", ".gdoc", ".gslides", ".gdraw", ".gform"}
EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}


def validate_excel_path(path: Path, *, label: str = "Fichier") -> None:
    if not path.exists():
        raise SystemExit(f"{label} introuvable: {path}")

    suffix = path.suffix.lower()
    if suffix in GOOGLE_SHORTCUT_SUFFIXES:
        raise SystemExit(
            f"Le fichier {path.name} est un raccourci Google Sheets, pas un .xlsx.\n"
            "Sur Drive, enregistrez le fichier au format Excel (.xlsx) "
            "ou placez une copie .xlsx dans un dossier synchronise."
        )
    if suffix not in EXCEL_SUFFIXES:
        raise SystemExit(
            f"Extension inattendue {suffix!r} pour {path.name}. "
            f"Attendu: {', '.join(sorted(EXCEL_SUFFIXES))}."
        )


def validate_workbook(path: Path) -> None:
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        wb.close()
    except Exception as exc:
        raise SystemExit(f"Fichier Excel illisible ({path.name}): {exc}") from exc


def copy_workbook(source_path: Path, dest_path: Path) -> None:
    validate_excel_path(source_path, label="Fichier source")
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


def backup_drive_source(drive_path: Path) -> Path | None:
    """Sauvegarde locale du fichier Drive avant ecrasement (push vers Drive)."""
    if not drive_path.exists():
        return None
    backup_dir = excel_dir() / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = backup_dir / f"{drive_path.stem}.drive.backup.{stamp}{drive_path.suffix}"
    shutil.copy2(drive_path, backup_path)
    return backup_path


def run_workbook_copy(
    source_path: Path,
    dest_path: Path,
    *,
    dry_run: bool = False,
    skip_backup: bool = False,
    backup_dest: Callable[[Path], Path | None] | None = None,
    action_label: str = "Copie",
) -> Path | None:
    validate_excel_path(source_path, label="Fichier source")

    backup_path = None
    if dest_path.exists() and not skip_backup:
        if backup_dest is not None:
            backup_path = backup_dest(dest_path)
        else:
            backup_path = backup_excel_timestamped(dest_path)

    if dry_run:
        print(f"(dry-run) {action_label} prevue:")
        print(f"  Source : {source_path}")
        print(f"  Cible  : {dest_path}")
        if backup_path:
            print(f"  Backup : {backup_path}")
        return backup_path

    copy_workbook(source_path, dest_path)

    print(f"{action_label}:")
    print(f"  Source : {source_path}")
    print(f"  Cible  : {dest_path}")
    if backup_path:
        print(f"  Backup : {backup_path}")
    return backup_path
