"""Utilitaires partages par preparer_excel.py et generer_site.py."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from outils.drive_config import try_resolve_source_path
from outils.excel_utils import default_excel_path

ROOT = Path(__file__).resolve().parent.parent
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


def resolve_workbook_paths(
    explicit_excel: str | None,
) -> tuple[Path, Path | None, list[str]]:
    if explicit_excel:
        local = Path(explicit_excel).resolve()
        return local, None, [str(local)]

    local = default_excel_path()
    drive = try_resolve_source_path()
    return local, drive, [str(local)]


def source_excel_args(
    local_excel: Path,
    drive_path: Path | None,
    explicit_excel: str | None,
) -> list[str]:
    if explicit_excel:
        return [str(Path(explicit_excel).resolve())]
    if drive_path:
        return [str(drive_path)]
    return [str(local_excel)]
