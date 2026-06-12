"""Chargement de la configuration Google Drive local (data/drive_config.json)."""

from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any

from outils.excel_utils import data_dir

DEFAULT_DRIVE_CONFIG: dict[str, Any] = {
    "source_path": "",
}


def default_drive_config_path() -> Path:
    return data_dir() / "drive_config.json"


def load_drive_config(path: Path | None = None) -> dict[str, Any]:
    config_path = path or default_drive_config_path()
    config = deepcopy(DEFAULT_DRIVE_CONFIG)
    if not config_path.exists():
        return config

    raw = json.loads(config_path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError(f"Configuration invalide dans {config_path}: objet JSON attendu")

    for key, value in raw.items():
        config[key] = value
    return config


def resolve_source_path(path: Path | None = None) -> Path:
    resolved = try_resolve_source_path(path)
    if resolved is None:
        config_path = path or default_drive_config_path()
        raise ValueError(
            f"source_path manquant dans {config_path} "
            "(copiez data/drive_config.example.json et indiquez le chemin sur le disque Google Drive)."
        )
    return resolved


def try_resolve_source_path(path: Path | None = None) -> Path | None:
    config_path = path or default_drive_config_path()
    if not config_path.exists():
        return None
    config = load_drive_config(config_path)
    source = str(config.get("source_path", "")).strip()
    if not source:
        return None
    return Path(source).expanduser().resolve()
