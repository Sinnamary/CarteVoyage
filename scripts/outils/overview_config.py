"""Chargement de la configuration Vue d'ensemble (data/overview_config.json)."""

from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any

from outils.excel_utils import data_dir

DEFAULT_OVERVIEW_CONFIG: dict[str, Any] = {
    "title": "Voyage été 2026",
    "start_date": "2026-08-03",
    "sheet_name": "Vue d'ensemble",
    "intro": "",
    "route": "",
    "notes": [],
    "sections": {
        "phases": True,
        "by_day": True,
        "by_ville": True,
        "totals": True,
    },
    "verify_markers": [],
    "phases": [],
    "day_resume_limit": 3,
    "write_snapshot": True,
}


def default_overview_config_path() -> Path:
    return data_dir() / "overview_config.json"


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    merged = deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = _deep_merge(merged[key], value)
        else:
            merged[key] = value
    return merged


def load_overview_config(path: Path | None = None) -> dict[str, Any]:
    config_path = path or default_overview_config_path()
    config = deepcopy(DEFAULT_OVERVIEW_CONFIG)
    if not config_path.exists():
        return config

    raw = json.loads(config_path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError(f"Configuration invalide dans {config_path}: objet JSON attendu")

    return _deep_merge(config, raw)


def resolve_overview_config(
    path: Path | None = None,
    *,
    title: str | None = None,
    start_date: str | None = None,
) -> dict[str, Any]:
    config = load_overview_config(path)
    if title is not None and title.strip():
        config["title"] = title.strip()
    if start_date is not None and start_date.strip():
        config["start_date"] = start_date.strip()
    return config
