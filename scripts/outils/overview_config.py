"""Chargement de la configuration Vue d'ensemble (data/overview_config.json)."""

from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import cast

from outils.app_types import OverviewConfig
from outils.excel_utils import data_dir

DEFAULT_OVERVIEW_CONFIG: OverviewConfig = {
    "start_date": "2026-08-03",
    "sheet_name": "Vue d'ensemble",
    "verify_markers": [],
    "day_resume_limit": 3,
    "write_snapshot": True,
}


def default_overview_config_path() -> Path:
    """Chemin par defaut de data/overview_config.json."""
    return data_dir() / "overview_config.json"


def _as_overview_config(raw: dict[str, object]) -> OverviewConfig:
    merged = deepcopy(DEFAULT_OVERVIEW_CONFIG)
    if "start_date" in raw:
        merged["start_date"] = str(raw["start_date"])
    if "sheet_name" in raw:
        merged["sheet_name"] = str(raw["sheet_name"])
    if "verify_markers" in raw:
        markers = raw["verify_markers"]
        if isinstance(markers, list):
            merged["verify_markers"] = [str(marker) for marker in markers]
    if "day_resume_limit" in raw:
        limit = raw["day_resume_limit"]
        if isinstance(limit, bool):
            merged["day_resume_limit"] = int(limit)
        elif isinstance(limit, int):
            merged["day_resume_limit"] = limit
        elif isinstance(limit, str) and limit.strip().isdigit():
            merged["day_resume_limit"] = int(limit.strip())
    if "write_snapshot" in raw:
        merged["write_snapshot"] = bool(raw["write_snapshot"])
    if "domicile" in raw:
        merged["domicile"] = str(raw["domicile"])
    if "banner_title" in raw:
        merged["banner_title"] = str(raw["banner_title"])
    return merged


def load_overview_config(path: Path | None = None) -> OverviewConfig:
    """Charge la configuration Vue d'ensemble depuis le JSON, avec valeurs par defaut."""
    config_path = path or default_overview_config_path()
    if not config_path.exists():
        return deepcopy(DEFAULT_OVERVIEW_CONFIG)

    raw = cast(object, json.loads(config_path.read_text(encoding="utf-8")))
    if not isinstance(raw, dict):
        raise ValueError(f"Configuration invalide dans {config_path}: objet JSON attendu")

    return _as_overview_config(raw)


def resolve_overview_config(
    path: Path | None = None,
    *,
    start_date: str | None = None,
) -> OverviewConfig:
    """Charge la configuration et applique une surcharge eventuelle de date de depart."""
    config = load_overview_config(path)
    if start_date is not None and start_date.strip():
        config["start_date"] = start_date.strip()
    return config
