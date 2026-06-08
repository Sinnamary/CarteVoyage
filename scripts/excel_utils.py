"""Utilitaires partagés pour lire/écrire le fichier Excel de voyage."""

from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

PLANNING_SHEETS = {"Amsterdam"}

BASE_COLUMNS = [
    "Action",
    "Nom",
    "Type",
    "Billet",
    "Prix",
    "City Card",
    "Ouverture",
    "Fermeture",
    "Remarque",
    "Site",
]
MAP_COLUMNS = ["Ordre", "Latitude", "Longitude", "Lien"]

DAY_COLORS = [
    "#e74c3c",
    "#27ae60",
    "#3498db",
    "#9b59b6",
    "#e67e22",
    "#1abc9c",
    "#f39c12",
    "#2c3e50",
    "#d35400",
    "#16a085",
]


def project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def excel_dir() -> Path:
    path = project_root() / "excel"
    path.mkdir(exist_ok=True)
    return path


def data_dir() -> Path:
    path = project_root() / "data"
    path.mkdir(exist_ok=True)
    return path


def web_dir() -> Path:
    path = project_root() / "web"
    path.mkdir(exist_ok=True)
    return path


def default_excel_path() -> Path:
    return excel_dir() / "Voyage Amsterdam.xlsx"


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


ORDRE_RE = re.compile(r"^\s*(\d+)[.,](\d+)\s*$")


def parse_ordre(value: Any) -> dict[str, int] | None:
    """Extrait jour et visite depuis la colonne Ordre (ex. 6.5 -> jour 6, visite 5)."""
    if value is None or value == "":
        return None

    text = normalize_text(value).replace(",", ".")
    if not text:
        return None

    match = ORDRE_RE.match(text)
    if not match and isinstance(value, (int, float)):
        text = f"{float(value):g}".replace(",", ".")
        match = ORDRE_RE.match(text)

    if not match:
        return None

    return {
        "jour": int(match.group(1)),
        "visite": int(match.group(2)),
    }


def jour_color(jour: int) -> str:
    return DAY_COLORS[(jour - 1) % len(DAY_COLORS)]


def is_activity_sheet(sheet_name: str) -> bool:
    return sheet_name not in PLANNING_SHEETS


def find_header_row(ws: Worksheet) -> int | None:
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        values = [normalize_text(c) for c in row]
        if "Nom" in values:
            return row_idx
    return None


def build_column_index(header_row: tuple[Any, ...]) -> dict[str, int]:
    index: dict[str, int] = {}
    for col_idx, value in enumerate(header_row):
        name = normalize_text(value)
        if name and name not in index:
            index[name] = col_idx
    return index


def ensure_map_columns(ws: Worksheet, header_row_idx: int) -> dict[str, int]:
    header = [cell.value for cell in ws[header_row_idx]]
    col_index = build_column_index(tuple(header))

    next_col = len(header) + 1
    for col_name in MAP_COLUMNS:
        if col_name not in col_index:
            ws.cell(row=header_row_idx, column=next_col, value=col_name)
            col_index[col_name] = next_col - 1
            next_col += 1
        else:
            next_col = max(next_col, col_index[col_name] + 2)

    return col_index


def cell_value(row: tuple[Any, ...], col_index: dict[str, int], column: str) -> Any:
    if column not in col_index:
        return None
    idx = col_index[column]
    if idx >= len(row):
        return None
    return row[idx]


def parse_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def parse_int(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def has_coordinates(row: tuple[Any, ...], col_index: dict[str, int]) -> bool:
    lat = parse_float(cell_value(row, col_index, "Latitude"))
    lon = parse_float(cell_value(row, col_index, "Longitude"))
    return lat is not None and lon is not None


def backup_excel(excel_path: Path) -> Path:
    backup_dir = excel_path.parent / "backups"
    backup_dir.mkdir(exist_ok=True)
    backup_path = backup_dir / (excel_path.stem + ".backup" + excel_path.suffix)
    shutil.copy2(excel_path, backup_path)
    return backup_path


def iter_activity_rows(wb: openpyxl.Workbook):
    for sheet_name in wb.sheetnames:
        if not is_activity_sheet(sheet_name):
            continue

        ws = wb[sheet_name]
        header_row_idx = find_header_row(ws)
        if header_row_idx is None:
            continue

        header = next(ws.iter_rows(min_row=header_row_idx, max_row=header_row_idx, values_only=True))
        col_index = build_column_index(header)
        if "Nom" not in col_index:
            continue

        for row_idx, row in enumerate(
            ws.iter_rows(min_row=header_row_idx + 1, values_only=True),
            start=header_row_idx + 1,
        ):
            nom = normalize_text(cell_value(row, col_index, "Nom"))
            if not nom:
                continue

            ordre_raw = cell_value(row, col_index, "Ordre")
            parsed = parse_ordre(ordre_raw)
            if not parsed:
                continue

            ordre_label = normalize_text(ordre_raw).replace(",", ".")
            if not ORDRE_RE.match(ordre_label):
                ordre_label = f"{parsed['jour']}.{parsed['visite']}"

            yield {
                "row_idx": row_idx,
                "jour": parsed["jour"],
                "visite": parsed["visite"],
                "ordre": parsed["visite"],
                "ordre_label": ordre_label,
                "nom": nom,
                "col_index": col_index,
                "row": row,
                "ws": ws,
                "header_row_idx": header_row_idx,
            }


def row_to_point(item: dict[str, Any]) -> dict[str, Any] | None:
    row = item["row"]
    col_index = item["col_index"]

    lat = parse_float(cell_value(row, col_index, "Latitude"))
    lon = parse_float(cell_value(row, col_index, "Longitude"))
    if lat is None or lon is None:
        return None

    lien = normalize_text(cell_value(row, col_index, "Lien"))
    site = normalize_text(cell_value(row, col_index, "Site"))
    url = lien or site or None

    def field(name: str) -> Any:
        value = cell_value(row, col_index, name)
        if value is None or value == "":
            return None
        return value

    jour = item["jour"]
    visite = item["visite"]
    return {
        "id": f"{jour}.{visite}-{item['row_idx']}",
        "ordre": item["ordre"],
        "ordre_label": item["ordre_label"],
        "jour": jour,
        "visite": visite,
        "nom": item["nom"],
        "lat": lat,
        "lon": lon,
        "lien": url,
        "couleur": jour_color(jour),
        "popup": {
            "action": field("Action"),
            "type": field("Type"),
            "billet": field("Billet"),
            "prix": field("Prix"),
            "city_card": field("City Card"),
            "ouverture": field("Ouverture"),
            "fermeture": field("Fermeture"),
            "remarque": field("Remarque"),
        },
    }


def build_voyage_data(wb: openpyxl.Workbook) -> dict[str, Any]:
    points: list[dict[str, Any]] = []
    jours: set[int] = set()

    for item in iter_activity_rows(wb):
        jours.add(item["jour"])
        point = row_to_point(item)
        if point:
            points.append(point)

    points.sort(key=lambda p: (p["jour"], p["visite"], p["id"]))

    return {
        "jours": sorted(jours),
        "points": points,
    }
