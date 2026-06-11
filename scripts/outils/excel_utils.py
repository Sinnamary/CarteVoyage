"""Utilitaires pour lire/écrire le classeur de planning (feuilles Jour 1, Jour 2, …)."""

from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

DEFAULT_EXCEL_NAME = "Voyage Aout 2026.xlsx"

OVERVIEW_SHEET = "Vue d'ensemble"
IGNORE_SHEETS = {OVERVIEW_SHEET, "Listes"}
DAY_HEADER_ROW = 2
DATA_START_ROW = 3

PLANNING_COLUMNS = {
    "Ordre": "N° étape",
    "Nom": "Lieu",
    "Action": "Nature",
    "Type": "Catégorie",
    "Remarque": "Quartier",
    "Ville": "Ville",
    "Billet": "Réservation",
    "Prix": "Prix (€)",
    "Ouverture": "Heure début",
    "Fermeture": "Heure fin",
    "Site": "Site web",
}

MAP_EXTRA_COLUMNS = ["Latitude", "Longitude", "Lien"]
SKIP_LIEU = {"", "Journée à planifier"}

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

ORDRE_RE = re.compile(r"^\s*(\d+)[.,](\d+)\s*$")


def project_root() -> Path:
    return Path(__file__).resolve().parent.parent.parent


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
    return excel_dir() / DEFAULT_EXCEL_NAME


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_ordre(value: Any) -> dict[str, int] | None:
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

    return {"jour": int(match.group(1)), "visite": int(match.group(2))}


def jour_color(jour: int) -> str:
    return DAY_COLORS[(jour - 1) % len(DAY_COLORS)]


def is_day_sheet(sheet_name: str) -> bool:
    return sheet_name.startswith("Jour ") and sheet_name not in IGNORE_SHEETS


def day_sheets(wb: openpyxl.Workbook) -> list[str]:
    return [name for name in wb.sheetnames if is_day_sheet(name)]


def build_planning_col_index(header_row: tuple[Any, ...]) -> dict[str, int]:
    index: dict[str, int] = {}
    excel_to_internal = {excel: internal for internal, excel in PLANNING_COLUMNS.items()}

    for col_idx, value in enumerate(header_row):
        name = normalize_text(value)
        if not name:
            continue
        if name in excel_to_internal:
            internal = excel_to_internal[name]
            if internal not in index:
                index[internal] = col_idx
        elif name in MAP_EXTRA_COLUMNS and name not in index:
            index[name] = col_idx

    return index


def ensure_map_columns(ws: Worksheet) -> dict[str, int]:
    header = [cell.value for cell in ws[DAY_HEADER_ROW]]
    col_index = build_planning_col_index(tuple(header))

    next_col = len(header) + 1
    while next_col > 1 and header[next_col - 2] is None:
        next_col -= 1
    if next_col <= len(header):
        next_col = len(header) + 1

    for col_name in MAP_EXTRA_COLUMNS:
        if col_name not in col_index:
            ws.cell(row=DAY_HEADER_ROW, column=next_col, value=col_name)
            col_index[col_name] = next_col - 1
            next_col += 1

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


def parse_prix(value: Any) -> float | None:
    return parse_float(value)


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


LISTES_SHEET = "Listes"
LISTES_DATA_START_ROW = 2
# Colonnes de la feuille Listes -> plages utilisees par les listes deroulantes.
LISTES_COLUMN_LETTERS = ("A", "B", "C", "D", "E")


def listes_end_row(ws: Worksheet, col_idx: int) -> int:
    """Derniere ligne non vide d'une colonne de listes (a partir de la ligne 2)."""
    end_row = LISTES_DATA_START_ROW - 1
    for row_idx in range(LISTES_DATA_START_ROW, ws.max_row + 1):
        if normalize_text(ws.cell(row_idx, col_idx).value):
            end_row = row_idx
    return end_row


def listes_range_formula(col_letter: str, end_row: int) -> str:
    return f"Listes!${col_letter}${LISTES_DATA_START_ROW}:${col_letter}${end_row}"


def build_listes_ranges(wb: openpyxl.Workbook) -> dict[str, str]:
    """Calcule les plages Listes!$X$2:$X$N pour chaque colonne de listes."""
    ws = wb[LISTES_SHEET]
    ranges: dict[str, str] = {}
    for col_idx, col_letter in enumerate(LISTES_COLUMN_LETTERS, start=1):
        end_row = listes_end_row(ws, col_idx)
        if end_row >= LISTES_DATA_START_ROW:
            ranges[col_letter] = listes_range_formula(col_letter, end_row)
    return ranges


def sync_listes_validations(wb: openpyxl.Workbook) -> list[str]:
    """
    Met a jour les validations des feuilles Jour pour couvrir toute la feuille Listes.
    Retourne la liste des changements effectues.
    """
    ranges = build_listes_ranges(wb)
    changes: list[str] = []

    for sheet_name in day_sheets(wb):
        ws = wb[sheet_name]
        for dv in ws.data_validations.dataValidation:
            formula = str(dv.formula1 or "")
            if not formula.startswith("Listes!$"):
                continue
            for col_letter, new_range in ranges.items():
                prefix = f"Listes!${col_letter}$"
                if not formula.startswith(prefix):
                    continue
                if formula != new_range:
                    dv.formula1 = new_range
                    changes.append(f"{sheet_name}: {formula} -> {new_range}")
                break

    return changes


def ville_for_row(row: tuple[Any, ...], col_index: dict[str, int]) -> str:
    return normalize_text(cell_value(row, col_index, "Ville"))


def iter_activity_rows(wb: openpyxl.Workbook):
    for sheet_name in day_sheets(wb):
        ws = wb[sheet_name]
        header = next(
            ws.iter_rows(min_row=DAY_HEADER_ROW, max_row=DAY_HEADER_ROW, values_only=True)
        )
        col_index = build_planning_col_index(header)
        if "Nom" not in col_index or "Ordre" not in col_index:
            continue

        for row_idx, row in enumerate(
            ws.iter_rows(min_row=DATA_START_ROW, values_only=True),
            start=DATA_START_ROW,
        ):
            nom = normalize_text(cell_value(row, col_index, "Nom"))
            if nom in SKIP_LIEU:
                continue

            ordre_raw = cell_value(row, col_index, "Ordre")
            parsed = parse_ordre(ordre_raw)
            if not parsed:
                continue

            ordre_label = normalize_text(ordre_raw).replace(",", ".")
            if not ordre_label:
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
                "header_row_idx": DAY_HEADER_ROW,
                "sheet_name": sheet_name,
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
        "ville": field("Ville"),
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


def find_ordre_collisions(wb: openpyxl.Workbook) -> dict[tuple[int, int], list[str]]:
    seen: dict[tuple[int, int], list[str]] = {}
    for item in iter_activity_rows(wb):
        seen.setdefault((item["jour"], item["visite"]), []).append(item["nom"])
    return {key: noms for key, noms in seen.items() if len(noms) > 1}


def find_duplicate_ordre_labels(wb: openpyxl.Workbook) -> dict[str, list[str]]:
    seen: dict[str, list[str]] = {}
    for item in iter_activity_rows(wb):
        label = item["ordre_label"]
        seen.setdefault(label, []).append(f"{item['sheet_name']}:{item['nom']}")
    return {label: locs for label, locs in seen.items() if len(locs) > 1}


def build_voyage_data(wb: openpyxl.Workbook) -> dict[str, Any]:
    points: list[dict[str, Any]] = []
    jours: set[int] = set()

    for item in iter_activity_rows(wb):
        jours.add(item["jour"])
        point = row_to_point(item)
        if point:
            points.append(point)

    points.sort(key=lambda p: (p["jour"], p["visite"], p["id"]))
    return {"jours": sorted(jours), "points": points}
