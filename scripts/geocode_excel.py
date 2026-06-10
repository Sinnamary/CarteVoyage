#!/usr/bin/env python3
"""Géocode les lieux du classeur de planning et remplit Latitude/Longitude."""

from __future__ import annotations

import argparse
import csv
import json
import time
from pathlib import Path

import openpyxl
import requests

from excel_utils import (
    backup_excel,
    cell_value,
    data_dir,
    day_sheets,
    default_excel_path,
    ensure_map_columns,
    has_coordinates,
    iter_activity_rows,
    normalize_text,
    ville_for_row,
)

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
USER_AGENT = "CarteVoyage/1.0 (voyage planning map)"
REQUEST_DELAY = 1.1

MANUAL_COORDS = {
    "joods historisch museum": (52.367015, 4.903445),
    "Allard Pierson": (52.368944, 4.892944),
}

NAME_ALIASES = {
    "Het Bejinhof": "Begijnhof Amsterdam",
    "Risjkmuseum": "Rijksmuseum Amsterdam",
    "Rembrandhuis": "Museum Rembrandthuis Amsterdam",
    "Joods Historisch Museum": "Joods Historisch Museum Amsterdam",
    "Verzetmuseum": "Verzetsmuseum Amsterdam",
    "Het Scheepvaart Museum": "Het Scheepvaartmuseum Amsterdam",
    "Croisière sur les canaux": "Amsterdam canal cruise",
    "Rudi's Stroopwaffels": "Rudi's Original Stroopwafels Amsterdam",
    "Balade dans l'ancien quartier juif": "Jodenbreestraat Amsterdam",
    "Les grands canaux": "Herengracht Amsterdam",
    "Tour A'DAM": "A'DAM Lookout Amsterdam",
    "Allard Pierson": "Allard Pierson Museum Amsterdam",
}

COUNTRY_BY_VILLE: dict[str, str] = {
    "amsterdam": "nl",
    "cologne": "de",
    "köln": "de",
    "lille": "fr",
    "strasbourg": "fr",
    "ennevelin": "fr",
}

NOM_ALIASES_BY_VILLE: dict[tuple[str, str], str] = {
    ("cathédrale", "cologne"): "Kölner Dom",
    ("cathédrale", "köln"): "Kölner Dom",
}


def load_cache(cache_path: Path) -> dict:
    if cache_path.exists():
        return json.loads(cache_path.read_text(encoding="utf-8"))
    return {}


def save_cache(cache_path: Path, cache: dict) -> None:
    cache_path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def resolve_nom(nom: str) -> str:
    return NAME_ALIASES.get(nom, nom)


def resolve_nom_for_ville(nom: str, ville: str) -> str:
    key = (normalize_text(nom).lower(), normalize_text(ville).lower())
    return NOM_ALIASES_BY_VILLE.get(key, resolve_nom(nom))


def country_for_ville(ville: str) -> str:
    return COUNTRY_BY_VILLE.get(ville.strip().lower(), "nl")


def build_queries(nom: str, remarque: str, action: str = "") -> list[str]:
    search_name = resolve_nom(nom)
    queries: list[str] = []

    if remarque and action.lower() == "balade":
        queries.append(remarque)
    queries.append(search_name)
    if remarque:
        queries.append(f"{search_name}, {remarque}")

    seen: set[str] = set()
    unique: list[str] = []
    for q in queries:
        q = q.strip(" ,")
        if q and q not in seen:
            seen.add(q)
            unique.append(q)
    return unique


def nominatim_search(query: str, country: str) -> tuple[float, float] | None:
    params = {
        "q": query,
        "format": "json",
        "limit": 1,
        "countrycodes": country,
    }
    response = requests.get(
        NOMINATIM_URL,
        params=params,
        headers={"User-Agent": USER_AGENT},
        timeout=30,
    )
    response.raise_for_status()
    results = response.json()
    if not results:
        return None
    return float(results[0]["lat"]), float(results[0]["lon"])


def geocode_place(
    nom: str,
    remarque: str,
    cache: dict,
    country: str,
    action: str = "",
    use_cache: bool = True,
    ville: str = "",
) -> tuple[float, float] | None:
    search_nom = resolve_nom_for_ville(nom, ville)
    if search_nom in MANUAL_COORDS:
        return MANUAL_COORDS[search_nom]
    if nom.lower() in MANUAL_COORDS:
        return MANUAL_COORDS[nom.lower()]

    cache_key = f"{country}|{search_nom}|{remarque}"
    if use_cache and cache_key in cache:
        entry = cache[cache_key]
        return entry["lat"], entry["lon"]

    for query in build_queries(search_nom, remarque, action):
        query_key = f"query|{country}|{query}"
        if use_cache and query_key in cache:
            entry = cache[query_key]
            coords = (entry["lat"], entry["lon"])
            cache[cache_key] = {"lat": coords[0], "lon": coords[1], "source": "cache"}
            return coords

        time.sleep(REQUEST_DELAY)
        try:
            coords = nominatim_search(query, country)
        except requests.RequestException as exc:
            print(f"WARN requete Nominatim en echec pour '{query}': {exc}")
            coords = None

        if coords:
            cache[query_key] = {"lat": coords[0], "lon": coords[1], "source": "nominatim"}
            cache[cache_key] = {
                "lat": coords[0],
                "lon": coords[1],
                "source": "nominatim",
                "query": query,
            }
            return coords

    return None


def ensure_columns(wb: openpyxl.Workbook) -> None:
    for sheet_name in day_sheets(wb):
        ensure_map_columns(wb[sheet_name])


def run_geocoding(excel_path: Path, dry_run: bool = False, force: bool = False) -> None:
    wb = openpyxl.load_workbook(excel_path)
    ensure_columns(wb)

    cache_path = data_dir() / "geocode_cache.json"
    cache = load_cache(cache_path)
    errors: list[dict[str, str]] = []
    updated = 0
    skipped = 0

    for item in iter_activity_rows(wb):
        row = item["row"]
        col_index = item["col_index"]
        ws = item["ws"]
        row_idx = item["row_idx"]

        if has_coordinates(row, col_index) and not force:
            skipped += 1
            continue

        remarque = normalize_text(cell_value(row, col_index, "Remarque"))
        action = normalize_text(cell_value(row, col_index, "Action"))
        ville = ville_for_row(row, col_index)
        country = country_for_ville(ville)

        coords = geocode_place(
            item["nom"],
            remarque,
            cache,
            country=country,
            action=action,
            use_cache=not force,
            ville=ville,
        )

        label = item["ordre_label"]
        if coords:
            if not dry_run:
                ws.cell(row=row_idx, column=col_index["Latitude"] + 1, value=coords[0])
                ws.cell(row=row_idx, column=col_index["Longitude"] + 1, value=coords[1])
            updated += 1
            print(f"OK  [{label}] {item['nom']} ({ville or country}) -> {coords[0]:.6f}, {coords[1]:.6f}")
        else:
            errors.append(
                {
                    "jour": str(item["jour"]),
                    "visite": str(item["visite"]),
                    "nom": item["nom"],
                    "ville": ville,
                    "pays": country,
                    "requete": build_queries(item["nom"], remarque, action)[0],
                }
            )
            print(f"ERR [{label}] {item['nom']} -> non trouve")

    save_cache(cache_path, cache)

    errors_path = data_dir() / "geocode_errors.csv"
    with errors_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f, fieldnames=["jour", "visite", "nom", "ville", "pays", "requete"]
        )
        writer.writeheader()
        writer.writerows(errors)

    if not dry_run and updated > 0:
        backup_path = backup_excel(excel_path)
        wb.save(excel_path)
        print(f"Backup: {backup_path}")

    print(f"\nTermine: {updated} geocode(s), {skipped} ignore(s), {len(errors)} erreur(s)")
    print(f"Cache: {cache_path}")
    print(f"Erreurs: {errors_path}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Geocode les lieux dans le fichier Excel de voyage.")
    parser.add_argument("excel", nargs="?", default=str(default_excel_path()))
    parser.add_argument("--dry-run", action="store_true", help="Simuler sans ecrire dans Excel")
    parser.add_argument("--force", action="store_true", help="Re-geocoder meme si Lat/Lon presentes")
    args = parser.parse_args()

    excel_path = Path(args.excel).resolve()
    if not excel_path.exists():
        raise SystemExit(f"Fichier introuvable: {excel_path}")

    run_geocoding(excel_path, dry_run=args.dry_run, force=args.force)


if __name__ == "__main__":
    main()
