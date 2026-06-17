#!/usr/bin/env python3
"""Genere voyages.json et les pages HTML statiques de la carte."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl

from outils.excel_utils import (
    build_voyage_data,
    data_dir,
    default_excel_path,
    find_duplicate_ordre_labels,
    find_ordre_collisions,
    iter_activity_rows,
    row_to_point,
    web_dir,
)
from site_web.site_nav import render_header

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
  <title>{title}</title>
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" crossorigin="">
  <link rel="stylesheet" href="assets/css/map.css">
</head>
<body>
  {header}

  <div class="app-layout">
    <aside class="filters-panel" id="filters-panel">
      <details class="filters-details" open>
        <summary class="filters-toggle">Filtres</summary>
        <div class="filters-body">
          <section class="filter-section">
            <h3>Jours</h3>
            <div id="filter-jours" class="filter-group"></div>
          </section>
          <section class="filter-section filter-section-map">
            <h3>Carte</h3>
            <label class="map-view-toggle" for="toggle-exclude-car">
              <input type="checkbox" id="toggle-exclude-car" checked>
              Centrer sur les activites du jour
            </label>
            <p class="filter-hint">Masque les trajets voiture et les points de transport inter-villes.</p>
          </section>
          <section class="filter-section">
            <h3>Visites</h3>
            <p class="filter-hint">Cliquez sur une visite pour la localiser sur la carte.</p>
            <div id="filter-visites" class="filter-visites-list"></div>
          </section>
          <section class="filter-section filter-section-trajets">
            <h3>Trajets</h3>
            <p class="filter-hint">Cochez un ou plusieurs trajets entre visites consecutives du meme jour (a pied en ville, en voiture si changement de ville, transport ou plus de 5 km).</p>
            <div class="trajets-actions">
              <button type="button" id="btn-trajets-clear" class="btn-secondary">Effacer les trajets</button>
            </div>
            <div id="filter-trajets" class="filter-trajets-list"></div>
            <p id="trajets-status" class="filter-hint" hidden></p>
          </section>
          <button type="button" id="btn-reset" class="btn-reset">Tout afficher</button>
          <button type="button" id="btn-close-filters" class="btn-close-filters">Voir la carte</button>
        </div>
      </details>
    </aside>
    <main class="map-container">
      <div id="map"></div>
    </main>
  </div>

  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js" crossorigin=""></script>
  <script>
    window.VOYAGE_DATA = {voyage_json};
  </script>
  <script src="assets/js/map.js"></script>
</body>
</html>
"""


def write_missing_coords_report(wb: openpyxl.Workbook) -> tuple[Path, int]:
    report_path = data_dir() / "lignes_sans_coords.csv"
    rows: list[dict[str, str]] = []

    for item in iter_activity_rows(wb):
        if row_to_point(item) is None:
            rows.append(
                {
                    "jour": str(item["jour"]),
                    "visite": str(item["visite"]),
                    "ordre": item["ordre_label"],
                    "nom": item["nom"],
                    "feuille": item["sheet_name"],
                }
            )

    with report_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f, fieldnames=["jour", "visite", "ordre", "nom", "feuille"]
        )
        writer.writeheader()
        writer.writerows(rows)

    return report_path, len(rows)


def check_lodging_consistency(voyage_data: dict) -> list[str]:
    """Vérifie la cohérence des hébergements dans les données de voyage.

    Règles attendues par jour :
    - Chaque jour possède au plus UN hébergement unique (même nom, mêmes coords).
    - Si un jour a 2 lignes hébergement (matin + soir), la première doit avoir une
      heure de début <= 10h (ou "00h00") et la dernière une heure de début >= 17h.
    - Si un jour n'a qu'une seule ligne hébergement, c'est acceptable (premier ou
      dernier jour du séjour à cet hébergement).
    - Toutes les lignes hébergement d'un même jour doivent pointer vers les mêmes
      coordonnées GPS (lat/lon identiques).

    Retourne la liste des avertissements (vide si tout est correct).
    """
    warnings: list[str] = []

    def parse_hour(heure: str | None) -> int | None:
        """Extrait l'heure depuis "HHhMM" ou "HH:MM". Retourne None si invalide."""
        if not heure:
            return None
        m = re.match(r"(\d{1,2})[h:H]", str(heure))
        return int(m.group(1)) if m else None

    # Grouper par jour → liste de points hébergement
    by_day: dict[int, list[dict]] = {}
    for pt in voyage_data.get("points", []):
        action = ((pt.get("popup") or {}).get("action") or "")
        if str(action).strip().lower() == "hébergement":
            jour = pt["jour"]
            by_day.setdefault(jour, []).append(pt)

    for jour, pts in sorted(by_day.items()):
        # Regrouper par nom (un même hébergement peut avoir 2 lignes : matin + soir)
        by_nom: dict[str, list[dict]] = {}
        for pt in pts:
            by_nom.setdefault(pt["nom"], []).append(pt)

        for nom, entries in by_nom.items():
            # Vérifier cohérence des coordonnées
            lats = {round(e["lat"], 6) for e in entries}
            lons = {round(e["lon"], 6) for e in entries}
            if len(lats) > 1 or len(lons) > 1:
                warnings.append(
                    f"J{jour} - {nom!r} : coordonnees GPS incoherentes"
                    f" lat={lats} lon={lons}"
                )

            if len(entries) < 2:
                continue  # Une seule entrée = acceptable (1er ou dernier jour)

            # Trier par numéro de visite
            entries.sort(key=lambda e: e["visite"])
            first, last = entries[0], entries[-1]

            # Entrée du matin : ouverture devrait être 00h00 ou fermeture <= 12h
            ouv_first = parse_hour(first["popup"].get("ouverture"))
            fer_first = parse_hour(first["popup"].get("fermeture"))
            is_morning = (
                (first["popup"].get("ouverture") or "").strip() == "00h00"
                or (fer_first is not None and 0 < fer_first <= 12)
            )
            if not is_morning:
                warnings.append(
                    f"J{jour} - {nom!r} - visite {first['visite']} : entree de"
                    f" depart (matin) horaires suspects"
                    f" ({first['popup'].get('ouverture')} -> {first['popup'].get('fermeture')})"
                    f" -- attendu ouverture 00h00 ou fermeture <= 12h"
                )

            # Entrée du soir : ouverture devrait être >= 17h
            ouv_last = parse_hour(last["popup"].get("ouverture"))
            is_evening = ouv_last is not None and ouv_last >= 17
            if not is_evening:
                warnings.append(
                    f"J{jour} - {nom!r} - visite {last['visite']} : entree"
                    f" d'arrivee (soir) horaires suspects"
                    f" ({last['popup'].get('ouverture')} -> {last['popup'].get('fermeture')})"
                    f" -- attendu ouverture >= 17h"
                )

    return warnings


def build_html_pages(voyage_data: dict) -> None:
    voyage_json = json.dumps(voyage_data, ensure_ascii=False)
    (web_dir() / "index.html").write_text(
        HTML_TEMPLATE.format(
            title="Carte du voyage",
            header=render_header("map"),
            voyage_json=voyage_json,
        ),
        encoding="utf-8",
    )


def run_build(excel_path: Path) -> None:
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    for (jour, visite), noms in sorted(find_ordre_collisions(wb).items()):
        print(f"ATTENTION: ordre {jour}.{visite} en double: {', '.join(noms)}")

    for label, locs in sorted(find_duplicate_ordre_labels(wb).items()):
        print(f"ATTENTION: N° étape {label} répété: {', '.join(locs)}")

    voyage_data = build_voyage_data(wb)
    json_path = data_dir() / "voyages.json"
    json_path.write_text(
        json.dumps(voyage_data, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # Vérification cohérence des hébergements (base nuitée matin/soir).
    lodging_warnings = check_lodging_consistency(voyage_data)
    for w in lodging_warnings:
        print(f"HÉBERGEMENT: {w}")

    missing_path, missing_count = write_missing_coords_report(wb)
    build_html_pages(voyage_data)

    total_rows = sum(1 for _ in iter_activity_rows(wb))
    print(f"Lignes planning: {total_rows}")
    print(f"Points sur la carte: {len(voyage_data['points'])}")
    print(f"Jours: {voyage_data['jours']}")
    print(f"JSON: {json_path}")
    print(f"HTML: {web_dir() / 'index.html'}")
    if missing_count > 0:
        print(f"Lignes sans coords: {missing_count} (voir {missing_path})")


def main() -> None:
    parser = argparse.ArgumentParser(description="Genere la carte HTML depuis le fichier Excel.")
    parser.add_argument("excel", nargs="?", default=str(default_excel_path()))
    args = parser.parse_args()

    excel_path = Path(args.excel).resolve()
    if not excel_path.exists():
        raise SystemExit(f"Fichier introuvable: {excel_path}")

    run_build(excel_path)


if __name__ == "__main__":
    main()
