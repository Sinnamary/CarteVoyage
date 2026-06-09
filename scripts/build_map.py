#!/usr/bin/env python3
"""Genere voyages.json et les pages HTML statiques de la carte."""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path

import openpyxl

from excel_utils import (
    build_voyage_data,
    data_dir,
    default_excel_path,
    find_ordre_collisions,
    iter_activity_rows,
    row_to_point,
    web_dir,
)

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
  <header class="app-header">
    <div class="header-content">
      <img src="assets/img/logo-cartevoyage.png" alt="CarteVoyage" class="header-logo">
      <div class="header-titles">
        <h1>{heading}</h1>
        <p class="subtitle">{subtitle}</p>
      </div>
    </div>
  </header>

  <div class="app-layout">
    <aside class="filters-panel" id="filters-panel">
      <details class="filters-details" open>
        <summary class="filters-toggle">Filtres</summary>
        <div class="filters-body">
          <section class="filter-section">
            <h3>Jours</h3>
            <div id="filter-jours" class="filter-group"></div>
          </section>
          <section class="filter-section">
            <h3>Visites</h3>
            <p class="filter-hint">Cliquez sur une visite pour la localiser sur la carte.</p>
            <div id="filter-visites" class="filter-visites-list"></div>
          </section>
          <section class="filter-section filter-section-trajets">
            <h3>Trajets a pied</h3>
            <p class="filter-hint">Cochez un ou plusieurs trajets entre visites consecutives du meme jour.</p>
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
                }
            )

    with report_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["jour", "visite", "ordre", "nom"])
        writer.writeheader()
        writer.writerows(rows)

    return report_path, len(rows)


def build_html_pages(voyage_data: dict) -> None:
    voyage_json = json.dumps(voyage_data, ensure_ascii=False)
    (web_dir() / "index.html").write_text(
        HTML_TEMPLATE.format(
            title="Carte du voyage",
            heading="Carte du voyage",
            subtitle="Points de visite par jour",
            voyage_json=voyage_json,
        ),
        encoding="utf-8",
    )


def run_build(excel_path: Path) -> None:
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    collisions = find_ordre_collisions(wb)
    for (jour, visite), noms in sorted(collisions.items()):
        print(f"ATTENTION: ordre {jour}.{visite} utilise par plusieurs lignes: {', '.join(noms)}")

    voyage_data = build_voyage_data(wb)
    json_path = data_dir() / "voyages.json"
    json_path.write_text(json.dumps(voyage_data, ensure_ascii=False, indent=2), encoding="utf-8")

    missing_path, missing_count = write_missing_coords_report(wb)
    build_html_pages(voyage_data)

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
