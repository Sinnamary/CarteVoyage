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
    iter_activity_rows,
    row_to_point,
    web_dir,
)

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title}</title>
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" crossorigin="">
  <link rel="stylesheet" href="{assets_prefix}assets/css/map.css">
</head>
<body>
  <header class="app-header">
    <div class="header-content">
      <h1>{heading}</h1>
      <p class="subtitle">{subtitle}</p>
      <nav class="page-nav">
        <a href="{home_href}">Tout le voyage</a>
        {ville_links}
      </nav>
    </div>
  </header>

  <div class="app-layout">
    <aside class="filters-panel" id="filters-panel">
      <h2>Filtres</h2>
      <section class="filter-section">
        <h3>Villes</h3>
        <div id="filter-villes" class="filter-group"></div>
      </section>
      <section class="filter-section">
        <h3>Jours</h3>
        <div id="filter-jours" class="filter-group"></div>
      </section>
      <section class="filter-section">
        <h3>Onglets / quartiers</h3>
        <div id="filter-onglets" class="filter-group"></div>
      </section>
      <button type="button" id="btn-reset" class="btn-reset">Tout afficher</button>
    </aside>
    <main class="map-container">
      <div id="map"></div>
    </main>
  </div>

  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js" crossorigin=""></script>
  <script>
    window.VOYAGE_DATA = {voyage_json};
    window.PAGE_FILTER = {page_filter_json};
  </script>
  <script src="{assets_prefix}assets/js/map.js"></script>
</body>
</html>
"""


def write_missing_coords_report(wb: openpyxl.Workbook) -> Path:
    report_path = data_dir() / "lignes_sans_coords.csv"
    rows: list[dict[str, str]] = []

    for item in iter_activity_rows(wb):
        if row_to_point(item) is None:
            rows.append(
                {
                    "onglet": item["sheet_name"],
                    "nom": item["nom"],
                    "ville": item["ville"],
                    "jour": str(item["jour"] or ""),
                    "ordre": str(item["ordre"]),
                }
            )

    with report_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["onglet", "nom", "ville", "jour", "ordre"])
        writer.writeheader()
        writer.writerows(rows)

    return report_path


def build_ville_links(villes: list[str], current_ville: str | None, in_villes_folder: bool) -> str:
    links = []
    for ville in villes:
        slug = ville.lower().replace(" ", "-")
        href = f"villes/{slug}.html" if not in_villes_folder else f"{slug}.html"
        if in_villes_folder:
            home_href = "../index.html"
        else:
            home_href = "index.html"

        active = ' class="active"' if ville == current_ville else ""
        links.append(f'<a href="{href}"{active}>{ville}</a>')

    if in_villes_folder:
        return " ".join(links)
    return " ".join(links)


def render_html(
    voyage_data: dict,
    title: str,
    heading: str,
    subtitle: str,
    page_filter: dict,
    assets_prefix: str,
    ville_links: str,
    home_href: str,
) -> str:
    voyage_json = json.dumps(voyage_data, ensure_ascii=False)
    page_filter_json = json.dumps(page_filter, ensure_ascii=False)

    nav_villes = []
    for ville in voyage_data["villes"]:
        slug = ville.lower().replace(" ", "-")
        href = f"{assets_prefix}villes/{slug}.html"
        active = ' class="active"' if page_filter.get("ville") == ville else ""
        nav_villes.append(f'<a href="{href}"{active}>{ville}</a>')

    return HTML_TEMPLATE.format(
        title=title,
        heading=heading,
        subtitle=subtitle,
        assets_prefix=assets_prefix,
        voyage_json=voyage_json,
        page_filter_json=page_filter_json,
        home_href=home_href,
        ville_links=" ".join(nav_villes),
    )


def build_html_pages(voyage_data: dict) -> None:
    web = web_dir()
    villes_dir = web / "villes"
    villes_dir.mkdir(exist_ok=True)

    (web / "index.html").write_text(
        render_html(
            voyage_data=voyage_data,
            title="Carte du voyage",
            heading="Carte du voyage",
            subtitle="Tous les deplacements et points de visite",
            page_filter={"mode": "all"},
            assets_prefix="",
            ville_links="",
            home_href="index.html",
        ),
        encoding="utf-8",
    )

    for ville in voyage_data["villes"]:
        slug = ville.lower().replace(" ", "-")
        (villes_dir / f"{slug}.html").write_text(
            render_html(
                voyage_data=voyage_data,
                title=f"Carte — {ville}",
                heading=ville,
                subtitle=f"Points de visite a {ville}",
                page_filter={"mode": "ville", "ville": ville},
                assets_prefix="../",
                ville_links="",
                home_href="../index.html",
            ),
            encoding="utf-8",
        )


def run_build(excel_path: Path) -> None:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    voyage_data = build_voyage_data(wb)
    json_path = data_dir() / "voyages.json"
    json_path.write_text(json.dumps(voyage_data, ensure_ascii=False, indent=2), encoding="utf-8")

    missing_path = write_missing_coords_report(wb)
    build_html_pages(voyage_data)

    print(f"Points sur la carte: {len(voyage_data['points'])}")
    print(f"JSON: {json_path}")
    print(f"HTML: {web_dir() / 'index.html'}")
    if missing_path.exists():
        print(f"Lignes sans coords: {missing_path}")


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
