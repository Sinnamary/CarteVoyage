#!/usr/bin/env python3
"""Fusionne les JSON du projet et genere la page de controle inspect.html."""

from __future__ import annotations

import argparse
import csv
import json
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from outils.excel_utils import data_dir, default_excel_path, normalize_text, project_root, web_dir
from outils.overview_config import default_overview_config_path, load_overview_config, resolve_overview_config
from site_web.site_nav import render_header

STATUS_OK = "ok"
STATUS_WARN = "warn"
STATUS_ERROR = "error"


def is_trajet_line(nom: str) -> bool:
    lower = normalize_text(nom).lower()
    return lower.startswith(("trajet ", "retour "))

INSPECT_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Contrôle du voyage</title>
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" crossorigin="">
  <link rel="stylesheet" href="assets/css/map.css">
  <link rel="stylesheet" href="assets/css/stats.css">
  <link rel="stylesheet" href="assets/css/inspect.css">
</head>
<body class="inspect-page">
  {header}
  <main class="inspect-main">
    <div class="inspect-intro">
      <h1>Contrôle du voyage</h1>
      <p class="inspect-generated" id="inspect-generated"></p>
    </div>
    <div id="inspect-root"></div>
  </main>
  <script>
    window.INSPECT_DATA = {inspect_json};
  </script>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js" crossorigin=""></script>
  <script src="assets/js/inspect.js"></script>
</body>
</html>
"""


def parse_iso_date(value: str) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(value[:10])
    except ValueError:
        return None


def activity_key(jour: int, ordre: str) -> str:
    return f"{jour}:{ordre}"


def load_json_file(path: Path) -> dict[str, Any] | None:
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def file_source_meta(path: Path) -> dict[str, Any]:
    rel = path.relative_to(project_root()).as_posix()
    if not path.exists():
        return {"path": rel, "present": False, "mtime": None, "size": 0}
    stat = path.stat()
    mtime = datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).strftime("%d/%m/%Y %H:%M UTC")
    return {"path": rel, "present": True, "mtime": mtime, "size": stat.st_size}


def load_csv_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def add_check(
    checks: list[dict[str, Any]],
    check_id: str,
    status: str,
    message: str,
    *,
    details: str = "",
) -> None:
    checks.append(
        {
            "id": check_id,
            "status": status,
            "message": message,
            "details": details,
        }
    )


def collect_activities(
    voyages: dict[str, Any] | None,
    stats: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    items: dict[str, dict[str, Any]] = {}

    for point in (voyages or {}).get("points", []):
        key = activity_key(point["jour"], point["ordre_label"])
        popup = point.get("popup") or {}
        items[key] = {
            "key": key,
            "jour": point["jour"],
            "ordre": point["ordre_label"],
            "nom": point["nom"],
            "ville": point.get("ville") or "",
            "action": popup.get("action") or "",
            "prix": popup.get("prix"),
            "on_map": True,
            "in_stats": True,
            "is_trajet": False,
        }

    for row in (stats or {}).get("missing_coords", []):
        key = activity_key(row["jour"], row["ordre"])
        items[key] = {
            "key": key,
            "jour": row["jour"],
            "ordre": row["ordre"],
            "nom": row["nom"],
            "ville": "",
            "action": "",
            "prix": None,
            "on_map": False,
            "in_stats": True,
            "is_trajet": False,
        }

    for row in (stats or {}).get("trajet_lines", []):
        key = activity_key(row["jour"], row["ordre"])
        items[key] = {
            "key": key,
            "jour": row["jour"],
            "ordre": row["ordre"],
            "nom": row["nom"],
            "ville": row.get("ville") or "",
            "action": "Transport",
            "prix": None,
            "on_map": False,
            "in_stats": True,
            "is_trajet": True,
        }

    return sorted(items.values(), key=lambda item: (item["jour"], item["ordre"]))


def overview_day_index(overview: dict[str, Any] | None) -> dict[int, dict[str, Any]]:
    if not overview:
        return {}
    index: dict[int, dict[str, Any]] = {}
    for row in overview.get("by_day", []):
        index[int(row["jour"])] = row
    return index


def date_for_jour(start: date, jour: int) -> date:
    return start + timedelta(days=jour - 1)


def build_days_timeline(
    stats: dict[str, Any] | None,
    overview: dict[str, Any] | None,
    config: dict[str, Any],
) -> list[dict[str, Any]]:
    if not stats:
        return []

    start = parse_iso_date(config.get("start_date", ""))
    overview_days = overview_day_index(overview)
    days: list[dict[str, Any]] = []

    for jour in stats.get("jours", []):
        day_stats = stats.get("by_day", {}).get(str(jour), {})
        overview_day = overview_days.get(jour)
        day_date = None
        if overview_day and overview_day.get("date"):
            day_date = overview_day["date"]
        elif start:
            day_date = date_for_jour(start, jour).isoformat()

        days.append(
            {
                "jour": jour,
                "date": day_date,
                "ville": (overview_day or {}).get("ville") or "",
                "resume": (overview_day or {}).get("resume") or "",
                "activities": day_stats.get("activities", 0),
                "geocoded": day_stats.get("geocoded", 0),
                "prix": day_stats.get("prix", 0),
                "foot_km": day_stats.get("foot_km", 0),
                "car_km": day_stats.get("car_km", 0),
                "couleur": day_stats.get("couleur", "#2c3e50"),
                "in_overview": jour in overview_days,
            }
        )

    return days


def overview_text_blob(overview: dict[str, Any] | None, config: dict[str, Any]) -> str:
    parts: list[str] = []
    if config.get("intro"):
        parts.append(str(config["intro"]))
    for note in config.get("notes") or []:
        parts.append(str(note))
    if not overview:
        return " ".join(parts)

    summary = overview.get("summary") or {}
    for value in summary.values():
        if value is not None:
            parts.append(str(value))
    for row in overview.get("by_day", []):
        for value in row.values():
            if value is not None:
                parts.append(str(value))
    for row in overview.get("by_ville", []):
        for value in row.values():
            if value is not None:
                parts.append(str(value))
    for phase in overview.get("phases", []):
        for value in phase.values():
            if value is not None:
                parts.append(str(value))
    return " ".join(parts)


def run_checks(
    checks: list[dict[str, Any]],
    *,
    config: dict[str, Any],
    overview: dict[str, Any] | None,
    voyages: dict[str, Any] | None,
    stats: dict[str, Any] | None,
    activities: list[dict[str, Any]],
    geocode_errors: list[dict[str, str]],
    missing_csv: list[dict[str, str]],
    sources: dict[str, dict[str, Any]],
) -> None:
    required = {
        "overview_config": ("Configuration", STATUS_ERROR),
        "voyages": ("Carte (voyages.json)", STATUS_ERROR),
        "stats": ("Statistiques (stats.json)", STATUS_ERROR),
    }
    for key, (label, severity) in required.items():
        if sources[key]["present"]:
            add_check(checks, f"source_{key}", STATUS_OK, f"{label} présent")
        else:
            add_check(checks, f"source_{key}", severity, f"{label} manquant")

    if not sources["overview"]["present"]:
        add_check(
            checks,
            "source_overview",
            STATUS_WARN,
            "overview.json absent — lancer generer_site.py ou build_overview.py --snapshot-only",
        )
    else:
        add_check(checks, "source_overview", STATUS_OK, "Synthèse (overview.json) présente")

    if voyages and stats:
        v_jours = list(voyages.get("jours", []))
        s_jours = list(stats.get("jours", []))
        if v_jours == s_jours:
            add_check(
                checks,
                "days_align",
                STATUS_OK,
                f"{len(v_jours)} jours alignés (carte ↔ stats)",
            )
        else:
            add_check(
                checks,
                "days_align",
                STATUS_ERROR,
                "Jours incohérents entre carte et stats",
                details=f"carte={v_jours}, stats={s_jours}",
            )

        geocoded_stats = stats.get("summary", {}).get("geocoded", 0)
        map_points = len(voyages.get("points", []))
        if geocoded_stats == map_points:
            add_check(
                checks,
                "geocoded_map",
                STATUS_OK,
                f"{map_points} points carte = activités géolocalisées",
            )
        else:
            add_check(
                checks,
                "geocoded_map",
                STATUS_ERROR,
                "Écart points carte / géolocalisées",
                details=f"carte={map_points}, stats={geocoded_stats}",
            )

        stats_activities = stats.get("summary", {}).get("activities", 0)
        if stats_activities == len(activities):
            add_check(
                checks,
                "activities_inventory",
                STATUS_OK,
                f"{stats_activities} activités recensées",
            )
        else:
            add_check(
                checks,
                "activities_inventory",
                STATUS_WARN,
                "Inventaire d'activités incomplet",
                details=f"stats={stats_activities}, inspect={len(activities)}",
            )

        missing_stats = stats.get("summary", {}).get("missing_coords", 0)
        csv_activity_missing = [
            row for row in missing_csv if not is_trajet_line(row.get("nom", ""))
        ]
        if missing_stats == len(csv_activity_missing):
            if missing_stats == 0:
                add_check(checks, "missing_coords", STATUS_OK, "Aucune activité hors carte")
            else:
                add_check(
                    checks,
                    "missing_coords",
                    STATUS_WARN,
                    f"{missing_stats} activité(s) sans coordonnées (CSV cohérent)",
                )
        else:
            add_check(
                checks,
                "missing_coords",
                STATUS_ERROR,
                "Écart missing_coords stats / CSV",
                details=(
                    f"stats={missing_stats}, csv_activites={len(csv_activity_missing)}, "
                    f"csv_total={len(missing_csv)}"
                ),
            )

        if overview:
            overview_days = overview.get("summary", {}).get("jours_count", 0)
            if overview_days == len(s_jours):
                add_check(
                    checks,
                    "overview_days",
                    STATUS_OK,
                    "Jours overview alignés avec stats",
                )
            else:
                add_check(
                    checks,
                    "overview_days",
                    STATUS_WARN,
                    "Écart jours overview / stats",
                    details=f"overview={overview_days}, stats={len(s_jours)}",
                )

            stats_budget = round(stats.get("budget", {}).get("total", 0), 2)
            overview_budget = round(overview.get("summary", {}).get("prix_total", 0), 2)
            if abs(stats_budget - overview_budget) < 0.01:
                add_check(
                    checks,
                    "budget_align",
                    STATUS_OK,
                    f"Budget cohérent ({stats_budget:.2f} €)",
                )
            else:
                add_check(
                    checks,
                    "budget_align",
                    STATUS_WARN,
                    "Budget overview ≠ stats",
                    details=f"overview={overview_budget:.2f} €, stats={stats_budget:.2f} €",
                )

    markers = config.get("verify_markers") or []
    if markers:
        blob = overview_text_blob(overview, config)
        missing = [marker for marker in markers if marker not in blob]
        if not overview:
            add_check(
                checks,
                "verify_markers",
                STATUS_WARN,
                "Marqueurs non vérifiables (overview.json absent)",
            )
        elif not missing:
            add_check(checks, "verify_markers", STATUS_OK, "Marqueurs de vérification présents")
        else:
            add_check(
                checks,
                "verify_markers",
                STATUS_WARN,
                f"{len(missing)} marqueur(s) absent(s)",
                details=", ".join(missing),
            )

    phases = (overview or {}).get("phases") or []
    if phases and stats:
        uncovered = []
        for jour in stats.get("jours", []):
            if not any(int(p["from_jour"]) <= jour <= int(p["to_jour"]) for p in phases):
                uncovered.append(jour)
        if not uncovered:
            add_check(checks, "phases_cover", STATUS_OK, "Tous les jours couverts par une phase")
        else:
            add_check(
                checks,
                "phases_cover",
                STATUS_WARN,
                f"Jours hors phases config : {uncovered}",
            )

    if geocode_errors:
        add_check(
            checks,
            "geocode_errors",
            STATUS_WARN,
            f"{len(geocode_errors)} erreur(s) de géocodage",
        )
    elif sources["geocode_cache"]["present"]:
        add_check(checks, "geocode_errors", STATUS_OK, "Aucune erreur de géocodage (CSV vide)")


def build_coverage_rows(
    activities: list[dict[str, Any]],
    overview: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    overview_days = overview_day_index(overview)
    has_overview = overview is not None
    rows: list[dict[str, Any]] = []
    for item in activities:
        jour = item["jour"]
        in_overview = jour in overview_days if has_overview else None
        rows.append(
            {
                "jour": jour,
                "ordre": item["ordre"],
                "nom": item["nom"],
                "ville": item["ville"],
                "in_stats": item["in_stats"],
                "on_map": item["on_map"],
                "in_overview": in_overview,
                "is_trajet": item["is_trajet"],
                "status": _coverage_status(item, in_overview, has_overview),
            }
        )
    return rows


def _coverage_status(
    item: dict[str, Any],
    in_overview: bool | None,
    has_overview: bool,
) -> str:
    if not item["in_stats"]:
        return STATUS_ERROR
    if item["is_trajet"]:
        if has_overview and in_overview is False:
            return STATUS_WARN
        return STATUS_OK
    if item["on_map"]:
        if has_overview and in_overview is False:
            return STATUS_WARN
        return STATUS_OK
    return STATUS_WARN


def build_inspect_data(paths: dict[str, Path]) -> dict[str, Any]:
    config = load_overview_config(paths["overview_config"])
    overview = load_json_file(paths["overview"])
    voyages = load_json_file(paths["voyages"])
    stats = load_json_file(paths["stats"])

    geocode_errors = load_csv_rows(paths["geocode_errors"])
    missing_csv = load_csv_rows(paths["missing_coords"])

    sources = {
        "overview_config": file_source_meta(paths["overview_config"]),
        "overview": file_source_meta(paths["overview"]),
        "voyages": file_source_meta(paths["voyages"]),
        "stats": file_source_meta(paths["stats"]),
        "geocode_cache": file_source_meta(paths["geocode_cache"]),
        "route_stats_cache": file_source_meta(paths["route_stats_cache"]),
        "geocode_errors": file_source_meta(paths["geocode_errors"]),
        "missing_coords": file_source_meta(paths["missing_coords"]),
    }

    activities = collect_activities(voyages, stats)
    for item in activities:
        item["in_overview"] = item["jour"] in overview_day_index(overview)

    days = build_days_timeline(stats, overview, config)
    coverage = build_coverage_rows(activities, overview)

    checks: list[dict[str, Any]] = []
    run_checks(
        checks,
        config=config,
        overview=overview,
        voyages=voyages,
        stats=stats,
        activities=activities,
        geocode_errors=geocode_errors,
        missing_csv=missing_csv,
        sources=sources,
    )

    anomalies = [row for row in coverage if row["status"] != STATUS_OK]
    error_count = sum(1 for check in checks if check["status"] == STATUS_ERROR)
    warn_count = sum(1 for check in checks if check["status"] == STATUS_WARN)
    overall = STATUS_OK
    if error_count:
        overall = STATUS_ERROR
    elif warn_count:
        overall = STATUS_WARN

    summary = {
        "title": (overview or {}).get("summary", {}).get("route", "Voyage"),
        "period": (overview or {}).get("summary", {}).get("period", ""),
        "route": (overview or {}).get("summary", {}).get("route", ""),
        "activities": (stats or {}).get("summary", {}).get("activities", len(activities)),
        "on_map": len((voyages or {}).get("points", [])),
        "budget": (stats or {}).get("budget", {}).get("total", 0),
        "jours": len((stats or {}).get("jours", [])),
    }

    map_points = [
        {
            "jour": p["jour"],
            "ordre": p["ordre_label"],
            "nom": p["nom"],
            "ville": p.get("ville"),
            "lat": p["lat"],
            "lon": p["lon"],
            "couleur": p.get("couleur", "#2c3e50"),
        }
        for p in (voyages or {}).get("points", [])
    ]

    return {
        "generated_at": datetime.now(timezone.utc).strftime("%d/%m/%Y %H:%M UTC"),
        "overall_status": overall,
        "has_overview": overview is not None,
        "summary": summary,
        "sources": sources,
        "checks": checks,
        "days": days,
        "activities": activities,
        "coverage": coverage,
        "anomalies": anomalies,
        "geocode_errors": geocode_errors[:50],
        "missing_coords": missing_csv[:50],
        "config": {
            "start_date": config.get("start_date"),
            "verify_markers": config.get("verify_markers", []),
        },
        "map_points": map_points,
    }


def render_html(inspect: dict[str, Any]) -> str:
    return INSPECT_HTML_TEMPLATE.format(
        header=render_header(active="inspect"),
        inspect_json=json.dumps(inspect, ensure_ascii=False),
    )


def ensure_overview_snapshot(config_path: Path, excel_path: Path | None = None) -> bool:
    """Genere overview.json depuis Excel s'il est absent."""
    overview_path = data_dir() / "overview.json"
    if overview_path.exists():
        return False

    path = excel_path or default_excel_path()
    if not path.exists():
        return False

    config = resolve_overview_config(config_path)
    if not config.get("write_snapshot", True):
        return False

    from site_web.build_overview import write_snapshot_from_excel

    write_snapshot_from_excel(path, config)
    print(f"overview.json genere : {overview_path}")
    return True


def run_build(config_path: Path | None = None, excel_path: Path | None = None) -> None:
    cfg_path = config_path or default_overview_config_path()
    ensure_overview_snapshot(cfg_path, excel_path)

    root = data_dir()
    paths = {
        "overview_config": cfg_path,
        "overview": root / "overview.json",
        "voyages": root / "voyages.json",
        "stats": root / "stats.json",
        "geocode_cache": root / "geocode_cache.json",
        "route_stats_cache": root / "route_stats_cache.json",
        "geocode_errors": root / "geocode_errors.csv",
        "missing_coords": root / "lignes_sans_coords.csv",
    }

    inspect = build_inspect_data(paths)

    json_path = root / "inspect.json"
    json_path.write_text(json.dumps(inspect, ensure_ascii=False, indent=2), encoding="utf-8")

    html_path = web_dir() / "inspect.html"
    html_path.write_text(render_html(inspect), encoding="utf-8")

    ok = sum(1 for c in inspect["checks"] if c["status"] == STATUS_OK)
    warn = sum(1 for c in inspect["checks"] if c["status"] == STATUS_WARN)
    err = sum(1 for c in inspect["checks"] if c["status"] == STATUS_ERROR)

    print(f"Controle genere : {inspect['overall_status'].upper()}")
    print(f"  Verifications : {ok} ok, {warn} avert., {err} erreur(s)")
    print(f"  Activites : {len(inspect['activities'])} · Anomalies : {len(inspect['anomalies'])}")
    print(f"  JSON : {json_path}")
    print(f"  HTML : {html_path}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Genere inspect.json et inspect.html.")
    parser.add_argument(
        "excel",
        nargs="?",
        default=str(default_excel_path()),
        help="Classeur Excel pour generer overview.json si absent.",
    )
    parser.add_argument(
        "--config",
        default=str(default_overview_config_path()),
        help="Fichier overview_config.json",
    )
    args = parser.parse_args()
    excel = Path(args.excel).resolve() if args.excel else None
    run_build(Path(args.config).resolve(), excel_path=excel if excel.exists() else None)


if __name__ == "__main__":
    main()
