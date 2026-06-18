#!/usr/bin/env python3
"""Fusionne les JSON du projet et genere la page de controle inspect.html."""

from __future__ import annotations

import argparse
import csv
import json
import sys
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import cast

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl

from outils.app_types import (
    DayStatsPublic,
    FileSourceMeta,
    InspectActivity,
    InspectCheck,
    InspectCoverageRow,
    InspectDayTimeline,
    InspectMapPoint,
    InspectPayload,
    InspectSources,
    InspectSummary,
    OverviewConfig,
    OverviewDaySnapshot,
    OverviewPayload,
    StatsPayload,
    VoyageData,
)
from outils.cli_args import BuildInspectArgs
from outils.excel_utils import (
    activity_rows_from_workbook,
    build_lodging_audit,
    data_dir,
    default_excel_path,
    is_trajet_line,
    normalize_text,
    project_root,
    web_dir,
)
from outils.overview_config import (
    default_overview_config_path,
    load_overview_config,
    resolve_overview_config,
)
from site_web.build_overview import write_snapshot_from_excel
from site_web.site_nav import render_header

STATUS_OK = "ok"
STATUS_WARN = "warn"
STATUS_ERROR = "error"


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


def load_json_file(path: Path) -> object | None:
    if not path.exists():
        return None
    return cast(object, json.loads(path.read_text(encoding="utf-8")))


def parse_overview(raw: object | None) -> OverviewPayload | None:
    if isinstance(raw, dict):
        return cast(OverviewPayload, raw)
    return None


def parse_voyages(raw: object | None) -> VoyageData | None:
    if isinstance(raw, dict):
        return cast(VoyageData, raw)
    return None


def parse_stats(raw: object | None) -> StatsPayload | None:
    if isinstance(raw, dict):
        return cast(StatsPayload, raw)
    return None


def file_source_meta(path: Path) -> FileSourceMeta:
    rel = path.relative_to(project_root()).as_posix()
    if not path.exists():
        return {"path": rel, "present": False, "mtime": None, "size": 0}
    stat = path.stat()
    mtime = datetime.fromtimestamp(stat.st_mtime, tz=UTC).strftime("%d/%m/%Y %H:%M UTC")
    return {"path": rel, "present": True, "mtime": mtime, "size": stat.st_size}


def load_csv_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def add_check(
    checks: list[InspectCheck],
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
    voyages: VoyageData | None,
    stats: StatsPayload | None,
) -> list[InspectActivity]:
    items: dict[str, InspectActivity] = {}

    for point in voyages.get("points", []) if voyages else []:
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

    for row in stats.get("missing_coords", []) if stats else []:
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

    for row in stats.get("trajet_lines", []) if stats else []:
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


def overview_day_index(overview: OverviewPayload | None) -> dict[int, OverviewDaySnapshot]:
    if not overview:
        return {}
    index: dict[int, OverviewDaySnapshot] = {}
    for row in overview.get("by_day", []):
        index[int(row["jour"])] = row
    return index


def date_for_jour(start: date, jour: int) -> date:
    return start + timedelta(days=jour - 1)


def build_days_timeline(
    stats: StatsPayload | None,
    overview: OverviewPayload | None,
    config: OverviewConfig,
) -> list[InspectDayTimeline]:
    if not stats:
        return []

    start = parse_iso_date(config.get("start_date", ""))
    overview_days = overview_day_index(overview)
    days: list[InspectDayTimeline] = []

    empty_day_stats: DayStatsPublic = {
        "activities": 0,
        "geocoded": 0,
        "foot_km": 0.0,
        "car_km": 0.0,
        "foot_min": 0,
        "car_min": 0,
        "prix": 0.0,
        "couleur": "#2c3e50",
    }
    for jour in stats.get("jours", []):
        day_stats = stats.get("by_day", {}).get(str(jour), empty_day_stats)
        overview_day = overview_days.get(jour)
        day_date = None
        if overview_day and overview_day.get("date"):
            day_date = overview_day.get("date")
        elif start:
            day_date = date_for_jour(start, jour).isoformat()

        days.append(
            {
                "jour": jour,
                "date": day_date,
                "ville": overview_day.get("ville") or "" if overview_day else "",
                "resume": overview_day.get("resume") or "" if overview_day else "",
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


def overview_text_blob(overview: OverviewPayload | None, config: OverviewConfig) -> str:
    parts: list[str] = []
    intro = config.get("intro")
    if intro:
        parts.append(str(intro))
    parts.extend(str(note) for note in config.get("notes") or [])
    if not overview:
        return " ".join(parts)

    summary = overview.get("summary") or {}
    parts.extend(str(value) for value in summary.values() if value is not None)
    for day_row in overview.get("by_day", []):
        parts.extend(str(value) for value in day_row.values() if value is not None)
    for ville_row in overview.get("by_ville", []):
        parts.extend(str(value) for value in ville_row.values() if value is not None)
    for phase in overview.get("phases", []):
        parts.extend(str(value) for value in phase.values() if value is not None)
    return " ".join(parts)


def run_checks(
    checks: list[InspectCheck],
    *,
    config: OverviewConfig,
    overview: OverviewPayload | None,
    voyages: VoyageData | None,
    stats: StatsPayload | None,
    activities: list[InspectActivity],
    geocode_errors: list[dict[str, str]],
    missing_csv: list[dict[str, str]],
    sources: InspectSources,
) -> None:
    required: dict[str, tuple[str, str]] = {
        "overview_config": ("Configuration", STATUS_ERROR),
        "voyages": ("Carte (voyages.json)", STATUS_ERROR),
        "stats": ("Statistiques (stats.json)", STATUS_ERROR),
    }
    for key in ("overview_config", "voyages", "stats"):
        label, severity = required[key]
        if sources[key]["present"]:
            add_check(checks, f"source_{key}", STATUS_OK, f"{label} présent")
        else:
            add_check(checks, f"source_{key}", severity, f"{label} manquant")

    if not sources["overview"]["present"]:
        add_check(
            checks,
            "source_overview",
            STATUS_WARN,
            "overview.json absent — lancer preparer_excel.py ou build_overview.py --snapshot-only",
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
        uncovered = [
            jour
            for jour in stats.get("jours", [])
            if not any(int(p["from_jour"]) <= jour <= int(p["to_jour"]) for p in phases)
        ]
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
    activities: list[InspectActivity],
    overview: OverviewPayload | None,
) -> list[InspectCoverageRow]:
    overview_days = overview_day_index(overview)
    has_overview = overview is not None
    rows: list[InspectCoverageRow] = []
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
    item: InspectActivity,
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


def build_lodging_audit_section(
    excel_path: Path | None,
    config: OverviewConfig,
    overview: OverviewPayload | None,
    stats: StatsPayload | None,
) -> dict[str, object] | None:
    if not excel_path or not excel_path.exists():
        return None

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    rows = activity_rows_from_workbook(wb)
    jours = list(stats.get("jours") if stats else sorted({r["jour"] for r in rows}))
    domicile = normalize_text(config.get("domicile"))
    return build_lodging_audit(
        rows,
        domicile=domicile,
        jours=jours,
        overview_by_day=cast(dict[int, dict[str, object]], overview_day_index(overview)),
    )


def build_inspect_data(
    paths: dict[str, Path],
    excel_path: Path | None = None,
) -> InspectPayload:
    config = load_overview_config(paths["overview_config"])
    overview = parse_overview(load_json_file(paths["overview"]))
    voyages = parse_voyages(load_json_file(paths["voyages"]))
    stats = parse_stats(load_json_file(paths["stats"]))

    geocode_errors = load_csv_rows(paths["geocode_errors"])
    missing_csv = load_csv_rows(paths["missing_coords"])

    sources: InspectSources = {
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

    checks: list[InspectCheck] = []
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

    lodging_audit = build_lodging_audit_section(excel_path, config, overview, stats)
    if lodging_audit:
        audit_days = lodging_audit.get("days")
        mismatches = [
            d
            for d in (audit_days if isinstance(audit_days, list) else [])
            if isinstance(d, dict) and d.get("match_overview_nuit") is False
        ]
        if mismatches:
            add_check(
                checks,
                "lodging_audit",
                STATUS_WARN,
                f"{len(mismatches)} jour(s) : nuit calculée ≠ overview",
                details=", ".join(f"J{d.get('jour', '?')}" for d in mismatches),
            )
        elif overview:
            add_check(
                checks,
                "lodging_audit",
                STATUS_OK,
                "Hébergements recalculés cohérents avec overview.json",
            )

    anomalies = [row for row in coverage if row["status"] != STATUS_OK]
    error_count = sum(1 for check in checks if check["status"] == STATUS_ERROR)
    warn_count = sum(1 for check in checks if check["status"] == STATUS_WARN)
    overall = STATUS_OK
    if error_count:
        overall = STATUS_ERROR
    elif warn_count:
        overall = STATUS_WARN

    overview_summary = overview.get("summary", {}) if overview else {}
    stats_summary = stats["summary"] if stats else None
    stats_budget = stats["budget"] if stats else None
    summary: InspectSummary = {
        "title": overview_summary.get("route", "Voyage"),
        "period": overview_summary.get("period", ""),
        "route": overview_summary.get("route", ""),
        "activities": stats_summary["activities"] if stats_summary else len(activities),
        "on_map": len(voyages["points"]) if voyages else 0,
        "budget": stats_budget["total"] if stats_budget else 0,
        "jours": len(stats["jours"]) if stats else 0,
    }

    map_points: list[InspectMapPoint] = [
        {
            "jour": p["jour"],
            "ordre": p["ordre_label"],
            "nom": p["nom"],
            "ville": p.get("ville"),
            "lat": p["lat"],
            "lon": p["lon"],
            "couleur": p.get("couleur", "#2c3e50"),
        }
        for p in (voyages["points"] if voyages else [])
    ]

    return {
        "generated_at": datetime.now(UTC).strftime("%d/%m/%Y %H:%M UTC"),
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
            "domicile": config.get("domicile"),
            "verify_markers": config.get("verify_markers", []),
        },
        "lodging_audit": lodging_audit,
        "map_points": map_points,
    }


def render_html(inspect: InspectPayload) -> str:
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

    _ = write_snapshot_from_excel(path, config)
    print(f"overview.json genere : {overview_path}")
    return True


def run_build(config_path: Path | None = None, excel_path: Path | None = None) -> None:
    cfg_path = config_path or default_overview_config_path()
    _ = ensure_overview_snapshot(cfg_path, excel_path)

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

    inspect = build_inspect_data(paths, excel_path=excel_path)

    json_path = root / "inspect.json"
    _ = json_path.write_text(json.dumps(inspect, ensure_ascii=False, indent=2), encoding="utf-8")

    html_path = web_dir() / "inspect.html"
    _ = html_path.write_text(render_html(inspect), encoding="utf-8")

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
    _ = parser.add_argument(
        "excel",
        nargs="?",
        default=str(default_excel_path()),
        help="Classeur Excel pour generer overview.json si absent.",
    )
    _ = parser.add_argument(
        "--config",
        default=str(default_overview_config_path()),
        help="Fichier overview_config.json",
    )
    args = parser.parse_args(namespace=BuildInspectArgs())
    excel_path: Path | None = None
    if args.excel:
        excel = Path(args.excel).resolve()
        if excel.exists():
            excel_path = excel
    run_build(Path(args.config).resolve(), excel_path=excel_path)


if __name__ == "__main__":
    main()
