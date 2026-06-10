#!/usr/bin/env python3
"""Genere stats.json et stats.html depuis le classeur Excel."""

from __future__ import annotations

import argparse
import json
import math
import sys
import time
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import openpyxl
import requests

from excel_utils import (
    cell_value,
    data_dir,
    default_excel_path,
    iter_activity_rows,
    jour_color,
    normalize_text,
    row_to_point,
    web_dir,
)
from site_nav import render_header

WALKING_MAX_AIR_DISTANCE_M = 5000
ROUTE_DELAY_S = 0.12
OSRM_SERVERS = {
    "foot": [
        "https://routing.openstreetmap.de/routed-foot/route/v1/foot",
        "https://router.project-osrm.org/route/v1/foot",
    ],
    "car": [
        "https://routing.openstreetmap.de/routed-car/route/v1/car",
        "https://router.project-osrm.org/route/v1/driving",
    ],
}

STATS_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Statistiques du voyage</title>
  <link rel="stylesheet" href="assets/css/map.css">
  <link rel="stylesheet" href="assets/css/stats.css">
</head>
<body class="stats-page">
  {header}
  <main class="stats-main">
    <div class="stats-intro">
      <h1>Statistiques du voyage</h1>
      <p class="stats-generated">Genere le {generated_at}</p>
    </div>
    {content}
  </main>
</body>
</html>
"""


def haversine_m(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    r = 6371000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def is_transport(action: str) -> bool:
    return normalize_text(action).lower() == "transport"


def is_trajet_line(nom: str) -> bool:
    lower = normalize_text(nom).lower()
    return lower.startswith(("trajet ", "retour "))


def parse_prix(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def row_entry(item: dict[str, Any]) -> dict[str, Any]:
    row = item["row"]
    ci = item["col_index"]
    pt = row_to_point(item)
    nom = item["nom"]
    return {
        "jour": item["jour"],
        "visite": item["visite"],
        "ordre": item["ordre_label"],
        "nom": nom,
        "ville": normalize_text(cell_value(row, ci, "Ville")),
        "action": normalize_text(cell_value(row, ci, "Action")),
        "type": normalize_text(cell_value(row, ci, "Type")),
        "billet": normalize_text(cell_value(row, ci, "Billet")),
        "prix": parse_prix(cell_value(row, ci, "Prix")),
        "ouverture": cell_value(row, ci, "Ouverture"),
        "fermeture": cell_value(row, ci, "Fermeture"),
        "remarque": normalize_text(cell_value(row, ci, "Remarque")),
        "has_coords": pt is not None,
        "lat": pt["lat"] if pt else None,
        "lon": pt["lon"] if pt else None,
        "couleur": jour_color(item["jour"]),
        "is_trajet_line": is_trajet_line(nom),
    }


def segment_travel_mode(from_r: dict, to_r: dict) -> str:
    if is_transport(from_r["action"]) or is_transport(to_r["action"]):
        return "car"
    ville_from = from_r["ville"].strip().lower()
    ville_to = to_r["ville"].strip().lower()
    if ville_from and ville_to and ville_from != ville_to:
        return "car"
    if from_r["has_coords"] and to_r["has_coords"]:
        dist = haversine_m(
            from_r["lat"], from_r["lon"], to_r["lat"], to_r["lon"]
        )
        if dist > WALKING_MAX_AIR_DISTANCE_M:
            return "car"
    return "foot"


def load_route_cache(path: Path) -> dict[str, Any]:
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {}


def save_route_cache(path: Path, cache: dict[str, Any]) -> None:
    path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def route_cache_key(
    mode: str, lat1: float, lon1: float, lat2: float, lon2: float
) -> str:
    return f"{mode}:{lat1:.6f},{lon1:.6f}->{lat2:.6f},{lon2:.6f}"


def fetch_osrm_route(
    lat1: float, lon1: float, lat2: float, lon2: float, mode: str
) -> dict[str, Any] | None:
    servers = OSRM_SERVERS.get(mode, OSRM_SERVERS["foot"])
    for base in servers:
        url = (
            f"{base}/{lon1},{lat1};{lon2},{lat2}"
            "?overview=false&geometries=geojson"
        )
        try:
            response = requests.get(url, timeout=30)
            if not response.ok:
                continue
            payload = response.json()
            if payload.get("code") != "Ok" or not payload.get("routes"):
                continue
            route = payload["routes"][0]
            return {
                "distance_m": round(route["distance"]),
                "duration_s": round(route["duration"]),
                "source": "osrm",
            }
        except requests.RequestException:
            continue
    return None


def resolve_route(
    from_r: dict,
    to_r: dict,
    mode: str,
    cache: dict[str, Any],
    use_osrm: bool,
) -> dict[str, Any]:
    air_m = haversine_m(from_r["lat"], from_r["lon"], to_r["lat"], to_r["lon"])
    key = route_cache_key(mode, from_r["lat"], from_r["lon"], to_r["lat"], to_r["lon"])

    if key in cache:
        result = dict(cache[key])
        result["cached"] = True
        return result

    result: dict[str, Any] = {
        "distance_m": round(air_m),
        "duration_s": None,
        "source": "air",
        "cached": False,
    }

    if use_osrm:
        osrm = fetch_osrm_route(
            from_r["lat"], from_r["lon"], to_r["lat"], to_r["lon"], mode
        )
        if osrm:
            result = osrm
            result["cached"] = False
        time.sleep(ROUTE_DELAY_S)

    if mode == "car" and result["duration_s"] is None:
        result["duration_s"] = round(air_m / 22)
    elif mode == "foot" and result["duration_s"] is None:
        result["duration_s"] = round(air_m / 1.4)

    cache[key] = {
        "distance_m": result["distance_m"],
        "duration_s": result["duration_s"],
        "source": result["source"],
    }
    return result


def build_segments(
    rows: list[dict[str, Any]],
    cache: dict[str, Any],
    use_osrm: bool,
) -> list[dict[str, Any]]:
    by_day: dict[int, list[dict]] = defaultdict(list)
    for row in rows:
        by_day[row["jour"]].append(row)
    for jour in by_day:
        by_day[jour].sort(key=lambda x: x["visite"])

    segments: list[dict[str, Any]] = []
    for jour, day_rows in sorted(by_day.items()):
        for i in range(len(day_rows) - 1):
            from_r, to_r = day_rows[i], day_rows[i + 1]
            mode = segment_travel_mode(from_r, to_r)
            seg: dict[str, Any] = {
                "jour": jour,
                "from_ordre": from_r["ordre"],
                "to_ordre": to_r["ordre"],
                "from_nom": from_r["nom"],
                "to_nom": to_r["nom"],
                "from_ville": from_r["ville"],
                "to_ville": to_r["ville"],
                "mode": mode,
                "calculable": from_r["has_coords"] and to_r["has_coords"],
            }
            if seg["calculable"]:
                route = resolve_route(from_r, to_r, mode, cache, use_osrm)
                seg["distance_m"] = route["distance_m"]
                seg["duration_s"] = route["duration_s"]
                seg["source"] = route["source"]
                seg["air_m"] = round(
                    haversine_m(
                        from_r["lat"], from_r["lon"], to_r["lat"], to_r["lon"]
                    )
                )
            segments.append(seg)
    return segments


def counter_dict(counter: Counter) -> dict[str, int]:
    return dict(counter.most_common())


def format_distance(meters: int | float | None) -> str:
    if meters is None:
        return "—"
    if meters < 950:
        return f"{round(meters)} m"
    return f"{meters / 1000:.1f} km".replace(".", ",")


def format_duration(seconds: int | float | None) -> str:
    if seconds is None:
        return "—"
    minutes = round(seconds / 60)
    if minutes < 60:
        return f"{minutes} min"
    hours = minutes // 60
    rest = minutes % 60
    return f"{hours} h {rest} min" if rest else f"{hours} h"


def format_euro(amount: float | None) -> str:
    if amount is None:
        return "—"
    if amount == 0:
        return "0 €"
    text = f"{amount:.2f}".replace(".", ",")
    return f"{text} €"


def escape_html(text: Any) -> str:
    if text is None:
        return ""
    s = str(text)
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def render_stat_cards(cards: list[tuple[str, str, str]]) -> str:
    items = []
    for label, value, hint in cards:
        items.append(
            f'<article class="stat-card">'
            f'<p class="stat-card-label">{escape_html(label)}</p>'
            f'<p class="stat-card-value">{escape_html(value)}</p>'
            f'<p class="stat-card-hint">{escape_html(hint)}</p>'
            f"</article>"
        )
    return f'<section class="stats-cards">{"".join(items)}</section>'


def render_table(headers: list[str], rows: list[list[Any]]) -> str:
    head = "".join(f"<th>{escape_html(h)}</th>" for h in headers)
    body_rows = []
    for row in rows:
        cells = "".join(f"<td>{escape_html(c)}</td>" for c in row)
        body_rows.append(f"<tr>{cells}</tr>")
    return (
        '<div class="stats-table-wrap"><table class="stats-table">'
        f"<thead><tr>{head}</tr></thead>"
        f'<tbody>{"".join(body_rows)}</tbody></table></div>'
    )


def render_bar_list(items: list[tuple[str, int]], total: int) -> str:
    if not items:
        return "<p class='stats-empty'>Aucune donnee.</p>"
    max_count = max(count for _, count in items) or 1
    bars = []
    for label, count in items:
        width = max(4, round((count / max_count) * 100))
        pct = round((count / total) * 100) if total else 0
        bars.append(
            f'<div class="stats-bar-row">'
            f'<span class="stats-bar-label">{escape_html(label)}</span>'
            f'<div class="stats-bar-track"><div class="stats-bar-fill" style="width:{width}%"></div></div>'
            f'<span class="stats-bar-count">{count} ({pct}%)</span>'
            f"</div>"
        )
    return f'<div class="stats-bars">{"".join(bars)}</div>'


def render_section(title: str, body: str) -> str:
    return (
        f'<section class="stats-section">'
        f"<h2>{escape_html(title)}</h2>{body}</section>"
    )


def build_statistics(wb: openpyxl.Workbook, use_osrm: bool) -> dict[str, Any]:
    rows = [row_entry(item) for item in iter_activity_rows(wb)]
    cache_path = data_dir() / "route_stats_cache.json"
    cache = load_route_cache(cache_path)
    segments = build_segments(rows, cache, use_osrm)
    save_route_cache(cache_path, cache)

    geocoded = [r for r in rows if r["has_coords"]]
    trajet_lines = [r for r in rows if r["is_trajet_line"]]
    missing_coords = [r for r in rows if not r["has_coords"] and not r["is_trajet_line"]]

    calc_segments = [s for s in segments if s["calculable"]]
    foot_segments = [s for s in calc_segments if s["mode"] == "foot"]
    car_segments = [s for s in calc_segments if s["mode"] == "car"]
    foot_mode_count = sum(1 for s in segments if s["mode"] == "foot")
    car_mode_count = sum(1 for s in segments if s["mode"] == "car")
    non_calc_count = len(segments) - len(calc_segments)

    foot_distance = sum(s["distance_m"] for s in foot_segments)
    car_distance = sum(s["distance_m"] for s in car_segments)
    foot_duration = sum(s.get("duration_s") or 0 for s in foot_segments)
    car_duration = sum(s.get("duration_s") or 0 for s in car_segments)

    by_day: dict[int, dict[str, Any]] = defaultdict(
        lambda: {
            "activities": 0,
            "geocoded": 0,
            "foot_m": 0,
            "car_m": 0,
            "foot_min": 0,
            "car_min": 0,
            "prix": 0.0,
            "prix_count": 0,
        }
    )
    for row in rows:
        d = by_day[row["jour"]]
        d["activities"] += 1
        if row["has_coords"]:
            d["geocoded"] += 1
        if row["prix"] is not None:
            d["prix"] += row["prix"]
            d["prix_count"] += 1

    for seg in calc_segments:
        d = by_day[seg["jour"]]
        if seg["mode"] == "foot":
            d["foot_m"] += seg["distance_m"]
            d["foot_min"] += (seg.get("duration_s") or 0) / 60
        else:
            d["car_m"] += seg["distance_m"]
            d["car_min"] += (seg.get("duration_s") or 0) / 60

    by_ville_activities: Counter = Counter()
    by_ville_foot: defaultdict[str, int] = defaultdict(int)
    by_ville_car: defaultdict[str, int] = defaultdict(int)
    for row in rows:
        ville = row["ville"] or "Non renseignee"
        if not row["is_trajet_line"]:
            by_ville_activities[ville] += 1
    for seg in calc_segments:
        ville = seg["from_ville"] or seg["to_ville"] or "Non renseignee"
        if seg["mode"] == "foot":
            by_ville_foot[ville] += seg["distance_m"]
        else:
            by_ville_car[ville] += seg["distance_m"]

    by_action: Counter = Counter(r["action"] or "Non renseignee" for r in rows)
    by_type: Counter = Counter(r["type"] or "Non renseignee" for r in rows)
    by_billet: Counter = Counter(r["billet"] or "Non renseignee" for r in rows)

    prix_rows = [r for r in rows if r["prix"] is not None]
    prix_total = sum(r["prix"] for r in prix_rows)
    prix_visites = [
        r for r in prix_rows if r["action"].lower() in {"visite", "croisière", "croisiere"}
    ]
    prix_visites_total = sum(r["prix"] for r in prix_visites)

    villes = sorted({r["ville"] for r in rows if r["ville"] and not r["is_trajet_line"]})
    jours = sorted({r["jour"] for r in rows})

    longest_foot = max(foot_segments, key=lambda s: s["distance_m"], default=None)
    longest_car = max(car_segments, key=lambda s: s["distance_m"], default=None)
    most_walked_day = max(
        jours,
        key=lambda j: by_day[j]["foot_m"],
        default=None,
    )

    osrm_count = sum(1 for s in calc_segments if s.get("source") == "osrm")
    air_count = sum(1 for s in calc_segments if s.get("source") == "air")

    return {
        "generated_at": datetime.now(timezone.utc).strftime("%d/%m/%Y %H:%M UTC"),
        "summary": {
            "jours": len(jours),
            "activities": len(rows),
            "geocoded": len(geocoded),
            "on_map": len(geocoded),
            "trajet_lines": len(trajet_lines),
            "missing_coords": len(missing_coords),
            "villes": len(villes),
            "segments_total": len(segments),
            "segments_calculable": len(calc_segments),
            "segments_foot": len(foot_segments),
            "segments_car": len(car_segments),
            "segments_foot_mode": foot_mode_count,
            "segments_car_mode": car_mode_count,
            "segments_non_calculable": non_calc_count,
        },
        "distances": {
            "foot_m": foot_distance,
            "car_m": car_distance,
            "total_m": foot_distance + car_distance,
            "foot_duration_s": foot_duration,
            "car_duration_s": car_duration,
            "osrm_routes": osrm_count,
            "air_fallback": air_count,
        },
        "budget": {
            "total": prix_total,
            "entries": len(prix_rows),
            "visits_total": prix_visites_total,
            "visits_entries": len(prix_visites),
        },
        "by_day": {
            str(j): {
                "activities": by_day[j]["activities"],
                "geocoded": by_day[j]["geocoded"],
                "foot_km": round(by_day[j]["foot_m"] / 1000, 2),
                "car_km": round(by_day[j]["car_m"] / 1000, 2),
                "foot_min": round(by_day[j]["foot_min"]),
                "car_min": round(by_day[j]["car_min"]),
                "prix": round(by_day[j]["prix"], 2),
                "couleur": jour_color(j),
            }
            for j in jours
        },
        "by_ville": {
            ville: {
                "activities": by_ville_activities[ville],
                "foot_km": round(by_ville_foot.get(ville, 0) / 1000, 2),
                "car_km": round(by_ville_car.get(ville, 0) / 1000, 2),
            }
            for ville in sorted(by_ville_activities.keys())
        },
        "by_action": counter_dict(by_action),
        "by_type": counter_dict(by_type),
        "by_billet": counter_dict(by_billet),
        "trajet_lines": [
            {
                "jour": r["jour"],
                "ordre": r["ordre"],
                "nom": r["nom"],
                "ville": r["ville"],
                "ouverture": r["ouverture"],
                "fermeture": r["fermeture"],
                "billet": r["billet"],
            }
            for r in trajet_lines
        ],
        "missing_coords": [
            {"jour": r["jour"], "ordre": r["ordre"], "nom": r["nom"]}
            for r in missing_coords
        ],
        "highlights": {
            "longest_foot": (
                {
                    "jour": longest_foot["jour"],
                    "label": f"{longest_foot['from_ordre']} -> {longest_foot['to_ordre']}",
                    "distance_m": longest_foot["distance_m"],
                    "from_nom": longest_foot["from_nom"],
                    "to_nom": longest_foot["to_nom"],
                }
                if longest_foot
                else None
            ),
            "longest_car": (
                {
                    "jour": longest_car["jour"],
                    "label": f"{longest_car['from_ordre']} -> {longest_car['to_ordre']}",
                    "distance_m": longest_car["distance_m"],
                    "from_nom": longest_car["from_nom"],
                    "to_nom": longest_car["to_nom"],
                }
                if longest_car
                else None
            ),
            "most_walked_day": (
                {
                    "jour": most_walked_day,
                    "foot_km": round(by_day[most_walked_day]["foot_m"] / 1000, 2),
                }
                if most_walked_day is not None
                else None
            ),
        },
        "segments": segments,
        "villes": villes,
        "jours": jours,
    }


def render_html(stats: dict[str, Any]) -> str:
    s = stats["summary"]
    d = stats["distances"]
    b = stats["budget"]
    h = stats["highlights"]

    cards = render_stat_cards(
        [
            ("Jours de voyage", str(s["jours"]), f"{s['villes']} villes visitées"),
            ("Activités", str(s["activities"]), f"{s['geocoded']} géolocalisées"),
            (
                "Distance à pied",
                format_distance(d["foot_m"]),
                format_duration(d["foot_duration_s"]),
            ),
            (
                "Distance voiture",
                format_distance(d["car_m"]),
                format_duration(d["car_duration_s"]),
            ),
            (
                "Distance totale",
                format_distance(d["total_m"]),
                (
                    f"{s['segments_calculable']} seg. calcules "
                    f"({s['segments_foot']} pied, {s['segments_car']} voiture)"
                ),
            ),
            (
                "Budget renseigné",
                format_euro(b["total"]),
                f"{b['entries']} lignes · visites : {format_euro(b['visits_total'])}",
            ),
        ]
    )

    highlights_html = "<ul class='stats-list'>"
    if h["longest_foot"]:
        lf = h["longest_foot"]
        highlights_html += (
            f"<li>Plus long segment à pied (J{lf['jour']}) : "
            f"<strong>{format_distance(lf['distance_m'])}</strong> "
            f"({escape_html(lf['from_nom'])} → {escape_html(lf['to_nom'])})</li>"
        )
    if h["longest_car"]:
        lc = h["longest_car"]
        highlights_html += (
            f"<li>Plus long segment voiture (J{lc['jour']}) : "
            f"<strong>{format_distance(lc['distance_m'])}</strong> "
            f"({escape_html(lc['from_nom'])} → {escape_html(lc['to_nom'])})</li>"
        )
    if h["most_walked_day"]:
        mw = h["most_walked_day"]
        highlights_html += (
            f"<li>Jour le plus marché : <strong>Jour {mw['jour']}</strong> "
            f"({mw['foot_km']} km à pied)</li>"
        )
    highlights_html += "</ul>"

    day_rows = []
    for jour in stats["jours"]:
        dj = stats["by_day"][str(jour)]
        day_rows.append(
            [
                f"Jour {jour}",
                dj["activities"],
                dj["geocoded"],
                f"{dj['foot_km']} km",
                f"{dj['car_km']} km",
                f"{dj['foot_min']} min",
                f"{dj['car_min']} min",
                format_euro(dj["prix"]) if dj["prix"] else "—",
            ]
        )

    ville_rows = []
    for ville, dv in sorted(
        stats["by_ville"].items(),
        key=lambda x: x[1]["activities"],
        reverse=True,
    ):
        ville_rows.append(
            [
                ville,
                dv["activities"],
                f"{dv['foot_km']} km",
                f"{dv['car_km']} km",
            ]
        )

    trajet_rows = []
    for t in stats["trajet_lines"]:
        trajet_rows.append(
            [
                f"J{t['jour']} {t['ordre']}",
                t["nom"],
                t["ville"] or "—",
                f"{t['ouverture'] or '?'} - {t['fermeture'] or '?'}",
                t["billet"] or "—",
            ]
        )

    segment_rows = []
    for seg in stats["segments"]:
        if not seg["calculable"]:
            segment_rows.append(
                [
                    f"J{seg['jour']}",
                    f"{seg['from_ordre']}->{seg['to_ordre']}",
                    seg["mode"],
                    "—",
                    "—",
                    "Coords manquantes",
                ]
            )
            continue
        segment_rows.append(
            [
                f"J{seg['jour']}",
                f"{seg['from_ordre']}->{seg['to_ordre']}",
                "À pied" if seg["mode"] == "foot" else "Voiture",
                format_distance(seg["distance_m"]),
                format_duration(seg.get("duration_s")),
                seg.get("source", "—"),
            ]
        )

    missing_rows = [
        [f"J{m['jour']} {m['ordre']}", m["nom"]] for m in stats["missing_coords"]
    ]

    total_actions = sum(stats["by_action"].values())
    total_types = sum(stats["by_type"].values())

    content = (
        cards
        + render_section("Points saillants", highlights_html)
        + render_section(
            "Distances et itinéraires",
            "<p class='stats-note'>Distances calculees via OSRM (OpenStreetMap) "
            f"quand disponible, sinon a vol d'oiseau. "
            f"{d['osrm_routes']} itineraires OSRM, {d['air_fallback']} estimations directes. "
            f"{s['segments_non_calculable']} segments non calculables "
            f"(coordonnees manquantes sur {s['segments_car_mode']} segments voiture au total).</p>",
        )
        + render_section(
            "Répartition par jour",
            render_table(
                [
                    "Jour",
                    "Activités",
                    "Géoloc.",
                    "Pied",
                    "Voiture",
                    "Durée pied",
                    "Durée voit.",
                    "Budget",
                ],
                day_rows,
            ),
        )
        + render_section(
            "Répartition par ville",
            render_table(["Ville", "Activités", "Pied", "Voiture"], ville_rows),
        )
        + render_section(
            "Activités par nature",
            render_bar_list(
                list(stats["by_action"].items()),
                total_actions,
            ),
        )
        + render_section(
            "Activités par catégorie",
            render_bar_list(
                list(stats["by_type"].items()),
                total_types,
            ),
        )
        + render_section(
            "Réservations",
            render_bar_list(
                list(stats["by_billet"].items()),
                sum(stats["by_billet"].values()),
            ),
        )
        + render_section(
            "Trajets logistiques",
            render_table(
                ["Étape", "Trajet", "Ville", "Horaire", "Réservation"],
                trajet_rows or [["—", "Aucun", "—", "—", "—"]],
            ),
        )
        + render_section(
            "Tous les segments",
            render_table(
                ["Jour", "Segment", "Mode", "Distance", "Durée", "Source"],
                segment_rows,
            ),
        )
    )

    if missing_rows:
        content += render_section(
            "Lieux sans coordonnées",
            render_table(["Étape", "Lieu"], missing_rows),
        )

    return STATS_HTML_TEMPLATE.format(
        header=render_header("stats"),
        generated_at=escape_html(stats["generated_at"]),
        content=content,
    )


def run_build(excel_path: Path, use_osrm: bool = True) -> None:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    stats = build_statistics(wb, use_osrm=use_osrm)

    json_path = data_dir() / "stats.json"
    json_path.write_text(
        json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    html_path = web_dir() / "stats.html"
    html_path.write_text(render_html(stats), encoding="utf-8")

    d = stats["distances"]
    s = stats["summary"]
    print(f"Statistiques generees : {len(stats['segments'])} segments")
    print(f"  Activites : {s['activities']} ({s['geocoded']} geolocalisees)")
    print(f"  A pied : {format_distance(d['foot_m'])} ({format_duration(d['foot_duration_s'])})")
    print(f"  Voiture : {format_distance(d['car_m'])} ({format_duration(d['car_duration_s'])})")
    print(f"  JSON : {json_path}")
    print(f"  HTML : {html_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Genere la page de statistiques du voyage."
    )
    parser.add_argument("excel", nargs="?", default=str(default_excel_path()))
    parser.add_argument(
        "--no-osrm",
        action="store_true",
        help="N'appelle pas OSRM (distances a vol d'oiseau uniquement).",
    )
    args = parser.parse_args()

    excel_path = Path(args.excel).resolve()
    if not excel_path.exists():
        raise SystemExit(f"Fichier introuvable: {excel_path}")

    run_build(excel_path, use_osrm=not args.no_osrm)


if __name__ == "__main__":
    main()
