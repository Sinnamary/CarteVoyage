#!/usr/bin/env python3
"""Regenere la feuille Excel Vue d'ensemble depuis les feuilles Jour N."""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from outils.excel_utils import (
    backup_excel,
    cell_value,
    data_dir,
    default_excel_path,
    iter_activity_rows,
    jour_color,
    normalize_text,
    parse_prix,
)
from outils.overview_config import default_overview_config_path, resolve_overview_config

MOIS_FR = (
    "",
    "janvier",
    "février",
    "mars",
    "avril",
    "mai",
    "juin",
    "juillet",
    "août",
    "septembre",
    "octobre",
    "novembre",
    "décembre",
)

HEADER_FILL = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=16)
SUBTITLE_FONT = Font(bold=True, size=12)
BOLD_FONT = Font(bold=True)


def is_trajet_line(nom: str) -> bool:
    lower = normalize_text(nom).lower()
    return lower.startswith(("trajet ", "retour "))


def parse_start_date(value: str) -> date:
    match = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", value.strip())
    if not match:
        raise ValueError(f"Date invalide {value!r} (attendu AAAA-MM-JJ)")
    return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))


def format_date_fr(d: date) -> str:
    return f"{d.day} {MOIS_FR[d.month]} {d.year}"


def format_date_range_fr(start: date, end: date) -> str:
    if start == end:
        return format_date_fr(start)
    if start.year == end.year and start.month == end.month:
        return f"{start.day} au {end.day} {MOIS_FR[start.month]} {start.year}"
    if start.year == end.year:
        return f"{start.day} {MOIS_FR[start.month]} au {end.day} {MOIS_FR[end.month]} {start.year}"
    return f"{format_date_fr(start)} au {format_date_fr(end)}"


def date_for_jour(start: date, jour: int) -> date:
    return start + timedelta(days=jour - 1)


def hex_to_fill(hex_color: str) -> PatternFill:
    color = hex_color.lstrip("#").upper()
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def row_entry(item: dict[str, Any]) -> dict[str, Any]:
    row = item["row"]
    ci = item["col_index"]
    nom = item["nom"]
    return {
        "jour": item["jour"],
        "visite": item["visite"],
        "nom": nom,
        "ville": normalize_text(cell_value(row, ci, "Ville")),
        "action": normalize_text(cell_value(row, ci, "Action")),
        "type": normalize_text(cell_value(row, ci, "Type")),
        "prix": parse_prix(cell_value(row, ci, "Prix")),
        "is_trajet_line": is_trajet_line(nom),
    }


def primary_ville(rows: list[dict[str, Any]]) -> str:
    counts: Counter[str] = Counter()
    first_order: dict[str, int] = {}
    for row in rows:
        if row["is_trajet_line"] or not row["ville"]:
            continue
        counts[row["ville"]] += 1
        first_order.setdefault(row["ville"], row["visite"])
    if not counts:
        return ""
    return max(counts.keys(), key=lambda v: (counts[v], -first_order[v]))


def villes_for_day(rows: list[dict[str, Any]]) -> list[str]:
    seen: list[str] = []
    for row in sorted(rows, key=lambda r: r["visite"]):
        ville = row["ville"]
        if row["is_trajet_line"] or not ville or ville in seen:
            continue
        seen.append(ville)
    return seen


def day_resume(rows: list[dict[str, Any]], limit: int = 3) -> str:
    visites = [
        row["nom"]
        for row in sorted(rows, key=lambda r: r["visite"])
        if not row["is_trajet_line"] and row["action"].lower() == "visite"
    ]
    if visites:
        head = ", ".join(visites[:limit])
        extra = len(visites) - limit
        return head + (f" (+{extra})" if extra > 0 else "")
    actions = Counter(
        row["action"] or "Activité"
        for row in rows
        if not row["is_trajet_line"]
    )
    if not actions:
        return "Trajets / logistique"
    return ", ".join(f"{n} {label.lower()}" for label, n in actions.most_common(3))


def build_phases_from_config(
    config_phases: list[dict[str, Any]],
    start_date: date,
) -> list[dict[str, Any]]:
    phases: list[dict[str, Any]] = []
    for entry in config_phases:
        from_jour = int(entry["from_jour"])
        to_jour = int(entry["to_jour"])
        label = normalize_text(entry.get("label")) or normalize_text(entry.get("ville"))
        phases.append(
            {
                "ville": label,
                "start": date_for_jour(start_date, from_jour),
                "end": date_for_jour(start_date, to_jour),
                "jours": f"Jour {from_jour}–{to_jour}",
            }
        )
    return phases


def build_phases_auto(day_summaries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    phases: list[dict[str, Any]] = []
    if not day_summaries:
        return phases

    phase_start_idx = 0
    prev_ville = day_summaries[0]["ville"]
    for idx, summary in enumerate(day_summaries[1:], start=1):
        if summary["ville"] != prev_ville:
            start = day_summaries[phase_start_idx]
            end = day_summaries[idx - 1]
            phases.append(
                {
                    "ville": prev_ville,
                    "start": start["date"],
                    "end": end["date"],
                    "jours": f"Jour {start['jour']}–{end['jour']}",
                }
            )
            phase_start_idx = idx
            prev_ville = summary["ville"]
    start = day_summaries[phase_start_idx]
    end = day_summaries[-1]
    phases.append(
        {
            "ville": prev_ville,
            "start": start["date"],
            "end": end["date"],
            "jours": f"Jour {start['jour']}–{end['jour']}",
        }
    )
    return phases


def collect_overview_data(
    wb: openpyxl.Workbook,
    config: dict[str, Any],
) -> dict[str, Any]:
    start_date = parse_start_date(config["start_date"])
    trip_title = config["title"]
    resume_limit = int(config.get("day_resume_limit") or 3)

    rows = [row_entry(item) for item in iter_activity_rows(wb)]
    by_jour: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        by_jour[row["jour"]].append(row)

    jours = sorted(by_jour)
    day_summaries: list[dict[str, Any]] = []
    route_cities: list[str] = []

    for jour in jours:
        day_rows = by_jour[jour]
        main_ville = primary_ville(day_rows)
        day_villes = villes_for_day(day_rows)
        if main_ville and (not route_cities or route_cities[-1] != main_ville):
            route_cities.append(main_ville)

        activities = [r for r in day_rows if not r["is_trajet_line"]]
        visites = [r for r in activities if r["action"].lower() == "visite"]
        prix_rows = [r for r in day_rows if r["prix"] is not None]
        prix_total = sum(r["prix"] for r in prix_rows)

        day_summaries.append(
            {
                "jour": jour,
                "date": date_for_jour(start_date, jour),
                "ville": main_ville,
                "villes": day_villes,
                "activities": len(activities),
                "visites": len(visites),
                "prix": prix_total,
                "resume": day_resume(day_rows, limit=resume_limit),
                "couleur": jour_color(jour),
            }
        )

    config_phases = config.get("phases") or []
    if config_phases:
        phases = build_phases_from_config(config_phases, start_date)
    else:
        phases = build_phases_auto(day_summaries)

    by_ville: dict[str, dict[str, Any]] = defaultdict(
        lambda: {"activities": 0, "visites": 0, "prix": 0.0, "jours": set()}
    )
    for row in rows:
        if row["is_trajet_line"] or not row["ville"]:
            continue
        entry = by_ville[row["ville"]]
        entry["activities"] += 1
        entry["jours"].add(row["jour"])
        if row["action"].lower() == "visite":
            entry["visites"] += 1
        if row["prix"] is not None:
            entry["prix"] += row["prix"]

    ville_rows = []
    for ville in sorted(by_ville, key=lambda v: min(by_ville[v]["jours"])):
        entry = by_ville[ville]
        ville_rows.append(
            {
                "ville": ville,
                "activities": entry["activities"],
                "visites": entry["visites"],
                "prix": entry["prix"],
                "jours": len(entry["jours"]),
            }
        )

    prix_total = sum(r["prix"] for r in rows if r["prix"] is not None)
    visites_total = sum(1 for r in rows if r["action"].lower() == "visite")
    activities_total = sum(1 for r in rows if not r["is_trajet_line"])

    if jours:
        period = format_date_range_fr(
            date_for_jour(start_date, jours[0]),
            date_for_jour(start_date, jours[-1]),
        )
    else:
        period = ""

    route = normalize_text(config.get("route")) or " → ".join(route_cities)

    return {
        "title": trip_title,
        "intro": normalize_text(config.get("intro")),
        "notes": [normalize_text(note) for note in (config.get("notes") or []) if normalize_text(note)],
        "sections": config.get("sections") or {},
        "generated_at": datetime.now(timezone.utc).strftime("%d/%m/%Y %H:%M UTC"),
        "period": period,
        "route": route,
        "jours_count": len(jours),
        "activities_total": activities_total,
        "visites_total": visites_total,
        "prix_total": prix_total,
        "day_summaries": day_summaries,
        "phases": phases,
        "ville_rows": ville_rows,
    }


def json_default(value: Any) -> Any:
    if isinstance(value, date):
        return value.isoformat()
    raise TypeError(f"Type non serialisable: {type(value)!r}")


def write_overview_snapshot(data: dict[str, Any], config: dict[str, Any]) -> Path:
    snapshot = {
        "generated_at": data["generated_at"],
        "config": {
            "title": config["title"],
            "start_date": config["start_date"],
            "sheet_name": config["sheet_name"],
        },
        "summary": {
            "period": data["period"],
            "route": data["route"],
            "jours_count": data["jours_count"],
            "activities_total": data["activities_total"],
            "visites_total": data["visites_total"],
            "prix_total": data["prix_total"],
        },
        "phases": data["phases"],
        "by_day": data["day_summaries"],
        "by_ville": data["ville_rows"],
    }
    path = data_dir() / "overview.json"
    path.write_text(
        json.dumps(snapshot, ensure_ascii=False, indent=2, default=json_default),
        encoding="utf-8",
    )
    return path


def set_cell(
    ws: Worksheet,
    row: int,
    col: int,
    value: Any,
    *,
    font: Font | None = None,
    fill: PatternFill | None = None,
    alignment: Alignment | None = None,
) -> None:
    cell = ws.cell(row=row, column=col, value=value)
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment


def write_table_header(ws: Worksheet, row: int, headers: list[str]) -> None:
    for col, header in enumerate(headers, start=1):
        set_cell(
            ws,
            row,
            col,
            header,
            font=HEADER_FONT,
            fill=HEADER_FILL,
            alignment=Alignment(horizontal="center"),
        )


def autosize_columns(ws: Worksheet, max_col: int, min_width: int = 10, max_width: int = 48) -> None:
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        max_len = min_width
        for cell in ws[letter]:
            if cell.value is None:
                continue
            max_len = max(max_len, min(len(str(cell.value)), max_width))
        ws.column_dimensions[letter].width = max_len + 2


def clear_sheet(ws: Worksheet) -> None:
    if ws.max_row:
        ws.delete_rows(1, ws.max_row)


def ensure_overview_sheet(wb: openpyxl.Workbook, sheet_name: str) -> Worksheet:
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    return wb.create_sheet(sheet_name, 0)


def section_enabled(sections: dict[str, Any], name: str, default: bool = True) -> bool:
    value = sections.get(name, default)
    return bool(value)


def render_overview_sheet(ws: Worksheet, data: dict[str, Any]) -> None:
    clear_sheet(ws)
    sections = data.get("sections") or {}
    row = 1
    max_col = 8

    set_cell(ws, row, 1, data["title"], font=TITLE_FONT)
    row += 1
    set_cell(ws, row, 1, f"Période : {data['period']} ({data['jours_count']} jours)", font=SUBTITLE_FONT)
    row += 1
    if data["intro"]:
        set_cell(ws, row, 1, data["intro"])
        row += 1
    if data["route"]:
        set_cell(ws, row, 1, f"Itinéraire : {data['route']}")
        row += 1
    if section_enabled(sections, "totals", True):
        set_cell(
            ws,
            row,
            1,
            f"{data['activities_total']} activités · {data['visites_total']} visites · "
            f"{data['prix_total']:.2f} € budget renseigné",
        )
        row += 1
    set_cell(ws, row, 1, f"Généré le {data['generated_at']}")
    row += 1

    for note in data.get("notes") or []:
        set_cell(ws, row, 1, f"• {note}")
        row += 1
    row += 1

    if section_enabled(sections, "phases", True) and data["phases"]:
        set_cell(ws, row, 1, "Phases du voyage", font=SUBTITLE_FONT)
        row += 1
        write_table_header(ws, row, ["Phase", "Jours", "Période"])
        row += 1
        for phase in data["phases"]:
            set_cell(ws, row, 1, phase["ville"])
            set_cell(ws, row, 2, phase["jours"])
            set_cell(ws, row, 3, format_date_range_fr(phase["start"], phase["end"]))
            row += 1
        row += 1

    header_row = row
    if section_enabled(sections, "by_day", True):
        set_cell(ws, row, 1, "Planning par jour", font=SUBTITLE_FONT)
        row += 1
        day_headers = [
            "Jour",
            "Date",
            "Ville principale",
            "Autres villes",
            "Activités",
            "Visites",
            "Budget (€)",
            "Résumé",
        ]
        write_table_header(ws, row, day_headers)
        header_row = row
        row += 1

        for summary in data["day_summaries"]:
            autres = ", ".join(v for v in summary["villes"] if v != summary["ville"])
            day_fill = hex_to_fill(summary["couleur"])
            set_cell(ws, row, 1, summary["jour"], fill=day_fill, font=Font(bold=True, color="FFFFFF"))
            set_cell(ws, row, 2, format_date_fr(summary["date"]))
            set_cell(ws, row, 3, summary["ville"])
            set_cell(ws, row, 4, autres)
            set_cell(ws, row, 5, summary["activities"], alignment=Alignment(horizontal="center"))
            set_cell(ws, row, 6, summary["visites"], alignment=Alignment(horizontal="center"))
            prix = summary["prix"]
            set_cell(
                ws,
                row,
                7,
                round(prix, 2) if prix else "",
                alignment=Alignment(horizontal="right"),
            )
            set_cell(ws, row, 8, summary["resume"])
            row += 1

        if section_enabled(sections, "totals", True):
            set_cell(ws, row, 1, "Total", font=BOLD_FONT)
            set_cell(ws, row, 5, data["activities_total"], font=BOLD_FONT, alignment=Alignment(horizontal="center"))
            set_cell(ws, row, 6, data["visites_total"], font=BOLD_FONT, alignment=Alignment(horizontal="center"))
            set_cell(
                ws,
                row,
                7,
                round(data["prix_total"], 2) if data["prix_total"] else "",
                font=BOLD_FONT,
                alignment=Alignment(horizontal="right"),
            )
            row += 1
        row += 1

    if section_enabled(sections, "by_ville", True):
        set_cell(ws, row, 1, "Répartition par ville", font=SUBTITLE_FONT)
        row += 1
        ville_headers = ["Ville", "Jours", "Activités", "Visites", "Budget (€)"]
        write_table_header(ws, row, ville_headers)
        row += 1
        for ville_row in data["ville_rows"]:
            set_cell(ws, row, 1, ville_row["ville"])
            set_cell(ws, row, 2, ville_row["jours"], alignment=Alignment(horizontal="center"))
            set_cell(ws, row, 3, ville_row["activities"], alignment=Alignment(horizontal="center"))
            set_cell(ws, row, 4, ville_row["visites"], alignment=Alignment(horizontal="center"))
            prix = ville_row["prix"]
            set_cell(
                ws,
                row,
                5,
                round(prix, 2) if prix else "",
                alignment=Alignment(horizontal="right"),
            )
            row += 1
        max_col = max(max_col, len(ville_headers))

    if section_enabled(sections, "by_day", True):
        ws.freeze_panes = f"A{header_row + 1}"
    autosize_columns(ws, max_col=max_col)


def run_build(
    excel_path: Path,
    config: dict[str, Any],
    *,
    config_path: Path | None = None,
    dry_run: bool = False,
) -> None:
    wb = openpyxl.load_workbook(excel_path)
    data = collect_overview_data(wb, config)
    sheet_name = config["sheet_name"]

    if dry_run:
        print(f"(dry-run) Vue d'ensemble pour {excel_path.name}")
        print(f"  Config : {config_path or default_overview_config_path()}")
        print(f"  Période : {data['period']}")
        print(f"  Itinéraire : {data['route']}")
        print(f"  {len(data['day_summaries'])} jour(s), {data['activities_total']} activité(s)")
        return

    ws = ensure_overview_sheet(wb, sheet_name)
    render_overview_sheet(ws, data)
    backup_path = backup_excel(excel_path)
    wb.save(excel_path)

    snapshot_path = None
    if config.get("write_snapshot", True):
        snapshot_path = write_overview_snapshot(data, config)

    print(f"Vue d'ensemble regeneree : {excel_path}")
    print(f"  Feuille : {sheet_name}")
    print(f"  Période : {data['period']}")
    print(f"  Itinéraire : {data['route']}")
    print(f"  Backup : {backup_path}")
    if snapshot_path:
        print(f"  Snapshot : {snapshot_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Regenere la feuille Vue d'ensemble depuis les feuilles Jour N."
    )
    parser.add_argument("excel", nargs="?", default=str(default_excel_path()))
    parser.add_argument(
        "--config",
        default=str(default_overview_config_path()),
        help="Fichier JSON de configuration (defaut: data/overview_config.json).",
    )
    parser.add_argument(
        "--start-date",
        default=None,
        help="Surcharge la date du jour 1 (AAAA-MM-JJ).",
    )
    parser.add_argument(
        "--title",
        default=None,
        help="Surcharge le titre affiche en tete de la feuille.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Afficher sans ecrire dans Excel")
    args = parser.parse_args()

    excel_path = Path(args.excel).resolve()
    if not excel_path.exists():
        raise SystemExit(f"Fichier introuvable: {excel_path}")

    config_path = Path(args.config).resolve()
    try:
        config = resolve_overview_config(
            config_path,
            title=args.title,
            start_date=args.start_date,
        )
        parse_start_date(config["start_date"])
    except (ValueError, json.JSONDecodeError) as exc:
        raise SystemExit(str(exc)) from exc

    run_build(excel_path, config, config_path=config_path, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
