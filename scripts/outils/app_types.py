"""Types partages (config JSON, cellules Excel, structures metier)."""

from __future__ import annotations

from datetime import date, datetime
from typing import NotRequired, TypedDict

ExcelCellValue = str | int | float | bool | datetime | date | None
ExcelRow = tuple[object, ...]


class DriveConfig(TypedDict):
    source_path: str


class OverviewConfig(TypedDict):
    start_date: str
    sheet_name: str
    verify_markers: list[str]
    day_resume_limit: int
    write_snapshot: bool
    domicile: NotRequired[str]
    banner_title: NotRequired[str]
    intro: NotRequired[str]
    notes: NotRequired[list[str]]


class LodgingFields(TypedDict):
    jour: int
    visite: int
    nom: str
    ville: str
    action: str


class ActivityFields(LodgingFields):
    ordre: str


class OverviewRow(LodgingFields):
    type: str
    prix: float | None
    remarque: str
    ouverture: str
    fermeture: str
    lat: float | None
    lon: float | None
    lien: str
    site: str
    is_trajet_line: bool


class LodgingVilleBucket(TypedDict):
    ville: str
    nom: str
    jours: set[int]
    lat: float | None
    lon: float | None
    lien: str


class MapPopup(TypedDict, total=False):
    action: str | None
    type: str | None
    billet: str | None
    prix: float | None
    city_card: str | None
    ouverture: str | None
    fermeture: str | None
    remarque: str | None


class MapPoint(TypedDict):
    id: str
    ordre: int
    ordre_label: str
    jour: int
    visite: int
    nom: str
    ville: str | None
    lat: float
    lon: float
    lien: str | None
    couleur: str
    popup: MapPopup


class VoyageData(TypedDict):
    jours: list[int]
    points: list[MapPoint]


class GeocodeCacheEntry(TypedDict, total=False):
    lat: float
    lon: float
    source: str
    query: str


class StatsRow(TypedDict):
    jour: int
    visite: int
    ordre: str
    nom: str
    ville: str
    action: str
    type: str
    billet: str
    prix: float | None
    ouverture: ExcelCellValue
    fermeture: ExcelCellValue
    remarque: str
    has_coords: bool
    lat: float | None
    lon: float | None
    couleur: str
    is_trajet_line: bool


class RouteCacheEntry(TypedDict):
    distance_m: int | float
    duration_s: int | float | None
    source: str


class ResolvedRoute(RouteCacheEntry, total=False):
    cached: bool


class OsrmRoute(TypedDict):
    distance_m: int | float
    duration_s: int | float
    source: str


class TravelSegment(TypedDict):
    jour: int
    from_ordre: str
    to_ordre: str
    from_nom: str
    to_nom: str
    from_ville: str
    to_ville: str
    mode: str
    calculable: bool
    distance_m: NotRequired[int | float]
    duration_s: NotRequired[int | float | None]
    source: NotRequired[str]
    air_m: NotRequired[int]


class DayStatsAccumulator(TypedDict):
    activities: int
    geocoded: int
    foot_m: int | float
    car_m: int | float
    foot_min: int | float
    car_min: int | float
    prix: float
    prix_count: int


class DayStatsPublic(TypedDict):
    activities: int
    geocoded: int
    foot_km: float
    car_km: float
    foot_min: int | float
    car_min: int | float
    prix: float
    couleur: str


class VilleStats(TypedDict):
    activities: int
    foot_km: float
    car_km: float


class StatsSummary(TypedDict):
    jours: int
    activities: int
    geocoded: int
    on_map: int
    trajet_lines: int
    missing_coords: int
    villes: int
    segments_total: int
    segments_calculable: int
    segments_foot: int
    segments_car: int
    segments_foot_mode: int
    segments_car_mode: int
    segments_non_calculable: int


class StatsDistances(TypedDict):
    foot_m: int | float
    car_m: int | float
    total_m: int | float
    foot_duration_s: int | float
    car_duration_s: int | float
    osrm_routes: int
    air_fallback: int


class StatsBudget(TypedDict):
    total: int | float
    entries: int
    visits_total: int | float
    visits_entries: int


class SegmentHighlight(TypedDict):
    jour: int
    label: str
    distance_m: int | float
    from_nom: str
    to_nom: str


class WalkedDayHighlight(TypedDict):
    jour: int
    foot_km: float


class StatsHighlights(TypedDict):
    longest_foot: SegmentHighlight | None
    longest_car: SegmentHighlight | None
    most_walked_day: WalkedDayHighlight | None


class TrajetLineStat(TypedDict):
    jour: int
    ordre: str
    nom: str
    ville: str
    ouverture: ExcelCellValue
    fermeture: ExcelCellValue
    billet: str


class MissingCoordStat(TypedDict):
    jour: int
    ordre: str
    nom: str


class StatsPayload(TypedDict):
    generated_at: str
    summary: StatsSummary
    distances: StatsDistances
    budget: StatsBudget
    by_day: dict[str, DayStatsPublic]
    by_ville: dict[str, VilleStats]
    by_action: dict[str, int]
    by_type: dict[str, int]
    by_billet: dict[str, int]
    trajet_lines: list[TrajetLineStat]
    missing_coords: list[MissingCoordStat]
    highlights: StatsHighlights
    segments: list[TravelSegment]
    villes: list[str]
    jours: list[int]


class OverviewDaySummary(TypedDict):
    jour: int
    date: date
    ville: str
    villes: list[str]
    activities: int
    visites: int
    prix: float
    resume: str
    couleur: str
    theme: NotRequired[str]
    lodging_villes_label: NotRequired[str]
    highlight: NotRequired[bool]
    nuit: NotRequired[dict[str, str] | None]


class OverviewPhaseData(TypedDict):
    ville: str
    start: date
    end: date
    jours: str
    from_jour: int
    to_jour: int


class OverviewTripStep(TypedDict):
    ville: str
    dates: str
    description: str


class OverviewLodgingRow(TypedDict):
    ville: str
    nom: str
    dates: str
    start: date
    maps_url: NotRequired[str | None]


class OverviewVilleRow(TypedDict):
    ville: str
    activities: int
    visites: int
    prix: int | float
    jours: int


class OverviewVilleAccumulator(TypedDict):
    activities: int
    visites: int
    prix: float
    jours: set[int]


class OverviewCollectedData(TypedDict):
    banner_title: str
    generated_at: str
    period: str
    route: str
    jours_count: int
    activities_total: int
    visites_total: int
    prix_total: int | float
    day_summaries: list[OverviewDaySummary]
    phases: list[OverviewPhaseData]
    trip_steps: list[OverviewTripStep]
    lodging_rows: list[OverviewLodgingRow]
    ville_rows: list[OverviewVilleRow]


class OverviewSnapshotSummary(TypedDict, total=False):
    period: str
    route: str
    jours_count: int
    activities_total: int
    visites_total: int
    prix_total: int | float


class OverviewDaySnapshot(TypedDict):
    jour: int
    date: NotRequired[str]
    ville: NotRequired[str]
    resume: NotRequired[str]
    nuit: NotRequired[str]


class OverviewPhase(TypedDict):
    from_jour: int
    to_jour: int
    ville: NotRequired[str]


class OverviewPayload(TypedDict, total=False):
    generated_at: str
    summary: OverviewSnapshotSummary
    by_day: list[OverviewDaySnapshot]
    by_ville: list[dict[str, object]]
    phases: list[OverviewPhase]


class FileSourceMeta(TypedDict):
    path: str
    present: bool
    mtime: str | None
    size: int


class InspectCheck(TypedDict):
    id: str
    status: str
    message: str
    details: str


class InspectActivity(TypedDict):
    key: str
    jour: int
    ordre: str
    nom: str
    ville: str
    action: str
    prix: float | None
    on_map: bool
    in_stats: bool
    is_trajet: bool
    in_overview: NotRequired[bool]


class InspectDayTimeline(TypedDict):
    jour: int
    date: str | None
    ville: str
    resume: str
    activities: int | float
    geocoded: int | float
    prix: int | float
    foot_km: int | float
    car_km: int | float
    couleur: str
    in_overview: bool


class InspectCoverageRow(TypedDict):
    jour: int
    ordre: str
    nom: str
    ville: str
    in_stats: bool
    on_map: bool
    in_overview: bool | None
    is_trajet: bool
    status: str


class InspectMapPoint(TypedDict):
    jour: int
    ordre: str
    nom: str
    ville: str | None
    lat: float
    lon: float
    couleur: str


class InspectSummary(TypedDict):
    title: str
    period: str
    route: str
    activities: int
    on_map: int
    budget: int | float
    jours: int


class InspectConfigSnippet(TypedDict):
    start_date: str | None
    domicile: str | None
    verify_markers: list[str]


class InspectSources(TypedDict):
    overview_config: FileSourceMeta
    overview: FileSourceMeta
    voyages: FileSourceMeta
    stats: FileSourceMeta
    geocode_cache: FileSourceMeta
    route_stats_cache: FileSourceMeta
    geocode_errors: FileSourceMeta
    missing_coords: FileSourceMeta


class InspectPayload(TypedDict):
    generated_at: str
    overall_status: str
    has_overview: bool
    summary: InspectSummary
    sources: InspectSources
    checks: list[InspectCheck]
    days: list[InspectDayTimeline]
    activities: list[InspectActivity]
    coverage: list[InspectCoverageRow]
    anomalies: list[InspectCoverageRow]
    geocode_errors: list[dict[str, str]]
    missing_coords: list[dict[str, str]]
    config: InspectConfigSnippet
    lodging_audit: dict[str, object] | None
    map_points: list[InspectMapPoint]
