"""Namespaces argparse types pour basedpyright."""

from __future__ import annotations

import argparse


class ExcelPathArgs(argparse.Namespace):
    excel: str


class SyncListesArgs(ExcelPathArgs):
    dry_run: bool


class SyncFromDriveArgs(argparse.Namespace):
    dest: str
    source: str | None
    config: str
    dry_run: bool
    no_backup: bool


class SyncToDriveArgs(argparse.Namespace):
    source: str
    dest: str | None
    config: str
    dry_run: bool
    no_backup: bool
    backup_drive: bool


class VerifyWorkbookArgs(ExcelPathArgs):
    pass


class GeocodeArgs(ExcelPathArgs):
    dry_run: bool
    force: bool


class BuildMapArgs(ExcelPathArgs):
    pass


class BuildStatsArgs(ExcelPathArgs):
    no_osrm: bool


class BuildOverviewArgs(ExcelPathArgs):
    config: str
    start_date: str | None
    dry_run: bool
    snapshot_only: bool


class BuildInspectArgs(ExcelPathArgs):
    config: str


class GenererSiteArgs(argparse.Namespace):
    excel: str | None
    drive_pull: bool
    no_osrm: bool


class PreparerExcelArgs(argparse.Namespace):
    excel: str | None
    skip_drive_pull: bool
    skip_drive_push: bool
    skip_sync: bool
    skip_overview: bool
    skip_verify: bool
    skip_geocode: bool
    geocode_force: bool
