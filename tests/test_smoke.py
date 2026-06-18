"""Tests de fumee : imports, chemins et utilitaires sans effet de bord."""

from __future__ import annotations

import importlib
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"

if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


@pytest.fixture
def root() -> Path:
    return ROOT


def test_project_layout(root: Path) -> None:
    assert (root / "generer_site.py").is_file()
    assert (root / "preparer_excel.py").is_file()
    assert (root / "qualite.py").is_file()
    assert (root / "scripts" / "pipeline_common.py").is_file()
    assert (root / "pyproject.toml").is_file()


@pytest.mark.parametrize(
    "module_name",
    [
        "generer_site",
        "preparer_excel",
        "qualite",
        "pipeline_common",
        "outils.excel_utils",
        "outils.drive_config",
        "outils.overview_config",
        "site_web.site_nav",
    ],
)
def test_import_core_modules(module_name: str) -> None:
    _ = importlib.import_module(module_name)


def test_pipeline_common_root(root: Path) -> None:
    from pipeline_common import ROOT, SCRIPTS

    assert ROOT == root
    assert SCRIPTS == root / "scripts"


def test_default_excel_path_under_root(root: Path) -> None:
    from outils.excel_utils import default_excel_path

    path = default_excel_path()
    assert path.is_absolute()
    assert path.parent == root / "excel"


def test_parse_ordre_rejects_excel_date() -> None:
    from datetime import datetime

    from outils.excel_utils import is_ordre_read_as_date, parse_ordre

    dt = datetime(2026, 8, 1, 0, 0)
    assert is_ordre_read_as_date(dt)
    assert parse_ordre(dt) is None
    assert parse_ordre("1.1") == {"jour": 1, "visite": 1}


def test_ordre_cell_format_issue_detects_date_format() -> None:
    from outils.excel_utils import ordre_cell_format_issue

    issue = ordre_cell_format_issue(
        "Jour 1",
        3,
        value="1.1",
        number_format="d/m",
    )
    assert issue is not None
    assert "format date" in issue

    assert (
        ordre_cell_format_issue(
            "Jour 1",
            3,
            value="1.1",
            number_format="@",
        )
        is None
    )
