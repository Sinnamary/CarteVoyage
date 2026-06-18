#!/usr/bin/env python3
"""Orchestrateur qualite code pour CarteVoyage.

Lance une batterie d'outils Python modernes (ruff, vulture, bandit, mypy,
basedpyright, pytest, radon, interrogate, pip-audit, deptry, compileall).

Usage :
  python qualite.py              # toutes les verifications
  python qualite.py --fix        # corrige format + lint auto-fixable
  python qualite.py --install    # installe requirements-dev.txt
  python qualite.py --list       # liste les controles disponibles
  python qualite.py --only ruff,pytest
  python qualite.py --skip radon,interrogate
"""

from __future__ import annotations

import argparse
import importlib.util
import shutil
import subprocess
import sys
import textwrap
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path


class QualiteArgs(argparse.Namespace):
    fix: bool
    install: bool
    list: bool
    only: str | None
    skip: str | None
    strict: bool
    verbose: bool


ROOT = Path(__file__).resolve().parent
PYTHON = sys.executable
SOURCE_TARGETS = ("scripts", "generer_site.py", "preparer_excel.py", "qualite.py")
REQUIREMENTS_DEV = ROOT / "requirements-dev.txt"


@dataclass
class Check:
    """Un controle qualite."""

    id: str
    label: str
    category: str
    runner: Callable[[QualiteArgs], tuple[bool, str]]
    optional: bool = False
    warn_only: bool = False


def _run(
    cmd: Sequence[str],
    *,
    cwd: Path = ROOT,
    allow_missing: bool = False,
) -> tuple[int, str]:
    try:
        proc = subprocess.run(
            list(cmd),
            cwd=cwd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except FileNotFoundError:
        if allow_missing:
            return 127, f"commande introuvable : {cmd[0]}"
        raise
    out = (proc.stdout or "") + (proc.stderr or "")
    return proc.returncode, out.strip()


def _module_available(name: str) -> bool:
    return importlib.util.find_spec(name) is not None


def _tool_version(module: str) -> str:
    if not _module_available(module):
        return "absent"
    code, out = _run([PYTHON, "-m", module, "--version"], allow_missing=True)
    if code == 0 and out:
        return out.splitlines()[0]
    return "installe"


def _source_args() -> list[str]:
    return [p for p in SOURCE_TARGETS if (ROOT / p).exists()]


def _ensure_dev_tools(_: QualiteArgs) -> tuple[bool, str]:
    missing = [
        mod
        for mod in (
            "ruff",
            "vulture",
            "bandit",
            "pytest",
            "radon",
        )
        if not _module_available(mod)
    ]
    if missing:
        return False, (
            "Outils manquants : " + ", ".join(missing) + f". Lancez : {PYTHON} qualite.py --install"
        )
    return True, "Outils principaux disponibles"


def _check_compileall(_: QualiteArgs) -> tuple[bool, str]:
    code, out = _run(
        [PYTHON, "-m", "compileall", "-q", "-j", "0", *_source_args()],
        allow_missing=True,
    )
    return code == 0, out or "Syntaxe Python OK"


def _check_ruff_format(args: QualiteArgs) -> tuple[bool, str]:
    cmd = [PYTHON, "-m", "ruff", "format"]
    if not args.fix:
        cmd.append("--check")
    cmd.extend(_source_args())
    code, out = _run(cmd, allow_missing=True)
    if code == 127:
        return False, out
    return code == 0, out or ("Format applique" if args.fix else "Format OK")


def _check_ruff_lint(args: QualiteArgs) -> tuple[bool, str]:
    cmd = [PYTHON, "-m", "ruff", "check", *_source_args()]
    if args.fix:
        cmd.append("--fix")
    code, out = _run(cmd, allow_missing=True)
    if code == 127:
        return False, out
    return code == 0, out or "Lint OK"


def _check_vulture(_: QualiteArgs) -> tuple[bool, str]:
    if not _module_available("vulture"):
        return False, "vulture non installe"
    pyproject = ROOT / "pyproject.toml"
    cmd = [PYTHON, "-m", "vulture"]
    if pyproject.is_file():
        cmd.extend(["--config", "pyproject.toml"])
    else:
        cmd.extend([*_source_args(), "--min-confidence", "80"])
    code, out = _run(cmd, allow_missing=True)
    return code == 0, out or "Aucun code mort detecte"


def _check_bandit(_: QualiteArgs) -> tuple[bool, str]:
    if not _module_available("bandit"):
        return False, "bandit non installe"
    cmd = [
        PYTHON,
        "-m",
        "bandit",
        "-r",
        *_source_args(),
        "-c",
        "pyproject.toml",
        "-q",
    ]
    code, out = _run(cmd, allow_missing=True)
    if code == 127:
        return False, out
    return code == 0, out or "Aucun probleme de securite detecte"


def _check_mypy(_: QualiteArgs) -> tuple[bool, str]:
    if not _module_available("mypy"):
        return False, "mypy non installe (optionnel : python qualite.py --install)"
    code, out = _run(
        [PYTHON, "-m", "mypy", "--config-file", "pyproject.toml"],
        allow_missing=True,
    )
    return code == 0, out or "Typage mypy OK"


def _check_pyright(_: QualiteArgs) -> tuple[bool, str]:
    module = "basedpyright" if _module_available("basedpyright") else "pyright"
    if not _module_available(module):
        return False, "basedpyright/pyright non installe (optionnel : python qualite.py --install)"
    code, out = _run(
        [PYTHON, "-m", module, "--project", str(ROOT)],
        allow_missing=True,
    )
    return code == 0, out or "Typage pyright OK"


def _check_radon_cc(_: QualiteArgs) -> tuple[bool, str]:
    if not _module_available("radon"):
        return False, "radon non installe"
    code, out = _run(
        [
            PYTHON,
            "-m",
            "radon",
            "cc",
            *_source_args(),
            "-a",
            "-nc",
        ],
        allow_missing=True,
    )
    # -nc : n'affiche que complexite >= C
    return code == 0, out or "Complexite cyclomatique acceptable"


def _check_radon_mi(_: QualiteArgs) -> tuple[bool, str]:
    if not _module_available("radon"):
        return False, "radon non installe"
    code, out = _run(
        [PYTHON, "-m", "radon", "mi", *_source_args(), "-nc"],
        allow_missing=True,
    )
    return code == 0, out or "Indice de maintenabilite acceptable"


def _check_interrogate(_: QualiteArgs) -> tuple[bool, str]:
    if not _module_available("interrogate"):
        return False, "interrogate non installe"
    code, out = _run(
        [
            PYTHON,
            "-m",
            "interrogate",
            *_source_args(),
            "-c",
            "pyproject.toml",
        ],
        allow_missing=True,
    )
    return code == 0, out or "Couverture docstrings OK"


def _check_pytest(args: QualiteArgs) -> tuple[bool, str]:
    if not _module_available("pytest"):
        return False, "pytest non installe"
    tests_dir = ROOT / "tests"
    if not tests_dir.is_dir() or not any(tests_dir.glob("test_*.py")):
        return True, "Aucun test trouve (dossier tests/ vide)"
    cmd = [
        PYTHON,
        "-m",
        "pytest",
        "tests",
        "--cov=scripts",
        "--cov=generer_site.py",
        "--cov=preparer_excel.py",
        "--cov=qualite.py",
        "--cov-report=term-missing:skip-covered",
        "--cov-fail-under=0",
    ]
    if args.verbose:
        cmd.append("-v")
    else:
        cmd.append("-q")
    code, out = _run(cmd, allow_missing=True)
    return code == 0, out or "Tests OK"


def _check_pip_audit(_: QualiteArgs) -> tuple[bool, str]:
    if not _module_available("pip_audit"):
        return False, "pip-audit non installe"
    req_files = [
        p
        for p in (
            ROOT / "scripts" / "requirements.txt",
            REQUIREMENTS_DEV,
        )
        if p.is_file()
    ]
    if not req_files:
        return True, "Aucun fichier requirements a auditer"
    messages: list[str] = []
    ok = True
    for req in req_files:
        code, out = _run(
            [PYTHON, "-m", "pip_audit", "-r", str(req), "--progress-spinner=off"],
            allow_missing=True,
        )
        if code != 0:
            ok = False
        header = f"=== {req.relative_to(ROOT)} ==="
        messages.append(header)
        messages.append(out or ("OK" if code == 0 else "Vulnerabilites detectees"))
    return ok, "\n".join(messages)


def _check_deptry(_: QualiteArgs) -> tuple[bool, str]:
    if not _module_available("deptry"):
        return False, "deptry non installe"
    code, out = _run(
        [
            PYTHON,
            "-m",
            "deptry",
            ".",
            "--config",
            "pyproject.toml",
        ],
        allow_missing=True,
    )
    if code == 127:
        return False, out
    # deptry peut echouer sans pyproject [project] complet
    if "Could not find" in out and code != 0:
        return True, out + "\n(deptry ignore : configuration projet minimale)"
    return code == 0, out or "Dependances coherentes"


def _check_requirements_pinned(_: QualiteArgs) -> tuple[bool, str]:
    """Verifie que les requirements utilisent des bornes de version."""
    issues: list[str] = []
    for req_file in (
        ROOT / "scripts" / "requirements.txt",
        REQUIREMENTS_DEV,
    ):
        if not req_file.is_file():
            continue
        for line in req_file.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith(("#", "-r")):
                continue
            if "==" in stripped:
                issues.append(f"{req_file.name}: version figee (==) : {stripped}")
            elif ">=" not in stripped and "~=" not in stripped:
                issues.append(f"{req_file.name}: sans borne de version : {stripped}")
    if issues:
        return False, "\n".join(issues)
    return True, "Requirements avec bornes de version"


CHECKS: list[Check] = [
    Check("tools", "Outils installes", "prerequis", _ensure_dev_tools),
    Check("compileall", "Syntaxe Python (compileall)", "syntaxe", _check_compileall),
    Check("ruff-format", "Formatage (ruff format)", "style", _check_ruff_format),
    Check("ruff", "Lint (ruff check)", "lint", _check_ruff_lint),
    Check("vulture", "Code mort (vulture)", "analyse", _check_vulture),
    Check("bandit", "Securite (bandit)", "securite", _check_bandit),
    Check("mypy", "Typage (mypy)", "types", _check_mypy, optional=True),
    Check("pyright", "Typage (basedpyright)", "types", _check_pyright, optional=True),
    Check(
        "radon-cc",
        "Complexite cyclomatique (radon cc)",
        "metriques",
        _check_radon_cc,
        warn_only=True,
    ),
    Check(
        "radon-mi",
        "Maintenabilite (radon mi)",
        "metriques",
        _check_radon_mi,
        warn_only=True,
    ),
    Check(
        "interrogate",
        "Docstrings (interrogate)",
        "documentation",
        _check_interrogate,
        optional=True,
        warn_only=True,
    ),
    Check("pytest", "Tests unitaires (pytest)", "tests", _check_pytest),
    Check(
        "pip-audit",
        "Vulnerabilites deps (pip-audit)",
        "securite",
        _check_pip_audit,
        optional=True,
    ),
    Check(
        "deptry",
        "Deps inutilisees (deptry)",
        "deps",
        _check_deptry,
        optional=True,
        warn_only=True,
    ),
    Check("requirements", "Conventions requirements.txt", "deps", _check_requirements_pinned),
]

CHECK_BY_ID = {c.id: c for c in CHECKS}


def _install_dev_requirements() -> int:
    if not REQUIREMENTS_DEV.is_file():
        print(f"Fichier introuvable : {REQUIREMENTS_DEV}", file=sys.stderr)
        return 1
    print(f"Installation : pip install -r {REQUIREMENTS_DEV.name}", flush=True)
    proc = subprocess.run(
        [PYTHON, "-m", "pip", "install", "-r", str(REQUIREMENTS_DEV)],
        cwd=ROOT,
        check=False,
    )
    return proc.returncode


def _print_header(title: str) -> None:
    line = "=" * min(72, max(len(title) + 4, 40))
    print(f"\n{line}\n  {title}\n{line}", flush=True)


def _print_versions() -> None:
    tools = (
        ("ruff", "ruff"),
        ("vulture", "vulture"),
        ("bandit", "bandit"),
        ("mypy", "mypy"),
        ("basedpyright", "basedpyright"),
        ("pytest", "pytest"),
        ("radon", "radon"),
        ("interrogate", "interrogate"),
        ("pip_audit", "pip-audit"),
        ("deptry", "deptry"),
    )
    print("Versions detectees :", flush=True)
    for label, mod in tools:
        print(f"  - {label:14} {_tool_version(mod)}", flush=True)


def _select_checks(args: QualiteArgs) -> list[Check]:
    selected = list(CHECKS)
    if args.only:
        wanted = {x.strip() for x in args.only.split(",") if x.strip()}
        unknown = wanted - CHECK_BY_ID.keys()
        if unknown:
            raise SystemExit(f"Controles inconnus : {', '.join(sorted(unknown))}")
        selected = [CHECK_BY_ID[i] for i in (c.id for c in CHECKS) if i in wanted]
    if args.skip:
        skipped = {x.strip() for x in args.skip.split(",") if x.strip()}
        selected = [c for c in selected if c.id not in skipped]
    return selected


def _run_checks(args: QualiteArgs) -> int:
    selected = _select_checks(args)
    failures = 0
    warnings = 0
    skipped = 0

    _print_versions()

    for check in selected:
        _print_header(f"{check.label} [{check.id}]")
        try:
            ok, message = check.runner(args)
        except FileNotFoundError as exc:
            ok, message = False, str(exc)

        if not ok and check.optional and not args.strict:
            print(f"IGNORE (outil optionnel) : {message}", flush=True)
            skipped += 1
            continue

        if message:
            print(textwrap.indent(message, "  "), flush=True)

        if ok:
            print("  -> OK", flush=True)
        elif check.warn_only and not args.strict:
            print("  -> AVERTISSEMENT (non bloquant)", flush=True)
            warnings += 1
        else:
            print("  -> ECHEC", flush=True)
            failures += 1

    _print_header("Resume")
    total = len(selected)
    passed = total - failures - warnings - skipped
    print(
        f"  {passed} OK | {failures} echec(s) | {warnings} avertissement(s) | {skipped} ignore(s)",
        flush=True,
    )

    if failures:
        print("\nCorrections rapides :", flush=True)
        print(f"  {PYTHON} qualite.py --fix          # ruff format + lint auto-fix", flush=True)
        print(f"  {PYTHON} qualite.py --install       # installer tous les outils", flush=True)
        return 1
    if warnings and args.strict:
        return 1
    return 0


def _list_checks() -> None:
    print("Controles disponibles :\n")
    current_cat = ""
    for check in CHECKS:
        if check.category != current_cat:
            current_cat = check.category
            print(f"[{current_cat}]")
        flags = []
        if check.optional:
            flags.append("optionnel")
        if check.warn_only:
            flags.append("avertissement")
        suffix = f" ({', '.join(flags)})" if flags else ""
        print(f"  {check.id:14} {check.label}{suffix}")
    print(
        textwrap.dedent(
            f"""
            Exemples :
              {PYTHON} qualite.py
              {PYTHON} qualite.py --only ruff,pytest,vulture
              {PYTHON} qualite.py --fix
            """
        ).strip()
    )


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Lance les outils de qualite du projet CarteVoyage.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent(
            """
            Outils couverts : compileall, ruff (format+lint), vulture, bandit,
            mypy, basedpyright, radon, interrogate, pytest+coverage, pip-audit,
            deptry, verification des requirements.
            """
        ).strip(),
    )
    _ = parser.add_argument(
        "--fix",
        action="store_true",
        help="Applique ruff format et ruff check --fix avant les autres controles",
    )
    _ = parser.add_argument(
        "--install",
        action="store_true",
        help="Installe requirements-dev.txt puis quitte",
    )
    _ = parser.add_argument(
        "--list",
        action="store_true",
        help="Affiche la liste des controles et quitte",
    )
    _ = parser.add_argument(
        "--only",
        metavar="IDS",
        help="Controles a lancer (virgules), ex. ruff,pytest,vulture",
    )
    _ = parser.add_argument(
        "--skip",
        metavar="IDS",
        help="Controles a ignorer, ex. radon-cc,interrogate",
    )
    _ = parser.add_argument(
        "--strict",
        action="store_true",
        help="Echoue aussi sur outils optionnels manquants et avertissements",
    )
    _ = parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Sortie pytest detaillee",
    )
    return parser


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args(namespace=QualiteArgs())

    if args.install:
        raise SystemExit(_install_dev_requirements())

    if args.list:
        _list_checks()
        return

    if not shutil.which(PYTHON):
        raise SystemExit(f"Python introuvable : {PYTHON}")

    raise SystemExit(_run_checks(args))


if __name__ == "__main__":
    main()
