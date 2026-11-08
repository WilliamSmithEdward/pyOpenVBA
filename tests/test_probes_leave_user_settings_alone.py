"""No live probe sets what outlasts the probe or reaches outside Excel: a user-level setting kept in the registry, a mail envelope, saving to
the cloud, or a password a save encrypts or locks the file with (scripts/probe_limits.py). On 2026-11-07 a held-out round set
DoNotPromptForConvert False and left it so for William until it was put back. A probe may read these; a set is an assignment in its VBA or a
pair in a list of cases it sets. Only a probe that reads a user-level setting first and puts it back may set it, and it is named here."""

from __future__ import annotations

import ast
import importlib.util
import re
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
_SPEC = importlib.util.spec_from_file_location("probe_limits", ROOT / "scripts/probe_limits.py")
assert _SPEC is not None and _SPEC.loader is not None
_LIMITS: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(_LIMITS)
USER_LEVEL: frozenset[str] = _LIMITS.USER_LEVEL
NEVER_SET: frozenset[str] = _LIMITS.NEVER_SET

SCRIPTS = sorted((ROOT / "scripts").glob("measure_*.py"))
#: Probes that keep a user-level setting's value and put it back after each case.
RESTORES = {"measure_settable_audit": frozenset({"DoNotPromptForConvert"})}


def _sets(names: frozenset[str]) -> re.Pattern[str]:
    alternatives = "|".join(sorted(names))
    # .Name = value in VBA; ('Name', value) or 'Name': value in a list or table of cases a probe sets.
    return re.compile(rf"\.(?:{alternatives})\s*=(?!=)|\(\s*['\"](?:{alternatives})['\"]\s*,|['\"](?:{alternatives})['\"]\s*:")


def _code_strings(path: Path) -> list[str]:
    """Every string in a script's code, and its code itself without the docstring: a case list is Python, a probe's VBA is a string."""
    source = path.read_text(encoding="utf-8")
    tree = ast.parse(source)
    docstring = tree.body[0].value if tree.body and isinstance(tree.body[0], ast.Expr) else None
    strings = [node.value for node in ast.walk(tree) if isinstance(node, ast.Constant) and isinstance(node.value, str) and node is not docstring]
    body = source if docstring is None else source.replace(ast.get_source_segment(source, docstring) or "", "", 1)
    return [*strings, body]


def _found(path: Path) -> list[str]:
    allowed = RESTORES.get(path.stem, frozenset())
    pattern = _sets((USER_LEVEL - allowed) | NEVER_SET)
    return sorted({hit for text in _code_strings(path) for hit in pattern.findall(text)})


def test_there_are_probes_to_check() -> None:
    assert len(SCRIPTS) > 100


@pytest.mark.parametrize("path", SCRIPTS, ids=lambda path: path.stem)
def test_a_probe_sets_no_user_level_setting_or_password(path: Path) -> None:
    assert _found(path) == []


@pytest.mark.parametrize(("code", "caught"), [
    ('x = "wb.DoNotPromptForConvert = False"', True), ("x = [('EnvelopeVisible', 'True')]", True), ('x = {"Password": "p"}', True),
    ('x = "wb.WritePassword = \\"w\\""', True), ('x = "r = wb.DoNotPromptForConvert"', False), ("x = ['AutoSaveOn', 'Name']", False),
    ('x = "If wb.Password = \\"\\" Then"', True), ('x = "r = (wb.ReadOnly = True)"', False), ('"""It never sets AutoSaveOn = True."""', False)])
def test_the_check_catches_what_it_should(code: str, caught: bool, tmp_path: Path) -> None:
    path = tmp_path / "measure_x.py"
    path.write_text(code, encoding="utf-8")
    assert bool(_found(path)) is caught
