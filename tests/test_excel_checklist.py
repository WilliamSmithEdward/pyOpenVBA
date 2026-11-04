"""docs/excel_checklist.csv agrees with what the Excel model registers.

The checklist records a status per member (docs/host_completeness.md
defines them). These tests keep it honest as the model grows: a member
registered in code cannot be listed as missing, a listed pass needs
evidence that exists, and no member of an implemented class is left out.
"""

from __future__ import annotations

import ast
import csv
import inspect
import json
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import (
    _autofilter,
    _cell_format,
    _chart_axes,
    _chart_data_labels,
    _chart_error_bars,
    _chart_fonts,
    _chart_format,
    _chart_points,
    _chart_trendlines,
    _charts,
    _colour_scales,
    _connections,
    _custom_properties,
    _data_bars,
    _document_properties,
    _error_checks,
    _formats,
    _hyperlinks,
    _icon_sets,
    _model,
    _notes,
    _outline,
    _page_breaks,
    _page_setup,
    _pivot_calculated,
    _pivot_cells,
    _pivot_criteria,
    _pivot_items,
    _old_controls,
    _old_drawings,
    _pivots,
    _query_tables,
    _protection,
    _rich_text,
    _scenarios,
    _shape_format,
    _shapes,
    _sheet_tabs,
    _sparklines,
    _sort,
    _tables,
    _threads,
    _validation,
    _windows,
    _allow_edit_ranges,
    _dialogs,
    _workbook_more,
    _workbook_settings,
)
from pyopenvba.interpreter._inventory_data import MEMBERS
from pyopenvba.interpreter._objects import MemberSpec, VBAObject

ROOT = Path(__file__).resolve().parents[1]
CHECKLIST = ROOT / "docs" / "excel_checklist.csv"
ACCEPTED_GAPS = ROOT / "tests" / "fixtures" / "verified_gaps_accepted.json"
STATUSES = {"VERIFIED", "PARTIAL", "UNASSESSED", "MISSING", "EXCLUDED"}
IMPLEMENTED = {"VERIFIED", "PARTIAL", "UNASSESSED"}


def _rows() -> list[dict[str, str]]:
    with CHECKLIST.open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def _registered(cls: type[VBAObject]) -> dict[str, MemberSpec]:
    members: dict[str, MemberSpec] = getattr(cls, "_vba_members")
    return members


def _implemented(cls: type[VBAObject]) -> set[str]:
    """The members a class registers and the events the model raises for it (see _events)."""
    raised: frozenset[str] = getattr(cls, "raised_events", frozenset())
    return set(_registered(cls)) | {name.lower() for name in raised}


def _classes() -> dict[str, type[VBAObject]]:
    found: dict[str, type[VBAObject]] = {}
    for module in (_model, _shapes, _formats, _sort, _autofilter, _protection, _tables, _validation, _windows, _notes,
                   _hyperlinks, _chart_axes, _old_controls, _old_drawings, _query_tables, _pivots, _pivot_criteria, _pivot_calculated, _pivot_cells, _sheet_tabs, _page_setup,
                   _page_breaks, _outline, _colour_scales, _data_bars, _icon_sets, _document_properties, _rich_text,
                   _custom_properties, _error_checks, _scenarios, _cell_format, _sparklines,
                   _connections, _charts, _chart_data_labels, _chart_error_bars, _chart_trendlines, _pivot_items,
                   _threads, _allow_edit_ranges, _dialogs, _workbook_more, _workbook_settings, _shape_format, _chart_format, _chart_points, _chart_fonts):
        for cls in vars(module).values():
            if not inspect.isclass(cls) or cls.__module__ != module.__name__:
                continue
            if module is _pivots and cls.__name__ == '_Collection':
                continue
            if not issubclass(cls, VBAObject) or not _registered(cls) or not cls.vba_type_name:
                # A class with no type name is a base the types share, not a type.
                continue
            # Rows and Columns views share Range's type name; the class
            # with the most members is the one that stands for the type.
            current = found.get(cls.vba_type_name)
            if current is None or len(_registered(cls)) > len(_registered(current)):
                found[cls.vba_type_name] = cls
    return found


ROWS = _rows()
CLASSES = _classes()


def test_every_row_has_a_known_status_and_no_duplicates() -> None:
    assert ROWS
    assert {row["status"] for row in ROWS} <= STATUSES
    keys = [(row["object"], row["member"].lower()) for row in ROWS]
    assert len(keys) == len(set(keys))


def test_every_row_has_its_eight_fields() -> None:
    """A quote left undoubled inside a quoted field splits a row into more fields, which DictReader keeps under
    the key None."""
    split = [(row["object"], row["member"]) for row in ROWS if len(row) != 8]
    assert not split, split


@pytest.mark.parametrize("type_name", sorted(CLASSES))
def test_the_status_matches_the_registration(type_name: str) -> None:
    registered = _implemented(CLASSES[type_name])
    wrong = [
        f"{row['member']}: {row['status']}"
        for row in ROWS
        if row["object"] == type_name and (row["member"].lower() in registered) != (row["status"] in IMPLEMENTED)
    ]
    assert not wrong, "run scripts/build_excel_checklist.py, then review: " + ", ".join(wrong)


@pytest.mark.parametrize("type_name", sorted(CLASSES))
def test_every_member_of_an_implemented_class_is_listed(type_name: str) -> None:
    listed = {row["member"].lower() for row in ROWS if row["object"] == type_name}
    expected = set(_registered(CLASSES[type_name]))
    lowered = type_name.lower()
    for key in (f"excel:{lowered}", f"excel:i{lowered}", f"excel:_{lowered}"):
        expected.update(MEMBERS.get(key) or frozenset())
    assert expected <= listed, "run scripts/build_excel_checklist.py: " + ", ".join(sorted(expected - listed))


def test_a_pass_names_evidence_that_exists() -> None:
    for row in ROWS:
        if row["status"] not in {"VERIFIED", "PARTIAL"}:
            continue
        paths = [one.strip() for one in row["evidence"].split(";") if one.strip()]
        assert paths, f"{row['object']}.{row['member']} is {row['status']} with no evidence"
        for path in paths:
            assert (ROOT / path).exists(), f"{row['object']}.{row['member']}: {path} does not exist"


def test_verified_leaves_no_interface_gap_and_excluded_says_why() -> None:
    for row in ROWS:
        where = f"{row['object']}.{row['member']}"
        if row["status"] == "VERIFIED":
            assert not row["interface_gaps"], f"{where} is VERIFIED but lacks {row['interface_gaps']}"
        if row["status"] == "EXCLUDED":
            assert row["note"], f"{where} is EXCLUDED without a reason"


def _interface_gaps(row: dict[str, str], spec: MemberSpec | None) -> str:
    """What scripts/build_excel_checklist.py would write for a row, from the classes as they are now and the kind and
    signature it took from the type library."""
    if spec is None:
        return ""
    gaps: list[str] = []
    if row["kind"] == "property" and not spec.settable:
        gaps.append("no setter")
    opened = row["signature"].find("(")
    listed = row["signature"][opened + 1:row["signature"].rfind(")")] if row["kind"] != "property" and opened >= 0 else ""
    wanted = [re.sub(r"^(Optional |ByVal |ByRef |ParamArray )+", "", one.strip(" []")).split(" ")[0] for one in listed.split(",") if one.strip(" []")]
    wanted = [name for place, name in enumerate(wanted)
              if name != "lcid" and not (place == len(wanted) - 1 and re.fullmatch(r"pp[A-Z]\w*", name))]
    if wanted and not spec.varargs:
        have = {name.lower() for name in spec.parameters}
        absent = [name for name in wanted if name.lower() not in have]
        if absent:
            gaps.append("no " + ", ".join(absent))
    return "; ".join(gaps)


@pytest.mark.parametrize("type_name", sorted(CLASSES))
def test_the_interface_gaps_column_matches_the_classes(type_name: str) -> None:
    """The column is written by a script nobody has to run, so a setter removed by hand would leave it blank and the row
    VERIFIED: this works the gaps out again."""
    registered = _registered(CLASSES[type_name])
    wrong = [
        f"{row['member']}: says {row['interface_gaps']!r}, is {_interface_gaps(row, registered.get(row['member'].lower()))!r}"
        for row in ROWS
        if row["object"] == type_name and row["interface_gaps"] != _interface_gaps(row, registered.get(row["member"].lower()))
    ]
    assert not wrong, "run scripts/build_excel_checklist.py, then review: " + "; ".join(wrong)


def _gaps_in(function: object) -> tuple[list[str], list[str]]:
    """The VBAUnsupportedError messages a function raises itself, and the names of the self.x(...) it calls."""
    try:
        source = inspect.getsource(inspect.unwrap(function))  # type: ignore[arg-type]
    except (OSError, TypeError):
        return [], []
    raised: list[str] = []
    called: list[str] = []
    for node in ast.walk(ast.parse("class _:\n" + source if source[0].isspace() else source)):
        if isinstance(node, ast.Raise) and isinstance(node.exc, ast.Call) and getattr(node.exc.func, "id", "") == "VBAUnsupportedError":
            first = node.exc.args[0]
            raised.append(str(first.value) if isinstance(first, ast.Constant) else ast.unparse(first))
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and getattr(node.func.value, "id", "") == "self":
            called.append(node.func.attr)
    return raised, called


def _reachable_gaps(cls: type[VBAObject], function: object, seen: set[object]) -> list[str]:
    if function is None or function in seen:
        return []
    seen.add(function)
    raised, called = _gaps_in(function)
    for name in called:
        found = inspect.getattr_static(cls, name, None)
        if inspect.isfunction(found):
            raised += _reachable_gaps(cls, found, seen)
    return raised


def _verified_gaps() -> dict[str, set[str]]:
    gaps: dict[str, set[str]] = {}
    for row in ROWS:
        cls = CLASSES.get(row["object"])
        spec = _registered(cls).get(row["member"].lower()) if cls is not None and row["status"] == "VERIFIED" else None
        if cls is None or spec is None:
            continue
        seen: set[object] = set()
        for function in (spec.getter, spec.setter, spec.ref_setter):
            for gap in _reachable_gaps(cls, function, seen):
                gaps.setdefault(gap, set()).add(f"{row['object']}.{row['member']}")
    return gaps


def test_a_verified_row_reaches_no_reported_gap_but_an_accepted_one() -> None:
    """An unsupported error is a reported gap, not a passing gate (docs/host_completeness.md). A gap a macro could not
    plausibly hit may stay under a VERIFIED row if tests/fixtures/verified_gaps_accepted.json gives the reason.

    The scan reads source only: it follows self.x(...) calls within the row's own class, so a gap raised in a module's
    function, in another class or by a value the interpreter turns away is not seen. A pass does not say there is none."""
    accepted = json.loads(ACCEPTED_GAPS.read_text(encoding="utf-8"))
    gaps = _verified_gaps()
    wrong = [f"{gap!r}: {sorted(gaps[gap])}" for gap in sorted(gaps) if gap not in accepted]
    assert not wrong, "PARTIAL these, or accept the gap with a reason: " + "; ".join(wrong)
    assert all(accepted.values()), "an accepted gap needs its reason"
    assert not set(accepted) - set(gaps), f"accepted but no VERIFIED row reaches it: {sorted(set(accepted) - set(gaps))}"
