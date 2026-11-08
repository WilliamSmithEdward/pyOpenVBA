"""docs/excel_checklist.csv agrees with what the Excel model registers.

The checklist records a status per member (docs/host_completeness.md
defines them). These tests keep it honest as the model grows: a member
registered in code cannot be listed as missing, a listed pass needs
evidence that exists, and no member of an implemented class is left out.
"""

from __future__ import annotations

import ast
import csv
import dataclasses
import functools
import importlib
import inspect
import json
import re
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import (
    _autofilter,
    _cell_format,
    _cell_styles,
    _chart_axes,
    _chart_data_labels,
    _chart_error_bars,
    _chart_fonts,
    _chart_format,
    _chart_points,
    _chart_trendlines,
    _charts,
    _colour_scales,
    _conditional,
    _conditional_display,
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
from pyopenvba.exceptions import VBAUnsupportedError
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
                   _threads, _allow_edit_ranges, _dialogs, _workbook_more, _workbook_settings, _shape_format, _chart_format, _chart_points, _chart_fonts,
                   _cell_styles, _conditional, _conditional_display):
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


@dataclasses.dataclass(frozen=True)
class _Source:
    """What one function's source says: the gaps it raises itself and the calls the scan can name."""

    raised: list[str]
    on_self: list[str]
    named: list[str]
    imported: dict[str, tuple[str, str]]


def _tree_of(function: object) -> ast.AST | None:
    try:
        source = inspect.getsource(inspect.unwrap(function))  # type: ignore[arg-type]
    except (OSError, TypeError):
        return None
    return ast.parse("class _:\n" + source if source[0].isspace() else source)


def _imports_in(tree: ast.AST) -> dict[str, tuple[str, str]]:
    return {alias.asname or alias.name: (node.module, alias.name) for node in ast.walk(tree)
            if isinstance(node, ast.ImportFrom) and node.module and _scanned_module(node.module) for alias in node.names}


def _dotted(node: ast.expr) -> str:
    """A name or a module's attribute as the scan names calls, or "" for anything else."""
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name):
        return f"{node.value.id}.{node.attr}"
    return ""


def _is_refusal_class(found: object) -> bool:
    return inspect.isclass(found) and issubclass(found, VBAUnsupportedError)


def _is_refusal_partial(found: object) -> bool:
    is_partial = isinstance(found, functools.partial)
    made: object = getattr(found, "func", None)
    return is_partial and _is_refusal_class(made)


def _message(call: ast.Call, name: str) -> str:
    """What a refusal says, as the scan reports it: its first argument, or the name it was made by."""
    if not call.args:
        return name
    first = call.args[0]
    return str(first.value) if isinstance(first, ast.Constant) else ast.unparse(first)


def _refusals_made(node: ast.expr | None, function: object, imported: dict[str, tuple[str, str]],
                   kept: dict[str, list[str]]) -> list[str] | None:
    """The refusals an expression makes, or None where it makes none: a VBAUnsupportedError (or a subclass) called or
    named, a name bound to one, a functools.partial of the class, a function that returns or raises one, or a variable
    of the function that keeps one."""
    if node is None:
        return None
    if isinstance(node, ast.Name) and node.id in kept:
        return kept[node.id]
    if isinstance(node, ast.Call):
        name = _dotted(node.func)
        found = _called(name, function, imported) if name else None
        if _is_refusal_class(found) or _is_refusal_partial(found):
            return [_message(node, name)]
        if inspect.isfunction(found) and _scanned_module(found.__module__) and _makes_refusal(found):
            return [_message(node, name)]
        return None
    name = _dotted(node)
    found = _called(name, function, imported) if name else None
    if isinstance(found, VBAUnsupportedError):
        return [str(found)]
    if _is_refusal_class(found):
        return [name]
    return None


#: Whether each function looked at returns a refusal; False while it is still being looked at, which ends a loop.
_MAKES_REFUSAL: dict[object, bool] = {}


def _makes_refusal(function: object) -> bool:
    """Whether a function hands back a refusal to raise: it returns one, made in any of the ways _refusals_made sees,
    through other functions that do as well."""
    if function in _MAKES_REFUSAL:
        return _MAKES_REFUSAL[function]
    _MAKES_REFUSAL[function] = False
    tree = _tree_of(function)
    if tree is not None:
        imported, kept = _imports_in(tree), _kept_refusals(tree, function)
        _MAKES_REFUSAL[function] = any(isinstance(node, ast.Return) and _refusals_made(node.value, function, imported, kept) is not None
                                       for node in ast.walk(tree))
    return _MAKES_REFUSAL[function]


def _kept_refusals(tree: ast.AST, function: object) -> dict[str, list[str]]:
    """The function's variables given a refusal, which a later `raise name` raises."""
    imported: dict[str, tuple[str, str]] = _imports_in(tree)
    kept: dict[str, list[str]] = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign) and all(isinstance(target, ast.Name) for target in node.targets):
            made = _refusals_made(node.value, function, imported, kept)
            if made is not None:
                for target in node.targets:
                    kept[target.id] = made  # type: ignore[attr-defined]
    return kept


@functools.cache
def _source_of(function: object) -> _Source:
    found = _Source([], [], [], {})
    tree = _tree_of(function)
    if tree is None:
        return found
    found.imported.update(_imports_in(tree))
    kept = _kept_refusals(tree, function)
    for node in ast.walk(tree):
        if isinstance(node, ast.Raise):
            found.raised.extend(_refusals_made(node.exc, function, found.imported, kept) or [])
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            found.named.append(node.func.id)
        elif isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and isinstance(node.func.value, ast.Name):
            if node.func.value.id == "self":
                found.on_self.append(node.func.attr)
            else:
                found.named.append(f"{node.func.value.id}.{node.func.attr}")
    return found


def _scanned_module(name: str) -> bool:
    return name == __name__ or name == "pyopenvba" or name.startswith("pyopenvba.")


def _called(name: str, function: object, imported: dict[str, tuple[str, str]]) -> object:
    """The plain function a name in `function` stands for: a global of its module, or a name imported inside it, or a function of a module either holds."""
    head, _, attribute = name.partition(".")
    if head in imported:
        module, member = imported[head]
        found = getattr(importlib.import_module(module), member, None)
    else:
        found = getattr(inspect.unwrap(function), "__globals__", {}).get(head)  # type: ignore[arg-type]
    if attribute:
        return getattr(found, attribute, None) if inspect.ismodule(found) and _scanned_module(found.__name__) else None
    return found


def _reachable_gaps(cls: type | None, function: object, seen: set[object]) -> list[str]:
    """The gaps `function` raises or reaches: through self.x(...) on `cls`, and through plain functions anywhere in the package.

    A gap is any raise of a VBAUnsupportedError or a subclass, however it is made: named in the raise, by a name bound to
    the class or to one, by a functools.partial of the class, by a function that returns one (followed through the
    functions it returns from), or kept in a variable first (see _refusals_made).

    A floor, not a count. It cannot see a method of another object, a function passed along as a callback, code the interpreter
    runs, a value turned away with some other error, or which of the gaps a given call can really arrive at."""
    if function is None or function in seen:
        return []
    seen.add(function)
    source = _source_of(function)
    raised = list(source.raised)
    for name in source.on_self:
        found = inspect.getattr_static(cls, name, None) if cls is not None else None
        if inspect.isfunction(found):
            raised += _reachable_gaps(cls, found, seen)
    for name in source.named:
        found = _called(name, function, source.imported)
        if inspect.isfunction(found) and _scanned_module(found.__module__):
            raised += _reachable_gaps(None, found, seen)
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


def test_the_rows_that_read_formula2_still_reach_its_gap() -> None:
    """Formula2's reading stands apart from Formula's, which never asks for it. The scan has to go on seeing it from Formula2's rows."""
    cls = CLASSES["Range"]
    for member in ("formula2", "formula2r1c1", "formula2local", "formula2r1c1local"):
        spec = _registered(cls)[member]
        seen: set[object] = set()
        assert any("Formula2 of" in gap for gap in _reachable_gaps(cls, spec.getter, seen)), member
    spec = _registered(cls)["formula"]
    assert not any("Formula2 of" in gap for gap in _reachable_gaps(cls, spec.getter, set()))


def test_a_verified_row_reaches_no_reported_gap_but_an_accepted_one() -> None:
    """An unsupported error is a reported gap, not a passing gate (docs/host_completeness.md). A gap a macro could not
    plausibly hit may stay under a VERIFIED row if tests/fixtures/verified_gaps_accepted.json gives the reason.

    The scan reads source only and is a floor: see _reachable_gaps for what it cannot see. A pass does not say there is none."""
    accepted = json.loads(ACCEPTED_GAPS.read_text(encoding="utf-8"))
    gaps = _verified_gaps()
    wrong = [f"{gap!r}: {sorted(gaps[gap])}" for gap in sorted(gaps) if gap not in accepted]
    assert not wrong, "PARTIAL these, or accept the gap with a reason: " + "; ".join(wrong)
    assert all(accepted.values()), "an accepted gap needs its reason"
    assert not set(accepted) - set(gaps), f"accepted but no VERIFIED row reaches it: {sorted(set(accepted) - set(gaps))}"


def _planted_gap() -> None:
    raise VBAUnsupportedError("planted in a function")


def _planted_step() -> None:
    _planted_gap()


def _planted_loop() -> None:
    _planted_loop_back()
    raise VBAUnsupportedError("planted in a loop")


def _planted_loop_back() -> None:
    _planted_loop()


def _call(function: Any) -> None:
    function()


class _PlantedRefusal(VBAUnsupportedError):
    def __init__(self, message: str = "planted with no message") -> None:
        super().__init__(message)


def _planted_factory(what: str) -> VBAUnsupportedError:
    return VBAUnsupportedError(f"{what} is not implemented")


def _planted_factory_of_factory(what: str) -> VBAUnsupportedError:
    return _planted_factory(what)


def _planted_returns_itself(what: str) -> Any:
    return _planted_returns_itself(what)


def _planted_other_error(what: str) -> ValueError:
    return ValueError(what)


_PLANTED_ALIAS = VBAUnsupportedError
_PLANTED_PARTIAL = functools.partial(VBAUnsupportedError)
_PLANTED_INSTANCE = VBAUnsupportedError("planted as an instance")


class _PlantedForms:
    """Each way a refusal is raised other than by naming VBAUnsupportedError in the raise."""

    def subclass(self) -> None:
        raise _PlantedRefusal("planted as a subclass")

    def helper(self) -> None:
        raise _planted_factory("planted by a helper")

    def helper_of_helper(self) -> None:
        raise _planted_factory_of_factory("planted by a helper's helper")

    def alias(self) -> None:
        raise _PLANTED_ALIAS("planted by an alias")

    def partial(self) -> None:
        raise _PLANTED_PARTIAL("planted by a partial")

    def instance(self) -> None:
        raise _PLANTED_INSTANCE

    def variable(self) -> None:
        refusal = VBAUnsupportedError("planted in a variable")
        raise refusal

    def variable_from_helper(self) -> None:
        refusal = _planted_factory("planted in a variable by a helper")
        raise refusal

    def bare_class(self) -> None:
        raise _PlantedRefusal

    def other_error(self) -> None:
        raise _planted_other_error("not a refusal")

    def returns_itself(self) -> None:
        raise _planted_returns_itself("not a refusal")


class _PlantedOther:
    def gap(self) -> None:
        raise VBAUnsupportedError("planted on another object")


class _Planted:
    def own(self) -> None:
        raise VBAUnsupportedError("planted on the class")

    def through_self(self) -> None:
        self.own()

    def through_function(self) -> None:
        _planted_step()

    def through_function_import(self) -> None:
        from test_excel_checklist import _planted_step as renamed  # pyright: ignore[reportMissingImports, reportUnknownVariableType, reportPrivateUsage]

        renamed()

    def through_module(self) -> None:
        _THIS._planted_step()  # pyright: ignore[reportPrivateUsage]

    def through_loop(self) -> None:
        _planted_loop()

    def on_another_object(self) -> None:
        _PlantedOther().gap()

    def as_a_callback(self) -> None:
        list(map(_call, [_planted_gap]))

    def nothing(self) -> None:
        return None


_THIS = sys.modules[__name__]


@pytest.mark.parametrize(
    ("method", "gaps"),
    [
        ("own", {"planted on the class"}),
        ("through_self", {"planted on the class"}),
        ("through_function", {"planted in a function"}),
        ("through_function_import", {"planted in a function"}),
        ("through_module", {"planted in a function"}),
        ("through_loop", {"planted in a loop"}),
        ("nothing", set[str]()),
    ],
)
def test_the_scan_finds_a_planted_gap_by_every_way_it_follows(method: str, gaps: set[str]) -> None:
    assert set(_reachable_gaps(_Planted, getattr(_Planted, method), set())) == gaps


@pytest.mark.parametrize(
    ("method", "gaps"),
    [
        ("subclass", {"planted as a subclass"}),
        ("helper", {"planted by a helper"}),
        ("helper_of_helper", {"planted by a helper's helper"}),
        ("alias", {"planted by an alias"}),
        ("partial", {"planted by a partial"}),
        ("instance", {"planted as an instance"}),
        ("variable", {"planted in a variable"}),
        ("variable_from_helper", {"planted in a variable by a helper"}),
        ("bare_class", {"_PlantedRefusal"}),
        ("other_error", set[str]()),
        ("returns_itself", set[str]()),
    ],
)
def test_the_scan_finds_a_refusal_raised_in_every_form(method: str, gaps: set[str]) -> None:
    assert set(_reachable_gaps(_PlantedForms, getattr(_PlantedForms, method), set())) == gaps


@pytest.mark.parametrize("method", ["on_another_object", "as_a_callback"])
def test_the_scan_is_a_floor_and_misses_what_its_docstring_says(method: str) -> None:
    """Should these start to be found, the docstring is out of date. A row may reach a gap this way and stay VERIFIED."""
    assert _reachable_gaps(_Planted, getattr(_Planted, method), set()) == []


@pytest.mark.parametrize(
    ("object_name", "member", "gap", "through"),
    [
        ("Border", "Color", "editing a filtered sheet whose rows are hidden by default is not implemented", "a function of another module"),
        ("Range", "Style", "giving a range a Style object of another workbook is not implemented", "the class's own code"),
        ("HPageBreaks", "Count", "Automatic or partial page breaks are not implemented", "a function"),
        ("ShapeRange", "TextFrame", "TextFrame of a ShapeRange of several shapes", "a helper that makes the refusal"),
        ("ListObject", "Delete", "with invalid table XML", "a helper of another module that makes the refusal"),
    ],
)
def test_the_scan_finds_real_gaps(object_name: str, member: str, gap: str, through: str) -> None:
    cls = CLASSES[object_name]
    spec = _registered(cls)[member.lower()]
    seen: set[object] = set()
    found = [one for function in (spec.getter, spec.setter, spec.ref_setter) for one in _reachable_gaps(cls, function, seen)]
    assert gap in found, through
