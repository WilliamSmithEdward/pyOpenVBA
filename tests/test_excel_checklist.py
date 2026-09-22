"""docs/excel_checklist.csv agrees with what the Excel model registers.

The checklist records a status per member (docs/host_completeness.md
defines them). These tests keep it honest as the model grows: a member
registered in code cannot be listed as missing, a listed pass needs
evidence that exists, and no member of an implemented class is left out.
"""

from __future__ import annotations

import csv
import inspect
from pathlib import Path

import pytest

from pyopenvba.apps.excel import _formats, _model, _shapes
from pyopenvba.interpreter._inventory_data import MEMBERS
from pyopenvba.interpreter._objects import MemberSpec, VBAObject

ROOT = Path(__file__).resolve().parents[1]
CHECKLIST = ROOT / "docs" / "excel_checklist.csv"
STATUSES = {"VERIFIED", "PARTIAL", "UNASSESSED", "MISSING", "EXCLUDED"}
IMPLEMENTED = {"VERIFIED", "PARTIAL", "UNASSESSED"}


def _rows() -> list[dict[str, str]]:
    with CHECKLIST.open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def _registered(cls: type[VBAObject]) -> dict[str, MemberSpec]:
    members: dict[str, MemberSpec] = getattr(cls, "_vba_members")
    return members


def _classes() -> dict[str, type[VBAObject]]:
    found: dict[str, type[VBAObject]] = {}
    for module in (_model, _shapes, _formats):
        for cls in vars(module).values():
            if not inspect.isclass(cls) or cls.__module__ != module.__name__:
                continue
            if not issubclass(cls, VBAObject) or not _registered(cls):
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


@pytest.mark.parametrize("type_name", sorted(CLASSES))
def test_the_status_matches_the_registration(type_name: str) -> None:
    registered = _registered(CLASSES[type_name])
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
