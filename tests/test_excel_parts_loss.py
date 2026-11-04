"""What the model keeps of a file's other parts from a file Excel saved: what it reads, and what is left after setting a
cell nothing reads and saving (scripts/measure_parts_loss.py)."""

from __future__ import annotations

import importlib.util
import json
import re
import shutil
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parents[1]
HERE = ROOT / "tests/fixtures/parts_loss"
MEASURED: dict[str, Any] = json.loads((HERE / "parts_loss.json").read_text(encoding="utf-8"))
_spec = importlib.util.spec_from_file_location("measure_parts_loss", ROOT / "scripts/measure_parts_loss.py")
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)

ROWS: list[dict[str, Any]] = [row for row in MEASURED["rows"] if "failed" not in row]
#: Files or members the model says it cannot read, which a save keeps all the same.
REPORTED = {"chart_styled", "pivot_slicer", "vba_project", "vba_form", "format_xls", "custom_views", "ignored_errors"}
#: The model answers this wrongly rather than reporting: a header picture read back as none.
WRONG = {"image_header": "a header picture in a loaded file reads as none"}
#: The model does not open these, so nothing is saved.
REFUSED = {"format_xls", "format_xlsb"}


def _lines(text: object) -> str:
    """The harness's answer turns a CR into a LF, so line ends are not compared."""
    return re.sub(r"[\r\n]+", "\n", str(text))


def _reading(row: dict[str, Any], file: Path) -> str:
    source = MEASURED["help"].format(cases=PROBE.cases(row["reads"])) + PROBE.MODEL_READ.format(last=len(row["reads"]) - 1)
    return str(PROBE.model_answer(source, "Read1", file))


@pytest.mark.parametrize("row", [row for row in ROWS if row["name"] not in REPORTED | set(WRONG) | {"format_xlsb"}], ids=lambda row: row["name"])
def test_the_model_reads_a_file_as_excel_does(row: dict[str, Any]) -> None:
    assert _lines(_reading(row, HERE / f"{row['name']}.{row['format']}")) == _lines(row["excel"])


@pytest.mark.parametrize("name", sorted(REPORTED))
def test_what_the_model_cannot_read_it_reports(name: str) -> None:
    row = next(one for one in ROWS if one["name"] == name)
    assert _reading(row, HERE / f"{name}.{row['format']}").startswith("reports:")


@pytest.mark.parametrize("name", sorted(WRONG))
def test_answers_still_wrong(name: str) -> None:
    row = next(one for one in ROWS if one["name"] == name)
    answer = _lines(_reading(row, HERE / f"{name}.{row['format']}"))
    if answer == _lines(row["excel"]):
        pytest.fail(f"{name} now matches Excel; move it out of WRONG")
    pytest.xfail(WRONG[name])


@pytest.mark.parametrize("row", [row for row in ROWS if row["name"] not in REFUSED], ids=lambda row: row["name"])
@pytest.mark.parametrize("route", ["macro", "python"])
def test_an_unrelated_edit_and_a_save_keep_every_part(row: dict[str, Any], route: str, tmp_path: Path) -> None:
    given = HERE / f"{row['name']}.{row['format']}"
    file = tmp_path / given.name
    shutil.copy(given, file)
    answer = PROBE.model_answer(PROBE.MODEL_EDIT, "Edit1", file) if route == "macro" else PROBE.python_edit(file)
    assert not str(answer).startswith(("reports", "crash")), answer
    assert PROBE.package_loss(given, file) == {}


def test_an_old_format_is_reported_from_python_too() -> None:
    with pytest.raises(VBAUnsupportedError):
        ExcelApplication.open(HERE / "format_xls.xls")


@pytest.mark.parametrize("name", sorted(PROBE.SEQUENCES))
def test_templates_and_add_ins_behave_as_in_excel(name: str) -> None:
    assert PROBE.run_sequence(name) == MEASURED["sequences"][name]
