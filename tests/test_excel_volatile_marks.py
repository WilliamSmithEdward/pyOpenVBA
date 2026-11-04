"""A formula reading a volatile one is marked ca="1" in a file, as Excel otherwise leaves it behind on opening (scripts/measure_volatile_marks.py)."""

from __future__ import annotations

import importlib.util
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.formula._prefixes import calculated_always, calls_only_ordinary_functions

FIXTURES = Path(__file__).parent / "fixtures"
BOOK = FIXTURES / "volatile_marks.xlsx"
MEASURED: dict[str, Any] = json.loads((FIXTURES / "volatile_marks.json").read_text(encoding="utf-8"))
EDITS: dict[str, dict[str, Any]] = MEASURED["edits"]
#: Excel marks a formula that read a volatile one as it was worked out, and the model one that names it: IF(FALSE,A1,5),
#: CHOOSE(1,5,A1) and ROW(A1) are marked here and not by Excel, which costs working them out on opening.
NAMES_BUT_DOES_NOT_READ = {"E1", "U2", "U3", "Y1"}
#: With the volatile formula gone, a sheet nobody wrote to keeps the marks its file had, which costs the same.
KEEPS_A_MARK = {"a volatile formula made plain", "a value in place of a volatile formula"}
_FORMULA = re.compile(r'<c r="([A-Z]+\d+)"[^>]*><f([^>]*?)(?:/>|>)')


def _saved(edit: str, target: Path) -> None:
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module(f'Public Sub Edit()\n    Dim wb As Object, ws As Object\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets("V")\n    {edit}\nEnd Sub\n', name="E")
    app.run("E.Edit")
    app.save(target)


def _marks(path: Path) -> dict[str, dict[str, str]]:
    with zipfile.ZipFile(path) as package:
        return {name: {reference: head.strip() for reference, head in _FORMULA.findall(package.read(part).decode("utf-8"))}
                for name, part in (("V", "xl/worksheets/sheet1.xml"), ("O", "xl/worksheets/sheet2.xml"))}


def _behind(answers: str) -> list[str]:
    """The formulas that do not agree with what they read."""
    return [one.split("=")[0] for one in answers.strip("|").split("|") if one.split("=")[1] != "0"]


def test_a_formula_without_the_mark_falls_behind() -> None:
    """Why it matters: on opening, Excel works out the volatile formula and not what reads it, unless that is marked."""
    assert not _behind(MEASURED["effect"]["marked"])
    behind = _behind(MEASURED["effect"]["marks_gone"])
    assert {"V!B1", "V!C1", "V!D1", "V!T1", "O!A1", "O!B1"} <= set(behind)
    assert "V!K1" not in behind and "O!C1" not in behind


@pytest.mark.parametrize("name", list(EDITS))
def test_every_mark_excel_writes_is_written(name: str, tmp_path: Path) -> None:
    target = tmp_path / "saved.xlsx"
    _saved(EDITS[name]["edit"], target)
    model = _marks(target)
    for sheet, cells in EDITS[name]["excel_marks"].items():
        assert set(cells) == set(model[sheet]), sheet
        for reference, head in cells.items():
            if head == model[sheet][reference]:
                continue
            assert head.replace(' ca="1"', "").replace('ca="1"', "").strip() == model[sheet][reference].replace(' ca="1"', "").replace('ca="1"', "").strip(), (sheet, reference)
            # What differs is only a mark the model adds.
            assert 'ca="1"' in model[sheet][reference] and 'ca="1"' not in head, (sheet, reference)
            assert reference in NAMES_BUT_DOES_NOT_READ or (name in KEEPS_A_MARK and sheet == "O"), (sheet, reference)


@pytest.mark.parametrize("name", [name for name in EDITS if name != "a name made volatile"])
def test_excel_finds_every_formula_up_to_date_on_opening(name: str) -> None:
    assert not _behind(EDITS[name]["model_agrees"])
    assert EDITS[name]["model_agrees"] == EDITS[name]["excel_agrees"]


def test_a_sheet_nobody_wrote_to_gains_the_mark(tmp_path: Path) -> None:
    """O!C1 reads V!J1, made volatile: only the mark and the value change on O."""
    target = tmp_path / "saved.xlsx"
    _saved('ws.Range("J1").Formula = "=RAND()"', target)
    assert _marks(target)["O"]["C1"] == 'ca="1"'
    assert _marks(target)["O"]["D1"] == 'ca="1"'
    assert _marks(BOOK)["O"]["C1"] == ""


def test_a_save_with_no_edit_loses_no_mark(tmp_path: Path) -> None:
    """A sheet Excel would tidy is written whole with no edit at all, which took every mark on a reader off it."""
    target = tmp_path / "saved.xlsx"
    ExcelApplication.open(BOOK, with_vba=False).save(target)
    model, excel = _marks(target), _marks(BOOK)
    for sheet, cells in excel.items():
        for reference, head in cells.items():
            assert head == model[sheet][reference] or reference in NAMES_BUT_DOES_NOT_READ, (sheet, reference)


@pytest.mark.parametrize("formula", ["=SUM(A1:A3)*ROW()", "=A1+1", "=IF(A1>0,1,2)", "=_xlfn.SINGLE(A1)", '="RAND"&A1'])
def test_the_quick_look_passes_over_ordinary_formulas(formula: str) -> None:
    assert calls_only_ordinary_functions(formula)
    assert not calculated_always(formula, lambda _: False)


@pytest.mark.parametrize("formula", ["=RAND()", "=A1+NOW()", "=OFFSET(A1,1,0)", "=indirect(A1)", "=MyMacro(A1)", "=_xlfn.SHEET()"])
def test_the_quick_look_stops_at_what_may_be_marked(formula: str) -> None:
    assert not calls_only_ordinary_functions(formula)
    assert calculated_always(formula, lambda _: False)


LIVE = os.environ.get("RUN_LIVE_EXCEL") == "1" and importlib.util.find_spec("pyvbaharness") is not None


@pytest.mark.skipif(not LIVE, reason="requires explicitly enabled native Excel")
def test_native_excel_finds_the_models_save_up_to_date(tmp_path: Path) -> None:
    harness: Any = pytest.importorskip("pyvbaharness")
    expressions = MEASURED["agrees"]
    lines = ['Public Function Agrees(ByVal path As String) As String', '    Dim wb As Workbook, v As Worksheet, o As Worksheet, out As String', '    Application.DisplayAlerts = False',
             '    Set wb = Workbooks.Open(path): Set v = wb.Worksheets("V"): Set o = wb.Worksheets("O")']
    lines += [f'    out = out & "|{name}=" & Round({expression}, 8)' for name, expression in expressions.items()]
    lines += ['    wb.Close False', '    Application.DisplayAlerts = True', '    Agrees = out', 'End Function', '']
    behind: dict[str, list[str]] = {}
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for number, (name, row) in enumerate(EDITS.items()):
            if name == "a name made volatile":
                continue
            target = tmp_path / f"saved{number}.xlsx"
            _saved(row["edit"], target)
            answer = excel.run_vba("\n".join(lines), "Agrees", (str(target),), timeout=60.0, line_numbers=False)
            if _behind(str(answer.value)):
                behind[name] = _behind(str(answer.value))
    assert not behind
