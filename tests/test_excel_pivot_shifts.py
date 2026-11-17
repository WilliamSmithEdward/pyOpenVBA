"""Pivot tables under inserts and deletes of rows, columns and cells, replayed against live Excel
(scripts/measure_pivot_shifts.py): where the table goes, what a PivotTable held from before the edit and one fetched again
read, where a source list on the edited sheet goes, and what a refresh, a save and a reopen leave.

A cell insert or delete inside a table is an edit of the report in Excel, which the model refuses, and so is a refresh
once a column has gone into or out of the source list; every other case answers as Excel did.
"""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = tuple(stem for stem in ('pivot_shifts_a', 'pivot_shifts_b', 'pivot_shifts1', 'pivot_shifts2',
                                  'pivot_shifts3', 'pivot_shifts4')
               if (FIXTURES / f'{stem}.json').exists())
REFUSED = ('inside a pivot table', 'changed source schema', 'Deleting all of a pivot table')


def _script() -> ModuleType:
    path = Path(__file__).parents[1] / 'scripts/measure_pivot_shifts.py'
    spec = importlib.util.spec_from_file_location('measure_pivot_shifts', path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


SCRIPT = _script()


def _cases() -> list[tuple[str, list[Any], str]]:
    found: list[tuple[str, list[Any], str]] = []
    for stem in ROUNDS:
        record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
        found += [(stem, case, answer) for case, answer in zip(record['cases'], record['answers'], strict=True)]
    return found


CASES = _cases()


@pytest.fixture(autouse=True)
def own_temp(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """The probe saves its book under %TEMP% by one name: each test gets its own folder, so parallel runs keep apart."""
    monkeypatch.setenv('TEMP', str(tmp_path))


def _model(case: list[Any]) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SCRIPT.SOURCE, name='Probe')
    return str(app.run('Probe.Go', *case))


@pytest.mark.parametrize('stem, case, answer', CASES,
                         ids=[f'{one[0]}-{one[1][0]}-{one[1][1]}-{"typed" if one[1][2] else "late"}' for one in CASES])
def test_it_answers_as_excel_did(stem: str, case: list[Any], answer: str) -> None:
    try:
        found = _model(case)
    except VBAUnsupportedError as refusal:
        assert any(reason in str(refusal) for reason in REFUSED), refusal
        return
    assert found == answer


def test_a_cut_changes_nothing() -> None:
    """Excel's 1004 comes before anything moves: the cell beside the table and the table stay where they were."""
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('''Function Edit() As String
    Dim wb As Object, ws As Object, pt As Object
    Set wb = ActiveWorkbook: Set ws = wb.Worksheets(1)
    ws.Range("A1:B3").Value = Array("k", "v"): ws.Range("A2:B2").Value = Array("a", 1): ws.Range("A3:B3").Value = Array("b", 2)
    Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B3")).CreatePivotTable(ws.Range("D3"), "PT")
    pt.PivotFields("k").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum
    ws.Range("D1").Value = 7
    On Error Resume Next
    ws.Range("D1").Insert xlShiftDown
    Edit = Err.Number & " " & ws.Range("D1").Value & " " & pt.TableRange1.Address & " " & pt.SourceData
End Function
''', name='P')
    assert app.run('P.Edit') == '1004 7 $D$3:$E$6 Sheet1!R1C1:R3C2'


def test_an_edit_inside_a_table_reports_itself() -> None:
    """Deleting an item's label hides the item in Excel, which the model does not do."""
    with pytest.raises(VBAUnsupportedError, match='inside a pivot table'):
        _model(['other', 'delete C4 up', True])
