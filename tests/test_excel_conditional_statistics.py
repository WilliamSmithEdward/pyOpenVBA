"""Native statistical FormatConditions API and specialized object identity."""
from __future__ import annotations

import importlib.util
import json
import os
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._conditional import AboveAverage, Top10, UniqueValues

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/conditional/statistics_api.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_statistical_api(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']


def _saved_expectation(case: dict[str, Any]) -> tuple[str, str]:
    expected = 'class=' + case['result'].split('class=', 1)[1].split('|count=', 1)[0]
    return expected, expected.split(';', 1)[0].removeprefix('class=')


def _saved_project(case: dict[str, Any], path: Path) -> Path:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    expected, typename = _saved_expectation(case)
    inspect_source = RECORD['inspection_source'].replace('Inspect', 'SavedInspect').replace('Prop', 'SavedProp')
    app.add_module('''Function SavedRule(Wanted As String) As String
Dim f As Object
For Each f In Cells.FormatConditions
If TypeName(f)=Wanted Then
SavedRule=SavedInspect(f)
Exit Function
End If
Next
End Function
''' + inspect_source, name='Saved', workbook=app.workbook)
    macro = "'" + app.workbook.name.replace("'", "''") + "'!SavedRule"
    assert app.run(macro, typename) == expected
    path = app.save(path, workbook=app.workbook)
    reopened = ExcelApplication.open(path)
    assert reopened.run('SavedRule', typename) == expected
    return path


@pytest.mark.parametrize('case', [case for case in RECORD['probes'] if 'class=' in case['result']],
                         ids=lambda case: case['name'])
def test_statistical_rule_roundtrip_dispatch(case: dict[str, Any], tmp_path: Path) -> None:
    _saved_project(case, tmp_path / 'statistics.xlsm')


@pytest.fixture(scope='module')
def native_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_saved_statistics_read_by_native_excel(case: dict[str, Any], tmp_path: Path, native_excel: Any) -> None:
    if 'class=' not in case['result']:
        pytest.skip('native construction error left no rule to reopen')
    expected, typename = _saved_expectation(case)
    native_excel.open_document(_saved_project(case, tmp_path / 'statistics.xlsm'))
    assert native_excel.run_raw('Saved.SavedRule', typename) == expected
    assert not any(event.get('dialogs') for event in native_excel.events)


@pytest.mark.parametrize('method,typename', [('AddTop10', 'Top10'), ('AddAboveAverage', 'AboveAverage'),
                                          ('AddUniqueValues', 'UniqueValues')])
def test_specialized_member_surface_and_lifecycle(method: str, typename: str, tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('''Function Probe() As String
Dim f As Object
Set f=Range("A1:A6").FormatConditions.''' + method + '''
f.Interior.Color=255
f.ModifyAppliesToRange Range("C3:C8")
f.Priority=3
f.StopIfTrue=True
Probe=TypeName(f) & "|" & f.AppliesTo.Address & "|" & CStr(f.Priority) & "|" & CStr(f.Interior.Color)
End Function
Function ReadSaved() As String
Dim f As Object
Set f=Range("C3:C8").FormatConditions(1)
ReadSaved=TypeName(f) & "|" & f.AppliesTo.Address & "|" & CStr(f.Priority) & "|" & CStr(f.Interior.Color)
End Function
Function DeleteSaved() As String
Range("C3:C8").FormatConditions(1).Delete
DeleteSaved=CStr(Cells.FormatConditions.Count)
End Function''', name='Harness')
    expected = typename + '|$C$3:$C$8|3|255'
    assert app.run('Probe') == expected
    reopened = ExcelApplication.open(app.save(tmp_path / 'styled-statistic.xlsm'))
    assert reopened.run('ReadSaved') == expected
    assert reopened.run('DeleteSaved') == '0'
    # Specialized rules inherit common lifecycle/styles, never cellIs formulas.
    rule = book.sheets_[0].conditional_formats[0]
    wrapped = {'Top10': Top10, 'AboveAverage': AboveAverage, 'UniqueValues': UniqueValues}[typename](book.sheets_[0], rule)
    assert wrapped.vba_type_name == typename
    for name in ('Formula1', 'Formula2', 'Operator', 'Modify'):
        assert wrapped.vba_member(name) is None
