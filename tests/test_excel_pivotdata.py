"""GETPIVOTDATA from native reports, through cells and VBA Evaluate."""
from __future__ import annotations
import json
import importlib.util
import os
from pathlib import Path
from typing import Any
from zipfile import ZipFile
import pytest
from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/pivots'
RECORD: dict[str, Any] = json.loads((ROOT / 'pivotdata.json').read_text(encoding='utf-8'))
ARRAYS: dict[str, Any] = json.loads((ROOT / 'pivotdata-arrays.json').read_text(encoding='utf-8'))
FORMULA_TEXT: dict[str, Any] = json.loads((ROOT / 'pivotdata-formula-text.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['file'] + ':' + case['formula'])
def test_native_pivotdata_cell(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    formula = case['formula'].replace('"', '""')
    address = 'Z20' if case['file'].startswith('native-advanced-') else 'J20'
    app.add_module(f'Function Probe() As String\nRange("{address}").Formula="{formula}"\n'
                   f'Probe=CStr(Range("{address}").Value2)\nEnd Function', name='Harness')
    assert app.run('Probe') == case['cell_result']


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['file'] + ':' + case['formula'])
def test_native_pivotdata_evaluate(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    app.add_module(case['source'], name='Harness')
    assert app.run(case['source'].split('(')[0].split()[-1]) == case['result']


@pytest.mark.parametrize('case', ARRAYS['probes'], ids=lambda case: case['mode'] + ':' + case['formula'])
def test_native_pivotdata_arrays(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / ARRAYS['file'], with_vba=False)
    app.add_module(case['source'], name='Harness')
    assert app.run(case['source'].split('(')[0].split()[-1]) == case['result']


@pytest.mark.parametrize('case', FORMULA_TEXT['probes'], ids=lambda case: case['mode'] + ':' + case['formula'])
def test_native_pivotdata_formula_text(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / FORMULA_TEXT['file'], with_vba=False)
    app.add_module(case['source'], name='Harness')
    assert app.run(case['source'].split('(')[0].split()[-1]) == case['result']


FILES = list(dict.fromkeys(case['file'] for case in RECORD['probes']))
ARRAY_FILES = [case for case in ARRAYS['probes'] if case['mode'] in ('Formula2', 'FormulaArray') and case['formula'] in {
    '=GETPIVOTDATA({"Revenue","Count Units";"Amount","Units"},H3)',
    '=GETPIVOTDATA("Revenue",H3,"Region",B20:B21)',
    '=GETPIVOTDATA(IF(TRUE,{"Revenue","Count Units"},"Revenue"),H3)',
    '=GETPIVOTDATA(INDEX({"Revenue";"Count Units"},0),H3)',
    '=GETPIVOTDATA(ReportCaptions,H3)',
}]


def _array_file(case: dict[str, Any], path: Path) -> str:
    app = ExcelApplication.open(ROOT / ARRAYS['file'], with_vba=False)
    app.add_module(case['source'], name='Harness')
    assert app.run(case['source'].split('(')[0].split()[-1]) == case['result']
    app.save(path)
    if case['mode'] == 'FormulaArray':
        read = 'Set result=Range("Z20:Z21")\n'
    else:
        read = 'On Error Resume Next\nSet result=Range("Z20").SpillingToRange\nOn Error GoTo 0\n'
        read += 'If result Is Nothing Then Set result=Range("Z20")\n'
    return 'Function ReadStored() As String\nDim result As Range\n' + read + \
           'ReadStored=Describe(result.Value2)\nEnd Function\n' + ARRAYS['helper']


@pytest.mark.parametrize('case', ARRAY_FILES, ids=lambda case: case['mode'] + ':' + case['formula'])
def test_pivotdata_array_roundtrip(case: dict[str, Any], tmp_path: Path) -> None:
    path = tmp_path / 'array.xlsx'
    source = _array_file(case, path)
    reopened = ExcelApplication.open(path, with_vba=False)
    reopened.add_module(source, name='Harness')
    assert reopened.run('ReadStored') == case['result']


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', ARRAY_FILES, ids=lambda case: case['mode'] + ':' + case['formula'])
def test_native_excel_reads_saved_pivotdata_arrays(case: dict[str, Any], tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'array.xlsx'
    source = _array_file(case, path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        result = excel.run_vba(source, 'ReadStored', line_numbers=False)
        assert result.ok and not result.dialogs, result
        assert result.value == case['result']


def _save_results(file: str, path: Path) -> list[str]:
    app = ExcelApplication.open(ROOT / file, with_vba=False)
    cases = [case for case in RECORD['probes'] if case['file'] == file]
    lines = ['Sub Store()']
    for row, case in enumerate(cases, 20):
        formula = case['formula'].replace('"', '""')
        lines.append(f'Cells({row},26).Formula="{formula}"')
    lines.append('End Sub')
    app.add_module('\n'.join(lines), name='Harness')
    app.run('Store')
    app.save(path)
    return [case['cell_result'] for case in cases]


@pytest.mark.parametrize('file', FILES)
def test_pivotdata_project_roundtrip(file: str, tmp_path: Path) -> None:
    path = tmp_path / file
    expected = _save_results(file, path)
    reopened = ExcelApplication.open(path, with_vba=False)
    actual = [reopened.evaluate(f'CStr(Cells({row},26).Value2)') for row in range(20, 20 + len(expected))]
    assert actual == expected
    with ZipFile(ROOT / file) as original, ZipFile(path) as saved:
        for part in original.namelist():
            if part.startswith(('xl/pivotTables/', 'xl/pivotCache/')):
                assert saved.read(part) == original.read(part)


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('file', FILES)
def test_native_excel_reads_saved_pivotdata(file: str, tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / file
    expected = _save_results(file, path)
    source = 'Function ReadResults() As String\nDim r As Long\n'
    source += f'For r=20 To {19 + len(expected)}\nReadResults=ReadResults & CStr(Cells(r,26).Value2) & "|"\nNext\nEnd Function'
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        result = excel.run_vba(source, 'ReadResults', line_numbers=False)
        assert result.ok and not result.dialogs, result
        assert result.value == ''.join(value + '|' for value in expected)
