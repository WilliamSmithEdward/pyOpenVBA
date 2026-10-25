"""Imported PivotTable runtime views compared with an Excel-authored workbook."""
from __future__ import annotations

import json
import importlib.util
import os
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures/pivots'
RECORD: dict[str, Any] = json.loads((ROOT / 'import.json').read_text(encoding='utf-8'))
# This existing layout method is deliberately unimplemented, rather than pretending
# that layout edits can succeed without rewriting the report and its XML.
PROBES = [case for case in RECORD['probes'] if not case['expression'].endswith('.RowAxisLayout')]
AXES: list[dict[str, Any]] = json.loads((ROOT / 'axes.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', PROBES, ids=lambda case: case['expression'])
def test_native_imported_pivot(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / 'native-simple.xlsx', with_vba=False)
    app.add_module(case['source'], name='Harness')
    result = app.run(case['source'].split('(')[0].split()[-1])
    # The fixture intentionally records the original workbook's name.
    expected = case['result'].replace('pivot-preview.xlsx', 'native-simple.xlsx')
    assert result == expected


@pytest.mark.parametrize('variant,case', [(variant, case) for variant in AXES for case in variant['probes']],
                         ids=[variant['file'] + ':' + case['expression'] for variant in AXES for case in variant['probes']])
def test_native_pivot_report_variants(variant: dict[str, Any], case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / variant['file'], with_vba=False)
    app.add_module(case['source'], name='Harness')
    assert app.run(case['source'].split('(')[0].split()[-1]) == case['result']


def test_pivot_collections_enumerate_reports_fields_and_items() -> None:
    app = ExcelApplication.open(ROOT / 'native-simple.xlsx', with_vba=False)
    app.add_module('''Function Probe() As String
Dim p As PivotTable, f As PivotField, i As PivotItem
For Each p In Worksheets(1).PivotTables
For Each f In p.RowFields
For Each i In f.PivotItems
Probe=Probe & p.Name & ":" & f.Name & ":" & i.Name & "|"
Next i
Next f
Next p
End Function''', name='Harness')
    assert app.run('Probe') == 'Summary:Region:East|Summary:Region:West|'


def test_blank_workbook_has_empty_pivot_collections() -> None:
    app = ExcelApplication()
    app.add_workbook()
    assert app.evaluate('ThisWorkbook.PivotCaches.Count') == 0
    assert app.evaluate('Worksheets(1).PivotTables.Count') == 0


def test_unimplemented_pivot_source_and_layout_mutations_remain_explicitly_unsupported() -> None:
    app = ExcelApplication.open(ROOT / 'native-simple.xlsx', with_vba=False)
    with pytest.raises(VBAUnsupportedError):
        app.evaluate('ThisWorkbook.PivotCaches.Create(2)')



def test_imported_simple_pivot_refreshes_edited_source_and_persists_result(tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / 'native-simple.xlsx', with_vba=False)
    pivot = 'Worksheets(1).PivotTables(1)'
    assert app.evaluate(pivot + '.GetPivotData("Amount").Value2') == 150
    app.workbook.sheets_[0].Range('B2').vba_set('Value2', 110)
    # Source edits do not change the cached report before explicit refresh.
    assert app.evaluate(pivot + '.GetPivotData("Amount").Value2') == 150
    assert app.evaluate(pivot + '.RefreshTable()') is True
    assert app.evaluate(pivot + '.GetPivotData("Amount").Value2') == 250
    result = app.save(tmp_path / 'legacy-refreshed.xlsx')
    reopened = ExcelApplication.open(result, with_vba=False)
    assert reopened.evaluate(pivot + '.GetPivotData("Amount").Value2') == 250


def test_pivot_parts_survive_cell_edit_and_reopen(tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / 'native-simple.xlsx', with_vba=False)
    app.add_module('Sub Edit()\nRange("J10").Value2=42\nEnd Sub', name='Harness')
    app.run('Edit')
    saved = app.save(tmp_path / 'saved.xlsx')
    with ZipFile(ROOT / 'native-simple.xlsx') as original, ZipFile(saved) as result:
        parts = [part for part in original.namelist() if part.startswith(('xl/pivotTables/', 'xl/pivotCache/'))]
        assert len(parts) == 5
        for part in parts:
            assert result.read(part) == original.read(part)
    reopened = ExcelApplication.open(saved, with_vba=False)
    assert reopened.evaluate('Worksheets(1).PivotTables("Summary").DataBodyRange.Cells(1,1).Value2') == 90
    assert reopened.evaluate('ThisWorkbook.PivotCaches(1).RecordCount') == 5
    assert reopened.evaluate('Range("J10").Value2') == 42


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='set RUN_LIVE_EXCEL=1 with pyvbaharness and native Excel')
@pytest.mark.parametrize('file,expected', [('native-simple.xlsx', '1|1|Summary|90'),
                                         ('native-column.xlsx', '1|1|Summary|90'),
                                         ('native-page.xlsx', '1|1|Summary|150'),
                                         ('native-multiple-data.xlsx', '1|1|Summary|90'),
                                         ('native-shared-cache.xlsx', '2|1|Other|90')])
def test_native_excel_reopens_model_saved_pivot(tmp_path: Path, file: str, expected: str) -> None:
    harness = pytest.importorskip('pyvbaharness')

    app = ExcelApplication.open(ROOT / file, with_vba=False)
    saved = app.save(tmp_path / 'saved.xlsx')
    source = '''Function Probe() As String
Probe=CStr(Worksheets(1).PivotTables.Count) & "|" & CStr(ThisWorkbook.PivotCaches.Count) & "|" & Worksheets(1).PivotTables(1).Name & "|" & CStr(Worksheets(1).PivotTables(1).DataBodyRange.Cells(1,1).Value2)
End Function'''
    with harness.ExcelSession() as excel:
        excel.open_document(saved)
        result = excel.run_vba(source, 'Probe', line_numbers=False)
        assert result.ok and not result.dialogs, result
        assert result.value == expected
