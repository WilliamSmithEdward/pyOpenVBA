"""Native cache/table creation lifecycle and persisted reports."""
from __future__ import annotations
import json
import importlib.util
import os
from pathlib import Path
from typing import Any
import pytest
from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/pivots'
RECORD: dict[str, Any] = json.loads((ROOT / 'creation.json').read_text(encoding='utf-8'))
SAVED = [case for case in RECORD['probes'] if 'ERR' not in case['result']]
INSPECT = '''Function Snapshot() As String
Dim p As PivotTable, f As PivotField
Snapshot=CStr(ThisWorkbook.PivotCaches.Count) & "|" & CStr(ThisWorkbook.PivotCaches(1).RecordCount) & "|" & Worksheets(1).UsedRange.Address
For Each p In Worksheets(1).PivotTables
Snapshot=Snapshot & "|" & p.Name & ":" & p.TableRange1.Address & ":" & CStr(p.CacheIndex) & ":" & CStr(p.TableRange1.Cells(1,1).Value2)
For Each f In p.PivotFields
Snapshot=Snapshot & ":" & f.Name
Next
Next
End Function'''


def _add_case(app: ExcelApplication, case: dict[str, Any]) -> None:
    for name, source in case.get('classes', {}).items():
        app.add_module(source, name=name, kind='class')
    app.add_module(case['source'], name='Harness')


def _saved_project(case: dict[str, Any], path: Path) -> str:
    app = ExcelApplication()
    app.add_workbook()
    _add_case(app, case)
    assert app.run('Probe') == case['result']
    app.add_module(INSPECT, name='Inspection')
    expected = app.run('Snapshot')
    app.save(path)
    assert isinstance(expected, str)
    return expected


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_pivot_creation(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    _add_case(app, case)
    actual = app.run('Probe')
    # Native workbook numbering advances within one Excel instance.
    expected = case['result']
    if case['name'] == 'source-union':
        import re
        expected = re.sub(r'\[Book\d+\]', '[Book1]', expected)
    assert actual == expected


def test_created_pivot_roundtrip(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['probes'][0]['source'], name='Harness')
    assert app.run('Probe') == RECORD['probes'][0]['result']
    path = app.save(tmp_path / 'created.xlsx')
    reopened = ExcelApplication.open(path, with_vba=False)
    assert reopened.evaluate('ThisWorkbook.PivotCaches.Count') == 1
    assert reopened.evaluate('Worksheets(1).PivotTables(1).Name') == 'Made'
    assert reopened.evaluate('ThisWorkbook.PivotCaches(1).RecordCount') == 5
    assert reopened.evaluate('Worksheets(1).PivotTables(1).TableRange1.Address') == '$E$3:$G$20'


def test_shared_cache_after_save_and_reopen(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['probes'][0]['source'], name='Harness')
    app.run('Probe')
    path = app.save(tmp_path / 'created.xlsx')
    for current in (app, ExcelApplication.open(path, with_vba=False)):
        current.add_module('''Sub AddSecond()
Worksheets(1).Range("A1").Value = "Changed"
Worksheets(1).Range("J3:L20").Value = 123
ThisWorkbook.PivotCaches(1).CreatePivotTable Worksheets(1).Range("J3"), "Second"
End Sub''', name='AddReport')
        current.run('AddSecond')
        assert current.evaluate('Worksheets(1).PivotTables(1).Name') == 'Second'
        assert current.evaluate('Worksheets(1).PivotTables(1).PivotFields(1).Name') == 'Region'
        assert current.evaluate('IsEmpty(Worksheets(1).Range("J3").Value)') is True
        output = current.save(tmp_path / ('shared-' + str(id(current)) + '.xlsx'))
        for _ in range(2):
            reopened = ExcelApplication.open(output, with_vba=False)
            assert reopened.evaluate('ThisWorkbook.PivotCaches.Count') == 1
            assert reopened.evaluate('Worksheets(1).PivotTables(1).Name') == 'Second'
            assert reopened.evaluate('Worksheets(1).PivotTables(2).Name') == 'Made'
            output = reopened.save(output)


@pytest.mark.parametrize('case', SAVED, ids=lambda case: case['name'])
def test_created_project_roundtrip(case: dict[str, Any], tmp_path: Path) -> None:
    path = tmp_path / 'created.xlsm'
    expected = _saved_project(case, path)
    assert ExcelApplication.open(path).run('Snapshot') == expected


def test_creation_preserves_unrecognized_parts_and_relationships(tmp_path: Path) -> None:
    from pyopenvba.apps.excel._charts_file import xml_tree
    from pyopenvba.powerquery._sheets import add_relationship

    app = ExcelApplication()
    book = app.add_workbook()
    app.save(tmp_path / 'empty.xlsx')
    package = book.package
    records_part = 'xl/pivotCache/pivotCacheRecords1.xml'
    records = b'<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>'
    package.write(records_part, records)
    report_rels = 'xl/pivotTables/_rels/pivotTable1.xml.rels'
    relationships = b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    package.write(report_rels, relationships)
    sheet_rels = 'xl/worksheets/_rels/sheet1.xml.rels'
    target = '../pivotTables/unrecognized.xml'
    identifier = add_relationship(package, sheet_rels,
                                  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable', target)
    _add_case(app, RECORD['probes'][0])
    app.run('Probe')
    path = app.save(tmp_path / 'created.xlsx')
    assert package.read(records_part) == records
    assert package.read(report_rels) == relationships
    original = next(entry for entry in xml_tree(package.read(sheet_rels)) if entry.get('Id') == identifier)
    assert original.get('Target') == target
    reopened = ExcelApplication.open(path, with_vba=False)
    assert reopened.evaluate('Worksheets(1).PivotTables.Count') == 1
    assert reopened.evaluate('Worksheets(1).PivotTables(1).Name') == 'Made'
    assert reopened.evaluate('ThisWorkbook.PivotCaches(1).RecordCount') == 5


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', SAVED, ids=lambda case: case['name'])
def test_native_excel_opens_created_project(case: dict[str, Any], tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'created.xlsm'
    expected = _saved_project(case, path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('Inspection.Snapshot') == expected
        if case['name'] == 'default' or case['name'].startswith('shared-'):
            # Exercise the generated record types, not just the report metadata.
            result = excel.run_vba('''Function AggregateCreated() As String
Dim p As PivotTable
Set p=Worksheets(1).PivotTables(1)
p.AddDataField p.PivotFields("Amount"), "Total", xlSum
AggregateCreated=CStr(p.GetPivotData("Total").Value2)
End Function''', 'AggregateCreated', line_numbers=False)
            assert result.ok and result.value == '150' and not result.dialogs, result
        assert not any(event.get('dialogs') for event in excel.events)


def _mixed_project(path: Path) -> str:
    case = next(case for case in SAVED if case['name'] == 'edge-merge-inner')
    _saved_project(case, path)
    app = ExcelApplication.open(path)
    app.add_module('''Sub AddArrays()
Range("J1").Formula2="=SEQUENCE(2)"
Range("M2").Value=99
Range("M1").Formula2="=SEQUENCE(2)"
End Sub
Function MixedSnapshot() As String
MixedSnapshot=CStr(Range("J1").Value2) & "|" & CStr(Range("J2").Value2) & "|" & CStr(Range("M1").Value2) & "|" & CStr(Range("M2").Value2) & "|" & CStr(Range("E3").Value2)
End Function''', name='Mixed')
    app.run('AddArrays')
    expected = app.run('MixedSnapshot')
    assert expected == '1|2|Error 2045|99|Error 2045'
    assert isinstance(expected, str)
    app.save(path)
    return expected


def test_blocked_pivot_and_dynamic_errors_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / 'mixed.xlsm'
    expected = _mixed_project(path)
    assert ExcelApplication.open(path).run('MixedSnapshot') == expected


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
def test_native_excel_opens_pivot_and_dynamic_errors(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'mixed.xlsm'
    expected = _mixed_project(path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('Mixed.MixedSnapshot') == expected
        assert not any(event.get('dialogs') for event in excel.events)
