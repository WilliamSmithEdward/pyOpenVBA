"""Statistical conditional rules retain their settings and rendering through sheet copies."""
from __future__ import annotations

import os
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._model import Workbook

CASES = [
    ('Top10', 'AddTop10', 'f.Rank=2', [1, 2, 3, 4], '0011'),
    ('AboveAverage', 'AddAboveAverage', '', [1, 2, 3, 4], '0011'),
    ('UniqueValues', 'AddUniqueValues', '', [1, 2, 2, 3], '1001'),
    ('UniqueValues', 'AddUniqueValues', 'f.DupeUnique=xlDuplicate', [1, 2, 2, 3], '0110'),
]
INSPECT = '''Function InspectStatistical() As String
Dim s As Worksheet,c As Range,f As Object
For Each s In Worksheets
Set f=s.Range("A1").FormatConditions(1)
InspectStatistical=InspectStatistical & s.Name & ":" & TypeName(f) & ":"
For Each c In s.Range("A1:A4")
If c.DisplayFormat.Font.Bold Then
InspectStatistical=InspectStatistical & "1"
Else
InspectStatistical=InspectStatistical & "0"
End If
Next
InspectStatistical=InspectStatistical & ":" & s.Range("A4").Text & "|"
Next
End Function'''


def copied_project(method: str, settings: str, values: list[int], external: bool, path: Path) -> str:
    app = ExcelApplication()
    app.add_workbook()
    assignments = '\n'.join(f'Range("A{row}").Value={value}' for row, value in enumerate(values, 1))
    app.add_module('Sub MakeStatistical()\nDim f As ' + method.removeprefix('Add') + '''
ActiveSheet.Name="Data"
Range("A1:A4").ColumnWidth=30
''' + assignments + f'\nSet f=Range("A1:A4").FormatConditions.{method}\n' + settings + '''
f.Font.Bold=True
f.NumberFormat="0.00 ""match"""
End Sub''', name='SetupStatistical')
    app.run('MakeStatistical')
    app.add_module('Sub CopyStatistical()\n' + ('ActiveSheet.Copy' if external else
                   'ActiveSheet.Copy After:=ActiveSheet') + '\nEnd Sub', name='CopyStatisticalModule')
    app.run('CopyStatistical')
    # Standard modules are supplied explicitly to the newly active workbook.
    app.add_module(INSPECT, name='StatisticalInspection')
    if external:
        target = app.application.ActiveWorkbook()
        assert isinstance(target, Workbook)
        app.add_module(INSPECT, name='StatisticalInspection', workbook=target)
    expected = app.run('InspectStatistical')
    assert isinstance(expected, str)
    app.save(path)
    assert ExcelApplication.open(path).run('InspectStatistical') == expected
    return expected


@pytest.mark.parametrize('external', [False, True])
@pytest.mark.parametrize(('kind', 'method', 'settings', 'values', 'bits'), CASES)
def test_statistical_sheet_copy_and_roundtrip(kind: str, method: str, settings: str,
                                            values: list[int], bits: str, external: bool, tmp_path: Path) -> None:
    result = copied_project(method, settings, values, external, tmp_path / 'statistical-copy.xlsm')
    text = f'{values[-1]}.00 match' if bits[-1] == '1' else str(values[-1])
    wanted = f'Data:{kind}:{bits}:{text}|'
    if not external:
        wanted += f'Data (2):{kind}:{bits}:{text}|'
    assert result == wanted


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1', reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize(('kind', 'method', 'settings', 'values', 'bits'), CASES)
def test_native_reopens_copied_statistical_rules(kind: str, method: str, settings: str,
                                                values: list[int], bits: str, tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'native-statistical-copy.xlsm'
    expected = copied_project(method, settings, values, True, path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        result = excel.run_raw('StatisticalInspection.InspectStatistical')
        assert result == expected
        assert not any(event.get('dialogs') for event in excel.events)
