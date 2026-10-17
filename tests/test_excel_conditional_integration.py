"""Conditional formats participate in clearing and ordinary worksheet copies."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._conditional import read_rules

ROOT = Path(__file__).parent / 'fixtures/conditional'
CASES: list[dict[str, Any]] = json.loads((ROOT / 'integration.json').read_text(encoding='utf-8'))['probes']


@pytest.mark.parametrize('kind', ['dataBar', 'colorScale', 'iconSet'])
def test_imported_visual_rule_does_not_break_text_number_format(kind: str, tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('''Function ReadText() As String
Range("A1").Value=12
Range("A1").NumberFormat="0.00"
ReadText=Range("A1").Text
End Function''', name='Harness')
    sheet = book.sheets_[0]
    sheet.conditional_formats = read_rules(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<conditionalFormatting sqref="A1"><cfRule type="{kind}" priority="1"/>'
        '</conditionalFormatting></worksheet>')
    sheet.conditional_changed = True
    assert app.run('ReadText') == '12.00'
    reopened = ExcelApplication.open(app.save(tmp_path / 'visual.xlsm'))
    assert reopened.run('ReadText') == '12.00'


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_clear_and_sheet_copy(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']


def test_copied_rule_styles_and_formulas_survive_roundtrip(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Function Create() As String
Dim f As FormatCondition
Worksheets(1).Name="Source"
Set f=Range("A1:A5").FormatConditions.Add(xlExpression,Formula1:="=A1>0")
f.Font.Bold=False
f.Font.Italic=True
f.Interior.Color=RGB(255,0,0)
f.NumberFormat="0.00"
Worksheets(1).Copy
Create=ActiveSheet.Name
End Function
Function Inspect() As String
Dim f As FormatCondition
Set f=Worksheets(1).Range("A1").FormatConditions(1)
Inspect=f.Formula1 & "|" & CStr(f.Font.Bold) & "|" & CStr(f.Font.Italic) & "|" & CStr(f.Interior.Color) & "|" & f.NumberFormat
End Function''', name='Harness')
    assert app.run('Create') == 'Source'
    # Save the new active workbook and inspect its rules independently.
    path = app.save(tmp_path / 'copied.xlsm')
    result = ExcelApplication.open(path)
    result.add_module('''Function InspectCopied() As String
Dim f As FormatCondition
Set f=Worksheets(1).Range("A1").FormatConditions(1)
InspectCopied=f.Formula1 & "|" & CStr(f.Font.Bold) & "|" & CStr(f.Font.Italic) & "|" & CStr(f.Interior.Color) & "|" & f.NumberFormat
End Function''', name='Inspection')
    assert result.run('InspectCopied') == '=A1>0|False|True|255|0.00'
