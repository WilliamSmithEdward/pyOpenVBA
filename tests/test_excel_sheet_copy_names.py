"""What a copied sheet is called when its name is taken (scripts/measure_sheet_copy_names.py, tests/fixtures/sheet_copy_names.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/sheet_copy_names.json').read_text(encoding='utf-8'))
#: Excel gives a number this long no successor, at a length not measured.
REFUSED = [['S (99999)']]


def _copied(names: list[str]) -> str:
    made = ''.join(f'    wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = "{name}"\n' for name in names[1:])
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'''Public Function Go() As String
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "{names[0]}"
{made}    On Error Resume Next
    wb.Worksheets(wb.Worksheets.Count).Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Go = wb.Worksheets.Count & "|" & wb.Worksheets(wb.Worksheets.Count).Name & "|" & Err.Number
End Function
''', name='Names')
    return str(app.run('Names.Go'))


@pytest.mark.parametrize(('names', 'answer'), [case for case in RECORD['cases'] if case[0] not in REFUSED], ids=lambda value: str(value))
def test_each_copy_is_named_as_excel_named_it(names: list[str], answer: str) -> None:
    assert _copied(names) == answer


@pytest.mark.parametrize('names', REFUSED)
def test_a_number_too_long_to_follow_says_so(names: list[str]) -> None:
    with pytest.raises(VBAUnsupportedError):
        _copied(names)
