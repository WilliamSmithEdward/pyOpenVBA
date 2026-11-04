"""How a value fits where it goes, against live Excel (scripts/measure_excel_fits.py)."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).resolve().parents[1]
RECORD = json.loads((ROOT / 'tests/fixtures/excel_fits.json').read_text(encoding='utf-8'))
ANSWERS: dict[str, str] = RECORD['answers']


@pytest.mark.parametrize('key', sorted(ANSWERS))
def test_a_value_fits_as_excel_fitted_it(key: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['helpers'] + RECORD['sources'][key], name='Harness')
    try:
        assert app.run('Harness.Probe') == ANSWERS[key]
    except VBAUnsupportedError:
        pytest.fail('reports itself as not implemented')


#: What Excel saved after each Range.Sort, without the namespace it declares on the element.
SAVED = {
    'ws.Range("A1:C5").Sort ws.Range("A1"), xlAscending, Header:=xlYes': 'ref="A2:C5"><sortCondition ref="A1"/>',
    'ws.Range("A1:C5").Sort ws.Range("A1"), xlAscending, ws.Range("B1"), , xlDescending, Header:=xlYes': 'ref="A2:C5"><sortCondition ref="A1"/><sortCondition descending="1" ref="B1"/>',
    'ws.Range("A1:C5").Sort ws.Range("A1"), Orientation:=xlLeftToRight, Header:=xlNo': 'columnSort="1" ref="A1:C5"><sortCondition ref="A1"/>',
    'ws.Range("A1:C5").Sort ws.Range("A1"), Header:=xlYes, MatchCase:=True': 'caseSensitive="1" ref="A2:C5"><sortCondition ref="A1"/>',
}


@pytest.mark.parametrize('line', SAVED)
def test_a_range_sort_is_saved_as_excel_saved_it(line: str, tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Sub Probe()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A1:C1").Value = Array("x", "y", "z")
    ws.Range("A2:C2").Value = Array(3, 1, 2): ws.Range("A3:C3").Value = Array(1, 2, 3): ws.Range("A4:C4").Value = Array(2, 3, 1): ws.Range("A5:C5").Value = Array(4, 4, 4)
    {line}
End Sub''', name='Harness')
    app.run('Harness.Probe')
    path = tmp_path / 'saved.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        text = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<sortState (.*?)</sortState>', text)
    assert found is not None
    assert re.sub(r' xmlns:\w+="[^"]*"', '', found.group(1)) == SAVED[line]
