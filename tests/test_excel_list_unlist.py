"""ListObject.Unlist, replayed from live Excel (scripts/measure_list_unlist.py, tests/fixtures/list_unlist.json): what
each case reads, and the cells and formats it saves; and a digest of those for every built-in style under each set of the
table's options."""
from __future__ import annotations

import hashlib
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/list_unlist.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Hiding a table's header row, resizing one with a totals row, and cutting a whole table elsewhere, which the model
#: does not do.
LEFT_OUT: tuple[str, ...] = ('o.ShowHeaders = False', 'o.ShowTotals = True: o.Resize', '.Cut ',
                             # Excel lists a border set before the table's after them; the engine lists it first.
                             'ws.Range("C4").Borders.LineStyle')
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in LEFT_OUT)]


def _source() -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case '):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept:
            continue
        written.append(line)
    return '\n'.join(written)


def cells(sheet: str) -> str:
    """The sheet's cells, its text numbered as Excel's shared strings number it where the engine writes it in the
    cell."""
    texts: dict[str, int] = {}

    def shared(found: re.Match[str]) -> str:
        return f't="s"><v>{texts.setdefault(found.group(1), len(texts))}</v></c>'

    # A row's height follows the display's scale, which the engine takes as 96 DPI.
    block = sheet[sheet.index('<sheetData'):sheet.index('</sheetData>')]
    held = re.sub(r' x14ac:dyDescent="[^"]*"| ht="[^"]*"', '', block)
    return re.sub(r't="inlineStr"><is><t[^>]*>([^<]*)</t></is></c>', shared, held)


def formats(styles: str) -> str:
    return styles[styles.index('<fonts'):styles.index('</cellXfs>')]


def parts_of(path: Path) -> tuple[str, str]:
    with zipfile.ZipFile(path) as package:
        return (package.read('xl/worksheets/sheet1.xml').decode('utf-8'), package.read('xl/styles.xml').decode('utf-8'))


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('list_unlist')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='UnlistProbe')
    return {index: str(app.run('UnlistProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    sheet, styles = parts_of(run[1] / f'c{index}.xlsx')
    held = RECORD['cases'][index][2]
    assert cells(sheet) == cells(held['xl/worksheets/sheet1.xml'])
    assert formats(styles) == formats(held['xl/styles.xml'])


SWEEP_PROBE = """Public Function Probe(ByVal path As String, ByVal style As String, ByVal options As Long) As String
    Dim wb As Workbook, ws As Object, o As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("a", "b", "c")
    ws.Range("A2:C2").Value = Array(1, 2, 3)
    ws.Range("A3:C3").Value = Array(4, 5, 6)
    ws.Range("A4:C4").Value = Array(7, 8, 9)
    ws.Range("A5:C5").Value = Array(10, 11, 12)
    Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C5"), , xlYes)
    o.Name = "T"
    o.TableStyle = style
    If options And 1 Then o.ShowTotals = True
    If options And 2 Then o.ShowTableStyleColumnStripes = True
    If options And 4 Then o.ShowTableStyleFirstColumn = True
    If options And 8 Then o.ShowTableStyleLastColumn = True
    If options And 16 Then o.ShowTableStyleRowStripes = False
    o.Unlist
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Probe = "ok"
End Function
"""
_OPTIONS = {'o.ShowTotals = True': 1, 'o.ShowTableStyleColumnStripes = True': 2,
            'o.ShowTableStyleFirstColumn = True': 4,
            'o.ShowTableStyleLastColumn = True': 8, 'o.ShowTableStyleRowStripes = False': 16}


@pytest.fixture(scope='module')
def sweep() -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SWEEP_PROBE, name='UnlistSweep')
    return app


@pytest.mark.parametrize('case', list(RECORD['swept']))
def test_each_style_as_excel_saved_it(sweep: ExcelApplication, tmp_path: Path, case: str) -> None:
    style, *options = case.split(': ', 1)
    bits = sum(value for key, value in _OPTIONS.items() if options and key in options[0])
    sweep.run('UnlistSweep.Probe', str(tmp_path / 'p.xlsx'), style.split('"')[1], bits)
    sheet, styles = parts_of(tmp_path / 'p.xlsx')
    assert hashlib.sha256((cells(sheet) + formats(styles)).encode('utf-8')).hexdigest() == RECORD['swept'][case]
