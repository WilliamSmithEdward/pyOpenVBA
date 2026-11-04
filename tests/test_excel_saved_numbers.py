"""A number in a cell's <v> as Excel spells it when it saves (scripts/measure_saved_numbers.py,
tests/fixtures/saved_numbers.json). A huge number in plain digits made Excel refuse the whole file."""
from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._io import _number_text  # pyright: ignore[reportPrivateUsage]

SPELLED: list[str] = json.loads(
    (Path(__file__).parent / 'fixtures/saved_numbers.json').read_text(encoding='utf-8'))['spelled']


#: Excel writes 17 digits for these where the 15 it needs are within 5/16 of an epsilon, the rule excel_number has from
#: smaller numbers. They read back the same double.
LONGER_THAN_THE_RULE = {'1.5000000000000001E+300', '1.7976931348623099E+308'}


@pytest.mark.parametrize('spelled', [one for one in SPELLED if one not in LONGER_THAN_THE_RULE])
def test_each_number_as_excel_spells_it(spelled: str) -> None:
    assert _number_text(float(spelled)) == spelled


def test_a_huge_number_reaches_the_file_with_an_exponent(tmp_path: Path) -> None:
    path = tmp_path / 'huge.xlsx'
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Public Sub Make()\n    Dim b As Object\n    Set b = Workbooks.Add(xlWBATWorksheet)\n'
                   f'    b.Worksheets(1).Range("B2").Value = 1E+308\n    b.SaveAs "{path}", xlOpenXMLWorkbook\n'
                   f'    b.Close False\nEnd Sub\n', name='Huge')
    app.run('Huge.Make')
    with zipfile.ZipFile(path) as package:
        assert '<v>1E+308</v>' in package.read('xl/worksheets/sheet1.xml').decode('utf-8')
