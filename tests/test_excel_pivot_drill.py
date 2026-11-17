"""Range.ShowDetail on a pivot table's value, the records behind it on a new sheet (drill-through), replayed from live
Excel (scripts/measure_pivot_drill.py, tests/fixtures/pivot_drill.json): the events a sink saw, what failed, and every
sheet but the source's: its place, its table, its cells with their types and formats, and the active cell and
selection; and the table part Excel saved.

Excel ran on a 144-DPI display, where a column's standard width reads 8.09; the model's is a 96-DPI one's, so each
sheet's StandardHeight and widths are compared as one fact, that the detail's columns keep the standard width.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/pivot_drill.json').read_text(encoding='utf-8'))
CASES: list[list[Any]] = RECORD['cases']
#: Values standing for part of a table's records, which Excel lists in an order not worked out: not built.
PARTS = {0, 2, 3, 4, 6, 7, 8, 9, 10, 15, 21, 22, 25, 26, 27, 28, 29}
#: ShowDetail on a cell beside the table, which Excel refuses with 1004; the model takes it for an outline's and
#: reports outlines over part of a row unsupported.
BESIDE = {14}
_WIDTHS = re.compile(r' sh ([\d.]+) widths ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+)')
_UID = re.compile(r' xr:uid="\{[0-9A-F-]+\}"')


def _dpi_free(answer: str) -> str:
    """The answer with each sheet's StandardHeight and widths put as whether its columns keep one width."""
    def kept(found: re.Match[str]) -> str:
        return ' widths ' + ('standard' if len(set(found.groups()[1:])) == 1 else 'fitted')

    return _WIDTHS.sub(kept, answer)


def _run(index: int, folder: Path) -> tuple[str, dict[str, str]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['log'], name='EventLog')
    app.add_module(RECORD['sink'], name='ExcelSink', kind='class')
    app.add_module(RECORD['source'], name='ProbeModule')
    answer = str(app.run('ProbeModule.Probe', str(folder), index)).replace(str(folder), '<folder>')
    parts: dict[str, str] = {}
    saved = folder / f'c{index}.xlsx'
    if saved.exists():
        with zipfile.ZipFile(saved) as package:
            parts = {name: package.read(name).decode('utf-8') for name in package.namelist()
                     if re.fullmatch(r'xl/tables/table\d+\.xml', name)}
    return answer, parts


def _table(text: str) -> str:
    """A table part from its root on, without the revision markup Excel adds."""
    text = text[text.find('<table '):].replace(' />', '/>')
    return _UID.sub('', re.sub(r' xmlns:(mc|xr|xr3)="[^"]*"| mc:Ignorable="[^"]*"| xr3:uid="[^"]*"', '', text))


@pytest.mark.parametrize('index', sorted(set(range(len(CASES))) - PARTS - BESIDE), ids=str)
def test_a_values_detail_is_excels(index: int, tmp_path: Path) -> None:
    answer, parts = _run(index, tmp_path)
    assert _dpi_free(answer) == _dpi_free(CASES[index][1])
    excel = {name: _table(text) for name, text in CASES[index][2].items() if name.startswith('xl/tables/')}
    assert {name: _table(text) for name, text in parts.items()} == excel


@pytest.mark.parametrize('index', sorted(PARTS), ids=str)
def test_a_part_of_the_records_reports_itself(index: int, tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError, match='part of its records'):
        _run(index, tmp_path)


@pytest.mark.parametrize('index', sorted(BESIDE), ids=str)
def test_a_cell_beside_the_table_reports_itself(index: int, tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError, match='outlines'):
        _run(index, tmp_path)
