"""One edit at a time on a loaded chart, against what Excel answers and saves (scripts/measure_loaded_chart_edits.py,
tests/fixtures/loaded_chart_edits.json). Until each edit is built it stays refused, so nothing half-built loses what a file holds. A case the
model does answer matches Excel's answer and the bytes of its chart part, but for a chart Excel did not write: Excel rewrites that part
whole at any save, with fresh ids, and the model keeps it as it came.

For whoever writes the serializer those charts wait for: Excel gives each series a c16:uniqueId of {0000000N-XXXX-4XXX-[89AB]XXX-XXXXXXXXXXXX}.
The first 8 digits are the series number and do not vary. The other 24 are shared by a chart's series, new at every save, and show a version 4
GUID's pattern in all 37 saves measured (24 here, 13 repeated in new sessions, in one session, and with A1 or C3 selected). Mask those 24 and
nothing else. A selected range is still to be tried. tests/fixtures/chart_markup.json's selection rule sets the first 8, on a chart made in
a session."""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_chart_edits.json').read_text(encoding='utf-8'))
#: What the model answers on any loaded chart: no edit, and the title activating may earn.
ANSWERED = ('nothing', 'activate', 'legend position read')
#: What is built for a chart Excel wrote. A chart it did not write refuses these, being rewritten whole by Excel's save.
BUILT = ('HasTitle True', 'HasTitle False', 'HasTitle True then False', 'title delete', 'HasTitle False then True', 'HasTitle True False True',
         'title gone then read', 'title gone then activate', 'new title text read', 'title delete then True',
         'HasLegend False', 'HasLegend True', 'legend position', 'HasLegend False then True', 'legend gone position', 'legend back position',
         'legend position bottom', 'legend position same', 'legend moved position read',
         'axis maximum', 'axis minimum', 'axis both', 'axis maximum twice', 'axis maximum auto', 'axis maximum auto read',
         'axis gridlines', 'gridlines back', 'category gridlines', 'minor gridlines', 'minor gridlines off',
         'series name', 'series name formula', 'series name cell')
#: Built, but not for a chart with a style: new gridlines take the style's look, which is not read.
NEEDS_NO_STYLE = ('category gridlines', 'minor gridlines', 'minor gridlines off')
NOT_EXCELS = 'loaded_chart_openpyxl.xlsx'


def what(case: list[Any]) -> str:
    return RECORD['edits'][case[1]][0]


def label(case: list[Any]) -> str:
    return f'{case[0]} {what(case)}'


def answered(case: list[Any]) -> bool:
    if what(case) in NEEDS_NO_STYLE and 'dressed' in case[0]:
        return False
    return what(case) in ANSWERED or (what(case) in BUILT and case[0] != NOT_EXCELS)


def run(case: list[Any], folder: Path) -> tuple[str, dict[str, bytes]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    book, out = folder / 'in.xlsx', folder / 'out.xlsx'
    shutil.copy(FIXTURES / case[0], book)
    answer = str(app.run('P.Act', str(book), str(out), case[1]))
    with zipfile.ZipFile(out) as package:
        return answer, {name: package.read(name) for name in package.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if not answered(case)], ids=label)
def test_an_edit_not_built_stays_refused(case: list[Any], tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError):
        run(case, tmp_path)


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if answered(case)], ids=label)
def test_what_is_answered_is_excels_answer(case: list[Any], tmp_path: Path) -> None:
    assert run(case, tmp_path)[0] == case[2]


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if answered(case) and case[0] != NOT_EXCELS], ids=label)
def test_a_chart_excel_wrote_saves_as_excel_saved_it(case: list[Any], tmp_path: Path) -> None:
    saved = {name: text.encode('utf-8') for name, text in case[3].items() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}
    assert run(case, tmp_path)[1] == saved


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if what(case) in ANSWERED and case[0] == NOT_EXCELS], ids=label)
def test_a_chart_excel_did_not_write_stays_as_it_came(case: list[Any], tmp_path: Path) -> None:
    came = {name: text.encode('utf-8') for name, text in RECORD['originals'][NOT_EXCELS].items() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}
    assert run(case, tmp_path)[1] == came
