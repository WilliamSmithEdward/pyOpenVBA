"""A loaded chart due its automatic title, replayed from live Excel (scripts/measure_chart_loaded_title.py,
tests/fixtures/chart_loaded_title.json): what earns it, and that its part saves as Excel saved it, the title and nothing else added.
Without autoTitleDeleted a chart is taken to have had its title deleted, so nothing earns it."""

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
RECORD: dict[str, Any] = json.loads((FIXTURES / 'chart_loaded_title.json').read_text(encoding='utf-8'))
#: Excel's save writes the flag a chart was missing; a loaded part the macro did not edit stays as it came.
FLAG_ONLY = re.compile(r'<c:autoTitleDeleted val="1"/>')


def replayed(case: list[Any], tmp_path: Path) -> tuple[str, dict[str, str]]:
    name, which = case[0], case[1]
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    book = tmp_path / 'in.xlsx'
    shutil.copy(FIXTURES / name, book)
    out = tmp_path / 'out.xlsx'
    answer = str(app.run('P.Act', str(book), str(out), which))
    with zipfile.ZipFile(out) as package:
        return answer, {part: package.read(part).decode('utf-8') for part in package.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', part)}


def label(case: list[Any]) -> str:
    return f"{case[0]} {RECORD['labels'][case[1]]}"


def refused(case: list[Any]) -> bool:
    """A loaded chart's sheet cannot be copied here, and two charts cannot be selected together."""
    what = RECORD['labels'][case[1]]
    return what == 'Worksheet.Copy' or (what == 'ChartObjects.Select' and 'two' in case[0])


ANSWERED = [case for case in RECORD['cases'] if not refused(case)]


@pytest.mark.parametrize('case', ANSWERED, ids=label)
def test_each_case_answers_as_excel_did(case: list[Any], tmp_path: Path) -> None:
    assert replayed(case, tmp_path)[0] == case[2]


@pytest.mark.parametrize('case', [case for case in ANSWERED if 'no_flag' not in case[0]], ids=label)
def test_each_chart_part_is_the_bytes_excel_saved(case: list[Any], tmp_path: Path) -> None:
    assert replayed(case, tmp_path)[1] == case[3]


@pytest.mark.parametrize('case', [case for case in ANSWERED if 'no_flag' in case[0]], ids=label)
def test_a_chart_without_the_flag_earns_nothing_and_stays_as_it_came(case: list[Any], tmp_path: Path) -> None:
    ours = replayed(case, tmp_path)[1]
    assert {part: FLAG_ONLY.sub('', text) for part, text in case[3].items()} == ours


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if refused(case)], ids=label)
def test_what_is_left_out_reports_itself(case: list[Any], tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError):
        replayed(case, tmp_path)


def test_an_edit_to_a_loaded_chart_is_still_refused(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('''Public Sub Try(ByVal path As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(path)
    wb.Worksheets(1).ChartObjects(1).Chart.ChartType = xlArea
End Sub
''', name='P')
    book = tmp_path / 'in.xlsx'
    shutil.copy(FIXTURES / 'chart_loaded_title_one.xlsx', book)
    with pytest.raises(VBAUnsupportedError, match='imported chart'):
        app.run('P.Try', str(book))
