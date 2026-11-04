"""One edit at a time on a loaded chart, against what Excel answers and saves (scripts/measure_loaded_chart_edits.py,
tests/fixtures/loaded_chart_edits.json). Until each edit is built it stays refused, so nothing half-built loses what a file holds. A case the
model does answer matches Excel's answer and the bytes of its chart part, but for a chart Excel did not write: Excel rewrites that part
whole at any save, with fresh ids, and the model keeps it as it came."""

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
#: What the model answers on a loaded chart: no edit, and the title activating may earn.
ANSWERED = ('nothing', 'activate')
NOT_EXCELS = 'loaded_chart_openpyxl.xlsx'


def what(case: list[Any]) -> str:
    return RECORD['edits'][case[1]][0]


def label(case: list[Any]) -> str:
    return f'{case[0]} {what(case)}'


def run(case: list[Any], folder: Path) -> tuple[str, dict[str, bytes]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    book, out = folder / 'in.xlsx', folder / 'out.xlsx'
    shutil.copy(FIXTURES / case[0], book)
    answer = str(app.run('P.Act', str(book), str(out), case[1]))
    with zipfile.ZipFile(out) as package:
        return answer, {name: package.read(name) for name in package.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if what(case) not in ANSWERED], ids=label)
def test_an_edit_not_built_stays_refused(case: list[Any], tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError, match='imported chart'):
        run(case, tmp_path)


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if what(case) in ANSWERED], ids=label)
def test_what_is_answered_is_excels_answer(case: list[Any], tmp_path: Path) -> None:
    assert run(case, tmp_path)[0] == case[2]


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if what(case) in ANSWERED and case[0] != NOT_EXCELS], ids=label)
def test_a_chart_excel_wrote_saves_as_excel_saved_it(case: list[Any], tmp_path: Path) -> None:
    saved = {name: text.encode('utf-8') for name, text in case[3].items() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}
    assert run(case, tmp_path)[1] == saved


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if what(case) in ANSWERED and case[0] == NOT_EXCELS], ids=label)
def test_a_chart_excel_did_not_write_stays_as_it_came(case: list[Any], tmp_path: Path) -> None:
    came = {name: text.encode('utf-8') for name, text in RECORD['originals'][NOT_EXCELS].items() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}
    assert run(case, tmp_path)[1] == came
