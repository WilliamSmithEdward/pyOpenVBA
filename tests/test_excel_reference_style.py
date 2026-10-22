"""Application.ReferenceStyle, measured by scripts/measure_reference_style.py.

The style reads xlA1 (1) or xlR1C1 (-4150) and any other number leaves it
as it was. Formula, FormulaLocal and Address read as A1 in either style;
Evaluate reads R1C1 references, relative ones from A1, only in the R1C1
style, where A1 references still read, and in the A1 style R1C1 text is
unreadable, #VALUE!. A workbook saved in the R1C1 style says so with
refMode on calcPr; one opened while another is open leaves the
application's style alone.
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/reference_style.json').read_text(encoding='utf-8'))


def test_the_style_and_what_it_changes() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['runtime']['macro'], name='Harness')
    assert app.run('Harness.Probe') == RECORD['runtime']['result']


def _calc_pr(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        workbook = package.read('xl/workbook.xml').decode('utf-8')
    found = re.search(r'<calcPr\b[^>]*/?>', workbook)
    assert found is not None
    return found.group(0)


def _refmode(element: str) -> str | None:
    found = re.search(r'\brefMode="([^"]*)"', element)
    return found.group(1) if found else None


def _set_style(app: ExcelApplication, style: str) -> None:
    app.add_module(f'Sub SetStyle()\n    Application.ReferenceStyle = {style}\nEnd Sub\n', name='Style')
    app.run('Style.SetStyle')


def test_a_workbook_keeps_the_style_it_was_saved_in(tmp_path: Path) -> None:
    saved = RECORD['file']['calcPr']
    app = ExcelApplication()
    app.add_workbook()
    _set_style(app, 'xlR1C1')
    r1c1 = app.save(tmp_path / 'style-r1c1.xlsx')
    assert _refmode(_calc_pr(r1c1)) == _refmode(saved['r1c1']) == 'R1C1'
    _set_style(app, 'xlA1')
    a1 = app.save(tmp_path / 'style-a1.xlsx')
    assert _refmode(_calc_pr(a1)) is _refmode(saved['a1']) is None


def test_the_first_workbook_brings_its_style_and_a_later_one_does_not(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    _set_style(app, 'xlR1C1')
    path = app.save(tmp_path / 'style-r1c1.xlsm')
    reopened = ExcelApplication.open(path)
    # As the first workbook opened it gives the application its style, as it does the calculation mode.
    assert reopened.evaluate('Application.ReferenceStyle') == -4150
    first = ExcelApplication()
    first.add_workbook()
    first.open_workbook(path)
    # Opened after another, as Excel opened it in the probe, it leaves the style at A1.
    assert first.evaluate('Application.ReferenceStyle') == 1
    assert RECORD['file']['opened'] == '1|1|1'
