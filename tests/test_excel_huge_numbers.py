"""Huge numbers, an unallocated array and awkward counts answer as Excel did, each line bound late and typed (scripts/measure_excel_huge_numbers.py), where some once let a
Python error out and some never returned. The answers show an array's size rather than its cells, so nothing here builds more than Excel measured."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_excel_huge_numbers', ROOT / 'scripts/measure_excel_huge_numbers.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
RECORD = json.loads((ROOT / 'tests/fixtures/excel_huge_numbers.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('binding', ['late', 'typed'])
def test_each_line_answers_as_excel_did(binding: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(MEASURE.macro(typed=binding == 'typed'), name='H')
    found = str(app.run('H.Probe')).split(MEASURE.SEPARATOR)
    wanted = RECORD['answers'][binding].split(MEASURE.SEPARATOR)
    assert len(found) == len(wanted)
    assert not [(asked[0], excel, model) for asked, excel, model in zip(MEASURE.LINES, wanted, found) if excel != model]


@pytest.mark.parametrize('size', [5001, 99999])
@pytest.mark.parametrize('how', ['WorksheetFunction.MUnit', 'Application.MUnit'])
def test_munit_between_what_was_measured_refuses_openly(how: str, size: int) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Function Probe()\n    Dim x As Variant\n    x = {how}({size})\nEnd Function\n', name='H')
    with pytest.raises(VBAUnsupportedError):
        app.run('H.Probe')


def test_a_cell_asking_for_a_vast_munit_refuses_rather_than_building_it() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Function Probe()\n    Range("A1").Formula = "=SUM(MUNIT(99999))"\n    Probe = Range("A1").Value\nEnd Function\n', name='H')
    with pytest.raises(VBAUnsupportedError):
        app.run('H.Probe')


@pytest.mark.parametrize('call', ['ActiveWorkbook.SaveAs ""', 'ActiveWorkbook.Close True, ""', 'Worksheets(1).SaveAs ""'])
def test_a_blank_file_name_refuses_openly(call: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Not measured: Excel would save under a name it chooses, and no probe should write one. The refusal touches no file."""
    monkeypatch.chdir(tmp_path)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Sub Probe()\n    {call}\nEnd Sub\n', name='H')
    with pytest.raises(VBAUnsupportedError):
        app.run('H.Probe')
    assert not list(tmp_path.iterdir())
