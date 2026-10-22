"""What-if data tables, measured by scripts/measure_data_tables.py.

Range.Table makes a table of a block's cells but its first row and column:
each is the block's formula worked out with the values its row and column
try put in the input cells, everything reading them worked out again and
then put back. Its cells say =TABLE(), are an array to HasArray and
CurrentArray, are guarded as an array formula is, move with rows, columns
and shifted cells only whole, and copy as their values; in semi-automatic
calculation they wait for a calculation asked for. A file keeps the table
on its first cell, every cell's value beside it.
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/data_tables.json').read_text(encoding='utf-8'))
PROBES: dict[str, dict[str, Any]] = {probe['name']: probe for probe in RECORD['probes']}
#: The macro that builds the probes' model and its three tables, from the first probe's module.
BUILD = PROBES['values']['macro'].split('\nFunction Probe()')[0]

#: Probes the model does not answer as Excel does, and why.
GAPS: dict[str, str] = {}


@pytest.mark.parametrize('name', [
    pytest.param(name, marks=[pytest.mark.xfail(reason=GAPS[name], strict=True)] if name in GAPS else [])
    for name in PROBES])
def test_a_data_table(name: str) -> None:
    probe = PROBES[name]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    if 'error' in probe:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Harness.Probe')
        assert raised.value.number == probe['error']
    else:
        assert app.run('Harness.Probe') == probe['result']


_CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"[^>]*?(?:/>|>.*?</c>)', re.DOTALL)


def _saved(tmp_path: Path, name: str) -> tuple[dict[str, str], str]:
    """The cells and calcPr of the workbook Build fills, saved by the model after what Excel's was saved after."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(BUILD + f'\nSub Save()\nDim wb As Workbook\nSet wb = ActiveWorkbook\nBuild wb.Worksheets(1)\n'
                   f'{RECORD["files"][name]["step"]}\nEnd Sub\n', name='Harness')
    app.run('Harness.Save')
    saved = app.save(tmp_path / f'{name}.xlsx')
    with zipfile.ZipFile(saved) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
        workbook = package.read('xl/workbook.xml').decode('utf-8')
    found = re.search(r'<calcPr\b[^>]*/?>', workbook)
    return {match.group(1): match.group(0) for match in _CELL.finditer(sheet)}, found.group(0) if found else ''


def _parts(xml: str) -> tuple[str, str, str]:
    """A cell's type, its <f> without the ca Excel sets on some tables and not others, and its value."""
    kind = re.search(r'\st="([^"]*)"', xml.split('>')[0])
    formula = re.search(r'<f\b[^>]*?(?:/>|>.*?</f>)', xml)
    value = re.search(r'<v>(.*?)</v>', xml)
    return (kind.group(1) if kind else '', re.sub(r' ca="1"', '', formula.group(0)) if formula else '',
            value.group(1) if value else '')


@pytest.mark.parametrize('name', list(RECORD['files']))
def test_the_cells_a_file_keeps(tmp_path: Path, name: str) -> None:
    cells, calc = _saved(tmp_path, name)
    expected = RECORD['files'][name]
    assert set(cells) == set(expected['cells'])
    for reference, xml in expected['cells'].items():
        kind, formula, value = _parts(xml)
        found_kind, found_formula, found_value = _parts(cells[reference])
        assert (found_kind, found_formula) == (kind, formula), (reference, cells[reference], xml)
        if kind == 'e':
            assert found_value == value, (reference, cells[reference], xml)
        else:
            # Excel spells some numbers with a digit more than the shortest that reads back (see excel_number).
            assert float(found_value) == float(value), (reference, cells[reference], xml)
    assert ('calcMode="autoNoTable"' in calc) == ('calcMode="autoNoTable"' in expected['calcPr'])


#: Edits the model does not carry through a data table, each reporting itself unsupported rather than changing the
#: table's cells as cells of their own.
UNSUPPORTED = {
    'protected': 'ws.Protect: ws.Range("R2:S4").Table ColumnInput:=ws.Range("A1")',
    'fill_into': 'ws.Range("D6:D7").FillDown',
    'fill_from': 'ws.Range("D7").FillDown',
    'sort': 'ws.Range("C3:E6").Sort Key1:=ws.Range("D3")',
    'cut': 'ws.Range("C3").Cut ws.Range("R10")',
    'cut_onto': 'ws.Range("R10").Cut ws.Range("D4")',
    'merge': 'ws.Range("C2:D3").Merge',
    'replace': 'ws.Range("D3:E6").Replace "A1", "A2", xlPart',
    'copy_onto': 'ws.Range("A1").Copy ws.Range("D4")',
    'paste_onto': 'ws.Range("A1").Copy: ws.Range("D4").PasteSpecial xlPasteValues',
    'spilling_reader': 'ws.Range("R1").Formula2 = "=SEQUENCE(2)*A1": ws.Range("R5").Value = ws.Range("D3").Value',
}


@pytest.mark.parametrize('name', list(UNSUPPORTED))
def test_an_edit_the_model_does_not_carry_through_a_table(name: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(BUILD + f'\nSub Edit()\nDim ws As Worksheet\nSet ws = ActiveWorkbook.Worksheets(1)\nBuild ws\n'
                   f'ws.Range("A3").Value = 2000\n{UNSUPPORTED[name]}\nEnd Sub\n', name='Harness')
    with pytest.raises(VBAUnsupportedError, match=r'data table|Range\.Table'):
        app.run('Harness.Edit')


def _reader(app: ExcelApplication) -> str:
    app.add_module('Function Probe() As String\nDim ws As Worksheet\nSet ws = ActiveWorkbook.Worksheets(1)\n'
                   'ws.Range("A3").Value = 2000\n'
                   'Probe = ws.Range("D3").Formula & "|" & ws.Range("N3").FormulaR1C1 & "|" & ws.Range("D3").Value & '
                   '"|" & ws.Range("P5").Value & "|" & ws.Range("H3").CurrentArray.Address & "|" & '
                   'Application.Calculation\nEnd Function\n', name='Reading')
    return str(app.run('Reading.Probe'))


def test_a_saved_table_opens_and_works_again(tmp_path: Path) -> None:
    _saved(tmp_path, 'tables')
    app = ExcelApplication.open(tmp_path / 'tables.xlsx')
    model_changes = PROBES['model_changes']['result'].split('|')
    assert _reader(app) == '|'.join(['=TABLE(,A1)', '=TABLE(R2C1,R1C1)', model_changes[0], model_changes[4],
                                     '$G$2:$K$3', '-4105'])


def test_excel_finds_what_the_model_saved(tmp_path: Path) -> None:
    # Excel opened the model's three files without a repair, beside a workbook already open; the model opening them
    # the same way finds every cell's formula and value as Excel did, and after the model changes.
    for name in RECORD['files']:
        _saved(tmp_path, name)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['opened']['macro'].replace('{folder}', f'{tmp_path}{os.sep}'), name='Opening')
    assert app.run('Opening.Probe') == RECORD['opened']['result']


def test_a_semiautomatic_workbook_opens_waiting(tmp_path: Path) -> None:
    _saved(tmp_path, 'semiautomatic')
    app = ExcelApplication.open(tmp_path / 'semiautomatic.xlsx')
    values = PROBES['values']['result'].split('#')
    # Opened alone it brings its calculation, and the tables keep what the file holds when the model changes.
    assert _reader(app) == '|'.join(['=TABLE(,A1)', '=TABLE(R2C1,R1C1)', values[0].split('|')[0],
                                     values[2].split('|')[8], '$G$2:$K$3', '2'])
