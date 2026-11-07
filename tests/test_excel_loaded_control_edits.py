"""A setting changed on a form control loaded from a file, replayed from live Excel (scripts/measure_loaded_control_edits.py,
tests/fixtures/loaded_control_edits.json): what the set and a reopen answer, what the save writes, and what Excel read of the model's
save. `excel_reads_model` is Excel's reading of the model's save, recorded when the fixture was measured.

Excel writes the whole VML part again on every save, so the model's VML is compared with Excel's in the changed control's own
x:ClientData; the sheet's part and the control's own part are compared whole."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_control_edits.json').read_text(encoding='utf-8'))
#: The settings the model writes on a loaded control, so far.
MODELLED = ('Enabled', 'PrintObject', 'LockedText', 'DropDownLines')
CASES = [one for one in RECORD['cases'] if one['member'] in MODELLED]
CONSTANTS = {'xlRTL': -5004, 'xlLeft': -4131, 'xlTop': -4160, 'xlUpward': -4171, 'xlFreeFloating': 3, 'xlSimple': 1}


def _value(text: str) -> object:
    if text in ('True', 'False'):
        return text == 'True'
    if text.startswith('"'):
        return text.strip('"')
    return CONSTANTS.get(text, int(text) if text.lstrip('-').isdigit() else text)


def _saved(case: dict[str, Any], folder: Path) -> tuple[str, dict[str, str]]:
    book, saved = folder / 'in.xlsx', folder / 'out.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Probe')
    answer = str(app.run('Probe.Row', str(book), str(saved), case['kind'], case['member'], _value(case['value']), True))
    with zipfile.ZipFile(saved) as package:
        return answer, {name: package.read(name).decode('utf-8') for name in package.namelist() if name in case['parts']}


def _blocks(vml: str) -> list[str]:
    return re.findall(r'<x:ClientData\b.*?</x:ClientData>', vml, re.DOTALL)


@pytest.mark.parametrize('case', CASES, ids=[f'{one["kind"]} {one["member"]}' for one in CASES])
def test_each_setting_answers_and_saves_as_excel_did(case: dict[str, Any], tmp_path: Path) -> None:
    answer, parts = _saved(case, tmp_path)
    assert answer == case['excel']
    for name, texts in case['parts'].items():
        if name.endswith('.vml'):
            before, after = _blocks(texts['before']), _blocks(texts['after'])
            changed = next(at for at, (one, two) in enumerate(zip(before, after)) if one != two)
            assert _blocks(parts[name])[changed] == after[changed]
        else:
            assert parts[name] == texts['after'], name


@pytest.mark.parametrize('case', [one for one in CASES if 'excel_reads_model' in one], ids=lambda one: f'{one["kind"]} {one["member"]}')
def test_excel_read_the_model_save_as_its_own(case: dict[str, Any]) -> None:
    assert case['excel_reads_model'] == case['excel'].split('|reopened ')[1]


def test_a_loaded_control_nobody_changed_is_saved_as_its_file_has_it(tmp_path: Path) -> None:
    """Setting what a control already holds writes nothing new: no part of it differs from the workbook saved untouched."""
    book = tmp_path / 'in.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    untouched, again = tmp_path / 'untouched.xlsx', tmp_path / 'again.xlsx'
    first = ExcelApplication.open(book, with_vba=False)
    first.save(untouched)
    second = ExcelApplication.open(book, with_vba=False)
    second.add_module('Public Sub Same()\n    With ActiveWorkbook.Worksheets(1)\n        .Buttons(1).Enabled = .Buttons(1).Enabled\n'
                      '        .CheckBoxes(1).PrintObject = .CheckBoxes(1).PrintObject\n        .DropDowns(1).DropDownLines = .DropDowns(1).DropDownLines\n'
                      '        .OptionButtons(1).LockedText = .OptionButtons(1).LockedText\n    End With\nEnd Sub\n', name='Same')
    second.run('Same.Same')
    second.save(again)
    with zipfile.ZipFile(untouched) as one, zipfile.ZipFile(again) as two:
        for name in one.namelist():
            if re.search(r'ctrlProp|vmlDrawing|worksheets/sheet', name):
                assert one.read(name) == two.read(name), name
