"""A setting changed on a form control loaded from a file, replayed from live Excel (scripts/measure_loaded_control_edits.py,
tests/fixtures/loaded_control_edits.json): what the set and a reopen answer, what the save writes, and what Excel read of the model's
save. `excel_reads_model` is Excel's reading of the model's save, recorded when the fixture was measured.

The sheet's part and the control's own part are compared whole. Excel writes its VML again on every save and renumbers every shape's
creationId in the drawing part, so in those two what the set changed is compared: the model's save against its own untouched save,
element by element, and Excel's after against its before. A part Excel's set left alone is the model's untouched save byte for byte."""
from __future__ import annotations

import difflib
import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_control_edits.json').read_text(encoding='utf-8'))
#: scripts/measure_button_text.py: every value of a button's alignments, Orientation and ReadingOrder, and PhoneticAccelerator, with what
#: each reads in the session.
TEXT: dict[str, Any] = json.loads((FIXTURES / 'button_text.json').read_text(encoding='utf-8'))
#: The settings the model does not yet set on a loaded control as Excel does.
NOT_YET: dict[tuple[str, str], str] = {('Buttons', 'AutoSize'): 'refused: the size it fits to is the caption measured in its font'}
CASES = ([{**one, 'source': RECORD['source']} for one in RECORD['cases'] if (one['kind'], one['member']) not in NOT_YET]
         + [{**one, 'source': TEXT['source']} for one in TEXT['cases']])
#: The parts a control's settings live in.
CONTROL_PARTS = re.compile(r'xl/(?:worksheets/sheet|drawings/drawing|drawings/vmlDrawing|ctrlProps/ctrlProp)\d+\.(?:xml|vml)$')
CONSTANTS = {'xlLeft': -4131, 'xlCenter': -4108, 'xlRight': -4152, 'xlJustify': -4130, 'xlDistributed': -4117, 'xlTop': -4160,
             'xlBottom': -4107, 'xlHorizontal': -4128, 'xlUpward': -4171, 'xlDownward': -4170, 'xlVertical': -4166, 'xlRTL': -5004,
             'xlLTR': -5003, 'xlContext': -5002, 'xlFreeFloating': 3, 'xlSimple': 1}


def _value(text: str) -> object:
    if text in ('True', 'False'):
        return text == 'True'
    if text.startswith('"'):
        return text.strip('"')
    return CONSTANTS.get(text, int(text) if text.lstrip('-').isdigit() else text)


def _saved(case: dict[str, Any], folder: Path, change: bool = True) -> tuple[str, dict[str, str]]:
    book, saved = folder / 'in.xlsx', folder / f'out_{change}.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(case.get('source', RECORD['source']), name='Probe')
    answer = str(app.run('Probe.Row', str(book), str(saved), case['kind'], case['member'], _value(case['value']), change))
    with zipfile.ZipFile(saved) as package:
        return answer, {name: package.read(name).decode('utf-8') for name in package.namelist() if CONTROL_PARTS.match(name)}


def _changes(before: str, after: str) -> list[tuple[str, list[str], list[str]]]:
    """What changed from one part to the other, tag by tag and text by text."""
    one, two = (re.split(r'(?<=>)|(?=<)', text) for text in (before, after))
    return [(tag, one[i1:i2], two[j1:j2]) for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, one, two, autojunk=False).get_opcodes()
            if tag != 'equal']


@pytest.mark.parametrize('case', CASES, ids=[f'{one["kind"]} {one["member"]} {one["value"]}' for one in CASES])
def test_each_setting_answers_and_saves_as_excel_did(case: dict[str, Any], tmp_path: Path) -> None:
    answer, parts = _saved(case, tmp_path)
    assert answer == case['excel']
    _, untouched = _saved(case, tmp_path, change=False)
    for name, part in parts.items():
        texts = case['parts'].get(name)
        if texts is None:
            assert part == untouched[name], name
        elif name.endswith('.vml') or '/drawings/drawing' in name:
            assert _changes(untouched[name], part) == _changes(texts['before'], texts['after']), name
        else:
            assert part == texts['after'], name


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


FLAGS: dict[str, Any] = json.loads((FIXTURES / 'button_flags.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', FLAGS['cases'], ids=lambda one: f'{"new" if one["new"] else "loaded"} {one["member"]}')
def test_a_button_flag_reads_in_the_session_as_excel_did(case: dict[str, Any], tmp_path: Path) -> None:
    """scripts/measure_button_flags.py: the roles read True for the session and False after a reopen, AddIndent False throughout. A new
    button's roles are not saved yet (tests/test_excel_new_control_edits.py), and the save reports itself."""
    book, saved = tmp_path / 'in.xlsx', tmp_path / 'out.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(FLAGS['source'], name='Probe')
    if case['new'] and case['member'] != 'AddIndent':
        with pytest.raises(VBAUnsupportedError, match="Saving a new button's"):
            app.run('Probe.Row', str(book), str(saved), case['member'], True)
        return
    assert str(app.run('Probe.Row', str(book), str(saved), case['member'], case['new'])) == case['excel']


def _loaded(tmp_path: Path, body: str) -> ExcelApplication:
    book = tmp_path / 'in.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication.open(book, with_vba=False)
    app.add_module(f'Public Sub Go()\n    With ActiveWorkbook.Worksheets(1)\n{body}    End With\nEnd Sub\n', name='Go')
    return app


def test_a_second_text_setting_on_one_button_is_refused(tmp_path: Path) -> None:
    """Only one of a button's alignments and Orientation at a time was measured from its layout, so a second is refused; the first can
    be set again."""
    app = _loaded(tmp_path, '        .Buttons(1).HorizontalAlignment = xlLeft\n        .Buttons(1).HorizontalAlignment = xlRight\n'
                            '        .Buttons(1).VerticalAlignment = xlTop\n')
    with pytest.raises(VBAUnsupportedError, match='laid out otherwise'):
        app.run('Go.Go')


def test_saving_a_new_button_text_setting_is_refused(tmp_path: Path) -> None:
    app = _loaded(tmp_path, '        .Buttons.Add(300, 300, 60, 30).Orientation = xlUpward\n')
    app.run('Go.Go')
    with pytest.raises(VBAUnsupportedError, match="Saving a new button's"):
        app.save(tmp_path / 'out.xlsx')


@pytest.mark.parametrize('member', ['HorizontalAlignment = xlDistributed', 'VerticalAlignment = xlBottom', 'Orientation = xlUpward'])
def test_saving_a_button_text_setting_twice_writes_it_once(member: str, tmp_path: Path) -> None:
    app = _loaded(tmp_path, f'        .Buttons(1).{member}\n')
    app.run('Go.Go')
    first, second = tmp_path / 'first.xlsx', tmp_path / 'second.xlsx'
    app.save(first)
    app.save(second)
    with zipfile.ZipFile(first) as one, zipfile.ZipFile(second) as two:
        for name in one.namelist():
            if CONTROL_PARTS.match(name):
                assert one.read(name) == two.read(name), name
