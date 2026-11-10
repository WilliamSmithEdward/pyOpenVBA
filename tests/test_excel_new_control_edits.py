"""A setting changed on a form control a macro has just made, replayed from live Excel (scripts/measure_new_control_edits.py,
tests/fixtures/new_control_edits.json): what the set answers, what the control reads in the session and after a reopen, and what the save
says of it. `excel_reads_model` is Excel's reading of the model's save, recorded when the fixture was measured.

The model writes a new control's parts in its own layout, so a save is compared by what it says: each VML client setting, the div's
text-align and the textbox's layout flow, the control part's attributes, and the DrawingML text frame. What the set changed from a save
with nothing set is the same in the model's files as in Excel's."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'new_control_edits.json').read_text(encoding='utf-8'))
#: The settings the model does not yet make on a new control as Excel does, each with why.
NOT_YET: dict[tuple[str, str], str] = {
    ('Buttons', 'Formula'): 'refused: the caption it shows is the cell text, which a new control does not take yet',
    ('Buttons', 'AutoSize'): 'refused: the size it fits to is the caption measured in its font',
}
#: Settings a new button takes in the session but does not save yet: Excel gives the button a DrawingML text body and reads these from
#: it in place of the VML, and the model writes none, so Excel would read the model's file otherwise. The save reports itself.
UNSAVED = {('Buttons', member, 'True') for member in ('DefaultButton', 'CancelButton', 'DismissButton', 'HelpButton')} | {
    ('Buttons', 'HorizontalAlignment', 'xlLeft'), ('Buttons', 'HorizontalAlignment', 'xlDistributed'), ('Buttons', 'VerticalAlignment', 'xlTop'),
    ('Buttons', 'VerticalAlignment', 'xlBottom'), ('Buttons', 'Orientation', 'xlUpward')}
CASES = [one for one in RECORD['cases'] if (one['kind'], one['member']) not in NOT_YET and (one['kind'], one['member'], one['value']) not in UNSAVED]
HELD = [one for one in RECORD['cases'] if (one['kind'], one['member'], one['value']) in UNSAVED]
CONSTANTS = {'xlRTL': -5004, 'xlContext': -5002, 'xlLeft': -4131, 'xlCenter': -4108, 'xlDistributed': -4117, 'xlTop': -4160, 'xlBottom': -4107,
             'xlUpward': -4171, 'xlHorizontal': -4128}


def _value(text: str) -> object:
    if text in ('True', 'False'):
        return text == 'True'
    if text.startswith('"'):
        return text.strip('"')
    return CONSTANTS[text]


def facts(parts: dict[str, str]) -> set[str]:
    """What a save says of its one control's settings."""
    found: set[str] = set()
    for name, text in parts.items():
        if name.endswith('.vml'):
            client = re.search(r'<x:ClientData\b[^>]*>(.*?)</x:ClientData>', text, re.DOTALL)
            for tag, inner in re.findall(r'<x:(\w+)\b[^>]*?(?:/>|>(.*?)</x:\1>)', client.group(1) if client else '', re.DOTALL):
                if tag != 'Anchor':
                    found.add(f'vml {tag}={inner.strip()}')
            division = re.search(r"<div\b[^>]*>", text)
            align = re.search(r'text-align:(\w+)', division.group(0)) if division else None
            found.add(f'vml div={align.group(1) if align else ""}')
            box = re.search(r'<v:textbox\b[^>]*>', text)
            flow = re.search(r'mso-layout-flow-alt:([\w-]+)|layout-flow:(\w+)', box.group(0)) if box else None
            found.add(f'vml flow={flow.group(0) if flow else ""}')
        elif 'ctrlProps/' in name:
            head = re.search(r'<formControlPr\b([^>]*)', text)
            found |= {f'part {key}={value}' for key, value in re.findall(r'(\w+)="([^"]*)"', head.group(1) if head else '')
                      if key != 'xmlns' and value}
        elif '/drawings/drawing' in name:
            body = re.search(r'<a:bodyPr\b[^>]*>', text)
            for key in ('anchor', 'vert'):
                value = re.search(rf'\b{key}="(\w+)"', body.group(0)) if body else None
                found.add(f'dml {key}={value.group(1) if value else ""}')
            paragraph = re.search(r'<a:pPr\b[^>]*\balgn="(\w+)"', text)
            found.add(f'dml algn={paragraph.group(1) if paragraph else ""}')
    return found


def _change(before: dict[str, str], after: dict[str, str]) -> tuple[set[str], set[str]]:
    one, two = facts(before), facts(after)
    return two - one, one - two


def _saved(case: dict[str, Any], folder: Path, change: bool) -> tuple[str, dict[str, str]]:
    saved = folder / f'out_{change}.xlsx'
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Probe')
    answer = str(app.run('Probe.Row', str(saved), case['kind'], RECORD['boxes'][case['kind']], case['member'], _value(case['value']), change))
    with zipfile.ZipFile(saved) as package:
        return answer, {name: package.read(name).decode('utf-8') for name in package.namelist()
                        if re.search(r'ctrlProp|vmlDrawing|drawings/drawing', name)}


@pytest.mark.parametrize('case', CASES, ids=[f'{one["kind"]} {one["member"]} {one["value"]}' for one in CASES])
def test_each_setting_answers_and_saves_as_excel_did(case: dict[str, Any], tmp_path: Path) -> None:
    answer, parts = _saved(case, tmp_path, True)
    assert answer == case['excel']
    _, untouched = _saved(case, tmp_path, False)
    # A part Excel's set left alone is not in the fixture; what it says did not change.
    wanted = _change({name: texts['before'] for name, texts in case['parts'].items()}, {name: texts['after'] for name, texts in case['parts'].items()})
    assert _change(untouched, parts) == wanted


@pytest.mark.parametrize('case', [one for one in CASES if 'excel_reads_model' in one], ids=lambda one: f'{one["kind"]} {one["member"]} {one["value"]}')
def test_excel_read_the_model_save_as_its_own(case: dict[str, Any]) -> None:
    assert case['excel_reads_model'] == case['excel'].split('|reopened ')[1]


#: What a new button reads in the session, before any save, for each setting it takes but does not save.
SESSION = r'''Public Function Session(ByVal member As String, ByVal value As Variant) As String
    Dim o As Object, v As Variant
    Set o = Workbooks.Add(xlWBATWorksheet).Worksheets(1).Buttons.Add(100, 70, 60, 30)
    CallByName o, member, VbLet, value
    v = CallByName(o, member, VbGet)
    Session = "E0|" & TypeName(v) & " " & CStr(v)
End Function
'''


@pytest.mark.parametrize('case', HELD, ids=lambda one: f'{one["member"]} {one["value"]}')
def test_a_setting_a_new_button_cannot_save_reads_as_excel_did_then_its_save_reports_itself(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SESSION, name='Session')
    assert str(app.run('Session.Session', case['member'], _value(case['value']))) == case['excel'].split('|reopened ')[0]
    with pytest.raises(VBAUnsupportedError, match="Saving a new button's"):
        _saved(case, tmp_path, True)


def test_not_yet_names_measured_settings_only() -> None:
    measured = {(one['kind'], one['member'], one['value']) for one in RECORD['cases']}
    assert set(NOT_YET) <= {(kind, member) for kind, member, _ in measured}
    assert UNSAVED <= measured
