"""The sheet's old collections and their items on a workbook Excel saved, each property read typed and late, replayed from live Excel
(scripts/measure_old_collection_members.py, tests/fixtures/old_collection_members.json)."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError, VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
BOOK = FIXTURES / 'old_collections.xlsx'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'old_collection_members.json').read_text(encoding='utf-8'))
CASES: list[dict[str, Any]] = RECORD['cases']
#: These report themselves: the text layout of a drawing from a file, and a form control's Border and Interior.
REFUSED = re.compile(r'o\.(AutoScaleFont|HorizontalAlignment|VerticalAlignment|Orientation|Font\b|Characters\b|Border|Interior)')
DRIVER = '''
Public Function InModel(ByVal routine As String, ByVal which As Long) As String
    InModel = Application.Run(routine, ActiveWorkbook.Worksheets(1), which)
End Function

Public Sub Trimmed(ByVal routine As String)
    Application.Run routine, ActiveWorkbook.Worksheets(1), 0
End Sub
'''


def compiled() -> str:
    source: str = RECORD['source']
    for routine in {one['routine'] for one in CASES if one['excel'].startswith('compile')}:
        source = re.sub(rf'Public Function {routine}\(.*?End Function\n\n', '', source, flags=re.DOTALL)
    return source + DRIVER


SOURCE = compiled()
GROUPS: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
for _case in CASES:
    if not _case['excel'].startswith('compile'):
        GROUPS.setdefault((_case['state'], _case['class'], _case['binding']), []).append(_case)


def answered(app: ExcelApplication, case: dict[str, Any]) -> str:
    try:
        return str(app.run('Probe.InModel', case['routine'], case['which']))
    except VBAUnsupportedError:
        return 'reports itself'


def allowed(case: dict[str, Any], got: str) -> bool:
    if got == 'reports itself':
        return REFUSED.search(case['read']) is not None
    return False


@pytest.mark.parametrize('key', GROUPS, ids=' '.join)
def test_each_read_as_excel_answered(key: tuple[str, str, str]) -> None:
    group = GROUPS[key]
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module(SOURCE, name='Probe')
    if group[0].get('trim'):
        app.run('Probe.Trimmed', group[0]['trim'])
    wrong = [(one['read'], one['excel'], got) for one in group
             if (got := answered(app, one)) != one['excel'] and not allowed(one, got)]
    assert not wrong


@pytest.mark.xfail(reason='The compile step does not yet check a typed default member against the type library')
@pytest.mark.parametrize('name', ['CheckBoxes', 'DropDowns', 'ListBoxes', 'OptionButtons'])
def test_a_typed_index_on_a_collection_of_valued_controls_does_not_compile(name: str) -> None:
    """Their default member is Value, which takes no index. `o.Item(1)` compiles, and late `o(1)` runs."""
    app = ExcelApplication.open(BOOK, with_vba=False)
    source = f'Public Function Tried(ws As Worksheet) As String\n    Dim o As Excel.{name}\n    Set o = ws.{name}\n    Tried = o(1).Name\nEnd Function\n'
    with pytest.raises(VBACompileError, match='Wrong number of arguments'):
        app.add_module(source, name='Probe')


def _vml(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read('xl/drawings/vmlDrawing1.vml').decode('utf-8')


def _shapes(vml: str) -> list[str]:
    return re.findall(r'<v:shape id="_x0000_s\d+".*?</v:shape>', vml, re.DOTALL)


def test_reading_a_control_settings_leaves_its_markup_alone_at_a_save(tmp_path: Path) -> None:
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module('Public Sub Edit()\n    ActiveWorkbook.Worksheets(1).Range("H30").Value = 1\nEnd Sub\n', name='Probe')
    app.run('Probe.Edit')
    app.interpreter.sandbox = False
    app.save(tmp_path / 'saved.xlsx')
    assert _vml(tmp_path / 'saved.xlsx') == _vml(BOOK)


def test_a_new_control_leaves_the_loaded_controls_markup_alone(tmp_path: Path) -> None:
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module('Public Sub Edit()\n    ActiveWorkbook.Worksheets(1).Buttons.Add 10, 500, 60, 30\nEnd Sub\n', name='Probe')
    app.run('Probe.Edit')
    app.save(tmp_path / 'saved.xlsx')
    assert _shapes(_vml(tmp_path / 'saved.xlsx'))[:10] == _shapes(_vml(BOOK))
