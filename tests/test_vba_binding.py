"""How a call binds, against what Excel's compiler refused (scripts/measure_vba_binding.py)."""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.interpreter import _ast as A
from pyopenvba.interpreter._compile import nodes
from pyopenvba.interpreter._type_library import build, find_class

ANSWERS: dict[str, dict[str, object]] = json.loads((Path(__file__).parent / 'fixtures/vba_binding.json').read_text(encoding='utf-8'))['answers']
TOP = '''Public HeldRange As Range
Public HeldObject As Object
Public HeldVariant As Variant
Public Type Rec
    R As Range
    O As Object
End Type

Function FRange() As Range
End Function
Function FObject() As Object
End Function
Function FVariant() As Variant
End Function
Function FNone()
End Function
Property Get PRange() As Range
End Property
Property Get PObject() As Object
End Property
'''
BODY = '''Function Probe(Optional pr As Range, Optional po As Object, Optional pv As Variant) As String
    Dim ws As Worksheet, wb As Workbook, r As Range, co As ChartObject, ch As Chart, sh As Shape, lo As ListObject, o As Object, v As Variant
    Dim arr(1) As Range, arrv(1) As Variant, rec As Rec, coll As Collection, cobj As Object
    If False Then
        @@LINE@@
    End If
    Probe = "compiled"
End Function
'''
FORMS = {
    'with_range': 'With r\n.@\nEnd With', 'with_object': 'With o\n.@\nEnd With', 'with_expression': 'With ws.Range("A1")\n.@\nEnd With', 'with_active_sheet': 'With ActiveSheet\n.@\nEnd With', 'with_sheet_range': 'With ActiveSheet\n.Range("A1").@\nEnd With', 'with_nested': 'With r\nWith .Font\n.@\nEnd With\nEnd With', 'with_variant': 'With v\n.@\nEnd With',
    'for_each_range': 'Dim c As Range\nFor Each c In r\nc.@\nNext', 'for_each_variant': 'Dim k As Variant\nFor Each k In r\nk.@\nNext', 'for_each_object': 'Dim k As Object\nFor Each k In r\nk.@\nNext',
    'set_keeps_declared': 'Set o = r\no.@', 'typed_from_late': 'Set r = ActiveSheet.Range("A1")\nr.@',
    'param_range': 'pr.@', 'param_object': 'po.@', 'param_variant': 'pv.@',
}
#: Collection is early but has no Creator and takes any name, so Excel's compiler shows nothing either way.
TOLD_APART = sorted({key.split(' ', 1)[1] for key in ANSWERS if key.startswith('creator ')} - {'coll'})


def early_in_excel(expression: str) -> bool:
    return 'compile' in ANSWERS['nope ' + expression] or 'compile' in ANSWERS['creator ' + expression]


def early_here(expression: str) -> bool:
    """Whether the probe's own member call was resolved through a library class."""
    line = FORMS[expression].replace('@', 'Nope') if expression in FORMS else f'v = {expression}.Nope'
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(TOP + BODY.replace('@@LINE@@', line), name='Harness')
    assert app.run('Harness.Probe') == 'compiled'
    assert app.interpreter is not None
    module = app.interpreter.modules['harness']
    procedure = module.procedure('Probe')
    assert procedure is not None
    program = module.programs[id(procedure)]
    probes = [node for node in nodes(procedure.body) if isinstance(node, A.Member) and node.name == 'Nope']
    assert len(probes) == 1
    return id(probes[0]) in program.early_bound_members


def test_the_expressions_asked_hold_both_bindings() -> None:
    early = [expression for expression in TOLD_APART if early_in_excel(expression)]
    assert 100 < len(early) < len(TOLD_APART) - 50


@pytest.mark.parametrize('expression', TOLD_APART)
def test_a_call_binds_as_excels_compiler_read_it(expression: str) -> None:
    assert early_here(expression) == early_in_excel(expression)


def test_the_table_says_what_it_was_read_from() -> None:
    office, versions = build()
    assert re.fullmatch(r'\d+\.\d+\.\d+\.\d+', office)
    assert set(versions) == {'excel', 'office'}


@pytest.mark.parametrize(('expression', 'nonextensible'), [('Worksheet', True), ('Chart', True), ('Sheets', True), ('Workbooks', True), ('FillFormat', True), ('Range', False), ('Workbook', False), ('Font', False), ('Shape', False)])
def test_the_compiler_takes_any_name_on_an_extensible_class(expression: str, nonextensible: bool) -> None:
    found = find_class('excel', expression)
    assert found is not None
    assert found.nonextensible == nonextensible


@pytest.mark.parametrize(('declared', 'message'), [('ws', "Can't assign to read-only property"), ('r', 'Wrong number of arguments or invalid property assignment')])
def test_a_dual_class_says_read_only_where_a_dispatch_class_says_wrong_number(declared: str, message: str) -> None:
    assert ANSWERS[f'creator {declared}'] == {'compile': message[:70]}
    found = find_class('excel', {'ws': 'Worksheet', 'r': 'Range'}[declared])
    assert found is not None and found.dual == (declared == 'ws')


def test_the_type_of_a_project_class_is_not_the_librarys() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Value As Long', name='Range', kind='class')
    app.add_module('Function Probe() As Long\n    Dim r As Range\n    Set r = New Range\n    r.Value = 4\n    Probe = r.Value\nEnd Function', name='Harness')
    assert app.run('Harness.Probe') == 4


#: The line runs with ws the first sheet, which has no comments: what Comments(Null) answers says how the call was bound.
SETUP = 'Dim ws As Worksheet, wo As Object, cs As Comments, co As Object, n As Long\nSet ws = ActiveSheet: Set wo = ws: Set cs = ws.Comments: Set co = cs\nOn Error Resume Next\n'


@pytest.mark.parametrize(('line', 'expected'), [
    ('n = TypeName(ws.Comments(Null))', 94),
    ('n = TypeName(wo.Comments(Null))', 13),
    ('n = TypeName(cs(Null))', 94),
    ('n = TypeName(co(Null))', 13),
    ('n = TypeName(cs.Item(Null))', 94),
    ('n = TypeName(co.Item(Null))', 13),
    ('With cs\nn = TypeName(.Item(Null))\nEnd With', 94),
    ('With co\nn = TypeName(.Item(Null))\nEnd With', 13),
    ('With ActiveSheet\nn = TypeName(.Comments(Null))\nEnd With', 13),
    ('n = TypeName(ActiveCell.Worksheet.Comments(Null))', 94),
    ('n = TypeName(Worksheets(1).Comments(Null))', 13),
])
def test_a_call_answers_by_how_it_was_bound(line: str, expected: int) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Function Probe() As Long\n{SETUP}{line}\nProbe = Err.Number\nEnd Function', name='Harness')
    assert app.run('Harness.Probe') == expected


def test_a_parameter_binds_as_it_is_declared() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('''Function Typed(c As Comments) As Long
    Dim v As Variant
    On Error Resume Next
    v = TypeName(c.Item(Null))
    Typed = Err.Number
End Function
Function Loose(c As Object) As Long
    Dim v As Variant
    On Error Resume Next
    v = TypeName(c.Item(Null))
    Loose = Err.Number
End Function
Function Probe() As String
    Probe = Typed(ActiveSheet.Comments) & " " & Loose(ActiveSheet.Comments)
End Function''', name='Harness')
    assert app.run('Harness.Probe') == '94 13'
