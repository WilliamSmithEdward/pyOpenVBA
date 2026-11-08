"""The second probe of an object given by Let (no Set), on classes and properties the first did not use, with the receiving rule's predictions
and the model's written down first (tests/fixtures/late_object_lets.json was the first).

Each case makes a new workbook with 7 in A1, "abc" in B1, 9 in C1, the date 1/2/2020 in A2 and a name Base, and gives one object to one
property, typed and late. The objects: a one-cell Range holding a number, one holding a date, one area of two cells, an Interior, a
Worksheet, a Name, a Collection and Nothing. Widths and heights are left out: they depend on the display's DPI.

The receiving rule's prediction for each case, from the first probe, by the kind of property:
  a one-cell Range gives its value, so the set answers as that value would;
  a String property refuses another object or Nothing with 13;
  a Boolean property refuses another object with 1004, and Nothing with 5;
  a Variant property refuses another object with 1004, and Nothing with 5;
  a Long enum is either: Workbook.UpdateLinks said 13 and Worksheet.Visible 1004, so this probe decides between them;
  typed, a String, Boolean or Long property gets the object's default value, and one with none is 438; a Collection or Nothing does not
  compile.

    python scripts/measure_late_object_lets_2.py predict
    python scripts/measure_late_object_lets_2.py measure

writes tests/fixtures/late_object_lets_2.json.
"""
from __future__ import annotations

import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/late_object_lets_2.json'
SEPARATOR = '%%'
#: Each target: the declared class, how the probe reaches it, the property, and its kind as the type library declares it.
TARGETS = [('Worksheet', 'ws', 'ScrollArea', 'String'), ('Worksheet', 'ws', 'EnableSelection', 'enum'),
           ('Worksheet', 'ws', 'DisplayPageBreaks', 'Boolean'), ('Workbook', 'wb', 'CheckCompatibility', 'Boolean'),
           ('Workbook', 'wb', 'ConflictResolution', 'enum'), ('Range', 'ws.Range("D1")', 'Formula', 'Variant'),
           ('Range', 'ws.Range("D1")', 'HorizontalAlignment', 'Variant'), ('Range', 'ws.Range("D1")', 'IndentLevel', 'Variant'),
           ('Range', 'ws.Range("D1")', 'WrapText', 'Variant'), ('Range', 'ws.Range("D1")', 'Locked', 'Variant'),
           ('Interior', 'ws.Range("D1").Interior', 'ColorIndex', 'Variant'), ('Interior', 'ws.Range("D1").Interior', 'Pattern', 'Variant'),
           ('Font', 'ws.Range("D1").Font', 'Italic', 'Variant'), ('Font', 'ws.Range("D1").Font', 'Color', 'Variant'),
           ('Name', 'wb.Names("Base")', 'RefersTo', 'Variant'), ('Name', 'wb.Names("Base")', 'Visible', 'Boolean')]
OBJECTS = ['ws.Range("A1")', 'ws.Range("A2")', 'ws.Range("A1:B1")', 'ws.Range("A1").Interior', 'ws', 'wb.Names("Base")', 'New Collection',
           'Nothing']
#: Typed, a Let of these does not compile (tests/fixtures/late_object_lets.json), so a typed probe leaves them out.
NOT_COMPILING = frozenset({'New Collection', 'Nothing'})
#: Objects with no default member a Let can read.
NO_VALUE = frozenset({'ws.Range("A1").Interior', 'ws', 'New Collection'})
PART = 16


def rule(kind: str, value: str, typed: bool) -> str:
    """The receiving rule's predicted error for a case, before Excel answers: a number, 'value' where the object's value decides, or a
    choice where the first probe left it open."""
    if typed and value in NOT_COMPILING:
        return 'does not compile'
    if value.startswith('ws.Range') and not value.endswith('Interior') and value != 'ws.Range("A1:B1")':
        return 'value'
    if typed and kind != 'Variant':
        return '438' if value in NO_VALUE else 'value'
    if value == 'Nothing':
        return {'String': '13', 'enum': '13 or 5'}.get(kind, '5')
    if value == 'ws.Range("A1:B1")':
        return 'open'
    return {'String': '13', 'enum': '13 or 1004'}.get(kind, '1004')


def _case(target: tuple[str, str, str, str], value: str, typed: bool, compiles_only: bool) -> str:
    kind_class, path, prop, _kind = target
    holder = 't' if typed else 'o'
    if compiles_only and typed and value in NOT_COMPILING:
        return f'    out = out & "{SEPARATOR}" & "does not compile"\n'
    return f'''    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = 7: ws.Range("B1").Value = "abc": ws.Range("C1").Value = 9: ws.Range("A2").Value = DateSerial(2020, 1, 2)
    wb.Names.Add Name:="Base", RefersTo:="=Sheet1!$A$1"
    Set {holder}{kind_class} = {path}
    Err.Clear
    {holder}{kind_class}.{prop} = {value}
    s = "e" & Err.Number
    Err.Clear: v = Empty
    v = {holder}{kind_class}.{prop}
    s = s & "/" & IIf(Err.Number, "e" & Err.Number, TypeName(v) & ":" & CStr(v))
    wb.Close SaveChanges:=False
    out = out & "{SEPARATOR}" & s
'''


def cases() -> list[tuple[str, str, str, str, str]]:
    return [(*target, value) for target in TARGETS for value in OBJECTS]


def macro(typed: bool, compiles_only: bool = False) -> str:
    every = cases()
    parts = [every[at:at + PART] for at in range(0, len(every), PART)]
    declare = ('    Dim tWorkbook As Workbook, tWorksheet As Worksheet, tRange As Range, tFont As Font, tInterior As Interior, tName As Name\n'
               if typed else '    Dim oWorkbook As Object, oWorksheet As Object, oRange As Object, oFont As Object, oInterior As Object, oName As Object\n')
    procedures = ''.join(f'''Private Sub Part{index}(ByRef out As String)
    Dim wb As Workbook, ws As Worksheet, v As Variant, s As String
{declare}    On Error Resume Next
{''.join(_case(case[:4], case[4], typed, compiles_only) for case in part)}End Sub
''' for index, part in enumerate(parts))
    calls = ''.join(f'    Part{index} out\n' for index in range(len(parts)))
    return f'''{procedures}Function Probe() As String
    Dim out As String
    Application.DisplayAlerts = False
{calls}    Application.DisplayAlerts = True
    Probe = Mid(out, {len(SEPARATOR) + 1})
End Function
'''


def predict() -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication

    model: dict[str, list[str]] = {}
    for typed in (True, False):
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(macro(typed), name='H')
        model['typed' if typed else 'late'] = str(app.run('H.Probe')).split(SEPARATOR)
    table = {('typed' if typed else 'late'): [rule(case[3], case[4], typed) for case in cases()] for typed in (True, False)}
    text = json.dumps({'cases': cases(), 'rule': table, 'model': model}, indent=1) + '\n'
    OUT.write_text(text, encoding='utf-8', newline='\n')
    print('predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def _verdict(predicted: str, answer: str) -> bool | None:
    """Whether the rule's prediction held: None where it left the case open or to the object's value."""
    if predicted in ('value', 'open'):
        return None
    if predicted == 'does not compile':
        return answer == predicted
    return answer.split('/')[0].removeprefix('e') in predicted.split(' or ')


def measure() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, Any] = json.loads(OUT.read_text(encoding='utf-8'))
    answers: dict[str, list[str]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.run_vba('Function Hello() As String\n    Hello = "x"\nEnd Function\n', 'Hello', (), timeout=60)
        for typed in (True, False):
            result = excel.run_vba(macro(typed, compiles_only=True), 'Probe', (), timeout=400, line_numbers=False)
            assert result.ok, (result.error, [one.message for one in result.dialogs])
            answers['typed' if typed else 'late'] = str(result.value).split(SEPARATOR)
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel': answers})
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    for binding in answers:
        verdicts = [_verdict(p, a) for p, a in zip(record['rule'][binding], answers[binding])]
        print(binding, 'rule held', verdicts.count(True), 'failed', verdicts.count(False), 'open', verdicts.count(None),
              '; model misses', sum(a != b for a, b in zip(record['model'][binding], answers[binding])))
        for case, predicted, answer, verdict in zip(record['cases'], record['rule'][binding], answers[binding], verdicts):
            if verdict is False:
                print('  ', binding, case[2], case[4], 'rule', predicted, 'excel', answer)


if __name__ == '__main__':
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
