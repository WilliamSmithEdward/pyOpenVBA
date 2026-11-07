"""An object given by Let (no Set) to a property, typed and late, in live Excel, with the model's answers and each rival rule's written
down first.

Each case makes a new workbook, puts 7 in A1, "abc" in B1 and 9 in C1, and gives one object to one property: a one-cell Range holding a
number, one holding text, one area of two cells, two areas, a Font, a Style, a Collection, or Nothing. The properties are of each kind a
macro sets: Boolean (Workbook.Date1904, Font.Bold), a Long enum (Workbook.UpdateLinks, Worksheet.Visible), String (Worksheet.Name), Double
(Worksheet.StandardWidth), Variant (Range.Value, Range.NumberFormat, Range.RowHeight, Font.Size, Font.Name, Application.StatusBar) and
one that takes an object as well as a name (Range.Style). Each answer is the set's error and the property read after, as TypeName:value.
Application.StatusBar is put back to False after its cases.

The rival rules, and what each predicts before Excel answers:
  value: VBA hands the property the object's default value, typed and late alike, so every late answer equals the typed one.
  object: bound late, the property receives the object itself and answers for it, so late answers differ from typed ones wherever the
          property does not read an object's value itself.

    python scripts/measure_late_object_lets.py predict
    python scripts/measure_late_object_lets.py measure

writes tests/fixtures/late_object_lets.json.
"""
from __future__ import annotations

import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/late_object_lets.json'
SEPARATOR = '%%'
#: Each target: the object a typed probe declares it as, how the probe reaches it, and the property.
TARGETS = [('Workbook', 'wb', 'Date1904'), ('Workbook', 'wb', 'UpdateLinks'), ('Worksheet', 'ws', 'Visible'), ('Worksheet', 'ws', 'Name'),
           ('Worksheet', 'ws', 'StandardWidth'), ('Range', 'ws.Range("D1")', 'Value'), ('Range', 'ws.Range("D1")', 'NumberFormat'),
           ('Range', 'ws.Range("D1")', 'RowHeight'), ('Range', 'ws.Range("D1")', 'Style'), ('Font', 'ws.Range("D1").Font', 'Size'),
           ('Font', 'ws.Range("D1").Font', 'Name'), ('Font', 'ws.Range("D1").Font', 'Bold'), ('Application', 'Application', 'StatusBar')]
OBJECTS = ['ws.Range("A1")', 'ws.Range("B1")', 'ws.Range("A1:C1")', 'ws.Range("A1,C1")', 'ws.Range("A1").Font', 'wb.Styles("Good")',
           'New Collection', 'Nothing']
#: Cases per procedure, so none is too large to compile.
PART = 16


#: Typed, a Let of these does not compile: a Collection's default member, Item, takes an index ("Argument not optional"), and Nothing is
#: "Invalid use of object". A module that
#: does not compile runs nothing, so a typed probe in live Excel leaves these cases out and records them as not compiling.
NOT_COMPILING = frozenset({'New Collection', 'Nothing'})


def _case(target: tuple[str, str, str], value: str, declared: bool, compiles_only: bool = False) -> str:
    kind, path, prop = target
    holder = 't' if declared else 'o'
    if compiles_only and declared and value in NOT_COMPILING:
        return f'    out = out & "{SEPARATOR}" & "does not compile"\n'
    restore = '    Application.StatusBar = False\n' if prop == 'StatusBar' else ''
    return f'''    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = 7: ws.Range("B1").Value = "abc": ws.Range("C1").Value = 9
    Set {holder}{kind} = {path}
    Err.Clear
    {holder}{kind}.{prop} = {value}
    s = "e" & Err.Number
    Err.Clear: v = Empty
    v = {holder}{kind}.{prop}
    s = s & "/" & IIf(Err.Number, "e" & Err.Number, TypeName(v) & ":" & CStr(v))
{restore}    wb.Close SaveChanges:=False
    out = out & "{SEPARATOR}" & s
'''


def cases() -> list[tuple[str, str, str, str]]:
    return [(kind, path, prop, value) for kind, path, prop in TARGETS for value in OBJECTS]


def macro(typed: bool, compiles_only: bool = False) -> str:
    every = cases()
    parts = [every[at:at + PART] for at in range(0, len(every), PART)]
    declare = ('    Dim tWorkbook As Workbook, tWorksheet As Worksheet, tRange As Range, tFont As Font, tApplication As Application\n' if typed
               else '    Dim oWorkbook As Object, oWorksheet As Object, oRange As Object, oFont As Object, oApplication As Object\n')
    procedures = ''.join(f'''Private Sub Part{index}(ByRef out As String)
    Dim wb As Workbook, ws As Worksheet, v As Variant, s As String
{declare}    On Error Resume Next
{''.join(_case((kind, path, prop), value, typed, compiles_only) for kind, path, prop, value in part)}End Sub
''' for index, part in enumerate(parts))
    calls = ''.join(f'    Part{index} out\n' for index in range(len(parts)))
    return f'''{procedures}Function Probe() As String
    Dim out As String
    Application.DisplayAlerts = False
{calls}    Application.DisplayAlerts = True
    Application.StatusBar = False
    Probe = Mid(out, {len(SEPARATOR) + 1})
End Function
'''


def predict() -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication

    answers: dict[str, list[str]] = {}
    for typed in (True, False):
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(macro(typed), name='H')
        answers['typed' if typed else 'late'] = str(app.run('H.Probe')).split(SEPARATOR)
    text = json.dumps({'cases': cases(), 'predicted': answers}, indent=1) + '\n'
    OUT.write_text(text, encoding='utf-8', newline='\n')
    same = sum(a == b for a, b in zip(answers['typed'], answers['late']))
    print('predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest(), '; model late = typed in', same, 'of', len(cases()))


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
    same = sum(a == b for a, b in zip(answers['typed'], answers['late']))
    print('excel late = typed in', same, 'of', len(cases()))
    for binding in answers:
        print(binding, 'model misses', sum(a != b for a, b in zip(record['predicted'][binding], answers[binding])))


if __name__ == '__main__':
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
