"""Every property the type library lists on the sheet's old collections and their items, each read typed and late, on a workbook Excel saved.

Worksheet.Arcs, Buttons, CheckBoxes, DropDowns, Lines, ListBoxes, OptionButtons, Ovals, Pictures, Rectangles and TextBoxes
were measured on items a macro had just added, through an Object. Rival rules for a typed read: the same answer as the late
one, or the answer differing as it does on other classes (tests/fixtures/settable_audit*.json). Rival rules for an item from
a file: the same as a new item, or differing where the file leaves a setting out. Left, Top, Width and Height and the cells
under an item are left out, as they depend on the display's DPI (see docs/live_measurement.md if present).

    python scripts/measure_old_collection_members.py

makes tests/fixtures/old_collections.xlsx in Excel and writes tests/fixtures/old_collection_members.json, which
tests/test_excel_old_collection_members.py replays.
"""
from __future__ import annotations

import csv
import json
import struct
import sys
import tempfile
import time
import zlib
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig, HarnessError

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / 'tests/fixtures/old_collection_members.json'
BOOK = ROOT / 'tests/fixtures/old_collections.xlsx'
KINDS = {'Arcs': 'Arc', 'Buttons': 'Button', 'CheckBoxes': 'CheckBox', 'DropDowns': 'DropDown', 'Lines': 'Line', 'ListBoxes': 'ListBox',
         'OptionButtons': 'OptionButton', 'Ovals': 'Oval', 'Pictures': 'Picture', 'Rectangles': 'Rectangle', 'TextBoxes': 'TextBox'}
SKIPPED = {'Left', 'Top', 'Width', 'Height', 'BottomRightCell', 'TopLeftCell'}
SCALARS = {'String', 'Long', 'Integer', 'Boolean', 'Double', 'Single', 'Variant', 'Date', 'Byte', 'Currency'}
COLLECTION_EXTRAS = ['o(1).Name', 'o.Item(1).Name', 'o(2).Name', 'o(o(1).Name).Name', 'o(0).Name', 'o(99).Name', 'o("nope").Name', 'o.Item(0).Name',
                     'o.Item(99).Name', 'o.Item("nope").Name', 'TypeName(o(1))', 'TypeName(o(Array(1, 2)))', 'o(Array(1, 2)).Count']
#: A typed o(1) on these does not compile, "Wrong number of arguments or invalid property assignment": each was run once, and each costs a new Excel.
NO_INDEX = {'CheckBoxes', 'DropDowns', 'ListBoxes', 'OptionButtons'}
COMPILE_ERROR = 'compile Wrong number of arguments or invalid property assignment'
PARTIAL = Path(tempfile.gettempdir()) / 'old_collection_members_partial.json'
#: Reads through what a collection gives as an object.
THROUGH = ['o.Font.Bold', 'o.Font.Name', 'o.Font.Size', 'o.Font.ColorIndex', 'o.Interior.ColorIndex', 'o.Interior.Pattern', 'o.Border.LineStyle', 'o.Border.ColorIndex',
           'o.Border.Weight', 'o.Characters.Text', 'o.Characters.Count', 'o.ShapeRange.Count', 'o.ShapeRange(1).Name', 'o.ShapeRange.Name', 'TypeName(o.ShapeRange(1))']
MAKER = r'''Public Sub Make(ByVal picture As String, ByVal out As String)
    Dim wb As Workbook, ws As Worksheet, i As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    For i = 1 To 2
        ws.Arcs.Add 10, 10 + 60 * i, 60, 50 + 60 * i
        ws.Buttons.Add 100, 10 + 60 * i, 60, 30
        ws.CheckBoxes.Add 180, 10 + 60 * i, 60, 30
        ws.DropDowns.Add 260, 10 + 60 * i, 60, 20
        ws.Lines.Add 340, 10 + 60 * i, 400, 40 + 60 * i
        ws.ListBoxes.Add 420, 10 + 60 * i, 60, 50
        ws.OptionButtons.Add 500, 10 + 60 * i, 60, 30
        ws.Ovals.Add 580, 10 + 60 * i, 60, 30
        ws.Rectangles.Add 660, 10 + 60 * i, 60, 30
        ws.TextBoxes.Add 740, 10 + 60 * i, 60, 30
        ws.Shapes.AddPicture picture, False, True, 820, 10 + 60 * i, 20, 20
    Next
    ws.Range("A20").Value = "one": ws.Range("A21").Value = "two": ws.Range("A22").Value = "three"
    ws.Buttons(1).Caption = "Press"
    ws.CheckBoxes(1).Caption = "Tick": ws.CheckBoxes(1).Value = xlOn: ws.CheckBoxes(1).LinkedCell = "$C$20"
    ws.OptionButtons(1).Caption = "Pick": ws.OptionButtons(1).Value = xlOn
    ws.DropDowns(1).ListFillRange = "$A$20:$A$22": ws.DropDowns(1).Value = 2: ws.DropDowns(1).LinkedCell = "$C$21"
    ws.ListBoxes(1).ListFillRange = "$A$20:$A$22": ws.ListBoxes(1).Value = 3
    ws.TextBoxes(1).Text = "Boxed": ws.TextBoxes(2).Formula = "=$A$20"
    ws.Rectangles(1).Text = "Rect": ws.Ovals(1).Text = "Round"
    ws.Rectangles(2).Interior.ColorIndex = 6: ws.Rectangles(2).Border.ColorIndex = 3
    ws.Ovals(2).Visible = False: ws.Lines(2).Locked = False: ws.Arcs(2).PrintObject = False: ws.Buttons(2).Enabled = False
    ws.Pictures(2).Placement = xlFreeFloating
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
End Sub
'''
SHOW = r'''Private Function Show(v As Variant) As String
    If IsObject(v) Then
        Show = "obj " & TypeName(v)
    ElseIf IsArray(v) Then
        Show = "array"
    ElseIf IsNull(v) Then
        Show = "null"
    ElseIf IsEmpty(v) Then
        Show = "empty"
    ElseIf IsError(v) Then
        Show = "error " & CStr(v)
    Else
        Show = TypeName(v) & " " & CStr(v)
    End If
End Function
'''


def members() -> dict[str, list[str]]:
    """Each class's reads, as expressions on ``o``."""
    found: dict[str, list[str]] = {}
    with (ROOT / 'docs/excel_checklist.csv').open(encoding='utf-8', newline='') as file:
        for row in csv.DictReader(file):
            name = row['object']
            if name not in KINDS and name not in KINDS.values():
                continue
            member, kind, signature = row['member'], row['kind'], row['signature']
            if not kind.startswith('property') or member in SKIPPED or member.startswith('_') or signature.startswith('('):
                continue
            returned = signature.rsplit(' As ', 1)[-1].strip() if ' As ' in signature else 'Variant'
            found.setdefault(name, []).append(f'o.{member}' if returned in SCALARS else f'TypeName(o.{member})')
    return found


def procedure(title: str, name: str, get: str, binding: str, reads: list[str]) -> str:
    declared, sheet = (f'Excel.{name}', 'Worksheet') if binding == 'Typed' else ('Object', 'Object')
    body = ''.join(f'    Case {index}\n        v = {expression}\n' for index, expression in enumerate(reads))
    return (f'Public Function {title}(ws As {sheet}, ByVal which As Long) As String\n    Dim o As {declared}, v As Variant\n    Set o = {get}\n'
            f'    On Error Resume Next\n    Select Case which\n{body}    End Select\n    If Err.Number Then\n        {title} = "E" & Err.Number\n'
            f'    Else\n        {title} = Show(v)\n    End If\nEnd Function\n\n')


def reader(table: dict[str, list[str]]) -> tuple[str, list[dict[str, str | int]]]:
    """The module and its cases. A form that may not compile typed gets a procedure of its own, as a compile error stops only the procedure run."""
    parts = [SHOW]
    cases: list[dict[str, str | int]] = []
    for collection, item in KINDS.items():
        for name, get in ((collection, f'ws.{collection}'), (item, f'ws.{collection}(1)')):
            for binding in ('Typed', 'Late'):
                reads = table.get(name, [])
                title = f'{binding}{name}'
                parts.append(procedure(title, name, get, binding, reads))
                cases.extend({'state': 'first', 'class': name, 'binding': binding, 'read': read, 'routine': title, 'which': index} for index, read in enumerate(reads))
                if name == collection:
                    for index, read in enumerate(COLLECTION_EXTRAS):
                        parts.append(procedure(f'{title}X{index}', name, get, binding, [read]))
                        cases.append({'state': 'first', 'class': name, 'binding': binding, 'read': read, 'routine': f'{title}X{index}', 'which': 0})
    # Item 2 of each kind holds the settings a first item leaves at their defaults; a collection down to one item may answer for it;
    # TypeName tells an object with no default member from no member, which read alike as 438.
    for collection, item in KINDS.items():
        reads = table.get(item, [])
        parts.append(procedure(f'Second{item}', item, f'ws.{collection}(2)', 'Late', reads))
        cases.extend({'state': 'second', 'class': item, 'binding': 'Late', 'read': read, 'routine': f'Second{item}', 'which': index} for index, read in enumerate(reads))
        parts.append(f'Public Function Trim{collection}(ws As Object, ByVal which As Long) As String\n    ws.{collection}(2).Delete\nEnd Function\n\n')
        reads = table.get(collection, [])
        cases.extend({'state': 'one', 'class': collection, 'binding': 'Late', 'read': read, 'routine': f'Late{collection}', 'which': index, 'trim': f'Trim{collection}'}
                     for index, read in enumerate(reads))
        for name, get in ((collection, f'ws.{collection}'), (item, f'ws.{collection}(1)')):
            kinds = [read if read.startswith('TypeName') else f'TypeName({read})' for read in table.get(name, [])]
            parts.append(procedure(f'Kinds{name}', name, get, 'Late', kinds))
            cases.extend({'state': 'kind', 'class': name, 'binding': 'Late', 'read': read, 'routine': f'Kinds{name}', 'which': index} for index, read in enumerate(kinds))
    for collection in KINDS:
        parts.append(procedure(f'Through{collection}', collection, f'ws.{collection}', 'Late', THROUGH))
        for state, trim in (('through', {}), ('through one', {'trim': f'Trim{collection}'})):
            cases.extend({'state': state, 'class': collection, 'binding': 'Late', 'read': read, 'routine': f'Through{collection}', 'which': index, **trim} for index, read in enumerate(THROUGH))
    for collection, item in KINDS.items():
        for place in (1, 2):
            parts.append(procedure(f'Through{item}{place}', item, f'ws.{collection}({place})', 'Late', THROUGH))
            cases.extend({'state': f'through item {place}', 'class': item, 'binding': 'Late', 'read': read, 'routine': f'Through{item}{place}', 'which': index} for index, read in enumerate(THROUGH))
    return ''.join(parts), cases


def tiny_png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))

    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\x00\xff\x00\x00')) + chunk(b'IEND', b'')


def run(session: list[ExcelSession], source: str, routine: str, *arguments: object) -> tuple[str | None, str]:
    """The answer of one walk, or None and why not. A blocked or dead Excel gets a new session for the next."""
    for _ in range(3):
        try:
            done = session[0].run_vba(source, routine, arguments, timeout=60.0, line_numbers=False)
        except HarnessError as failed:
            print('restarting after', str(failed)[:100], flush=True)
            session[0].close()
            time.sleep(5)
            session[0] = ExcelSession(HarnessConfig(exclusive=False))
            session[0].new_document()
            continue
        if done.ok:
            return str(done.value), done.outcome
        return None, f'{done.outcome} {[one.message[:120] for one in done.dialogs]}'
    return None, 'no session'


def main() -> None:
    source, cases = reader(members())
    driver = ('\nPublic Function Walk(ByVal path As String, ByVal trim As String, ByVal list As String) As String\n'
              '    Dim wb As Workbook, one As Variant, s As String\n    Application.DisplayAlerts = False\n    Set wb = Workbooks.Open(path, 0)\n'
              '    If Len(trim) Then Application.Run "\'" & ThisWorkbook.Name & "\'!" & trim, wb.Worksheets(1), 0\n'
              '    For Each one In Split(list, "|")\n        s = s & Application.Run("\'" & ThisWorkbook.Name & "\'!" & Split(one, ":")(0), wb.Worksheets(1), CLng(Split(one, ":")(1))) & "~~"\n    Next\n'
              '    wb.Close False\n    Walk = s\nEnd Function\n')

    def call(chosen: list[dict[str, str | int]]) -> str:
        return '|'.join(f'{one["routine"]}:{one["which"]}' for one in chosen)

    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
        session = [ExcelSession(HarnessConfig(exclusive=False))]
        session[0].new_document()
        picture, saved = Path(folder) / 'dot.png', Path(folder) / 'old_collections.xlsx'
        picture.write_bytes(tiny_png())
        if not BOOK.exists():
            made = session[0].run_vba(MAKER, 'Make', (str(picture), str(saved)), timeout=60.0, line_numbers=False)
            assert made.ok, made
            BOOK.write_bytes(saved.read_bytes())
            strip_save_path(BOOK)
        done: dict[str, str] = json.loads(PARTIAL.read_text(encoding='utf-8')) if PARTIAL.exists() else {}
        if not done and OUT.exists():
            done = {f'{one["state"]} {one["class"]} {one["binding"]} {one["read"]}': one['excel'] for one in json.loads(OUT.read_text(encoding='utf-8'))['cases']}
        groups: dict[tuple[str, str], list[dict[str, str | int]]] = {}
        for case in cases:
            groups.setdefault((str(case['state']), str(case['class']), str(case['binding'])), []).append(case)
        for (state, name, binding), whole in groups.items():
            group: list[dict[str, str | int]] = []
            for case in whole:
                key = f'{state} {name} {binding} {case["read"]}'
                if key in done:
                    case['excel'] = done[key]
                elif binding == 'Typed' and name in NO_INDEX and str(case['read']).startswith(('o(', 'TypeName(o(')):
                    case['excel'] = COMPILE_ERROR
                else:
                    group.append(case)
            if not group:
                continue
            trim = str(group[0].get('trim', ''))
            value, outcome = run(session, source + driver, 'Walk', str(BOOK), trim, call(group))
            print(state, name, binding, outcome[:150], len(group), flush=True)
            if value is not None:
                for case, answer in zip(group, value.split('~~'), strict=False):
                    case['excel'] = answer
            else:
                for case in group:
                    one, why = run(session, source + driver, 'Walk', str(BOOK), trim, call([case]))
                    case['excel'] = one.split('~~')[0] if one is not None else f'failed {why}'
                    if one is None:
                        print('  ', case['read'], why[:170], flush=True)
            done.update({f'{state} {name} {binding} {one["read"]}': str(one['excel']) for one in whole})
            PARTIAL.write_text(json.dumps(done), encoding='utf-8')
        session[0].close()
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'cases': cases}, indent=0) + '\n', encoding='utf-8')
    failed = [one for one in cases if str(one['excel']).startswith('failed')]
    print(len(cases), 'cases', len(failed), 'failed')
    if not failed:
        PARTIAL.unlink()


if __name__ == '__main__':
    main()
