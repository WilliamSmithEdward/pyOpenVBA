"""A property read on a whole old collection, typed and late, with answer types: on the workbook Excel saved
(tests/fixtures/old_collections.xlsx) and on one the macro makes in the session, as each is, with every item set alike, and with the second
item alone set.

Rival rules, written before measuring: a read over items that agree answers their value, or 1004 as it does over items that differ;
a book the macro made answers as one from a file, or differs where the file leaves a setting out. A member that cannot be set is
read as the book has it. Left, Top, Width and Height are left out, as they depend on the display's DPI.

Each walk runs one collection, binding, book and state, a fresh book for each member. Typed, a member the collection's type does not
have does not compile, so only members the checklist lists on that collection are read, typed and late alike.

    python scripts/measure_old_collection_reads.py [seconds]

writes tests/fixtures/old_collection_reads.json, which tests/test_excel_old_collection_reads.py replays. It keeps its place in a file
under TEMP, so a run that stops can go on.
"""
from __future__ import annotations

import csv
import importlib.util
import json
import struct
import sys
import tempfile
import time
import zlib
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/old_collection_reads.json'
BOOK = ROOT / 'tests/fixtures/old_collections.xlsx'
PARTIAL = Path(tempfile.gettempdir()) / 'old_collection_reads_partial.json'
BUDGET = 520.0
COLLECTIONS = ('Arcs', 'Buttons', 'CheckBoxes', 'DropDowns', 'Lines', 'ListBoxes', 'OptionButtons', 'Ovals', 'Pictures', 'Rectangles', 'TextBoxes', 'DrawingObjects')
SKIPPED = {'Left', 'Top', 'Width', 'Height', 'BottomRightCell', 'TopLeftCell', '_Default', '_NewEnum'}
SCALARS = {'String', 'Long', 'Integer', 'Boolean', 'Double', 'Single', 'Variant', 'Date', 'Byte', 'Currency'}
_lets = importlib.util.spec_from_file_location('measure_old_collection_lets', ROOT / 'scripts/measure_old_collection_lets.py')
assert _lets is not None and _lets.loader is not None
_lets_module = importlib.util.module_from_spec(_lets)
_lets.loader.exec_module(_lets_module)
VALUES: dict[str, str] = getattr(_lets_module, 'VALUES')
#: The book a macro makes: two of each kind, as old_collections.xlsx was made but with nothing set on them.
MAKE = r'''Private Function Made(ByVal picture As String) As Workbook
    Dim wb As Workbook, ws As Worksheet, i As Long
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
    Set Made = wb
End Function

Private Function Show(ByVal v As Variant) As String
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


def members() -> dict[str, list[tuple[str, bool]]]:
    """Each collection's properties, and whether each answers a value (or an object, read by its TypeName)."""
    found: dict[str, list[tuple[str, bool]]] = {}
    with (ROOT / 'docs/excel_checklist.csv').open(encoding='utf-8', newline='') as file:
        for row in csv.DictReader(file):
            if row['object'] not in COLLECTIONS or not row['kind'].startswith('property') or row['member'] in SKIPPED or row['signature'].startswith('('):
                continue
            returned = row['signature'].rsplit(' As ', 1)[-1].strip() if ' As ' in row['signature'] else 'Variant'
            found.setdefault(row['object'], []).append((row['member'], returned in SCALARS))
    return found


def walk_source(collection: str, binding: str, book: str, state: str, chosen: list[tuple[str, bool]]) -> str:
    """One walk: for each member a fresh book, its two items set as ``state`` says, and the collection's read."""
    declared = f'Excel.{collection}' if binding == 'Typed' else 'Object'
    opened = 'Set wb = Workbooks.Open(path, 0)' if book == 'saved' else 'Set wb = Made(picture)'
    body = ''
    for index, (member, scalar) in enumerate(chosen):
        value = VALUES.get(member)
        sets = ''
        if state != 'as is' and value is not None:
            # Agree: every item set alike, which for DrawingObjects is every kind on the sheet. Differ: only the second.
            change = (f'        For Each x In ws.{collection}\n            x.{member} = {value}\n        Next\n' if state == 'agree'
                      else f'        ws.{collection}(2).{member} = {value}\n')
            sets = f'        On Error Resume Next\n{change}        s = s & "set" & Err.Number & " "\n        Err.Clear\n        On Error GoTo 0\n'
        read = f'o.{member}' if scalar else f'TypeName(o.{member})'
        body += (f'    Case {index}\n{sets}        On Error Resume Next\n        v = {read}\n'
                 f'        If Err.Number Then s = s & "E" & Err.Number Else s = s & Show(v)\n        On Error GoTo 0\n')
    return MAKE + f'''
Public Function Walk(ByVal path As String, ByVal picture As String, ByVal count As Long) As String
    Dim wb As Workbook, ws As Object, o As {declared}, v As Variant, s As String, i As Long, x As Object
    Application.DisplayAlerts = False
    For i = 0 To count - 1
        {opened}
        Set ws = wb.Worksheets(1)
        Set o = ws.{collection}
        Select Case i
{body}        End Select
        s = s & "~~"
        wb.Close False
    Next
    Walk = s
End Function
'''


def walks() -> list[dict[str, object]]:
    table = members()
    found: list[dict[str, object]] = []
    for collection in COLLECTIONS:
        everything = table.get(collection, [])
        settable = [one for one in everything if one[0] in VALUES]
        for book in ('saved', 'made'):
            for state, chosen in (('as is', everything), ('agree', settable), ('differ', settable)):
                for binding in ('Typed', 'Late'):
                    if chosen:
                        found.append({'collection': collection, 'book': book, 'state': state, 'binding': binding, 'members': [one[0] for one in chosen],
                                      'scalars': [one[1] for one in chosen]})
    return found


def tiny_png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))

    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\x00\xff\x00\x00')) + chunk(b'IEND', b'')


def source_of(walk: dict[str, object]) -> str:
    chosen = list(zip(walk['members'], walk['scalars']))  # type: ignore[call-overload]
    return walk_source(str(walk['collection']), str(walk['binding']), str(walk['book']), str(walk['state']), chosen)  # pyright: ignore[reportArgumentType]


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig, HarnessError

    budget = float(sys.argv[1]) if len(sys.argv) > 1 else BUDGET
    every = walks()
    done: dict[str, list[str]] = json.loads(PARTIAL.read_text(encoding='utf-8')) if PARTIAL.exists() else {}
    started = time.monotonic()
    session = ExcelSession(HarnessConfig(exclusive=False))
    session.new_document()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        picture = Path(held) / 'dot.png'
        picture.write_bytes(tiny_png())
        book = Path(held) / 'old_collections.xlsx'
        book.write_bytes(BOOK.read_bytes())
        for walk in every:
            key = f'{walk["collection"]} {walk["book"]} {walk["state"]} {walk["binding"]}'
            if key in done:
                continue
            if time.monotonic() - started > budget:
                print('stopping for now', len(done), 'of', len(every), flush=True)
                break
            count = len(walk['members'])  # type: ignore[arg-type]
            try:
                ran = session.run_vba(source_of(walk), 'Walk', (str(book), str(picture), count), timeout=120.0, line_numbers=False)
            except HarnessError as failed:
                print('restarting', key, str(failed)[:80], flush=True)
                session.close()
                time.sleep(5)
                session = ExcelSession(HarnessConfig(exclusive=False))
                session.new_document()
                done[key] = [f'failed harness {str(failed)[:60]}'] * count
                continue
            answers = str(ran.value).split('~~')[:count] if ran.ok and not ran.dialogs else [f'failed {ran.outcome}'] * count
            done[key] = answers
            print(key, answers[:4], flush=True)
            PARTIAL.write_text(json.dumps(done), encoding='utf-8')
    session.close()
    if len(done) == len(every):
        cases = [{'collection': walk['collection'], 'book': walk['book'], 'state': walk['state'], 'binding': walk['binding'], 'member': member, 'excel': answer}
                 for walk in every for member, answer in zip(walk['members'], done[f'{walk["collection"]} {walk["book"]} {walk["state"]} {walk["binding"]}'], strict=True)]  # type: ignore[call-overload]
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': cases}, indent=0) + '\n', encoding='utf-8')
        PARTIAL.unlink()
        print('wrote', len(cases), 'cases', flush=True)


if __name__ == '__main__':
    main()
