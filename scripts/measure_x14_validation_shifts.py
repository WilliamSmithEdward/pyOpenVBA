"""What inserts and deletes do to data validation Excel keeps in a worksheet's x14 extension: lists that read another sheet.

    python scripts/measure_x14_validation_shifts.py build      # Excel makes tests/fixtures/x14_validation_base.xlsx
    python scripts/measure_x14_validation_shifts.py            # the first look: tests/fixtures/x14_validation_shifts.json
    python scripts/measure_x14_validation_shifts.py predict 2  # tests/fixtures/x14_validation_shifts2_predictions.json, from the model
    python scripts/measure_x14_validation_shifts.py measure 2  # asks Excel those and scores them
    python scripts/measure_x14_validation_shifts.py pieces     # a range's pieces through a shift: tests/fixtures/validation_pieces.json
    python scripts/measure_x14_validation_shifts.py order      # where a changed rule goes: tests/fixtures/validation_list_order.json
    python scripts/measure_x14_validation_shifts.py formula    # where a rule whose formula changes goes: validation_formula_order.json

The base book has numbers in Sheet1!A1:H12 and a list in Sheet2!A1:A6. Sheet1 holds two lists that read Sheet2, which Excel
keeps in the x14 extension (D2:E6 from Sheet2!$A$1:$A$5, and G3:G8 from Sheet2!$A$2:$A$4), and two that read its own sheet or
none, which it keeps in the plain dataValidations (a list on B2:B4 from $A$1:$A$3, and whole numbers on B7:C9). Each case opens a copy, makes one edit on either sheet,
reads every validated cell of Sheet1!A1:H16 (its type and Formula1), saves the book, and keeps both elements of Sheet1's part:
dataValidations and the extension's x14:dataValidations, as the saved file spells them.
"""
from __future__ import annotations

import json
import random
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / 'tests' / 'fixtures'
BASE = FIXTURES / 'x14_validation_base.xlsx'
BUILD = r'''Public Function Build(ByVal path As String) As String
    Dim wb As Object, ws As Object, other As Object, i As Long, j As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Sheet1"
    Set other = wb.Worksheets.Add(After:=ws)
    other.Name = "Sheet2"
    For i = 1 To 6
        other.Cells(i, 1).Value = "item" & i
    Next
    For i = 1 To 12
        For j = 1 To 8
            ws.Cells(i, j).Value = i * 10 + j
        Next
    Next
    ws.Range("D2:E6").Validation.Add Type:=xlValidateList, Formula1:="=Sheet2!$A$1:$A$5"
    ws.Range("G3:G8").Validation.Add Type:=xlValidateList, Formula1:="=Sheet2!$A$2:$A$4"
    ws.Range("B2:B4").Validation.Add Type:=xlValidateList, Formula1:="=$A$1:$A$3"
    ws.Range("B7:C9").Validation.Add Type:=xlValidateWholeNumber, Operator:=xlBetween, Formula1:="1", Formula2:="100"
    ws.Activate
    wb.SaveAs path, 51
    wb.Close False
    Build = "built"
End Function
'''
SOURCE = r'''Private Function Rules(ws As Object) As String
    Dim c As Object, out As String, kind As Long, formula As String
    For Each c In ws.Range("A1:H16").Cells
        kind = -1
        formula = ""
        On Error Resume Next
        kind = c.Validation.Type
        formula = c.Validation.Formula1
        On Error GoTo 0
        If kind <> -1 Then out = out & c.Address(False, False) & ":" & kind & ":" & formula & ";"
    Next
    Rules = out
End Function

Public Function Go(ByVal path As String, ByVal saved As String, ByVal edit As String, ByVal typed As Boolean) As String
    Dim wb As Object, ws As Object, wt As Worksheet, place As Object, out As String, parts() As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path)
    Set ws = wb.Worksheets("Sheet1")
    parts = Split(edit, " ")
    If parts(0) = "other" Then
        Set place = wb.Worksheets("Sheet2")
    Else
        Set place = ws
    End If
    Set wt = place
    On Error Resume Next
    Select Case parts(1) & " " & parts(3)
        Case "insert down": If typed Then wt.Range(parts(2)).Insert Shift:=xlShiftDown Else place.Range(parts(2)).Insert Shift:=xlShiftDown
        Case "insert right": If typed Then wt.Range(parts(2)).Insert Shift:=xlShiftToRight Else place.Range(parts(2)).Insert Shift:=xlShiftToRight
        Case "delete up": If typed Then wt.Range(parts(2)).Delete Shift:=xlShiftUp Else place.Range(parts(2)).Delete Shift:=xlShiftUp
        Case "delete left": If typed Then wt.Range(parts(2)).Delete Shift:=xlShiftToLeft Else place.Range(parts(2)).Delete Shift:=xlShiftToLeft
        Case "insert lines": If typed Then wt.Range(parts(2)).Insert Else place.Range(parts(2)).Insert
        Case "delete lines": If typed Then wt.Range(parts(2)).Delete Else place.Range(parts(2)).Delete
    End Select
    out = "err=" & Err.Number & " rules=" & Rules(ws)
    Err.Clear
    On Error GoTo 0
    wb.SaveAs saved, 51
    wb.Close False
    Go = out
End Function
'''
EDITS = ('sheet insert D3 down', 'sheet insert D1 down', 'sheet insert A1:H1 down', 'sheet insert C1:E1 down', 'sheet delete D3 up',
         'sheet delete D2 up', 'sheet delete D1:H1 up', 'sheet delete A3:H3 up', 'sheet delete D2:E6 up', 'sheet insert C2:C6 right',
         'sheet insert A3:A4 right', 'sheet delete A2:C8 left', 'sheet delete D2:D6 left', 'sheet delete E3 left', 'sheet insert 3:3 lines',
         'sheet delete 3:3 lines', 'sheet delete 2:6 lines', 'sheet insert D:D lines', 'sheet delete E:E lines', 'sheet delete D:D lines',
         'sheet insert E3 down', 'sheet delete E3 up', 'sheet insert D4:E4 down', 'sheet insert E4 right', 'sheet delete D3:D4 left',
         'sheet insert B8 down', 'sheet delete C8 left', 'sheet insert C8:C9 right', 'sheet insert A5:H5 down', 'sheet delete B3:E3 up',
         'sheet insert D2:E2 down', 'sheet delete E2:E6 left', 'sheet insert G5 down', 'sheet delete A4:H4 left', 'sheet insert B2 right',
         'other insert A1 down', 'other insert A3 down', 'other delete A2 up', 'other delete A1:A6 up', 'other insert 1:1 lines',
         'other delete 3:3 lines', 'other insert A:A lines', 'other insert A2 right')


def first_look() -> list[list[object]]:
    return [[edit, typed] for edit in EDITS for typed in (True, False)]


def held_out(round_: str) -> list[list[object]]:
    chooser = random.Random('x14-validation' + round_)
    columns = 'ABCDEFGH'
    cases: list[list[object]] = []
    while len(cases) < 100:
        where = chooser.choice(('sheet', 'sheet', 'sheet', 'other'))
        verb = chooser.choice(('insert', 'delete'))
        kind = chooser.choice(('cells', 'cells', 'lines'))
        if kind == 'lines':
            if chooser.random() < .5:
                top = chooser.randint(1, 10)
                target = f'{top}:{top + chooser.choice((0, 0, 1, 3))}'
            else:
                left = chooser.randrange(7)
                target = f'{columns[left]}:{columns[min(7, left + chooser.choice((0, 0, 1)))]}'
            edit = f'{where} {verb} {target} lines'
        else:
            left, top = chooser.randrange(7), chooser.randint(1, 10)
            target = f'{columns[left]}{top}:{columns[min(7, left + chooser.choice((0, 0, 1, 3)))]}{top + chooser.choice((0, 0, 1, 4))}'
            how = chooser.choice(('down', 'right')) if verb == 'insert' else chooser.choice(('up', 'left'))
            edit = f'{where} {verb} {target} {how}'
        cases.append([edit, chooser.random() < .5])
    return cases


def parts_of(path: Path) -> str:
    """Sheet1's validation as the saved file spells it: the plain element, then the extension's."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    plain = re.search(r'<dataValidations\b.*?</dataValidations>', xml, re.DOTALL)
    extended = re.search(r'<x14:dataValidations\b.*?</x14:dataValidations>', xml, re.DOTALL)
    return (plain.group(0) if plain else '') + ' | ' + (extended.group(0) if extended else '')


def build() -> None:
    sys.path.insert(0, str(ROOT / 'scripts'))
    from fixture_workbook import copy_saved
    from pyvbaharness import ExcelSession, HarnessConfig

    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temp:
        made = Path(temp) / 'base.xlsx'
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(BUILD, 'Build', (str(made),), timeout=60.0, line_numbers=False)
            assert result.ok, result
        copy_saved(made, BASE)
    print('built', BASE.name, BASE.stat().st_size)


def _ask(cases: list[list[object]]) -> list[str]:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: list[str] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temp:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            for number, case in enumerate(cases):
                opened, saved = Path(temp) / f'in{number}.xlsx', Path(temp) / f'out{number}.xlsx'
                shutil.copyfile(BASE, opened)
                result = excel.run_vba(SOURCE, 'Go', (str(opened), str(saved), *case), timeout=60.0, line_numbers=False)
                answers.append(f'{result.value} ~ {parts_of(saved)}' if result.ok else f'ERROR {str(result)[:200]}')
    return answers


def _write(path: Path, cases: list[list[object]], answers: list[str]) -> None:
    path.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'cases': cases, 'answers': answers},
                               indent=0), encoding='utf-8', newline='\n')


def model_answer(case: list[object], temp: Path) -> str:
    """What the model answers for one case, opening the base book from the folder given and saving beside it."""
    from pyopenvba.apps.excel import ExcelApplication

    opened, saved = temp / 'in.xlsx', temp / 'out.xlsx'
    shutil.copyfile(BASE, opened)
    saved.unlink(missing_ok=True)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='Probe')
    found = str(app.run('Probe.Go', str(opened), str(saved), *case))
    return f'{found} ~ {parts_of(saved)}'


def predict(round_: str) -> None:
    from pyopenvba.exceptions import VBAUnsupportedError

    found = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temp:
        for case in held_out(round_):
            try:
                model: str | None = model_answer(case, Path(temp))
            except VBAUnsupportedError:
                model = None
            found.append({'call': case, 'model': model})
    (FIXTURES / f'x14_validation_shifts{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0), encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused')


def measure(round_: str) -> None:
    predicted = json.loads((FIXTURES / f'x14_validation_shifts{round_}_predictions.json').read_text('utf-8'))['cases']
    cases = [one['call'] for one in predicted]
    answers = _ask(cases)
    _write(FIXTURES / f'x14_validation_shifts{round_}.json', cases, answers)
    wrong = [(one['call'], one['model'], answer) for one, answer in zip(predicted, answers) if one['model'] is not None and one['model'] != answer]
    refused = sum(one['model'] is None for one in predicted)
    print(f'{len(answers) - len(wrong) - refused} right, {len(wrong)} wrong, {refused} refused of {len(answers)}')
    for one in wrong[:20]:
        print(str(one)[:700])


# --- two narrower first looks, each in a new book: a range's pieces, and where a changed rule goes in the list ------------

#: One rule of whole numbers on the range given, then one edit; the file's sqref is read back.
PIECES_SOURCE = r'''Public Function Go(ByVal areas As String, ByVal edit As String, ByVal path As String) As String
    Dim wb As Object, ws As Object, parts() As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range(areas).Validation.Add Type:=xlValidateWholeNumber, Operator:=xlBetween, Formula1:="1", Formula2:="100"
    parts = Split(edit, " ")
    Select Case parts(1)
        Case "down": ws.Range(parts(0)).Insert Shift:=xlShiftDown
        Case "right": ws.Range(parts(0)).Insert Shift:=xlShiftToRight
        Case "up": ws.Range(parts(0)).Delete Shift:=xlShiftUp
        Case "left": ws.Range(parts(0)).Delete Shift:=xlShiftToLeft
    End Select
    wb.SaveAs path, 51
    wb.Close False
    Go = "ok"
End Function
'''
PIECE_RULES = ('D2:D6', 'D2:E6', 'D2:F6', 'D3:D6', 'D2:D8', 'D4:D6', 'D2:D4', 'B4:F4', 'B4:F5', 'B4:F6', 'C4:F4', 'B4:H4', 'D4:F4')
PIECE_EDITS = ('A3:A4 right', 'A4 right', 'A2:A4 right', 'A4:A6 right', 'A5 right', 'D4 right', 'C3 right', 'A3:A4 left', 'A4 left',
               'C1 down', 'C1:D1 down', 'D1 down', 'B1:D1 down', 'E1 down', 'C3 down', 'D3 down', 'C1 up', 'D1 up', 'D3 up')
#: R1 on C3:D6 first, then R2 on A1 and R3 on A2, told apart by their Formula2; each edit changes R1 alone or moves the rest whole.
ORDER_SOURCE = r'''Public Function Go(ByVal edit As String, ByVal path As String) As String
    Dim wb As Object, ws As Object, parts() As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("C3:D6").Validation.Add Type:=xlValidateWholeNumber, Operator:=xlBetween, Formula1:="1", Formula2:="100"
    ws.Range("A1").Validation.Add Type:=xlValidateWholeNumber, Operator:=xlBetween, Formula1:="1", Formula2:="200"
    ws.Range("A2").Validation.Add Type:=xlValidateWholeNumber, Operator:=xlBetween, Formula1:="1", Formula2:="300"
    parts = Split(edit, " ")
    Select Case parts(1)
        Case "down": ws.Range(parts(0)).Insert Shift:=xlShiftDown
        Case "right": ws.Range(parts(0)).Insert Shift:=xlShiftToRight
        Case "up": ws.Range(parts(0)).Delete Shift:=xlShiftUp
        Case "left": ws.Range(parts(0)).Delete Shift:=xlShiftToLeft
        Case "insert": ws.Range(parts(0)).Insert
        Case "delete": ws.Range(parts(0)).Delete
    End Select
    wb.SaveAs path, 51
    wb.Close False
    Go = "ok"
End Function
'''
ORDER_EDITS = ('C1:D1 down', 'C4:D4 down', 'C7:D7 down', 'C3:D3 down', 'C6:D6 down', 'C2:D2 down', 'C3:D3 up', 'C4:D4 up', 'C6:D6 up',
               'C2:D2 up', 'C4 down', 'D5 up', 'C4:D5 up', 'C5:D7 up', 'C1:D3 up', 'A3:A6 right', 'D3:D6 right', 'E3:E6 right', 'C3:C6 right',
               'C3:C6 left', 'D3:D6 left', 'B3:C6 left', 'D3:E6 left', 'A3:B6 left', 'C3:D4 down', 'C4:D6 up',
               '4:4 insert', '7:7 insert', '3:3 insert', '6:6 insert', '3:3 delete', '4:4 delete', '6:6 delete', '2:3 delete', '6:7 delete',
               'C:C delete', 'D:D delete', 'D:D insert', 'C:C insert', 'E:E insert', 'B:C delete', 'D:E delete', '1:1 insert')
#: R1, a list on C2:C4 reading $A$3:$A$6, first, then R2 on E1 and R3 on E2; each edit shifts cells in column A alone, so only R1's
#: formula changes. Told apart by their Formula2 as above, R1's list carrying none: its own reads back as an empty Formula2.
FORMULA_SOURCE = r'''Public Function Go(ByVal edit As String, ByVal path As String) As String
    Dim wb As Object, ws As Object, parts() As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:A10").Value = 5
    ws.Range("C2:C4").Validation.Add Type:=xlValidateWholeNumber, Operator:=xlBetween, Formula1:="=MIN($A$3:$A$6)", Formula2:="100"
    ws.Range("E1").Validation.Add Type:=xlValidateWholeNumber, Operator:=xlBetween, Formula1:="1", Formula2:="200"
    ws.Range("E2").Validation.Add Type:=xlValidateWholeNumber, Operator:=xlBetween, Formula1:="1", Formula2:="300"
    parts = Split(edit, " ")
    Select Case parts(1)
        Case "down": ws.Range(parts(0)).Insert Shift:=xlShiftDown
        Case "up": ws.Range(parts(0)).Delete Shift:=xlShiftUp
    End Select
    wb.SaveAs path, 51
    wb.Close False
    Go = "ok"
End Function
'''
FORMULA_EDITS = ('A3 up', 'A4 up', 'A6 up', 'A2 up', 'A7 up', 'A3:A4 up', 'A5:A6 up', 'A5:A7 up', 'A2:A3 up', 'A3:A6 up', 'A1:A10 up',
                 'A3 down', 'A4 down', 'A6 down', 'A7 down', 'A2 down', 'A4:A5 down', 'A1 down')
NARROW = {'pieces': ('validation_pieces', PIECES_SOURCE, [[rule, edit] for rule in PIECE_RULES for edit in PIECE_EDITS]),
          'order': ('validation_list_order', ORDER_SOURCE, [[edit] for edit in ORDER_EDITS]),
          'formula': ('validation_formula_order', FORMULA_SOURCE, [[edit] for edit in FORMULA_EDITS])}


def read_rules(path: Path, *, formulas: bool = False) -> str:
    """The saved sheet's plain rules in file order, each as its Formula2 (and with ``formulas`` its Formula1) then its sqref."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.findall(r'sqref="([^"]*)"[^>]*><formula1>([^<]*)</formula1><formula2>([^<]*)</formula2>', xml)
    return ' '.join(f'{second}:{first}:{ranges}' if formulas else f'{second}:{ranges}' for ranges, first, second in found)


def narrow(name: str) -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    stem, source, cases = NARROW[name]
    answers: list[str] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temp, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for number, case in enumerate(cases):
            path = Path(temp) / f'o{number}.xlsx'
            result = excel.run_vba(source, 'Go', (*case, str(path)), timeout=60.0, line_numbers=False)
            answers.append(read_rules(path, formulas=name == 'formula') if result.ok else f'ERROR {str(result)[:200]}')
    (FIXTURES / f'{stem}.json').write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
                                                       'cases': cases, 'answers': answers}, indent=0), encoding='utf-8', newline='\n')
    print(len(cases), 'measured')


def model_narrow(source: str, case: list[object], temp: Path) -> str:
    """What the model saves for one case of a narrow look."""
    from pyopenvba.apps.excel import ExcelApplication

    path = temp / 'out.xlsx'
    path.unlink(missing_ok=True)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Probe')
    app.run('Probe.Go', *case, str(path))
    return read_rules(path, formulas='MIN($A$3:$A$6)' in source)


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'build':
        build()
    elif len(sys.argv) > 1 and sys.argv[1] in NARROW:
        narrow(sys.argv[1])
    elif len(sys.argv) > 2 and sys.argv[1] in ('predict', 'measure'):
        {'predict': predict, 'measure': measure}[sys.argv[1]](sys.argv[2])
    else:
        found = first_look()
        _write(FIXTURES / 'x14_validation_shifts.json', found, _ask(found))
        print(len(found), 'measured')
