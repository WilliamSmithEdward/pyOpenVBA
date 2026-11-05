"""What a font's held-out round asks: the same VBA run in the model before Excel is asked (`font_round.py predict Arial`), then in Excel (`font_round.py measure Arial`).

Everything that reads the width tables is asked: column AutoFit, what a narrow column shows, wrapped rows, plain rows, Justify, a pivot table's columns, a shape's lines, and the Normal style in the font. Writes tests/fixtures/font_round_<font>_predictions.json
and tests/fixtures/font_round_<font>.json. A measuring run keeps what it has and goes on from there, and stops if StandardHeight is not 15 at either end of a call.
"""
from __future__ import annotations

import json
import random
import sys
from datetime import datetime, timezone
from pathlib import Path

FIXTURES = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures'
SOURCE = r'''Private Sub Dress(ByVal c As Object, ByVal font As String, ByVal size As Double, ByVal isBold As Boolean, ByVal isItalic As Boolean, ByVal indent As Long)
    c.Font.Name = font
    c.Font.Size = size
    c.Font.Bold = isBold
    c.Font.Italic = isItalic
    If indent > 0 Then c.IndentLevel = indent
End Sub

Private Function Fresh() As Object
    Dim ws As Object, s As Object
    Set ws = ActiveSheet
    For Each s In ws.Shapes
        s.Delete
    Next
    ws.Cells.Clear
    ws.Rows(1).UseStandardHeight = True
    ws.Rows(2).UseStandardHeight = True
    ws.Columns(1).UseStandardWidth = True
    ws.Columns(2).UseStandardWidth = True
    Set Fresh = ws
End Function

Private Function Said(ByVal x As Variant) As String
    Said = TypeName(x) & ":" & CStr(x)
End Function

Public Function Fit(ByVal text As String, ByVal font As String, ByVal size As Double, ByVal isBold As Boolean, ByVal isItalic As Boolean, ByVal indent As Long) As String
    Dim ws As Object
    Set ws = Fresh()
    ws.Range("A1").Value = "'" & text
    Dress ws.Range("A1"), font, size, isBold, isItalic, indent
    ws.Columns(1).AutoFit
    Fit = Said(ws.Columns(1).ColumnWidth) & "|" & Said(ws.Columns(1).Width) & "|" & Said(ws.Rows(1).RowHeight)
End Function

Public Function Shown(ByVal value As Variant, ByVal format As String, ByVal font As String, ByVal size As Double, ByVal isBold As Boolean, ByVal isItalic As Boolean, ByVal widths As String) As String
    Dim ws As Object, one As Variant, r As String
    Set ws = Fresh()
    ws.Range("A1").NumberFormat = format
    ws.Range("A1").Value = value
    Dress ws.Range("A1"), font, size, isBold, isItalic, 0
    For Each one In Split(widths, ",")
        ws.Columns(1).ColumnWidth = Val(one)
        r = r & "|" & ws.Range("A1").Text
    Next
    Shown = r
End Function

Public Function Wrapped(ByVal text As String, ByVal font As String, ByVal size As Double, ByVal isBold As Boolean, ByVal isItalic As Boolean, ByVal indent As Long, ByVal widths As String) As String
    Dim ws As Object, one As Variant, r As String
    Set ws = Fresh()
    ws.Range("A1").WrapText = True
    Dress ws.Range("A1"), font, size, isBold, isItalic, indent
    For Each one In Split(widths, ",")
        ws.Columns(1).ColumnWidth = Val(one)
        ws.Range("A1").Value = "'" & Replace(text, "~", vbLf)
        r = r & "|" & ws.Rows(1).RowHeight
        ws.Columns(1).ColumnWidth = Val(one) + 2
        r = r & "," & ws.Rows(1).RowHeight
        ws.Rows(1).AutoFit
        r = r & "," & ws.Rows(1).RowHeight
    Next
    Wrapped = r
End Function

Public Function Plain(ByVal font As String, ByVal size As Double, ByVal isBold As Boolean, ByVal isItalic As Boolean) As String
    Dim ws As Object
    Set ws = Fresh()
    ws.Range("A1").Value = "x"
    Dress ws.Range("A1"), font, size, isBold, isItalic, 0
    Plain = Said(ws.Rows(1).RowHeight)
End Function

Public Function Lined(ByVal text As String, ByVal font As String, ByVal size As Double, ByVal isBold As Boolean, ByVal isItalic As Boolean, ByVal widths As String) As String
    Dim ws As Object, s As Object, one As Variant, r As String, i As Long
    Set ws = Fresh()
    Set s = ws.Shapes.AddShape(1, 10, 10, 100, 400)
    s.TextFrame2.TextRange.Text = text
    With s.TextFrame2.TextRange.Font
        .Name = font
        .Size = size
        .Bold = IIf(isBold, msoTrue, msoFalse)
        .Italic = IIf(isItalic, msoTrue, msoFalse)
    End With
    For Each one In Split(widths, ",")
        s.Width = CDbl(one)
        r = r & "|"
        For i = 1 To s.TextFrame2.TextRange.Lines.Count
            r = r & s.TextFrame2.TextRange.Lines(i).Length & ","
        Next
    Next
    s.Delete
    Lined = r
End Function

Public Function Justified(ByVal text As String, ByVal font As String, ByVal size As Double, ByVal width As Double) As String
    Dim ws As Object, r As String, i As Long
    Set ws = Fresh()
    ws.Columns(1).ColumnWidth = width
    ws.Range("A1").Value = text
    Dress ws.Range("A1"), font, size, False, False, 0
    ws.Range("A1:A12").Justify
    For i = 1 To 12
        r = r & "|" & ws.Cells(i, 1).Value
    Next
    Justified = r
End Function

Public Function Pivoted(ByVal font As String, ByVal size As Double, ByVal isBold As Boolean) As String
    Dim ws As Object, p As Object, wb As Object, data As Object, r As String, i As Long
    Set wb = ActiveWorkbook
    Set data = wb.Worksheets.Add
    data.Range("A1:B1").Value = Array("Region of sale", "Amount")
    data.Range("A2:B2").Value = Array("North east corner", 1234.5)
    data.Range("A3:B3").Value = Array("South", 99)
    Dress data.Range("A1:B3"), font, size, isBold, False, 0
    Set ws = wb.Worksheets.Add
    Dress ws.Range("A1:C12"), font, size, isBold, False, 0
    Set p = wb.PivotCaches.Create(1, data.Range("A1:B3")).CreatePivotTable(ws.Range("A3"), "P")
    p.PivotFields("Region of sale").Orientation = 1
    p.AddDataField p.PivotFields("Amount")
    For i = 1 To 3
        r = r & "|" & ws.Columns(i).ColumnWidth
    Next
    Application.DisplayAlerts = False
    ws.Delete
    data.Delete
    Application.DisplayAlerts = True
    Pivoted = r
End Function

Public Function Normal(ByVal font As String, ByVal size As Double) As String
    Dim ws As Object, r As String, name As String, tall As Double
    Set ws = Fresh()
    name = ActiveWorkbook.Styles("Normal").Font.Name
    tall = ActiveWorkbook.Styles("Normal").Font.Size
    ActiveWorkbook.Styles("Normal").Font.Name = font
    ActiveWorkbook.Styles("Normal").Font.Size = size
    r = Said(ws.StandardWidth) & "|" & Said(ws.StandardHeight) & "|" & Said(ws.Columns(3).Width) & "|" & Said(ws.Rows(3).RowHeight)
    ws.Range("C1").Value = "Some text in the Normal font"
    ws.Columns(3).AutoFit
    r = r & "|" & Said(ws.Columns(3).ColumnWidth)
    ws.Columns(3).UseStandardWidth = True
    ws.Range("C1").Clear
    ActiveWorkbook.Styles("Normal").Font.Name = name
    ActiveWorkbook.Styles("Normal").Font.Size = tall
    Normal = r
End Function
'''
WORDS = ('Harbour', 'ledger', 'Quince', 'volume', 'Sixty', 'tariff', 'WAREHOUSE', 'mill', 'jetty', 'Oxbow', 'rate-card', 'pre-tax', 'illicit', 'Wimble', 'fifty', 'Ampere', 'naive', 'caf\xe9', '1,250.75', '(net)', 'A&B', 'zig-zag', 'lily', 'MMXX', '9%')
SIZES = (8.0, 9.0, 10.0, 10.5, 11.0, 11.5, 12.0, 14.0, 16.0, 18.0, 20.0, 24.0, 36.0)
STYLES = ((False, False), (True, False), (False, True), (True, True))


def items(font: str) -> list[list[object]]:
    """Every question of this font's round, as a procedure and its arguments. Different for each font."""
    chooser = random.Random(font)
    found: list[list[object]] = []

    def text(most: int = 7) -> str:
        return ' '.join(chooser.choices(WORDS, k=chooser.randint(1, most)))

    for size in SIZES:
        for bold, italic in STYLES:
            for _ in range(5):
                found.append(['Fit', text(4), font, size, bold, italic, chooser.choice((0, 0, 1, 3))])
            found.append(['Plain', font, size, bold, italic])
            widths = ','.join(str(chooser.randint(1, 45) + chooser.choice((0, 0.5))) for _ in range(9))
            found.append(['Wrapped', text(), font, size, bold, italic, chooser.choice((0, 0, 1)), widths])
            found.append(['Wrapped', f'{text(3)}~{text(4)}', font, size, bold, italic, 0, widths])
            pixels = sorted(chooser.sample(range(24, 500), 10))
            found.append(['Lined', text(), font, size, bold, italic, ','.join(str(pixel * .75) for pixel in pixels)])
            for value, shape in ((1234567.891, 'General'), (0.000123456, 'General'), (98765.4321, '#,##0.00'), (45123.0, 'yyyy-mm-dd'), (12.5, '0.00%')):
                found.append(['Shown', value, shape, font, size, bold, italic, ','.join(str(one) for one in (0.5, 1, 1.5, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 14, 18))])
    for size in (9.0, 11.0, 14.0):
        for width in (8.0, 14.0, 30.0):
            found.append(['Justified', ' '.join(chooser.choices(WORDS, k=14)), font, size, width])
        for bold in (False, True):
            found.append(['Pivoted', font, size, bold])
        found.append(['Normal', font, size])
    chooser.shuffle(found)
    return found


def _stem(font: str) -> Path:
    return FIXTURES / f"font_round_{font.lower().replace(' ', '_')}"


def predict(font: str) -> None:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    cases = []
    refused = 0
    for item in items(font):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(SOURCE, name='Round')
        try:
            answer: str | None = str(app.run(f'Round.{item[0]}', *item[1:]))
            reason = None
        except VBAUnsupportedError as refusal:
            answer, reason = None, str(refusal)
            refused += 1
        cases.append({'call': item, 'model': answer, 'refused': reason})
    path = Path(f'{_stem(font)}_predictions.json')
    path.write_text(json.dumps({'font': font, 'cases': cases}, indent=0), encoding='utf-8', newline='\n')
    print(len(cases), 'questions,', refused, 'refused')


def _check(excel: object, stage: str) -> None:
    result = excel.run_vba('Public Function H() As String\n    H = CStr(ActiveSheet.StandardHeight)\nEnd Function\n', 'H', (), timeout=30.0, line_numbers=False)  # type: ignore[attr-defined]
    if not result.ok or str(result.value) != '15':
        print(f'Excel is not at 96 DPI {stage}: StandardHeight {result.value}; what is measured is kept')
        sys.exit(3)


def measure(font: str) -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases = json.loads(Path(f'{_stem(font)}_predictions.json').read_text('utf-8'))['cases']
    path = Path(f'{_stem(font)}.json')
    answers: list[str] = json.loads(path.read_text('utf-8'))['answers'] if path.exists() else []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        _check(excel, 'at the start')
        for number in range(len(answers), len(cases)):
            call = cases[number]['call']
            result = excel.run_vba(SOURCE, str(call[0]), tuple(call[1:]), timeout=120.0, line_numbers=False)
            answers.append(str(result.value) if result.ok and not result.dialogs else f'ERROR {result}')
            if number % 25 == 24:
                _check(excel, f'at question {number}')
                path.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': '15', 'answers': answers}, indent=0), encoding='utf-8', newline='\n')
        _check(excel, 'at the end')
    path.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': '15', 'answers': answers}, indent=0), encoding='utf-8', newline='\n')
    score(font)


def score(font: str) -> None:
    cases = json.loads(Path(f'{_stem(font)}_predictions.json').read_text('utf-8'))['cases']
    answers = json.loads(Path(f'{_stem(font)}.json').read_text('utf-8'))['answers']
    tally: dict[str, list[int]] = {}
    wrong: list[object] = []
    for case, answer in zip(cases, answers):
        row = tally.setdefault(str(case['call'][0]), [0, 0, 0])
        if case['model'] is None:
            row[2] += 1
        elif case['model'] == answer:
            row[0] += 1
        else:
            row[1] += 1
            wrong.append((case['call'], case['model'], answer))
    print(f'{len(answers)} of {len(cases)} asked; right, wrong, refused by procedure:', tally)
    for one in wrong[:int(sys.argv[3]) if len(sys.argv) > 3 else 12]:
        print(str(one)[:400])
    reasons = sorted({str(case['refused'])[:90] for case in cases if case['refused']})
    print(reasons)


if __name__ == '__main__':
    {'predict': predict, 'measure': measure, 'score': score}[sys.argv[1]](sys.argv[2])
