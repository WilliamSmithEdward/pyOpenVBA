"""Ask Excel where text breaks in a shape, at the widths a predict script wrote down, and score the predictions: `measure_shape_lines.py shape_lines2` reads
tests/fixtures/shape_lines2_predictions.json and writes tests/fixtures/shape_lines2.json.

Each call reads StandardHeight at both ends, so a move away from 96 DPI during a run stops it.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures'
SOURCE = r'''Public Function Go(ByVal codes As String, ByVal fontName As String, ByVal size As Double, ByVal isBold As Boolean, ByVal isItalic As Boolean, ByVal kind As String, ByVal leftMargin As Double, ByVal rightMargin As Double, ByVal wraps As Boolean, ByVal widths As String) As String
    Dim ws As Object, s As Object, t As Object, r As String, one As Variant, text As String, i As Long, before As String
    Set ws = ActiveSheet
    before = ws.StandardHeight
    For Each one In Split(codes, ",")
        text = text & ChrW(CLng(one))
    Next
    If kind = "text box" Then
        Set s = ws.Shapes.AddTextbox(1, 10, 10, 100, 400)
    Else
        Set s = ws.Shapes.AddShape(1, 10, 10, 100, 400)
    End If
    s.TextFrame2.TextRange.Text = text
    With s.TextFrame2.TextRange.Font
        .Name = fontName
        .Size = size
        .Bold = IIf(isBold, msoTrue, msoFalse)
        .Italic = IIf(isItalic, msoTrue, msoFalse)
    End With
    s.TextFrame2.MarginLeft = leftMargin
    s.TextFrame2.MarginRight = rightMargin
    If Not wraps Then s.TextFrame2.WordWrap = msoFalse
    For Each one In Split(widths, ",")
        s.Width = CDbl(one)
        Set t = s.TextFrame2.TextRange.Lines
        r = r & "|" & s.Width & ":" & TypeName(t) & ":" & TypeName(t.Count) & ":" & t.Count & ":"
        For i = 1 To t.Count
            r = r & s.TextFrame2.TextRange.Lines(i).Start & "+" & s.TextFrame2.TextRange.Lines(i).Length & ","
        Next
    Next
    s.Delete
    Go = before & "~" & ws.StandardHeight & r
End Function
'''


def main() -> None:
    stem = sys.argv[1]
    predictions = json.loads((FIXTURES / f'{stem}_predictions.json').read_text('utf-8'))
    measured = []
    right = total = rival_right = rival_total = 0
    missed: list[object] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for number, case in enumerate(predictions['cases']):
            codes = ','.join(str(ord(character)) for character in case['text'])
            widths = ','.join(str(pixel * 0.75) for pixel in case['pixels'])
            margins = case.get('margins', [7.2, 7.2])
            kind, wraps = case.get('kind', 'rectangle'), case.get('wrap', True)
            result = excel.run_vba(SOURCE, 'Go', (codes, case['font'], case['size'], case['bold'], case['italic'], kind, margins[0], margins[1], wraps, widths), timeout=120.0, line_numbers=False)
            assert result.ok and not result.dialogs, (case['text'], result)
            heights, *found = str(result.value).split('|')
            if heights != '15~15':
                print(f'Excel is not at 96 DPI, StandardHeight {heights}, at case {number}: nothing kept from this run')
                sys.exit(1)
            answers = []
            rivals = case.get('rival') or [None] * len(case['pixels'])
            for pixel, spans, rival, one in zip(case['pixels'], case['lines'], rivals, found, strict=True):
                width, kind_name, count_kind, count, parts = one.split(':')
                lines = [[int(part.split('+')[0]) - 1, int(part.split('+')[1])] for part in parts.split(',') if part]
                answers.append({'pixel': pixel, 'width': width, 'type': kind_name, 'count_type': count_kind, 'count': int(count), 'lines': lines})
                total += 1
                if rival is not None:
                    rival_total += 1
                    rival_right += lines == rival
                if lines == spans:
                    right += 1
                else:
                    missed.append((case['text'], case['font'], case['size'], case['bold'], case['italic'], kind, margins, wraps, pixel, width, spans, lines))
            measured.append({key: case[key] for key in ('text', 'font', 'size', 'bold', 'italic')} | {'kind': kind, 'margins': margins, 'wrap': wraps, 'answers': answers})
    (FIXTURES / f'{stem}.json').write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': '15', 'source': SOURCE, 'cases': measured}, indent=0), encoding='utf-8', newline='\n')
    print(f'{right} of {total}; where the rounding decides, the unrounded rival {rival_right} of {rival_total}')
    for one in missed[:40]:
        print(one)


if __name__ == '__main__':
    main()
