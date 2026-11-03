"""TableStyle.TableStyleElements on the built-in styles: how many there are, how an element is found, what it reads, and what
a macro may change on a style it did not make.

    python scripts/measure_table_style_elements.py

writes tests/fixtures/table_style_elements.json, which tests/test_excel_table_style_elements.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/table_style_elements.json'
S = 'wb.TableStyles("TableStyleMedium2")'
E = f'{S}.TableStyleElements'
READS = ['HasFormat', 'Font.Bold', 'Font.Italic', 'Font.Color', 'Font.ThemeColor', 'Font.TintAndShade', 'Font.ColorIndex', 'Interior.Color', 'Interior.ThemeColor', 'Interior.TintAndShade', 'Interior.Pattern',
         'Borders(xlEdgeTop).LineStyle', 'Borders(xlEdgeTop).Weight', 'Borders(xlEdgeTop).Color', 'Borders(xlEdgeBottom).LineStyle', 'Borders(xlEdgeLeft).LineStyle', 'Borders(xlEdgeRight).LineStyle', 'Borders(xlInsideHorizontal).LineStyle', 'Borders(xlInsideVertical).LineStyle', 'StripeSize']
CASES = [
    f'r = TypeName({E})', f'r = {E}.Count', f'r = TypeName({E}(1))', f'r = TypeName({E}(0))', f'r = TypeName({E}(xlWholeTable))', f'r = TypeName({E}(13))', f'r = TypeName({E}(-1))', f'r = TypeName({E}(43))', f'r = TypeName({E}(44))', f'r = TypeName({E}("x"))', f'r = TypeName({E}.Item(1))', f'r = TypeName({E}.Parent)', f'r = TypeName({E}.Application)', f'r = {E}.Creator', f'r = {E}(1).Creator', f'r = TypeName({E}(1).Parent)', f'r = TypeName({E}(1).Application)', f'For Each x In {E}: i = i + 1: Next: r = i', f'r = TypeName({E}(1).Font)', f'r = TypeName({E}(1).Interior)', f'r = TypeName({E}(1).Borders)', f'r = TypeName({E}(1).Borders(xlEdgeTop))',
    *[f'r = {S}.TableStyleElements({element}).{read}' for element in (0, 1, 2, 5, 6, 7, 9) for read in READS],
    *[f'r = wb.TableStyles("{style}").TableStyleElements({element}).HasFormat' for style in ('TableStyleLight1', 'TableStyleLight21', 'TableStyleDark11', 'PivotStyleLight16', 'SlicerStyleLight1') for element in (0, 1, 3, 12, 20, 41)],
    f'r = wb.TableStyles("TableStyleLight1").TableStyleElements.Count', 'r = wb.TableStyles("PivotStyleLight16").TableStyleElements.Count', 'r = wb.TableStyles("SlicerStyleLight1").TableStyleElements.Count', 'r = wb.TableStyles("TimeSlicerStyleLight1").TableStyleElements.Count',
    # Changing a built-in style.
    f'{E}(1).Font.Bold = True', f'{E}(1).Interior.Color = 255', f'{E}(1).StripeSize = 2', f'{E}(5).StripeSize = 2', f'{E}(1).Clear', f'r = {E}(1).Clear', f'{E}(1).Borders(xlEdgeTop).LineStyle = xlContinuous',
]

MODULE = r'''Option Explicit

Private Function Shown(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        Shown = CStr(v)
    ElseIf IsArray(v) Then
        Shown = "array"
    Else
        Shown = TypeName(v) & ":" & CStr(v)
    End If
End Function

Public Function Tried(ByVal first As Long, ByVal last As Long) As String
    Dim wb As Object, r As Variant, s As String, which As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    For which = first To last
        r = Empty
        On Error GoTo Failed
        Run which, wb, r
        s = s & "@@ok " & Shown(r)
Answered:
        On Error GoTo 0
    Next
    wb.Close False
    Tried = Mid(s, 3)
    Exit Function
Failed:
    s = s & "@@Err " & Err.Number & " " & Err.Description
    Resume Answered
End Function
'''


def _source() -> str:
    lines = ['', 'Private Sub Run(ByVal which As Long, wb As Object, r As Variant)', '    Dim x As Object, i As Long', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '']
    return MODULE + '\n'.join(lines)


def main() -> None:
    source = _source()
    answers: list[str] = []
    index = 0
    while index < len(CASES):
        with ExcelSession(HarnessConfig(exclusive=False)) as excel:
            excel.new_document()
            while index < len(CASES):
                last = min(index + 24, len(CASES) - 1)
                result = excel.run_vba(source, 'Tried', (index, last), timeout=60.0)
                if not result.ok or result.dialogs:
                    text = '; '.join(getattr(dialog, 'message', '') for dialog in result.dialogs)
                    answers.append(f'breaks: {result.outcome} {text}')
                    print(index, CASES[index], '=>', answers[-1], flush=True)
                    index += 1
                    break
                got = str(result.value).split('@@')
                answers.extend(got)
                for offset, answer in enumerate(got):
                    print(index + offset, CASES[index + offset], '=>', answer, flush=True)
                index = last + 1
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'reads': READS,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
