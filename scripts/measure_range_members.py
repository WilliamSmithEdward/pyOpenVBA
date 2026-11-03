"""Members of a range and a sheet still missing: Show, Justify, NavigateArrow, SortSpecial, Run, Phonetic and the rest. What each
answers or raises, and the cells it leaves.

Each case makes a workbook with text in A1:A4, numbers in B1:B4, a formula in D1 reading B1:B2, one in E1 reading D1, and a
sheet Other with a number in A1 that F1 reads. It runs the case's statements, which set r to what they want recorded. The
answer is "ok " or "Err " and the error's number, then r, then READS, each as its value or "E" and an error number.

    python scripts/measure_range_members.py

writes tests/fixtures/range_members.json, which tests/test_excel_range_members.py replays. It leaves out what would open a
window or speak on this machine: Speak, FunctionWizard, DialogBox, CheckSpelling, PrintPreview and ShowDataForm.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_members.json'
CASES = [
    # Bringing it into view.
    'r = d.Range("A1").Show', 'd.Range("A1").Show', 'd.Range("Z500").Show: r = ActiveCell.Address', 'd.Range("Z500").Select: d.Range("A1").Show: r = ActiveCell.Address', 'r = d.Range("A1:B2").Show', 'r = o.Range("A1").Show', 'o.Range("A1").Show: r = ActiveSheet.Name',
    # Filling a column with text.
    'd.Range("A1:A4").Justify: r = d.Range("A1").Value & "|" & d.Range("A2").Value & "|" & d.Range("A3").Value & "|" & d.Range("A4").Value', 'r = d.Range("A1:A4").Justify', 'r = TypeName(d.Range("A1:A4").Justify)',
    'd.Range("A1").Value = "the quick brown fox jumps over the lazy dog again and again": d.Range("A1").Justify: r = d.Range("A1").Value & "|" & d.Range("A2").Value & "|" & d.Range("A3").Value & "|" & d.Range("A4").Value & "|" & d.Range("A5").Value & "|" & d.Range("A6").Value',
    'd.Range("A1").Value = "the quick brown fox jumps over the lazy dog again and again": d.Range("A1:A2").Justify: r = d.Range("A1").Value & "|" & d.Range("A2").Value & "|" & d.Range("A3").Value', 'd.Range("A1:B4").Justify: r = d.Range("A1").Value & "|" & d.Range("A2").Value & "|" & d.Range("B1").Value', 'd.Range("B1:B4").Justify: r = d.Range("B1").Value & "|" & d.Range("B2").Value', 'd.Range("D1").Justify: r = d.Range("D1").Formula',
    # Tracer arrows.
    'r = d.Range("D1").NavigateArrow', 'r = TypeName(d.Range("D1").NavigateArrow(True))', 'r = d.Range("D1").NavigateArrow(True).Address', 'd.Range("D1").ShowPrecedents: r = d.Range("D1").NavigateArrow(True).Address', 'd.Range("D1").ShowPrecedents: r = d.Range("D1").NavigateArrow(True, 1).Address', 'd.Range("D1").ShowPrecedents: r = d.Range("D1").NavigateArrow(True, 1, 1).Address', 'd.Range("D1").ShowPrecedents: r = d.Range("D1").NavigateArrow(True, 1, 2).Address', 'd.Range("D1").ShowPrecedents: r = d.Range("D1").NavigateArrow(True, 2, 1).Address',
    'd.Range("D1").ShowPrecedents: d.Range("D1").NavigateArrow True: r = ActiveCell.Address', 'd.Range("D1").ShowDependents: r = d.Range("D1").NavigateArrow(False).Address', 'd.Range("D1").ShowDependents: r = d.Range("D1").NavigateArrow(False, 1, 1).Address', 'd.Range("D1").ShowPrecedents: r = d.Range("D1").NavigateArrow(False).Address', 'd.Range("D1").ShowDependents: r = d.Range("D1").NavigateArrow(True).Address', 'd.Range("F1").ShowPrecedents: r = d.Range("F1").NavigateArrow(True).Address(External:=True)', 'd.Range("A1").ShowPrecedents: r = d.Range("A1").NavigateArrow(True).Address', 'd.ClearArrows: r = d.Range("D1").NavigateArrow(True).Address', 'd.Range("D1").ShowPrecedents: d.ClearArrows: r = d.Range("D1").NavigateArrow(True).Address', 'd.Range("D1").ShowPrecedents True: r = d.Range("D1").ShowPrecedents(True)', 'r = d.Range("D1").ShowPrecedents', 'r = d.Range("D1").ShowDependents', 'r = d.Range("D1").ShowErrors',
    # Sorting by phonetic order.
    'd.Range("A1:A4").SortSpecial: r = d.Range("A1").Value', 'r = d.Range("A1:A4").SortSpecial(xlPinYin)', 'd.Range("A1:A4").SortSpecial xlPinYin, d.Range("A1"), xlDescending: r = d.Range("A1").Value & d.Range("A4").Value', 'd.Range("A1:A4").SortSpecial xlStroke, d.Range("A1"), xlAscending: r = d.Range("A1").Value & d.Range("A4").Value', 'd.Range("A1:B4").SortSpecial xlPinYin, d.Range("B1"), xlDescending: r = d.Range("B1").Value & d.Range("B4").Value & d.Range("A1").Value', 'd.Range("A1:A4").SortSpecial SortMethod:=9: r = d.Range("A1").Value',
    # A macro by cell.
    'r = d.Range("A1").Run', 'd.Range("A1").Run', 'r = d.Range("A1:A2").Run', 'd.Range("A1").Value = "=SUM(1,2)": r = d.Range("A1").Run',
    # Phonetics.
    'r = TypeName(d.Range("A1").Phonetic)', 'r = TypeName(d.Range("A1").Phonetics)', 'r = d.Range("A1").Phonetic.Text', 'r = d.Range("A1").Phonetics.Count', 'd.Range("A1").SetPhonetic: r = d.Range("A1").Phonetics.Count', 'r = d.Range("A1").Phonetic.Visible', 'd.Range("A1").Phonetic.Visible = True: r = d.Range("A1").Phonetic.Visible',
    # Other things it can do.
    'r = TypeName(d.Range("A1").SoundNote)', 'r = TypeName(d.Range("A1").XPath)', 'r = d.Range("A1").XPath.Value', 'r = d.Range("A1").CellControl Is Nothing', 'r = d.Range("A1").ResetContents',
    # A sheet's paste.
    'd.Range("B1:B2").Copy: d.Activate: d.Range("G1").Select: d.PasteSpecial: r = d.Range("G1").Value & d.Range("G2").Value', 'd.Range("B1:B2").Copy: d.Range("G1").Select: d.PasteSpecial "Text": r = d.Range("G1").Value & TypeName(d.Range("G1").Value)', 'd.Range("B1:B2").Copy: d.Range("G1").Select: r = d.PasteSpecial(Format:="Text")', 'd.Range("B1:B2").Copy: d.Range("G1").Select: d.PasteSpecial Format:="Nothing"', 'd.PasteSpecial', 'Application.CutCopyMode = False: d.PasteSpecial',
    'r = TypeName(d.CircularReference)',
]
# Which arrow leads where, and which links an arrow to another sheet has.
for formula in ('=B1+B2+B3', '=B3+B1', '=SUM(B1:B2)+B4', '=B1+B1', '=B1+D1', '=5', '=E9'):
    CASES += [f'd.Range("H1").Formula = "{formula}": d.Range("H1").ShowPrecedents: r = d.Range("H1").NavigateArrow(True, {arrow}, 1).Address' for arrow in (1, 2, 3, 4)]
for formula in ('=Other!A1', '=Other!A1+Other!B1', '=B1+Other!A1', '=Other!A1+B1+B2'):
    CASES += [f'd.Range("H1").Formula = "{formula}": d.Range("H1").ShowPrecedents: r = d.Range("H1").NavigateArrow(True, {arrow}, {link}).Address(External:=False)' for arrow, link in ((1, 1), (1, 2), (2, 1), (2, 2), (3, 1))]
    CASES += [f'd.Range("H1").Formula = "{formula}": d.Range("H1").ShowPrecedents: r = d.Range("H1").NavigateArrow(True, 1, 1).Parent.Name']
SPREAD = 'd.Range("G3").Formula = "=D1+1": d.Range("C7").Formula = "=D1": o.Range("B1").Formula = "=Data!D1": d.Range("D1").ShowDependents: '
CASES += [f'{SPREAD}r = d.Range("D1").NavigateArrow(False, {arrow}, {link}).Address' for arrow, link in ((1, 1), (2, 1), (3, 1), (4, 1), (5, 1), (1, 2), (2, 2), (3, 2), (4, 2))]
CASES += [f'{SPREAD}d.Range("D1").ShowDependents: r = d.Range("D1").NavigateArrow(False, {arrow}, 1).Address' for arrow in (1, 2, 3, 4)]
CASES += [f'd.Range("D1").ShowPrecedents: d.Range("D1").ShowPrecedents: r = d.Range("D1").NavigateArrow(True, {arrow}, 1).Address' for arrow in (1, 2)]
CASES += ['d.Range("E1").ShowPrecedents: d.Range("E1").ShowPrecedents: r = d.Range("D1").NavigateArrow(True, 1, 1).Address', 'd.Range("E1").ShowPrecedents: d.Range("E1").ShowPrecedents: r = d.Range("E1").NavigateArrow(True, 2, 1).Address', 'd.Range("E1").ShowPrecedents: d.Range("E1").ShowPrecedents True: r = d.Range("D1").NavigateArrow(True, 1, 1).Address', 'd.Range("D1").ShowPrecedents: d.Range("D1").ShowPrecedents True: r = d.Range("D1").NavigateArrow(True, 1, 1).Address', 'd.Range("D1:E1").ShowPrecedents: r = d.Range("E1").NavigateArrow(True, 1, 1).Address', 'd.Range("D1").ShowPrecedents: d.Range("B1").Value = 9: r = d.Range("D1").NavigateArrow(True, 1, 1).Address', 'd.Range("D1").ShowPrecedents: d.Range("H1").Value = 9: r = d.Range("D1").NavigateArrow(True, 1, 1).Address', 'd.Range("D1").ShowPrecedents: d.Rows(1).Insert: r = d.Range("D2").NavigateArrow(True, 1, 1).Address', 'd.Range("D1").ShowPrecedents: r = d.Range("D1:E1").NavigateArrow(True, 1, 1).Address', 'd.Range("D1").ShowPrecedents: o.Activate: r = d.Range("D1").NavigateArrow(True, 1, 1).Address']
# Where Justify breaks text.
TEXTS = ('aaaaaaaa bbbbbbbb cccccccc', 'iiiiiiiiiiiiii wwwwwwww iiiii', 'a b c d e f g h i j k l m n o p q r s t', 'oneveryveryveryverylongwordindeed and more', 'Mixed Case Words With Numbers 12345 678', 'trailing spaces   and   inner   runs  ', '=not a formula because of text')
for width in ('', 'd.Columns(1).ColumnWidth = 20: ', 'd.Columns(1).ColumnWidth = 3: '):
    CASES += [f'd.Range("A1").NumberFormat = "@": d.Range("A1").Value = "{value}": {width}d.Range("A1").Justify: r = d.Range("A1").Text & "|" & d.Range("A2").Text & "|" & d.Range("A3").Text & "|" & d.Range("A4").Text & "|" & d.Range("A5").Text & "|" & d.Range("A6").Text & "|" & d.Range("A7").Text & "|" & d.Range("A8").Text' for value in TEXTS]
READS = ['d.Range("A1").Value', 'd.Range("A2").Value', 'd.Range("A3").Value', 'd.Range("A4").Value', 'd.Range("B1").Value', 'd.Range("D1").Formula', 'ActiveSheet.Name', 'ActiveCell.Address', 'd.UsedRange.Address', 'Selection.Address']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, d As Object, o As Object, s As String, v As String, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set d = wb.Worksheets(1): d.Name = "Data"', '    Set o = wb.Worksheets.Add(After:=d): o.Name = "Other": d.Activate',
             '    d.Range("A1:A4").Value = Application.Transpose(Array("pear", "apple", "fig", "date"))', '    d.Range("B1:B4").Value = Application.Transpose(Array(3, 1, 4, 2))', '    d.Range("D1").Formula = "=SUM(B1:B2)"', '    d.Range("E1").Formula = "=D1"', '    o.Range("A1").Value = 5', '    d.Range("F1").Formula = "=Other!A1"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    Application.CutCopyMode = False', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 30):
            last = min(first + 29, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
