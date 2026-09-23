"""Copy, Cut, Paste and PasteSpecial through the clipboard, as Excel's object model answers them.

Every layout sits on a sheet of its own in one new workbook: a setup that
fills a few cells and copies or cuts them, pasting in one of the ways a
macro can, then a dump of A1:F6 -- each cell's formula, value type,
number format and bold -- with the used range, Application.CutCopyMode
and whether column D is as wide as column A, and what the last call
returned or the error it raised. The layouts ask what each paste type takes along, what an
operation does to the cells under it, what skipping blanks and
transposing do, how a paste is sized, how long a copy lasts, and what a
cut moves and which formulas follow it.

    python scripts/measure_paste.py

writes tests/fixtures/paste.json, which tests/test_excel_paste.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "paste.json"

HELPER = '''Private Function Show(v As Variant) As String
    If IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    ElseIf IsObject(v) Then
        If v Is Nothing Then Show = "Nothing" Else Show = "Object"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, c As Object
    For Each c In ws.Range("A1:F6").Cells
        out = out & c.Formula & "~" & TypeName(c.Value) & "~" & c.NumberFormat & "~" & c.Font.Bold & "<>"
    Next
    ' Widths depend on the display's DPI, so only whether D is as wide as A is kept.
    Dump = out & ws.UsedRange.Address & "<>" & Application.CutCopyMode & "<>" & _
        (ws.Columns("D").ColumnWidth = ws.Columns("A").ColumnWidth)
End Function
'''

#: A small block to copy: a bold number, a formula under it in 0.00, text beside them, and a blank.
SOURCE = ('ws.Range("A1").Value = 1\nws.Range("A1").Font.Bold = True\nws.Range("A2").Formula = "=A1*2"\n'
          'ws.Range("A2").NumberFormat = "0.00"\nws.Range("B1").Value = "x"\n')
#: What stands where the block is pasted, to see what each paste keeps of it.
UNDER = ('ws.Range("D1").Value = 10\nws.Range("D2").Value = 20\nws.Range("E1").Value = "old"\n'
         'ws.Range("E2").Value = 30\nws.Range("E2").NumberFormat = "$#,##0"\nws.Range("D1").Font.Bold = True\n')

#: name -> setup; each works on ws, the layout's own sheet and the active one, and keeps a result in v.
LAYOUTS: dict[str, str] = {
    "paste_at_selection": SOURCE + 'ws.Range("A1:B2").Copy\nws.Range("D1").Select\nv = ws.Paste',
    "paste_destination": SOURCE + 'ws.Range("A1:B2").Copy\nv = ws.Paste(Destination:=ws.Range("D3"))',
    "paste_twice": SOURCE + 'ws.Range("A1:B2").Copy\nws.Paste Destination:=ws.Range("D1")\n'
                            'v = ws.Paste(Destination:=ws.Range("D4"))',
    "copy_returns": SOURCE + 'v = ws.Range("A1:B2").Copy',
    "nothing_copied": 'Application.CutCopyMode = False\nws.Range("A1").Value = 1\nv = ws.Paste',
    "all": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\nv = ws.Range("D1").PasteSpecial(xlPasteAll)',
    "values": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\nv = ws.Range("D1").PasteSpecial(xlPasteValues)',
    "formulas": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\nv = ws.Range("D1").PasteSpecial(xlPasteFormulas)',
    "formats": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\nv = ws.Range("D1").PasteSpecial(xlPasteFormats)',
    "values_and_number_formats": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\n'
                                 'v = ws.Range("D1").PasteSpecial(xlPasteValuesAndNumberFormats)',
    "formulas_and_number_formats": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\n'
                                   'v = ws.Range("D1").PasteSpecial(xlPasteFormulasAndNumberFormats)',
    "all_except_borders": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\n'
                          'v = ws.Range("D1").PasteSpecial(xlPasteAllExceptBorders)',
    "column_widths": SOURCE + 'ws.Columns("A").ColumnWidth = 20\nws.Range("A1:B2").Copy\n'
                     'v = ws.Range("D1").PasteSpecial(xlPasteColumnWidths)',
    "default_paste": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\nv = ws.Range("D1").PasteSpecial()',
    "add": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\nv = ws.Range("D1").PasteSpecial(xlPasteValues, xlPasteSpecialOperationAdd)',
    "subtract": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\n'
                'v = ws.Range("D1").PasteSpecial(xlPasteValues, xlPasteSpecialOperationSubtract)',
    "multiply_all": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\n'
                    'v = ws.Range("D1").PasteSpecial(xlPasteAll, xlPasteSpecialOperationMultiply)',
    "divide_formulas": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\n'
                       'v = ws.Range("D1").PasteSpecial(xlPasteFormulas, xlPasteSpecialOperationDivide)',
    "operation_on_formula": SOURCE + 'ws.Range("D1").Formula = "=5+5"\nws.Range("D2").Formula = "=A1"\n'
                            'ws.Range("A1:A2").Copy\n'
                            'v = ws.Range("D1").PasteSpecial(xlPasteFormulas, xlPasteSpecialOperationAdd)',
    "skip_blanks": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\n'
                   'v = ws.Range("D1").PasteSpecial(xlPasteAll, SkipBlanks:=True)',
    "transpose": SOURCE + 'ws.Range("A3").Formula = "=B1&A1"\nws.Range("A1:B3").Copy\n'
                 'v = ws.Range("D1").PasteSpecial(xlPasteAll, Transpose:=True)',
    "tiled": SOURCE + 'ws.Range("A1:A2").Copy\nv = ws.Range("D1:D4").PasteSpecial(xlPasteAll)',
    "tiled_across": SOURCE + 'ws.Range("A1:A2").Copy\nv = ws.Range("D1:F2").PasteSpecial(xlPasteFormulas)',
    "odd_size": SOURCE + 'ws.Range("A1:A2").Copy\nv = ws.Range("D1:D3").PasteSpecial(xlPasteAll)',
    "smaller_target": SOURCE + 'ws.Range("A1:B2").Copy\nv = ws.Range("D1:D1").PasteSpecial(xlPasteAll)',
    "one_cell_many": SOURCE + 'ws.Range("A1").Copy\nv = ws.Range("D1:E3").PasteSpecial(xlPasteValues)',
    "source_changed": SOURCE + 'ws.Range("A1:A2").Copy\nws.Range("A1").Value = 7\n'
                      'v = ws.Range("D1").PasteSpecial(xlPasteValues)',
    "other_cell_written": SOURCE + 'ws.Range("A1:A2").Copy\nws.Range("F6").Value = 1\n'
                          'v = ws.Range("D1").PasteSpecial(xlPasteValues)',
    "mode_after_write": SOURCE + 'ws.Range("A1:A2").Copy\nws.Range("F6").Value = 1\nv = Application.CutCopyMode',
    "mode_cleared": SOURCE + 'ws.Range("A1:A2").Copy\nApplication.CutCopyMode = False\n'
                    'v = ws.Range("D1").PasteSpecial(xlPasteValues)',
    "cut_mode": SOURCE + 'ws.Range("A1:A2").Cut\nv = Application.CutCopyMode',
    "cut_paste": SOURCE + 'ws.Range("C1").Formula = "=A1*10"\nws.Range("C2").Formula = "=SUM(A1:A2)"\n'
                 'ws.Range("A1:B2").Cut\nws.Range("D3").Select\nv = ws.Paste',
    "cut_destination": SOURCE + 'ws.Range("C1").Formula = "=A1*10"\nws.Range("C2").Formula = "=SUM(A1:A2)"\n'
                       'v = ws.Range("A1:B2").Cut(ws.Range("D3"))',
    "cut_then_special": SOURCE + 'ws.Range("A1:B2").Cut\nv = ws.Range("D1").PasteSpecial(xlPasteValues)',
    "cut_formula_out": SOURCE + 'ws.Range("A2").Cut ws.Range("E5")',
    "cut_overlapping": SOURCE + 'v = ws.Range("A1:B2").Cut(ws.Range("B2"))',
    "cut_whole_column": SOURCE + 'ws.Range("C1").Formula = "=A1*10"\nv = ws.Columns("A").Cut(ws.Columns("E"))',
    "cut_partial_range": SOURCE + 'ws.Range("A3").Value = 5\nws.Range("C1").Formula = "=SUM(A1:A5)"\n'
                         'ws.Range("C2").Formula = "=SUM(A1:B1)"\nv = ws.Range("A1").Cut(ws.Range("D1"))',
    "cut_whole_range": SOURCE + 'ws.Range("A3").Value = 5\nws.Range("C1").Formula = "=SUM(A1:A3)"\n'
                       'ws.Range("C2").Formula = "=SUM($A$1:$A$3)"\nv = ws.Range("A1:A3").Cut(ws.Range("D1"))',
    "cut_onto_referenced": SOURCE + 'ws.Range("E1").Value = 9\nws.Range("C1").Formula = "=E1"\n'
                           'ws.Range("C2").Formula = "=SUM(E1:E2)"\nv = ws.Range("A1").Cut(ws.Range("E1"))',
    "cut_to_other_sheet": SOURCE + 'ws.Range("C1").Formula = "=A1+1"\n'
                          'ws.Range("A1").Cut ws.Parent.Worksheets(1).Range("F10")\n'
                          'v = ws.Parent.Worksheets(1).Range("F10").Formula & "," & ws.Parent.Worksheets(1).Range("F10").Value',
    "cut_name": SOURCE + 'ws.Names.Add Name:="Moved", RefersTo:="=$A$1:$A$2"\n'
                'ws.Range("A1:A2").Cut ws.Range("D1")\nv = ws.Names("Moved").RefersTo',
    "transpose_outside": 'ws.Range("C5").Value = 3\nws.Range("A1").Formula = "=C5"\nws.Range("B1").Formula = "=$C$5"\n'
                         'ws.Range("A2").Formula = "=A1+B1"\nws.Range("B2").Formula = "=C$5+$C6"\nws.Range("A1:B2").Copy\n'
                         'v = ws.Range("D1").PasteSpecial(xlPasteFormulas, Transpose:=True)',
    "paste_at_big_selection": SOURCE + 'ws.Range("A1:A2").Copy\nws.Range("D1:D4").Select\nv = ws.Paste',
    "paste_blank_under": SOURCE + UNDER + 'ws.Range("A1:B2").Copy\nws.Range("D1").Select\nv = ws.Paste',
    "values_of_text_and_errors": 'ws.Range("A1").Formula = "=1/0"\nws.Range("A2").Formula = "=""t"""\n'
                                 'ws.Range("A3").Value = "\'5"\nws.Range("A1:A3").Copy\n'
                                 'v = ws.Range("D1").PasteSpecial(xlPasteValues)',
    "add_to_blank_and_text": 'ws.Range("A1").Value = 1\nws.Range("A2").Value = "t"\nws.Range("A3").Value = 2\n'
                             'ws.Range("D2").Value = 5\nws.Range("D3").Value = "u"\nws.Range("A1:A3").Copy\n'
                             'v = ws.Range("D1").PasteSpecial(xlPasteValues, xlPasteSpecialOperationAdd)',
    "copy_areas": SOURCE + 'ws.Range("A1,A2").Copy\nv = ws.Range("D1").PasteSpecial(xlPasteValues)',
    "copy_areas_apart": SOURCE + 'ws.Range("A1,B2").Copy\nv = Application.CutCopyMode',
}


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String",
             "Dim failed As String, v As Variant, out As String", "On Error Resume Next", "Err.Clear",
             *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             'out = failed & Show(v) & ";" & Dump(ws)', "Application.CutCopyMode = False", f"Case{index} = out",
             "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, setup) in enumerate(LAYOUTS.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  "ws.Activate", "Application.CutCopyMode = False", f'ws.Name = "{name[:31]}"',
                  f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, setup))
    build += ["wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")[: len(LAYOUTS)]
    record = {"helper": HELPER, "layouts": [{"name": name, "setup": setup, "answers": answer}
                                            for (name, setup), answer in zip(LAYOUTS.items(), answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, answer in zip(LAYOUTS, answers, strict=True):
        head, _, dump = answer.partition(";")
        parts = dump.split("<>")
        cells = [f"{'ABCDEF'[i % 6]}{i // 6 + 1}={cell}" for i, cell in enumerate(parts[:36])
                 if not cell.startswith("~Empty~General~False")]
        print(f"{name}: {head} | {' | '.join(cells)} | used {parts[36]} mode {parts[37]} D as wide as A {parts[38]}")


if __name__ == "__main__":
    main()
