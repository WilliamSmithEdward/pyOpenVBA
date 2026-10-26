"""Workbook.SaveAs to CSV and text formats, in live Excel.

Each case fills a new workbook as the probe's Select Case says, saves it with SaveAs in a text FileFormat, and
records what the workbook is afterwards -- its name, FileFormat, Saved, ReadOnly, how many sheets, the active
sheet's name and whether a cell still holds its formula -- and the script records the bytes the save wrote. Some
cases save again with Save, to show what a second save writes.

    python scripts/measure_text_saves.py

writes tests/fixtures/text_saves.json, which tests/test_excel_text_saves.py replays.
"""
from __future__ import annotations

import base64
import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/text_saves.json"
#: Each case: its name, the file it saves as, the FileFormat, and what fills the workbook (PROBE's Select Case).
CASES: list[tuple[str, str, int]] = [
    ("basic", "basic.csv", 6),
    ("formats", "formats.csv", 6),
    ("general_numbers", "general.csv", 6),
    ("narrow_column", "narrow.csv", 6),
    ("quoting", "quoting.csv", 6),
    ("formulas", "formulas.csv", 6),
    ("offset", "offset.csv", 6),
    ("ragged", "ragged.csv", 6),
    ("ragged_long", "ragged_long.csv", 6),
    ("empty_sheet", "empty.csv", 6),
    ("second_sheet", "second.csv", 6),
    ("hidden", "hidden.csv", 6),
    ("merged", "merged.csv", 6),
    ("prefixed", "prefixed.csv", 6),
    ("resave", "resave.csv", 6),
    ("no_extension", "noext", 6),
    ("txt_extension", "csv_as.txt", 6),
    ("utf8", "utf8.csv", 62),
    ("unicode_text", "unicode.txt", 42),
    ("text", "text.txt", -4158),
    ("text_windows", "windows.txt", 20),
    ("text_msdos", "msdos.txt", 21),
    ("csv_mac", "mac.csv", 22),
    ("csv_windows", "windows.csv", 23),
    ("csv_msdos", "msdos.csv", 24),
    ("text_mac", "mac.txt", 19),
    ("printer", "printer.prn", 36),
    ("local", "local.csv", 6),
    ("accounting", "accounting.csv", 6),
    ("fills", "fills.csv", 6),
    ("formatted_edge", "edge.csv", 6),
    ("resave_two_sheets", "twosheets.csv", 6),
    ("unicode_no_extension", "unicode_noext", 42),
    ("msdos_no_extension", "msdos_noext", 24),
    ("utf8_no_extension", "utf8_noext", 62),
    ("text_no_extension", "text_noext", -4158),
    ("windows_text_no_extension", "wintext_noext", 20),
    ("msdos_text_no_extension", "dostext_noext", 21),
    ("windows_csv_no_extension", "wincsv_noext", 23),
    ("name_clash", "clash.csv", 6),
    ("negative_date", "negdate.csv", 6),
    ("conditional_format", "conditional.csv", 6),
    ("text_section", "section.csv", 6),
]

PROBE = r'''Private Function Esc(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        If code < 32 Or code > 126 Or ch = "|" Then out = out & "<" & code & ">" Else out = out & ch
    Next
    Esc = out
End Function

Private Sub Mixed(ws As Worksheet)
    ws.Range("A1").Value = "caf" & ChrW(233)
    ws.Range("B1").Value = "omega " & ChrW(937)
    ws.Range("C1").Value = "a,b"
    ws.Range("D1").Value = "say ""hi"""
    ws.Range("A2").Value = "tab" & Chr(9) & "in"
    ws.Range("B2").Value = "two" & Chr(10) & "lines"
    ws.Range("C2").Value = 1234.5
    ws.Range("C2").NumberFormat = "#,##0.00"
    ws.Range("D2").Value = 7
    ws.Range("A3").Value = "a;b"
    ws.Range("B3").Value = " pad "
    ws.Range("C3").Value = "euro " & ChrW(8364)
End Sub

Public Function Probe(name As String, path As String, fileFormat As Long) As String
    Dim wb As Workbook, ws As Worksheet, s As String, i As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Select Case name
        Case "basic", "resave", "no_extension", "txt_extension", "local"
            ws.Range("A1").Value = "name": ws.Range("B1").Value = "amount"
            ws.Range("A2").Value = "alpha": ws.Range("B2").Value = 1
            ws.Range("A3").Value = "beta": ws.Range("B3").Value = 2.5
        Case "formats"
            ws.Range("A1").Value = 1234.5: ws.Range("A1").NumberFormat = "#,##0.00"
            ws.Range("B1").Value = DateSerial(2026, 1, 2): ws.Range("B1").NumberFormat = "m/d/yyyy"
            ws.Range("C1").Value = 0.125: ws.Range("C1").NumberFormat = "0.00%"
            ws.Range("D1").Value = 5: ws.Range("D1").NumberFormat = "$#,##0.00"
            ws.Range("E1").Value = DateSerial(2026, 1, 2) + TimeSerial(10, 30, 0)
            ws.Range("F1").Value = 0.5: ws.Range("F1").NumberFormat = "h:mm AM/PM"
            ws.Range("G1").Value = -5: ws.Range("G1").NumberFormat = "0.00_);(0.00)"
            ws.Range("H1").Value = 42: ws.Range("H1").NumberFormat = "@"
            ws.Range("I1").Value = 3.14159: ws.Range("I1").NumberFormat = "0.0"
            ws.Range("J1").Value = 1.5: ws.Range("J1").NumberFormat = "# ?/?"
            ws.Range("K1").Value = 12345: ws.Range("K1").NumberFormat = "0.00E+00"
            ws.Range("L1").Value = 7: ws.Range("L1").NumberFormat = """N""0"
        Case "general_numbers"
            ws.Range("A1").Value = 1 / 3: ws.Range("B1").Value = 0.1 + 0.2: ws.Range("C1").Value = 123456789012345#
            ws.Range("D1").Value = 1.23456789012346E+17: ws.Range("E1").Value = 0.000001234
            ws.Range("F1").Value = -2.5: ws.Range("G1").Value = 1E+20: ws.Range("H1").Value = 12345678901#
            ws.Range("A2").Value = 1 / 3: ws.Columns(1).ColumnWidth = 40
        Case "narrow_column"
            ws.Range("A1").Value = 123456789: ws.Columns(1).ColumnWidth = 2
            ws.Range("B1").Value = 1 / 3: ws.Columns(2).ColumnWidth = 3
            ws.Range("C1").Value = DateSerial(2026, 1, 2): ws.Columns(3).ColumnWidth = 2
            ws.Range("D1").Value = "long text here": ws.Columns(4).ColumnWidth = 2
            ws.Range("E1").Value = 1234.5: ws.Range("E1").NumberFormat = "#,##0.00": ws.Columns(5).ColumnWidth = 2
        Case "quoting"
            ws.Range("A1").Value = "a,b": ws.Range("B1").Value = "say ""hi"""
            ws.Range("C1").Value = "two" & Chr(10) & "lines": ws.Range("D1").Value = "cr" & Chr(13) & "only"
            ws.Range("E1").Value = " lead": ws.Range("F1").Value = "trail ": ws.Range("G1").Value = "tab" & Chr(9) & "x"
            ws.Range("H1").Value = "semi;colon": ws.Range("I1").Value = "'quoted'": ws.Range("J1").Value = """"
            ws.Range("K1").Value = "a""": ws.Range("L1").NumberFormat = "@": ws.Range("L1").Value = "=not formula"
        Case "formulas"
            ws.Range("A1").Formula = "=1+2": ws.Range("B1").Formula = "=""x"""
            ws.Range("C1").Formula = "=1/0": ws.Range("D1").Value = True: ws.Range("E1").Value = False
            ws.Range("F1").Formula = "=A1>2": ws.Range("G1").Formula = "=NA()"
            ws.Range("H1").Formula = "=""a,b"""
        Case "offset"
            ws.Range("C3").Value = "x": ws.Range("D4").Value = "y"
        Case "ragged"
            ws.Range("A1").Value = "a": ws.Range("B2").Value = "b": ws.Range("D2").Value = "d"
            ws.Range("A4").Value = "e"
        Case "ragged_long"
            ws.Range("A1").Value = 1: ws.Range("E1").Value = 5
            For i = 2 To 25: ws.Cells(i, 1).Value = i: Next
        Case "empty_sheet"
        Case "second_sheet"
            ws.Range("A1").Value = "first"
            Set ws = wb.Worksheets.Add(After:=ws)
            ws.Name = "Data"
            ws.Range("A1").Value = "second"
            ws.Activate
        Case "hidden"
            ws.Range("A1").Value = "a": ws.Range("B1").Value = "b": ws.Range("C1").Value = "c"
            ws.Range("A2").Value = "d": ws.Range("B2").Value = "e": ws.Range("C2").Value = "f"
            ws.Rows(2).Hidden = True: ws.Columns(2).Hidden = True
        Case "merged"
            ws.Range("A1").Value = "m": ws.Range("A1:B2").Merge: ws.Range("C1").Value = "c"
        Case "prefixed"
            ws.Range("A1").Value = "'007": ws.Range("B1").Value = "'=1+1": ws.Range("C1").Value = "'"
        Case "accounting"
            ws.Range("A1").Value = 5: ws.Range("A1").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            ws.Range("B1").Value = -5: ws.Range("B1").NumberFormat = ws.Range("A1").NumberFormat
            ws.Range("C1").Value = 0: ws.Range("C1").NumberFormat = ws.Range("A1").NumberFormat
            ws.Range("D1").Value = "x": ws.Range("D1").NumberFormat = ws.Range("A1").NumberFormat
            ws.Range("E1").Value = 1234.5: ws.Range("E1").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
            ws.Columns(1).ColumnWidth = 20
        Case "fills"
            ws.Range("A1").Value = 5: ws.Range("A1").NumberFormat = "0* "
            ws.Range("B1").Value = "ab": ws.Range("B1").NumberFormat = "@*-"
            ws.Range("C1").Value = 7: ws.Range("C1").NumberFormat = "*x0"
            ws.Columns(2).ColumnWidth = 20
        Case "formatted_edge"
            ws.Range("A1").Value = "a": ws.Range("B2").Value = "b"
            ws.Range("D4").Interior.Color = RGB(255, 0, 0): ws.Range("E1").Font.Bold = True
        Case "resave_two_sheets"
            ws.Range("A1").Value = "first"
            wb.Worksheets.Add(After:=ws).Range("A1").Value = "second"
            ws.Activate
        Case "unicode_no_extension", "msdos_no_extension", "utf8_no_extension", "text_no_extension", _
             "windows_text_no_extension", "msdos_text_no_extension", "windows_csv_no_extension"
            ws.Range("A1").Value = "a": ws.Range("B1").Value = 1
        Case "name_clash"
            ws.Range("A1").Value = "a"
            wb.Worksheets.Add(After:=ws).Name = "clash"
            ws.Activate
        Case "negative_date"
            ws.Range("A1").Value = -1: ws.Range("A1").NumberFormat = "m/d/yyyy": ws.Range("B1").Value = "after"
        Case "conditional_format"
            ws.Range("A1").Value = 5
            ws.Range("A1").FormatConditions.Add(xlCellValue, xlGreater, "0").NumberFormat = "0.000"
            ws.Range("B1").Value = -5
            ws.Range("B1").FormatConditions.Add(xlCellValue, xlGreater, "0").NumberFormat = "0.000"
        Case "text_section"
            ws.Range("A1").Value = "x": ws.Range("A1").NumberFormat = "General;General;General;""Name: ""@"
            ws.Range("B1").Value = 5: ws.Range("B1").NumberFormat = "0;-0;""zero"";""T:""@"
            ws.Range("C1").Value = 0: ws.Range("C1").NumberFormat = "0;-0;""zero"";""T:""@"
        Case Else
            Mixed ws
    End Select
    If name = "local" Then
        wb.SaveAs Filename:=path, FileFormat:=fileFormat, Local:=True
    Else
        wb.SaveAs Filename:=path, FileFormat:=fileFormat
    End If
    s = Esc(wb.Name) & "|" & wb.FileFormat & "|" & wb.Saved & "|" & wb.ReadOnly & "|" & wb.Worksheets.Count & "|" & _
        Esc(wb.ActiveSheet.Name) & "|" & wb.ActiveSheet.Range("A1").HasFormula
    If name = "resave" Or name = "resave_two_sheets" Then
        wb.ActiveSheet.Range("C1").Value = "added"
        s = s & "|dirty=" & wb.Saved
        wb.Save
        s = s & "|after save=" & wb.Saved & " " & wb.FileFormat & " " & Esc(wb.Name) & " " & Esc(wb.ActiveSheet.Name)
    End If
    If name = "name_clash" Then s = s & "|sheets=" & Esc(wb.Worksheets(1).Name) & "," & Esc(wb.Worksheets(2).Name)
    wb.Close False
    Application.DisplayAlerts = True
    Probe = s
End Function
'''


def main() -> None:
    answers: dict[str, str] = {}
    saved: dict[str, str] = {}
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function V() As String\nV = Application.Version & "|" & Application.Build\n'
                                'End Function', "V")
        for name, file, file_format in CASES:
            path = Path(folder) / file
            result = excel.run_vba(PROBE, "Probe", [name, str(path), file_format], timeout=300.0)
            assert result.ok and not result.dialogs, f"{name}: {result.outcome}: {result.message} {result.error}"
            answers[name] = str(result.value)
            written = sorted(Path(folder).glob(file + "*"))
            assert written, f"{name}: nothing written for {file}"
            saved[name] = base64.b64encode(written[0].read_bytes()).decode("ascii")
            print(name, answers[name], written[0].name, written[0].read_bytes()[:120], flush=True)
    OUT.write_text(json.dumps({
        "excel_version_build": version.value, "measured_at": datetime.now(timezone.utc).isoformat(),
        "probe": PROBE, "cases": [list(case) for case in CASES], "answers": answers, "files": saved}, indent=2) + "\n",
        encoding="utf-8")


if __name__ == "__main__":
    main()
