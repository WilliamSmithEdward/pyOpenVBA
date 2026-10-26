"""Workbooks.Open of CSV and text files, and Open's other arguments, in live Excel.

Each case writes a file, opens it with Workbooks.Open as the case says, and records the workbook Excel makes of
it: its name, its sheet's name, how many sheets, FileFormat, Saved, ReadOnly and the used range, then every cell
that holds something, as TypeName, Value2, NumberFormat, its formula and its prefix character. Text is spelled
with every character outside printable ASCII, and < > | ;, as <code>, since the harness's answer cannot carry
control characters. Cases that try what may fail record the error instead.

Sequences open several files in one call, in order, to show what one open leaves for the next. Typing writes
strings with a currency or a percent sign through Range.Value, and Columns types the same through TextToColumns,
beside the "signed" file that holds them. A read-only workbook is never saved by Save:
with alerts off Excel saves a copy into Application.DefaultFilePath, outside the probe's folder.

    python scripts/measure_text_files.py

writes tests/fixtures/text_files.json, which tests/test_excel_text_files.py replays.
"""
from __future__ import annotations

import base64
import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/text_files.json"
BOM = b"\xef\xbb\xbf"
WIDE = b"\xff\xfe"
MIXED = "a\tb,c;d e|f\r\n"


def _ansi(text: str) -> bytes:
    return text.encode("cp1252")


#: Strings typed through Range.Value, and the same in a file, where a currency or a percent sign leads.
SIGNED = ["%5", "% 5", "%-5", "-%5", "%5%", "5%%", "%", "(%5)", "\u20ac5", "\u20ac 5", "-\u20ac5", "\u20ac-5",
          "(\u20ac5)", "5\u20ac", "5 \u20ac", "\u20ac1,234.50", "\u20ac.5", "\u20ac5%", "\u00a35", "\u00a55", "5$",
          "$ 5", "$-5", "5.5\u20ac", "1,234 \u20ac", "-5\u20ac", "(5\u20ac)", "\u20ac$5", "$\u20ac5", "\u20ac",
          "\u20ac1/2", "%1e2", "%1 1/2", "%5.5", "%5.50", "%.5", "%0", "%1,234", "%1,234.5", "%$5", "$%5", "%1/2",
          "%(5)", "% -5", "%5 ", " %5", "1e2%", "1 1/2%", "\u20ac1e2", "\u20ac5.0", "5.0\u20ac", "%\u20ac5",
          "\u20ac%5", "\u20ac1 1/2", "\u20ac 1,234", "(\u20ac 5)", "\u20ac(5)", "-\u20ac 5", "5-\u20ac", "5\u20ac-"]

#: Each case: its name, the file it writes and the file's bytes; PROBE's Select Case says how each opens.
CASES: list[tuple[str, str, bytes]] = [
    # How a CSV splits into rows and fields.
    ("basic", "data.csv", _ansi("name,amount\r\nalpha,1\r\nbeta,2.5\r\n")),
    ("lf_only", "lf.csv", _ansi("a,b\nc,d\n")),
    ("cr_only", "cr.csv", _ansi("a,b\rc,d\r")),
    ("no_final_newline", "last.csv", _ansi("a,b\r\nc,d")),
    ("empty_lines", "gaps.csv", _ansi("a\r\n\r\nb\r\n\r\n")),
    ("quoted", "quoted.csv", _ansi('"a,b","c""d",""\r\n')),
    ("quoted_newline", "lines.csv", _ansi('"line1\r\nline2",x\r\n"cr\ronly",y\r\n"lf\nonly",z\r\n')),
    ("quote_inside", "inside.csv", _ansi('ab"cd,e\r\n')),
    ("quote_after_text", "after.csv", _ansi('"ab"cd,e\r\n"x""y"z,"""",""""""\r\n')),
    ("spaces", "spaces.csv", _ansi(' a , b ," c ",   \r\n')),
    ("space_before_quote", "spacequote.csv", _ansi('x, "a,b"\r\n')),
    ("trailing_commas", "trailing.csv", _ansi("a,,\r\n,,b\r\n")),
    ("ragged", "ragged.csv", _ansi("a\r\nb,c,d\r\n")),
    ("unterminated_quote", "open.csv", _ansi('x,"abc\r\ndef\r\n')),
    ("tab_in_field", "tab.csv", _ansi("a\tb,c\r\n")),
    ("semicolons", "semi.csv", _ansi("a;b;c\r\n")),
    ("sep_line", "sepline.csv", _ansi('sep=;\r\na;b;"c;d"\r\ne,f\r\n')),
    ("sep_pipe", "seppipe.csv", _ansi("sep=|\r\na|b,c\r\n")),
    ("sep_upper", "sepupper.csv", _ansi("SEP=;\r\na;b\r\n")),
    ("sep_extra", "sepextra.csv", _ansi("sep=;x\r\na;b\r\n")),
    ("sep_empty", "sepempty.csv", _ansi("sep=\r\na,b\r\n")),
    ("sep_spaced", "sepspaced.csv", _ansi(" sep=;\r\na;b\r\n")),
    ("sep_second", "sepsecond.csv", _ansi("a,b\r\nsep=;\r\nc;d\r\n")),
    ("sep_bom", "sepbom.csv", BOM + b"sep=;\r\na;b\r\n"),
    # How each field is typed.
    ("numbers", "numbers.csv", _ansi("1,2.5,-3,+4,.5,1e5,1E-3,007,1.0,-0,(5),1.,5-\r\n")),
    ("big_numbers", "big.csv", _ansi("12345678901234567890,123456789012345,1234567890123456,"
                                     "0.123456789012345678,1e308,1e309\r\n")),
    ("thousands", "thousands.csv", _ansi('"1,234","1,234.5","12,34","1,2345"\r\n')),
    ("percent_currency", "money.csv", _ansi('5%,-5%,$5,"$1,234.50",($5),5 %\r\n')),
    ("signed", "signed.csv", _ansi(",".join(f'"{text}"' for text in SIGNED) + "\r\n")),
    ("dates", "dates.csv", _ansi("1/2/2026,2026-01-02,1/2,Jan 2,2-Jan-2026,13/1/2026,1/13/2026,2026/1/2,"
                                 "Jan-26,1-2\r\n")),
    ("times", "times.csv", _ansi("10:30,10:30:15,10:30 PM,1/2/2026 10:30,25:00,10:30:15.5\r\n")),
    ("booleans", "bools.csv", _ansi("TRUE,false,True,yes\r\n")),
    ("errors", "errors.csv", _ansi("#N/A,#DIV/0!,#VALUE!,#REF!,#NAME?,#NUM!,#NULL!,#n/a\r\n")),
    ("formulas", "formulas.csv", _ansi('=1+2,=A1*2,"=SUM(1,2)",="x", =1+1,=\r\n')),
    ("formula_constant", "constant.csv", _ansi("=1\r\n")),
    ("apostrophes", "quotes.csv", _ansi("'abc,'123,''x,'\r\n")),
    ("text_like", "textlike.csv", _ansi("1 1/2,1/0,  12  ,12 34,0x1F,1.2.3,1/4\r\n")),
    ("quoted_typed", "quotedtyped.csv", _ansi('"12","TRUE","=1+2","1/2/2026"," 7"\r\n')),
    ("long_text", "long.csv", _ansi("x" * 300 + "," + "y" * 40000 + "\r\nend\r\n")),
    ("field_32759", "f32759.csv", _ansi("y" * 32759 + "\r\nend\r\n")),
    ("field_32760", "f32760.csv", _ansi("y" * 32760 + "\r\nend\r\n")),
    ("field_32768", "f32768.csv", _ansi("y" * 32768 + "\r\nend\r\n")),
    ("quoted_long", "qlong.csv", _ansi('"' + "q" * 40000 + '",after\r\n')),
    ("line_past_field", "linepast.csv", _ansi("x" * 300 + "," + "y" * 32759 + "\r\n")),
    # How the bytes are read.
    ("utf8_bom", "bom.csv", BOM + "caf\u00e9,na\u00efve,\u20ac5\r\n".encode("utf-8")),
    ("utf8_no_bom", "nobom.csv", "caf\u00e9,\u20ac\r\n".encode("utf-8")),
    ("ansi", "ansi.csv", _ansi("caf\u00e9,\u20ac5\r\n")),
    ("utf16", "wide.csv", WIDE + "a,b\tc\r\nd\te\r\n".encode("utf-16-le")),
    ("utf16_big_endian", "big_endian.csv", b"\xfe\xff" + "a,b\tc\r\n".encode("utf-16-be")),
    ("txt_utf8_bom", "bom.txt", BOM + "a\tcaf\u00e9\r\n".encode("utf-8")),
    ("txt_utf16", "wide.txt", WIDE + "a\tb,c\r\n".encode("utf-16-le")),
    ("tiny_one", "tiny1.csv", b"a"),
    ("tiny_three", "tiny3.csv", b"a,b"),
    ("tiny_line", "tiny_line.csv", b"a\r\n"),
    ("tiny_four", "tiny4.csv", b"a,bc"),
    ("empty", "empty.csv", b""),
    ("only_newline", "newline.csv", b"\r\n"),
    # What the workbook and its sheet are called.
    ("long_name", "a_file_name_well_past_thirty_one_characters.csv", _ansi("a,b\r\n")),
    ("bracket_name", "a[b].csv", _ansi("a,b\r\n")),
    ("dotted_name", "my.data.csv", _ansi("a,b\r\n")),
    ("upper_extension", "UPPER.CSV", _ansi("a,b\r\n")),
    # Text files and the arguments that read them.
    ("txt_default", "plain.txt", _ansi(MIXED)),
    ("txt_commas", "commas.txt", _ansi(MIXED)),
    ("txt_semicolons", "semis.txt", _ansi(MIXED)),
    ("txt_spaces", "spaced.txt", _ansi(MIXED)),
    ("txt_spaces_run", "spacerun.txt", _ansi("a  b   c\r\n d \r\n")),
    ("txt_nothing", "whole.txt", _ansi("a\tb,c;d e|f\r\nab\tc\r\nabcdefgh\ti\r\n\tx\r\nabcdefg\t\ty\r\n")),
    ("txt_custom", "custom.txt", _ansi(MIXED)),
    ("txt_custom_long", "customlong.txt", _ansi("a||b|c\r\n")),
    ("txt_tabs", "tabs.txt", _ansi(MIXED)),
    ("txt_quoted", "quoted.txt", _ansi('"a\tb"\tc\r\n')),
    ("csv_custom", "custom.csv", _ansi(MIXED)),
    ("csv_local", "local.csv", _ansi("a,b;c\r\n")),
    ("prn", "data.prn", _ansi("a\tb c\r\n")),
    ("other_extension", "data.dat", _ansi("a\tb,c\r\n")),
    ("readonly", "readonly.csv", _ansi("a,1\r\n")),
]

#: Cases that open a workbook the probe makes first, at the case's path, rather than a file the script writes.
MADE = {"xlsx_args": "book.xlsx", "readonly_saveas_same": "same.xlsx", "readonly_saveas_other": "other.xlsx"}
#: Cases that record the error Workbooks.Open raises.
FAILING = {"missing_file": "absent.csv", "garbage_xlsx": "garbage.xlsx"}
#: Opens one after another in one call, each as a file and a way to open it, to show what an open leaves behind.
SEQUENCES = {
    "txt_delimiter": [("s1.txt", "plain"), ("s2.txt", "semicolons"), ("s3.txt", "plain"), ("s4.csv", "plain"),
                      ("s5.txt", "plain"), ("s6.csv", "spaces"), ("s7.txt", "plain"), ("s8.txt", "tabs"),
                      ("s9.txt", "plain"), ("s10.txt", "commas"), ("s11.txt", "plain"), ("s12.csv", "plain")],
}

PROBE = r'''Private Function Esc(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        If code < 32 Or code > 126 Or ch = "<" Or ch = ">" Or ch = "|" Or ch = ";" Then
            out = out & "<" & code & ">"
        Else
            out = out & ch
        End If
    Next
    Esc = out
End Function

Private Function Shown(v As Variant) As String
    If VarType(v) = vbString Then
        If Len(v) > 300 Then
            Shown = "len" & Len(v) & "[" & Esc(Left$(v, 20)) & "..." & Esc(Right$(v, 20)) & "]"
        Else
            Shown = "[" & Esc(CStr(v)) & "]"
        End If
    Else
        Shown = CStr(v)
    End If
End Function

Private Function Kind(c As Range) As String
    Kind = TypeName(c.Value) & ":" & Shown(c.Value2) & ":" & Esc(c.NumberFormat)
    If c.HasFormula Then Kind = Kind & ":" & Esc(c.Formula)
    If c.PrefixCharacter <> "" Then Kind = Kind & ":p" & Esc(c.PrefixCharacter)
End Function

Private Function Book(wb As Workbook) As String
    Dim ws As Worksheet, c As Range, s As String
    Set ws = wb.Worksheets(1)
    s = Esc(wb.Name) & "|" & Esc(ws.Name) & "|" & wb.Worksheets.Count & "|" & wb.FileFormat & "|" & wb.Saved & _
        "|" & wb.ReadOnly & "|" & ws.UsedRange.Address(False, False)
    For Each c In ws.UsedRange.Cells
        If Not IsEmpty(c.Value) Or c.HasFormula Then s = s & "|" & c.Address(False, False) & "=" & Kind(c)
    Next
    Book = s
End Function

Private Function Opened(path As String, how As String) As Workbook
    Select Case how
        Case "plain": Set Opened = Workbooks.Open(path)
        Case "tabs": Set Opened = Workbooks.Open(path, Format:=1)
        Case "commas": Set Opened = Workbooks.Open(path, Format:=2)
        Case "spaces": Set Opened = Workbooks.Open(path, Format:=3)
        Case "semicolons": Set Opened = Workbooks.Open(path, Format:=4)
        Case "nothing": Set Opened = Workbooks.Open(path, Format:=5)
        Case "pipe": Set Opened = Workbooks.Open(path, Format:=6, Delimiter:="|")
        Case "pipes": Set Opened = Workbooks.Open(path, Format:=6, Delimiter:="||")
        Case "semicolon": Set Opened = Workbooks.Open(path, Format:=6, Delimiter:=";")
        Case "local": Set Opened = Workbooks.Open(path, Local:=True)
        Case "readonly": Set Opened = Workbooks.Open(path, ReadOnly:=True)
    End Select
End Function

Public Function Probe(name As String, path As String) As String
    Dim wb As Workbook, s As String, how As String
    Application.DisplayAlerts = False
    Select Case name
        Case "txt_commas": how = "commas"
        Case "txt_semicolons": how = "semicolons"
        Case "txt_spaces", "txt_spaces_run": how = "spaces"
        Case "txt_nothing": how = "nothing"
        Case "txt_custom": how = "pipe"
        Case "txt_custom_long": how = "pipes"
        Case "txt_tabs", "txt_quoted": how = "tabs"
        Case "csv_custom": how = "semicolon"
        Case "csv_local": how = "local"
        Case "readonly": how = "readonly"
        Case Else: how = "plain"
    End Select
    Select Case name
        Case "xlsx_args", "readonly_saveas_same", "readonly_saveas_other"
            Set wb = Workbooks.Add(xlWBATWorksheet)
            wb.Worksheets(1).Range("A1").Value = 1
            wb.SaveAs path, xlOpenXMLWorkbook
            wb.Close False
            If name = "xlsx_args" Then
                Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=False, Password:="pw", _
                    WriteResPassword:="wr", IgnoreReadOnlyRecommended:=True, Editable:=False, Notify:=False, _
                    AddToMru:=False, Local:=False, CorruptLoad:=xlNormalLoad)
            Else
                Set wb = Workbooks.Open(path, ReadOnly:=True)
                wb.Worksheets(1).Range("A1").Value = 2
                On Error Resume Next
                If name = "readonly_saveas_same" Then
                    wb.SaveAs path, xlOpenXMLWorkbook
                Else
                    wb.SaveAs Replace(path, ".xlsx", "_copy.xlsx"), xlOpenXMLWorkbook
                End If
                s = "|saveas=" & Err.Number & ":" & Esc(Err.Description) & " readonly=" & wb.ReadOnly & _
                    " saved=" & wb.Saved & " name=" & Esc(wb.Name)
                On Error GoTo 0
            End If
        Case "missing_file", "garbage_xlsx"
            On Error Resume Next
            Set wb = Workbooks.Open(path)
            Probe = "error=" & Err.Number & ":" & Esc(Err.Description) & " books=" & Workbooks.Count
            On Error GoTo 0
            If Not wb Is Nothing Then wb.Close False
            Application.DisplayAlerts = True
            Exit Function
        Case Else: Set wb = Opened(path, how)
    End Select
    s = Book(wb) & s
    wb.Close False
    Application.DisplayAlerts = True
    Probe = s
End Function

Public Function Sequence(steps As String) As String
    Dim wb As Workbook, s As String, step As Variant, parts() As String
    Application.DisplayAlerts = False
    For Each step In Split(steps, Chr(2))
        parts = Split(step, Chr(1))
        Set wb = Opened(CStr(parts(0)), CStr(parts(1)))
        s = s & "~" & Book(wb)
        wb.Close False
    Next
    Application.DisplayAlerts = True
    Sequence = Mid$(s, 2)
End Function

Public Function Typing(values As String) As String
    Dim wb As Workbook, ws As Worksheet, s As String, value As Variant
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    For Each value In Split(values, Chr(1))
        ws.Range("A1").Clear
        ws.Range("A1").Value = value
        s = s & "~" & Esc(CStr(value)) & "=" & Kind(ws.Range("A1"))
    Next
    wb.Close False
    Application.DisplayAlerts = True
    Typing = Mid$(s, 2)
End Function

Public Function Columns(values As String) As String
    Dim wb As Workbook, ws As Worksheet, s As String, items() As String, i As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    items = Split(values, Chr(1))
    For i = 0 To UBound(items)
        ws.Cells(i + 1, 1).Value = "'" & items(i)
    Next
    ws.Range(ws.Cells(1, 1), ws.Cells(UBound(items) + 1, 1)).TextToColumns DataType:=xlDelimited, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False
    For i = 0 To UBound(items)
        s = s & "~" & Esc(items(i)) & "=" & Kind(ws.Cells(i + 1, 1))
    Next
    wb.Close False
    Application.DisplayAlerts = True
    Columns = Mid$(s, 2)
End Function
'''


def sequence_steps(root: Path, steps: list[tuple[str, str]]) -> str:
    """A sequence as Sequence takes it: each file's path and how it opens, apart by Chr(1), the steps by Chr(2)."""
    return "\x02".join(f"{root / file}\x01{how}" for file, how in steps)


def main() -> None:
    answers: dict[str, str] = {}
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        root = Path(folder)
        excel.new_document()
        version = excel.run_vba('Function V() As String\nV = Application.Version & "|" & Application.Build\n'
                                'End Function', "V")
        for _, file, data in CASES:
            (root / file).write_bytes(data)
        (root / FAILING["garbage_xlsx"]).write_bytes(b"hello")
        for steps in SEQUENCES.values():
            for file, _ in steps:
                (root / file).write_bytes(_ansi(MIXED))
        # The sequences run first, in a session no Format has touched.
        for name, steps in SEQUENCES.items():
            result = excel.run_vba(PROBE, "Sequence", [sequence_steps(root, steps)], timeout=300.0)
            assert result.ok and not result.dialogs, f"{name}: {result.outcome}: {result.message} {result.error}"
            answers[name] = str(result.value)
            print(name, answers[name][:300], flush=True)
        for proc in ("Typing", "Columns"):
            typed = excel.run_vba(PROBE, proc, ["\x01".join(SIGNED)], timeout=300.0)
            assert typed.ok and not typed.dialogs, f"{proc}: {typed.outcome}: {typed.message} {typed.error}"
            answers[proc.lower()] = str(typed.value)
        planned = [(name, file) for name, file, _ in CASES] + list(MADE.items()) + list(FAILING.items())
        for name, file in planned:
            path = str(root / file)
            result = excel.run_vba(PROBE, "Probe", [name, path], timeout=300.0)
            assert result.ok and not result.dialogs, f"{name}: {result.outcome}: {result.message} {result.error}"
            # The folder is this machine's; the test puts its own back.
            answers[name] = str(result.value).replace(path, "{path}")
            print(name, answers[name][:300], flush=True)
    OUT.write_text(json.dumps({
        "excel_version_build": version.value, "measured_at": datetime.now(timezone.utc).isoformat(),
        "probe": PROBE, "signed": SIGNED, "mixed": MIXED,
        "files": {name: {"file": file, "data": base64.b64encode(data).decode("ascii")} for name, file, data in CASES},
        "made": MADE, "failing": FAILING, "sequences": SEQUENCES, "answers": answers}, indent=2) + "\n",
        encoding="utf-8")


if __name__ == "__main__":
    main()
