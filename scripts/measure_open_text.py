"""Workbooks.OpenText in live Excel: a text file read with the delimiters, qualifier, start row, field kinds, fixed
widths, separators and code page a macro gives.

Each case writes a file, opens it with OpenText as the case says, and records the workbook Excel makes, as
scripts/measure_text_files.py records one: its name, its sheet's name, how many sheets, FileFormat, Saved and
ReadOnly, the used range and every cell that holds something. The cases that read with Excel's defaults run first,
in a session no OpenText has touched. A sequence then opens several files in one call, to show what OpenText
remembers of what it was given, and Remembered what that has to do with TextToColumns and Workbooks.Open, before
putting the settings back to a tab alone; the other cases follow, in order, the one that sets a Mac code page last.

    python scripts/measure_open_text.py

writes tests/fixtures/open_text.json, which tests/test_excel_open_text.py replays.
"""
from __future__ import annotations

import base64
import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/open_text.json"
BOM, WIDE = b"\xef\xbb\xbf", b"\xff\xfe"
MIXED = "a\tb,c;d e|f\r\n1,2;3 4|5\r\n"
OPEN = "Workbooks.OpenText Filename:=path"


def _ansi(text: str) -> bytes:
    return text.encode("cp1252")


#: Each case: its name, the file it writes, the file's bytes, and the OpenText call that reads it.
CASES: list[tuple[str, str, bytes, str]] = [
    # Excel's defaults, first in the session.
    ("default_txt", "default.txt", _ansi(MIXED), OPEN),
    ("default_csv", "default.csv", _ansi(MIXED), OPEN),
    ("default_bom", "defaultbom.txt", BOM + "caf\u00e9\tx\r\n".encode("utf-8"), OPEN),
    ("delimited_none", "delimited.txt", _ansi(MIXED), f"{OPEN}, DataType:=xlDelimited"),
    ("fixed_default", "fixeddefault.txt", _ansi("abc 123 x\r\nde  45  yz\r\n"), f"{OPEN}, DataType:=xlFixedWidth"),
    ("default_aligned", "aligned.txt", _ansi("abc 123 x\r\nde  45  yz\r\n"), OPEN),
    # Delimiters.
    ("tab", "tab.txt", _ansi(MIXED), f"{OPEN}, DataType:=xlDelimited, Tab:=True"),
    ("comma", "comma.txt", _ansi(MIXED), f"{OPEN}, DataType:=xlDelimited, Comma:=True"),
    ("semicolon", "semicolon.txt", _ansi(MIXED), f"{OPEN}, DataType:=xlDelimited, Semicolon:=True"),
    ("space", "space.txt", _ansi(MIXED), f"{OPEN}, DataType:=xlDelimited, Space:=True"),
    ("other", "other.txt", _ansi(MIXED), f'{OPEN}, DataType:=xlDelimited, Other:=True, OtherChar:="|"'),
    ("other_long", "otherlong.txt", _ansi(MIXED), f'{OPEN}, DataType:=xlDelimited, Other:=True, OtherChar:="|;"'),
    ("several", "several.txt", _ansi(MIXED), f"{OPEN}, DataType:=xlDelimited, Tab:=True, Comma:=True, Semicolon:=True"),
    ("consecutive", "consecutive.txt", _ansi("a  b   c\r\n d \r\n"),
     f"{OPEN}, DataType:=xlDelimited, Space:=True, ConsecutiveDelimiter:=True"),
    ("consecutive_quoted", "consecutiveq.txt", _ansi('a,,"",b\r\n,,x\r\n'),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True, ConsecutiveDelimiter:=True"),
    ("qualifier_single", "single.txt", _ansi("'a,b',c\r\n\"d,e\",f\r\n"),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True, TextQualifier:=xlTextQualifierSingleQuote"),
    ("qualifier_none", "noqualifier.txt", _ansi('"a,b",c\r\n'),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True, TextQualifier:=xlTextQualifierNone"),
    ("quoted_newline", "quotednewline.txt", _ansi('"line1\r\nline2",x\r\ny,z\r\n'),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True"),
    ("csv_with_args", "args.csv", _ansi(MIXED), f"{OPEN}, DataType:=xlDelimited, Semicolon:=True, Comma:=False"),
    # Where reading starts.
    ("start_row", "start.txt", _ansi("h1\r\nh2\r\na,b\r\nc,d\r\n"), f"{OPEN}, StartRow:=3, DataType:=xlDelimited, Comma:=True"),
    ("start_row_past", "past.txt", _ansi("a\r\nb\r\n"), f"{OPEN}, StartRow:=10, DataType:=xlDelimited, Comma:=True"),
    ("start_row_quoted", "startquoted.txt", _ansi('"x\r\ny",1\r\nz,2\r\nw,3\r\n'),
     f"{OPEN}, StartRow:=2, DataType:=xlDelimited, Comma:=True"),
    # Field kinds.
    ("field_text", "fieldtext.txt", _ansi("007,007,1/2\r\n"),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True, FieldInfo:=Array(Array(1, 2), Array(2, 1), Array(3, 2))"),
    ("field_skip", "fieldskip.txt", _ansi("a,b,c\r\n"),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True, FieldInfo:=Array(Array(1, 1), Array(2, 9), Array(3, 1))"),
    ("field_dates", "fielddates.txt", _ansi("13/1/2026,2026-01-13,1/13/2026,13/1/26\r\n"),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True, FieldInfo:=Array(Array(1, 4), Array(2, 5), Array(3, 3), "
     "Array(4, 4))"),
    ("field_short", "fieldshort.txt", _ansi("007,007,007\r\n"),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True, FieldInfo:=Array(Array(1, 2))"),
    # Fixed widths.
    ("fixed", "fixed.txt", _ansi("abc123  x\r\nde 45678yz\r\n"),
     f"{OPEN}, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 1), Array(3, 1), Array(6, 2))"),
    ("fixed_skip", "fixedskip.txt", _ansi("abc123  x\r\nde 45678yz\r\n"),
     f"{OPEN}, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 2), Array(3, 9), Array(6, 1))"),
    # Numbers.
    ("separators", "separators.txt", _ansi("1.234,5;2,5;1,234.5\r\n"),
     f'{OPEN}, DataType:=xlDelimited, Semicolon:=True, DecimalSeparator:=",", ThousandsSeparator:="."'),
    ("trailing_minus", "minus.txt", _ansi("5-,-5,5- \r\n"),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True, TrailingMinusNumbers:=True"),
    ("trailing_minus_off", "minusoff.txt", _ansi("5-,-5,5- \r\n"),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True, TrailingMinusNumbers:=False"),
    ("typed", "typed.txt", _ansi("1,2.5,TRUE,1/2/2026,=1+2,#N/A,$5,5%,\u20ac5\r\n"),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True"),
    # Code pages.
    ("origin_utf8", "utf8.txt", "caf\u00e9,\u20ac5\r\n".encode("utf-8"),
     f"{OPEN}, Origin:=65001, DataType:=xlDelimited, Comma:=True"),
    ("origin_msdos", "dos.txt", "caf\u00e9,x\r\n".encode("cp437"),
     f"{OPEN}, Origin:=xlMSDOS, DataType:=xlDelimited, Comma:=True"),
    ("origin_windows", "windows.txt", _ansi("caf\u00e9,\u20ac5\r\n"),
     f"{OPEN}, Origin:=xlWindows, DataType:=xlDelimited, Comma:=True"),
    ("origin_1252", "cp1252.txt", _ansi("caf\u00e9,\u20ac5\r\n"),
     f"{OPEN}, Origin:=1252, DataType:=xlDelimited, Comma:=True"),
    ("bom_default", "bom.txt", BOM + "caf\u00e9,x\r\n".encode("utf-8"), f"{OPEN}, DataType:=xlDelimited, Comma:=True"),
    ("bom_with_origin", "bomorigin.txt", BOM + "caf\u00e9,x\r\n".encode("utf-8"),
     f"{OPEN}, Origin:=1252, DataType:=xlDelimited, Comma:=True"),
    ("utf16", "wide.txt", WIDE + "a,b\tc\r\n".encode("utf-16-le"), f"{OPEN}, DataType:=xlDelimited, Comma:=True"),
    ("long_field", "long.txt", _ansi("x" * 40000 + ",y\r\nend\r\n"), f"{OPEN}, DataType:=xlDelimited, Comma:=True"),
    ("long_field_kept", "longkept.txt", _ansi("x" * 10000 + ",y\r\nend\r\n"),
     f"{OPEN}, DataType:=xlDelimited, Comma:=True"),
    ("local", "local.txt", _ansi("1,5;2\r\n"), f"{OPEN}, DataType:=xlDelimited, Semicolon:=True, Local:=True"),
    ("bracket_name", "a[b].txt", _ansi("a,b\r\n"), f"{OPEN}, DataType:=xlDelimited, Comma:=True"),
    # Last, as Excel remembers its Origin for the opens after it.
    ("origin_mac", "mac.txt", "caf\u00e9,x\r\n".encode("mac-roman"),
     f"{OPEN}, Origin:=xlMacintosh, DataType:=xlDelimited, Comma:=True"),
]
#: The cases that run first, in a session no OpenText or TextToColumns has touched.
DEFAULTS = ("default_txt", "default_csv", "default_bom", "delimited_none", "fixed_default", "default_aligned")
#: Opens in one call, in order: what an OpenText leaves for the next OpenText, and for Workbooks.Open.
SEQUENCE = [
    ("remember1.txt", f"{OPEN}, DataType:=xlDelimited, Semicolon:=True"),
    ("remember2.txt", f"{OPEN}, DataType:=xlDelimited"),
    ("remember3.txt", OPEN),
    ("remember4.txt", "Workbooks.Open path"),
]
#: The files the Remembered function opens, in its order.
SHARED_FILES = [f"shared{index}.txt" for index in range(1, 11)]
#: What the settings OpenText remembers have to do with TextToColumns and Workbooks.Open: each step's answer, the
#: settings put back to a tab alone at the end for the cases after it.
SHARED = r'''
Private Function Split1(cellText As String, how As String) As String
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Range("A1").Value = cellText
    Select Case how
        Case "none": wb.Worksheets(1).Range("A1").TextToColumns DataType:=xlDelimited
        Case "space": wb.Worksheets(1).Range("A1").TextToColumns DataType:=xlDelimited, Space:=True
        Case "fixed": wb.Worksheets(1).Range("A1").TextToColumns DataType:=xlFixedWidth, _
            FieldInfo:=Array(Array(0, 1), Array(3, 1))
    End Select
    Split1 = Book(wb)
    wb.Close False
End Function

Private Function Opened() As String
    Opened = Book(ActiveWorkbook)
    ActiveWorkbook.Close False
End Function

Public Function Remembered(paths As String) As String
    Dim s As String, p() As String, wb As Workbook
    p = Split(paths, Chr(1))
    Application.DisplayAlerts = False
    s = "columns none~" & Split1("p" & Chr(9) & "q,r;t", "none")
    Workbooks.OpenText Filename:=p(0), DataType:=xlDelimited, Tab:=False, Semicolon:=False, Comma:=True
    ActiveWorkbook.Close False
    s = s & "~columns none after comma~" & Split1("p" & Chr(9) & "q,r;t", "none")
    s = s & "~columns space~" & Split1("s t,u", "space")
    Workbooks.OpenText Filename:=p(1)
    s = s & "~opentext after columns space~" & Opened()
    Workbooks.OpenText Filename:=p(2), Origin:=xlMSDOS
    ActiveWorkbook.Close False
    Workbooks.OpenText Filename:=p(3)
    s = s & "~opentext after msdos~" & Opened()
    Set wb = Workbooks.Open(p(4))
    wb.Close False
    Workbooks.OpenText Filename:=p(5)
    s = s & "~opentext after open~" & Opened()
    Workbooks.OpenText Filename:=p(6), Space:=True, ConsecutiveDelimiter:=True, Other:=True, OtherChar:="|"
    ActiveWorkbook.Close False
    Set wb = Workbooks.Open(p(7), Format:=4)
    wb.Close False
    Workbooks.OpenText Filename:=p(8)
    s = s & "~opentext after open format 4~" & Opened()
    s = s & "~columns fixed~" & Split1("ab  cd ef", "fixed")
    Workbooks.OpenText Filename:=p(9), DataType:=xlDelimited, Tab:=True, Comma:=False, Semicolon:=False, _
        Space:=False, Other:=False, ConsecutiveDelimiter:=False, Origin:=xlWindows
    s = s & "~reset~" & Opened()
    Application.DisplayAlerts = True
    Remembered = s
End Function

Public Function Failing(path As String) As String
    Application.DisplayAlerts = False
    On Error Resume Next
    Workbooks.OpenText Filename:=path, DataType:=xlDelimited, Comma:=True
    Failing = "error=" & Err.Number & ":" & Esc(Err.Description) & " books=" & Workbooks.Count
    On Error GoTo 0
    Application.DisplayAlerts = True
End Function
'''
#: The bytes each Remembered file holds: a line showing the delimiters, one with a run of spaces, and E9, a letter
#: whose code page shows.
SHARED_TEXT = b"a\tb,c;d e|f\r\ng  h|i;j\r\n\xe9\r\n"

HELPERS = r'''Private Function Esc(ByVal s As String) As String
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
'''


def probe() -> str:
    """The probe's module: Probe opens one case's file, Sequence the sequence's files one after another."""
    cases = "".join(f'        Case "{name}": {call}\n' for name, _, _, call in CASES)
    steps = "".join(f'        Case "{file}": {call}\n' for file, call in SEQUENCE)
    return (HELPERS + SHARED + "\nPublic Function Probe(name As String, path As String) As String\n"
            "    Application.DisplayAlerts = False\n    Select Case name\n" + cases + "    End Select\n"
            "    Probe = Book(ActiveWorkbook)\n    ActiveWorkbook.Close False\n    Application.DisplayAlerts = True\n"
            "End Function\n\nPublic Function Sequence(steps As String) As String\n"
            "    Dim s As String, piece As Variant, file As String, path As String\n"
            "    Application.DisplayAlerts = False\n"
            "    For Each piece In Split(steps, Chr(1))\n        file = Split(piece, Chr(2))(0)\n"
            "        path = Split(piece, Chr(2))(1)\n        Select Case file\n" + steps +
            "        End Select\n        s = s & \"~\" & Book(ActiveWorkbook)\n        ActiveWorkbook.Close False\n"
            "    Next\n    Application.DisplayAlerts = True\n    Sequence = Mid$(s, 2)\nEnd Function\n")


def sequence_steps(root: Path) -> str:
    """The sequence as Sequence takes it: each file's name and path apart by Chr(2), the steps by Chr(1)."""
    return "\x01".join(f"{file}\x02{root / file}" for file, _ in SEQUENCE)


def shared_paths(root: Path) -> str:
    """The Remembered function's files, by path, apart by Chr(1)."""
    return "\x01".join(str(root / file) for file in SHARED_FILES)


def main() -> None:
    source = probe()
    answers: dict[str, str] = {}
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        root = Path(folder)
        excel.new_document()
        version = excel.run_vba('Function V() As String\nV = Application.Version & "|" & Application.Build\n'
                                'End Function', "V")
        for _, file, data, _ in CASES:
            (root / file).write_bytes(data)
        for file, _ in SEQUENCE:
            (root / file).write_bytes(_ansi(MIXED))
        for file in SHARED_FILES:
            (root / file).write_bytes(SHARED_TEXT)
        for name, file, _, _ in [case for case in CASES if case[0] in DEFAULTS]:
            result = excel.run_vba(source, "Probe", [name, str(root / file)], timeout=300.0)
            assert result.ok and not result.dialogs, f"{name}: {result.outcome}: {result.message} {result.error}"
            answers[name] = str(result.value)
            print(name, answers[name][:300], flush=True)
        result = excel.run_vba(source, "Sequence", [sequence_steps(root)], timeout=300.0)
        assert result.ok and not result.dialogs, f"sequence: {result.outcome}: {result.message} {result.error}"
        answers["sequence"] = str(result.value)
        print("sequence", answers["sequence"][:600], flush=True)
        result = excel.run_vba(source, "Remembered", [shared_paths(root)], timeout=300.0)
        assert result.ok and not result.dialogs, f"shared: {result.outcome}: {result.message} {result.error}"
        answers["shared"] = str(result.value)
        print("shared", answers["shared"].replace("~", "\n  ~ ")[:3000], flush=True)
        for name, file, _, _ in CASES:
            if name in answers:
                continue
            result = excel.run_vba(source, "Probe", [name, str(root / file)], timeout=300.0)
            assert result.ok and not result.dialogs, f"{name}: {result.outcome}: {result.message} {result.error}"
            answers[name] = str(result.value)
            print(name, answers[name][:300], flush=True)
        missing = str(root / "absent.txt")
        result = excel.run_vba(source, "Failing", [missing], timeout=300.0)
        assert result.ok and not result.dialogs, f"missing: {result.outcome}: {result.message} {result.error}"
        # The folder is this machine's; the test puts its own back.
        answers["missing"] = str(result.value).replace(missing, "{path}")
        print("missing", answers["missing"], flush=True)
    OUT.write_text(json.dumps({
        "excel_version_build": version.value, "measured_at": datetime.now(timezone.utc).isoformat(),
        "probe": source, "mixed": MIXED, "defaults": list(DEFAULTS),
        "files": {name: {"file": file, "data": base64.b64encode(data).decode("ascii")} for name, file, data, _ in CASES},
        "sequence": [file for file, _ in SEQUENCE], "shared": SHARED_FILES,
        "shared_text": base64.b64encode(SHARED_TEXT).decode("ascii"), "answers": answers}, indent=2) + "\n",
        encoding="utf-8")


if __name__ == "__main__":
    main()
