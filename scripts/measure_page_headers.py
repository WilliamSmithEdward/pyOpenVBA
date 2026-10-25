"""PageSetup headers and footers in live Excel: what each property reads, what a save writes, how a file's are read.

Part one makes a workbook for each case, runs the case's steps on its sheet's PageSetup, an error's number and
description noted where a step raises one, and reads every header and footer property, the four header flags, and
the first and even pages' sections; then it saves the workbook. Part two opens workbooks whose headers the
measurement wrote into a file Excel saved, reads the same, and saves each twice: as it is, and after a step, most
often its CenterHeader set to itself. Each file's headerFooter element is kept. Last, Excel tries to open files
whose header is longer than a macro may set.

    python scripts/measure_page_headers.py

writes tests/fixtures/page_headers.json, which tests/test_excel_page_headers.py replays.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "page_headers.json"
PS = "ps"

HELPER = '''Private Function Shown(text As String) As String
    Shown = Replace(Replace(text, vbCr, "<CR>"), vbLf, "<LF>")
End Function

Private Function Sections(page As Object) As String
    Dim f As String
    f = ChrW(166)
    Sections = Shown(page.LeftHeader.Text) & f & Shown(page.CenterHeader.Text) & f & Shown(page.RightHeader.Text) & _
        f & Shown(page.LeftFooter.Text) & f & Shown(page.CenterFooter.Text) & f & Shown(page.RightFooter.Text)
End Function

Private Function Dump(ps As Object) As String
    Dim f As String
    f = ChrW(166)
    Dump = Shown(ps.LeftHeader) & f & Shown(ps.CenterHeader) & f & Shown(ps.RightHeader) & f & _
        Shown(ps.LeftFooter) & f & Shown(ps.CenterFooter) & f & Shown(ps.RightFooter) & f & _
        ps.DifferentFirstPageHeaderFooter & f & ps.OddAndEvenPagesHeaderFooter & f & _
        ps.ScaleWithDocHeaderFooter & f & ps.AlignMarginsHeaderFooter & f & _
        Sections(ps.FirstPage) & f & Sections(ps.EvenPage)
End Function

Private Function Longest(ps As PageSetup, member As String, first As Long, last As Long) As String
    Dim n As Long, kept As Long
    kept = -1
    On Error Resume Next
    For n = first To last
        Err.Clear
        Select Case member
            Case "CenterHeader": ps.CenterHeader = String(n, "a")
            Case "LeftHeader": ps.LeftHeader = String(n, "a")
            Case "RightFooter": ps.RightFooter = String(n, "a")
        End Select
        If Err.Number <> 0 Then Exit For
        kept = n
    Next
    Longest = "longest " & member & "=" & kept & " then E" & Err.Number & ":" & Err.Description & " "
    Err.Clear
End Function
'''

#: name -> the steps, run on ``ps``, the PageSetup of a new workbook's one sheet ``ws``.
CASES: dict[str, list[str]] = {
    "default": [],
    "each": ['ps.LeftHeader = "L"', 'ps.CenterHeader = "C"', 'ps.RightHeader = "R"', 'ps.LeftFooter = "l"',
             'ps.CenterFooter = "c"', 'ps.RightFooter = "r"'],
    "center_only": ['ps.CenterHeader = "Report"'],
    "right_only": ['ps.RightHeader = "R"'],
    "left_right": ['ps.LeftHeader = "a"', 'ps.RightHeader = "b"'],
    "set_order": ['ps.RightHeader = "r"', 'ps.LeftHeader = "l"', 'ps.CenterHeader = "c"'],
    "footer_only": ['ps.CenterFooter = "Page &P of &N"'],
    "codes": ['ps.CenterHeader = "Page &P of &N"', 'ps.LeftFooter = "&D &T"', 'ps.RightFooter = "&F &A &Z"'],
    "font": ['ps.CenterHeader = "&""Arial,Bold""&14Title&B x&I y&U&S&X&Y&K00FF00z"'],
    "embedded_left": ['ps.CenterHeader = "a&Lb"'],
    "embedded_after": ['ps.CenterHeader = "c"', 'ps.LeftHeader = "x&Cy"'],
    "embedded_right": ['ps.LeftHeader = "l"', 'ps.CenterHeader = "c&Rr"'],
    "trailing_amp": ['ps.CenterHeader = "a&"', 'ps.RightHeader = "b"'],
    "double_amp": ['ps.CenterHeader = "Tom && Jerry"'],
    "single_amp": ['ps.CenterHeader = "Tom & Jerry"'],
    "lower_codes": ['ps.CenterHeader = "a&lb&cc&rd"'],
    "newlines": ['ps.CenterHeader = "a" & vbLf & "b"', 'ps.LeftHeader = "c" & vbCrLf & "d"',
                 'ps.RightHeader = "e" & vbCr & "f"'],
    "xml_text": ['ps.CenterHeader = "<a> ""q"" \'x\' " & ChrW(&H263A) & ChrW(233)'],
    "clear": ['ps.CenterHeader = "x"', 'ps.CenterHeader = ""'],
    "clear_all": ['ps.LeftHeader = "L"', 'ps.CenterFooter = "c"', 'ps.LeftHeader = ""', 'ps.CenterFooter = ""'],
    "limit_center": ['notes = notes & Longest(ps, "CenterHeader", 245, 260)'],
    "limit_left": ['notes = notes & Longest(ps, "LeftHeader", 245, 260)'],
    "limit_footer": ['notes = notes & Longest(ps, "RightFooter", 245, 260)'],
    "limit_shared": ['ps.LeftHeader = String(100, "b")', 'notes = notes & Longest(ps, "CenterHeader", 140, 160)'],
    "limit_other_side": ['ps.LeftFooter = String(200, "b")', 'notes = notes & Longest(ps, "CenterHeader", 245, 260)'],
    "too_long_kept": ['ps.CenterHeader = "keep"', 'ps.CenterHeader = String(300, "a")'],
    "replace_long": ['ps.LeftHeader = String(200, "a")', 'ps.LeftHeader = String(200, "b")'],
    "two_lefts": ['ps.CenterHeader = "a&Lb&Lc"', 'notes = notes & "[" & ps.LeftHeader & "] "', 'ps.LeftHeader = "z"'],
    "coerced": ['ps.CenterHeader = 5', 'ps.LeftHeader = True', 'ps.RightHeader = 1.5', 'ps.LeftFooter = Null',
                'ps.CenterFooter = Empty', 'ps.RightFooter = #1/2/2024#'],
    "flags": ['ps.DifferentFirstPageHeaderFooter = True', 'ps.OddAndEvenPagesHeaderFooter = True',
              'ps.ScaleWithDocHeaderFooter = False', 'ps.AlignMarginsHeaderFooter = False'],
    "flags_with_text": ['ps.CenterHeader = "x"', 'ps.ScaleWithDocHeaderFooter = False'],
    "first_page": ['ps.DifferentFirstPageHeaderFooter = True', 'ps.FirstPage.CenterHeader.Text = "F1"',
                   'ps.FirstPage.LeftFooter.Text = "f"', 'ps.CenterHeader = "O"',
                   'notes = notes & TypeName(ps.FirstPage) & ":" & TypeName(ps.FirstPage.CenterHeader) & " "'],
    "first_page_flag_off": ['ps.FirstPage.CenterHeader.Text = "F1"'],
    "first_then_off": ['ps.DifferentFirstPageHeaderFooter = True', 'ps.FirstPage.CenterHeader.Text = "F1"',
                       'ps.DifferentFirstPageHeaderFooter = False'],
    "even_page": ['ps.OddAndEvenPagesHeaderFooter = True', 'ps.EvenPage.RightHeader.Text = "E"',
                  'ps.CenterHeader = "O"'],
    "even_page_flag_off": ['ps.EvenPage.RightHeader.Text = "E"'],
    "page_code_text": ['ps.FirstPage.CenterHeader.Text = "a&Rb"', 'ps.DifferentFirstPageHeaderFooter = True'],
    "page_too_long": ['ps.FirstPage.CenterHeader.Text = "keep"', 'ps.FirstPage.CenterHeader.Text = String(300, "a")'],
    "copied": ['ps.CenterHeader = "c&Rr"', 'ps.LeftFooter = "f"', 'ps.DifferentFirstPageHeaderFooter = True',
               'ps.FirstPage.CenterHeader.Text = "F"', 'ps.ScaleWithDocHeaderFooter = False', 'ws.Copy After:=ws',
               'notes = notes & "[copy " & Dump(wb.Worksheets(2).PageSetup) & "] "'],
    "flag_coerced": ['ps.DifferentFirstPageHeaderFooter = 1', 'ps.OddAndEvenPagesHeaderFooter = "True"',
                     'ps.ScaleWithDocHeaderFooter = 0', 'ps.AlignMarginsHeaderFooter = Null'],
}

REASSIGN = "ps.CenterHeader = ps.CenterHeader"
#: name -> a headerFooter element's attributes and children, written into a file Excel saved, and the step run on
#: it once it is saved as it is.
CRAFTED: dict[str, tuple[str, dict[str, str], str]] = {
    "reordered": ("", {"oddHeader": "&Rr&Ll&Cc"}, REASSIGN),
    "repeated": ("", {"oddHeader": "&La&Cb&Lc"}, REASSIGN),
    "repeated_left_set": ("", {"oddHeader": "&La&Cb&Lc"}, 'ps.LeftHeader = "z"'),
    "left_emptied": ("", {"oddHeader": "&La&Cb"}, 'ps.LeftHeader = ""'),
    "no_codes": ("", {"oddHeader": "abc"}, REASSIGN),
    "text_before": ("", {"oddHeader": "abc&Rr"}, REASSIGN),
    "text_before_right_set": ("", {"oddHeader": "abc&Rr"}, 'ps.RightHeader = "q"'),
    "prefix_and_center": ("", {"oddHeader": "abc&Cdef"}, REASSIGN),
    "double_amp": ("", {"oddHeader": "&Ca&&Rb"}, REASSIGN),
    "escape_at_start": ("", {"oddHeader": "&C&&Rb"}, REASSIGN),
    "trailing_amp": ("", {"oddHeader": "&Ca&"}, REASSIGN),
    "empty_sections": ("", {"oddHeader": "&L&C&R"}, REASSIGN),
    "code_at_start": ("", {"oddHeader": "&L&Cx"}, REASSIGN),
    "lower_codes": ("", {"oddHeader": "&lx&cy&rz"}, REASSIGN),
    "font": ("", {"oddHeader": '&L&"Arial,Bold"&14x&Cy'}, REASSIGN),
    "line_breaks": ("", {"oddHeader": "&Ca\r\nb&Lc_x000D_d"}, REASSIGN),
    "footers": ("", {"oddFooter": "&RPage &P", "evenFooter": "&Ceven", "firstFooter": "&Lfirst"}, REASSIGN),
    "first_unflagged": ("", {"oddHeader": "&Codd", "firstHeader": "&Cfirst"}, REASSIGN),
    "first_flagged": (' differentFirst="1"', {"oddHeader": "&Codd", "firstHeader": "&Cfirst"}, REASSIGN),
    "even_flagged": (' differentOddEven="1"', {"oddHeader": "&Codd", "evenHeader": "&Ceven"}, REASSIGN),
    "even_set": (' differentOddEven="1"', {"evenHeader": "&Ceven"}, 'ps.EvenPage.CenterHeader.Text = "E2"'),
    "flags_off": (' scaleWithDoc="0" alignWithMargins="0"', {"oddHeader": "&Cx"}, REASSIGN),
    "longest": ("", {"oddHeader": "&C" + "a" * 253}, REASSIGN),
    "amp_prefix": ("", {"oddHeader": "&&&Lx"}, REASSIGN),
    "amp_prefix_text": ("", {"oddHeader": "&&x&Ly"}, REASSIGN),
    "lone_amp": ("", {"oddHeader": "&"}, REASSIGN),
    "code_only": ("", {"oddHeader": "&C"}, 'ps.LeftHeader = "l"'),
}
#: Headers longer than Excel opens, each tried in a call of its own: the length of the header's text.
TOO_LONG = {"too_long_256": 256, "too_long_300": 300}
CHILDREN = ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter")
LOOK = '''Public Function Look(folder As String, file As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(folder & "\\" & file)
    Look = "E" & Err.Number & ":" & Err.Description
    If Not wb Is Nothing Then wb.Close False
End Function
'''


def vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def case_code(index: int, name: str, steps: list[str]) -> str:
    lines = [f"Private Function Make{index}(folder As String) As String",
             "    Dim wb As Workbook, ws As Worksheet, ps As PageSetup, notes As String",
             "    Set wb = Workbooks.Add(xlWBATWorksheet)", "    Set ws = wb.Worksheets(1)", "    Set ps = ws.PageSetup",
             "    On Error Resume Next"]
    for step in steps:
        lines += ["    Err.Clear", f"    {step}",
                  '    If Err.Number <> 0 Then notes = notes & "E" & Err.Number & ":" & Err.Description & " "']
    lines += ["    On Error GoTo 0", f'    Make{index} = "{name}^" & notes & "^" & Dump(ps) & "~"',
              f'    wb.SaveAs folder & "\\{name}.xlsx", 51', "    wb.Close False", "End Function"]
    return "\n".join(lines) + "\n"


def opened_code(index: int, name: str, step: str) -> str:
    return "\n".join([
        f"Private Function Open{index}(folder As String) As String",
        "    Dim wb As Workbook, ps As PageSetup, notes As String",
        f'    Set wb = Workbooks.Open(folder & "\\crafted_{name}.xlsx")',
        "    Set ps = wb.Worksheets(1).PageSetup",
        f'    Open{index} = "{name}^^" & Dump(ps) & "~"',
        f'    wb.SaveAs folder & "\\resaved_{name}.xlsx", 51',
        "    On Error Resume Next",
        f"    {step}",
        '    If Err.Number <> 0 Then notes = "E" & Err.Number & ":" & Err.Description',
        "    On Error GoTo 0",
        f'    Open{index} = Open{index} & "{name}_then^" & notes & "^" & Dump(ps) & "~"',
        f'    wb.SaveAs folder & "\\then_{name}.xlsx", 51',
        "    wb.Close False",
        "End Function"]) + "\n"


def program(functions: str, entry: str, count: int, prefix: str) -> str:
    calls = "".join(f"    out = out & {prefix}{index}(folder)\n" for index in range(count))
    return (HELPER + functions + f"Public Function {entry}(folder As String) As String\n    Dim out As String\n"
            "    Application.DisplayAlerts = False\n" + calls + f"    {entry} = out\nEnd Function\n")


def escaped(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def crafted(base: Path, target: Path, attributes: str, children: dict[str, str]) -> None:
    """``base`` with a headerFooter element of ``attributes`` and ``children`` where the schema puts it: after the
    sheet's pageSetup, or its pageMargins where it has none."""
    inner = "".join(f"<{name}>{escaped(children[name])}</{name}>" for name in CHILDREN if name in children)
    element = f"<headerFooter{attributes}>{inner}</headerFooter>"
    with zipfile.ZipFile(base) as source, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                found = re.search(r"<pageSetup\b[^>]*/>", text) or re.search(r"<pageMargins\b[^>]*/>", text)
                assert found is not None and "<headerFooter" not in text
                text = text[:found.end()] + element + text[found.end():]
                data = text.encode("utf-8")
            out.writestr(info, data)


def header_footer(path: Path) -> list[str | None]:
    """Each sheet's headerFooter element as the file has it, None where a sheet has none."""
    found: list[str | None] = []
    with zipfile.ZipFile(path) as package:
        for index in range(1, 10):
            part = f"xl/worksheets/sheet{index}.xml"
            if part not in package.namelist():
                break
            sheet = package.read(part).decode("utf-8")
            element = re.search(r"<headerFooter\b[^>]*?(?:/>|>.*?</headerFooter>)", sheet, re.S)
            found.append(element.group(0) if element else None)
    return found


def records(text: str) -> list[list[str]]:
    out = []
    for entry in text.split("~"):
        if entry:
            name, notes, dump = entry.split("^")
            out.append([name, notes, *dump.split("¦")])
    return out


def main() -> None:
    build = program("".join(case_code(index, name, steps) for index, (name, steps) in enumerate(CASES.items())),
                    "Build", len(CASES), "Make")
    opened = program("".join(opened_code(index, name, step) for index, (name, (_a, _c, step))
                             in enumerate(CRAFTED.items())), "Opened", len(CRAFTED), "Open")
    refused: dict[str, str] = {}
    with tempfile.TemporaryDirectory(prefix="pyopenvba-headers-") as name:
        folder = Path(name).resolve()
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            made = excel.run_vba(build, "Build", [str(folder)], timeout=900.0)
            assert made.ok, f"{made.outcome}: {made.message} {made.error}"
            for key, (attributes, children, _step) in CRAFTED.items():
                crafted(folder / "default.xlsx", folder / f"crafted_{key}.xlsx", attributes, children)
            read = excel.run_vba(opened, "Opened", [str(folder)], timeout=600.0)
            assert read.ok, f"{read.outcome}: {read.message} {read.error}"
            for key, length in TOO_LONG.items():
                crafted(folder / "default.xlsx", folder / f"{key}.xlsx", "", {"oddHeader": "&C" + "a" * (length - 2)})
                looked = excel.run_vba(LOOK, "Look", [str(folder), f"{key}.xlsx"], timeout=120.0)
                assert looked.ok, f"{looked.outcome}: {looked.message} {looked.error}"
                refused[key] = str(looked.value)
        files = {path.stem: header_footer(path) for path in sorted(folder.glob("*.xlsx"))
                 if not path.stem.startswith("too_long")}
    record = {"helper": HELPER, "build": build, "opened": opened, "look": LOOK,
              "crafted": {key: [attributes, children, step] for key, (attributes, children, step) in CRAFTED.items()},
              "too_long": {key: [length, refused[key]] for key, length in TOO_LONG.items()},
              "cases": records(str(made.value)) + records(str(read.value)), "files": files}
    OUT.write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for case in record["cases"]:
        print(ascii(case))
    for key, seen in files.items():
        print(key, ascii(seen))
    print(record["too_long"])


if __name__ == "__main__":
    main()
