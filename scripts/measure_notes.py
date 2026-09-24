"""How Excel saves the notes a macro adds, which VBA calls comments.

A workbook of its own gets notes of every shape a macro makes: one per
cell of a sheet called Data, text with line feeds, a carriage return,
markup, spaces, nothing and a lot, a note shown, one edited and one
deleted, beside a wide column and a tall row, in a merged cell, and at the
sheet's last column and row; a sheet whose rows and columns move under
its notes; and a sheet whose notes share their VML part with a form
control. Each note is read back -- its cell, text, whether it shows, its
shape's name and where the shape is -- and the workbook is saved, its
sheets, comments and VML parts kept.

    python scripts/measure_notes.py

writes tests/fixtures/notes.json and notes.xlsx, which
tests/test_excel_notes.py replays.
"""

from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import copy_saved

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "notes.json"
WORKBOOK = ROOT / "tests" / "fixtures" / "notes.xlsx"

#: The writes, sheet by sheet, each run with ws the sheet.
SHEETS: dict[str, str] = {
    "Data": """ws.Range("B2").AddComment "Hello"
ws.Range("A1").AddComment "Top left"
ws.Range("D5").AddComment "Line1" & vbLf & "Line2"
ws.Range("C3").AddComment "<b>&amp;" & Chr(34) & "'"
ws.Range("F10").AddComment "Shown"
ws.Range("F10").Comment.Visible = True
ws.Range("E4").AddComment ""
ws.Range("H2").AddComment "  spaced  "
ws.Range("G7").AddComment "Edited"
ws.Range("G7").Comment.Text "XY", 3, False
ws.Range("J2").AddComment "gone"
ws.Range("J2").Comment.Delete
ws.Range("K2").AddComment "after delete"
ws.Columns("L").ColumnWidth = 30
ws.Range("L3").AddComment "wide column"
ws.Rows(20).RowHeight = 40
ws.Range("M20").AddComment "tall row"
ws.Range("XFD3").AddComment "last column"
ws.Range("N2:O3").Merge
ws.Range("N2").AddComment "merged"
ws.Range("A1048576").AddComment "last row"
ws.Range("P2").AddComment "Line1" & vbCrLf & "Line2"
ws.Range("Q2").AddComment "a" & vbCr & "b"
ws.Range("R2").AddComment Replace(Space(80), " ", "abc ")
ws.Range("S2").NoteText "via NoteText"
ws.Range("T2").Value = 5
ws.Range("T2").AddComment "on a value"
""",
    "Moved": """ws.Range("B2").AddComment "moved"
ws.Range("D10").AddComment "deleted"
ws.Rows(1).Insert
ws.Columns(1).Insert
ws.Rows(11).Delete
ws.Range("C3").Copy ws.Range("F5")
""",
    "Controls": """ws.Range("B2").AddComment "before the control"
ws.Shapes.AddFormControl 0, 150, 60, 80, 20
ws.Range("C5").AddComment "after the control"
""",
    # Boxes whose style runs 69 to 72 characters before its z-index, which is where Excel's VML breaks a line;
    # and boxes by the sheet's last rows and columns, which do not fit where a box goes.
    "Edges": """ws.Range("M10").AddComment "69"
ws.Range("X10").AddComment "70"
ws.Range("X100").AddComment "71"
ws.Range("X1000").AddComment "72"
ws.Range("A1048573").AddComment "fits"
ws.Range("A1048574").AddComment "over by 9"
ws.Range("A1048575").AddComment "over by 29"
ws.Range("XFA3").AddComment "fits"
ws.Range("XFB3").AddComment "over by 31"
ws.Range("XFC3").AddComment "over by 95"
ws.Range("XFD1048576").AddComment "corner"
""",
}
#: What separates one note's answers, one note from the next, a sheet's name from its notes and one sheet from
#: the next: characters no note holds, as the harness turns control characters into spaces.
_FIELD, _NOTE, _NAMED, _SHEET = "|", "~", "`", "^"


def _module(path: Path) -> str:
    lines = ["Public Function Report() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    for index, (name, writes) in enumerate(SHEETS.items()):
        lines.append("Set ws = wb.Worksheets(1)" if not index else
                     "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))")
        lines += [f'ws.Name = "{name}"', writes]
    # What each sheet's notes read as written, and again once the saved file is opened.
    lines += ["out = Application.UserName & \"" + _SHEET + "\" & Notes(wb)",
              f'wb.SaveAs Filename:="{path}", FileFormat:=51', "wb.Close False",
              f'Set wb = Workbooks.Open("{path}")', 'out = out & "' + _SHEET * 2 + '" & Notes(wb)', "wb.Close False",
              "Report = out", "End Function", "",
              "Private Function Notes(wb As Object) As String",
              "Dim ws As Object, c As Object, out As String",
              "For Each ws In wb.Worksheets", f'    out = out & ws.Name & "{_NAMED}"', "    For Each c In ws.Comments",
              # A note's carriage returns and line feeds are spelled out, \r and \n, so the harness passes them on.
              f'        out = out & c.Parent.Address(False, False) & "{_FIELD}" & '
              r'Replace(Replace(c.Text, vbCr, "\r"), vbLf, "\n")'
              f' & "{_FIELD}" & c.Visible & "{_FIELD}" & c.Shape.Name & "{_FIELD}" & c.Shape.Left & "{_FIELD}" '
              f'& c.Shape.Top & "{_FIELD}" & c.Shape.Width & "{_FIELD}" & c.Shape.Height & "{_FIELD}" '
              f'& (c.Author = Application.UserName) & "{_NOTE}"', "    Next", f'    out = out & "{_SHEET}"', "Next",
              "Notes = out", "End Function"]
    return "\n".join(lines) + "\n"


def _reads(reported: str) -> tuple[str, dict[str, list[dict[str, str]]], dict[str, list[dict[str, str]]]]:
    """The author, and each sheet's notes as the collection lists them, as written and as the saved file opens."""
    author, _, rest = reported.partition(_SHEET)
    written, _, opened = rest.partition(_SHEET * 3)
    return author, _sheet_notes(written), _sheet_notes(opened)


def _sheet_notes(reported: str) -> dict[str, list[dict[str, str]]]:
    sheets = reported.split(_SHEET)
    found: dict[str, list[dict[str, str]]] = {}
    for sheet in sheets:
        if not sheet:
            continue
        name, _, notes = sheet.partition(_NAMED)
        found[name] = []
        for note in notes.split(_NOTE)[:-1]:
            cell, text, visible, shape, left, top, width, height, by_user = note.split(_FIELD)
            found[name].append({"cell": cell, "text": text.replace(r"\r", "\r").replace(r"\n", "\n"),
                                "visible": visible, "shape": shape, "left": left, "top": top, "width": width,
                                "height": height, "by_user": by_user})
    return found


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        path = Path(folder) / "notes.xlsx"
        excel.new_document()
        result = excel.run_vba(_module(path), "Report", timeout=300.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        author, notes, opened = _reads(str(result.value))
        # A workbook of its own, not the probe's, so the file carries no macro; the model opens it.
        copy_saved(path, WORKBOOK)
        with zipfile.ZipFile(path) as package:
            parts = {name: package.read(name).decode("utf-8") for name in package.namelist()
                     if name.startswith(("xl/worksheets/", "xl/comments", "xl/drawings/", "xl/ctrlProps/"))
                     or name == "[Content_Types].xml"}
    OUT.write_text(json.dumps({"author": author, "sheets": SHEETS, "notes": notes, "opened": opened, "parts": parts},
                              indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
