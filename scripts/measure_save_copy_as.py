"""Workbook.SaveCopyAs in live Excel: what it writes, and what it leaves alone.

A workbook saved as named.xlsx, with a hyperlink to a file beside it, is
changed and copied, and each copy records the workbook's Saved flag before
and after, whether its Name and FullName stayed, the save events the
workbook and the application raised, and any error's number and
description; a copy made two seconds after the last records whether the
workbook's Last save time property moved. The copies go to a new name, a
name without an extension, a .bak and a .xlsm name, a missing folder, the
workbook's own path, another open workbook's path, the same name in
another folder, an empty name, Null, a relative name, a named argument,
and a file that already exists, once with alerts off and once with them
on. A copy under manual calculation records the cell the workbook holds
before and after. A workbook never saved, with a hyperlink too, a
macro-enabled one and the harness's own document are copied as well, and
each hyperlink's address is read before and after a copy to another folder.

Afterwards every file the probe wrote is read: its workbook content type,
whether it holds a VBA project, its first sheet's A1 and B1, its hyperlink
target, its absPath, and for the first copy the parts that differ from the
Save that followed it. The harness's own copy is read and never kept.

    python scripts/measure_save_copy_as.py

writes tests/fixtures/save_copy_as.json, which tests/test_excel_save_copy_as.py replays.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "save_copy_as.json"

LOG = "Public Log As String\n"

SINK = '''Public WithEvents Book As Workbook
Public WithEvents App As Application
Private Sub Book_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Log = Log & "book-before:" & SaveAsUI & "|"
End Sub
Private Sub Book_AfterSave(ByVal Success As Boolean)
    Log = Log & "book-after:" & Success & "|"
End Sub
Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Log = Log & "app-before:" & SaveAsUI & "|"
End Sub
Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    Log = Log & "app-after:" & Success & "|"
End Sub
'''

#: Each copy's record: name ^ Saved before > after ^ Name kept ^ FullName kept ^ events ^ error ^ Last save time.
#: A timed copy waits two seconds first, so its Last save time reads same, moved or none (unreadable) for certain;
#: an untimed one records no time.
PROBE = '''Private Sink As CopySink

Private Function Copy(label As String, wb As Workbook, target As Variant, folder As String, _
                      Optional timed As Boolean) As String
    Dim was As Boolean, n As String, f As String, e As String, stamp As Variant, moved As String
    was = wb.Saved
    n = wb.Name
    f = wb.FullName
    On Error Resume Next
    If timed Then
        stamp = wb.BuiltinDocumentProperties("Last save time").Value
        If Err.Number <> 0 Then stamp = Empty
        Err.Clear
        Application.Wait Now + TimeSerial(0, 0, 2)
    End If
    Log = ""
    wb.SaveCopyAs target
    e = Err.Number & ":" & Err.Description
    Err.Clear
    If timed Then
        moved = "none"
        If IsEmpty(stamp) Then
            If Not IsEmpty(wb.BuiltinDocumentProperties("Last save time").Value) Then moved = "set"
        ElseIf wb.BuiltinDocumentProperties("Last save time").Value = stamp Then
            moved = "same"
        ElseIf wb.BuiltinDocumentProperties("Last save time").Value > stamp Then
            moved = "moved"
        End If
        If Err.Number <> 0 Then moved = "none"
        Err.Clear
    End If
    On Error GoTo 0
    Copy = label & "^" & was & ">" & wb.Saved & "^" & (wb.Name = n) & "^" & (wb.FullName = f) & "^" & Log & "^" & _
        Replace(e, folder, "<folder>") & "^" & moved & "~"
End Function

Private Function Linked(wb As Workbook, folder As String) As String
    Linked = Replace(wb.Worksheets(1).Hyperlinks(1).Address, folder, "<folder>")
End Function

Public Function Probe(folder As String) As String
    Dim wb As Workbook, other As Workbook, fresh As Workbook, out As String, e As String
    Set Sink = New CopySink
    ChDrive Left(folder, 1)
    ChDir folder
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    out = "settings^" & Application.CalculateBeforeSave & "~"
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Range("A1").Value = 1
    wb.Worksheets(1).Range("B1").Formula = "=A1*2"
    wb.Worksheets(1).Hyperlinks.Add wb.Worksheets(1).Range("A3"), folder & "\\target.txt"
    wb.SaveAs folder & "\\named.xlsx", 51
    Set Sink.Book = wb
    Set Sink.App = Application
    Application.EnableEvents = True
    wb.Worksheets(1).Range("A1").Value = 5
    out = out & Copy("dirty", wb, folder & "\\copy1.xlsx", folder, True)
    Log = ""
    wb.Save
    out = out & "save^" & Log & "~"
    out = out & Copy("clean", wb, folder & "\\copy2.xlsx", folder, True)
    wb.Worksheets(1).Range("A1").Value = 6
    out = out & Copy("no_extension", wb, folder & "\\noext", folder)
    out = out & Copy("bak", wb, folder & "\\copy.bak", folder)
    out = out & Copy("macro_name", wb, folder & "\\as_macro.xlsm", folder)
    out = out & Copy("missing_folder", wb, folder & "\\nope\\missing.xlsx", folder, True)
    out = out & Copy("own_path", wb, wb.FullName, folder)
    Application.EnableEvents = False
    Set other = Workbooks.Add(xlWBATWorksheet)
    other.SaveAs folder & "\\other.xlsx", 51
    Application.EnableEvents = True
    out = out & Copy("open_path", wb, folder & "\\other.xlsx", folder)
    Application.EnableEvents = False
    other.Close False
    Application.EnableEvents = True
    MkDir folder & "\\sub"
    out = out & "link_before^" & Linked(wb, folder) & "~"
    out = out & Copy("same_name_elsewhere", wb, folder & "\\sub\\named.xlsx", folder)
    out = out & "link_after^" & Linked(wb, folder) & "~"
    out = out & Copy("empty_name", wb, "", folder, True)
    out = out & Copy("null_name", wb, Null, folder)
    out = out & Copy("relative", wb, "relative.xlsx", folder)
    out = out & Copy("quiet_first", wb, folder & "\\over.xlsx", folder)
    wb.Worksheets(1).Range("A1").Value = 8
    out = out & Copy("quiet_again", wb, folder & "\\over.xlsx", folder)
    On Error Resume Next
    wb.SaveCopyAs Filename:=folder & "\\named_argument.xlsx"
    e = Err.Number & ":" & Err.Description
    On Error GoTo 0
    out = out & "named_argument^" & Replace(e, folder, "<folder>") & "~"
    Application.Calculation = xlCalculationManual
    wb.Worksheets(1).Range("A1").Value = 7
    out = out & "manual_before^" & wb.Worksheets(1).Range("B1").Value & "~"
    out = out & Copy("manual", wb, folder & "\\manual.xlsx", folder)
    out = out & "manual_after^" & wb.Worksheets(1).Range("B1").Value & "~"
    wb.Worksheets(1).Range("A1").Value = 11
    out = out & Copy("manual_empty", wb, "", folder)
    out = out & "manual_empty_after^" & wb.Worksheets(1).Range("B1").Value & "~"
    Application.Calculation = xlCalculationAutomatic
    wb.BuiltinDocumentProperties("Last author").Value = "Someone else"
    out = out & Copy("author_empty", wb, "", folder)
    out = out & "author_empty_after^" & (wb.BuiltinDocumentProperties("Last author").Value = "Someone else") & "~"
    wb.BuiltinDocumentProperties("Last author").Value = "Someone else"
    out = out & Copy("author", wb, folder & "\\author.xlsx", folder)
    out = out & "author_after^" & (wb.BuiltinDocumentProperties("Last author").Value = "Someone else") & "~"
    Application.EnableEvents = False
    Set fresh = Workbooks.Add(xlWBATWorksheet)
    fresh.Worksheets(1).Range("A1").Value = 3
    fresh.Worksheets(1).Hyperlinks.Add fresh.Worksheets(1).Range("A3"), folder & "\\target.txt"
    Set Sink.Book = fresh
    Application.EnableEvents = True
    out = out & "fresh_format^" & fresh.FileFormat & "^" & (fresh.Path = "") & "~"
    out = out & "fresh_link_before^" & Linked(fresh, folder) & "~"
    out = out & Copy("fresh", fresh, folder & "\\sub\\fresh.xlsx", folder, True)
    out = out & "fresh_link_after^" & Linked(fresh, folder) & "~"
    out = out & Copy("fresh_macro_name", fresh, folder & "\\fresh_macro.xlsm", folder, True)
    out = out & "fresh_after^" & (fresh.Path = "") & "^" & (Left(fresh.Name, 4) = "Book") & "~"
    Application.EnableEvents = False
    fresh.Close False
    Set fresh = Workbooks.Add(xlWBATWorksheet)
    fresh.SaveAs folder & "\\macro.xlsm", 52
    Set Sink.Book = fresh
    Application.EnableEvents = True
    fresh.Worksheets(1).Range("A1").Value = 4
    out = out & Copy("macro_book", fresh, folder & "\\macro_copy.xlsx", folder)
    Application.EnableEvents = False
    fresh.Close False
    Set Sink.Book = ThisWorkbook
    Application.EnableEvents = True
    out = out & "harness_format^" & ThisWorkbook.FileFormat & "~"
    out = out & Copy("harness", ThisWorkbook, folder & "\\harness_copy.xlsm", folder)
    Application.EnableEvents = False
    Probe = out
End Function
'''

#: The overwrite with alerts on, in a call of its own: a prompt would stop the call, not the record before it.
LOUD = '''Public Function Loud(folder As String) As String
    Dim wb As Workbook, e As String
    Set wb = Workbooks("named.xlsx")
    wb.Worksheets(1).Range("A1").Value = 9
    Application.DisplayAlerts = True
    On Error Resume Next
    wb.SaveCopyAs folder & "\\over.xlsx"
    e = Err.Number & ":" & Err.Description
    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.Close False
    Loud = "loud^" & Replace(e, folder, "<folder>") & "~"
End Function
'''

_OVERRIDE = re.compile(r'<Override\b[^>]*PartName="/xl/workbook\.xml"[^>]*ContentType="([^"]+)"')
_CELL = re.compile(r'<c r="([AB]1)"[^>]*?(?:/>|>(.*?)</c>)', re.S)
_VALUE = re.compile(r'<v>(.*?)</v>', re.S)
_ABS = re.compile(r'<x15ac:absPath\b[^>]*\burl="([^"]*)"')
_MODIFIED = re.compile(r'<dcterms:modified\b[^>]*>[^<]*</dcterms:modified>')
_LINK = re.compile(r'<Relationship\b[^>]*Type="[^"]*/hyperlink"[^>]*Target="([^"]*)"')
_LINK_RELS = "xl/worksheets/_rels/sheet1.xml.rels"


def records(text: str) -> list[list[str]]:
    return [entry.split("^") for entry in text.split("~") if entry]


def anonymous(text: str, folders: list[str]) -> str:
    """The temporary folder, as Excel spells it in a path or a file URL, written <folder>."""
    for folder in folders:
        text = text.replace(folder, "<folder>").replace(folder.replace("\\", "/"), "<folder>")
    return text


def read_file(path: Path, folders: list[str]) -> dict[str, object]:
    """What a written file holds: its workbook content type, VBA project, A1 and B1, hyperlink target and absPath."""
    if not zipfile.is_zipfile(path):
        return {"zip": False, "size": path.stat().st_size}
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        types = package.read("[Content_Types].xml").decode("utf-8")
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        book = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read(_LINK_RELS).decode("utf-8") if _LINK_RELS in names else ""
        core = package.read("docProps/core.xml").decode("utf-8")
    found = _OVERRIDE.search(types)
    cells: dict[str, str | None] = {}
    for match in _CELL.finditer(sheet):
        value = _VALUE.search(match.group(2) or "")
        cells[match.group(1)] = value.group(1) if value else None
    absolute = _ABS.search(book)
    link = _LINK.search(rels)
    return {"zip": True, "content_type": found.group(1) if found else None, "vba": "xl/vbaProject.bin" in names,
            "cells": cells, "link": anonymous(link.group(1), folders) if link else None,
            "abs_path": anonymous(absolute.group(1), folders) if absolute else None,
            "last_author_someone_else": "<cp:lastModifiedBy>Someone else</cp:lastModifiedBy>" in core}


def differing_parts(first: Path, second: Path) -> list[str]:
    """The parts whose bytes differ, docProps/core.xml compared without its modified time."""
    with zipfile.ZipFile(first) as one, zipfile.ZipFile(second) as two:
        names = sorted(set(one.namelist()) | set(two.namelist()))
        out = []
        for name in names:
            if name not in one.namelist() or name not in two.namelist():
                out.append(name)
                continue
            left, right = one.read(name), two.read(name)
            if name == "docProps/core.xml":
                left = _MODIFIED.sub("", left.decode("utf-8")).encode("utf-8")
                right = _MODIFIED.sub("", right.decode("utf-8")).encode("utf-8")
            if left != right:
                out.append(name)
    return out


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="pyopenvba-copy-") as name:
        folder = Path(name).resolve()
        folders = sorted({str(folder), name}, key=len, reverse=True)
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            excel.add_module("CopyLog", LOG)
            excel.add_module("CopySink", SINK, kind="class")
            probe = excel.run_vba(PROBE, "Probe", [str(folder)], timeout=300.0)
            assert probe.ok, f"{probe.outcome}: {probe.message} {probe.error}"
            loud = excel.run_vba(LOUD, "Loud", [str(folder)], timeout=60.0)
            loud_record = records(str(loud.value)) if loud.ok else [["loud", f"stopped: {loud.outcome}"]]
        text = anonymous(str(probe.value), folders)
        files = {path.relative_to(folder).as_posix(): read_file(path, folders)
                 for path in sorted(folder.rglob("*")) if path.is_file()}
        compared = differing_parts(folder / "copy1.xlsx", folder / "named.xlsx")
    record = {"log": LOG, "sink": SINK, "probe": PROBE, "loud": LOUD, "cases": records(text) + loud_record,
              "files": files, "copy1_differs_from_save": compared}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for case in record["cases"]:
        print(case)
    for file, seen in files.items():
        print(file, seen)
    print("copy1 differs from the save in", compared)


if __name__ == "__main__":
    main()
