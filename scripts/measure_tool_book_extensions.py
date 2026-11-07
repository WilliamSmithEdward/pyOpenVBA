"""Excel's 2010 extensions as another tool writes them, and under other prefixes: sparklines, 2010 data bars and a conditional
format on another sheet's cells, read by Excel and by the model.

    PYTHONPATH=<folder holding xlsxwriter> python scripts/measure_tool_book_extensions.py

writes tests/fixtures/tool_books/extensions.xlsx and extensions.json, which tests/test_excel_tool_book_extensions.py replays.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from measure_tool_books import OUT, model_reads, tally  # noqa: E402

X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
READS: list[str] = [
    'ws.Range("F1").SparklineGroups.Count', 'ws.Range("F1").SparklineGroups(1).Type', 'ws.Range("F1").SparklineGroups(1).SourceData', 'ws.Range("F1").SparklineGroups(1).Location.Address',
    'ws.Range("F1").SparklineGroups(1).Points.Markers.Visible', 'ws.Range("F2").SparklineGroups(1).Type', 'ws.Range("G1").SparklineGroups.Count', 'ws.Cells.SparklineGroups.Count',
    'ws.Range("A1").FormatConditions.Count', 'ws.Range("A1").FormatConditions(1).Type', 'ws.Range("A1").FormatConditions(1).BarFillType', 'ws.Range("A1").FormatConditions(1).BarColor.Color',
    'ws.Range("A1").FormatConditions(1).Direction', 'ws.Range("A1").FormatConditions(1).MinPoint.Type', 'ws.Range("A1").FormatConditions(1).NegativeBarFormat.ColorType', 'ws.Range("A1").FormatConditions(1).BarBorder.Type',
    'ws.Range("A1").FormatConditions(1).AxisPosition', 'ws.Range("A1").Value', 'ws.Range("A3").Value', 'ws.Cells.FormatConditions.Count',
]
#: What is done before a save, both ways, each on the book as written and under other prefixes.
EDITS = {"row_inserted": "ws.Rows(1).Insert", "sparkline_type": 'ws.Range("F1").SparklineGroups(1).Type = 1', "sparkline_cleared": 'ws.Range("F2").SparklineGroups.Clear',
         "sparkline_added": 'ws.Range("G1").SparklineGroups.Add 1, "A3:D3"', "bar_colour": 'ws.Range("A1").FormatConditions(1).BarColor.Color = 255',
         "bar_added": 'ws.Range("C1:C6").FormatConditions.AddDatabar', "condition_deleted": 'ws.Range("A1").FormatConditions(1).Delete'}
ACT = """Sub Act(ByVal file As String)
    Dim wb As Workbook, ws As Worksheet
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file, UpdateLinks:=0)
    Set ws = wb.Worksheets("Data")
    {action}
    wb.Save
    wb.Close False
End Sub
"""
SOURCE = """Private Function Shown(ByRef v As Variant) As String
    If IsObject(v) Then
        Shown = "Object"
    ElseIf IsError(v) Then
        Shown = "Error:" & CStr(v)
    ElseIf IsNull(v) Then
        Shown = "Null"
    ElseIf IsEmpty(v) Then
        Shown = "Empty"
    Else
        Shown = TypeName(v) & ":" & CStr(v)
    End If
End Function
Function Go(ByVal file As String) As String
    Dim wb As Workbook, ws As Worksheet, v As Variant, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file, UpdateLinks:=0)
    Set ws = wb.Worksheets("Data")
{reads}    wb.Close False
    Go = Mid$(s, 2)
End Function
"""
READ_ONE = """    On Error Resume Next
    v = Empty
    v = {read}
    If Err.Number <> 0 Then
        s = s & vbTab & "E" & Err.Number
        Err.Clear
    Else
        On Error GoTo 0
        s = s & vbTab & Shown(v)
    End If
    On Error GoTo 0
"""


def source(reads: list[str]) -> str:
    return SOURCE.format(reads="".join(READ_ONE.format(read=read) for read in reads))


def respelled(data: bytes, old: str, new: str, uri: str) -> bytes:
    """A sheet with a prefix another: every tag, attribute and declaration, and mc:Ignorable's list where it names it."""
    text = data.decode("utf-8")
    if f'xmlns:{old}="{uri}"' not in text:
        return data
    text = re.sub(rf"(</?){old}:", rf"\g<1>{new}:", text)
    text = re.sub(rf'(\s){old}:([\w.-]+=)', rf"\g<1>{new}:\2", text)
    text = text.replace(f'xmlns:{old}="{uri}"', f'xmlns:{new}="{uri}"')
    return re.sub(r'(Ignorable=")([^"]*)(")', lambda m: m.group(1) + " ".join(new if one == old else one for one in m.group(2).split()) + m.group(3), text).encode("utf-8")


VARIANTS = {"as_written": [], "x14_prefixed": [("x14", "q", X14)], "xm_prefixed": [("xm", "w", XM)], "mc_prefixed": [("mc", "m2", MC)],
            "all_prefixed": [("x14", "q", X14), ("xm", "w", XM), ("mc", "m2", MC)]}


def by_xlsxwriter(file: Path) -> None:
    import xlsxwriter

    wb = xlsxwriter.Workbook(str(file))
    ws = wb.add_worksheet("Data")
    for row in range(6):
        ws.write_row(row, 0, [row + 1, 3 - row, row * 2, row % 3])
    ws.conditional_format("A1:A6", {"type": "data_bar", "data_bar_2010": True, "bar_solid": True, "bar_color": "#00B050", "bar_negative_color_same": True})
    ws.add_sparkline("F1", {"range": "Data!A1:D1", "type": "column"})
    ws.add_sparkline("F2", {"range": "Data!A2:D2", "markers": True})
    wb.close()


def made(base: Path, variant: str, file: Path) -> None:
    with zipfile.ZipFile(base) as given, zipfile.ZipFile(file, "w", zipfile.ZIP_DEFLATED) as out:
        for name in given.namelist():
            data = given.read(name)
            if name.startswith("xl/worksheets/sheet"):
                for old, new, uri in VARIANTS[variant]:
                    data = respelled(data, old, new, uri)
            out.writestr(name, data)


def main() -> None:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError
    from pyvbaharness import ExcelSession, HarnessConfig

    base = OUT / "extensions.xlsx"
    by_xlsxwriter(base)
    rows: dict[str, object] = {}
    with tempfile.TemporaryDirectory(prefix="pyopenvba_ext_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        for variant in VARIANTS:
            file = Path(folder) / f"{variant}.xlsx"
            made(base, variant, file)
            read = excel.run_vba(source(READS), "Go", (str(file),), timeout=120.0, line_numbers=False)
            theirs = str(read.value).split("\t") if read.ok else []
            mine: list[str] = []
            for one in READS:
                app = ExcelApplication(sandbox=False)
                app.add_workbook()
                app.add_module(source([one]), name="M")
                try:
                    mine.append(str(app.run("Go", file.as_posix())))
                except VBAUnsupportedError as failure:
                    mine.append(f"reports: {str(failure)[:90]}")
                except Exception as failure:  # noqa: BLE001 - what the model does with such a file is what is measured
                    mine.append(f"crash: {type(failure).__name__}: {str(failure)[:90]}")
            rows[variant] = {"excel": theirs, "model": mine, "opens": "" if read.ok else f"{read.outcome} {read.message}"}
            print(variant, read.outcome, tally(theirs, mine) if theirs else "no reads", flush=True)
            for name, one, other in zip(READS, theirs, mine):
                if one != other:
                    print("  ", name, one[:40], "|", other[:80], flush=True)
    edits: dict[str, object] = {}
    with tempfile.TemporaryDirectory(prefix="pyopenvba_ext_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        for variant in ("as_written", "all_prefixed"):
            for name, action in EDITS.items():
                theirs_file, ours_file = Path(folder) / f"{variant}_{name}_x.xlsx", Path(folder) / f"{variant}_{name}_m.xlsx"
                made(base, variant, theirs_file)
                made(base, variant, ours_file)
                app = ExcelApplication(sandbox=False)
                app.add_workbook()
                app.add_module(ACT.format(action=action), name="M")
                reported = ""
                try:
                    app.run("Act", ours_file.as_posix())
                except VBAUnsupportedError as failure:
                    reported = f"reports: {str(failure)[:100]}"
                except Exception as failure:  # noqa: BLE001 - as above
                    reported = f"crash: {type(failure).__name__}: {str(failure)[:100]}"
                acted = excel.run_vba(ACT.format(action=action), "Act", (str(theirs_file),), timeout=120.0)
                assert acted.ok, (variant, name, acted.outcome, acted.message)
                theirs_read = excel.run_vba(source(READS), "Go", (str(theirs_file),), timeout=120.0, line_numbers=False)
                ours_read = excel.run_vba(source(READS), "Go", (str(ours_file),), timeout=120.0, line_numbers=False) if not reported else None
                entry = {"reported": reported, "excel": str(theirs_read.value).split("\t"),
                         "model": (str(ours_read.value).split("\t") if ours_read.ok else [f"failed {ours_read.outcome}"]) if ours_read is not None else []}
                edits[f"{variant}|{name}"] = entry
                differing = [one for one, a, b in zip(READS, entry["excel"], entry["model"]) if a != b]
                print(variant, name, reported[:90] or ("saved" if len(entry["model"]) == len(READS) else entry["model"]), differing[:5], flush=True)
    record = {"reads": READS, "edits": edits, "variants": rows, "measured_at": datetime.now(timezone.utc).isoformat()}
    (OUT / "extensions.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
