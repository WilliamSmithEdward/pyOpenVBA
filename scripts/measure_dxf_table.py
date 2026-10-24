"""The order Excel writes a workbook's dxfs in, which it leaves out, and how every part that names one follows.

A dxf is a partial format that something names by its place in the
stylesheet's dxfs: a conditional format's rule, a sort on a cell's colour,
an AutoFilter's colour filter, a table's column, a custom table style and
a pivot table's format. This makes a workbook in Excel holding each, its
rules made on two sheets in an order apart from the sheets' and one rule
alike to an earlier one, and saves it. Excel opens the file, adds a rule
and saves it again. Then copies of the first file take a dxf nothing
names at the start, in the middle and at the end, every name moved to
follow, and one where two rules name one dxf; Excel opens each and saves
it with nothing changed. Each file's dxfs and every name of one are kept,
with what the rules and the sort show. Two more workbooks have a sheet with
rules copied, and one rule's cells pasted to another sheet.

    python scripts/measure_dxf_table.py

writes tests/fixtures/dxf_table/: each file, Excel's save of each, and
dxf_table.json, which tests/test_excel_dxf_table.py replays.
``python scripts/measure_dxf_table.py copies`` measures the copy and the
paste alone, into the record the rest wrote.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from collections.abc import Callable
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "dxf_table"

BUILD = '''Private Function Rule(ws As Object, area As String, colour As Long) As String
    Dim f As Object
    On Error Resume Next
    Set f = ws.Range(area).FormatConditions.Add(xlCellValue, xlGreater, "5")
    f.Interior.Color = colour
    Rule = IIf(Err.Number <> 0, "E" & Err.Number, "ok")
End Function

Public Function Build() As String
Dim wb As Object, one As Object, two As Object, three As Object, out As String, r As Long, pv As Object, pt As Object
Dim t As Object, s As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set one = wb.Worksheets(1)
one.Name = "one"
Set two = wb.Worksheets.Add(After:=one)
two.Name = "two"
Set three = wb.Worksheets.Add(After:=two)
three.Name = "three"
one.Range("A1:B3").Value = 9
two.Range("A1:B3").Value = 9
out = out & "r1:" & Rule(two, "A1:A3", RGB(255, 0, 0)) & ";"
out = out & "r2:" & Rule(one, "A1:A3", RGB(0, 0, 255)) & ";"
out = out & "r3:" & Rule(two, "B1:B3", RGB(0, 255, 0)) & ";"
out = out & "r4:" & Rule(one, "B1:B3", RGB(0, 0, 255)) & ";"
On Error Resume Next
three.Range("A1").Value = "Key"
three.Range("B1").Value = "Other"
For r = 2 To 6
    three.Cells(r, 1).Value = r
    three.Cells(r, 2).Value = 7 - r
Next
three.Range("A3").Interior.Color = RGB(255, 255, 0)
three.Range("B4").Interior.Color = RGB(255, 0, 255)
Err.Clear
three.Sort.SortFields.Clear
three.Sort.SortFields.Add(Key:=three.Range("A2:A6"), SortOn:=xlSortOnCellColor, Order:=xlAscending).SortOnValue.Color = RGB(255, 255, 0)
three.Sort.SetRange three.Range("A1:B6")
three.Sort.Header = xlYes
three.Sort.Apply
out = out & "sort:" & IIf(Err.Number <> 0, "E" & Err.Number, "ok") & ";"
Err.Clear
three.Range("A1:B6").AutoFilter Field:=2, Criteria1:=RGB(255, 0, 255), Operator:=xlFilterCellColor
out = out & "filter:" & IIf(Err.Number <> 0, "E" & Err.Number, "ok") & ";"
Err.Clear
one.Range("D1").Value = "Head"
one.Range("E1").Value = "More"
one.Range("D2:E4").Value = 3
Set t = one.ListObjects.Add(xlSrcRange, one.Range("D1:E4"), , xlYes)
t.Name = "Shaded"
t.ListColumns(1).DataBodyRange.Interior.Color = RGB(0, 255, 255)
out = out & "table:" & IIf(Err.Number <> 0, "E" & Err.Number, "ok") & ";"
Err.Clear
Set s = wb.TableStyles.Add("Mine")
s.TableStyleElements(xlWholeTable).Interior.Color = RGB(128, 0, 0)
out = out & "style:" & IIf(Err.Number <> 0, "E" & Err.Number, "ok") & ";"
Err.Clear
Set pv = wb.Worksheets.Add(After:=three)
pv.Name = "pivot"
two.Range("A1").Value = "Name"
two.Range("B1").Value = "Amount"
Set pt = wb.PivotCaches.Create(xlDatabase, "two!R1C1:R3C2").CreatePivotTable(pv.Range("A3"), "Totals")
pt.PivotFields("Name").Orientation = xlRowField
pt.AddDataField pt.PivotFields("Amount"), "Total", xlSum
pt.DataBodyRange.Interior.Color = RGB(0, 128, 0)
out = out & "pivot:" & IIf(Err.Number <> 0, "E" & Err.Number, "ok") & ";"
On Error GoTo 0
wb.SaveAs "{folder}\\made.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\made.xlsx")
out = out & "r5:" & Rule(wb.Worksheets("three"), "D1:D3", RGB(128, 128, 0)) & ";"
wb.SaveAs "{folder}\\made_added.xlsx", 51
wb.Close False
Build = out
End Function
'''

#: A second workbook, its dxfs made in an order apart from the one Excel reads them in: tables on two sheets, a
#: filter and a sort between them, rules on three sheets made out of their sheets' order, a pivot and two table
#: styles. Saved, it is opened and saved again as it is, with a rule's colour changed, and with a rule deleted.
ORDER_BUILD = '''Private Function Rule(ws As Object, area As String, colour As Long) As Object
    Set Rule = ws.Range(area).FormatConditions.Add(xlCellValue, xlGreater, "5")
    Rule.Interior.Color = colour
End Function

Public Function Build() As String
Dim wb As Object, alpha As Object, beta As Object, gamma As Object, pv As Object, pt As Object, t As Object
Dim r As Long, f As Object, out As String
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set alpha = wb.Worksheets(1)
alpha.Name = "alpha"
Set beta = wb.Worksheets.Add(After:=alpha)
beta.Name = "beta"
Set gamma = wb.Worksheets.Add(After:=beta)
gamma.Name = "gamma"
For r = 1 To 6
    alpha.Cells(r, 1).Value = 9
    beta.Cells(r, 1).Value = r
    beta.Cells(r, 2).Value = 7 - r
    beta.Cells(r, 3).Value = 9
    beta.Cells(r, 5).Value = 9
    gamma.Cells(r, 1).Value = 9
Next
beta.Range("A1").Value = "Key"
beta.Range("B1").Value = "Other"
Set f = Rule(gamma, "A1:A3", RGB(255, 128, 0))
Set f = Rule(beta, "C1:C3", RGB(0, 0, 255))
Set f = Rule(alpha, "A1:A3", RGB(255, 0, 0))
Set f = Rule(beta, "E1:E3", RGB(0, 255, 0))
beta.Range("A3").Interior.Color = RGB(255, 255, 0)
beta.Range("B4").Interior.Color = RGB(255, 0, 255)
beta.Sort.SortFields.Clear
beta.Sort.SortFields.Add(Key:=beta.Range("A2:A6"), SortOn:=xlSortOnCellColor, Order:=xlAscending).SortOnValue.Color = RGB(255, 255, 0)
beta.Sort.SetRange beta.Range("A1:B6")
beta.Sort.Header = xlYes
beta.Sort.Apply
beta.Range("A1:B6").AutoFilter Field:=2, Criteria1:=RGB(255, 0, 255), Operator:=xlFilterCellColor
beta.Range("G1").Value = "Head"
beta.Range("H1").Value = "More"
beta.Range("G2:H4").Value = 3
Set t = beta.ListObjects.Add(xlSrcRange, beta.Range("G1:H4"), , xlYes)
t.ListColumns(1).DataBodyRange.Interior.Color = RGB(0, 128, 128)
wb.TableStyles.Add("First").TableStyleElements(xlWholeTable).Interior.Color = RGB(128, 0, 0)
alpha.Range("D1").Value = "Head"
alpha.Range("E1").Value = "More"
alpha.Range("D2:E4").Value = 3
Set t = alpha.ListObjects.Add(xlSrcRange, alpha.Range("D1:E4"), , xlYes)
t.ListColumns(1).DataBodyRange.Interior.Color = RGB(0, 255, 255)
wb.TableStyles.Add("Second").TableStyleElements(xlHeaderRow).Interior.Color = RGB(0, 0, 128)
gamma.Range("C1").Value = "Name"
gamma.Range("D1").Value = "Amount"
gamma.Range("C2:C3").Value = "x"
gamma.Range("D2:D3").Value = 4
Set pv = wb.Worksheets.Add(After:=gamma)
pv.Name = "pivot"
Set pt = wb.PivotCaches.Create(xlDatabase, "gamma!R1C3:R3C4").CreatePivotTable(pv.Range("A3"), "Totals")
pt.PivotFields("Name").Orientation = xlRowField
pt.AddDataField pt.PivotFields("Amount"), "Total", xlSum
pt.DataBodyRange.Interior.Color = RGB(0, 100, 0)
wb.SaveAs "{folder}\\order.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\order.xlsx")
wb.SaveAs "{folder}\\order_again.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\order.xlsx")
wb.Worksheets("alpha").Range("A1").FormatConditions(1).Interior.Color = RGB(255, 192, 203)
wb.SaveAs "{folder}\\order_changed.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\order.xlsx")
wb.Worksheets("gamma").Range("A1").FormatConditions(1).Delete
wb.SaveAs "{folder}\\order_deleted.xlsx", 51
wb.Close False
Build = "ok"
End Function
'''

#: A third workbook, for the order Excel reads one sheet's tables and its filter or sort in, and two pivots made
#: out of their sheets' order: saved, then opened and saved again as it is.
SHEET_BUILD = '''Private Function Pivot(wb As Object, ws As Object, colour As Long) As Object
    Dim pt As Object
    ws.Range("A1").Value = "Name"
    ws.Range("B1").Value = "Amount"
    ws.Range("A2:A3").Value = "x"
    ws.Range("B2:B3").Value = 4
    Set pt = wb.PivotCaches.Create(xlDatabase, ws.Name & "!R1C1:R3C2").CreatePivotTable(ws.Range("D3"))
    pt.PivotFields("Name").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Amount"), "Total", xlSum
    pt.DataBodyRange.Interior.Color = colour
    Set Pivot = pt
End Function

Public Function Build() As String
Dim wb As Object, p As Object, q As Object, r As Object, s As Object, t As Object, row As Long, f As Object
Dim w As Variant
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set p = wb.Worksheets(1)
p.Name = "p"
Set q = wb.Worksheets.Add(After:=p)
q.Name = "q"
Set r = wb.Worksheets.Add(After:=q)
r.Name = "r"
Set s = wb.Worksheets.Add(After:=r)
s.Name = "s"
For row = 1 To 6
    p.Cells(row, 1).Value = row
    p.Cells(row, 2).Value = 7 - row
    q.Cells(row, 1).Value = row
    q.Cells(row, 2).Value = 7 - row
Next
p.Range("A1").Value = "Key"
q.Range("A1").Value = "Key"
For Each w In Array(p, q)
    w.Range("D1").Value = "Head"
    w.Range("E1").Value = "More"
    w.Range("D2:E4").Value = 3
Next
Set t = q.ListObjects.Add(xlSrcRange, q.Range("D1:E4"), , xlYes)
t.ListColumns(1).DataBodyRange.Interior.Color = RGB(0, 128, 128)
Set t = p.ListObjects.Add(xlSrcRange, p.Range("D1:E4"), , xlYes)
t.ListColumns(1).DataBodyRange.Interior.Color = RGB(0, 255, 255)
p.Range("B4").Interior.Color = RGB(255, 0, 255)
p.Range("A1:B6").AutoFilter Field:=2, Criteria1:=RGB(255, 0, 255), Operator:=xlFilterCellColor
q.Range("A3").Interior.Color = RGB(255, 255, 0)
q.Sort.SortFields.Clear
q.Sort.SortFields.Add(Key:=q.Range("A2:A6"), SortOn:=xlSortOnCellColor, Order:=xlAscending).SortOnValue.Color = RGB(255, 255, 0)
q.Sort.SetRange q.Range("A1:B6")
q.Sort.Header = xlYes
q.Sort.Apply
Set f = p.Range("G1:G3").FormatConditions.Add(xlCellValue, xlGreater, "5")
f.Interior.Color = RGB(255, 0, 0)
Set f = Pivot(wb, s, RGB(128, 128, 0))
Set f = Pivot(wb, r, RGB(0, 100, 0))
wb.SaveAs "{folder}\\sheets.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\sheets.xlsx")
wb.SaveAs "{folder}\\sheets_resaved.xlsx", 51
wb.Close False
Build = "ok"
End Function
'''

#: Two more workbooks, for the dxfs of rules a sheet copy and a paste make: rules on two sheets made out of their
#: sheets' order, then one sheet copied after the other, or one rule's cells pasted to the other sheet, and saved.
COPY_BUILD = '''Private Sub Rule(ws As Object, area As String, colour As Long)
    Dim f As Object
    Set f = ws.Range(area).FormatConditions.Add(xlCellValue, xlGreater, "5")
    f.Interior.Color = colour
End Sub

Private Function Made() As Object
Dim wb As Object, one As Object, two As Object
Set wb = Workbooks.Add(xlWBATWorksheet)
Set one = wb.Worksheets(1)
one.Name = "one"
Set two = wb.Worksheets.Add(After:=one)
two.Name = "two"
one.Range("A1:B3").Value = 9
two.Range("A1:B3").Value = 9
Rule one, "A1:A3", RGB(255, 0, 0)
Rule two, "A1:A3", RGB(0, 255, 0)
Rule one, "B1:B3", RGB(0, 0, 255)
Set Made = wb
End Function

Public Function Build() As String
Dim wb As Object
Application.DisplayAlerts = False
Set wb = Made()
wb.Worksheets("one").Copy After:=wb.Worksheets("two")
wb.SaveAs "{folder}\\copied.xlsx", 51
wb.Close False
Set wb = Made()
wb.Worksheets("one").Range("A1:A3").Copy wb.Worksheets("two").Range("D1")
wb.SaveAs "{folder}\\pasted.xlsx", 51
wb.Close False
Build = "ok"
End Function
'''

READ = '''Private Function Shown(ws As Object) As String
    Dim f As Object, out As String
    On Error Resume Next
    For Each f In ws.Cells.FormatConditions
        out = out & f.AppliesTo.Address(False, False) & "@" & f.Priority & "=" & f.Interior.Color & ","
    Next
    If Err.Number <> 0 Then out = out & "E" & Err.Number
    Shown = out
End Function

Public Function Read() As String
Dim wb As Object, out As String, name As Variant, ws As Object
Application.DisplayAlerts = False
For Each name In Array({names})
    On Error Resume Next
    Err.Clear
    Set wb = Nothing
    Set wb = Workbooks.Open("{folder}\\" & name & ".xlsx")
    If wb Is Nothing Then
        out = out & name & "^E" & Err.Number & "|"
    Else
        out = out & name & "^"
        For Each ws In wb.Worksheets
            out = out & ws.Name & ":" & Shown(ws) & ";"
        Next
        Err.Clear
        out = out & "sort=" & wb.Worksheets("three").Sort.SortFields(1).SortOnValue.Color & ";"
        If Err.Number <> 0 Then out = out & "E" & Err.Number & ";"
        Err.Clear
        out = out & "filter=" & wb.Worksheets("three").AutoFilter.Filters(2).Criteria1.Color & ";"
        If Err.Number <> 0 Then out = out & "E" & Err.Number & ";"
        Err.Clear
        out = out & "style=" & wb.TableStyles("Mine").TableStyleElements(xlWholeTable).Interior.Color & "|"
        If Err.Number <> 0 Then out = out & "E" & Err.Number & "|"
        On Error GoTo 0
        wb.SaveAs "{folder}\\" & name & "_resaved.xlsx", 51
        wb.Close False
    End If
    On Error GoTo 0
Next
Read = out
End Function
'''

#: The second workbook's reader, as READ: its rules, the sort's and the filter's colours, both table styles', each
#: table's first column's and the pivot's.
ORDER_READ = READ.split("Public Function Read")[0] + '''Public Function Read() As String
Dim wb As Object, out As String, name As Variant, ws As Object, t As Object
Application.DisplayAlerts = False
For Each name In Array({names})
    On Error Resume Next
    Err.Clear
    Set wb = Nothing
    Set wb = Workbooks.Open("{folder}\\" & name & ".xlsx")
    If wb Is Nothing Then
        out = out & name & "^E" & Err.Number & "|"
    Else
        out = out & name & "^"
        For Each ws In wb.Worksheets
            out = out & ws.Name & ":" & Shown(ws)
            For Each t In ws.ListObjects
                out = out & "table=" & t.ListColumns(1).DataBodyRange.Cells(1, 1).Interior.Color & ","
            Next
            out = out & ";"
        Next
        Err.Clear
        out = out & "sort=" & wb.Worksheets("beta").Sort.SortFields(1).SortOnValue.Color & ";"
        If Err.Number <> 0 Then out = out & "E" & Err.Number & ";"
        Err.Clear
        out = out & "filter=" & wb.Worksheets("beta").AutoFilter.Filters(2).Criteria1.Color & ";"
        If Err.Number <> 0 Then out = out & "E" & Err.Number & ";"
        Err.Clear
        out = out & "styles=" & wb.TableStyles("First").TableStyleElements(xlWholeTable).Interior.Color & "," & _
            wb.TableStyles("Second").TableStyleElements(xlHeaderRow).Interior.Color & ";"
        If Err.Number <> 0 Then out = out & "E" & Err.Number & ";"
        Err.Clear
        out = out & "pivot=" & wb.Worksheets("pivot").PivotTables(1).DataBodyRange.Cells(1, 1).Interior.Color & "|"
        If Err.Number <> 0 Then out = out & "E" & Err.Number & "|"
        On Error GoTo 0
        wb.SaveAs "{folder}\\" & name & "_resaved.xlsx", 51
        wb.Close False
    End If
    On Error GoTo 0
Next
Read = out
End Function
'''

#: An element's start tag, and the attributes that name a dxf by its place, in any part.
_ELEMENT = re.compile(r"<([\w:]+)\b([^>]*)>")
_NAMES = re.compile(r'\b(dxfId|dataDxfId|headerRowDxfId|totalsRowDxfId|headerRowBorderDxfId|tableBorderDxfId|'
                    r'totalsRowBorderDxfId)="(\d+)"')


def _dxfs(styles: str) -> list[str]:
    found = re.search(r"<dxfs\b[^>]*?(?:/>|>(.*?)</dxfs>)", styles, re.DOTALL)
    if found is None or found.group(1) is None:
        return []
    return re.findall(r"<dxf>.*?</dxf>|<dxf/>", found.group(1), re.DOTALL)


def holdings(path: Path) -> dict[str, object]:
    """A file's dxfs, and each part's names of one in the order the part holds them: the tag naming it, the
    attribute, the place, and for a rule its range and priority."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        names: dict[str, list[str]] = {}
        for part in sorted(package.namelist()):
            if not part.endswith(".xml") or part.startswith("xl/theme"):
                continue
            text = package.read(part).decode("utf-8")
            found: list[str] = []
            for element in _ELEMENT.finditer(text):
                tag, attributes = element.groups()
                for name in _NAMES.finditer(attributes):
                    extra = ""
                    if tag == "cfRule":
                        priority = re.search(r'\bpriority="(\d+)"', attributes)
                        area = text.rfind("<conditionalFormatting", 0, element.start())
                        sqref = re.search(r'sqref="([^"]*)"', text[area:element.start()]) if area >= 0 else None
                        extra = f":{sqref.group(1) if sqref else ''}@{priority.group(1) if priority else ''}"
                    found.append(f"{tag}.{name.group(1)}={name.group(2)}{extra}")
            if found:
                names[part] = found
    return {"dxfs": _dxfs(styles), "names": names}


def _renamed(text: str, change: Callable[[int], int]) -> str:
    """``text`` with every name of a dxf moved as ``change`` moves its place."""
    return _ELEMENT.sub(lambda element: _NAMES.sub(
        lambda name: f'{name.group(1)}="{change(int(name.group(2)))}"', element.group(0)), text)


def variants(base: Path) -> dict[str, dict[str, str]]:
    """Each copy's changed parts: a dxf nothing names at the start, in the middle and at the end, every name moved
    to follow; and one where the second rule of sheet "one" names the dxf of its first."""
    with zipfile.ZipFile(base) as package:
        parts = {part: package.read(part).decode("utf-8") for part in package.namelist()
                 if part.endswith(".xml") and not part.startswith("xl/theme")}
    styles = parts["xl/styles.xml"]
    count = len(_dxfs(styles))
    stray = '<dxf><font><color rgb="FF9C0006"/></font></dxf>'

    def with_stray(at: int) -> dict[str, str]:
        entries = _dxfs(styles)
        listed = [*entries[:at], stray, *entries[at:]]
        found = re.search(r'<dxfs count="(\d+)">(.*?)</dxfs>', styles, re.DOTALL)
        assert found is not None
        text = styles[:found.start()] + f'<dxfs count="{len(listed)}">' + "".join(listed) + "</dxfs>" \
            + styles[found.end():]
        out = {part: _renamed(content, lambda place: place + 1 if place >= at else place)
               for part, content in parts.items() if part != "xl/styles.xml"}
        out["xl/styles.xml"] = _renamed(text, lambda place: place + 1 if place >= at else place)
        return {part: content for part, content in out.items() if content != parts.get(part)}

    one = next(part for part, content in parts.items() if part.startswith("xl/worksheets/sheet")
               and content.count("<cfRule") == 2 and 'sqref="A1:A3"' in content and 'sqref="B1:B3"' in content
               and "tableParts" in content)
    rules = re.findall(r'<cfRule\b[^>]*?dxfId="(\d+)"', parts[one])
    shared = re.sub(r'(<cfRule\b[^>]*?dxfId=")(\d+)(")', lambda match: match.group(1) + rules[0] + match.group(3),
                    parts[one])
    return {"unused_first": with_stray(0), "unused_middle": with_stray(count // 2), "unused_last": with_stray(count),
            "shared": {one: shared}}


def order_variants(base: Path) -> dict[str, dict[str, str]]:
    """The second workbook's copies: sheet "beta"'s second rule naming its first rule's dxf, sheet "gamma"'s rule
    naming the sort's, and beta's rules listed in the file the other way round."""
    with zipfile.ZipFile(base) as package:
        parts = {part: package.read(part).decode("utf-8") for part in package.namelist()
                 if part.endswith(".xml") and not part.startswith("xl/theme")}
    beta = next(part for part, content in parts.items() if part.startswith("xl/worksheets/sheet")
                and 'sqref="C1:C3"' in content and 'sqref="E1:E3"' in content)
    gamma = next(part for part, content in parts.items() if part.startswith("xl/worksheets/sheet")
                 and content.count("<cfRule") == 1 and "pivot" not in part and part != beta
                 and "tableParts" not in content)
    blocks = re.findall(r"<conditionalFormatting\b.*?</conditionalFormatting>", parts[beta], re.DOTALL)
    first = re.search(r'<conditionalFormatting sqref="C1:C3">.*?dxfId="(\d+)"', parts[beta], re.DOTALL)
    sort = re.search(r'<sortCondition\b[^>]*?dxfId="(\d+)"', parts[beta])
    assert len(blocks) == 2 and first is not None and sort is not None
    shared = re.sub(r'(<conditionalFormatting sqref="E1:E3">.*?dxfId=")(\d+)(")',
                    lambda match: match.group(1) + first.group(1) + match.group(3), parts[beta], count=1,
                    flags=re.DOTALL)
    cross = re.sub(r'(<cfRule\b[^>]*?dxfId=")(\d+)(")', lambda match: match.group(1) + sort.group(1) + match.group(3),
                   parts[gamma], count=1)
    start = parts[beta].index(blocks[0])
    end = parts[beta].index(blocks[1]) + len(blocks[1])
    swapped = parts[beta][:start] + blocks[1] + blocks[0] + parts[beta][end:]
    return {"order_shared": {beta: shared}, "order_cross": {gamma: cross}, "order_listed": {beta: swapped}}


def write(base: Path, target: Path, changed: dict[str, str]) -> None:
    with zipfile.ZipFile(base) as package, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for info in package.infolist():
            data = changed[info.filename].encode("utf-8") if info.filename in changed else package.read(info.filename)
            out.writestr(info, data)


def copies() -> dict[str, object]:
    """Excel's saves of the workbooks COPY_BUILD makes, and what each holds."""
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(COPY_BUILD.replace("{folder}", str(OUT)), "Build", timeout=600.0)
        assert result.ok and result.value == "ok", f"{result.outcome}: {result.message} {result.error}"
    for name in ("copied", "pasted"):
        strip_save_path(OUT / f"{name}.xlsx")
    return {"copy_build": COPY_BUILD, "copies": {name: holdings(OUT / f"{name}.xlsx") for name in ("copied", "pasted")}}


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(BUILD.replace("{folder}", str(OUT)), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    built = str(result.value)
    for name in ("made", "made_added"):
        strip_save_path(OUT / f"{name}.xlsx")
    made = variants(OUT / "made.xlsx")
    for name, changed in made.items():
        write(OUT / "made.xlsx", OUT / f"{name}.xlsx", changed)
    names = ["made", *made]
    code = READ.replace("{names}", ", ".join(f'"{name}"' for name in names)).replace("{folder}", str(OUT))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Read", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = dict(one.split("^", 1) for one in str(result.value).split("|") if "^" in one)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(ORDER_BUILD.replace("{folder}", str(OUT)), "Build", timeout=600.0)
        assert result.ok and result.value == "ok", f"{result.outcome}: {result.message} {result.error}"
    for name in ("order", "order_again", "order_changed", "order_deleted"):
        strip_save_path(OUT / f"{name}.xlsx")
    ordered = order_variants(OUT / "order.xlsx")
    for name, changed in ordered.items():
        write(OUT / "order.xlsx", OUT / f"{name}.xlsx", changed)
    order_names = ["order", "order_again", "order_changed", "order_deleted", *ordered]
    code = ORDER_READ.replace("{names}", ", ".join(f'"{name}"' for name in order_names)).replace("{folder}", str(OUT))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Read", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers.update(dict(one.split("^", 1) for one in str(result.value).split("|") if "^" in one))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(SHEET_BUILD.replace("{folder}", str(OUT)), "Build", timeout=600.0)
        assert result.ok and result.value == "ok", f"{result.outcome}: {result.message} {result.error}"
    for name in ("sheets", "sheets_resaved"):
        strip_save_path(OUT / f"{name}.xlsx")
    cases: list[dict[str, object]] = [{"name": "sheets", "answer": "", "file": holdings(OUT / "sheets.xlsx"),
                                       "resaved": holdings(OUT / "sheets_resaved.xlsx")}]
    for name in [*names, *order_names]:
        resaved = OUT / f"{name}_resaved.xlsx"
        if resaved.exists():
            strip_save_path(resaved)
        cases.append({"name": name, "answer": answers.get(name, ""), "file": holdings(OUT / f"{name}.xlsx"),
                      "resaved": holdings(resaved) if resaved.exists() else {}})
    record = {"build": BUILD, "built": built, "read": READ, "added": holdings(OUT / "made_added.xlsx"),
              "order_build": ORDER_BUILD, "order_read": ORDER_READ, "sheet_build": SHEET_BUILD, "cases": cases,
              **copies()}
    (OUT / "dxf_table.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print("built:", built)
    print("== made")
    print(json.dumps(cases[0]["file"], indent=1)[:4000])
    print("== made_added")
    print(json.dumps(record["added"], indent=1)[:3000])
    for one in cases:
        print("==", one["name"], ascii(one["answer"])[:300])
        if one["file"] != one["resaved"]:
            print("   file   ", ascii(one["file"])[:1500])
            print("   resaved", ascii(one["resaved"])[:1500])


if __name__ == "__main__":
    if sys.argv[1:] == ["copies"]:
        # Measure the copies alone, into the record the rest of the measurement wrote.
        kept = json.loads((OUT / "dxf_table.json").read_text(encoding="utf-8"))
        kept.update(copies())
        (OUT / "dxf_table.json").write_text(json.dumps(kept, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
        print(json.dumps(kept["copies"], indent=1))
    else:
        main()
