"""Members of Workbook, Workbooks, Sheets, Worksheet and PageSetup that a macro meets with no screen, server or mail
system. What each answers or raises, on a new workbook and on one saved to a folder.

It leaves out what would print, send mail, open a browser or wait for someone: PrintOut, PrintPreview, SendMail,
FollowHyperlink, ShowDataForm and the like. Machine details are recorded as a yes or no.

    python scripts/measure_workbook_members.py

writes tests/fixtures/workbook_members.json, which tests/test_excel_workbook_members.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/workbook_members.json'

COLOURS = [f'r = wb.Theme.ThemeColorScheme.Colors({i}).RGB' for i in range(1, 13)]
CASES = [
    # Pivot tables of the whole workbook.
    'r = TypeName(wb.PivotTables)', 'r = wb.PivotTables.Count', 'Set x = wb.PivotTables(1)', 'r = TypeName(wb.PivotTables("x"))',
    # The theme.
    'r = TypeName(wb.Theme)', 'r = TypeName(wb.Theme.ThemeColorScheme)', 'r = wb.Theme.ThemeColorScheme.Count', 'r = TypeName(wb.Theme.ThemeColorScheme.Colors(1))', *COLOURS,
    'r = wb.Theme.ThemeColorScheme.Colors(0).RGB', 'r = wb.Theme.ThemeColorScheme.Colors(13).RGB', 'r = wb.Theme.ThemeColorScheme.Colors("x").RGB',
    'wb.Theme.ThemeColorScheme.Colors(5).RGB = 255: r = wb.Theme.ThemeColorScheme.Colors(5).RGB', 'wb.Theme.ThemeColorScheme.Colors(5).RGB = 255: ws.Range("A1").Interior.ThemeColor = 5: r = ws.Range("A1").Interior.Color',
    'r = TypeName(wb.Theme.ThemeFontScheme)', 'r = TypeName(wb.Theme.ThemeFontScheme.MajorFont)', 'r = wb.Theme.ThemeFontScheme.MajorFont.Count', 'r = wb.Theme.ThemeFontScheme.MajorFont(1).Name', 'r = wb.Theme.ThemeFontScheme.MinorFont(1).Name', 'r = wb.Theme.ThemeFontScheme.MajorFont(2).Name', 'r = wb.Theme.ThemeFontScheme.MajorFont(3).Name', 'r = wb.Theme.ThemeFontScheme.MajorFont(4).Name',
    'r = TypeName(wb.Theme.ThemeFontScheme.MajorFont(1))', 'wb.Theme.ThemeFontScheme.MinorFont(1).Name = "Arial": r = wb.Theme.ThemeFontScheme.MinorFont(1).Name & "/" & ws.Range("A1").Font.Name',
    'r = TypeName(wb.Theme.ThemeEffectScheme)', 'r = TypeName(wb.Theme.Parent)', 'r = TypeName(wb.Theme.Application)', 'wb.ApplyTheme ""', 'wb.ApplyTheme "nothing.thmx"', 'wb.ApplyTheme',
    # Custom XML.
    'r = TypeName(wb.CustomXMLParts)', 'r = wb.CustomXMLParts.Count', 'r = wb.CustomXMLParts(1).BuiltIn', 'r = wb.CustomXMLParts(1).NamespaceURI', 'r = wb.CustomXMLParts(2).NamespaceURI', 'r = wb.CustomXMLParts(3).NamespaceURI', 'r = wb.CustomXMLParts(4).NamespaceURI', 'r = TypeName(wb.CustomXMLParts(1))',
    'r = TypeName(wb.CustomXMLParts.Add("<a xmlns=""urn:x""><b>1</b></a>"))', 'wb.CustomXMLParts.Add "<a xmlns=""urn:x""><b>1</b></a>": r = wb.CustomXMLParts.Count', 'r = wb.CustomXMLParts.Add("<a xmlns=""urn:x""><b>1</b></a>").XML', 'r = wb.CustomXMLParts.Add("<a xmlns=""urn:x""/>").NamespaceURI', 'r = wb.CustomXMLParts.Add("<a/>").NamespaceURI', 'r = wb.CustomXMLParts.Add("<a/>").BuiltIn', 'r = wb.CustomXMLParts.Add().Count', 'r = wb.CustomXMLParts.Add("not xml").XML',
    'wb.CustomXMLParts.Add("<a xmlns=""urn:x""/>").Delete: r = wb.CustomXMLParts.Count', 'wb.CustomXMLParts(1).Delete', 'wb.CustomXMLParts.Add "<a xmlns=""urn:x""/>": r = wb.CustomXMLParts.SelectByNamespace("urn:x").Count', 'r = wb.CustomXMLParts.SelectByNamespace("urn:y").Count', 'wb.CustomXMLParts.Add "<a xmlns=""urn:x""/>": r = TypeName(wb.CustomXMLParts.SelectByID(wb.CustomXMLParts(4).ID))', 'r = wb.CustomXMLParts(4).ID <> ""', 'r = TypeName(wb.CustomXMLParts.SelectByID("nope"))',
    'r = wb.CustomXMLParts.Add("<a><b>1</b></a>").SelectSingleNode("/a/b").Text', 'r = wb.CustomXMLParts.Add("<a><b>1</b></a>").DocumentElement.BaseName',
    # Windows, access and server state.
    'r = TypeName(wb.NewWindow)', 'wb.NewWindow: r = wb.Windows.Count & "/" & wb.Windows(1).Caption & "/" & wb.Windows(2).Caption', 'r = wb.NewWindow.WindowNumber', 'r = TypeName(wb.WebOptions)', 'r = wb.WebOptions.Encoding', 'r = wb.WebOptions.PixelsPerInch', 'r = wb.[_CodeName]', 'wb.[_CodeName] = "Mine": r = wb.CodeName',
    'wb.ChangeFileAccess xlReadOnly', 'saved.ChangeFileAccess xlReadOnly: r = saved.ReadOnly', 'saved.ChangeFileAccess xlReadWrite: r = saved.ReadOnly', 'saved.ChangeFileAccess 5', 'saved.ChangeFileAccess Mode:=xlReadOnly, Notify:=False: r = saved.ReadOnly', 'saved.ChangeFileAccess', 'r = saved.ChangeFileAccess(xlReadOnly)',
    'r = wb.WriteReservedBy = ""', 'r = saved.WriteReservedBy = ""', 'r = IsEmpty(wb.UserStatus)', 'r = TypeName(wb.UserStatus)', 'r = UBound(wb.UserStatus, 1) & "/" & UBound(wb.UserStatus, 2)', 'r = wb.UserStatus(1, 1) = Application.UserName', 'r = wb.UserStatus(1, 3)',
    'r = TypeName(wb.Signatures)', 'r = wb.Signatures.Count', 'r = TypeName(wb.ServerViewableItems)', 'r = wb.ServerViewableItems.Count', 'r = TypeName(wb.DocumentInspectors)', 'r = wb.DocumentInspectors.Count > 0', 'r = TypeName(wb.Model)', 'r = wb.Model.ModelTables.Count', 'r = TypeName(wb.WorkIdentity)', 'r = TypeName(wb.SensitivityLabel)', 'r = TypeName(wb.SmartDocument)', 'r = TypeName(wb.CommandBars)',
    # Shared workbooks and review, with nothing shared.
    'wb.AcceptAllChanges', 'wb.RejectAllChanges', 'wb.EndReview', 'wb.Reply', 'wb.ReplyAll', 'wb.ReplyWithChanges', 'wb.UnprotectSharing', 'wb.RemoveUser 1', 'wb.PurgeChangeHistoryNow 30', 'wb.HighlightChangesOptions', 'wb.Sync', 'wb.ReloadAs', 'wb.LockServerFile', 'wb.CheckIn', 'wb.CheckInWithVersion', 'wb.ToggleFormsDesign', 'r = TypeName(wb.GetWorkflowTasks)', 'r = TypeName(wb.GetWorkflowTemplates)', 'wb.AddToFavorites', 'wb.Mailer', 'wb.UpdateFromFile', 'saved.UpdateFromFile', 'wb.Post', 'wb.Research',
    # Sheets and pages.
    'r = TypeName(wb.Sheets.HPageBreaks)', 'r = wb.Sheets.HPageBreaks.Count', 'r = wb.Sheets.VPageBreaks.Count', 'ws.Rows(5).PageBreak = xlPageBreakManual: r = wb.Sheets.HPageBreaks.Count', 'r = TypeName(wb.Worksheets.HPageBreaks)', 'wb.Worksheets.Add: wb.Sheets.Select: r = wb.Sheets.VPageBreaks.Count',
    'r = TypeName(ws.PageSetup.Pages)', 'r = ws.PageSetup.Pages.Count', 'r = TypeName(ws.PageSetup.Pages(1))', 'r = ws.PageSetup.Pages(1).CenterHeader.Text', 'r = TypeName(ws.PageSetup.Pages(1).CenterHeader)', 'r = TypeName(ws.PageSetup.LeftHeaderPicture)', 'r = ws.PageSetup.LeftHeaderPicture.Filename', 'r = TypeName(ws.PageSetup.RightFooterPicture)', 'r = ws.PrintedCommentPages', 'r = TypeName(ws.PrintedCommentPages)',
    'r = TypeName(ws.XmlDataQuery("/a"))', 'r = TypeName(ws.XmlMapQuery("/a"))', 'r = TypeName(ws.MailEnvelope)',
    'For Each x In Workbooks: i = i + 1: Next: r = i > 0', 'For Each x In wb.Sheets: i = i + 1: Next: r = i', 'Set x = wb.Sheets.[_NewEnum]: r = TypeName(x)', 'r = Workbooks.CanCheckOut("nope.xlsx")',
    # Scenarios.
    'For Each x In ws.Scenarios: i = i + 1: Next: r = i', 'ws.Scenarios.Add "A", ws.Range("A1"), Array(5): For Each x In ws.Scenarios: r = x.Name: Next', 'ws.Scenarios.Merge ws', 'ws.Scenarios.Merge "Nope"', 'ws.Scenarios.Merge',
    'other.Range("A1").Value = 3: other.Scenarios.Add "B", other.Range("A1"), Array(7), "Why": ws.Scenarios.Merge other: r = ws.Scenarios.Count & "/" & ws.Scenarios(1).Name & "/" & ws.Scenarios(1).Values(1) & "/" & ws.Scenarios(1).ChangingCells.Address & "/" & ws.Scenarios(1).ChangingCells.Parent.Name & "/" & ws.Scenarios(1).Comment & "/" & other.Scenarios.Count',
    'other.Scenarios.Add "B", other.Range("A1"), Array(7): ws.Scenarios.Merge "Other": r = ws.Scenarios.Count & "/" & ws.Scenarios(1).Name',
    'other.Scenarios.Add "B", other.Range("A1"), Array(7): ws.Scenarios.Add "B", ws.Range("A2"), Array(1): ws.Scenarios.Merge other: r = ws.Scenarios.Count & "/" & ws.Scenarios(1).Name & "/" & ws.Scenarios(2).Name & "/" & ws.Scenarios(2).Values(1)',
    'other.Scenarios.Add "B", other.Range("A1"), Array(7): r = ws.Scenarios.Merge(other)', 'r = ws.Scenarios.Merge(other)', 'other.Scenarios.Add "B", other.Range("A1"), Array(7): saved.Worksheets(1).Scenarios.Merge other: r = saved.Worksheets(1).Scenarios.Count & "/" & saved.Worksheets(1).Scenarios(1).ChangingCells.Address',
    'other.Scenarios.Add "B", other.Range("A1"), Array(7): other.Scenarios.Add "C", other.Range("A1:A2"), Array(1, 2), , True, True: ws.Scenarios.Merge other: r = ws.Scenarios.Count & "/" & ws.Scenarios(2).Name & "/" & ws.Scenarios(2).Locked & "/" & ws.Scenarios(2).Hidden & "/" & ws.Scenarios(2).ChangingCells.Address',
    'ws.Scenarios.CreateSummary',
    # And the rest.
    'r = wb.Sheets.HPageBreaks(1).Location.Address', 'r = TypeName(wb.Sheets.HPageBreaks.Item(1))', 'r = wb.Sheets.HPageBreaks("x").Type', 'r = wb.Sheets.VPageBreaks(1).Type', 'r = wb.Sheets.HPageBreaks.Add(ws.Range("A5")).Type', 'r = TypeName(ws.PageSetup.Pages("x"))', 'r = ws.PageSetup.Pages.Item(1).Type', 'r = TypeName(ws.PageSetup.Pages.Parent)',
    'saved.ChangeFileAccess xlReadOnly: saved.ChangeFileAccess xlReadWrite: r = saved.ReadOnly', 'saved.Worksheets(1).Range("A1").Value = 1: saved.ChangeFileAccess xlReadOnly: r = saved.Worksheets(1).Range("A1").Value & "/" & saved.ReadOnly & "/" & saved.Saved', 'saved.Worksheets(1).Range("A1").Value = 1: saved.ChangeFileAccess xlReadOnly: saved.ChangeFileAccess xlReadWrite: r = saved.Worksheets(1).Range("A1").Value & "/" & saved.ReadOnly & "/" & saved.Saved',
    'saved.ChangeFileAccess "x"', 'saved.ChangeFileAccess 3', 'saved.ChangeFileAccess Mode:=xlReadOnly, WritePassword:="p": r = saved.ReadOnly',
    'ws.Range("A1:B1").Value = Array("K", "V"): ws.Range("A2:B2").Value = Array("a", 1): Set x = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B2")): x.CreatePivotTable other.Range("A1"), "P1": r = wb.PivotTables.Count & "/" & other.PivotTables.Count & "/" & wb.PivotTables(1).Name & "/" & TypeName(wb.PivotTables("P1"))',
    'r = wb.Theme.ThemeColorScheme.Colors(msoThemeAccent1).RGB', 'r = wb.Theme.ThemeColorScheme.Colors(1.6).RGB', 'wb.Theme.ThemeColorScheme.Colors(5) = 255', 'r = wb.Theme.ThemeColorScheme.Colors(5)', 'wb.Theme.ThemeColorScheme.Colors(5).RGB = -1', 'wb.Theme.ThemeColorScheme.Colors(5).RGB = "x"', 'wb.Theme.ThemeColorScheme.Colors(5).RGB = 16777216', 'r = wb.Theme.ThemeColorScheme.Colors(5).Parent Is Nothing', 'r = TypeName(wb.Theme.ThemeColorScheme.Parent)',
    'wb.Theme.ThemeFontScheme.MinorFont(1).Name = "": r = wb.Theme.ThemeFontScheme.MinorFont(1).Name', 'wb.Theme.ThemeFontScheme.MajorFont(2).Name = "Arial": r = wb.Theme.ThemeFontScheme.MajorFont(2).Name', 'r = wb.Theme.ThemeFontScheme.MinorFont(0).Name', 'r = TypeName(wb.Theme.ThemeFontScheme.MinorFont.Item(1))', 'r = TypeName(wb.Theme.ThemeFontScheme.Parent)', 'wb.Theme.ThemeFontScheme.MinorFont(1).Name = 5: r = wb.Theme.ThemeFontScheme.MinorFont(1).Name',
    'saved.Theme.ThemeColorScheme.Colors(5).RGB = 255: saved.Save: Set x = saved: n = saved.FullName: saved.Close False: Set saved = Workbooks.Open(n): r = saved.Theme.ThemeColorScheme.Colors(5).RGB',
    'saved.Theme.ThemeFontScheme.MinorFont(1).Name = "Arial": saved.Save: n = saved.FullName: saved.Close False: Set saved = Workbooks.Open(n): r = saved.Theme.ThemeFontScheme.MinorFont(1).Name & "/" & saved.Worksheets(1).Range("A1").Font.Name',
    'r = wb.WorkIdentity = ""', 'r = TypeName(wb.UserStatus(1))', 'x2 = wb.UserStatus: r = TypeName(x2(1, 1)) & "/" & TypeName(x2(1, 2)) & "/" & TypeName(x2(1, 3)) & "/" & x2(1, 3) & "/" & LBound(x2, 1) & "/" & LBound(x2, 2) & "/" & (x2(1, 1) = Application.UserName)',
    'x2 = saved.UserStatus: r = x2(1, 3)', 'r = TypeName(wb.WebOptions.Parent)', 'r = wb.WebOptions.RelyOnCSS & "/" & wb.WebOptions.OrganizeInFolder & "/" & wb.WebOptions.UseLongFileNames & "/" & wb.WebOptions.DownloadComponents & "/" & wb.WebOptions.RelyOnVML & "/" & wb.WebOptions.AllowPNG & "/" & wb.WebOptions.ScreenSize & "/" & wb.WebOptions.TargetBrowser & "/" & wb.WebOptions.LocationOfComponents & "/" & wb.WebOptions.FolderSuffix',
    'r = TypeName(wb.Model.Parent)', 'r = TypeName(wb.Model.ModelTables)', 'r = wb.Model.ModelRelationships.Count', 'r = TypeName(wb.Signatures.Parent)', 'r = TypeName(wb.Signatures(1))', 'r = wb.DocumentInspectors.Count', 'r = wb.CustomXMLParts.Count & "/" & TypeName(wb.CustomXMLParts.Parent)', 'Set x = wb.CustomXMLParts.Add("<a xmlns=""urn:x"" />"): r = x.XML', 'Set x = wb.CustomXMLParts(1): r = Len(x.XML) > 0', 'Set x = wb.CustomXMLParts.Add("<a/>"): x.Delete: r = wb.CustomXMLParts.Count: r = r & "/" & x.BuiltIn',
    'Set x = wb.CustomXMLParts.Add("<a/>"): wb.CustomXMLParts.Add "<b/>": r = wb.CustomXMLParts(4).XML & "/" & wb.CustomXMLParts(5).XML', 'r = wb.CustomXMLParts.Item(1).BuiltIn', 'For Each x In wb.CustomXMLParts: i = i + 1: Next: r = i', 'r = wb.CustomXMLParts("x").BuiltIn', 'r = wb.CustomXMLParts(0).BuiltIn',
    'wb.CustomXMLParts.Add "<a xmlns=""urn:x""/>": r = wb.CustomXMLParts.SelectByNamespace("urn:X").Count', 'r = wb.CustomXMLParts.Add("").BuiltIn', 'r = TypeName(wb.CustomXMLParts.SelectByNamespace("urn:y"))', 'wb.CustomXMLParts.Add "<a xmlns=""urn:x""/>": wb.CustomXMLParts.Add "<b xmlns=""urn:x""/>": r = wb.CustomXMLParts.SelectByNamespace("urn:x")(2).XML',
    'saved.CustomXMLParts.Add "<a xmlns=""urn:x""><b>1</b></a>": saved.Save: n = saved.FullName: saved.Close False: Set saved = Workbooks.Open(n): r = saved.CustomXMLParts.Count & "/" & saved.CustomXMLParts(4).XML',
    'r = TypeName(ws.PageSetup.CenterHeaderPicture.Parent)', 'ws.PageSetup.LeftHeaderPicture.Filename = "nope.png"', 'r = ws.PageSetup.LeftHeaderPicture.Height', 'r = ws.PageSetup.LeftHeaderPicture.LockAspectRatio',
]

MODULE = r'''Option Explicit

Private Function Shown(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        Shown = CStr(v)
    ElseIf IsArray(v) Then
        Shown = "array"
    Else
        Shown = TypeName(v) & ":" & CStr(v)
    End If
End Function

Public Function Tried(ByVal which As Long, ByVal folder As String) As String
    Dim wb As Object, ws As Object, other As Object, saved As Object, r As Variant, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Set other = wb.Worksheets.Add(After:=ws)
    other.Name = "Other"
    Set saved = Workbooks.Add(xlWBATWorksheet)
    saved.SaveAs folder & "\Saved" & which & ".xlsx", xlOpenXMLWorkbook
    ws.Activate
    On Error GoTo Failed
    Run which, wb, ws, other, saved, r
    s = "ok " & Shown(r)
Answered:
    On Error Resume Next
    Application.DisplayAlerts = False
    Tried = s
    saved.Close False
    wb.Close False
    Exit Function
Failed:
    s = "Err " & Err.Number & " " & Err.Description
    Resume Answered
End Function
'''


def _source() -> str:
    lines = ['', 'Private Sub Run(ByVal which As Long, wb As Object, ws As Object, other As Object, saved As Object, r As Variant)',
             '    Dim x As Object, i As Long, n As String, x2 As Variant', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '']
    return MODULE + '\n'.join(lines)


def main() -> None:
    """Each case alone, in a session that a dialog or a broken run ends."""
    source = _source()
    answers: list[str] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
        index = 0
        while index < len(CASES):
            with ExcelSession(HarnessConfig(exclusive=False)) as excel:
                excel.new_document()
                while index < len(CASES):
                    result = excel.run_vba(source, 'Tried', (index, folder), timeout=25.0)
                    index += 1
                    if not result.ok or result.dialogs:
                        text = '; '.join(getattr(dialog, 'message', '') for dialog in result.dialogs)
                        answers.append(f'breaks: {result.outcome} {text}')
                        print(index - 1, CASES[index - 1], '=>', answers[-1], flush=True)
                        break
                    answers.append(re.sub(re.escape(folder), '<folder>', str(result.value), flags=re.IGNORECASE))
                    print(index - 1, CASES[index - 1], '=>', answers[-1], flush=True)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
