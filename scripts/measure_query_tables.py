"""Query tables over local text files: what each member reads and takes, what a refresh puts on the sheet, and the parts
Excel saves.

Each case makes a workbook with one sheet and runs its statements, which set r to what they want recorded and have q, a
query table over d.csv, once Q has been run. It records r, the sheet's cells and the parts a query table touches. The
files live in a temporary folder, spelled <folder> in what is recorded.

    python scripts/measure_query_tables.py

writes tests/fixtures/query_tables.json, which tests/test_excel_query_tables.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/query_tables.json'
FILES = {'d.csv': 'a,b,c\r\n1,2,3\r\n4,x,6\r\n', 'd.txt': 'a\tb\tc\r\n1\t2\t3\r\n4\tx\t6\r\n', 's.txt': 'a;b;c\r\n1;2;3\r\n', 'o.txt': 'a|b|c\r\n1|2|3\r\n', 'q.csv': '"a,1",\'b\',c\r\n1,"2",3\r\n', 'sp.txt': 'a b  c\r\n1 2  3\r\n',
         'n.csv': 'd,m,t\r\n1/2/2020,1.5-,001\r\n3/4/2021,"1,000",002\r\n', 'w.txt': 'abcdef\r\n123456\r\n', 'e.csv': '', 'long.csv': 'x\r\n1\r\n2\r\n3\r\n4\r\n5\r\n6\r\n', 'short.csv': 'x\r\n1\r\n'}
Q = 'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", ws.Range("A1")): '
R = 'q.TextFileCommaDelimiter = True: q.Refresh False: '
CASES: list[str] = []

# The collection.
CASES += ['r = ws.QueryTables.Count', 'r = TypeName(ws.QueryTables) & " " & TypeName(ws.QueryTables.Parent) & " " & TypeName(ws.QueryTables.Application)', f'{Q}r = ws.QueryTables.Count & " " & TypeName(q) & " " & q.Name & " " & TypeName(ws.QueryTables(1)) & " " & ws.QueryTables(1).Name',
          f'{Q}r = TypeName(ws.QueryTables(q.Name)) & " " & TypeName(ws.QueryTables.Item(1))', f'{Q}r = TypeName(ws.QueryTables(2))', f'{Q}r = TypeName(ws.QueryTables("nope"))', f'{Q}n = 0: For Each o In ws.QueryTables: n = n + 1: t = TypeName(o): Next: r = n & " " & t',
          'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", ws.Range("A1"), "SELECT"): r = q.CommandText', 'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", ws.Range("C5:D9")): r = q.Destination.Address',
          f'{Q}Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.txt", ws.Range("F1")): r = ws.QueryTables.Count & " " & q.Name & " " & ws.QueryTables(1).Name & " " & ws.QueryTables(2).Name', 'Set q = ws.QueryTables.Add("TEXT;", ws.Range("A1")): r = q.Connection', 'Set q = ws.QueryTables.Add("nonsense", ws.Range("A1")): r = q.Connection', 'Set q = ws.QueryTables.Add("", ws.Range("A1")): r = q.Connection',
          'Set q = ws.QueryTables.Add("URL;http://example.invalid/", ws.Range("A1")): r = q.QueryType', 'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", CVar("A1")): r = q.Name', 'Set q = ws.QueryTables.Add(Connection:="TEXT;" & folder & "\\d.csv", Destination:=ws.Range("A1")): r = q.Name', 'Set o = Worksheets.Add: Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", o.Range("A1")): r = q.Name']

# Each member fresh, and after a refresh.
READS_OF_ONE = ('AdjustColumnWidth', 'BackgroundQuery', 'CommandText', 'CommandType', 'Connection', 'Destination.Address', 'EnableEditing', 'EnableRefresh', 'FetchedRowOverflow', 'FieldNames', 'FillAdjacentFormulas', 'MaintainConnection', 'Name', 'PostText', 'PreserveColumnInfo', 'PreserveFormatting', 'QueryType', 'Refreshing', 'RefreshOnFileOpen', 'RefreshPeriod', 'RefreshStyle',
                'ResultRange.Address', 'RobustConnect', 'RowNumbers', 'SaveData', 'SavePassword', 'SourceConnectionFile', 'SourceDataFile', 'TextFileConsecutiveDelimiter', 'TextFileCommaDelimiter', 'TextFileDecimalSeparator', 'TextFileOtherDelimiter', 'TextFileParseType', 'TextFilePlatform', 'TextFilePromptOnRefresh', 'TextFileSemicolonDelimiter', 'TextFileSpaceDelimiter', 'TextFileStartRow', 'TextFileTabDelimiter',
                'TextFileTextQualifier', 'TextFileThousandsSeparator', 'TextFileTrailingMinusNumbers', 'TextFileVisualLayout', 'WebConsecutiveDelimitersAsOne', 'WebDisableDateRecognition', 'WebDisableRedirections', 'WebFormatting', 'WebPreFormattedTextToColumns', 'WebSelectionType', 'WebSingleBlockTextImport', 'WebTables')
TYPES_OF_ONE = ('Parent', 'Application', 'Destination', 'ListObject', 'Parameters', 'Recordset', 'ResultRange', 'Sort', 'WorkbookConnection')
ARRAYS = ('TextFileColumnDataTypes', 'TextFileFixedColumnWidths')
CASES += [f'{Q}r = q.{name}' for name in READS_OF_ONE]
CASES += [f'{Q}{R}r = q.{name}' for name in READS_OF_ONE]
CASES += [f'{Q}r = TypeName(q.{name})' for name in TYPES_OF_ONE]
CASES += [f'{Q}{R}r = TypeName(q.{name})' for name in TYPES_OF_ONE]
CASES += [f'{Q}r = TypeName(q.{name})' for name in ARRAYS]

# Refreshing.
CASES += [f'{Q}r = q.Refresh(False)', f'{Q}r = q.Refresh(True)', f'{Q}r = q.Refresh', f'{Q}{R}r = q.Refresh(False)', f'{Q}{R}{R}r = q.ResultRange.Address & " " & q.Name', f'{Q}q.AdjustColumnWidth = False: {R}r = ws.Columns(1).ColumnWidth', 'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\none.csv", ws.Range("A1")): r = q.Refresh(False)', 'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\e.csv", ws.Range("A1")): q.Refresh False: r = q.ResultRange.Address',
          'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", ws.Range("C5")): q.TextFileCommaDelimiter = True: q.Refresh False: r = q.ResultRange.Address & " " & q.Destination.Address', 'ws.Range("A1").Value = "old": ws.Range("B2").Value = "keep": ' + f'{Q}{R}r = ws.Range("B2").Value', f'ws.Range("D1").Value = "next": {Q}{R}r = ws.Range("D1").Value & " " & ws.Range("E1").Value', f'ws.Range("A5").Value = "below": {Q}{R}r = ws.Range("A5").Value & " " & ws.Range("A4").Value']
for style in ('xlInsertDeleteCells', 'xlOverwriteCells', 'xlInsertEntireRows'):
    CASES += [f'ws.Range("A10").Value = "z": Set q = ws.QueryTables.Add("TEXT;" & folder & "\\long.csv", ws.Range("A1")): q.RefreshStyle = {style}: q.Refresh False: q.Connection = "TEXT;" & folder & "\\short.csv": q.Refresh False: r = q.ResultRange.Address & " " & ws.UsedRange.Address',
              f'ws.Range("A10").Value = "z": Set q = ws.QueryTables.Add("TEXT;" & folder & "\\short.csv", ws.Range("A1")): q.RefreshStyle = {style}: q.Refresh False: q.Connection = "TEXT;" & folder & "\\long.csv": q.Refresh False: r = q.ResultRange.Address & " " & ws.UsedRange.Address']
CASES += [f'{Q}q.RowNumbers = True: {R}r = q.ResultRange.Address', f'{Q}q.FieldNames = False: {R}r = q.ResultRange.Address', f'{Q}{R}ws.Range("E2").Formula = "=A2*2": q.FillAdjacentFormulas = True: q.Refresh False: r = ws.Range("E3").Formula', f'{Q}q.Refresh False: r = q.ResultRange.Address']

# The parsing settings, each alone.
for filename, settings in (('d.txt', ('q.TextFileTabDelimiter = True',)), ('s.txt', ('q.TextFileSemicolonDelimiter = True',)), ('o.txt', ('q.TextFileOtherDelimiter = "|"',)), ('sp.txt', ('q.TextFileSpaceDelimiter = True', 'q.TextFileSpaceDelimiter = True: q.TextFileConsecutiveDelimiter = True')),
                           ('q.csv', ('q.TextFileCommaDelimiter = True', 'q.TextFileCommaDelimiter = True: q.TextFileTextQualifier = xlTextQualifierSingleQuote', 'q.TextFileCommaDelimiter = True: q.TextFileTextQualifier = xlTextQualifierNone')),
                           ('d.csv', ('q.TextFileCommaDelimiter = True: q.TextFileStartRow = 2', 'q.TextFileCommaDelimiter = True: q.TextFileStartRow = 9', 'q.TextFileCommaDelimiter = True: q.TextFileColumnDataTypes = Array(2, 9, 1)', 'q.TextFileCommaDelimiter = True: q.TextFileColumnDataTypes = Array(1)', '')),
                           ('n.csv', ('q.TextFileCommaDelimiter = True', 'q.TextFileCommaDelimiter = True: q.TextFileColumnDataTypes = Array(4, 1, 2)', 'q.TextFileCommaDelimiter = True: q.TextFileTrailingMinusNumbers = False', 'q.TextFileCommaDelimiter = True: q.TextFileThousandsSeparator = ";"', 'q.TextFileCommaDelimiter = True: q.TextFileDecimalSeparator = ","')),
                           ('w.txt', ('q.TextFileParseType = xlFixedWidth: q.TextFileFixedColumnWidths = Array(2, 3)', 'q.TextFileParseType = xlFixedWidth'))):
    for setting in settings:
        CASES.append(f'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\{filename}", ws.Range("A1")): {setting}{": " if setting else ""}q.Refresh False: r = q.ResultRange.Address')

# Setting each member, then reading it and what the file makes of it.
SETS = (('AdjustColumnWidth', 'False'), ('BackgroundQuery', 'True'), ('BackgroundQuery', 'False'), ('CommandText', '"x"'), ('CommandType', 'xlCmdSql'), ('EnableEditing', 'False'), ('EnableRefresh', 'False'), ('FieldNames', 'False'), ('FillAdjacentFormulas', 'True'), ('MaintainConnection', 'False'), ('Name', '"mine"'), ('Name', '"has space"'), ('PostText', '"x"'), ('PreserveColumnInfo', 'False'), ('PreserveFormatting', 'False'),
        ('RefreshOnFileOpen', 'True'), ('RefreshPeriod', '5'), ('RefreshPeriod', '-1'), ('RefreshStyle', 'xlOverwriteCells'), ('RefreshStyle', 'xlInsertEntireRows'), ('RefreshStyle', '9'), ('RobustConnect', 'xlNever'), ('RowNumbers', 'True'), ('SaveData', 'False'), ('SavePassword', 'True'), ('SourceConnectionFile', '"x"'), ('SourceDataFile', '"x"'), ('TextFileCommaDelimiter', 'True'), ('TextFileConsecutiveDelimiter', 'True'),
        ('TextFileDecimalSeparator', '","'), ('TextFileOtherDelimiter', '"|"'), ('TextFileOtherDelimiter', '"||"'), ('TextFileParseType', 'xlFixedWidth'), ('TextFilePlatform', '1252'), ('TextFilePlatform', 'xlWindows'), ('TextFilePromptOnRefresh', 'True'), ('TextFileSemicolonDelimiter', 'True'), ('TextFileSpaceDelimiter', 'True'), ('TextFileStartRow', '3'), ('TextFileStartRow', '0'), ('TextFileTabDelimiter', 'False'),
        ('TextFileTextQualifier', 'xlTextQualifierNone'), ('TextFileThousandsSeparator', '" "'), ('TextFileTrailingMinusNumbers', 'False'), ('TextFileVisualLayout', 'xlTextVisualRTL'), ('WebFormatting', 'xlWebFormattingAll'), ('WebSelectionType', 'xlAllTables'), ('WebTables', '"1"'), ('Connection', '"TEXT;" & folder & "\\d.txt"'), ('Connection', '"URL;http://example.invalid/"'), ('Connection', '"nonsense"'))
CASES += [f'{Q}q.{name} = {value}: r = q.{name}' for name, value in SETS]
CASES += [f'{Q}q.TextFileColumnDataTypes = Array(2, 1): r = UBound(q.TextFileColumnDataTypes) & " " & q.TextFileColumnDataTypes(0)', f'{Q}q.TextFileFixedColumnWidths = Array(2, 3): r = UBound(q.TextFileFixedColumnWidths) & " " & q.TextFileFixedColumnWidths(1)', f'{Q}q.TextFileColumnDataTypes = 2: r = TypeName(q.TextFileColumnDataTypes)']

# Other methods, and what becomes of the query table.
CASES += [f'{Q}r = q.Delete & " " & ws.QueryTables.Count', f'{Q}{R}q.Delete: r = ws.QueryTables.Count & " " & ws.Range("A1").Value', f'{Q}{R}r = q.CancelRefresh', f'{Q}r = q.ResetTimer', f'{Q}r = q.SaveAsODC(folder & "\\x.odc")', f'{Q}{R}r = TypeName(ws.Range("A1").QueryTable) & " " & ws.Range("A1").QueryTable.Name', f'{Q}{R}r = TypeName(ws.Range("Z9").QueryTable)', f'{Q}r = TypeName(ws.Range("A1").QueryTable)',
          f'{Q}{R}r = ws.Parent.Names.Count & " " & ws.Names.Count & " " & ws.Names(1).Name & " " & ws.Names(1).RefersTo & " " & ws.Names(1).Visible', f'{Q}{R}r = ws.Parent.Connections.Count & " " & ws.Parent.Connections(1).Name & " " & ws.Parent.Connections(1).Type', f'{Q}r = ws.Parent.Connections.Count', f'{Q}{R}ws.Rows(1).Insert: r = q.Destination.Address & " " & q.ResultRange.Address', f'{Q}{R}ws.Columns(1).Insert: r = q.Destination.Address & " " & q.ResultRange.Address', f'{Q}{R}ws.Range("A1:C3").ClearContents: r = q.ResultRange.Address',
          f'{Q}{R}ws.Range("A2").Value = "edit": r = ws.Range("A2").Value & " " & q.ResultRange.Address', f'{Q}{R}ws.Copy: r = ActiveSheet.QueryTables.Count', f'{Q}{R}ws.Name = "Other": r = ws.Names(1).RefersTo', f'{Q}{R}ws.Parent.Worksheets.Add: ws.Delete: r = ActiveSheet.Parent.Names.Count']

# A query table is saved only once refreshed, so each setting again, refreshed, for what the file makes of it.
KILLS_EXCEL = ('PostText',)
REFUSED_SETS = (('RefreshPeriod', '-1'), ('RefreshStyle', '9'), ('TextFileStartRow', '0'), ('CommandType', 'xlCmdSql'), ('WebFormatting', 'xlWebFormattingAll'), ('WebSelectionType', 'xlAllTables'), ('WebTables', '"1"'), ('Connection', '"nonsense"'), ('Connection', '"URL;http://example.invalid/"'))
CASES += [f'{Q}q.{name} = {value}: {R}r = q.{name}' for name, value in SETS if name not in KILLS_EXCEL and (name, value) not in REFUSED_SETS and not name.startswith('TextFileParse')]
CASES += [f'{Q}q.TextFilePromptOnRefresh = False: {R}r = 1', f'{Q}q.TextFileTabDelimiter = True: {R}r = 1', f'{Q}q.TextFileTextQualifier = xlTextQualifierDoubleQuote: {R}r = 1', f'{Q}q.TextFileTrailingMinusNumbers = True: {R}r = 1', f'{Q}q.TextFilePlatform = xlMacintosh: {R}r = 1', f'{Q}q.TextFilePlatform = xlMSDOS: {R}r = 1', f'{Q}q.TextFilePlatform = 65001: {R}r = q.TextFilePlatform', f'{Q}q.TextFileVisualLayout = xlTextVisualLTR: {R}r = q.TextFileVisualLayout', f'{Q}q.RobustConnect = xlAlways: {R}r = q.RobustConnect', f'{Q}q.RobustConnect = xlAsRequired: {R}r = q.RobustConnect']

# Growing and shrinking over what is around, by each style.
GROW = 'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\short.csv", ws.Range("A1")): q.RefreshStyle = {0}: q.Refresh False: {1}q.Connection = "TEXT;" & folder & "\\long.csv": q.Refresh False: r = q.ResultRange.Address'
SHRINK = 'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\long.csv", ws.Range("A1")): q.RefreshStyle = {0}: q.Refresh False: {1}q.Connection = "TEXT;" & folder & "\\short.csv": q.Refresh False: r = q.ResultRange.Address'
for style in ('xlInsertDeleteCells', 'xlOverwriteCells', 'xlInsertEntireRows'):
    CASES += [GROW.format(style, around) for around in ('ws.Range("A4").Value = "z": ', 'ws.Range("B4").Value = "z": ', 'ws.Range("A3").Value = "z": ', 'ws.Range("B1").Value = "z": ', 'ws.Range("C9").Value = "z": ')]
    CASES += [SHRINK.format(style, around) for around in ('ws.Range("A9").Value = "z": ', 'ws.Range("B9").Value = "z": ', 'ws.Range("B2").Value = "z": ', 'ws.Range("A4").Value = "edit": ')]
CASES += [GROW.format('xlInsertDeleteCells', 'ws.Range("B2").Formula = "=A2*2": q.FillAdjacentFormulas = True: '), GROW.format('xlInsertEntireRows', 'ws.Range("B2").Formula = "=A2*2": q.FillAdjacentFormulas = True: '), GROW.format('xlOverwriteCells', 'ws.Range("B2").Formula = "=A2*2": q.FillAdjacentFormulas = True: '),
          GROW.format('xlInsertDeleteCells', 'ws.Range("A2").Interior.ColorIndex = 3: ') + ' & " " & ws.Range("A3").Interior.ColorIndex & " " & ws.Range("A7").Interior.ColorIndex', GROW.format('xlInsertDeleteCells', 'ws.Range("A2").NumberFormat = "0.00": ') + ' & " " & ws.Range("A3").NumberFormat & " " & ws.Range("A7").NumberFormat']

# Widths, what is around a first refresh, and refreshing in other states.
WIDE = 'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\wide.csv", ws.Range("A1")): '
FILES['wide.csv'] = 'a much longer heading than fits,b\r\n1,2\r\n'
CASES += [f'{WIDE}{R}r = ws.Columns(1).ColumnWidth & " " & ws.Columns(2).ColumnWidth', f'{WIDE}q.AdjustColumnWidth = False: {R}r = ws.Columns(1).ColumnWidth', f'{WIDE}ws.Columns(2).ColumnWidth = 30: {R}r = ws.Columns(2).ColumnWidth', f'{Q}ws.Columns(1).ColumnWidth = 30: {R}r = ws.Columns(1).ColumnWidth',
          f'{Q}q.PreserveColumnInfo = True: {R}ws.Columns(1).ColumnWidth = 30: q.Refresh False: r = ws.Columns(1).ColumnWidth', f'{Q}{R}ws.Columns(1).ColumnWidth = 30: q.Refresh False: r = ws.Columns(1).ColumnWidth',
          f'ws.Range("A1:C1").Value = "x": {Q}{R}r = ws.Range("D1").Value & " " & ws.Range("A1").Value', f'ws.Range("C3").Value = "x": {Q}{R}r = ws.Range("D3").Value', f'ws.Range("A1").Font.Bold = True: {Q}{R}r = ws.Range("A1").Font.Bold & " " & ws.Range("D1").Font.Bold', f'{Q}q.EnableRefresh = False: {R}r = 1', f'{Q}{R}q.EnableRefresh = False: q.Refresh False: r = 2', f'{Q}{R}ws.Protect: q.Refresh False: r = 3', f'{Q}ws.Protect: {R}r = 4',
          f'{Q}{R}q.Name = "mine": r = q.Name & " " & ws.Names(1).Name', f'{Q}{R}q.Name = "has space": r = q.Name & " " & ws.Names(1).Name', f'{Q}q.Name = "mine": {R}r = q.Name & " " & ws.Names(1).Name', f'{Q}{R}ws.Names(1).Delete: r = q.ResultRange.Address', f'{Q}{R}ws.Names(1).RefersTo = "=Sheet1!$F$1:$G$2": r = q.ResultRange.Address', f'{Q}{R}ws.Rows(2).Delete: r = q.ResultRange.Address', f'{Q}{R}ws.Rows("1:3").Delete: r = q.Destination.Address', f'{Q}{R}ws.Range("A1").Value = "h": q.Refresh False: r = ws.Range("A1").Value',
          f'{Q}{R}q.Delete: r = ws.Names.Count & " " & ws.Parent.Connections.Count', f'{Q}{R}q.Delete: Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", ws.Range("F1")): r = q.Name', f'{Q}{R}Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", ws.Range("F1")): {R}r = q.Name & " " & ws.Parent.Connections.Count & " " & ws.Parent.Connections(2).Name', f'{Q}{R}Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.txt", ws.Range("F1")): q.Refresh False: r = q.WorkbookConnection.Name']


def _probe() -> str:
    lines: list[str] = []
    for first in range(0, len(CASES), 80):
        lines += [f'Private Sub Act{first // 80}(ByVal which As Long, ByVal folder As String, ws As Worksheet, r As Variant)', '    Dim q As Object, o As Object, n As Long, t As String', '    Select Case which']
        for index, statement in enumerate(CASES[first:first + 80], first):
            lines += [f'    Case {index}', f'        {statement}']
        lines += ['    End Select', 'End Sub', '']
    lines += ['Private Function Dumped(ws As Worksheet) As String', '    Dim c As Range, s As String', '    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Function', '    For Each c In ws.UsedRange.Cells',
              '        If Not IsEmpty(c.Value) Then s = s & c.Address(False, False) & "=" & TypeName(c.Value) & ":" & c.Text & ";"', '    Next', '    Dumped = s', 'End Function', '',
              'Private Function Tried(ByVal which As Long, ByVal folder As String) As String', '    Dim wb As Workbook, ws As Worksheet, s As String, v As String, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed', '    Select Case which \\ 80']
    for block in range((len(CASES) + 79) // 80):
        lines += [f'    Case {block}: Act{block} which, folder, ws, r']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v', '    Err.Clear', '    v = "": v = ws.QueryTables.Count', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v', '    Err.Clear', '    v = "": v = Dumped(ws)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v',
              '    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String', '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def parts(path: Path, folder: str) -> dict[str, str]:
    """What a query table touches in the file, with the folder spelled <folder>."""
    found: dict[str, str] = {}
    if not path.exists():
        return found
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith('xl/queryTables/') or name in ('xl/connections.xml', 'xl/worksheets/_rels/sheet1.xml.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', '[Content_Types].xml'):
                text = package.read(name).decode('utf-8')
                found[name] = re.sub(re.escape(folder), '<folder>', text, flags=re.IGNORECASE)
    return found


def _batch(source: str, folder: str, first: int, last: int) -> list[str]:
    """The answers of these cases, in a session of their own. A batch that fails runs again a case at a time, so that the
    case that took Excel down, or stopped at a dialog, is recorded as that."""
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0)
        if result.ok and not result.dialogs:
            return str(result.value).split('@@@')
    if first == last:
        return [f'STOPPED {result.outcome}']
    return [answer for index in range(first, last + 1) for answer in _batch(source, folder, index, index)]


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder:
        for name, text in FILES.items():
            (Path(folder) / name).write_bytes(text.encode('ascii'))
        for first in range(0, len(CASES), 25):
            last = min(first + 24, len(CASES) - 1)
            answers += _batch(source, folder, first, last)
        saved = [parts(Path(folder) / f'c{index}.xlsx', folder) for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'files': FILES, 'source': source,
              'cases': [[statement, re.sub(re.escape(folder), '<folder>', answer, flags=re.IGNORECASE), part] for statement, answer, part in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statement, answer, _ in record['cases']:
        print(statement, '=>', answer)


if __name__ == '__main__':
    main()
