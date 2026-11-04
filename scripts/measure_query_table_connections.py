"""The connection behind a query table over a text file: how it is named, what it reads and takes, and when it goes.

Runs through measure_query_tables, whose cases have q, a query table over d.csv, once Q has been run. Each records r, the
sheet's cells and the parts a query table touches.

    python scripts/measure_query_table_connections.py

writes tests/fixtures/query_table_connections.json, which tests/test_excel_query_table_connections.py replays.
"""
from __future__ import annotations

from pathlib import Path

import measure_query_tables as base

Q = base.Q
R = base.R
C = 'Set c = q.WorkbookConnection: '
OTHER = 'Set o = ws.QueryTables.Add("TEXT;" & folder & "\\{0}", ws.Range("{1}")): '
ALL = 'n = 0: t = "": For Each c In ws.Parent.Connections: t = t & c.Name & ",": Next: r = t'
READS = ('Name', 'Description', 'Type', 'InModel', 'RefreshWithRefreshAll', 'Ranges.Count', 'Ranges(1).Address', 'TextConnection.Connection')
TYPES = ('Parent', 'Application', 'TextConnection', 'OLEDBConnection', 'ODBCConnection', 'ModelConnection', 'WorksheetDataConnection', 'DataFeedConnection', 'ModelTables', 'Ranges')
TEXT_READS = ('Connection', 'TextFilePlatform', 'TextFileStartRow', 'RefreshOnFileOpen', 'SourceDataFile', 'BackgroundQuery')

CASES: list[str] = []
# What the connection reads, before and after a refresh.
for prepare in (Q, Q + R):
    CASES += [f'{prepare}{C}r = c.{name}' for name in READS]
    CASES += [f'{prepare}{C}r = TypeName(c.{name})' for name in TYPES]
    CASES += [f'{prepare}{C}r = c.TextConnection.{name}' for name in TEXT_READS]
    CASES += [f'{prepare}{C}r = TypeName(c)', f'{prepare}r = ws.Parent.Connections.Count & " " & ws.Parent.Connections(1).Name',
              f'{prepare}r = ws.Parent.Connections("D").Name', f'{prepare}r = ws.Parent.Connections.Item(1).Type']

# Which file name the first table gives, and how others take numbers.
CASES += ['Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.txt", ws.Range("A1")): r = q.WorkbookConnection.Name', 'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\s.txt", ws.Range("A1")): r = q.WorkbookConnection.Name',
          f'{Q}{OTHER.format("d.csv", "F1")}r = q.WorkbookConnection.Name & " " & o.WorkbookConnection.Name', f'{Q}{OTHER.format("s.txt", "F1")}{ALL}', f'{Q}{OTHER.format("a.txt", "F1")}{ALL}',
          'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\s.txt", ws.Range("A1")): ' + OTHER.format("d.csv", "F1") + ALL, f'{Q}{OTHER.format("d.csv", "F1")}{OTHER.format("d.csv", "K1")}{ALL}',
          f'{Q}{OTHER.format("d.csv", "F1")}q.Delete: r = o.WorkbookConnection.Name & " " & ws.Parent.Connections.Count', f'{Q}{OTHER.format("d.csv", "F1")}q.Delete: {OTHER.format("d.csv", "K1")}{ALL}',
          f'{Q}{OTHER.format("d.csv", "F1")}o.Delete: {OTHER.format("d.csv", "K1")}{ALL}']

# What moves the name: the table's name, its source, a save.
CASES += [f'{Q}q.Name = "renamed": r = q.WorkbookConnection.Name', f'{Q}{R}q.Name = "renamed": r = q.WorkbookConnection.Name', f'{Q}q.Connection = "TEXT;" & folder & "\\s.txt": r = q.WorkbookConnection.Name',
          f'{Q}{R}q.Connection = "TEXT;" & folder & "\\s.txt": q.Refresh False: r = q.WorkbookConnection.Name', f'{Q}{C}q.Connection = "TEXT;" & folder & "\\s.txt": r = c.TextConnection.Connection',
          f'{Q}{C}q.TextFileStartRow = 2: r = c.TextConnection.TextFileStartRow']

# Setting the connection's own members.
CASES += [f'{Q}{C}c.Name = "mine": r = c.Name & " " & q.WorkbookConnection.Name & " " & ws.Parent.Connections("mine").Name', f'{Q}{R}{C}c.Name = "mine": r = c.Name & " " & q.Name',
          f'{Q}{C}c.Name = "": r = c.Name', f'{Q}{OTHER.format("s.txt", "F1")}{C}c.Name = "s": r = c.Name', f'{Q}{OTHER.format("s.txt", "F1")}{C}c.Name = "S": r = c.Name', f'{Q}{C}c.Name = "D": r = c.Name',
          f'{Q}{C}c.Description = "said": r = c.Description', f'{Q}{R}{C}c.Description = "said": r = c.Description', f'{Q}{R}{C}c.Name = "mine": ws.Range("A1").Value = "z": q.Refresh False: r = c.Name',
          f'{Q}{C}c.RefreshWithRefreshAll = False: r = c.RefreshWithRefreshAll', f'{Q}{R}{C}c.RefreshWithRefreshAll = False: r = c.RefreshWithRefreshAll', f'{Q}{C}c.TextConnection.TextFileStartRow = 2: r = q.TextFileStartRow',
          f'{Q}{C}c.TextConnection.Connection = "TEXT;" & folder & "\\s.txt": r = q.Connection']

# Refreshing and deleting through the connection.
CASES += [f'{Q}q.TextFileCommaDelimiter = True: q.BackgroundQuery = False: {C}c.Refresh: r = ws.Range("A1").Value', f'{Q}{R}{C}ws.Range("A2").Value = "z": q.BackgroundQuery = False: c.Refresh: r = ws.Range("A2").Value',
          f'{Q}{C}c.Delete: r = ws.Parent.Connections.Count & " " & ws.QueryTables.Count', f'{Q}{R}{C}c.Delete: r = ws.Parent.Connections.Count & " " & ws.QueryTables.Count', f'{Q}{R}ws.Parent.Connections(1).Delete: r = ws.Names.Count & " " & ws.Range("A1").Value',
          f'{Q}{R}{C}c.Delete: r = q.Name', f'{Q}{R}{C}c.Delete: r = c.Name', f'{Q}{R}{C}q.Delete: r = c.Name', f'{Q}{C}q.Delete: r = c.Name', f'{Q}{C}q.Delete: r = TypeName(c)', f'{Q}{C}c.Delete: r = c.Name',
          f'{Q}{OTHER.format("d.csv", "F1")}{C}c.Delete: r = ws.QueryTables.Count & " " & ws.QueryTables(1).WorkbookConnection.Name']

# With the other ways a connection comes and goes.
CASES += [f'{Q}ws.Parent.Connections.Add2 "z", "", "OLEDB;x", "y", 2: {ALL}', f'{Q}ws.Parent.Connections.Add2 "a", "", "OLEDB;x", "y", 2: {ALL}', f'ws.Parent.Connections.Add2 "d", "", "OLEDB;x", "y", 2: {Q}{ALL}',
          f'ws.Parent.Connections.Add2 "D", "", "OLEDB;x", "y", 2: {Q}{ALL}', f'{Q}ws.Parent.Connections.Add2 "d", "", "OLEDB;x", "y", 2: {ALL}', f'{Q}{R}ws.Parent.Connections.Add2 "z", "", "OLEDB;x", "y", 2: {ALL}',
          f'Set w = ws: Set ws = ws.Parent.Worksheets.Add: {Q}{R}Application.DisplayAlerts = False: Set book = ws.Parent: ws.Delete: Set ws = w: r = book.Connections.Count',
          f'{Q}{R}ws.Copy After:=ws: {ALL}', f'{Q}ws.Copy After:=ws: {ALL}', f'{Q}{R}ws.Copy After:=ws: r = ws.Parent.Worksheets(2).QueryTables(1).WorkbookConnection.Name']

# The ones that say what is saved.
CASES += [f'{Q}r = ws.Parent.Connections.Count', f'{Q}{R}{C}c.Name = "mine"', f'{Q}{R}{C}c.Description = "said"', f'{Q}{C}c.Name = "mine": q.TextFileCommaDelimiter = True: q.Refresh False',
          f'{Q}{OTHER.format("d.csv", "F1")}q.TextFileCommaDelimiter = True: q.Refresh False', f'{Q}{OTHER.format("d.csv", "F1")}o.TextFileCommaDelimiter = True: o.Refresh False', f'{Q}{R}{OTHER.format("d.csv", "F1")}']

# Whether a name comes free again.
O2 = OTHER.format("d.csv", "F1")
O3 = OTHER.format("d.csv", "K1")
OR = 'o.TextFileCommaDelimiter = True: o.Refresh False: '
CASES += [f'{Q}{R}{O2}{OR}q.Delete: {O3}{ALL}', f'{Q}{R}{O2}{OR}o.Delete: {O3}{ALL}', f'{Q}{R}q.Delete: Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", ws.Range("F1")): {ALL}', f'{Q}{R}{C}c.Delete: {O2}{ALL}', f'{Q}{C}c.Delete: {O2}{ALL}',
          f'{Q}{C}c.Name = "mine": {O2}{ALL}', f'{Q}{R}{C}c.Delete: ws.Parent.Connections.Add2 "d", "", "OLEDB;x", "y", 2: {ALL}', f'{Q}{C}q.Delete: {O2}{ALL}']

# A table whose connection was deleted.
CASES += [f'{Q}{C}c.Delete: r = TypeName(q.WorkbookConnection)', f'{Q}{R}{C}c.Delete: r = TypeName(q.WorkbookConnection)', f'{Q}{R}{C}c.Delete: r = q.WorkbookConnection.Name', f'{Q}{R}{C}c.Delete: ws.Range("A2").Value = "z": q.Refresh False: r = ws.Range("A2").Value',
          f'{Q}{C}c.Delete: q.TextFileCommaDelimiter = True: q.Refresh False: r = ws.Range("A1").Value', f'{Q}{R}{C}c.Delete: r = q.Connection', f'{Q}{R}{C}c.Delete: q.TextFileStartRow = 2: r = q.TextFileStartRow', f'{Q}{R}{C}c.Delete: q.Delete: r = ws.QueryTables.Count & " " & ws.Parent.Connections.Count',
          f'{Q}{R}ws.Parent.Connections(1).Delete: ws.Copy After:=ws: r = ws.Parent.Connections.Count & " " & ws.Parent.Worksheets(2).QueryTables.Count']

# A connection left behind by a table never refreshed.
LEFT = f'{Q}{C}q.Delete: '
CASES += [f'{LEFT}r = c.Type & " " & ws.Parent.Connections.Count & " " & ws.Parent.Connections(1).Name', f'{LEFT}r = c.TextConnection.Connection', f'{LEFT}c.TextConnection.TextFileStartRow = 2: r = c.TextConnection.TextFileStartRow', f'{LEFT}c.Name = "mine": r = c.Name',
          f'{LEFT}c.Delete: r = ws.Parent.Connections.Count', f'{LEFT}c.Refresh: r = ws.QueryTables.Count & " " & ws.Range("A1").Value', f'{Q}q.TextFileStartRow = 2: q.TextFileCommaDelimiter = True: q.Delete', f'{Q}{C}c.Description = "said": q.Delete',
          f'Set w = ws: Set ws = ws.Parent.Worksheets.Add: {Q}Application.DisplayAlerts = False: Set book = ws.Parent: ws.Delete: Set ws = w: r = book.Connections.Count']

# Refreshing through it.
CASES += [f'{Q}q.TextFileCommaDelimiter = True: {C}c.Refresh: r = ws.Range("A1").Value & " " & q.Refreshing & " " & q.ResultRange.Address', f'{Q}{R}{C}ws.Range("A2").Value = "z": c.Refresh: r = ws.Range("A2").Value & " " & q.Refreshing', f'{Q}{R}{C}ws.Range("A2").Value = "z": q.BackgroundQuery = False: q.WorkbookConnection.Refresh: r = ws.Range("A2").Value',
          f'{Q}{R}ws.Range("A2").Value = "z": q.BackgroundQuery = False: ws.Parent.Connections(1).Refresh: r = ws.Range("A2").Value', f'{Q}q.TextFileCommaDelimiter = True: q.BackgroundQuery = False: {C}c.Refresh: c.Refresh: r = ws.Range("A1").Value & " " & q.Refreshing', f'{Q}{R}{C}c.Delete: q.BackgroundQuery = False: ws.Range("A2").Value = "z": r = ws.Range("A2").Value']

# RefreshAll and the setting that leaves a connection out of it.
CASES += [f'{Q}{R}ws.Range("A2").Value = "z": q.BackgroundQuery = False: ws.Parent.RefreshAll: r = ws.Range("A2").Value', f'{Q}{R}{C}c.RefreshWithRefreshAll = False: ws.Range("A2").Value = "z": q.BackgroundQuery = False: ws.Parent.RefreshAll: r = ws.Range("A2").Value',
          f'{Q}q.TextFileCommaDelimiter = True: q.BackgroundQuery = False: ws.Parent.RefreshAll: r = ws.Range("A1").Value', f'{Q}{R}ws.Range("A2").Value = "z": ws.Parent.RefreshAll: r = ws.Range("A2").Value', f'{Q}{R}{C}c.Delete: ws.Range("A2").Value = "z": q.BackgroundQuery = False: ws.Parent.RefreshAll: r = ws.Range("A2").Value']

# Every member the text connection might share with its table.
TEXT_MEMBERS = ('TextFileColumnDataTypes', 'TextFileCommaDelimiter', 'TextFileConsecutiveDelimiter', 'TextFileDecimalSeparator', 'TextFileFixedColumnWidths', 'TextFileHeaderRow', 'TextFileOtherDelimiter', 'TextFileParseType', 'TextFilePromptOnRefresh', 'TextFileSemicolonDelimiter',
                'TextFileSpaceDelimiter', 'TextFileTabDelimiter', 'TextFileTextQualifier', 'TextFileThousandsSeparator', 'TextFileTrailingMinusNumbers', 'TextFileVisualLayout')
CASES += [f'{Q}{C}r = TypeName(c.TextConnection.{name})' for name in TEXT_MEMBERS]
CASES += [f'{Q}{C}r = TypeName(c.TextConnection.{name})' for name in ('Parent', 'Application', 'Creator', 'Name', 'CommandText', 'Refresh')]
CASES += [f'{Q}{C}c.TextConnection.TextFileCommaDelimiter = True: r = q.TextFileCommaDelimiter', f'{Q}{C}c.TextConnection.TextFileStartRow = 0: r = q.TextFileStartRow', f'{Q}{C}c.TextConnection.Connection = "URL;http://example.invalid/": r = q.Connection', f'{Q}{C}c.TextConnection.Connection = "x": r = q.Connection',
          f'{Q}{R}{C}c.TextConnection.Connection = "TEXT;" & folder & "\\s.txt": r = q.Connection', f'{Q}{C}c.TextConnection.TextFileParseType = xlFixedWidth: r = q.TextFileParseType']

# What a file says once it is open again.
BACK = 'ws.Parent.SaveCopyAs folder & "\\x.xlsx": Set book = Workbooks.Open(folder & "\\x.xlsx"): '
SHUT = ': book.Close False: Kill folder & "\\x.xlsx"'
BOOKED = 'book.Connections.Count & " " & book.Worksheets(1).QueryTables.Count'
CASES += [f'{Q}{R}{C}c.Description = "said": {BACK}r = book.Connections(1).Description & " " & book.Connections(1).Name{SHUT}', f'{Q}{R}{C}c.Name = "mine": {BACK}r = book.Connections(1).Name & " " & book.Worksheets(1).QueryTables(1).WorkbookConnection.Name{SHUT}',
          f'{Q}{R}{C}c.Delete: {BACK}r = {BOOKED}{SHUT}', f'{Q}{R}{C}c.RefreshWithRefreshAll = False: {BACK}r = book.Connections(1).RefreshWithRefreshAll{SHUT}', f'{Q}{C}q.Delete: {BACK}r = {BOOKED} & " " & book.Connections(1).Name & " " & book.Connections(1).TextConnection.Connection{SHUT}',
          f'{Q}{R}{C}c.Description = "said": {BACK}book.Connections(1).Description = "again": r = book.Connections(1).Description{SHUT}', f'{Q}{R}{C}c.Delete: {BACK}r = TypeName(book.Worksheets(1).QueryTables(1).WorkbookConnection){SHUT}']

# The copy of a table never refreshed, and a table deleted on a copied sheet, each run three times to see that it holds.
TWIN = 'Set w = ws.Parent.Worksheets(2): '
CASES += [f'{Q}ws.Copy After:=ws: r = ws.Parent.Worksheets(2).QueryTables.Count', f'{Q}{R}ws.Copy After:=ws: {TWIN}w.QueryTables(1).Delete: r = ws.QueryTables.Count & " " & w.QueryTables.Count & " " & ws.Parent.Connections.Count & " " & ws.Parent.Connections(1).Name']
CASES += [f'{Q}{R}ws.Copy After:=ws: {TWIN}w.QueryTables(1).Delete: {BACK}r = {BOOKED} & " " & book.Worksheets(2).QueryTables.Count{SHUT}'] * 3
CASES += [f'{Q}{R}ws.Copy After:=ws: {TWIN}ws.QueryTables(1).Delete: {BACK}r = {BOOKED} & " " & book.Worksheets(2).QueryTables.Count{SHUT}'] * 3
CASES += [f'{Q}{R}ws.Copy After:=ws: {TWIN}w.QueryTables(1).Delete'] * 3
CASES += [f'{Q}{R}ws.Copy After:=ws: {TWIN}w.QueryTables(1).Name = "Other": w.QueryTables(1).Delete: {BACK}r = {BOOKED} & " " & book.Worksheets(2).QueryTables.Count{SHUT}']

# Copied, or only named the same? Tables added apart with one name, on two sheets and in two workbooks.
APART = 'Set w = ws.Parent.Worksheets.Add(After:=ws): Set o = w.QueryTables.Add("TEXT;" & folder & "\\s.txt", w.Range("A1")): o.Refresh False: '
ELSEWHERE = 'Set book = Workbooks.Add: Set w = book.Worksheets(1): Set o = w.QueryTables.Add("TEXT;" & folder & "\\s.txt", w.Range("A1")): o.Refresh False: '
CASES += [f'{Q}{R}{APART}r = q.Name & " " & o.Name', f'{Q}{R}{APART}o.Delete: r = ws.QueryTables.Count & " " & w.QueryTables.Count', f'{Q}{R}{APART}q.Delete: r = ws.QueryTables.Count & " " & w.QueryTables.Count',
          f'{Q}{R}{ELSEWHERE}r = q.Name & " " & o.Name: book.Close False', f'{Q}{R}{ELSEWHERE}o.Delete: r = ws.QueryTables.Count & " " & w.QueryTables.Count: book.Close False', f'{Q}{R}{ELSEWHERE}q.Delete: r = ws.QueryTables.Count & " " & w.QueryTables.Count: book.Close False']

# A copy of a copy, and a copy's table deleted once the original has been renamed.
THIRD = 'ws.Copy After:=ws: Set w = ws.Parent.Worksheets(2): w.Copy After:=w: Set o = ws.Parent.Worksheets(3): '
CASES += [f'{Q}{R}{THIRD}o.QueryTables(1).Delete', f'{Q}{R}{THIRD}w.QueryTables(1).Delete', f'{Q}{R}{THIRD}q.Delete', f'{Q}{R}ws.Copy After:=ws: q.Name = "Other": ws.Parent.Worksheets(2).QueryTables(1).Delete']

# Saving what the model would otherwise refuse.
CASES += [f'{Q}{R}{O3}o.Delete', f'{Q}{C}{O3}{OR}q.Delete', f'{Q}{R}{C}c.Delete: {O3}{OR}', f'{Q}{R}{O2}{OR}Set c = o.WorkbookConnection: c.Delete']

# The member a table lacks, and RefreshAll over a file that has gone.
GONE = 'FileCopy folder & "\\d.csv", folder & "\\m.csv": Set q = ws.QueryTables.Add("TEXT;" & folder & "\\m.csv", ws.Range("A1")): q.TextFileCommaDelimiter = True: q.Refresh False: Kill folder & "\\m.csv": '
CASES += [f'{Q}{C}r = c.TextConnection.TextFileHeaderRow', f'{Q}{C}c.TextConnection.TextFileHeaderRow = True: r = c.TextConnection.TextFileHeaderRow', f'{GONE}ws.Range("A2").Value = "z": ws.Parent.RefreshAll: r = ws.Range("A2").Value', f'{GONE}r = q.Refreshing: q.Refresh False: r = ws.Range("A2").Value',
          f'{Q}{R}q.EnableRefresh = False: ws.Range("A2").Value = "z": ws.Parent.RefreshAll: r = ws.Range("A2").Value']

# Declare what the cases above use.
DECLARE = '    Dim c As Object, w As Object, book As Object'


def main() -> None:
    probe = base._probe

    def declared() -> str:
        return probe().replace('    Dim q As Object, o As Object, n As Long, t As String', f'    Dim q As Object, o As Object, n As Long, t As String\n{DECLARE}')

    base.CASES[:] = CASES
    base._probe = declared
    base.OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/query_table_connections.json'
    base.main()


if __name__ == '__main__':
    main()
