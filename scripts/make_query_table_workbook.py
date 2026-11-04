"""Has Excel save a workbook holding query tables, for tests/test_excel_query_tables_saved.py.

The workbook has two query tables over text files on Sheet1, one with settings changed, an ordinary cell, and a second
sheet. Excel's temporary folder is spelled C:\\data in what is committed, and the save path is stripped.

    python scripts/make_query_table_workbook.py

writes tests/fixtures/query_table_saved.xlsx.
"""
from __future__ import annotations

import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/query_table_saved.xlsx'
SOURCE = '''
Public Function Make(ByVal folder As String) As String
    Dim wb As Workbook, ws As Worksheet, q As QueryTable
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A20").Value = "ordinary"
    Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", ws.Range("A1"))
    q.TextFileCommaDelimiter = True
    q.Refresh False
    Set q = ws.QueryTables.Add("TEXT;" & folder & "\\s.txt", ws.Range("F1"))
    q.Name = "second one"
    q.TextFileSemicolonDelimiter = True
    q.TextFileTabDelimiter = False
    q.TextFileStartRow = 2
    q.RefreshStyle = xlOverwriteCells
    q.AdjustColumnWidth = False
    q.BackgroundQuery = False
    q.Refresh False
    wb.Worksheets.Add After:=ws
    wb.SaveAs folder & "\\saved.xlsx", xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Make = "made"
End Function
'''


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        (Path(folder) / 'd.csv').write_bytes(b'a,b,c\r\n1,2,3\r\n4,x,6\r\n')
        (Path(folder) / 's.txt').write_bytes(b'skip\r\nh;i\r\n7;8\r\n')
        excel.new_document()
        result = excel.run_vba(SOURCE, 'Make', (folder,), timeout=120.0)
        assert result.ok and not result.dialogs, result
        with zipfile.ZipFile(Path(folder) / 'saved.xlsx') as package, zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as written:
            for name in package.namelist():
                data = package.read(name)
                if name.endswith(('.xml', '.rels')):
                    text = data.decode('utf-8')
                    text = re.sub(re.escape(folder), lambda _: 'C:\\data', text, flags=re.IGNORECASE)
                    # The folder Excel saved in, which names the machine.
                    text = re.sub(r'<mc:AlternateContent[^>]*>\s*<mc:Choice[^>]*>\s*<x15ac:absPath[^>]*/>\s*</mc:Choice>\s*</mc:AlternateContent>', '', text)
                    data = text.encode('utf-8')
                written.writestr(name, data)


if __name__ == '__main__':
    main()
