"""XML imported with no map, where Excel infers a schema: XmlImportXml and XmlImport with an empty map argument, the schema text
it writes, the map, the table or single cells it fills at Destination, the types and number formats the cells get, and the
parts a save holds.

Each case makes a workbook with one sheet, runs its statements, and records r, a digest of every map with its schema's text,
every table and mapped cell, the cells' values with their number formats, and the parts a save holds. Files are only ever read
from and written to a folder of TEMP.

What the sizing scout showed, and the rivals for what it did not, written before measuring:
- A leaf's type is the narrowest of integer, double, date, boolean and string its values all take (L1); "007" is an integer.
  Rivals: 1 and 0 are booleans (B1) or integers (B2), predict B2; a dateTime value infers dateTime (D1) or string (D2).
- A repeating element becomes a table at Destination with its leaves as headers, attributes first; a root that repeats
  nothing maps single cells across one row, no headers (scout). A leaf beside a repeating element goes into the table as
  a column repeated down every row (F1) or a single cell beside it (F2).
- No Destination fails with -2147024809 (scout); a second import of the same root makes root_Map1 (scout).

    python scripts/measure_xml_infer.py

writes tests/fixtures/xml_infer.json, which tests/test_excel_xml_infer.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from pyvbaharness.results import HarnessError

from fixture_workbook import copy_saved

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/xml_infer.json'
BOOKS = Path(__file__).resolve().parents[1] / 'tests/fixtures/xml_infer'
DATA = {
    'd.xml': '<list><item><n>1</n><s>a</s></item><item><n>2</n><s>b</s></item></list>',
}


def _vba(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def imported(data: str, destination: str = 'ws.Range("B2")', overwrite: str = 'True') -> str:
    """Import ``data`` with no map at ``destination``; the map made is m."""
    return f'Set m = Nothing: r = wb.XmlImportXml({_vba(data)}, m, {overwrite}, {destination}): '


def rows(*values: str) -> str:
    return '<list>' + ''.join(f'<item><v>{one}</v></item>' for one in values) + '</list>'


def typed(*values: str) -> str:
    return imported(rows(*values)) + 'r = r & " " & TypeName(m)'


CASES: list[str] = [
    # Types inferred from a leaf's values.
    typed('1', '2'),
    typed('007', '8'),
    typed('1.5', '2'),
    typed('1', 'x'),
    typed('true', 'false'),
    typed('1', '0'),
    typed('true', '1'),
    typed('2024-01-02', '2024-03-04'),
    typed('2024-01-02T10:00:00', '2024-01-02T11:00:00'),
    typed('10:00:00', '11:00:00'),
    typed('-3', '+4'),
    typed('1e3', '2'),
    typed('1.0', '2.0'),
    typed('', '2'),
    typed('12345678901234567890', '1'),
    typed('2024-01-02', 'x'),
    typed('TRUE', 'FALSE'),
    typed(' 1', '2'),
    typed('1,000', '2'),
    typed('abc', 'def'),
    typed('.5', '2'),
    typed('2024-13-01', '2024-01-02'),
    typed('NaN', '1'),
    typed('INF', '1'),
    # Shapes.
    imported('<list><title>T</title><item><n>1</n></item><item><n>2</n></item></list>'),
    imported('<list><a><n>1</n></a><a><n>2</n></a><b><m>3</m></b><b><m>4</m></b></list>'),
    imported('<list kind="k"><item><n>1</n></item><item><n>2</n></item></list>'),
    imported('<r><g><item><n>1</n></item><item><n>2</n></item></g></r>'),
    imported('<list><item><n>1</n><s>a</s></item><item><n>2</n></item></list>'),
    imported('<list><item><n>1</n><s>a</s></item><item><s>b</s><n>2</n></item></list>'),
    imported('<list><item><n>1</n><s>a</s></item></list>'),
    imported('<v>5</v>'),
    imported('<list/>'),
    imported('<list><item>text<n>1</n></item><item><n>2</n></item></list>'),
    imported('<list><item><n>1</n><tag>x</tag><tag>y</tag></item></list>'),
    imported('<list xmlns="urn:d"><item><n>1</n></item><item><n>2</n></item></list>'),
    imported('<p:list xmlns:p="urn:x"><p:item><p:n>1</p:n></p:item><p:item><p:n>2</p:n></p:item></p:list>'),
    imported('<doc><title>Hi</title><count>3</count></doc>'),
    imported('<list><item id="7"><n>1</n></item><item id="8"><n>2</n></item></list>'),
    # Destinations and arguments.
    imported(DATA['d.xml'], 'ws.Range("A1")'),
    imported(DATA['d.xml'], 'ws.Range("D5")'),
    'Set m = Nothing: r = wb.XmlImportXml(' + _vba(DATA['d.xml']) + ', m, True)',
    imported(DATA['d.xml'], overwrite='False'),
    imported(DATA['d.xml']) + 'Set m = Nothing: r = r & " " & wb.XmlImportXml(' + _vba(DATA['d.xml']) + ', m, True, ws.Range("H2")) & " " & m.Name',
    'ws.Range("B2").Value = "x": ' + imported(DATA['d.xml']),
    imported(DATA['d.xml'], 'ws.Range("B2:C3")'),
    'Set m = Nothing: r = wb.XmlImport(f & "d.xml", m, True, ws.Range("B2")) & " " & m.Name',
    imported(DATA['d.xml']) + 'r = r & " " & m.ExportXml(s) & " " & s',
    imported(DATA['d.xml']) + 'r = r & " " & m.ImportXml("<list><item><n>5</n><s>e</s></item></list>")',
    imported(DATA['d.xml']) + 'r = r & " " & m.Schemas(1).Name & " " & TypeName(m.Schemas(1).Namespace) & " " & m.Schemas(1).Namespace.URI',
    imported('<doc><title>Hi</title><count>3</count></doc>') + 'r = r & " " & m.ExportXml(s) & " " & s',
    'Set m = Nothing: r = wb.XmlImportXml(' + _vba(DATA['d.xml']) + ', m, , ws.Range("B2"))',
    'Set m = Nothing: r = wb.XmlImportXml("not xml", m, True, ws.Range("B2"))',
    # Held-out round A, predicted before it ran.
    typed('1', '1.5', 'x'),
    typed('2024-01-02', '2024-01-02T10:00:00'),
    typed('true', 'false', '0'),
    typed('-INF', '2'),
    typed('1.', '2'),
    imported('<a><b><c>1</c></b><d>x</d></a>'),
    imported('<r><item id="1"><y>2</y><x>1</x></item><item id="2"><x>3</x><y>4</y></item></r>'),
    imported('<doc><title>Hi</title><count>3</count></doc>', 'ws.Range("D5")'),
    imported('<list><item><n>1</n><s></s></item><item><n>2</n><s>b</s></item></list>') + 'r = r & " " & m.ExportXml(s) & " " & s',
    imported(DATA['d.xml']) + 'r = r & " " & m.ImportXml("<list><item><n>5</n><s>e</s></item><item><n>6</n><s>f</s></item>'
                              '<item><n>7</n><s>g</s></item></list>") & " " & ws.ListObjects(1).Range.Address',
    # Held-out round B, predicted before it ran.
    typed('0', '1', '1'),
    typed('2024-02-29', '2023-02-29'),
    typed('23:59:59', '24:00:00'),
    typed(' true', 'false'),
    imported('<list><item><n>1</n></item><item/></list>'),
    imported(DATA['d.xml']) + 'Set m = Nothing: r = r & " " & wb.XmlImportXml("<doc><t>x</t></doc>", m, True, ws.Range("H2")) & " " & m.Name',
    imported('<list><item><n>1</n><s>a</s></item><item><n>2</n><s>b</s></item><item><n>3</n><s>c</s></item></list>', 'ws.Range("A1")')
    + 'r = r & " " & m.ImportXml("<list><item><n>9</n><s>z</s></item></list>") & " " & ws.ListObjects(1).Range.Address',
    imported('<v>2024-01-02</v>') + 'r = r & " " & TypeName(ws.Range("B2").Value) & " " & ws.Range("B2").NumberFormat',
    # Held-out round C, predicted before it ran. Single cells mapped before a table: their ids go on from the table's (R1) or
    # run in the order they were made (R2).
    imported('<doc><t>x</t></doc>') + 'Set m = Nothing: r = r & " " & wb.XmlImportXml(' + _vba(DATA['d.xml']) + ', m, True, ws.Range("B5"))',
    imported('<list><item><n>1</n></item><item></item><item><n>3</n></item></list>'),
    imported(DATA['d.xml']) + 'ws.Range("B4").ClearContents: r = r & " " & m.ExportXml(s) & " " & s',
    imported('<doc><title>Hi</title><count>3</count></doc>') + 'ws.Range("B2").ClearContents: r = r & " " & m.ExportXml(s) & " " & s',
    typed('5', '-0'),
    typed('2024-01-02', ''),
    imported('<list><item><a>1</a></item><item><b>2</b></item></list>'),
    # Held-out round D, predicted before it ran.
    'ws.ListObjects.Add xlSrcRange, ws.Range("H1:H2"), , xlYes: ' + imported('<doc><t>x</t></doc>'),
    typed('3.0e1', '1'),
    imported('<list><item><n>1</n><s>a</s></item><item><n>2</n><s></s></item><item><n>3</n><s>c</s></item></list>')
    + 'r = r & " " & m.ExportXml(s) & " " & s',
    imported('<list><item><n>1</n></item><item> </item></list>'),
    imported(DATA['d.xml'], 'ws.Range("A1")') + 'r = r & " " & m.ImportXml("<list/>") & " " & ws.ListObjects(1).Range.Address',
    typed('-1.5', '2'),
]
#: The cases whose saved workbooks are kept, for the model to open as Excel saved them.
KEPT = (39, 37, 38, 26)
READS = '''Private Function Digest(wb As Workbook) As String
    Dim s As String, m As Object, lo As Object, c As Object, cell As Range
    On Error Resume Next
    s = "maps " & wb.XmlMaps.Count
    For Each m In wb.XmlMaps
        s = s & " [" & m.Name & " " & m.RootElementName & " " & m.RootElementNamespace.Prefix & " exportable " & m.IsExportable
        s = s & " " & m.ShowImportExportValidationErrors & m.SaveDataSourceDefinition & m.AdjustColumnWidth & m.PreserveColumnFilter
        s = s & m.PreserveNumberFormatting & m.AppendOnImport & " " & m.Schemas(1).Name & " " & m.Schemas(1).XML & "]"
        Err.Clear
    Next
    Err.Clear
    For Each lo In wb.Worksheets(1).ListObjects
        s = s & " {" & lo.Name & " " & lo.Range.Address & " " & lo.TableStyle & " " & lo.ShowAutoFilter & " " & TypeName(lo.XmlMap)
        For Each c In lo.ListColumns
            s = s & " " & c.Name & "=" & c.XPath.Value
        Next
        s = s & "}"
    Next
    Err.Clear
    For Each cell In wb.Worksheets(1).Range("A1:J12")
        If cell.XPath.Value <> "" Then s = s & " " & cell.Address(False, False) & ":" & cell.XPath.Value
        If Not IsEmpty(cell.Value) Then s = s & " " & cell.Address(False, False) & "=" & TypeName(cell.Value) & ":" & cell.Value & ":" & cell.NumberFormat
        Err.Clear
    Next
    Digest = s
End Function
'''


def probe() -> str:
    lines = [READS, 'Private Sub Act(ByVal which As Long, wb As Workbook, ws As Worksheet, ByVal f As String, r As Variant)',
             '    Dim m As XmlMap, s As String', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '',
              'Public Function Probe(ByVal folder As String, ByVal which As Long) As String',
              '    Dim wb As Workbook, ws As Worksheet, r As Variant, s As String, f As String',
              '    Application.DisplayAlerts = False',
              '    f = folder & "\\"',
              '    Set wb = Workbooks.Add(xlWBATWorksheet)',
              '    Set ws = wb.Worksheets(1)',
              '    On Error GoTo Failed',
              '    Act which, wb, ws, f, r',
              '    s = "ok"',
              'Answered:',
              '    On Error Resume Next',
              '    s = s & "|" & CStr(r) & "|" & Digest(wb)',
              '    Err.Clear',
              '    wb.SaveAs f & "c" & which & ".xlsx", xlOpenXMLWorkbook',
              '    If Err.Number Then s = s & "|saveE" & Err.Number',
              '    wb.Close False',
              '    Application.DisplayAlerts = True',
              '    Probe = s',
              '    Exit Function',
              'Failed:',
              '    s = "Err " & Err.Number',
              '    Resume Answered',
              'End Function']
    return '\n'.join(lines) + '\n'


def parts(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    keep = re.compile(r'xl/xmlMaps\.xml|xl/tables/.*\.xml|xl/worksheets/_rels/sheet1\.xml\.rels|\[Content_Types\]\.xml|xl/_rels/workbook\.xml\.rels|xl/connections\.xml')
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name).decode('utf-8') for name in package.namelist() if keep.fullmatch(name)}


def main() -> None:
    source = probe()
    answers: list[str] = []
    saved: list[dict[str, str]] = []
    with tempfile.TemporaryDirectory() as folder:
        for name, text in DATA.items():
            (Path(folder) / name).write_text(text, encoding='utf-8')
        for index in range(len(CASES)):
            try:
                with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                    excel.new_document()
                    done = excel.run_vba(source, 'Probe', (folder, index), timeout=120.0, line_numbers=False)
                    answers.append(str(done.value) if done.ok and not done.dialogs else f'STOPPED {done.outcome} {done.dialogs}')
            except HarnessError as failure:
                answers.append(f'STOPPED {failure}')
            # The folder's name changes every run.
            answers[-1] = answers[-1].replace(folder, '<folder>')
            saved.append({name: text.replace(folder, '<folder>') for name, text in parts(Path(folder) / f'c{index}.xlsx').items()})
            if index in KEPT and (Path(folder) / f'c{index}.xlsx').exists():
                BOOKS.mkdir(exist_ok=True)
                copy_saved(Path(folder) / f'c{index}.xlsx', BOOKS / f'c{index}.xlsx')
            print(index, CASES[index][:100], '=>', answers[-1][:300])
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'data': DATA,
              'cases': [[statement, answer, part] for statement, answer, part in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
