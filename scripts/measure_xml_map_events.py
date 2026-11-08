"""The XML map events: Workbook BeforeXmlImport, AfterXmlImport, BeforeXmlExport and AfterXmlExport, and Application's
WorkbookBefore/After twins, as each import and export raises them, and what Cancel does.

A sink class holds the workbook and Application WithEvents; every handler writes one line to a log, so a case's answer is
the order the events came in with what they were given, then what the call answered. Files go under a TEMP folder only.
The rivals, written before measuring: (C1) Cancel in a Before event stops the call, which answers as if nothing was imported
or exported and raises nothing, or (C2) raises an error; the Url an import from text gives is "" (R1).

    python scripts/measure_xml_map_events.py

writes tests/fixtures/xml_map_events.json, which tests/test_excel_xml_map_events.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/xml_map_events.json'
SCHEMA = ('<?xml version="1.0"?><xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:element name="list"><xs:complexType>'
          '<xs:sequence><xs:element name="item" maxOccurs="unbounded"><xs:complexType><xs:sequence><xs:element name="n" type="xs:integer"/>'
          '</xs:sequence></xs:complexType></xs:element></xs:sequence></xs:complexType></xs:element></xs:schema>')
DATA = '<list><item><n>1</n></item><item><n>2</n></item></list>'
SINK = '''Public WithEvents Book As Workbook
Public WithEvents App As Application

Private Sub Book_BeforeXmlImport(ByVal Map As XmlMap, ByVal Url As String, ByVal IsRefresh As Boolean, Cancel As Boolean)
    Log = Log & "[BeforeImport " & Map.Name & " " & Url & " " & IsRefresh & " " & Cancel & "]"
    If WantCancel = 1 Then Cancel = True
End Sub

Private Sub Book_AfterXmlImport(ByVal Map As XmlMap, ByVal IsRefresh As Boolean, ByVal Result As XlXmlImportResult)
    Log = Log & "[AfterImport " & Map.Name & " " & IsRefresh & " " & Result & "]"
End Sub

Private Sub Book_BeforeXmlExport(ByVal Map As XmlMap, ByVal Url As String, Cancel As Boolean)
    Log = Log & "[BeforeExport " & Map.Name & " " & Url & " " & Cancel & "]"
    If WantCancel = 2 Then Cancel = True
End Sub

Private Sub Book_AfterXmlExport(ByVal Map As XmlMap, ByVal Url As String, ByVal Result As XlXmlExportResult)
    Log = Log & "[AfterExport " & Map.Name & " " & Url & " " & Result & "]"
End Sub

Private Sub App_WorkbookBeforeXmlImport(ByVal Wb As Workbook, ByVal Map As XmlMap, ByVal Url As String, ByVal IsRefresh As Boolean, Cancel As Boolean)
    Log = Log & "[AppBeforeImport " & Map.Name & " " & Url & " " & Cancel & "]"
    If WantCancel = 3 Then Cancel = True
End Sub

Private Sub App_WorkbookAfterXmlImport(ByVal Wb As Workbook, ByVal Map As XmlMap, ByVal IsRefresh As Boolean, ByVal Result As XlXmlImportResult)
    Log = Log & "[AppAfterImport " & Map.Name & " " & Result & "]"
End Sub

Private Sub App_WorkbookBeforeXmlExport(ByVal Wb As Workbook, ByVal Map As XmlMap, ByVal Url As String, Cancel As Boolean)
    Log = Log & "[AppBeforeExport " & Map.Name & " " & Url & " " & Cancel & "]"
End Sub

Private Sub App_WorkbookAfterXmlExport(ByVal Wb As Workbook, ByVal Map As XmlMap, ByVal Url As String, ByVal Result As XlXmlExportResult)
    Log = Log & "[AppAfterExport " & Map.Name & " " & Url & " " & Result & "]"
End Sub
'''
LOG = 'Public Log As String\nPublic WantCancel As Long\n'
MAKE = ('Set m = wb.XmlMaps.Add(f & "s.xsd", "list"): Set lo = wb.Worksheets(1).ListObjects.Add(xlSrcRange, wb.Worksheets(1).Range("A1:A2"), , xlYes): '
        'lo.ListColumns(1).XPath.SetValue m, "/list/item/n": ')
CASES: list[tuple[int, str]] = [
    (0, MAKE + 'r = m.ImportXml(data)'),
    (0, MAKE + 'r = m.Import(f & "d.xml")'),
    (0, MAKE + 'r = wb.XmlImportXml(data, m, True)'),
    (0, MAKE + 'r = m.ImportXml(data): r = r & " " & m.ExportXml(s)'),
    (0, MAKE + 'r = m.ImportXml(data): r = r & " " & m.Export(f & "out.xml", True)'),
    (0, MAKE + 'r = m.ImportXml(data): wb.SaveAsXMLData f & "out2.xml", m: r = r & " saved"'),
    (1, MAKE + 'r = m.ImportXml(data) & " " & lo.Range.Address'),
    (2, MAKE + 'r = m.ImportXml(data): r = r & " " & m.ExportXml(s) & " " & Len(s)'),
    (3, MAKE + 'r = m.ImportXml(data) & " " & lo.Range.Address'),
    (0, MAKE + 'Application.EnableEvents = False: r = m.ImportXml(data) & " " & m.ExportXml(s)'),
    (0, MAKE + 'r = m.ImportXml("<other/>")'),
    (2, MAKE + 'r = m.ImportXml(data): r = r & " " & m.Export(f & "out3.xml", True) & " " & (Dir(f & "out3.xml") <> "")'),
    # Held out: a Cancel fails the call with 1004 (C2), and a text's Url is "Unspecified".
    (2, MAKE + 'r = m.ImportXml(data): wb.SaveAsXMLData f & "out4.xml", m: r = r & " " & (Dir(f & "out4.xml") <> "")'),
    (1, MAKE + 'r = wb.XmlImportXml(data, m, True) & " " & lo.Range.Address'),
    (0, MAKE + 'r = m.ImportXml(data, False) & " " & m.ImportXml("<list/>")'),
]
PROBE = '''Function Probe(ByVal folder As String, ByVal which As Long) As String
    Dim sink As New ExcelSink, wb As Workbook, m As Object, lo As Object, r As Variant, s As String, f As String, data As String, failure As String
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    f = folder & "\\"
    data = "{data}"
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set sink.Book = wb
    Set sink.App = Application
    Log = ""
    WantCancel = 0
    Select Case which
{cases}
    End Select
    Probe = Log & "|" & CStr(r) & failure
    Application.EnableEvents = False
    wb.Close False
    Exit Function
Failed:
    failure = "!E" & Err.Number
    Resume Next
End Function
'''


def probe() -> str:
    lines = []
    for index, (cancel, statement) in enumerate(CASES):
        lines += [f'    Case {index}', f'        WantCancel = {cancel}', '        Application.EnableEvents = True', '        On Error GoTo Failed',
                  f'        {statement}', '        On Error GoTo 0']
    return PROBE.replace('{data}', DATA.replace('"', '""')).replace('{cases}', '\n'.join(lines))


def install() -> str:
    lines = ['Function Install() As String', 'Dim text As String, part As Object',
             'Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)', 'part.Name = "ExcelSink"', 'text = ""']
    lines += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in SINK.splitlines()]
    lines += ['part.CodeModule.AddFromString text', 'Install = "ok"', 'End Function']
    return '\n'.join(lines)


def main() -> None:
    rows: list[list[str]] = []
    source = probe()
    with tempfile.TemporaryDirectory() as folder:
        (Path(folder) / 's.xsd').write_text(SCHEMA, encoding='utf-8')
        (Path(folder) / 'd.xml').write_text(DATA, encoding='utf-8')
        with ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
            for index, (_, statement) in enumerate(CASES):
                excel.new_document()
                excel.add_module('EventLog', LOG)
                installed = excel.run_vba(install(), 'Install', timeout=60.0, line_numbers=False)
                assert installed.ok, installed
                done = excel.run_vba(source, 'Probe', (folder, index), timeout=120.0, line_numbers=False)
                answer = (str(done.value) if done.ok and not done.dialogs else f'STOPPED {done.outcome} {done.dialogs}').replace(folder, '<folder>')
                rows.append([statement, answer])
                print(index, statement[-60:], '=>', answer)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'schema': SCHEMA, 'data': DATA, 'log': LOG, 'sink': SINK,
                               'source': source, 'cancels': [cancel for cancel, _ in CASES], 'cases': rows}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
