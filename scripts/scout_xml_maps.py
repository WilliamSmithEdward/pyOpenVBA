"""A first look at XML maps in live Excel, to size the work: what an import infers and maps, what an export writes, and the
parts a save holds. Files go under TEMP only. Prints what it saw and writes nothing into the repository.

    python scripts/scout_xml_maps.py
"""
from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

DATA = '<root><row><a>1</a><b>x</b></row><row><a>2</a><b>y</b></row></root>'
SCHEMA = ('<?xml version="1.0"?><xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:element name="list"><xs:complexType>'
          '<xs:sequence><xs:element name="item" maxOccurs="unbounded"><xs:complexType><xs:sequence>'
          '<xs:element name="n" type="xs:integer"/><xs:element name="s" type="xs:string"/></xs:sequence></xs:complexType>'
          '</xs:element></xs:sequence></xs:complexType></xs:element></xs:schema>')
SOURCE = '''Private Function Tried(f As String) As String
    Tried = f
End Function

Public Function Scout(ByVal folder As String, ByVal data As String, ByVal schemaPath As String) As String
    Dim wb As Workbook, ws As Worksheet, m As XmlMap, r As Variant, s As String, lo As Object, i As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    On Error Resume Next
    r = wb.XmlImportXml(data, m, True, ws.Range("A1"))
    s = s & "import " & Err.Number & " " & r & vbLf: Err.Clear
    s = s & "maps " & wb.XmlMaps.Count & vbLf
    Set m = wb.XmlMaps(1)
    s = s & "name " & m.Name & " root " & m.RootElementName & " ns " & m.RootElementNamespace.Prefix & vbLf: Err.Clear
    s = s & "schemas " & m.Schemas.Count & vbLf
    s = s & "schemaxml " & m.Schemas(1).XML & vbLf: Err.Clear
    s = s & "schemaname " & m.Schemas(1).Name & " ns " & m.Schemas(1).Namespace.URI & vbLf: Err.Clear
    s = s & "flags " & m.ShowImportExportValidationErrors & m.SaveDataSourceDefinition & m.AdjustColumnWidth & m.PreserveColumnFilter & m.PreserveNumberFormatting & m.AppendOnImport & m.IsExportable & vbLf: Err.Clear
    s = s & "lists " & ws.ListObjects.Count & vbLf
    Set lo = ws.ListObjects(1)
    s = s & "list " & lo.Name & " " & lo.Range.Address & " map " & lo.XmlMap.Name & vbLf: Err.Clear
    For i = 1 To lo.ListColumns.Count
        s = s & "col " & lo.ListColumns(i).Name & " " & lo.ListColumns(i).XPath.Value & " rep " & lo.ListColumns(i).XPath.Repeating & vbLf: Err.Clear
    Next
    s = s & "cells " & ws.Range("A2").Value & " " & TypeName(ws.Range("A2").Value) & " " & ws.Range("B3").Value & vbLf
    s = s & "a2xpath " & ws.Range("A2").XPath.Value & vbLf: Err.Clear
    Dim out As String
    r = m.ExportXml(out)
    s = s & "export " & Err.Number & " " & r & " " & out & vbLf: Err.Clear
    s = s & "binding " & TypeName(m.DataBinding) & " " & m.DataBinding.SourceUrl & vbLf: Err.Clear
    s = s & "namespaces " & wb.XmlNamespaces.Count & vbLf: Err.Clear
    wb.SaveAs folder & "\\scout.xlsx", xlOpenXMLWorkbook
    s = s & "saved " & Err.Number & vbLf: Err.Clear
    wb.SaveAsXMLData folder & "\\scout_data.xml", m
    s = s & "saveasxml " & Err.Number & vbLf: Err.Clear
    Set m = wb.XmlMaps.Add(schemaPath, "list")
    s = s & "added " & Err.Number & " " & m.Name & " " & wb.XmlMaps.Count & " root " & m.RootElementName & vbLf: Err.Clear
    s = s & "addedschema " & m.Schemas(1).XML & vbLf: Err.Clear
    ws.Range("E1").XPath.SetValue m, "/list/item/n"
    s = s & "setvalue " & Err.Number & " " & ws.Range("E1").XPath.Value & vbLf: Err.Clear
    r = m.ImportXml("<list><item><n>5</n><s>q</s></item></list>")
    s = s & "importmapped " & Err.Number & " " & r & " " & ws.Range("E1").Value & " " & ws.Range("E2").Value & " lists " & ws.ListObjects.Count & vbLf: Err.Clear
    wb.Save
    s = s & "saved2 " & Err.Number & vbLf: Err.Clear
    wb.Close False
    Scout = s
End Function
'''


def main() -> None:
    with tempfile.TemporaryDirectory() as folder:
        schema = Path(folder) / 'list.xsd'
        schema.write_text(SCHEMA, encoding='utf-8')
        with ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
            excel.new_document()
            done = excel.run_vba(SOURCE, 'Scout', (folder, DATA, str(schema)), timeout=180.0, line_numbers=False)
        print('ok', done.ok, done.outcome, done.dialogs)
        print(done.value)
        for name in ('scout.xlsx',):
            path = Path(folder) / name
            if path.exists():
                with zipfile.ZipFile(path) as package:
                    for part in package.namelist():
                        text = package.read(part).decode('utf-8', 'replace')
                        if any(word in part for word in ('xmlMaps', 'tables/', 'tableSingle', 'Content_Types', 'workbook.xml.rels', 'connections')):
                            print('==', part, len(text))
                            print(text[:1500])
        data = Path(folder) / 'scout_data.xml'
        if data.exists():
            print('== scout_data.xml', json.dumps(data.read_text(encoding='utf-8', errors='replace')[:800]))


if __name__ == '__main__':
    main()
