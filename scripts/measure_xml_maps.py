"""XML maps from a schema: XmlMaps.Add, mapping cells and table columns with XPath.SetValue, ImportXml, Import, ExportXml,
Export, SaveAsXMLData and XmlImportXml with a map, and the parts a save writes for them.

Each case makes a workbook with one sheet, writes the schemas below under a folder of TEMP, runs its statements, and records
r, a digest of every map, schema, table and mapped cell, the cells' values, and the parts a save holds. Files are only ever
read from and written to that folder.

The rivals, written before measuring:
- Add names a map after its root element with _Map, and a second of the same root with a number after it (N1), or refuses it
  (N2); it keeps the schema's text as given (K1) or rewrites it (K2).
- SetValue on more than one cell of a repeating element makes a table of them (T1) or maps them one by one (T2); on one
  cell it is a single mapping (S1) whatever repeats.
- ImportXml fills a mapped table row by row and grows it (G1), its values typed by the schema (V1) or as text reads them (V2);
  with Overwrite False or AppendOnImport it adds rows after (A1).
- ExportXml writes tab-indented XML with standalone="yes" and the xsi namespace (E1), as the first look had it.

    python scripts/measure_xml_maps.py [case ...]

measures every case, or only the cases named with the rest kept as they were measured, and writes
tests/fixtures/xml_maps.json, which tests/test_excel_xml_maps.py replays.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from pyvbaharness.results import HarnessError

from fixture_workbook import copy_saved

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/xml_maps.json'
BOOKS = Path(__file__).resolve().parents[1] / 'tests/fixtures/xml_maps'
_LIST = ('<xs:element name="list"><xs:complexType><xs:sequence><xs:element name="item" maxOccurs="unbounded"><xs:complexType>'
         '<xs:sequence>{leaves}</xs:sequence>{attributes}</xs:complexType></xs:element></xs:sequence></xs:complexType></xs:element>')
SCHEMAS: dict[str, str] = {
    's1.xsd': '<?xml version="1.0"?><xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">' + _LIST.format(
        leaves='<xs:element name="n" type="xs:integer"/><xs:element name="s" type="xs:string"/><xs:element name="d" type="xs:date"/>'
               '<xs:element name="b" type="xs:boolean"/><xs:element name="x" type="xs:decimal"/>', attributes='') + '</xs:schema>',
    's2.xsd': '<?xml version="1.0"?><xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" targetNamespace="urn:t" xmlns="urn:t" '
              'elementFormDefault="qualified">' + _LIST.format(leaves='<xs:element name="n" type="xs:integer"/>'
                                                                      '<xs:element name="s" type="xs:string"/>', attributes='') + '</xs:schema>',
    's3.xsd': '<?xml version="1.0"?><xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">' + _LIST.format(
        leaves='<xs:element name="s" type="xs:string"/>', attributes='<xs:attribute name="id" type="xs:integer"/>') + '</xs:schema>',
    's4.xsd': '<?xml version="1.0"?><xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:element name="doc"><xs:complexType>'
              '<xs:sequence><xs:element name="title" type="xs:string"/><xs:element name="count" type="xs:integer"/></xs:sequence>'
              '</xs:complexType></xs:element></xs:schema>',
    's5.xsd': '<?xml version="1.0"?><xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:element name="list"><xs:complexType>'
              '<xs:sequence><xs:element name="item" maxOccurs="unbounded"><xs:complexType><xs:sequence>'
              '<xs:element name="tag" type="xs:string" maxOccurs="unbounded"/></xs:sequence></xs:complexType></xs:element>'
              '</xs:sequence></xs:complexType></xs:element></xs:schema>',
}
DATA = {
    'd1.xml': '<list><item><n>1</n><s>a</s><d>2024-01-02</d><b>true</b><x>1.5</x></item>'
              '<item><n>2</n><s>b</s><d>2024-03-04</d><b>false</b><x>-2</x></item></list>',
    'd4.xml': '<doc><title>Hello</title><count>7</count></doc>',
}
ADD1 = 'Set m = wb.XmlMaps.Add(f & "s1.xsd", "list"): '
#: A table of five columns, each mapped to a leaf of the repeating item. Mapping one cell to a repeating element makes a single
#: mapping of the first item (round 1, kept as single_* cases below), so the columns are a table's.
COLUMNS1 = ('Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E2"), , xlYes): lo.ListColumns(1).XPath.SetValue m, "/list/item/n": '
            'lo.ListColumns(2).XPath.SetValue m, "/list/item/s": lo.ListColumns(3).XPath.SetValue m, "/list/item/d": '
            'lo.ListColumns(4).XPath.SetValue m, "/list/item/b": lo.ListColumns(5).XPath.SetValue m, "/list/item/x": ')
SINGLES1 = ('ws.Range("A1").XPath.SetValue m, "/list/item/n": ws.Range("B1").XPath.SetValue m, "/list/item/s": '
            'ws.Range("C1").XPath.SetValue m, "/list/item/d": ws.Range("D1").XPath.SetValue m, "/list/item/b": '
            'ws.Range("E1").XPath.SetValue m, "/list/item/x": ')
IMPORT1 = 'r = m.ImportXml(ReadAll(f & "d1.xml")): '
CASES: list[str] = [
    # Add, and what a map and its schema answer.
    ADD1 + 'r = TypeName(m) & " " & m.Name & " " & m.RootElementName & " " & TypeName(m.RootElementNamespace)',
    ADD1 + 'r = m.Schemas.Count & " " & m.Schemas(1).Name & " " & (m.Schemas(1).XML = Mid(ReadAll(f & "s1.xsd"), 22))',
    ADD1 + 'r = m.ShowImportExportValidationErrors & m.SaveDataSourceDefinition & m.AdjustColumnWidth & m.PreserveColumnFilter & '
           'm.PreserveNumberFormatting & m.AppendOnImport & m.IsExportable',
    ADD1 + 'Set m = wb.XmlMaps.Add(f & "s1.xsd", "list"): r = m.Name & " " & wb.XmlMaps.Count & " " & wb.XmlMaps(1).Name',
    ADD1 + 'r = wb.XmlMaps("list_Map").Name & " " & TypeName(wb.XmlMaps.Parent) & " " & TypeName(m.Parent)',
    ADD1 + 'm.Delete: r = wb.XmlMaps.Count',
    'Set m = wb.XmlMaps.Add(f & "s1.xsd"): r = m.Name & " " & m.RootElementName',
    'Set m = wb.XmlMaps.Add(f & "s1.xsd", "item"): r = m.Name',
    'Set m = wb.XmlMaps.Add(f & "s1.xsd", "nope"): r = m.Name',
    'Set m = wb.XmlMaps.Add(f & "missing.xsd", "list"): r = m.Name',
    'Set m = wb.XmlMaps.Add(f & "s2.xsd", "list"): r = m.Name & " " & m.RootElementNamespace.URI & " " & m.RootElementNamespace.Prefix & " " & '
    'wb.XmlNamespaces.Count & " " & m.Schemas(1).Namespace.URI',
    'Set m = wb.XmlMaps.Add(f & "s4.xsd", "doc"): r = m.Name & " " & m.IsExportable',
    ADD1 + 'm.Name = "Mine": r = m.Name & " " & wb.XmlMaps(1).Name',
    ADD1 + 'm.AppendOnImport = True: m.AdjustColumnWidth = False: r = m.AppendOnImport & m.AdjustColumnWidth',
    # Mapping cells.
    ADD1 + 'ws.Range("A1").XPath.SetValue m, "/list/item/n": r = ws.ListObjects.Count & " " & ws.Range("A1").XPath.Value & " " & ws.Range("A1").XPath.Repeating',
    ADD1 + COLUMNS1 + 'r = ws.ListObjects.Count & " " & ws.ListObjects(1).Range.Address & " " & m.IsExportable',
    ADD1 + 'ws.Range("A1:A3").XPath.SetValue m, "/list/item/n": r = ws.ListObjects.Count',
    'Set m = wb.XmlMaps.Add(f & "s4.xsd", "doc"): ws.Range("B2").XPath.SetValue m, "/doc/title": ws.Range("B3").XPath.SetValue m, "/doc/count": '
    'r = ws.ListObjects.Count & " " & ws.Range("B2").XPath.Value & " " & ws.Range("B2").XPath.Repeating & " " & ws.Range("B2").XPath.Map.Name & " " & m.IsExportable',
    ADD1 + 'ws.Range("A1").XPath.SetValue m, "/list/item/nope": r = ws.Range("A1").XPath.Value',
    ADD1 + COLUMNS1 + 'ws.Range("A1").XPath.Clear: r = ws.ListObjects(1).ListColumns.Count & " " & ws.Range("A1").XPath.Value',
    'Set m = wb.XmlMaps.Add(f & "s3.xsd", "list"): ws.Range("A1").XPath.SetValue m, "/list/item/@id": ws.Range("B1").XPath.SetValue m, "/list/item/s": '
    'r = ws.ListObjects(1).ListColumns(1).Name & " " & ws.ListObjects(1).ListColumns(1).XPath.Value',
    'Set m = wb.XmlMaps.Add(f & "s5.xsd", "list"): ws.Range("A1").XPath.SetValue m, "/list/item/tag": r = ws.ListObjects.Count',
    # Import.
    ADD1 + COLUMNS1 + IMPORT1 + 'r = r & " " & ws.ListObjects(1).Range.Address',
    ADD1 + COLUMNS1 + IMPORT1 + IMPORT1 + 'r = r & " " & ws.ListObjects(1).Range.Address',
    ADD1 + COLUMNS1 + 'm.AppendOnImport = True: ' + IMPORT1 + IMPORT1 + 'r = r & " " & ws.ListObjects(1).Range.Address',
    ADD1 + COLUMNS1 + 'r = m.Import(f & "d1.xml") & " " & ws.ListObjects(1).Range.Address & " " & m.DataBinding.SourceUrl',
    ADD1 + COLUMNS1 + 'r = wb.XmlImportXml(ReadAll(f & "d1.xml"), m, True) & " " & ws.ListObjects(1).Range.Address',
    ADD1 + COLUMNS1 + IMPORT1 + 'r = wb.XmlImportXml(ReadAll(f & "d1.xml"), m, False) & " " & ws.ListObjects(1).Range.Address',
    ADD1 + COLUMNS1 + 'r = wb.XmlImport(f & "d1.xml", m, True) & " " & ws.ListObjects(1).Range.Address',
    ADD1 + COLUMNS1 + 'r = m.ImportXml("<list><item><n>x</n></item></list>") & " " & ws.Range("A2").Value',
    ADD1 + COLUMNS1 + 'r = m.ImportXml("<other/>")',
    ADD1 + COLUMNS1 + 'r = m.ImportXml("not xml")',
    ADD1 + 'ws.Range("A1").XPath.SetValue m, "/list/item/n": ' + IMPORT1 + 'r = r & " " & ws.Range("A1").Value & " " & ws.Range("A2").Value',
    'Set m = wb.XmlMaps.Add(f & "s4.xsd", "doc"): ws.Range("B2").XPath.SetValue m, "/doc/title": ws.Range("B3").XPath.SetValue m, "/doc/count": '
    'r = m.Import(f & "d4.xml") & " " & ws.Range("B2").Value & " " & ws.Range("B3").Value & " " & TypeName(ws.Range("B3").Value)',
    'Set m = wb.XmlMaps.Add(f & "s2.xsd", "list"): ws.Range("A1").XPath.SetValue m, "/ns1:list/ns1:item/ns1:n": r = ws.Range("A1").XPath.Value',
    # Export.
    ADD1 + COLUMNS1 + IMPORT1 + 'r = m.ExportXml(s) & " " & s',
    ADD1 + COLUMNS1 + 'r = m.ExportXml(s) & " " & s',
    ADD1 + COLUMNS1 + IMPORT1 + 'r = m.Export(f & "out.xml") & " " & ReadAll(f & "out.xml")',
    ADD1 + COLUMNS1 + IMPORT1 + 'wb.SaveAsXMLData f & "out.xml", m: r = ReadAll(f & "out.xml")',
    'Set m = wb.XmlMaps.Add(f & "s4.xsd", "doc"): ws.Range("B2").XPath.SetValue m, "/doc/title": ws.Range("B3").XPath.SetValue m, "/doc/count": '
    'ws.Range("B2").Value = "Hi": ws.Range("B3").Value = 3: r = m.ExportXml(s) & " " & s',
    ADD1 + COLUMNS1 + IMPORT1 + 'ws.Range("A2").Value = 9: r = m.ExportXml(s) & " " & s',
    # A table's rows and the map after edits.
    ADD1 + COLUMNS1 + IMPORT1 + 'ws.ListObjects(1).ListRows.Add: r = ws.ListObjects(1).Range.Address & " " & m.ExportXml(s) & " " & s',
    ADD1 + COLUMNS1 + IMPORT1 + 'r = ws.ListObjects(1).XmlMap.Name & " " & ws.ListObjects(1).ListColumns(2).XPath.Map.Name & " " & ws.Range("B3").XPath.Value',
    # The leaves of a repeating item mapped one cell each: single mappings, each of the first item.
    ADD1 + SINGLES1 + 'r = ws.ListObjects.Count & " " & ws.Range("B1").XPath.Repeating',
    ADD1 + SINGLES1 + IMPORT1 + 'r = r & " " & ws.Range("A2").Value',
    ADD1 + SINGLES1 + IMPORT1 + 'r = m.ExportXml(s) & " " & s',
    ADD1 + SINGLES1 + 'r = m.ExportXml(s) & " " & s',
    ADD1 + SINGLES1 + IMPORT1 + 'r = m.Export(f & "out.xml") & " " & ReadAll(f & "out.xml")',
    # A table's columns, before and after.
    ADD1 + COLUMNS1 + 'r = lo.ListColumns(1).Name & " " & lo.ListColumns(1).XPath.Repeating & " " & lo.Range.Address & " " & TypeName(lo.XmlMap)',
    ADD1 + 'Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes): lo.ListColumns(1).XPath.SetValue m, "/list/item/n": '
    'r = lo.ListColumns(1).Name & " " & lo.ListColumns(2).XPath.Value & " " & lo.XmlMap.Name',
    ADD1 + 'Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes): lo.ListColumns(1).XPath.SetValue m, "/list/item/n": '
    + IMPORT1 + 'r = r & " " & lo.Range.Address',
]
# Held out: the rules from the cases above, on cases written after them. Rivals for the two not seen before: an import widens
# the mapped columns to fit (W1) or leaves them (W2); an export of a map with leaves left unmapped writes the mapped ones (U1)
# or refuses (U2).
S4 = 'Set m = wb.XmlMaps.Add(f & "s4.xsd", "doc"): '
CASES += [
    ADD1 + COLUMNS1 + IMPORT1 + 'r = ws.Columns(1).ColumnWidth & " " & ws.Columns(2).ColumnWidth & " " & ws.Columns(3).ColumnWidth',
    'Set m = wb.XmlMaps.Add(f & "s4.xsd"): r = m.Name & " " & m.RootElementName',
    ADD1 + S4 + 'r = wb.XmlMaps(2).Name & " " & m.Schemas(1).Name & " " & wb.XmlMaps(1).Schemas(1).Name',
    ADD1 + COLUMNS1 + 'r = m.ImportXml("<list></list>") & " " & lo.Range.Address',
    ADD1 + COLUMNS1 + 'r = m.ImportXml("<list><item><s>only</s></item></list>") & " " & IsEmpty(ws.Range("A2").Value) & " " & ws.Range("B2").Value',
    ADD1 + COLUMNS1 + 'r = m.ImportXml(ReadAll(f & "d1.xml"), False) & " " & lo.Range.Address',
    S4 + 'ws.Range("B2").XPath.SetValue m, "/doc/title": ws.Range("B3").XPath.SetValue m, "/doc/count": '
    'r = m.ImportXml(ReadAll(f & "d4.xml")) & m.ImportXml("<doc><title>Two</title><count>8</count></doc>") & " " & ws.Range("B2").Value & ws.Range("B3").Value',
    ADD1 + 'ws.Range("C3").XPath.SetValue m, "/list/item/s": r = ws.Range("C3").XPath.Repeating & " " & ws.Range("C3").XPath.Map.Name',
    ADD1 + 'ws.Range("A1:A2").XPath.SetValue m, "/list/item/s": ' + IMPORT1 + 'r = r & " " & ws.ListObjects(1).Range.Address & " " & ws.Range("A3").Value',
    ADD1 + COLUMNS1 + IMPORT1 + 'r = m.Export(f & "fresh.xml", True) & " " & Len(ReadAll(f & "fresh.xml"))',
    ADD1 + COLUMNS1 + 'ws.Range("A2").Value = 4: ws.Range("B2").Value = "q": ws.Range("C2").Value = DateSerial(2023, 5, 6): '
    'ws.Range("D2").Value = False: ws.Range("E2").Value = 0.25: r = m.ExportXml(s) & " " & s',
    S4 + 'ws.Range("B2").XPath.SetValue m, "/doc/title": r = m.ImportXml(ReadAll(f & "d4.xml")) & " " & ws.Range("B2").Value & " " & m.ExportXml(s) & " " & s',
    ADD1 + COLUMNS1 + IMPORT1 + 'r = TypeName(ws.Range("A3").Value) & " " & ws.Range("C3").Value & " " & ws.Range("E3").Value',
]
# Held out a second time, after the first round's fixes: an import missing what the schema requires answers 2 and an export
# without a required leaf 1; maps are listed by name; AdjustColumnWidth fits the columns. And how ExportXml's text ends its
# lines, which the harness hands back with CR LF as LF: (L1) LF, (L2) CR LF.
S3 = 'Set m = wb.XmlMaps.Add(f & "s3.xsd", "list"): '
CASES += [
    ADD1 + COLUMNS1 + IMPORT1 + 'r = m.ExportXml(s) & " " & Len(s) & " " & InStr(s, vbCr)',
    ADD1 + COLUMNS1 + 'r = m.ImportXml("<list><item><n>1</n><s>a</s></item></list>") & " " & lo.Range.Address',
    S3 + 'Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes): lo.ListColumns(1).XPath.SetValue m, "/list/item/@id": '
    'lo.ListColumns(2).XPath.SetValue m, "/list/item/s": r = m.ImportXml("<list><item><s>z</s></item></list>") & " " & ws.Range("B2").Value',
    S4 + 'Set m = wb.XmlMaps.Add(f & "s1.xsd", "list"): Set m = wb.XmlMaps.Add(f & "s3.xsd", "list"): '
    'r = wb.XmlMaps(1).Name & " " & wb.XmlMaps(2).Name & " " & wb.XmlMaps(3).Name & " " & wb.XmlMaps(3).Schemas(1).Name',
    S4 + 'ws.Range("B2").XPath.SetValue m, "/doc/title": ws.Range("B3").XPath.SetValue m, "/doc/count": '
    'r = m.ImportXml("<doc><title>T</title></doc>") & " " & IsEmpty(ws.Range("B3").Value)',
    ADD1 + COLUMNS1 + 'm.AdjustColumnWidth = False: ' + IMPORT1 + 'r = ws.Columns(2).ColumnWidth',
]
# The second held-out round was stopped before it finished (the stand-down of 2026-11-08), so the cases above are spent.
# Held-out round A, written and predicted with round B before either ran, on the same code.
NS_DATA = '<list xmlns="urn:t"><item><n>3</n><s>x</s></item></list>'
CASES += [
    ADD1 + COLUMNS1 + 'm.AppendOnImport = True: ' + IMPORT1 + 'r = m.ImportXml(ReadAll(f & "d1.xml")) & " " & lo.Range.Address & " " & ws.Range("A6").Value',
    ADD1 + COLUMNS1 + IMPORT1 + 'r = m.ExportXml(s) & " " & Len(s) & " " & InStr(s, vbCr) & " " & InStr(s, vbLf)',
    S4 + 'ws.Range("B2").XPath.SetValue m, "/doc/title": ws.Range("B3").XPath.SetValue m, "/doc/count": '
    'r = m.ImportXml("<doc><count>4</count></doc>") & " " & IsEmpty(ws.Range("B2").Value) & " " & ws.Range("B3").Value',
    ADD1 + 'r = wb.XmlMaps("LIST_MAP").Name & " " & wb.XmlMaps("list_map").RootElementName',
    ADD1 + 'ws.Range("A1:A3").XPath.SetValue m, "/list/item/n": ' + IMPORT1 + 'r = r & " " & ws.ListObjects(1).Range.Address & " " & ws.Range("A3").Value',
    ADD1 + COLUMNS1 + 'ws.Range("A2").Value = 7: ws.Range("B2").Value = "t": ws.Range("C2").Value = DateSerial(2022, 2, 3): ws.Range("D2").Value = True: '
    'ws.Range("E2").Value = 1: r = m.ExportXml(s) & " " & s',
    'Set m = wb.XmlMaps.Add(f & "s2.xsd", "list"): ws.Range("A1").XPath.SetValue m, "/ns1:list/ns1:item/ns1:n": '
    f'r = m.ImportXml("{NS_DATA.replace(chr(34), chr(34) * 2)}") & " " & ws.Range("A1").Value & " " & TypeName(ws.Range("A1").Value)',
]
# Held-out round B.
CASES += [
    ADD1 + COLUMNS1 + 'r = m.ImportXml("<list><item><n>+7</n><s>p</s><d>2024-02-29</d><b>1</b><x>2.50</x></item></list>") & " " & '
    'ws.Range("A2").Value & " " & ws.Range("C2").Value & " " & ws.Range("D2").Value & " " & ws.Range("E2").Value',
    ADD1 + COLUMNS1 + 'r = m.ImportXml("<list><item><n>1</n><s>p</s><d>2024-13-01</d><b>0</b><x>1</x></item></list>") & " " & '
    'TypeName(ws.Range("C2").Value) & " " & ws.Range("C2").Value & " " & ws.Range("D2").Value',
    ADD1 + COLUMNS1 + 'r = m.ImportXml("<list>" & String(3, "x") & "</list>")',
    ADD1 + COLUMNS1 + 'r = m.ImportXml(Replace(ReadAll(f & "d1.xml"), "</list>", "<item><n>3</n><s>c</s><d>2024-05-06</d><b>true</b><x>0</x></item></list>")) & '
    '" " & lo.Range.Address & " " & ws.Range("B4").Value',
    S3 + S4 + ADD1 + 'r = wb.XmlMaps(1).Name & " " & wb.XmlMaps(2).Name & " " & wb.XmlMaps(3).Name & " " & wb.XmlMaps.Count',
    'Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes): r = TypeName(lo.XmlMap) & " " & lo.ListColumns(1).XPath.Value & "|" & ws.Range("D4").XPath.Value',
    ADD1 + COLUMNS1 + IMPORT1 + 'ws.Range("B3").Value = "": r = m.ExportXml(s) & " " & s',
    # An empty value imported into a mapped string column: predicted Empty, as an inferred map's is (xml_infer.json).
    ADD1 + COLUMNS1 + 'r = m.ImportXml("<list><item><n>1</n><s></s><d>2024-01-02</d><b>true</b><x>1</x></item></list>") & " " & '
                      'TypeName(ws.Range("B2").Value)',
]
#: The cases whose saved workbooks are kept, for the model to open as Excel saved them.
KEPT = (16, 17, 19, 22, 34, 44, 49)
READS = '''Private Function ReadAll(ByVal path As String) As String
    Dim n As Integer, t As String
    n = FreeFile
    Open path For Binary Access Read As #n
    t = Space$(LOF(n))
    Get #n, , t
    Close #n
    ReadAll = t
End Function

Private Function Digest(wb As Workbook) As String
    Dim s As String, m As Object, k As Object, lo As Object, c As Object, i As Long, cell As Range
    On Error Resume Next
    s = "maps " & wb.XmlMaps.Count
    For Each m In wb.XmlMaps
        s = s & " [" & m.Name & " " & m.RootElementName & " schemas " & m.Schemas.Count & " exportable " & m.IsExportable & "]"
    Next
    Err.Clear
    For Each lo In wb.Worksheets(1).ListObjects
        s = s & " {" & lo.Name & " " & lo.Range.Address & " " & TypeName(lo.XmlMap)
        For Each c In lo.ListColumns
            s = s & " " & c.Name & "=" & c.XPath.Value
        Next
        s = s & "}"
    Next
    Err.Clear
    For Each cell In wb.Worksheets(1).Range("A1:F6")
        If cell.XPath.Value <> "" Then s = s & " " & cell.Address(False, False) & ":" & cell.XPath.Value
        If Not IsEmpty(cell.Value) Then s = s & " " & cell.Address(False, False) & "=" & TypeName(cell.Value) & ":" & cell.Value
        Err.Clear
    Next
    Digest = s
End Function
'''


def probe() -> str:
    lines = [READS, 'Private Sub Act(ByVal which As Long, wb As Workbook, ws As Worksheet, ByVal f As String, r As Variant)',
             '    Dim m As Object, s As String, lo As Object', '    Select Case which']
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


def write_inputs(folder: Path) -> None:
    for name, text in {**SCHEMAS, **DATA}.items():
        (folder / name).write_text(text, encoding='utf-8')


def main() -> None:
    """Measure every case, or with case numbers given, only those, kept beside the rest of the fixture as it was measured."""
    source = probe()
    wanted = [int(one) for one in sys.argv[1:]] or list(range(len(CASES)))
    before = json.loads(OUT.read_text(encoding='utf-8'))['cases'] if sys.argv[1:] else []
    answers: list[str] = [case[1] for case in before] + [''] * (len(CASES) - len(before))
    saved: list[dict[str, str]] = [case[2] for case in before] + [{}] * (len(CASES) - len(before))
    with tempfile.TemporaryDirectory() as folder:
        write_inputs(Path(folder))
        for index in wanted:
            try:
                with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                    excel.new_document()
                    done = excel.run_vba(source, 'Probe', (folder, index), timeout=120.0, line_numbers=False)
                    answer = str(done.value) if done.ok and not done.dialogs else f'STOPPED {done.outcome} {done.dialogs}'
            except HarnessError as failure:
                answer = f'STOPPED {failure}'
            # The folder's name changes every run, and the data file's text is the probe's own.
            answers[index] = answer.replace(folder, '<folder>')
            saved[index] = {name: text.replace(folder, '<folder>') for name, text in parts(Path(folder) / f'c{index}.xlsx').items()}
            if index in KEPT and (Path(folder) / f'c{index}.xlsx').exists():
                BOOKS.mkdir(exist_ok=True)
                copy_saved(Path(folder) / f'c{index}.xlsx', BOOKS / f'c{index}.xlsx')
            print(index, CASES[index][:100], '=>', answers[index][:300])
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'schemas': SCHEMAS, 'data': DATA,
              'cases': [[statement, answer, part] for statement, answer, part in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
