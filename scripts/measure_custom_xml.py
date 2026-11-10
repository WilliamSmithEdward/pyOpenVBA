"""CustomXMLPart and its nodes: what a macro reads from a part it added or loaded, what XPath selects, what AddNode,
LoadXML and Load do, and what Excel writes under customXml/ when the book is saved.

Each case is statements run on a new workbook (``p`` the part the first statement leaves, ``n`` a node, ``ns`` a
NamespaceManager), then expressions read one by one, each as its value, its TypeName or its error. The book is then
saved as .xlsx and the parts it holds under customXml/, the workbook's relationship to them and their content types
are gathered.

    python scripts/measure_custom_xml.py [round]
    python scripts/measure_custom_xml.py predict 1   # tests/fixtures/custom_xml1_predictions.json

writes tests/fixtures/custom_xml<round>.json, which tests/test_excel_custom_xml.py replays.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / 'tests/fixtures'

ROOT_XML = '<root xmlns="urn:a"><x k="1">one</x><x k="2">two</x><y>three</y></root>'
PLAIN_XML = '<book><title lang="en">Text</title><title lang="fr">Texte</title><n>5</n></book>'
ADD = f'Set p = ActiveWorkbook.CustomXMLParts.Add("{ROOT_XML.replace(chr(34), chr(34) * 2)}")'
ADD_PLAIN = f'Set p = ActiveWorkbook.CustomXMLParts.Add("{PLAIN_XML.replace(chr(34), chr(34) * 2)}")'
NODE = ('TypeName({0}) & "|" & {0}.BaseName & "|" & {0}.NamespaceURI & "|" & {0}.NodeType & "|" & {0}.Text & "|" & '
        '{0}.XPath')

#: Each case: a name, the statements, and the expressions read.
CASES: list[tuple[str, list[str], list[str]]] = [
    ('saved', [ADD], ['p.NamespaceURI', 'p.BuiltIn', 'ActiveWorkbook.CustomXMLParts.Count']),
    ('saved two', [ADD, ADD_PLAIN], ['ActiveWorkbook.CustomXMLParts.Count', 'p.NamespaceURI']),
    ('saved empty', ['Set p = ActiveWorkbook.CustomXMLParts.Add()'],
     ['p.XML', 'p.NamespaceURI', 'TypeName(p.DocumentElement)']),
    ('document element', [ADD], [NODE.format('p.DocumentElement'), 'p.DocumentElement.XML',
                                 'p.DocumentElement.ChildNodes.Count', 'p.DocumentElement.HasChildNodes',
                                 NODE.format('p.DocumentElement.FirstChild'),
                                 NODE.format('p.DocumentElement.LastChild'),
                                 'TypeName(p.DocumentElement.ParentNode)', 'TypeName(p.DocumentElement.OwnerPart)',
                                 'TypeName(p.DocumentElement.OwnerDocument)']),
    ('plain document', [ADD_PLAIN], [NODE.format('p.DocumentElement'),
                                     NODE.format('p.DocumentElement.FirstChild'),
                                     'p.DocumentElement.FirstChild.Attributes.Count',
                                     NODE.format('p.DocumentElement.FirstChild.Attributes(1)'),
                                     NODE.format('p.DocumentElement.FirstChild.FirstChild'),
                                     NODE.format('p.DocumentElement.FirstChild.NextSibling'),
                                     'TypeName(p.DocumentElement.FirstChild.PreviousSibling)']),
    ('select plain', [ADD_PLAIN], ['p.SelectNodes("/book/title").Count', 'p.SelectNodes("//title").Count',
                                   'p.SelectNodes("/book/*").Count', 'p.SelectNodes("//@lang").Count',
                                   'p.SelectNodes("/book/title[2]").Count',
                                   NODE.format('p.SelectSingleNode("/book/title[2]")'),
                                   NODE.format('p.SelectSingleNode("/book/title[@lang=\'fr\']")'),
                                   NODE.format('p.SelectSingleNode("/book/title/@lang")'),
                                   NODE.format('p.SelectSingleNode("/book/n/text()")'),
                                   'TypeName(p.SelectSingleNode("/book/none"))', 'p.SelectNodes("/book/none").Count',
                                   'p.SelectNodes("book/title").Count', 'p.SelectNodes("//title[last()]").Count',
                                   'p.SelectNodes("/book/title[position()>1]").Count',
                                   'p.SelectNodes("/book/n[.=5]").Count', 'p.SelectNodes("count(//title)").Count',
                                   'p.SelectNodes("/book/title[").Count', 'p.SelectNodes("").Count',
                                   'p.SelectNodes("/book/title")(1).Text', 'p.SelectNodes("/book/title")(3).Text',
                                   'p.SelectNodes("/book/title").Item(2).Text']),
    ('select namespaced', [ADD], ['p.SelectNodes("/root/x").Count', 'p.SelectNodes("//x").Count',
                                  'p.SelectNodes("/ns0:root/ns0:x").Count', 'p.NamespaceManager.Count',
                                  'p.NamespaceManager.LookupNamespace("ns0")', 'p.NamespaceManager.LookupPrefix("urn:a")',
                                  'p.SelectNodes("//*").Count', 'p.DocumentElement.XPath',
                                  'p.DocumentElement.FirstChild.XPath']),
    ('prefix', [ADD, 'p.NamespaceManager.AddNamespace "a", "urn:a"'],
     ['p.NamespaceManager.Count', 'p.NamespaceManager(1).Prefix & "=" & p.NamespaceManager(1).NamespaceURI',
      'p.SelectNodes("/a:root/a:x").Count', NODE.format('p.SelectSingleNode("/a:root/a:x[2]")'),
      'p.SelectNodes("//a:x/@k").Count', 'p.NamespaceManager.LookupPrefix("urn:a")',
      'p.DocumentElement.FirstChild.XPath', 'p.SelectNodes("/b:root").Count']),
    ('add node', [ADD_PLAIN, 'Set n = p.DocumentElement',
                  'p.AddNode n, "extra", "", , msoCustomXMLNodeElement, "added"'],
     ['p.XML', 'p.DocumentElement.ChildNodes.Count', NODE.format('p.DocumentElement.LastChild')]),
    ('add attribute', [ADD_PLAIN, 'Set n = p.DocumentElement',
                       'p.AddNode n, "id", "", , msoCustomXMLNodeAttribute, "7"'],
     ['p.XML', 'n.Attributes.Count']),
    ('add before', [ADD_PLAIN, 'Set n = p.DocumentElement',
                    'p.AddNode n, "first", "", n.FirstChild, msoCustomXMLNodeElement, ""'],
     ['p.XML', NODE.format('p.DocumentElement.FirstChild')]),
    ('add text', [ADD_PLAIN, 'Set n = p.SelectSingleNode("/book/n")',
                  'p.AddNode n, "", "", , msoCustomXMLNodeText, "6"'], ['p.XML', 'n.Text']),
    ('add comment', [ADD_PLAIN, 'Set n = p.DocumentElement',
                     'p.AddNode n, "", "", , msoCustomXMLNodeComment, "note"'], ['p.XML']),
    ('add namespaced', [ADD, 'Set n = p.DocumentElement', 'p.AddNode n, "z", "urn:a", , , "4"'],
     ['p.XML', 'p.SelectNodes("//*").Count']),
    ('add other namespace', [ADD, 'Set n = p.DocumentElement', 'p.AddNode n, "z", "urn:b", , , "4"'], ['p.XML']),
    ('add to nothing', [ADD], ['p.AddNode(Nothing, "z")']),
    ('load xml', ['Set p = ActiveWorkbook.CustomXMLParts.Add()', f'b = p.LoadXML("{ROOT_XML.replace(chr(34), chr(34) * 2)}")'],
     ['b', 'p.XML', 'p.NamespaceURI', 'p.DocumentElement.BaseName']),
    ('load xml twice', [ADD, f'b = p.LoadXML("<other/>")'], ['b', 'p.XML']),
    ('load xml bad', ['Set p = ActiveWorkbook.CustomXMLParts.Add()', 'b = p.LoadXML("<a><b></a>")'],
     ['b', 'p.XML', 'TypeName(p.DocumentElement)']),
    ('add bad', [], ['TypeName(ActiveWorkbook.CustomXMLParts.Add("<a>"))']),
    ('errors', [ADD], ['p.Errors.Count', 'TypeName(p.Errors)', 'TypeName(p.SchemaCollection)']),
    ('node edits', [ADD_PLAIN, 'Set n = p.SelectSingleNode("/book/n")', 'n.Text = "9"'],
     ['p.XML', 'n.NodeValue', 'n.FirstChild.NodeValue']),
    ('node delete', [ADD_PLAIN, 'p.SelectSingleNode("/book/title[1]").Delete'],
     ['p.XML', 'p.SelectNodes("//title").Count']),
    ('node append', [ADD_PLAIN, 'Set n = p.DocumentElement', 'n.AppendChildNode "more", "", msoCustomXMLNodeElement, "x"',
                     'n.AppendChildSubtree "<sub><a>1</a></sub>"'], ['p.XML']),
    ('node remove', [ADD_PLAIN, 'Set n = p.DocumentElement', 'n.RemoveChild n.LastChild'], ['p.XML']),
    ('node replace', [ADD_PLAIN, 'Set n = p.DocumentElement',
                      'n.ReplaceChildNode n.LastChild, "m", "", msoCustomXMLNodeElement, "8"'], ['p.XML']),
    ('select node from node', [ADD_PLAIN, 'Set n = p.DocumentElement'],
     ['n.SelectNodes("title").Count', 'n.SelectNodes("./title").Count', 'n.SelectSingleNode("n").Text',
      'n.FirstChild.SelectNodes("..").Count', 'n.FirstChild.SelectNodes("following-sibling::*").Count']),
    ('built in', [], ['ActiveWorkbook.CustomXMLParts.Count', 'ActiveWorkbook.CustomXMLParts(1).NamespaceURI',
                      'TypeName(ActiveWorkbook.CustomXMLParts(1).DocumentElement)',
                      'ActiveWorkbook.CustomXMLParts(1).SelectNodes("//*").Count']),
    ('deleted then saved', [ADD, ADD_PLAIN, 'p.Delete'], ['ActiveWorkbook.CustomXMLParts.Count']),
]


MIXED_XML = ('<?xml version="1.0" encoding="UTF-8"?><!--lead--><r a="1 &amp; 2">x &lt; y &gt; z "q"'
             '<![CDATA[c<d]]><?pi data?><e/><!--inner--></r>')
ADD_MIXED = f'Set p = ActiveWorkbook.CustomXMLParts.Add("{MIXED_XML.replace(chr(34), chr(34) * 2)}")'
TYPED = f'Set tp = ActiveWorkbook.CustomXMLParts.Add("{PLAIN_XML.replace(chr(34), chr(34) * 2)}")'
#: The second look: ids, the namespace manager, AddNode typed, node kinds and escaping, more XPath, Load.
CASES_B: list[tuple[str, list[str], list[str]]] = [
    ('ids', [ADD], ['p.ID', 'ActiveWorkbook.CustomXMLParts.SelectByID(p.ID).NamespaceURI']),
    ('manager plain', [ADD_PLAIN], ['p.NamespaceManager.Count',
                                    'p.NamespaceManager(1).Prefix & "=" & p.NamespaceManager(1).NamespaceURI',
                                    'p.NamespaceManager(2).Prefix & "=" & p.NamespaceManager(2).NamespaceURI',
                                    'p.NamespaceManager.LookupNamespace("xml")', 'TypeName(p.NamespaceManager)',
                                    'TypeName(p.NamespaceManager(1))', 'p.NamespaceManager(3).Prefix']),
    ('manager namespaced', [ADD], ['p.NamespaceManager(1).Prefix & "=" & p.NamespaceManager(1).NamespaceURI',
                                   'p.NamespaceManager(2).Prefix & "=" & p.NamespaceManager(2).NamespaceURI']),
    ('manager add twice', [ADD, 'p.NamespaceManager.AddNamespace "a", "urn:a"',
                           'p.NamespaceManager.AddNamespace "a", "urn:b"'],
     ['p.NamespaceManager.Count', 'p.NamespaceManager.LookupNamespace("a")']),
    ('typed add element', [TYPED, 'Set tn = tp.DocumentElement',
                           'tp.AddNode tn, "extra", "", , msoCustomXMLNodeElement, "added"'],
     ['tp.XML', 'tp.DocumentElement.LastChild.XPath']),
    ('typed add attribute', [TYPED, 'Set tn = tp.DocumentElement',
                             'tp.AddNode tn, "id", "", , msoCustomXMLNodeAttribute, "7"'], ['tp.XML']),
    ('typed add text', [TYPED, 'Set tn = tp.SelectSingleNode("/book/n")',
                        'tp.AddNode tn, "", "", , msoCustomXMLNodeText, "6"'], ['tp.XML', 'tn.Text']),
    ('typed add comment', [TYPED, 'Set tn = tp.DocumentElement',
                           'tp.AddNode tn, "", "", , msoCustomXMLNodeComment, "note"'], ['tp.XML']),
    ('typed add instruction', [TYPED, 'Set tn = tp.DocumentElement',
                               'tp.AddNode tn, "pi", "", , msoCustomXMLNodeProcessingInstruction, "d"'], ['tp.XML']),
    ('typed add namespaced', [TYPED, 'Set tn = tp.DocumentElement', 'tp.AddNode tn, "z", "urn:b", , , "4"'],
     ['tp.XML']),
    ('typed add defaults', [TYPED, 'Set tn = tp.DocumentElement', 'tp.AddNode tn, "z"'], ['tp.XML']),
    ('typed add no name', [TYPED, 'Set tn = tp.DocumentElement', 'tp.AddNode tn'], ['tp.XML']),
    ('late add nothing sibling', [ADD_PLAIN, 'Set n = p.DocumentElement', 'p.AddNode n, "z", "", Nothing, 1, "v"'],
     ['p.XML']),
    ('mixed', [ADD_MIXED], ['p.XML', 'p.DocumentElement.ChildNodes.Count', 'p.DocumentElement.Text',
                            'p.DocumentElement.ChildNodes(1).NodeType', 'p.DocumentElement.ChildNodes(2).NodeType',
                            'p.DocumentElement.ChildNodes(3).NodeType', 'p.DocumentElement.ChildNodes(4).NodeType',
                            'p.DocumentElement.ChildNodes(5).NodeType', 'p.DocumentElement.ChildNodes(2).NodeValue',
                            'p.DocumentElement.ChildNodes(3).BaseName & "=" & p.DocumentElement.ChildNodes(3).NodeValue',
                            'p.DocumentElement.ChildNodes(5).NodeValue', 'p.DocumentElement.Attributes(1).NodeValue',
                            'p.DocumentElement.Attributes(1).XML', 'p.DocumentElement.ChildNodes(1).XML',
                            'p.DocumentElement.ChildNodes(2).XPath', 'p.DocumentElement.ChildNodes(3).XPath',
                            'p.DocumentElement.ChildNodes(5).XPath', 'p.DocumentElement.ChildNodes(4).XPath']),
    ('spaces', ['Set p = ActiveWorkbook.CustomXMLParts.Add("<r>" & vbLf & "  <a>1</a>" & vbCrLf & "</r>")'],
     ['p.XML', 'p.DocumentElement.ChildNodes.Count', 'p.DocumentElement.Text', 'Len(p.XML)']),
    ('more paths', [ADD_PLAIN], [f'p.SelectNodes("{path}").Count' for path in (
        '//title[@lang]', '/book/title[text()=\'Text\']', '//n | //title', '/book/title[@lang=\'en\' or @lang=\'x\']',
        '/*', '/', '/book/title[contains(.,\'Tex\')]', '/book/n[number(.)>4]', 'string(//n)', '//title/..',
        '/book/title/following-sibling::n', '/book/descendant::*', '*', '/book/*[2]', '//title[1]', '(//title)[1]',
        '//text()', '//node()', '/book/title[starts-with(@lang,\'f\')]', '/book/title[not(@lang=\'en\')]',
        '//@*', '/book/title[@lang!=\'en\']', '/book/n[. > 4 and . < 6]', 'descendant::title',
        '/book/title[position()=last()]', '//title[@lang=\'en\']/following::*', '//*[local-name()=\'n\']',
        '/book/title[2]/preceding-sibling::*', 'ancestor::*', '//comment()', '/book/n[.=\'5\']')]),
    ('node members', [ADD_PLAIN, 'Set n = p.SelectSingleNode("/book/title[1]/@lang")'],
     ['n.NodeValue', 'n.Text', 'TypeName(n.ParentNode)', 'n.XML', 'n.XPath', 'TypeName(n.OwnerPart)',
      'n.HasChildNodes', 'n.ChildNodes.Count', 'TypeName(n.FirstChild)', 'p.DocumentElement.NodeValue',
      'p.SelectSingleNode("/book/title[1]").NodeValue', 'p.SelectSingleNode("/book/title[1]").XML',
      'p.SelectSingleNode("/book/title[1]/text()").NodeValue']),
    ('set values', [ADD_PLAIN, 'p.SelectSingleNode("/book/title[1]/@lang").NodeValue = "de"',
                    'p.SelectSingleNode("/book/title[2]").Text = "a<b"',
                    'p.SelectSingleNode("/book/n/text()").NodeValue = "8"'], ['p.XML']),
    ('delete root', [ADD_PLAIN, 'p.DocumentElement.Delete'], ['p.XML', 'TypeName(p.DocumentElement)']),
    ('load file', ['Open Environ("TEMP") & "\\pyovba_cx.xml" For Output As #1', 'Print #1, "<f><g>1</g></f>"',
                   'Close #1', 'Set p = ActiveWorkbook.CustomXMLParts.Add()',
                   'b = p.Load(Environ("TEMP") & "\\pyovba_cx.xml")', 'Kill Environ("TEMP") & "\\pyovba_cx.xml"'],
     ['b', 'p.XML', 'p.DocumentElement.BaseName']),
    ('load missing', ['Set p = ActiveWorkbook.CustomXMLParts.Add()',
                      'b = p.Load(Environ("TEMP") & "\\pyovba_cx_none.xml")'], ['b', 'p.XML']),
    ('load loaded', [ADD, 'b = p.Load(Environ("TEMP") & "\\pyovba_cx_none.xml")'], ['b']),
    ('errors and schemas', [ADD], ['p.Errors.Count', 'p.SchemaCollection.Count', 'TypeName(p.Errors(1))',
                                   'p.SchemaCollection.Validate', 'TypeName(p.SchemaCollection.Item(1))']),
    ('built in reads', [], ['ActiveWorkbook.CustomXMLParts(1).DocumentElement.BaseName',
                            'ActiveWorkbook.CustomXMLParts(2).NamespaceURI',
                            'ActiveWorkbook.CustomXMLParts(3).NamespaceURI',
                            'ActiveWorkbook.CustomXMLParts(1).XML', 'ActiveWorkbook.CustomXMLParts(2).XML',
                            'ActiveWorkbook.CustomXMLParts(3).XML']),
]


#: A class whose WithEvents variables log each event a case's statements raise.
SINK = '''Public WithEvents p As CustomXMLPart
Public WithEvents ps As CustomXMLParts
Public log As String

Private Function Nm(n As Object) As String
    On Error Resume Next
    Nm = "?"
    If n Is Nothing Then Nm = "-" Else Nm = n.BaseName & ":" & n.NodeType
End Function

Private Sub p_NodeAfterDelete(ByVal OldNode As CustomXMLNode, ByVal OldParentNode As CustomXMLNode, ByVal OldNextSibling As CustomXMLNode, ByVal InUndoRedo As Boolean)
    log = log & "D(" & Nm(OldNode) & "," & Nm(OldParentNode) & "," & Nm(OldNextSibling) & "," & InUndoRedo & ")"
End Sub

Private Sub p_NodeAfterInsert(ByVal NewNode As CustomXMLNode, ByVal InUndoRedo As Boolean)
    log = log & "I(" & Nm(NewNode) & "," & InUndoRedo & ")"
End Sub

Private Sub p_NodeAfterReplace(ByVal OldNode As CustomXMLNode, ByVal NewNode As CustomXMLNode, ByVal InUndoRedo As Boolean)
    log = log & "R(" & Nm(OldNode) & "," & Nm(NewNode) & "," & InUndoRedo & ")"
End Sub

Private Sub ps_PartAfterAdd(ByVal NewPart As CustomXMLPart)
    log = log & "A(" & NewPart.NamespaceURI & ")"
End Sub

Private Sub ps_PartAfterLoad(ByVal Part As CustomXMLPart)
    log = log & "L(" & Part.NamespaceURI & ")"
End Sub

Private Sub ps_PartBeforeDelete(ByVal OldPart As CustomXMLPart)
    log = log & "B(" & OldPart.NamespaceURI & ")"
End Sub
'''
LISTEN = 'Set k = New Sink: Set k.ps = ActiveWorkbook.CustomXMLParts'
HEAR = 'Set k.p = p'
#: The third look: events, carriage returns, schema collections.
CASES_C: list[tuple[str, list[str], list[str]]] = [
    ('events add', [LISTEN, ADD_PLAIN], ['k.log']),
    ('events add empty and load', [LISTEN, 'Set p = ActiveWorkbook.CustomXMLParts.Add()',
                                   'b = p.LoadXML("<q/>")'], ['k.log']),
    ('events delete part', [ADD, LISTEN, 'p.Delete'], ['k.log']),
    ('events add node', [ADD_PLAIN, LISTEN, HEAR, 'Set tp = p', 'Set tn = tp.DocumentElement',
                         'tp.AddNode tn, "z", "", , msoCustomXMLNodeElement, "v"'], ['k.log']),
    ('events add attribute', [ADD_PLAIN, LISTEN, HEAR, 'Set tp = p', 'Set tn = tp.DocumentElement',
                              'tp.AddNode tn, "id", "", , msoCustomXMLNodeAttribute, "7"'], ['k.log']),
    ('events delete node', [ADD_PLAIN, LISTEN, HEAR, 'p.SelectSingleNode("/book/title[1]").Delete'], ['k.log']),
    ('events text', [ADD_PLAIN, LISTEN, HEAR, 'p.SelectSingleNode("/book/n").Text = "9"'], ['k.log']),
    ('events node value', [ADD_PLAIN, LISTEN, HEAR, 'p.SelectSingleNode("/book/n/text()").NodeValue = "9"'],
     ['k.log']),
    ('events attribute value', [ADD_PLAIN, LISTEN, HEAR, 'p.SelectSingleNode("//@lang").NodeValue = "de"'],
     ['k.log']),
    ('events append', [ADD_PLAIN, LISTEN, HEAR, 'p.DocumentElement.AppendChildNode "m", "", 1, "x"',
                       'p.DocumentElement.AppendChildSubtree "<s><t/></s>"'], ['k.log']),
    ('events remove', [ADD_PLAIN, LISTEN, HEAR, 'p.DocumentElement.RemoveChild p.DocumentElement.LastChild'],
     ['k.log']),
    ('events replace', [ADD_PLAIN, LISTEN, HEAR,
                        'p.DocumentElement.ReplaceChildNode p.DocumentElement.LastChild, "m", "", 1, "8"'],
     ['k.log']),
    ('events load xml', [LISTEN, 'Set p = ActiveWorkbook.CustomXMLParts.Add()', HEAR, 'b = p.LoadXML("<q><r/></q>")'],
     ['k.log']),
    ('returns', ['Set p = ActiveWorkbook.CustomXMLParts.Add("<r>" & vbLf & "  <a>1</a>" & vbCrLf & "</r>")'],
     ['Replace(Replace(p.XML, vbCr, "~"), vbLf, "^")', 'Replace(Replace(p.DocumentElement.Text, vbCr, "~"), vbLf, "^")',
      'Replace(Replace(p.DocumentElement.XML, vbCr, "~"), vbLf, "^")',
      'Replace(Replace(p.DocumentElement.FirstChild.NodeValue, vbCr, "~"), vbLf, "^")']),
    ('returns in text', [ADD_PLAIN, 'p.SelectSingleNode("/book/n").Text = "a" & vbCrLf & "b" & vbLf & "c"'],
     ['Replace(Replace(p.XML, vbCr, "~"), vbLf, "^")',
      'Replace(Replace(p.SelectSingleNode("/book/n").Text, vbCr, "~"), vbLf, "^")']),
    ('schemas namespaced', [ADD], ['p.SchemaCollection.Count', 'p.SchemaCollection(1).NamespaceURI',
                                   'p.SchemaCollection(1).Location', 'p.SchemaCollection(1).XML',
                                   'TypeName(p.SchemaCollection(1).Parent)']),
    ('schemas plain', [ADD_PLAIN], ['p.SchemaCollection.Count', 'p.SchemaCollection.Validate',
                                    'TypeName(p.SchemaCollection.Item(1))']),
    ('schema set nothing', [ADD, 'Set p.SchemaCollection = Nothing'], ['p.SchemaCollection.Count']),
]


def source(cases: list[tuple[str, list[str], list[str]]]) -> str:
    """Go, which runs case ``which`` on a new workbook and saves it to ``path``; each case is a function of its own,
    as one procedure holding them all is too large for VBA."""
    lines = ['Public Function Go(ByVal which As Long, ByVal path As String) As String',
             '    Dim wb As Workbook, s As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    On Error Resume Next', '    Select Case which']
    lines += [f'    Case {number}: s = Case{number}(wb)' for number in range(len(cases))]
    lines += ['    End Select', '    Err.Clear', '    Application.DisplayAlerts = False', '    wb.SaveAs path, 51',
              '    s = s & " save=" & Err.Number', '    Application.DisplayAlerts = True', '    wb.Close False',
              '    Go = s', 'End Function']
    for number, (_name, statements, reads) in enumerate(cases):
        lines += ['', f'Private Function Case{number}(wb As Workbook) As String',
                  '    Dim tp As CustomXMLPart, tn As CustomXMLNode, k As Sink',
                  '    Dim p As Object, n As Object, b As Variant, s As String, v As Variant', '    On Error Resume Next']
        for statement in statements:
            lines += ['    Err.Clear', f'    {statement}',
                      '    If Err.Number <> 0 Then s = s & "[" & Err.Number & ":" & Err.Description & "]"']
        for read in reads:
            lines += ['    Err.Clear', '    v = Empty', f'    v = {read}',
                      '    If Err.Number <> 0 Then s = s & "|E" & Err.Number & ":" & Err.Description '
                      'Else s = s & "|" & TypeName(v) & ":" & CStr(v)']
        lines += [f'    Case{number} = s', 'End Function']
    return '\n'.join(lines) + '\n'


def saved(path: Path) -> dict[str, str]:
    """The parts under customXml/, the workbook's relationships to them, and their content types."""
    if not path.exists():
        return {}
    found: dict[str, str] = {}
    with zipfile.ZipFile(path) as book:
        for name in sorted(book.namelist()):
            if name.startswith('customXml/'):
                found[name] = book.read(name).decode('utf-8', errors='replace')
        rels = book.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        found['workbook rels'] = ''.join(re.findall(r'<Relationship [^>]*customXml[^>]*/>', rels))
        types = book.read('[Content_Types].xml').decode('utf-8')
        found['content types'] = ''.join(re.findall(r'<Override [^>]*customXml[^>]*/>', types))
    return found


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    stem = 'custom_xml' + (sys.argv[1] if len(sys.argv) > 1 else '')
    predicted = FIXTURES / f'{stem}_predictions.json'
    cases = ([tuple(case['call']) for case in json.loads(predicted.read_text('utf-8'))['cases']]
             if predicted.exists() else {'_b': CASES_B, '_c': CASES_C}.get(stem[len('custom_xml'):], CASES))
    text = source(cases)  # type: ignore[arg-type]
    answers: list[str] = []
    files: list[dict[str, str]] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.add_module('Sink', SINK, kind='class')
        for number in range(len(cases)):
            path = Path(folder) / f'c{number}.xlsx'
            result = excel.run_vba(text, 'Go', (number, str(path)), timeout=120.0, line_numbers=False)
            answers.append(str(result.value) if result.ok and not result.dialogs
                           else f'ERROR {result.ok} dialogs={len(result.dialogs)} {str(result)[:300]}')
            files.append(saved(path))
    (FIXTURES / f'{stem}.json').write_text(json.dumps({
        'measured_at': datetime.now(timezone.utc).isoformat(), 'source': text, 'sink': SINK,
        'cases': [[*case, answer, found] for case, answer, found in zip(cases, answers, files, strict=True)]},
        indent=1) + '\n', encoding='utf-8', newline='\n')
    print(len(answers), 'asked,', sum(answer.startswith('ERROR') for answer in answers), 'errors')
    if predicted.exists():
        guid = re.compile(r'\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}')

        def same(case: dict[str, Any], answer: str, found: dict[str, str]) -> bool:
            norm = lambda text: guid.sub('{ID}', text.replace('\r\n', '\n'))  # noqa: E731
            return norm(case['model']) == norm(answer) and \
                {key: norm(value) for key, value in case['files'].items()} == {key: norm(value) for key, value in found.items()}

        predictions = json.loads(predicted.read_text('utf-8'))['cases']
        wrong = [(case['call'][0], case['model'][:300], answer[:300]) for case, answer, found in
                 zip(predictions, answers, files, strict=True) if case['model'] is not None and not same(case, answer, found)]
        refused = sum(case['model'] is None for case in predictions)
        print(f'{len(answers) - len(wrong) - refused} right, {len(wrong)} wrong, {refused} refused of {len(answers)}')
        for one in wrong[:10]:
            print(one)


def model_answer(cases: list[Any], number: int, folder: Path) -> tuple[str | None, dict[str, str]]:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SINK, name='Sink', kind='class')
    app.add_module(source(cases), name='Probe')
    path = folder / f'c{number}.xlsx'
    path.unlink(missing_ok=True)
    try:
        answer = str(app.run('Probe.Go', number, str(path)))
    except VBAUnsupportedError:
        return None, {}
    return answer, saved(path)


def _tree(pick: Any, depth: int, names: str) -> str:
    """A random element: a name from ``names``, up to two attributes, text or children."""
    name = pick.choice(names)
    attributes = ''.join(f' {key}="{pick.randint(0, 3)}"' for key in pick.sample('kmq', pick.randint(0, 2)))
    if depth == 0 or pick.random() < 0.3:
        body = pick.choice(['', 'w', str(pick.randint(0, 9)), 'a b', 'x&amp;y'])
    else:
        body = ''.join(_tree(pick, depth - 1, names) for _ in range(pick.randint(1, 3)))
    return f'<{name}{attributes}/>' if not body else f'<{name}{attributes}>{body}</{name}>'


def held_out(round_: str) -> list[tuple[str, list[str], list[str]]]:
    """Fresh cases for the round: a random part, with a default namespace in one of three, then node edits, AddNode
    held typed, a prefix added, and reads by XPath drawn from the part's own names and its nodes' members."""
    import random

    pick = random.Random('custom_xml' + round_)
    cases: list[tuple[str, list[str], list[str]]] = []
    for number in range(60):
        names = 'abcd'
        namespaced = pick.random() < 0.34
        inner = ''.join(_tree(pick, 2, names) for _ in range(pick.randint(1, 4)))
        root = f'<r xmlns="urn:h">{inner}</r>' if namespaced else f'<r>{inner}</r>'
        quoted = root.replace('"', '""')
        p = 'ns0:' if namespaced else ''
        statements = [f'Set p = ActiveWorkbook.CustomXMLParts.Add("{quoted}")' if pick.random() < 0.8 else
                      f'Set p = ActiveWorkbook.CustomXMLParts.Add(): b = p.LoadXML("{quoted}")',
                      'Set tp = p', 'Set tn = tp.DocumentElement']
        target = f'//{p}{pick.choice(names)}'
        for _ in range(pick.randint(0, 3)):
            statements.append(pick.choice([
                f'tp.AddNode tn, "{pick.choice(names)}", "{"urn:h" if namespaced and pick.random() < 0.5 else ""}", , '
                f'msoCustomXMLNodeElement, "{pick.choice(["", "v", "1"])}"',
                f'tp.AddNode tn, "z{pick.randint(0, 9)}", "", , msoCustomXMLNodeAttribute, "{pick.randint(0, 9)}"',
                'tp.AddNode tn, "", "", , msoCustomXMLNodeComment, "note"',
                'tp.AddNode tn, "pi", "", , msoCustomXMLNodeProcessingInstruction, "d"',
                'tp.AddNode tn, "e", "", tn.FirstChild, msoCustomXMLNodeElement, "first"',
                f'tp.SelectSingleNode("{target}").Delete',
                f'tp.SelectSingleNode("{target}").Text = "t{pick.randint(0, 9)}"',
                f'tp.SelectSingleNode("{target}").AppendChildNode "n", "{"urn:h" if namespaced else ""}", 1, "c"',
                f'tp.SelectSingleNode("{target}").AppendChildSubtree "<s><u/></s>"' if not namespaced else
                'tp.NamespaceManager.AddNamespace "h", "urn:h"',
                'tn.RemoveChild tn.LastChild',
                'tn.ReplaceChildNode tn.FirstChild, "y", "" , 1, "8"' if not namespaced else 'tn.Text = "flat"',
                f'tn.InsertNodeBefore "w", "{"urn:h" if namespaced else ""}", 1, "9", tn.LastChild',
                'tp.SelectSingleNode("//@*").NodeValue = "5"',
            ]))
        paths = [f'/{p}r', f'//{p}{pick.choice(names)}', f'//{p}{pick.choice(names)}[1]', f'//{p}*[2]', '//@*',
                 f'//{p}{pick.choice(names)}[@{pick.choice("kmq")}]', f'//{p}{pick.choice(names)}[@k=\'{pick.randint(0, 3)}\']',
                 f'/{p}r/{p}*[last()]', '//text()', '//node()', f'//{p}{pick.choice(names)}/..', f'count(//{p}*)',
                 f'//{p}{pick.choice(names)}[position()>1]', f'/{p}r/*', f'//{p}*[{p}{pick.choice(names)}]',
                 f'//{p}{pick.choice(names)} | //{p}{pick.choice(names)}', '//comment()', '//processing-instruction()',
                 f'/{p}r/{p}*[1]/following-sibling::*', f'descendant::{p}{pick.choice(names)}',
                 f'//{p}*[text()=\'w\']', f'//{p}*[not(*)]', f'//{p}*[contains(., \'1\')]', f'//{p}{pick.choice(names)}[',
                 '/r', '//h:*', f'//{p}*[number(.)>2]']
        chosen = pick.sample(paths, 5)
        reads = ['p.XML', 'p.NamespaceURI', 'p.NamespaceManager.Count']
        reads += [f'p.SelectNodes("{path}").Count' for path in chosen]
        node = pick.choice(chosen[:3])
        reads += [f'TypeName(p.SelectSingleNode("{node}"))',
                  NODE.format(f'p.SelectSingleNode("{node}")'),
                  f'p.SelectSingleNode("{node}").XML', f'p.SelectSingleNode("{node}").NodeValue',
                  f'p.SelectSingleNode("{node}").ChildNodes.Count', f'p.SelectSingleNode("{node}").Attributes.Count',
                  f'TypeName(p.SelectSingleNode("{node}").NextSibling)', 'p.DocumentElement.LastChild.XPath']
        cases.append((f'held {number}', statements, reads))
    return cases


def predict(round_: str) -> None:
    import tempfile

    cases = held_out(round_)
    folder = Path(tempfile.mkdtemp())
    found = []
    for number, case in enumerate(cases):
        answer, files = model_answer(cases, number, folder)
        found.append({'call': list(case), 'model': answer, 'files': files})
    (FIXTURES / f'custom_xml{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0),
                                                                   encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused')


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == 'predict':
        predict(sys.argv[2])
    else:
        main()
