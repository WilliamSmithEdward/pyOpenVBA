"""Every Workbook property a macro reads, on a book Excel saved and opened again, bound typed (As Workbook) and late (As Object).

Live Excel builds a small book (three sheets, a defined name, a table, a chart sheet and a style), saves it to a temporary folder as .xlsx,
closes it and opens it again, then reads each property: its TypeName and value, or, for an object, its TypeName and Count where it has one,
or the error it raises. The model opens the same file and runs the same macro, and tests/test_excel_workbook_reads.py compares the two.

Left out: what reaches mail or a mail envelope (Mailer, EnvelopeVisible), servers and rights (Permission, ServerPolicy, SharedWorkspace,
DocumentLibraryVersions, ContentTypeProperties, SensitivityLabel), the Research pane and SmartDocument. Paths and the user's name are recorded
as a yes or no, never as what they are on this machine.

    python scripts/measure_workbook_reads.py

writes tests/fixtures/workbook_reads.json and, for each round, tests/fixtures/workbook_reads_<round>.xlsx or .xlsm, the book both sides
read. A round already in the fixture is kept as measured.
"""
from __future__ import annotations

import json
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/workbook_reads.json'
SEPARATOR = '%%'
#: Not read: mail, servers and rights, panes, and SmartDocument. See the docstring.
LEFT_OUT = frozenset({'Mailer', 'EnvelopeVisible', 'Permission', 'ServerPolicy', 'SharedWorkspace', 'DocumentLibraryVersions', 'ContentTypeProperties',
                      'SensitivityLabel', 'Research', 'SmartDocument'})
#: Read as whether they are blank, since their text is this machine's folder or user.
PRIVATE = frozenset({'FullName', 'FullNameURLEncoded', 'Path', 'UserStatus', 'WriteReservedBy'})


def properties() -> dict[str, str]:
    """Each property and the type it is declared as, from the type library's own signatures in the checklist."""
    import csv

    with (ROOT / 'docs/excel_checklist.csv').open(encoding='utf-8', newline='') as source:
        rows = [row for row in csv.DictReader(source) if row['object'] == 'Workbook' and 'property' in row['kind']]
    return {row['member']: row['signature'].removeprefix('As ').strip() for row in sorted(rows, key=lambda row: row['member'])
            if row['member'] not in LEFT_OUT}


DECLARED = properties()
NAMES = list(DECLARED)
#: Types that hold a value, not an object. Typed, `Set o = wb.Name` does not compile, so a read of one of these has no object branch, and a
#: read of an object has no value branch; a Variant has both.
VALUES = ('Boolean', 'String', 'Long', 'Integer', 'Double', 'Single', 'Date', 'Currency', 'Byte')


def _kind(name: str) -> str:
    declared = DECLARED[name]
    if declared == 'Variant' or not declared:
        return 'either'
    return 'value' if declared in VALUES or declared.startswith(('Xl', 'Mso')) else 'object'


#: The collections, whose Count is read too. Only these: Count of another object is a member it does not have.
COLLECTIONS = frozenset({'BuiltinDocumentProperties', 'Charts', 'CommandBars', 'Connections', 'CustomDocumentProperties', 'CustomViews', 'CustomXMLParts',
                         'DocumentInspectors', 'Excel4IntlMacroSheets', 'Excel4MacroSheets', 'IconSets', 'Names', 'PivotTables', 'PublishObjects',
                         'Queries', 'Sheets', 'Signatures', 'SlicerCaches', 'Styles', 'TableStyles', 'Windows', 'Worksheets', 'XmlMaps',
                         'XmlNamespaces'})
#: Reads per procedure, so none is too large to compile.
PART = 20


def _read(name: str) -> str:
    # A name VBA cannot spell bare, as _CodeName, is written in brackets.
    member = f'[{name}]' if name.startswith('_') else name
    shown = 'Blank(v)' if name in PRIVATE else 'Show(v)'
    counted = ' & CountOf(o)' if name in COLLECTIONS else ''
    as_object = f'''            Set o = wb.{member}: e = Err.Number
            If e = 0 Then s = "obj:" & TypeName(o){counted}
'''
    as_value = f'''            v = wb.{member}: e = Err.Number
            If e = 0 Then s = {shown}
'''
    kind = _kind(name)
    if kind == 'either':
        read = f'''    b = IsObject(wb.{member}): e = Err.Number
    If e = 0 Then
        If b Then
    {as_object.replace(chr(10) + '            ', chr(10) + '                ')}        Else
    {as_value.replace(chr(10) + '            ', chr(10) + '                ')}        End If
    End If
'''
    else:
        read = (as_object if kind == 'object' else as_value).replace('            ', '    ')
    return f'''    Err.Clear
    s = ""
{read}    If e <> 0 Then s = "e" & e
    out = out & "{SEPARATOR}" & s
'''


def macro(declared: str, names: list[str] | None = None) -> str:
    """The probe for these names, all by default; a test leaves out what the model refuses openly, to read the rest."""
    names = NAMES if names is None else names
    parts = [names[at:at + PART] for at in range(0, len(names), PART)]
    procedures = ''.join(f'''Private Sub Part{index}(ByVal wb As {declared}, ByRef out As String)
    Dim o As Object, v As Variant, s As String, e As Long, b As Boolean
    On Error Resume Next
{''.join(_read(name) for name in names)}End Sub
''' for index, names in enumerate(parts))
    calls = ''.join(f'    Part{index} wb, out\n' for index in range(len(parts)))
    return f'''{procedures}Private Function Show(ByVal v As Variant) As String
    If IsArray(v) Then
        Show = "array:" & LBound(v) & "-" & UBound(v)
    ElseIf IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function
Private Function Blank(ByVal v As Variant) As String
    If IsArray(v) Then
        Blank = "array:" & UBound(v, 1) & "x" & UBound(v, 2)
    Else
        Blank = TypeName(v) & ":" & IIf(Len(CStr(v)) = 0, "blank", "set")
    End If
End Function
Private Function CountOf(ByVal o As Object) As String
    On Error Resume Next
    Dim n As Long
    n = -1
    n = o.Count
    If n >= 0 Then CountOf = ":" & n
End Function
Function Probe(ByVal path As String, ByVal readOnly As Boolean) As String
    Dim wb As {declared}, out As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, ReadOnly:=readOnly, IgnoreReadOnlyRecommended:=True)
{calls}    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = Mid(out, {len(SEPARATOR) + 1})
End Function
'''


PLAIN = '''Function Build(ByVal path As String) As String
    Dim wb As Workbook, ws As Worksheet
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"
    ws.Range("A1").Value = "Item": ws.Range("B1").Value = "Amount"
    ws.Range("A2").Value = "a": ws.Range("B2").Value = 1
    ws.Range("A3").Value = "b": ws.Range("B3").Value = 2
    ws.Range("A4").Value = "c": ws.Range("B4").Value = 3
    ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B4"), , xlYes).Name = "Amounts"
    wb.Worksheets.Add(After:=ws).Name = "Second"
    wb.Worksheets.Add(After:=wb.Worksheets(2)).Name = "Third"
    wb.Names.Add Name:="Total", RefersTo:="=SUM(Data!$B$2:$B$4)"
    wb.Charts.Add(After:=wb.Worksheets(3)).SetSourceData ws.Range("A1:B4")
    wb.Styles.Add "Probe style"
    ws.Activate
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Build = "built"
End Function
'''
#: A macro-enabled book with workbook settings away from their defaults, its properties filled in and its structure protected.
SETTINGS = '''Function Build(ByVal path As String) As String
    Dim wb As Workbook, ws As Worksheet
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Figures"
    ws.Range("A1").Value = 1.25: ws.Range("A2").Value = 2.5: ws.Range("A3").Formula = "=A1+A2"
    wb.Worksheets.Add(After:=ws).Name = "Notes"
    wb.Names.Add Name:="Rate", RefersTo:="=0.05"
    wb.Names.Add Name:="Hidden", RefersTo:="=Figures!$A$1", Visible:=False
    wb.Date1904 = True
    wb.CheckCompatibility = True
    wb.ChartDataPointTrack = False
    wb.DisplayDrawingObjects = xlPlaceholders
    wb.InactiveListBorderVisible = False
    wb.UpdateRemoteReferences = False
    wb.SaveLinkValues = False
    wb.ForceFullCalculation = True
    wb.ShowPivotTableFieldList = False
    wb.EnableAutoRecover = False
    wb.RemovePersonalInformation = True
    wb.ConflictResolution = xlLocalSessionChanges
    wb.BuiltinDocumentProperties("Title").Value = "Round two"
    wb.BuiltinDocumentProperties("Subject").Value = "Workbook reads"
    wb.CustomDocumentProperties.Add Name:="Batch", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=7
    wb.CustomDocumentProperties.Add Name:="Owner", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="Desk"
    wb.CustomViews.Add "First view"
    wb.Styles.Add "Second style"
    wb.Protect Structure:=True, Windows:=False
    ws.Activate
    wb.SaveAs path, xlOpenXMLWorkbookMacroEnabled, CreateBackup:=True, ReadOnlyRecommended:=True
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Build = "built"
End Function
'''
#: Each round: the book Excel builds, the file type it is saved as, and whether the probe opens it read-only.
ROUNDS = {'plain': (PLAIN, 'xlsx', True), 'settings': (SETTINGS, 'xlsm', False)}


def book(round_name: str) -> Path:
    return ROOT / f'tests/fixtures/workbook_reads_{round_name}.{ROUNDS[round_name][1]}'


def main() -> None:
    import sys

    from pyvbaharness import ExcelSession, HarnessConfig

    sys.path.insert(0, str(ROOT / 'scripts'))
    from fixture_workbook import copy_saved

    record = json.loads(OUT.read_text(encoding='utf-8')) if OUT.exists() else {'names': NAMES, 'answers': {}}
    assert record['names'] == NAMES, 'the property list changed: measure every round again'
    folder = Path(tempfile.mkdtemp(prefix='wbreads_'))
    try:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            excel.run_vba('Function Hello() As String\n    Hello = "x"\nEnd Function\n', 'Hello', (), timeout=60)
            for round_name, (build, extension, read_only) in ROUNDS.items():
                if round_name in record['answers']:
                    continue
                built_at, path = folder / f'Built.{extension}', folder / f'Probe.{extension}'
                built = excel.run_vba(build, 'Build', (str(built_at),), timeout=120, line_numbers=False)
                assert built.ok and built.value == 'built', (built.error, built.dialogs)
                # Excel writes the save folder into the book; both sides read the copy without it.
                copy_saved(built_at, book(round_name))
                shutil.copyfile(book(round_name), path)
                answers: dict[str, list[str]] = {}
                for declared in ('Workbook', 'Object'):
                    result = excel.run_vba(macro(declared), 'Probe', (str(path), read_only), timeout=300, line_numbers=False)
                    assert result.ok, (result.error, result.dialogs)
                    answers[declared] = str(result.value).split(SEPARATOR)
                record['answers'][round_name] = answers
                record['measured_at'] = datetime.now(timezone.utc).isoformat()
                OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
                same = sum(typed == late for typed, late in zip(answers['Workbook'], answers['Object']))
                print(round_name, len(NAMES), 'properties;', same, 'answered the same typed and late')
    finally:
        shutil.rmtree(folder, ignore_errors=True)


if __name__ == '__main__':
    main()
