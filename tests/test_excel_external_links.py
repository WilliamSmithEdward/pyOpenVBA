"""Formulas that read another, closed workbook, from the copy of its cells the link keeps.

Measured by scripts/measure_external_links.py: Dest.xlsx read Source.xlsx
and was saved with both open, then reopened alone, its links not updated.
Range.Formula and FormulaR1C1 show the link as the path of the workbook it
points at; the formulas come to what the link keeps; a calculation works
out the cells that changed, COUNTIF over the link keeping its cached value
until a full calculation makes it #VALUE!; a macro may name the link by its
path or its file's name; and a row put in moves the workbook's own
references but not the link's. pyOfficeEditor's 25 cases, measured the same
way, are here too (its tests/test_excel_external_links.py, commit 691adba).
"""
from __future__ import annotations

import base64
import io
import json
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import ERRORS

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/external_links.json').read_text(encoding='utf-8'))
COUNT = len(RECORD['formulas'])


def _dest(folder: Path) -> Path:
    """The Dest.xlsx Excel saved, in ``folder``, beside where Source.xlsx was."""
    path = folder / 'Dest.xlsx'
    with zipfile.ZipFile(path, 'w') as package:
        for name, text in RECORD['dest_parts'].items():
            data = base64.b64decode(text[len('base64:'):]) if text.startswith('base64:') \
                else text.replace('<folder>', str(folder)).encode('utf-8')
            package.writestr(name, data)
    return path


def _probe() -> str:
    """The measuring macro's reading of the reopened workbook, its calculations, its writes and the row put in."""
    return f'''Function Probe() As String
Dim ws As Worksheet, s As String, i As Long
Set ws = ActiveSheet
For i = 1 To {COUNT}
s = s & "#" & ws.Cells(i, 1).Formula & "~" & ws.Cells(i, 1).FormulaR1C1 & "~" & CStr(ws.Cells(i, 1).Value)
Next
ws.Range("B1").Value = 20
Application.Calculate
For i = 1 To {COUNT}
s = s & "~" & CStr(ws.Cells(i, 1).Value)
Next
Application.CalculateFull
For i = 1 To {COUNT}
s = s & "~" & CStr(ws.Cells(i, 1).Value)
Next
ws.Range("C1").Formula = "='" & ActiveWorkbook.Path & "\\[Source.xlsx]Sheet1'!A3"
ws.Range("C2").Formula = "=[Source.xlsx]Sheet1!A4+1"
s = s & "@" & ws.Range("C1").Formula & "~" & CStr(ws.Range("C1").Value) & "~" & ws.Range("C2").Formula
s = s & "~" & CStr(ws.Range("C2").Value)
ws.Rows(1).Insert
s = s & "@" & ws.Range("A3").Formula & "~" & ws.Range("D2").Formula & "~" & CStr(ws.Range("A3").Value)
Probe = Mid(s, 2)
End Function
'''


def test_a_closed_workbook_as_excel_reads_it(tmp_path: Path) -> None:
    app = ExcelApplication.open(_dest(tmp_path))
    app.add_module(_probe(), name='Harness')
    got = str(app.run('Harness.Probe')).replace(str(tmp_path), '<folder>')
    assert got == RECORD['result']


def test_saving_keeps_the_link(tmp_path: Path) -> None:
    source = _dest(tmp_path)
    app = ExcelApplication.open(source)
    app.sheet(1).set_value('B1', 30)
    saved = app.save(tmp_path / 'Saved.xlsx')
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(saved) as after:
        link = before.read('xl/externalLinks/externalLink1.xml').decode('utf-8')
        # A formula reads one of Source's names, =Source.xlsx!Rate*100, so the link gains the sheet Excel adds on
        # saving it closed, named after Source (tests/fixtures/workbook_links.json, NamedAgain.xlsx).
        assert after.read('xl/externalLinks/externalLink1.xml').decode('utf-8') == link.replace(
            '</sheetNames>', '<sheetName val="Source"/></sheetNames>').replace(
            '</sheetDataSet>', '<sheetData sheetId="2" refreshError="1"/></sheetDataSet>')
        name = 'xl/externalLinks/_rels/externalLink1.xml.rels'
        assert after.read(name) == before.read(name)
        assert b'<f>[1]Sheet1!A2*B1</f>' in after.read('xl/worksheets/sheet1.xml')
    reopened = ExcelApplication.open(saved)
    assert reopened.sheet(1).value('A2') == 60
    assert reopened.sheet(1).formula('A2') == f"='{tmp_path}\\[Source.xlsx]Sheet1'!A2*B1"


def test_a_new_link_is_not_made(tmp_path: Path) -> None:
    app = ExcelApplication.open(_dest(tmp_path))
    with pytest.raises(VBAUnsupportedError):
        app.sheet(1).set_value('E1', '=[Other.xlsx]Sheet1!A1')


#: pyOfficeEditor's link, as Excel wrote it but for the absolute path newer Excel records beside the relative one.
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
LINK = (
    f'<externalLink xmlns="{MAIN}"><externalBook xmlns:r="{RELS}" r:id="rId1">'
    '<sheetNames><sheetName val="Sheet1"/><sheetName val="Other"/></sheetNames>'
    "<definedNames><definedName name=\"Rate\" refersTo=\"='Sheet1'!$C$1\"/></definedNames>"
    '<sheetDataSet><sheetData sheetId="0">'
    '<row r="1"><cell r="A1"><v>1</v></cell><cell r="B1" t="str"><v>a</v></cell><cell r="C1"><v>10</v></cell>'
    '<cell r="D1" t="str"><v>x</v></cell></row>'
    '<row r="2"><cell r="A2"><v>2</v></cell><cell r="B2" t="str"><v>b</v></cell><cell r="C2"><v>20</v></cell></row>'
    '<row r="3"><cell r="A3"><v>3</v></cell><cell r="B3" t="str"><v>c</v></cell><cell r="C3"><v>30</v></cell></row>'
    '<row r="4"><cell r="A4"><v>4</v></cell><cell r="B4" t="str"><v>d</v></cell><cell r="C4"><v>40</v></cell></row>'
    '<row r="5"><cell r="A5"><v>5</v></cell><cell r="B5" t="str"><v>e</v></cell><cell r="C5"><v>50</v></cell></row>'
    '</sheetData><sheetData sheetId="1"><row r="1"><cell r="A1"><v>7</v></cell></row></sheetData>'
    '</sheetDataSet></externalBook></externalLink>'
)
#: Each formula in the file's own spelling, and what Excel showed after B1 changed to 20 and everything
#: recalculated with the linked workbook closed.
FORMULAS: list[tuple[str, object]] = [
    ('[1]Sheet1!A1', 1.0), ('[1]Sheet1!A2*B1', 40.0), ('SUM([1]Sheet1!A1:A5)', 15.0),
    ('VLOOKUP("c",[1]Sheet1!B1:C5,2,FALSE)', 30.0), ('INDEX([1]Sheet1!A1:A5,3)', 3.0), ('[1]!Rate*100', 1000.0),
    ('[1]Other!A1*2', 14.0), ('SUM([1]Sheet1:Other!A1)', 8.0), ('[1]Sheet1!A6', 0.0), ('[1]Sheet1!D1', 'x'),
    ('ISREF([1]Sheet1!A1)', True), ('SUMPRODUCT([1]Sheet1!A1:A5*(B1+1))', 315.0), ('SUM([1]Sheet1!A:A)', 15.0),
    ('COUNTBLANK([1]Sheet1!A1:A6)', 1.0), ('RANK(3,[1]Sheet1!A1:A5)', 3.0), ('SUBTOTAL(9,[1]Sheet1!A1:A5)', 15.0),
    ('ROW([1]Sheet1!A3)', 3.0), ('COUNTIF([1]Sheet1!A1:A5,">2")', '#VALUE!'),
    ('SUMIFS([1]Sheet1!A1:A5,[1]Sheet1!B1:B5,"c")', '#VALUE!'), ('DSUM([1]Sheet1!A1:C5,3,[1]Sheet1!A1:A2)', '#VALUE!'),
    ('OFFSET([1]Sheet1!A1,1,0)', '#VALUE!'), ('INDIRECT("[Source.xlsx]Sheet1!A1")', '#REF!'),
    ('CELL("address",[1]Sheet1!A2)', '#N/A'), ('SHEET([1]Sheet1!A1)', '#N/A'), ('ISFORMULA([1]Sheet1!A1)', '#N/A'),
]


def _escaped(text: str) -> str:
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def _package() -> bytes:
    cells = ''.join(f'<row r="{row}"><c r="A{row}"><f>{_escaped(formula)}</f></c>'
                    + ('<c r="B1"><v>10</v></c>' if row == 1 else '') + '</row>'
                    for row, (formula, _) in enumerate(FORMULAS, start=1))
    parts = {
        '[Content_Types].xml': (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/externalLinks/externalLink1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>'),
        '_rels/.rels': ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                        f'<Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/>'
                        '</Relationships>'),
        'xl/workbook.xml': (f'<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets>'
                            '<sheet name="Dest" sheetId="1" r:id="rId1"/></sheets>'
                            '<externalReferences><externalReference r:id="rId2"/></externalReferences></workbook>'),
        'xl/_rels/workbook.xml.rels': (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{RELS}/externalLink" Target="externalLinks/externalLink1.xml"/>'
            '</Relationships>'),
        'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{MAIN}"><sheetData>{cells}</sheetData></worksheet>',
        'xl/externalLinks/externalLink1.xml': LINK,
        'xl/externalLinks/_rels/externalLink1.xml.rels': (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{RELS}/externalLinkPath" Target="Source.xlsx" TargetMode="External"/>'
            '</Relationships>'),
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w') as package:
        for name, text in parts.items():
            package.writestr(name, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + text)
    return buffer.getvalue()


def test_pyofficeeditors_cases(tmp_path: Path) -> None:
    path = tmp_path / 'Dest.xlsx'
    path.write_bytes(_package())
    app = ExcelApplication.open(path)
    app.sheet(1).set_value('B1', 20)
    wanted = [ERRORS[want] if isinstance(want, str) and want.startswith('#') else want for _, want in FORMULAS]
    got = [app.workbook.calculator.value_of('Dest', row, 1) for row in range(1, len(FORMULAS) + 1)]
    assert got == wanted
