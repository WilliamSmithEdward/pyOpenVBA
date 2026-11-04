"""ConnectorFormat.BeginConnect, EndConnect and their disconnects: where each end of a connector goes, what it reads, the
errors, and the connector each saved workbook holds.

    python scripts/measure_shape_connectors.py

writes tests/fixtures/shape_connectors.json, which tests/test_excel_shape_connectors.py replays. Anchors follow the display's
scale (rows read 14.5 points at some scales), so they are left out; the connector's own xfrm is what is compared.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_connectors.json'
MAKE = ('Set a = ws.Shapes.AddShape(1, 100, 100, 80, 40): Set b = ws.Shapes.AddShape(1, 300, 200, 60, 50): '
        'Set c = ws.Shapes.AddShape(9, 400, 50, 60, 60): Set t = ws.Shapes.AddShape(7, 50, 300, 60, 60): '
        'Set r = ws.Shapes.AddShape(5, 200, 300, 80, 40): Set x = ws.Shapes.AddTextbox(1, 500, 300, 80, 40): ')
STRAIGHT = 'Set sh = ws.Shapes.AddConnector(1, 10, 10, 50, 30): '
ELBOW = 'Set sh = ws.Shapes.AddConnector(2, 10, 10, 50, 30): '
CURVE = 'Set sh = ws.Shapes.AddConnector(3, 10, 10, 50, 30): '
LINE = 'Set sh = ws.Shapes.AddLine(10, 10, 50, 30): '
F = 'sh.ConnectorFormat.'
BEGIN_A = f'{F}BeginConnect a, 4: '
END_B = f'{F}EndConnect b, 2'
CASES = [
    f'{STRAIGHT}s2 = 0',
    *[f'{STRAIGHT}{F}BeginConnect a, {site}' for site in (1, 2, 3, 4)], *[f'{STRAIGHT}{F}EndConnect a, {site}' for site in (1, 2, 3, 4)],
    *[f'{STRAIGHT}{F}BeginConnect a, {site}' for site in (0, 5, -1)], *[f'{STRAIGHT}{F}EndConnect a, {site}' for site in (0, 5)],
    *[f'{STRAIGHT}{F}BeginConnect c, {site}' for site in (1, 2, 3, 4, 5, 6, 7, 8, 9)],
    *[f'{STRAIGHT}{F}BeginConnect t, {site}' for site in (1, 2, 3, 4, 5, 6, 7)],
    *[f'{STRAIGHT}{F}BeginConnect r, {site}' for site in (1, 2, 3, 4, 5)],
    *[f'{STRAIGHT}{F}BeginConnect x, {site}' for site in (1, 2, 3, 4, 5)],
    # Both ends, in each quarter, so the flips show.
    f'{STRAIGHT}{BEGIN_A}{END_B}', f'{STRAIGHT}{END_B}: {BEGIN_A.rstrip(": ")}', f'{STRAIGHT}{F}BeginConnect b, 2: {F}EndConnect a, 4',
    f'{STRAIGHT}{F}BeginConnect a, 1: {F}EndConnect c, 3', f'{STRAIGHT}{F}BeginConnect c, 3: {F}EndConnect a, 1',
    f'{STRAIGHT}{F}BeginConnect a, 3: {F}EndConnect t, 1', f'{STRAIGHT}{F}BeginConnect t, 1: {F}EndConnect a, 3',
    f'{STRAIGHT}{F}BeginConnect a, 1: {F}EndConnect a, 3', f'{STRAIGHT}{F}BeginConnect a, 2: {F}EndConnect a, 2', f'{STRAIGHT}{F}BeginConnect a, 1: {F}EndConnect a, 1',
    f'{STRAIGHT}{F}BeginConnect a, 1: {F}BeginConnect b, 2', f'{STRAIGHT}{F}BeginConnect a, 1: {F}BeginConnect a, 3',
    # Other kinds of connector and line.
    f'{ELBOW}{BEGIN_A}{END_B}', f'{ELBOW}{F}BeginConnect a, 4', f'{ELBOW}{F}EndConnect b, 2', f'{ELBOW}{F}BeginConnect a, 1: {F}EndConnect b, 1',
    f'{ELBOW}{F}BeginConnect a, 3: {F}EndConnect b, 1', f'{ELBOW}{F}BeginConnect a, 4: {F}EndConnect b, 4',
    f'{CURVE}{BEGIN_A}{END_B}', f'{CURVE}{F}BeginConnect a, 4', f'{CURVE}{F}BeginConnect a, 3: {F}EndConnect b, 1',
    f'{LINE}{BEGIN_A}{END_B}', f'{LINE}{F}BeginConnect a, 1',
    # Disconnecting.
    f'{STRAIGHT}{BEGIN_A}{F}BeginDisconnect', f'{STRAIGHT}{END_B}: {F}EndDisconnect', f'{STRAIGHT}{BEGIN_A}{END_B}: {F}BeginDisconnect', f'{STRAIGHT}{BEGIN_A}{END_B}: {F}EndDisconnect',
    f'{STRAIGHT}{BEGIN_A}{END_B}: {F}BeginDisconnect: {F}EndDisconnect', f'{STRAIGHT}{F}BeginDisconnect', f'{STRAIGHT}{F}EndDisconnect',
    f'{STRAIGHT}{BEGIN_A}{F}BeginDisconnect: {F}BeginDisconnect', f'{STRAIGHT}{BEGIN_A}{F}BeginDisconnect: {F}BeginConnect b, 1',
    # Wrong kinds of thing.
    f'{STRAIGHT}{F}BeginConnect sh, 1', f'{STRAIGHT}Set o = ws.Shapes.AddLine(1, 1, 2, 2): {F}BeginConnect o, 1', f'{STRAIGHT}Set o = ws.Shapes.AddConnector(1, 1, 1, 2, 2): {F}BeginConnect o, 1',
    f'{STRAIGHT}{F}BeginConnect "a", 1', f'{STRAIGHT}{F}BeginConnect a', f'{STRAIGHT}{F}EndConnect a', f'{STRAIGHT}{F}BeginConnect 1, 1', f'{STRAIGHT}{F}BeginConnect Nothing, 1',
    f'{STRAIGHT}{F}BeginConnect ConnectionSite:=2, ConnectedShape:=a',
    'Set sh = a: sh.ConnectorFormat.BeginConnect b, 1', 'Set sh = a: sh.ConnectorFormat.BeginDisconnect', 'Set sh = a: s2 = sh.ConnectorFormat.BeginConnectedShape.Name',
    # What moves with what.
    *[f'{STRAIGHT}{BEGIN_A}{END_B}: {one}' for one in (
        'a.Left = 150', 'a.Top = 20', 'a.Width = 200', 'a.Height = 100', 'b.Left = 20', 'a.IncrementLeft 10', 'a.IncrementTop 10', 'a.ScaleWidth 2, msoFalse', 'a.Rotation = 90', 'a.Rotation = 45', 'a.Flip msoFlipHorizontal',
        'a.Delete', 'b.Delete', 'a.Delete: b.Delete', 'sh.Left = 5', 'sh.Top = 5', 'sh.Width = 5', 'sh.Height = 5', 'sh.IncrementLeft 7', 'sh.Flip msoFlipHorizontal', 'sh.Rotation = 90', 'sh.RerouteConnections', 'a.Duplicate',
        'sh.Duplicate', 'a.AutoShapeType = 9', 'b.Left = 20: a.Top = 300', 'a.Name = "Renamed"', 'a.ZOrder msoBringToFront', 'sh.ConnectorFormat.Type = msoConnectorCurve')],
    *[f'{ELBOW}{BEGIN_A}{END_B}: {one}' for one in ('a.Left = 150', 'b.Left = 20', 'sh.RerouteConnections', 'sh.ConnectorFormat.Type = msoConnectorCurve')],
    f'{STRAIGHT}a.Rotation = 90: {BEGIN_A.rstrip(": ")}', f'{STRAIGHT}a.Flip msoFlipHorizontal: {BEGIN_A.rstrip(": ")}', f'{STRAIGHT}a.Rotation = 45: {F}BeginConnect a, 1',
    f'{STRAIGHT}{F}BeginConnect a, 1: {F}EndConnect b, 1: sh.RerouteConnections', f'{STRAIGHT}{F}BeginConnect a, 1: sh.RerouteConnections', f'{STRAIGHT}sh.RerouteConnections',
    f'{STRAIGHT}{F}BeginConnect a, 2: {F}EndConnect b, 4: sh.RerouteConnections', f'{LINE}sh.RerouteConnections', 'Set sh = a: sh.RerouteConnections',
    # Which sites RerouteConnections picks: my rule says the nearest pair, so a3 and t1, a4 and c3, r2 and t5, r1 and a3, and c5 and b1 or b4.
    f'{STRAIGHT}{F}BeginConnect a, 1: {F}EndConnect t, 3: sh.RerouteConnections', f'{STRAIGHT}{F}BeginConnect a, 1: {F}EndConnect c, 1: sh.RerouteConnections',
    f'{STRAIGHT}{F}BeginConnect r, 1: {F}EndConnect t, 1: sh.RerouteConnections', f'{STRAIGHT}{F}BeginConnect r, 3: {F}EndConnect a, 1: sh.RerouteConnections',
    f'{STRAIGHT}{F}BeginConnect c, 1: {F}EndConnect b, 3: sh.RerouteConnections', f'{STRAIGHT}{F}BeginConnect b, 1: {F}EndConnect c, 2: sh.RerouteConnections',
    f'{STRAIGHT}{F}BeginConnect a, 1: {F}EndConnect a, 3: sh.RerouteConnections',
    # A follow after the connector's own size was set, and a second move.
    f'{STRAIGHT}{BEGIN_A}{END_B}: sh.Width = 5: a.Left = 150', f'{STRAIGHT}{BEGIN_A}{END_B}: sh.Height = 5: b.Top = 300', f'{STRAIGHT}{BEGIN_A}{END_B}: a.Left = 150: a.Left = 90',
    f'{STRAIGHT}{BEGIN_A}{END_B}: {F}BeginDisconnect: a.Left = 150', f'{STRAIGHT}{BEGIN_A}{END_B}: {F}BeginDisconnect: b.Left = 500',
    # ShapeRange.RerouteConnections: my rule says each connector in it reroutes, and anything else in it raises as a shape alone does.
    f'{STRAIGHT}{F}BeginConnect a, 1: {F}EndConnect b, 1: ws.Shapes.Range(Array(sh.Name)).RerouteConnections', f'{STRAIGHT}ws.Shapes.Range(Array(sh.Name)).RerouteConnections',
    f'{STRAIGHT}{F}BeginConnect a, 1: ws.Shapes.Range(Array(sh.Name)).RerouteConnections', f'{STRAIGHT}{F}BeginConnect a, 1: {F}EndConnect b, 1: ws.Shapes.Range(Array(sh.Name, a.Name)).RerouteConnections',
    f'{STRAIGHT}{F}BeginConnect a, 1: {F}EndConnect b, 1: ws.Shapes.Range(Array(a.Name, sh.Name)).RerouteConnections', f'{STRAIGHT}ws.Shapes.Range(Array(a.Name)).RerouteConnections',
    f'{STRAIGHT}Set o = ws.Shapes.AddConnector(1, 1, 1, 2, 2): {F}BeginConnect a, 1: {F}EndConnect b, 1: o.ConnectorFormat.BeginConnect a, 3: o.ConnectorFormat.EndConnect c, 1: ws.Shapes.Range(Array(sh.Name, o.Name)).RerouteConnections',
    f'{STRAIGHT}Set o = ws.Shapes.AddConnector(1, 1, 1, 2, 2): {F}BeginConnect a, 1: {F}EndConnect b, 1: ws.Shapes.Range(Array(sh.Name, o.Name)).RerouteConnections',
    # The Creator of each object here.
    *[f'{STRAIGHT}s2 = {one}.Creator' for one in ('sh.ConnectorFormat', 'sh.TextFrame2.TextRange', 'sh.TextFrame2.TextRange.Font', 'a.PictureFormat', 'a.ConnectorFormat')],
    # A group's member.
    f'{STRAIGHT}Set o = ws.Shapes.Range(Array(a.Name, b.Name)).Group: {F}BeginConnect o, 1', f'{STRAIGHT}Set o = ws.Shapes.Range(Array(a.Name, b.Name)).Group: {F}BeginConnect o.GroupItems(1), 1',
]
READS = [f'{F}BeginConnected', f'{F}EndConnected', f'{F}BeginConnectedShape.Name', f'{F}EndConnectedShape.Name', f'{F}BeginConnectionSite', f'{F}EndConnectionSite',
         'sh.Left', 'sh.Top', 'sh.Width', 'sh.Height', 'sh.HorizontalFlip', 'sh.VerticalFlip', 'sh.Rotation', 'sh.Connector', 'sh.Name', 'sh.AutoShapeType', f'{F}Type',
         'sh.Adjustments.Count', 'sh.Adjustments.Item(1)', 'a.Left', 'a.Top', 'a.Width', 'a.Height', 'a.Connector', 'ws.Shapes.Count', 's2']


def _probe() -> str:
    lines: list[str] = []
    chunks = [range(at, min(at + 50, len(CASES))) for at in range(0, len(CASES), 50)]
    # Each in a procedure of its own, as one holding them all is too large to compile, taking its own error: passed up through
    # a call, Excel's -2147024809 turns into VBA's 5.
    for number, chunk in enumerate(chunks):
        lines += [f'Private Sub Act{number}(ByVal which As Long, ByVal ws As Worksheet, a As Object, b As Object, c As Object, t As Object, r As Object, x As Object, sh As Object, o As Object, s2 As Variant, failure As Long)',
                  '    On Error GoTo Failed', f'    {MAKE.rstrip(": ")}', '    Select Case which']
        for index in chunk:
            lines += [f'    Case {index}', f'        {CASES[index]}']
        lines += ['    End Select', '    Exit Sub', 'Failed:', '    failure = Err.Number', 'End Sub', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
              '    Dim wb As Workbook, ws As Worksheet, a As Object, b As Object, c As Object, t As Object, r As Object, x As Object, sh As Object, o As Object, s As String, v As String, s2 As Variant, failure As Long',
              '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)', '    Select Case which \\ 50']
    for number in range(len(chunks)):
        lines += [f'    Case {number}: Act{number} which, ws, a, b, c, t, r, x, sh, o, s2, failure']
    lines += ['    End Select', '    s = IIf(failure = 0, "ok", "Err " & failure)', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "~" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', 'End Function',
              '', 'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=560.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = []
        for index in range(len(CASES)):
            with zipfile.ZipFile(Path(folder) / f'c{index}.xlsx') as package:
                xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
            parts.append(re.findall(r'<xdr:cxnSp\b.*?</xdr:cxnSp>', xml, re.DOTALL))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'make': MAKE, 'source': source,
              'cases': [[statements, answer, part] for statements, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
