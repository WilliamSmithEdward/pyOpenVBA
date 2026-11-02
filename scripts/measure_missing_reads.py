"""Members the model was missing, read on a new workbook: each expression of READS as its text, TypeName for an
object, or "E" and the error number.

The workbook has a rectangle, a column chart of A1:B4 with an axis title and a legend, and nothing else. Reads that
depend on layout (the left, top, width and height of a chart's parts) or on the machine or the user, such as
WriteReservedBy, are left out.

    python scripts/measure_missing_reads.py

writes tests/fixtures/missing_reads.json, which tests/test_excel_missing_reads.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/missing_reads.json'
#: (object variable, member); objects are read as TypeName.
READS = [
    ('wb', 'Password'), ('wb', 'WritePassword'), ('wb', 'PasswordEncryptionProvider'),
    ('wb', 'PasswordEncryptionAlgorithm'), ('wb', 'PasswordEncryptionKeyLength'),
    ('wb', 'PasswordEncryptionFileProperties'), ('wb', 'Container'), ('wb', 'ContentTypeProperties'),
    ('wb', 'ServerPolicy'), ('wb', '_CodeName'), ('wb', 'ActiveChart'), ('wb', 'ActiveSlicer'),
    ('wb', 'ExternalCodeServiceTimeout'), ('wb', 'PivotTables'), ('wb', 'Theme'), ('wb', 'WebOptions'),
    ('wb', 'DocumentLibraryVersions'), ('wb', 'SharedWorkspace'), ('wb', 'SmartDocument'), ('wb', 'Permission'),
    ('ws', 'PrintedCommentPages'),
    ('ch', 'Index'), ('ch', 'CodeName'), ('ch', 'Visible'), ('ch', 'ProtectContents'), ('ch', 'ProtectData'),
    ('ch', 'ProtectDrawingObjects'), ('ch', 'ProtectFormatting'), ('ch', 'ProtectionMode'),
    ('ch', 'ProtectSelection'), ('ch', 'ShowAllFieldButtons'), ('ch', 'ShowAxisFieldButtons'),
    ('ch', 'ShowLegendFieldButtons'), ('ch', 'ShowReportFilterFieldButtons'), ('ch', 'ShowValueFieldButtons'),
    ('ch', 'ShowExpandCollapseEntireFieldButtons'), ('ch', 'Elevation'), ('ch', 'Rotation'), ('ch', 'Perspective'),
    ('ch', 'RightAngleAxes'), ('ch', 'DepthPercent'), ('ch', 'HeightPercent'), ('ch', 'GapDepth'),
    ('ch', 'AutoScaling'), ('ch', 'BarShape'), ('ch', 'CategoryLabelLevel'), ('ch', 'SeriesNameLevel'),
    ('ch', 'ChartStyle'), ('ch', 'ChartColor'), ('ch', 'PlotArea'), ('ch', 'ChartArea'), ('ch', 'Walls'),
    ('ch', 'Floor'), ('ch', 'BackWall'), ('ch', 'SideWall'), ('ch', 'PivotLayout'), ('ch', 'Shapes'),
    ('ch', 'Hyperlinks'), ('ch', 'PageSetup'), ('ch', 'Tab'), ('ch', 'Next'), ('ch', 'Previous'),
    ('ch', 'PrintedCommentPages'),
    ('ax', 'Application'), ('ax', 'CategorySortOrder'), ('ax', 'MinorUnitScale'), ('ax', 'DisplayUnitLabel'),
    ('ax', 'MajorGridlines'), ('ax', 'MinorGridlines'), ('ax', 'TickLabels'), ('ax', 'Border'),
    ('vx', 'MajorGridlines'), ('vx', 'DisplayUnitLabel'), ('vx', 'CategorySortOrder'),
    ('at', 'Application'), ('at', 'Name'), ('le', 'Application'), ('le', 'Index'), ('le', 'Font'),
    ('le', 'LegendKey'),
    ('sr', 'LeaderLines'), ('sr', 'BubbleSizes'),
    ('co', 'ZOrder'), ('co', 'ProtectChartObject'),
    ('sh', 'FormControlType'), ('sh', 'HasSmartArt'), ('sh', 'Child'), ('sh', 'ConnectionSiteCount'),
    ('sh', 'Connector'), ('sh', 'Hyperlink'), ('sh', 'Callout'), ('sh', 'BlackWhiteMode'), ('sh', 'GraphicStyle'),
    ('sh', 'ShapeStyle'), ('sh', 'BackgroundStyle'),
    ('sr2', 'Application'), ('sr2', 'Parent'), ('sr2', 'ID'), ('sr2', 'HasChart'), ('sr2', 'Child'),
    ('sr2', 'Connector'), ('sr2', 'ZOrderPosition'), ('sr2', 'GroupItems'), ('sr2', 'ParentGroup'),
    ('tf', 'Application'), ('tf', 'Parent'),
    ('wb.Sheets', 'HPageBreaks'), ('wb.Sheets', 'VPageBreaks'),
    ('ws.PageSetup', 'PrintQuality'), ('ws.PageSetup', 'Pages'),
    ('ws.Range("A1")', 'Phonetic'), ('ws.Range("A1")', 'Phonetics'), ('ws.Range("A1")', 'XPath'),
    ('ws.Range("A1")', 'SoundNote'),
]
#: Read after LATER_SETUP adds a group of two rectangles, a button and a check box.
LATER_SETUP = [
    '    Set g = ws.Shapes.Range(Array(ws.Shapes.AddShape(msoShapeOval, 10, 320, 40, 40).Name, '
    'ws.Shapes.AddShape(msoShapeOval, 60, 320, 40, 40).Name)).Group',
    '    Set bt = ws.Shapes.AddFormControl(xlButtonControl, 120, 320, 60, 20)',
    '    Set cb = ws.Shapes.AddFormControl(xlCheckBox, 200, 320, 60, 20)',
]
LATER_READS = [
    ('g', 'Child'), ('g.GroupItems(1)', 'Child'), ('g.GroupItems(1)', 'Connector'), ('g', 'FormControlType'),
    ('bt', 'FormControlType'), ('cb', 'FormControlType'), ('bt', 'HasSmartArt'), ('bt', 'BlackWhiteMode'),
    ('ws.Shapes.Range(Array(g.Name))', 'Child'), ('ws.Shapes.Range(Array(g.Name))', 'GroupItems'),
    ('ws.Shapes.Range(Array(sh.Name, g.Name))', 'ID'), ('ws.Shapes.Range(Array(sh.Name, g.Name))', 'HasChart'),
    ('ws.Shapes.Range(Array(sh.Name, g.Name))', 'ZOrderPosition'),
    ('ws.Shapes.Range(Array(sh.Name, g.Name))', 'Parent'),
    ('ws.Shapes.Range(Array(co.Name))', 'HasChart'),
]


def _probe() -> str:
    lines = ['Private Function Fetch(ByVal target As Object, ByVal name As String) As String',
             '    Dim v As Variant', '    On Error GoTo Failed',
             '    If IsObject(CallByName(target, name, VbGet)) Then',
             '        Fetch = TypeName(CallByName(target, name, VbGet))', '    Else',
             '        v = CallByName(target, name, VbGet)',
             '        If IsArray(v) Then Fetch = "Array" Else Fetch = TypeName(v) & ":" & CStr(v)', '    End If',
             '    Exit Function', 'Failed:', '    Fetch = "E" & Err.Number', 'End Function', '',
             'Public Function Probe() As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, co As Object, ax As Object, vx As Object',
             '    Dim at As Object, le As Object, sr As Object, sh As Object, sr2 As Object, tf As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    Set sh = ws.Shapes.AddShape(msoShapeRectangle, 300, 10, 100, 40): Set tf = sh.TextFrame2',
             '    Set sr2 = ws.Shapes.Range(Array(sh.Name))',
             '    Set co = ws.ChartObjects.Add(10, 100, 300, 200): Set ch = co.Chart',
             '    ch.SetSourceData ws.Range("A1:B4"): ch.HasLegend = True',
             '    Set ax = ch.Axes(xlCategory): Set vx = ch.Axes(xlValue)',
             '    ax.HasTitle = True: Set at = ax.AxisTitle', '    Set le = ch.Legend.LegendEntries(1)',
             '    Set sr = ch.SeriesCollection(1)']
    for target, name in READS:
        lines.append(f'    s = s & "@@@" & Fetch({target}, "{name}")')
    lines += ['    Dim g As Object, bt As Object, cb As Object', *LATER_SETUP]
    for target, name in LATER_READS:
        lines.append(f'    s = s & "@@@" & Fetch({target}, "{name}")')
    lines += ['    wb.Close False', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=200.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[f'{target}.{name}', answer]
                        for (target, name), answer in zip(READS + LATER_READS, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for read, answer in record['cases']:
        print(read, '->', answer)


if __name__ == '__main__':
    main()
