"""Font, Interior, Border, Borders and a Range's alignment and protection on cells: each settable property given each kind of value, on a fresh cell and on
one already formatted; what it then reads, and what it reads after a save and reopen; and what a range of two cells that
differ reads.

Each answer is the error from the assignment, then every property of that object, as TypeName and value, in the order
of the family's STATE, Null and Empty spelled out.

    python scripts/measure_format_properties.py

writes tests/fixtures/format_properties.json, which tests/test_excel_format_properties.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/format_properties.json'
BAD = ['Null', 'Empty', 'True', '"x"', '""']
COLOURS = ['0', '255', '16777215', '16777216', '-1', '-2', '1.5', '"255"', *BAD]
INDEXES = ['0', '1', '3', '56', '57', '-1', 'xlNone', 'xlAutomatic', '1.5', '"3"', *BAD]
THEMES = ['1', '2', '12', '0', '13', '-1', '1.5', '"2"', *BAD]
TINTS = ['0', '0.5', '-0.5', '1', '-1', '1.1', '-1.1', '0.123456', '"0.5"', *BAD]
SWITCHES = ['True', 'False', '1', '0', '-1', '2', '"True"', '"x"', 'Null', 'Empty', '""']
LINES = ['xlContinuous', 'xlDash', 'xlDashDot', 'xlDashDotDot', 'xlDot', 'xlDouble', 'xlSlantDashDot', 'xlLineStyleNone',
         '1', '0', '2', '99', '"1"', *BAD[1:]]
#: Null is left out of these two: the assignment ends Excel (the next call fails with -2147417848).
WEIGHTS = ['xlHairline', 'xlThin', 'xlMedium', 'xlThick', '3', '0', '5', 'xlAutomatic', '-1', '"2"', *BAD[1:]]
#: How each family reaches its object, what is read back, and what a formatted cell has first.
FAMILIES = {
    'Font': ('c.Font', ['Name', 'Size', 'Underline', 'ThemeFont', 'Background', 'Color', 'ColorIndex', 'ThemeColor',
                        'TintAndShade', 'Bold', 'Italic', 'FontStyle', 'Strikethrough', 'Subscript', 'Superscript'],
             ['', 'c.Font.Name = "Arial": c.Font.Size = 14: c.Font.Underline = xlUnderlineStyleSingle']),
    'Interior': ('c.Interior', ['Color', 'ColorIndex', 'Pattern', 'PatternColor', 'PatternColorIndex', 'ThemeColor',
                                'TintAndShade', 'PatternThemeColor', 'PatternTintAndShade', 'InvertIfNegative'],
                 ['', 'c.Interior.Color = 66051', 'c.Interior.Pattern = xlGray50: c.Interior.PatternColor = 197121']),
    'Border': ('c.Borders(xlEdgeBottom)', ['LineStyle', 'Weight', 'Color', 'ColorIndex', 'ThemeColor', 'TintAndShade'],
               ['', 'c.Borders(xlEdgeBottom).LineStyle = xlContinuous',
                'c.Borders(xlEdgeBottom).Color = 66051']),
    'Borders': ('c.Borders', ['LineStyle', 'Weight', 'Color', 'ColorIndex', 'ThemeColor', 'TintAndShade'],
                ['', 'c.Borders.LineStyle = xlContinuous', 'c.Borders(xlEdgeTop).LineStyle = xlContinuous']),
    'Range': ('c', ['HorizontalAlignment', 'VerticalAlignment', 'WrapText', 'Orientation', 'IndentLevel', 'AddIndent',
                    'ShrinkToFit', 'ReadingOrder', 'Locked', 'FormulaHidden'],
              ['', 'c.HorizontalAlignment = xlLeft: c.IndentLevel = 2: c.WrapText = True: c.Locked = False']),
}
#: Each family's property given the values to try.
VALUES = {
    'Font': {'Name': ['"Arial"', '"Calibri"', '"Times New Roman"', '"No Such Font Zzz"', '" Arial"', '5', *BAD],
             'Size': ['1', '8', '10.5', '11.25', '11.3', '72', '409', '409.5', '410', '0', '0.5', '-1', '"12"', *BAD],
             'Underline': ['xlUnderlineStyleNone', 'xlUnderlineStyleSingle', 'xlUnderlineStyleDouble',
                           'xlUnderlineStyleSingleAccounting', 'xlUnderlineStyleDoubleAccounting', 'False', '3', '0',
                           '1', '99', '"2"', *BAD],
             'ThemeFont': ['0', '1', '2', '3', '4', '-1', '1.5', '"1"', *BAD],
             'Background': ['1', '2', '3', 'xlAutomatic', '0', *BAD],
             'Bold': SWITCHES, 'Italic': SWITCHES, 'Strikethrough': SWITCHES, 'Subscript': SWITCHES,
             'Superscript': SWITCHES, 'Color': COLOURS, 'ColorIndex': INDEXES, 'ThemeColor': THEMES,
             'TintAndShade': TINTS,
             'FontStyle': ['"Regular"', '"Bold"', '"Italic"', '"Bold Italic"', '"bold"', '"No Such Style"', '5', *BAD]},
    'Interior': {'Color': COLOURS, 'ColorIndex': INDEXES,
                 'Pattern': ['xlSolid', 'xlNone', 'xlAutomatic', 'xlGray50', 'xlChecker', 'xlHorizontal', 'xlLightUp',
                             '0', '2', '99', '"1"', *BAD],
                 'PatternColor': COLOURS, 'PatternColorIndex': INDEXES, 'ThemeColor': THEMES, 'TintAndShade': TINTS,
                 'PatternThemeColor': THEMES, 'PatternTintAndShade': TINTS, 'InvertIfNegative': ['True', 'False', '1']},
    'Border': {'ColorIndex': INDEXES, 'TintAndShade': TINTS, 'Color': COLOURS, 'ThemeColor': THEMES,
               'LineStyle': LINES,
               'Weight': WEIGHTS},
    'Borders': {'ColorIndex': INDEXES, 'ThemeColor': THEMES, 'TintAndShade': TINTS, 'Color': COLOURS, 'LineStyle': LINES,
                'Value': LINES, 'Weight': WEIGHTS},
    'Range': {'HorizontalAlignment': ['xlGeneral', 'xlLeft', 'xlCenter', 'xlRight', 'xlFill', 'xlJustify',
                                      'xlCenterAcrossSelection', 'xlDistributed', '0', '1', '-1', '99', '"-4108"', *BAD],
              'VerticalAlignment': ['xlTop', 'xlCenter', 'xlBottom', 'xlJustify', 'xlDistributed', 'xlLeft', '0', '1',
                                    '-1', '"-4108"', *BAD],
              'WrapText': ['True', 'False', '1', '0', '-1', '2', '"True"', '"x"', 'Null', 'Empty', '""'],
              'Orientation': ['xlHorizontal', 'xlVertical', 'xlUpward', 'xlDownward', '0', '45', '-45', '90', '-90',
                              '91', '-91', '180', '255', '1.5', '"45"', *BAD],
              'IndentLevel': ['0', '1', '15', '250', '251', '-1', '1.5', '"3"', *BAD],
              'AddIndent': ['True', 'False', '1', '0', '"True"', '"x"', 'Null', 'Empty', '""'],
              'ShrinkToFit': ['True', 'False', '1', '0', '"True"', '"x"', 'Null', 'Empty', '""'],
              'ReadingOrder': ['xlContext', 'xlLTR', 'xlRTL', '0', '1', '-1', '"-5002"', *BAD],
              'Locked': ['True', 'False', '1', '0', '"True"', '"x"', 'Null', 'Empty', '""'],
              'FormulaHidden': ['True', 'False', '1', '0', '"True"', '"x"', 'Null', 'Empty', '""']},
}
#: Two cells differing in one property, read through the range: (family, the first's set, the second's set, read).
MIXED = [(family, f'{prop} = {a}', f'{prop} = {b}', prop) for family, prop, a, b in [
    ('Font', 'Name', '"Arial"', '"Calibri"'), ('Font', 'Size', '10', '12'),
    ('Font', 'Underline', 'xlUnderlineStyleSingle', 'xlUnderlineStyleNone'), ('Font', 'ThemeFont', '1', '2'),
    ('Interior', 'Color', '255', '65280'), ('Interior', 'ColorIndex', '3', '4'), ('Interior', 'Pattern', '1', '9'),
    ('Interior', 'PatternColor', '255', '65280'), ('Interior', 'PatternColorIndex', '3', '4'),
    ('Interior', 'ThemeColor', '1', '2'), ('Interior', 'TintAndShade', '0.5', '-0.5'),
    ('Interior', 'PatternThemeColor', '1', '2'), ('Interior', 'PatternTintAndShade', '0.5', '-0.5'),
    ('Borders', 'ColorIndex', '3', '4'), ('Borders', 'ThemeColor', '1', '2'), ('Borders', 'TintAndShade', '0.5', '-0.5'),
    ('Border', 'ColorIndex', '3', '4'), ('Border', 'TintAndShade', '0.5', '-0.5'), ('Border', 'Weight', '2', '4'),
    ('Range', 'HorizontalAlignment', 'xlLeft', 'xlRight'), ('Range', 'VerticalAlignment', 'xlTop', 'xlBottom'),
    ('Range', 'WrapText', 'True', 'False'), ('Range', 'Orientation', '45', '-45'), ('Range', 'IndentLevel', '1', '2'),
    ('Range', 'ShrinkToFit', 'True', 'False'), ('Range', 'ReadingOrder', 'xlLTR', 'xlRTL'),
    ('Range', 'Locked', 'True', 'False'), ('Range', 'FormulaHidden', 'True', 'False'),
]]


def cases() -> list[tuple[str, str, str, str]]:
    """Each (family, what the cell has first, property, value)."""
    return [(family, first, prop, value) for family, values in VALUES.items() for prop, tried in values.items()
            for first in FAMILIES[family][2] for value in tried]


def _state(family: str, name: str) -> str:
    """A function reading all of a family's properties from c."""
    target, props, _ = FAMILIES[family]
    reads = ''.join(f'    v = Empty\n    v = {target}.{prop}\n    If Err.Number <> 0 Then\n'
                    f'        s = s & "|e" & Err.Number\n        Err.Clear\n    Else\n'
                    f'        s = s & "|" & Show(v)\n    End If\n' for prop in props)
    return (f'\nPrivate Function {name}(c As Object) As String\n    Dim s As String, v As Variant\n'
            f'    On Error Resume Next\n{reads}    {name} = s\nEnd Function\n')


def source() -> str:
    """Set(i) makes case i and answers; Reopen saves and opens the book; After(i) reads a case again."""
    every = cases()
    body = ('Private wb As Object, ws As Object\n\nPublic Sub Setup()\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n'
            '    Set ws = wb.Worksheets(1)\nEnd Sub\n'
            '\nPublic Sub Reopen(ByVal folder As String)\n    Dim path As String\n'
            '    path = folder & Application.PathSeparator & "formats.xlsx"\n'
            '    wb.SaveAs path, xlOpenXMLWorkbook\n    wb.Close False\n'
            '    Set wb = Workbooks.Open(path)\n    Set ws = wb.Worksheets(1)\nEnd Sub\n'
            '\nPublic Sub Finish()\n    wb.Close False\nEnd Sub\n'
            '\nPrivate Function Show(v As Variant) As String\n    If IsNull(v) Then\n        Show = "Null"\n'
            '    ElseIf IsEmpty(v) Then\n        Show = "Empty"\n    Else\n        Show = TypeName(v) & " " & CStr(v)\n'
            '    End If\nEnd Function\n')
    body += ''.join(_state(family, f'State{family}') for family in FAMILIES)
    for number, (family, first, prop, value) in enumerate(every):
        target = FAMILIES[family][0]
        row = 2 * number + 2
        body += (f'\nPublic Function S{number}() As String\n    Dim c As Object, e As Long\n'
                 f'    Set c = ws.Cells({row}, 2)\n    c.Value = "x"\n' + (f'    {first}\n' if first else '')
                 + f'    On Error Resume Next\n    {target}.{prop} = {value}\n    e = Err.Number\n    Err.Clear\n'
                 f'    S{number} = "e" & e & State{family}(c)\nEnd Function\n'
                 f'\nPublic Function A{number}() As String\n    A{number} = State{family}(ws.Cells({row}, 2))\nEnd Function\n')
    for number, (family, one, two, prop) in enumerate(MIXED):
        target = FAMILIES[family][0]
        row = 2 * number + 2
        body += (f'\nPublic Function M{number}() As String\n    Dim c As Object, v As Variant\n'
                 f'    Set c = ws.Cells({row}, 6)\n    c.Value = "x"\n    {target}.{one}\n'
                 f'    Set c = ws.Cells({row}, 8)\n    c.Value = "x"\n    {target}.{two}\n'
                 f'    Set c = ws.Range(ws.Cells({row}, 6), ws.Cells({row}, 8))\n    On Error Resume Next\n'
                 f'    v = {target}.{prop}\n    If Err.Number <> 0 Then\n        M{number} = "e" & Err.Number\n    Else\n'
                 f'        M{number} = Show(v)\n    End If\nEnd Function\n')
    return body


def _runner(calls: list[str], name: str) -> str:
    """A function answering each of the calls joined by @@, split up since one procedure cannot hold them all."""
    parts = [calls[at:at + 300] for at in range(0, len(calls), 300)]
    body = ''
    for number, part in enumerate(parts):
        lines = ''.join(f'    out = out & "@@" & {call}()\n' for call in part)
        body += f'\nPrivate Function {name}{number}() As String\n    Dim out As String\n{lines}    {name}{number} = Mid(out, 3)\nEnd Function\n'
    joined = ' & "@@" & '.join(f'{name}{number}()' for number in range(len(parts)))
    return body + f'\nPrivate Function {name}() As String\n    {name} = {joined}\nEnd Function\n'


def main() -> None:
    every = cases()
    probe = source()
    runner = _runner(['S' + str(number) for number in range(len(every))]
                     + ['M' + str(number) for number in range(len(MIXED))], 'First')
    runner += _runner(['A' + str(number) for number in range(len(every))], 'Second')
    runner += ('\nPublic Function Probe(ByVal folder As String) As String\n    Dim out As String\n    Setup\n'
               '    out = First()\n    Reopen folder\n    out = out & "@@" & Second()\n    Finish\n'
               '    Probe = out\nEnd Function\n')
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        result = excel.run_vba(probe + runner, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
    answers = str(result.value).split('@@')
    count = len(every)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': probe, 'states': {
        family: FAMILIES[family][1] for family in FAMILIES},
        'cases': [[*case, answers[number], answers[count + len(MIXED) + number]] for number, case in enumerate(every)],
        'mixed': [[*case, answers[count + number]] for number, case in enumerate(MIXED)]}
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    print(len(every), 'cases', len(MIXED), 'mixed')


if __name__ == '__main__':
    main()
