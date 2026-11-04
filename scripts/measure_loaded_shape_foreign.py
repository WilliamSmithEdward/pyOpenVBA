"""What Excel reads of text markup it does not write itself, as other programs and older files have it.

    python scripts/measure_loaded_shape_foreign.py

Takes tests/fixtures/loaded_shape_text.xlsx, puts each txBody below in a shape of its own in the drawing (tests/fixtures/loaded_shape_foreign.xlsx), opens that in Excel
and reads each shape.

Writes tests/fixtures/loaded_shape_foreign.json. tests/test_excel_loaded_shape_text.py replays it.
"""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
sys.path.insert(0, str(ROOT / 'scripts'))
from measure_loaded_shape_text import FONT, FONT2, FRAME, FRAME2, PARAGRAPH  # noqa: E402
from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.exceptions import VBAUnsupportedError  # noqa: E402

BASE = ROOT / 'tests/fixtures/loaded_shape_text.xlsx'
BOOK = ROOT / 'tests/fixtures/loaded_shape_foreign.xlsx'
OUT = ROOT / 'tests/fixtures/loaded_shape_foreign.json'


def _p(runs: str, properties: str = '') -> str:
    return f'<a:p>{properties}{runs}</a:p>'


def _r(text: str, properties: str = '') -> str:
    return f'<a:r>{properties}<a:t>{text}</a:t></a:r>'


_BODY = '<a:bodyPr/><a:lstStyle/>'
#: name, body. No attributes at all, then one kind of thing a file can hold at a time.
BODIES: dict[str, str] = {
    'Bare': _BODY + _p(_r('Hello')),
    'NoSize': _BODY + _p(_r('Hello', '<a:rPr lang="en-US"/>')),
    'French': _BODY + _p(_r('Bonjour', '<a:rPr lang="fr-FR" sz="1200" dirty="0"/>')),
    'BodyDefaults': '<a:bodyPr vert="vert" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="b"/><a:lstStyle/>' + _p(_r('Hello')),
    'Justified': '<a:bodyPr anchor="just"/><a:lstStyle/>' + _p(_r('Hello')),
    'Distributed': '<a:bodyPr anchor="dist"/><a:lstStyle/>' + _p(_r('Hello')),
    'BaseAnchor': '<a:bodyPr/><a:lstStyle/>' + _p(_r('Hello'), '<a:pPr fontAlgn="b"/>'),
    'Scaled': '<a:bodyPr><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr><a:lstStyle/>' + _p(_r('Hello', '<a:rPr sz="1600"/>')),
    'Break': _BODY + _p(_r('One') + '<a:br><a:rPr lang="en-US" sz="1400" b="1"/></a:br>' + _r('Two')),
    'BareBreak': _BODY + _p(_r('One') + '<a:br/>' + _r('Two')),
    'Field': _BODY + _p('<a:fld id="{A1B2C3D4-0000-4000-8000-000000000000}" type="TxLink"><a:rPr lang="en-US" sz="1100"/><a:pPr/><a:t>Linked</a:t></a:fld>'),
    'EmptyPara': _BODY + _p('') + _p(_r('Two')) + '<a:p/>',
    'EndOnly': _BODY + _p('<a:endParaRPr lang="en-US" sz="2000" b="1"/>'),
    'EndAfter': _BODY + _p(_r('Hello', '<a:rPr lang="en-US" sz="1100"/>') + '<a:endParaRPr lang="en-US" sz="1100"/>'),
    'EndDiffers': _BODY + _p(_r('Hello', '<a:rPr lang="en-US" sz="1100"/>') + '<a:endParaRPr lang="en-US" sz="2000" b="1"/>'),
    'Sentences': _BODY + _p(_r('One two. Three four! Five? Six; seven: eight, nine.')) + _p(_r('Ten.  Eleven e.g. twelve.')),
    'Wordy': _BODY + _p(_r(' Lead  double-word don\'t 3.14 a,b (c) "q" end ')),
    'Sng': _BODY + _p(_r('Hello', '<a:rPr u="sng"/>')),
    'Heavy': _BODY + _p(_r('Hello', '<a:rPr u="heavy"/>')),
    'Wavy': _BODY + _p(_r('Hello', '<a:rPr u="wavy"/>')),
    'DblStrike': _BODY + _p(_r('Hello', '<a:rPr strike="dblStrike"/>')),
    'NoStrike': _BODY + _p(_r('Hello', '<a:rPr strike="noStrike" u="none" b="0" i="0"/>')),
    'SmallCaps': _BODY + _p(_r('Hello', '<a:rPr cap="small"/>')),
    'Baseline': _BODY + _p(_r('Hello', '<a:rPr baseline="10000"/>')),
    'Normalized': _BODY + _p(_r('Hello', '<a:rPr normalizeH="1"/>')),
    'Major': _BODY + _p(_r('Hello', '<a:rPr><a:latin typeface="+mj-lt"/></a:rPr>')),
    'Minor': _BODY + _p(_r('Hello', '<a:rPr><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:rPr>')),
    'Named': _BODY + _p(_r('Hello', '<a:rPr><a:latin typeface="Times New Roman"/><a:ea typeface="MS Gothic"/><a:cs typeface="Arial"/><a:sym typeface="Wingdings"/></a:rPr>')),
    'Tinted': _BODY + _p(_r('Hello', '<a:rPr><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:rPr>')),
    'Lit': _BODY + _p(_r('Hello', '<a:rPr><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill></a:rPr>')),
    'System': _BODY + _p(_r('Hello', '<a:rPr><a:solidFill><a:sysClr val="windowText" lastClr="000000"/></a:solidFill></a:rPr>')),
    'Unpainted': _BODY + _p(_r('Hello', '<a:rPr><a:noFill/></a:rPr>')),
    'Graded': _BODY + _p(_r('Hello', '<a:rPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></a:rPr>')),
    'Patterned': _BODY + _p(_r('Hello', '<a:rPr><a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></a:rPr>')),
    'Glowing': _BODY + _p(_r('Hello', '<a:rPr><a:effectLst><a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow></a:effectLst></a:rPr>')),
    'Shadowed': _BODY + _p(_r('Hello', '<a:rPr><a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="000000"><a:alpha val="43137"/></a:srgbClr></a:outerShdw></a:effectLst></a:rPr>')),
    'Marked': _BODY + _p(_r('Hello', '<a:rPr><a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr>')),
    'UnderInk': _BODY + _p(_r('Hello', '<a:rPr u="sng"><a:uFill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:uFill></a:rPr>')),
    'Linked': _BODY + _p(_r('Hello', '<a:rPr><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id=""/></a:rPr>')),
    'Listed': '<a:bodyPr/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="2400" b="1"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>' + _p(_r('Hello')),
    'Spaced': _BODY + _p(_r('Hello'), '<a:pPr marL="100000" marR="50000" indent="-25000" algn="just" fontAlgn="ctr" latinLnBrk="1" eaLnBrk="0" hangingPunct="0" rtl="1" lvl="2"><a:lnSpc><a:spcPts val="1800"/></a:lnSpc><a:spcBef><a:spcPct val="50000"/></a:spcBef><a:spcAft><a:spcPts val="300"/></a:spcAft></a:pPr>'),
    'Bulleted': _BODY + _p(_r('Hello'), '<a:pPr><a:buChar char="&#8226;"/></a:pPr>') + _p(_r('World'), '<a:pPr><a:buNone/></a:pPr>'),
    'Warped': '<a:bodyPr><a:prstTxWarp prst="textWave1"><a:avLst/></a:prstTxWarp></a:bodyPr><a:lstStyle/>' + _p(_r('Hello')),
    'Turned': '<a:bodyPr rot="1800000"/><a:lstStyle/>' + _p(_r('Hello')),
    'Thaid': _BODY + _p(_r('Hello'), '<a:pPr algn="thaiDist"/>'),
    'Marks': _BODY + _p(_r('Hello &amp; &lt;world&gt;')),
    'FontTop': _BODY + _p(_r('Hello'), '<a:pPr fontAlgn="t"/>'),
    'FontBase': _BODY + _p(_r('Hello'), '<a:pPr fontAlgn="base"/>'),
    'FontAuto': _BODY + _p(_r('Hello'), '<a:pPr fontAlgn="auto"/>'),
    'Bulleted2': _BODY + _p(_r('Hello'), '<a:pPr><a:buNone/></a:pPr>') + _p(_r('World'), '<a:pPr><a:buChar char="&#8226;"/></a:pPr>'),
    'Numbered': _BODY + _p(_r('Hello'), '<a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr>'),
    'Dist': _BODY + _p(_r('Hello'), '<a:pPr algn="dist"/>'),
    'JustLow': _BODY + _p(_r('Hello'), '<a:pPr algn="justLow"/>'),
    'Hyphens': _BODY + _p(_r('a-b c_d e.f g/h i\\j k@l m#n 100% $5 x+y=z')),
    'Dots': _BODY + _p(_r('Wait... what?! Yes. No')),
    'Titles': _BODY + _p(_r('Mr. Smith went. He said &quot;Hi.&quot; Then left.')),
    'Spaces': _BODY + _p(_r('  two   three  ')),
    'Tabbed': _BODY + _p(_r('one\ttwo')),
    'Numbers': _BODY + _p(_r('1,000.50 and 2.5. Next')),
    'Mixed2': _BODY + _p(_r('A b. ') + _r('C d', '<a:rPr b="1"/>') + _r('. E')),
    **{f'U{kind}': _BODY + _p(_r('Hello', f'<a:rPr u="{kind}"/>')) for kind in (
        'dotted', 'dottedHeavy', 'dash', 'dashHeavy', 'dashLong', 'dashLongHeavy', 'dotDash', 'dotDashHeavy', 'dotDotDash', 'dotDotDashHeavy', 'wavyHeavy', 'wavyDbl')},
}
READS: list[str] = []
for name in BODIES:
    head = f'ws.Shapes("{name}")'
    READS += [f'{head}.TextFrame.{member}' for member in FRAME]
    READS += [f'{head}.TextFrame2.{member}' for member in FRAME2 if not member.endswith('Lines.Count')]
    READS += [f'{head}.TextFrame2.TextRange.ParagraphFormat.{member}' for member in PARAGRAPH]
    READS += [f'{head}.TextFrame.Characters.Font.{member}' for member in FONT]
    READS += [f'{head}.TextFrame2.TextRange.Font.{member}' for member in FONT2]
for name in ('Sentences', 'Wordy', 'Hyphens', 'Dots', 'Titles', 'Spaces', 'Tabbed', 'Numbers', 'Mixed2', 'Break', 'EmptyPara', 'Bare'):
    head = f'ws.Shapes("{name}").TextFrame2.TextRange'
    READS += [f'{head}.Words({place}).Text' for place in range(1, 28)]
    READS += [f'{head}.Sentences({place}).Text' for place in range(1, 9)]
    READS += [f'{head}.Sentences({place}).Start' for place in range(1, 9)]


def built() -> None:
    """The base book's drawing with a new shape for each body."""
    with zipfile.ZipFile(BASE) as package:
        parts = {info.filename: package.read(info.filename) for info in package.infolist()}
    xml = parts['xl/drawings/drawing1.xml'].decode('utf-8')
    anchors = ''
    for number, (name, body) in enumerate(BODIES.items()):
        row = number * 3
        anchors += (f'<xdr:twoCellAnchor><xdr:from><xdr:col>20</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>23</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row + 3}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                    f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{100 + number}" name="{name}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
                    '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>'
                    f'<xdr:txBody>{body}</xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>')
    parts['xl/drawings/drawing1.xml'] = xml.replace('</xdr:wsDr>', anchors + '</xdr:wsDr>').encode('utf-8')
    with zipfile.ZipFile(BOOK, 'w', zipfile.ZIP_DEFLATED) as package:
        for name, data in parts.items():
            package.writestr(name, data)


def reader(reads: list[str]) -> str:
    lines = ['''Private Function Shown(v As Variant) As String
    If IsObject(v) Then
        Shown = "object"
    ElseIf IsArray(v) Or IsNull(v) Then
        Shown = TypeName(v)
    Else
        Shown = CStr(v)
    End If
End Function
''']
    names: list[str] = []
    for start in range(0, len(reads), 60):
        name = f'Part{start}'
        names.append(name)
        lines += [f'Private Function {name}(ws As Object) As String', '    Dim v As Variant, s As String', '    On Error Resume Next']
        for read in reads[start:start + 60]:
            lines += ['    s = s & "@@"', '    Err.Clear', f'    v = {read}', '    If Err.Number Then', '        s = s & "E" & Err.Number: Err.Clear', '    Else', '        s = s & Shown(v)', '    End If']
        lines += [f'    {name} = s', 'End Function']
    lines += ['Public Function Opened(ByVal path As String) As String', '    Dim wb As Workbook, s As String', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Open(path)']
    lines += [f'    s = s & {name}(wb.Worksheets(1))' for name in names]
    lines += ['    wb.Close False', '    Application.DisplayAlerts = True', '    Opened = s', 'End Function']
    return '\n'.join(lines) + '\n'


def one_read(number: int, read: str) -> str:
    return '\n'.join([f'Public Function R{number}() As String', '    Dim ws As Object, v As Variant', '    Set ws = ActiveWorkbook.Worksheets(1)', '    On Error Resume Next', f'    v = {read}', '    If Err.Number Then', f'        R{number} = "E" & Err.Number',
                      '    ElseIf IsObject(v) Then', f'        R{number} = "object"', '    ElseIf IsArray(v) Or IsNull(v) Then', f'        R{number} = TypeName(v)', '    Else', f'        R{number} = CStr(v)', '    End If', 'End Function', ''])


def in_model() -> list[str]:
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module(''.join(one_read(number, read) for number, read in enumerate(READS)), name='Reader')
    answers: list[str] = []
    for number in range(len(READS)):
        try:
            answers.append(str(app.run(f'Reader.R{number}')))
        except VBAUnsupportedError as refused:
            answers.append(f'UNSUPPORTED {str(refused)[:80]}')
    return answers


def main() -> None:
    built()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        copy = Path(folder) / 'book.xlsx'
        shutil.copyfile(BOOK, copy)
        excel.new_document()
        truth: list[str] = []
        # A module past about 15,000 lines times out on the way in.
        for start in range(0, len(READS), 1200):
            read = excel.run_vba(reader(READS[start:start + 1200]), 'Opened', (str(copy),), timeout=280.0, line_numbers=False)
            assert read.ok and not read.dialogs, read
            truth += str(read.value).split('@@')[1:]
    assert len(truth) == len(READS), (len(truth), len(READS))
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'bodies': BODIES, 'reads': READS, 'excel': truth}, indent=0) + '\n', encoding='utf-8')
    print(len(READS), 'reads measured')


if __name__ == '__main__':
    main()
