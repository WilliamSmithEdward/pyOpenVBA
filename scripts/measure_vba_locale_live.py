"""What VBA's conversions and Excel's cells do under the Windows Regional format in force (issue #252).

Run once under en-US and once with the Regional format switched to Japanese
(Japan), by the machine's owner (Settings > Time & language > Language &
region > Regional format, or Set-Culture):

    python scripts/measure_vba_locale_live.py

Each run records its answers in tests/fixtures/vba_locale_live.json under the
Regional format read from the registry, keeping the other runs, so a run that
missed the switch labels itself. tests/test_vba_locale.py replays the VBA
cases (kind "expr"); the cell, format, formula, International and sort cases
are the Excel host's. The VBA source is ASCII only, since the project's code
page is 1252 here: Japanese text is built with ChrW, and every answer comes
back as hex UTF-16 units so nothing is lost on the way.
"""
from __future__ import annotations

import json
import winreg
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_locale_live.json'
D = 'DateSerial(2026, 10, 28)'
J = 'DateSerial(2026, 1, 5)'
T = 'TimeSerial(13, 4, 5)'


def vb(text: str) -> str:
    """A VBA expression for text, with every character outside printable ASCII as ChrW."""
    parts: list[str] = []
    run = ''
    for character in text:
        if ' ' <= character <= '~':
            run += '""' if character == '"' else character
        else:
            if run:
                parts.append(f'"{run}"')
                run = ''
            parts.append(f'ChrW(&H{ord(character):04X})')
    if run or not parts:
        parts.append(f'"{run}"')
    return ' & '.join(parts)


DATE_TEXTS = ['令和8年10月28日', '令和元年5月1日', '平成31年4月30日', 'R8.10.28', 'H31.4.30', 'R8/10/28',
              '2026年1月5日', '2026/01/05', '2026/1/5', '1/5/2026', '1/5', '5-Jan', 'Jan 5, 2026', '13時4分',
              '13:04', '午後1:04', '1:04 PM']
NUMBER_TEXTS = ['\\1,200', '￥1,200', '¥1,200', '$1,200', '(12)', '１２', '1,200', '1.200,5', '１２％', '12%',
                '-12', '12-', '1e3']
CELL_TEXTS = DATE_TEXTS + NUMBER_TEXTS + ['（１２）', 'ＴＲＵＥ', 'TRUE', '＄１２', '１２：３０', '＇AB', "'AB",
                                          '’AB', '1-2', '3/4', '1 1/2', '０．５']
FORMATS = ['ggge年m月d日', 'gge', 'gg', 'g', 'e', 'ee', 'yyyy/mm/dd', 'aaa', 'aaaa', 'ddd', 'dddd', 'General Date',
           'Long Date', 'Medium Date', 'Short Date', 'Long Time', 'Medium Time', 'Short Time', 'c', 'ddddd',
           'dddddd', 'ttttt', 'AM/PM', 'h:nn AMPM', 'mmm', 'mmmm', 'oooo', 'w', 'ww', 'q', 'y', 'm/d/yy']
SORT_TEXTS = ['あ', 'ア', 'ｱ', 'い', 'か', 'が', 'カ', '亜', '阿', '一', 'Ａ', 'a', 'A', 'ａ', 'b', '１', '1', '2',
              'ー', 'ぁ', 'ゃ', 'や', 'ﾔ', '漢', '字']

#: Each case: (label, kind, payload). Kinds: "expr" evaluates a VBA expression; "cell" types text into a General
#: cell and reads it back; "intl" reads Application.International(n); "sort" sorts strings with Range.Sort;
#: "text" puts a formula in a cell and reads its value and text.
CASES: list[tuple[str, str, object]] = []
for expression in [f'CStr({J})', f'CStr({D} + {T})', f'CStr({T})', 'CStr(1234.5)', 'CStr(CCur(1234.5))',
                   'CStr(True)', f'CStr(CVar({J}))', f'FormatDateTime({J}, vbLongDate)',
                   f'FormatDateTime({J}, vbShortDate)', f'FormatDateTime({T}, vbLongTime)',
                   'FormatCurrency(1234.5)', 'FormatCurrency(-1234.5)', 'FormatNumber(1234.5)',
                   'FormatPercent(0.125)', 'MonthName(10)', 'MonthName(10, True)', 'WeekdayName(3)',
                   'WeekdayName(3, True)', 'Format(1234.5, "Currency")', 'Format(1234.5, "Standard")',
                   'Format(-1234.5, "Currency")', 'Format(1234.5)', 'Format(0.5, "Percent")',
                   'Format(True, "Yes/No")', f'Format({J})', f'DatePart("ww", {J})', f'Weekday({J})',
                   f'Year({D})', 'CStr(#10/28/2026#)']:
    CASES.append((expression, 'expr', expression))
for code in FORMATS:
    CASES.append((f'Format(D, {code!r})', 'expr', f'Format({D} + {T}, {vb(code)})'))
for text in DATE_TEXTS:
    CASES.append((f'IsDate({text!r})', 'expr', f'IsDate({vb(text)})'))
    CASES.append((f'CDate({text!r})', 'expr', f'CDbl(CDate({vb(text)}))'))
    CASES.append((f'DateValue({text!r})', 'expr', f'CDbl(DateValue({vb(text)}))'))
for text in NUMBER_TEXTS:
    for function in ('IsNumeric', 'CDbl', 'CCur', 'CDec', 'Val', 'CLng'):
        CASES.append((f'{function}({text!r})', 'expr', f'{function}({vb(text)})'))
for first, second in [('あ', 'ア'), ('Ａ', 'a'), ('ｱ', 'ア'), ('が', 'か'), ('１', '1')]:
    CASES.append((f'StrComp({first!r}, {second!r}, vbTextCompare)', 'expr',
                  f'StrComp({vb(first)}, {vb(second)}, vbTextCompare)'))
    CASES.append((f'{first!r} < {second!r}', 'expr', f'({vb(first)} < {vb(second)})'))
for text in CELL_TEXTS:
    CASES.append((f'Range.Value = {text!r}', 'cell', text))
for code in ['General', 'ggge"年"m"月"d"日"', '[$-ja-JP]ggge"年"m"月"d"日"', '[$-411]gge.m.d', 'm/d/yyyy',
             'yyyy/m/d', 'mmm d', '[$-F800]dddd, mmmm dd, yyyy', '[$-F400]h:mm:ss AM/PM', '#,##0', '$#,##0',
             '[$¥-411]#,##0', '\\#,##0']:
    CASES.append((f'NumberFormat {code!r} on 2026-10-28 13:04:05', 'fmt', code))
for formula in ['=TEXT(DATE(2026,10,28),"ggge年m月d日")', '=TEXT(DATE(2026,10,28),"[$-ja-JP]ggge年m月d日")',
                '=TEXT(1234.5,"$#,##0")', '=DATEVALUE("令和8年10月28日")', '=DATEVALUE("2026/1/5")',
                '=DATEVALUE("1/5/2026")', '=VALUE("１２")', '=VALUE("\\1,200")', '=VALUE("$1,200")',
                '=YEN(1234.5)', '=DOLLAR(1234.5)', '=ASC("ＡＢ")', '=JIS("AB")', '=PHONETIC(A1)', '=DATESTRING(46323)',
                '=NUMBERSTRING(123,1)', '=LEN("あ")', '=LENB("あ")', '=FIXED(1234.5)']:
    CASES.append((formula, 'text', formula))
for index in range(1, 46):
    CASES.append((f'International({index})', 'intl', index))
CASES.append(('Range.Sort of kana, kanji and widths', 'sort', SORT_TEXTS))
CASES.append(('Range.Sort of kana, kanji and widths, MatchCase', 'sort-case', SORT_TEXTS))


def _probe() -> str:
    lines = ['Private Function H(ByVal s As String) As String',
             '    Dim i As Long, r As String',
             '    For i = 1 To Len(s)',
             '        r = r & Right("000" & Hex(AscW(Mid(s, i, 1)) And &HFFFF&), 4)',
             '    Next',
             '    H = r',
             'End Function', '',
             'Private Function Shown(ByVal v As Variant) As String',
             '    If IsError(v) Then',
             '        Shown = "Error:" & H(CStr(v))',
             '    Else',
             '        Shown = TypeName(v) & ":" & H(CStr(v))',
             '    End If',
             'End Function', '',
             'Private Function Cell(ByVal c As Range) As String',
             '    Cell = Shown(c.Value) & "|" & Shown(c.Value2) & "|" & H(c.Text) & "|" & H(c.NumberFormat) & "|" & _',
             '        H(c.NumberFormatLocal) & "|" & H(c.Formula) & "|" & H(c.PrefixCharacter)',
             'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal ws As Worksheet) As String',
             '    Dim c As Range, i As Long, s As String, r As Range',
             '    Set c = ws.Cells(which + 1, 1)',
             '    c.ColumnWidth = 60',
             '    On Error GoTo Failed',
             '    Select Case which']
    for index, (_label, kind, payload) in enumerate(CASES):
        lines.append(f'    Case {index}')
        if kind == 'expr':
            lines.append(f'        Tried = Shown({payload})')
        elif kind == 'cell':
            lines.append(f'        c.Value = {vb(str(payload))}')
            lines.append('        Tried = Cell(c)')
        elif kind == 'fmt':
            lines.append(f'        c.Value = {D} + {T}')
            lines.append(f'        c.NumberFormat = {vb(str(payload))}')
            lines.append('        Tried = Cell(c)')
        elif kind == 'text':
            lines.append(f'        c.Formula = {vb(str(payload))}')
            lines.append('        Tried = Cell(c)')
        elif kind == 'intl':
            lines.append(f'        Tried = Shown(Application.International({payload}))')
        else:
            texts = list(payload)  # type: ignore[call-overload]
            for place, text in enumerate(texts):
                lines.append(f'        ws.Cells({place + 1}, 3).Value = {vb(text)}')
                lines.append(f'        ws.Cells({place + 1}, 4).Value = {place}')
            last = len(texts)
            match = 'True' if kind == 'sort-case' else 'False'
            lines.append(f'        Set r = ws.Range(ws.Cells(1, 3), ws.Cells({last}, 4))')
            lines.append(f'        r.Sort Key1:=ws.Cells(1, 3), Order1:=xlAscending, Header:=xlNo, MatchCase:={match}')
            lines.append(f'        For i = 1 To {last}')
            lines.append('            s = s & "," & ws.Cells(i, 4).Value')
            lines.append('        Next')
            lines.append('        r.ClearContents')
            lines.append('        Tried = Mid(s, 2)')
    lines += ['    End Select', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & H(Err.Description)', 'End Function', '',
              'Public Function Probe(ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long, wb As Workbook',
              '    Set wb = Workbooks.Add(xlWBATWorksheet)',
              '    For i = first To last',
              '        s = s & "@@@" & Tried(i, wb.Worksheets(1))',
              '    Next',
              '    wb.Close False',
              '    Probe = Mid(s, 4)',
              'End Function']
    return '\n'.join(lines) + '\n'


def _decoded(coded: str) -> str:
    return ''.join(chr(int(coded[place:place + 4], 16)) for place in range(0, len(coded), 4))


def _readable(answer: str, kind: str) -> object:
    """Hex fields turned back into text, so the JSON reads plainly."""
    if answer.startswith('Err '):
        number, _, description = answer[4:].partition(' ')
        return {'error': int(number), 'description': _decoded(description)}
    if kind in ('sort', 'sort-case'):
        return [int(place) for place in answer.split(',')]

    def shown(field: str) -> list[str]:
        type_name, _, value = field.partition(':')
        return [type_name, _decoded(value)]

    if kind in ('cell', 'fmt', 'text'):
        value, value2, text, number_format, local, formula, prefix = answer.split('|')
        return {'Value': shown(value), 'Value2': shown(value2), 'Text': _decoded(text),
                'NumberFormat': _decoded(number_format), 'NumberFormatLocal': _decoded(local),
                'Formula': _decoded(formula), 'PrefixCharacter': _decoded(prefix)}
    return shown(answer)


def _regional_format() -> str:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Control Panel\International') as key:
        return str(winreg.QueryValueEx(key, 'LocaleName')[0])


def main() -> None:
    locale = _regional_format()
    source = _probe()
    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 60):
            last = min(first + 59, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (first, last), timeout=300.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'regional_format': locale,
              'cases': [[label, kind, payload, _readable(answer, kind)]
                        for (label, kind, payload), answer in zip(CASES, answers, strict=True)]}
    runs = json.loads(OUT.read_text(encoding='utf-8')) if OUT.exists() else {}
    runs[locale] = record
    OUT.write_text(json.dumps(dict(sorted(runs.items())), indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    print(locale, len(record['cases']), 'cases ->', OUT)


if __name__ == '__main__':
    main()
