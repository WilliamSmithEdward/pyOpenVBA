"""What Excel's cells show and read under the Windows Regional format in force (issue #252, the Excel half).

Run once under en-US and once with the Regional format switched to Japanese
(Japan), by the machine's owner (Settings > Time & language > Language &
region > Regional format, or Set-Culture):

    python scripts/measure_excel_locale_live.py

Each run records its answers in tests/fixtures/excel_locale_live.json under
the Regional format read from the registry, keeping the other runs. The
cases: every built-in number format and a few more, each showing a date, a
time, numbers and text; text typed into a General cell; NumberFormatLocal
set and read back; worksheet functions that read or write dates, numbers
and currency; Application.International. The machinery is
scripts/measure_vba_locale_live.py's.
"""
from __future__ import annotations

from pathlib import Path

from measure_vba_locale_live import measure

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_locale_live.json'

#: The format codes of Excel's built-in formats 1-49, then codes that name a locale, a currency or an era.
CODES = ['General', '0', '0.00', '#,##0', '#,##0.00', '#,##0_);(#,##0)', '#,##0_);[Red](#,##0)',
         '#,##0.00_);(#,##0.00)', '#,##0.00_);[Red](#,##0.00)', '$#,##0_);($#,##0)', '$#,##0_);[Red]($#,##0)',
         '$#,##0.00_);($#,##0.00)', '$#,##0.00_);[Red]($#,##0.00)', '0%', '0.00%', '0.00E+00', '# ?/?',
         '# ??/??', 'm/d/yyyy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss',
         'm/d/yyyy h:mm', '#,##0 ;(#,##0)', '#,##0 ;[Red](#,##0)', '#,##0.00;(#,##0.00)',
         '#,##0.00;[Red](#,##0.00)', '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
         '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)', '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
         '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 'mm:ss', '[h]:mm:ss', 'mm:ss.0', '##0.0E+0', '@',
         'yyyy/m/d', 'yyyy/mm/dd', 'mmmm d, yyyy', 'dddd', 'ddd', 'mmm', 'mmmm', 'AM/PM h', 'h AM/PM', 'A/P h',
         '[$-F800]dddd, mmmm dd, yyyy', '[$-F400]h:mm:ss AM/PM', '[$-x-sysdate]dddd, mmmm dd, yyyy',
         '[$-x-systime]h:mm:ss AM/PM', '[$-411]ggge"年"m"月"d"日"', '[$-ja-JP]aaaa', '[$-409]mmmm', 'ggge', 'aaa',
         'e', '[$¥-411]#,##0', '[$$-409]#,##0', '[$€-x-euro2] #,##0.00', '\\#,##0', '"$"#,##0', '#,##0"円"',
         '[DBNum1]0', '[DBNum2]0', '[DBNum3]0', 'yyyy"年"m"月"d"日"', 'm"月"d"日"', '#,##0;-#,##0', '¥#,##0',
         '#,##0.00 $']
VALUES = ['DateSerial(2026, 1, 5) + TimeSerial(9, 4, 5)', 'TimeSerial(13, 4, 5)', '1234.5', '-1234.5', '0',
          '"abc"']
TYPED = ['2026/1/5', '2026/01/05', '1/5/2026', '1/5', '5/2026', '2026/1', '2026-1-5', '1-5', '5-Jan', 'Jan-26',
         'Jan 5, 2026', '5 Jan 2026', '2026/1/5 13:04', '13:04', '13:04:05', '1:04 PM', '1:04 AM', '1:04 午後',
         '午後1:04', '1:04:05.5', '2026/13/5', '13/5/2026', '31/12', '26/1/5', '0026/1/5', '$1,200', '¥1,200',
         '\\1,200', '￥1,200', '1,200', '1.200', '12%', '(12)', '-12', '1.5E3', '1 1/2', '3/4', '1,200円',
         '２０２６/１/５', '１２', '1:04', '2026年1月5日', '1月5日', '13時4分']
LOCALS = ['G/標準', 'General', 'yyyy/mm/dd', 'yyyy"年"m"月"d"日"', 'm/d/yyyy', '¥#,##0', '$#,##0', '\\#,##0',
          '#,##0', 'h:mm AM/PM', 'h:mm 午後', '[$-411]ggge"年"', 'ge.m.d', 'aaaa', 'dddd', '0.00', '0,00']
FORMULAS = ['=DOLLAR(1234.5)', '=DOLLAR(-1234.5,0)', '=DOLLAR(1234.567,2)', '=FIXED(1234.5)', '=YEN(1234.5)',
            '=TEXT(46027,"yyyy/mm/dd")', '=TEXT(46027,"m/d/yyyy")', '=TEXT(46027,"ggge")',
            '=TEXT(46027,"[$-411]ggge")', '=TEXT(46027,"mmm")', '=TEXT(46027,"dddd")', '=TEXT(46027,"aaaa")',
            '=TEXT(0.5445,"h:mm AM/PM")', '=TEXT(1234.5,"$#,##0")', '=TEXT(1234.5,"¥#,##0")',
            '=TEXT(1234.5,"Currency")', '=DATEVALUE("2026/1/5")', '=DATEVALUE("1/5/2026")', '=DATEVALUE("1/5")',
            '=DATEVALUE("5-Jan-2026")', '=DATEVALUE("26/1/5")', '=DATEVALUE("13/5/2026")',
            '=DATEVALUE("2026年1月5日")', '=TIMEVALUE("1:04 PM")', '=TIMEVALUE("13:04")', '=TIMEVALUE("午後1:04")',
            '=VALUE("$1,200")', '=VALUE("¥1,200")', '=VALUE("\\1,200")', '=VALUE("1/5/2026")', '=VALUE("2026/1/5")',
            '=--"2026/1/5"', '=--"1/5/2026"', '="1/5"+0', '="$5"+0', '="¥5"+0', '=YEAR("1/5/2026")',
            '=NUMBERVALUE("1,200.5")', '=DATE(2026,1,5)', '=NOW()>0', '=CELL("format",B1)']

CASES: list[tuple[str, str, object]] = []
for code in CODES:
    for value in VALUES:
        CASES.append((f'{code} on {value}', 'shown', [code, value]))
for text in TYPED:
    CASES.append((f'Range.Value = {text!r}', 'cell', text))
for code in LOCALS:
    CASES.append((f'NumberFormatLocal = {code!r}', 'local', code))
for formula in FORMULAS:
    CASES.append((formula, 'text', formula))
for index in range(1, 46):
    CASES.append((f'International({index})', 'intl', index))


if __name__ == '__main__':
    measure(CASES, OUT)
