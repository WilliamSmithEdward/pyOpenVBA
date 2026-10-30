"""Measure VBA's conversions under the ja-JP Regional format through oleaut32, which VBA calls for them.

VBA's Format, CStr of a date, CDate, IsDate and the number conversions take
the user's Regional format through oleaut32. Called with an LCID, oleaut32
answers as VBA does under that locale, so ja-JP (LCID 1041) can be measured
without changing any setting. That oleaut32 is the oracle was checked against
live VBA under each format:

- en-US: VarDateFromStr agrees with tests/fixtures/date_text.json on 180 of
  182 strings (the other two are VBA's own fallback to a number), and
  VarFormatFromTokens with live Format on 32 date codes.
- ja-JP: with the Regional format switched to Japanese, live VBA agreed with
  oleaut32 on every date it read and every code it formatted
  (tests/fixtures/vba_locale_live.json, scripts/measure_vba_locale_live.py).

    python scripts/measure_vba_locale.py

writes tests/fixtures/vba_locale.json, which tests/test_vba_locale.py replays.
Windows only.
"""
from __future__ import annotations

import ctypes
import datetime as dt
import itertools
import json
from ctypes import wintypes
from pathlib import Path
from typing import Any

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_locale.json'
LCID = 1041
VT_R8, VT_DATE = 5, 7

_windows: Any = ctypes
ole = _windows.WinDLL('oleaut32')
ole.SysFreeString.argtypes = [ctypes.c_void_p]


class _Variant(ctypes.Structure):
    _fields_ = [('vt', ctypes.c_ushort), ('reserved1', ctypes.c_ushort), ('reserved2', ctypes.c_ushort),
                ('reserved3', ctypes.c_ushort), ('value', ctypes.c_double), ('padding', ctypes.c_void_p)]


def _bstr(out: ctypes.c_void_p) -> str:
    text = ctypes.wstring_at(out.value) if out.value else ''
    ole.SysFreeString(out)
    return text


def formatted(value: float, vt: int, code: str) -> str | None:
    """Format(value, code), or None where oleaut32 refuses the code."""
    tokens = (ctypes.c_ubyte * 4096)()
    used = ctypes.c_int()
    if ole.VarTokenizeFormatString(ctypes.c_wchar_p(code), tokens, 4096, 1, 1, wintypes.DWORD(LCID),
                                   ctypes.byref(used)):
        return None
    variant = _Variant(vt=vt, value=value)
    out = ctypes.c_void_p()
    if ole.VarFormatFromTokens(ctypes.byref(variant), ctypes.c_wchar_p(code), tokens, 0, ctypes.byref(out),
                               wintypes.DWORD(LCID)):
        return None
    return _bstr(out)


def date_string(serial: float) -> str:
    out = ctypes.c_void_p()
    assert not ole.VarBstrFromDate(ctypes.c_double(serial), wintypes.DWORD(LCID), wintypes.DWORD(0),
                                   ctypes.byref(out))
    return _bstr(out)


def read_date(text: str) -> float | None:
    out = ctypes.c_double()
    hr = ole.VarDateFromStr(ctypes.c_wchar_p(text), wintypes.DWORD(LCID), wintypes.DWORD(0), ctypes.byref(out))
    return None if hr else out.value


def read_number(text: str) -> float | None:
    out = ctypes.c_double()
    hr = ole.VarR8FromStr(ctypes.c_wchar_p(text), wintypes.DWORD(LCID), wintypes.DWORD(0), ctypes.byref(out))
    return None if hr else out.value


def serial(*parts: int) -> float:
    delta = dt.datetime(*parts) - dt.datetime(1899, 12, 30)
    days, fraction = delta.days, delta.seconds / 86400
    # Before 30 December 1899 the day counts back and the time still forward: 29 December noon is -1.5.
    return days + fraction if days >= 0 else days - fraction


MOMENTS = [(100, 1, 1), (1867, 12, 31, 9, 5, 7), (1868, 1, 1), (1868, 12, 31, 23, 0, 0), (1869, 1, 1),
           (1912, 7, 29), (1912, 7, 30, 0, 0, 1), (1926, 12, 24), (1926, 12, 25), (1989, 1, 7), (1989, 1, 8),
           (1999, 3, 4, 12, 0, 0), (2019, 4, 30, 23, 59, 59), (2019, 5, 1), (2020, 1, 5, 9, 5, 7),
           (2024, 2, 29, 0, 30, 0), (2024, 10, 28, 13, 4, 5), (2024, 12, 31, 11, 59, 59), (9999, 12, 31, 18, 0, 0)]
DATE_CODES = ['g', 'gg', 'ggg', 'e', 'ee', 'ggge', 'gge', 'ggee', 'ggge年m月d日', '[$-411]ggge', 'gge.m.d',
              'G', 'GGG', 'E', 'EE', 'eee', 'gggg', 'aaa', 'aaaa', 'aa', 'a', 'ooo', 'oooo', 'oo', 'o', 'ddd',
              'dddd', 'mmm', 'mmmm', 'c', 'ddddd', 'dddddd', 'ttttt', 'General Date', 'Long Date',
              'Medium Date', 'Short Date', 'Long Time', 'Medium Time', 'Short Time', 'h:nn AM/PM', 'h:nn am/pm',
              'h:nn A/P', 'h:nn AMPM', 'AMPM h:nn', 'hh AMPM nn', 'h AMPM AMPM', 'yyyy/mm/dd hh:nn:ss',
              'yy-mm-dd', 'yyyy"年"m"月"d"日"', 'dddd, mmmm d', 'w ww q y', 'c AMPM', 'ddddd ttttt',
              'General Date;x', 'aaaa"です"', 'ggg"年"']
NUMBERS = [0.0, 0.5, 1.0, 1234.5, -1234.5, 1234567.891, -0.4, 0.125, 1e20]
NUMBER_CODES = ['Currency', 'Fixed', 'Standard', 'Percent', 'Scientific', 'General Number', 'Yes/No',
                'True/False', 'On/Off', '$#,##0', '¥#,##0', '\\##0', '#,##0.00 "円"', '0.00%']


def _date_texts() -> list[str]:
    """Text to read as a date: the Japanese forms and the usual ones, joined every way that matters."""
    eras = ['令和', '令', 'R', 'r', '平成', 'H', '昭和', 'S', '大正', 'T', '明治', 'M', 'm', 'X']
    marked_parts = {'年': ['2019', '24', '5', '元'], '月': ['10', '1', '13'], '日': ['28', '5', '31', '32']}
    marked: list[str] = []
    for count in (1, 2, 3):
        for units in itertools.permutations('年月日', count):
            for values in itertools.product(*[marked_parts[unit] for unit in units]):
                for joiner in ('', ' '):
                    marked.append(joiner.join(f'{value}{unit}' for value, unit in zip(values, units, strict=True)))
    marked = [text for text in marked if '元' not in text or text.startswith('元年')]
    plain = [f'{a}{sep}{b}{sep}{c}' for a, b, c in [('3', '4', '5'), ('13', '4', '5'), ('3', '13', '5'),
                                                       ('31', '12', '5'), ('2019', '3', '4'), ('3', '4', '2019'),
                                                       ('28', '10', '2019'), ('5', '3', '31'), ('99', '12', '31'),
                                                       ('12', '31', '99'), ('0', '1', '2')]
             for sep in ('/', '-', ' ', ',', '.', '/ ')]
    plain += [f'{a}{sep}{b}' for a, b in [('3', '4'), ('13', '4'), ('3', '13'), ('2019', '3'), ('3', '2019'),
                                          ('13', '13'), ('99', '3'), ('3', '99'), ('31', '4')]
              for sep in ('/', '-', ' ')]
    named = ['Oct 28', '28 Oct', 'Oct 28 2019', '28 Oct 2019', 'Oct 2019', '2019 Oct', 'Mar 4 5', '4 Mar 5',
             'Mar 13 5', '13 Mar 5', 'Mar 4 13', '40 Mar 4', 'Mar 40', 'March 4, 2019', '4-Mar-19', 'Oct 28日',
             '10月 Oct']
    dates = marked + plain + named
    with_eras: list[str] = []
    for era in eras:
        for body in ['8/10/28', '8年10月28日', '元年5月1日', '元/5/1', '8年', '8年10月', '8/10', '8', '8.10.28',
                     '32年1月1日', '1年1月1日', '45/7/29', ' 8/10/28', '8年10月28日 13時', '8-10-28',
                     '8 10 28', '8年28日10月', '10月28日', '8/10/28 13:04']:
            with_eras.append(era + body)
    clocks = ['13:04', '1:04:05', '9:05', '13:04:05', '25:00', '13:60']
    for count in (1, 2, 3):
        for units in itertools.permutations('時分秒', count):
            for values in itertools.product(*[{'時': ['13', '9', '25', '0', '12'], '分': ['4', '60'],
                                                '秒': ['5', '61']}[unit] for unit in units]):
                for joiner in ('', ' '):
                    clocks.append(joiner.join(f'{value}{unit}' for value, unit in zip(values, units, strict=True)))
    marks = ['午前', '午後', 'AM', 'PM', 'a', 'p']
    times = list(clocks)
    for mark in marks:
        for clock in ['13:04', '1:04', '9時', '13時', '12時', '1時4分', '0時', '9', '1:04:05', '4分', '9時5秒']:
            times += [f'{mark}{clock}', f'{mark} {clock}', f'{clock}{mark}', f'{clock} {mark}']
        times += [mark, f'{mark}2019/10/28', f'2019/10/28{mark}', f'2019/10/28 {mark}']
    joined: list[str] = []
    for date, time in itertools.product(['2019/10/28', '2019年10月28日', '10月28日', '令和1年10月28日', 'Oct 28 2019',
                                         '28/10/2019'],
                                        ['13:04', '13時', '13時4分5秒', '午後1:04', '1:04 PM', '午後1時', '4分']):
        for joiner in ('', ' ', '/'):
            joined += [f'{date}{joiner}{time}', f'{time}{joiner}{date}']
    other = ['２０１９年１０月２８日', '２０１９/１０/２８', '１３：０４', '１２．５', '１，２００', 'Ｒ８/１０/２８',
             '2019年10月28日(月)', '2019/10/28 (月)', '2019/10/28月', '月曜日', '2019年10月28日 月曜日', '時', '年',
             '', ' ', '2019年10月28日 ', ' 2019年10月28日', '2019 年 10 月 28 日', '2019年　10月', '十月',
             '2019年10月28日午後', '午後2019年10月28日', '2019年10月28日13', '13 時', '13時 4分']
    return list(dict.fromkeys(dates + with_eras + times + joined + other))


def _number_texts() -> list[str]:
    bodies = ['12', '1,200', '1200.5', '.5', '1e3', '12-', '-12', '+12', '(12)', '&H10', '&O17', '1,2,3', ',5',
              '1 200', '12%', '']
    signs = ['', '\\', '¥', '￥', '$', '＄', '\\\\', '¥¥']
    texts: list[str] = []
    for body in bodies:
        for sign in signs:
            texts += [sign + body, body + sign, f'({sign}{body})', f'{sign}({body})', f'-{sign}{body}',
                      f'{sign}-{body}', f' {sign} {body} ']
    wide = str.maketrans({chr(code): chr(code + 0xFEE0) for code in range(0x21, 0x7F)})
    texts += [text.translate(wide) for text in list(texts)[:400]]
    texts += ['１２', '１２．５', '－１２', '（１２）', '１，２００', '＆Ｈ１０', '１ｅ３', '１２％', '　12　', '12　',
              '¥　12', '1２', '１2.５']
    return list(dict.fromkeys(texts))


def main() -> None:
    year = dt.date.today().year
    date_texts = _date_texts()
    assert not any(str(year) in text for text in date_texts), 'a text spells out the year it is measured in'
    record = {
        'lcid': LCID,
        'measured_at': dt.datetime.now(dt.timezone.utc).isoformat(),
        'measured_year': year,
        'format_dates': [[list(moment), code, formatted(serial(*moment), VT_DATE, code)]
                         for moment in MOMENTS for code in DATE_CODES],
        'format_numbers': [[number, code, formatted(number, VT_R8, code)] for number in NUMBERS
                           for code in NUMBER_CODES],
        'date_strings': [[list(moment), date_string(serial(*moment))] for moment in MOMENTS],
        'read_dates': [[text, read_date(text)] for text in date_texts],
        'read_numbers': [[text, read_number(text)] for text in _number_texts()],
    }
    OUT.write_text(json.dumps(record, indent=0, ensure_ascii=False) + '\n', encoding='utf-8')
    print(OUT, {key: len(value) for key, value in record.items() if isinstance(value, list)})


if __name__ == '__main__':
    main()
