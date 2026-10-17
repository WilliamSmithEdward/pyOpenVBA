"""Native statistical populations, sparse unions and read-scoped recomputation."""
from __future__ import annotations

import json
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel import _conditional_stats as stats
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._conditional import NS, Rule
from pyopenvba.apps.excel._conditional_eval import conditional_number_format, display_style
from pyopenvba.apps.excel._model import Worksheet
from pyopenvba.exceptions import VBAUnsupportedError

NATIVE: list[dict[str, Any]] = []
for filename in ('statistics_display.json', 'statistics_keys.json', 'statistics_criteria.json', 'statistics_unicode.json'):
    record = json.loads((Path(__file__).parent / 'fixtures/conditional' / filename).read_text(encoding='utf-8'))
    NATIVE.extend(dict(case, name=filename + '-' + case['name']) for case in record['probes'])


@pytest.mark.parametrize('case', NATIVE, ids=lambda case: case['name'])
def test_native_statistical_display_controls(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']


def setup(values: str, *, kind: str, attributes: str = '', areas: list[Area] | None = None) -> tuple[ExcelApplication, Worksheet]:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('Sub Setup()\n' + values + '\nEnd Sub', name='SetupModule')
    app.run('Setup')
    sheet = book.sheets_[0]
    root = xml_tree(f'<cfRule xmlns="{NS[1:-1]}" type="{kind}" priority="1" {attributes}/>')
    patch = xml_tree(f'<dxf xmlns="{NS[1:-1]}"><font><b/></font><numFmt numFmtId="164" formatCode="0.00"/></dxf>')
    sheet.conditional_formats.append(Rule(areas or [Area(1, 1, 6, 1)], root, dxf=patch))
    return app, sheet


def vector(sheet: Worksheet, count: int) -> str:
    with stats.conditional_read():
        return ''.join('1' if display_style(sheet, row, 1).font.bold else '0' for row in range(1, count + 1))


@pytest.mark.parametrize(('attributes', 'expected'), [
    ('rank="2"', '0111'), ('rank="2" bottom="1"', '1110'),
])
def test_top_and_bottom_include_cutoff_ties(attributes: str, expected: str) -> None:
    _, sheet = setup('Range("A1").Value=10\nRange("A2:A3").Value=20\nRange("A4").Value=30',
                      kind='top10', attributes=attributes, areas=[Area(1, 1, 4, 1)])
    assert vector(sheet, 4) == expected


@pytest.mark.parametrize(('rank', 'expected'), [(1, '000001'), (10, '000001'), (20, '000001'),
                                              (25, '000001'), (50, '000111'), (100, '111111')])
def test_native_percent_floors_population_with_one_minimum(rank: int, expected: str) -> None:
    _, sheet = setup('\n'.join(f'Range("A{i}").Value={i}' for i in range(1, 7)),
                      kind='top10', attributes=f'rank="{rank}" percent="1"')
    assert vector(sheet, 6) == expected


@pytest.mark.parametrize(('attributes', 'expected'), [
    ('', '001000'), ('aboveAverage="0"', '100000'),
    ('equalAverage="1"', '011000'), ('aboveAverage="0" equalAverage="1"', '110000'),
])
def test_average_ignores_logical_text_and_blank_values(attributes: str, expected: str) -> None:
    _, sheet = setup('Range("A1").Value=1\nRange("A2").Value=2\nRange("A3").Value=3\nRange("A4").Value=True\nRange("A5").NumberFormat="@"\nRange("A5").Value="9"',
                      kind='aboveAverage', attributes=attributes)
    assert vector(sheet, 6) == expected


@pytest.mark.parametrize(('values', 'attributes', 'expected'), [
    ('Range("A1").Value=3\nRange("A2").Value=2\nRange("A3").Value=0', 'stdDev="1"', '100'),
    ('Range("A1").Value=0\nRange("A2").Value=2', 'stdDev="1"', '01'),
    ('Range("A1").Value=3', 'stdDev="1"', '1'),
])
def test_population_deviation_includes_exact_threshold(values: str, attributes: str, expected: str) -> None:
    _, sheet = setup(values, kind='aboveAverage', attributes=attributes, areas=[Area(1, 1, len(expected), 1)])
    assert vector(sheet, len(expected)) == expected


@pytest.mark.parametrize(('kind', 'expected'), [('uniqueValues', '0011100001001111'),
                                              ('duplicateValues', '1100001110110000')])
def test_native_unique_typed_population(kind: str, expected: str) -> None:
    values = '''Range("A1").Value=1
Range("A2").NumberFormat="@"
Range("A2").Value="1"
Range("A3").Value=True
Range("A4").Value=-1
Range("A5").NumberFormat="@"
Range("A5").Value="TRUE"
Range("A7").Formula="="""""
Range("A8").Value="Foo"
Range("A9").Value="foo"
Range("A10").Value="Bar"
Range("A11:A12").Value=CVErr(2042)
Range("A13").Value=CVErr(2007)
Range("A14").Value=" "
Range("A15").Value="foo "
Range("A16").Value=0'''
    _, sheet = setup(values, kind=kind, areas=[Area(1, 1, 16, 1)])
    assert vector(sheet, 16) == expected


def test_sparse_whole_column_counts_blank_partner_without_walking_empty_cells() -> None:
    _, sheet = setup('Range("A1").Formula="="""""', kind='duplicateValues', areas=[Area(1, 1, 1048576, 1)])
    assert display_style(sheet, 1, 1).font.bold
    assert not display_style(sheet, 900000, 1).font.bold
    assert len(sheet.cells_) == 1


def test_read_scope_builds_population_once_and_next_read_recomputes(monkeypatch: pytest.MonkeyPatch) -> None:
    app, sheet = setup('Range("A1:A100").Value=1', kind='top10', attributes='rank="2"', areas=[Area(1, 1, 100, 1)])
    original: Callable[[Worksheet, Rule], stats.Population] = getattr(stats, '_population')
    calls = 0

    def measured(owner: Worksheet, rule: Rule) -> stats.Population:
        nonlocal calls
        calls += 1
        return original(owner, rule)

    monkeypatch.setattr(stats, '_population', measured)
    assert vector(sheet, 100) == '1' * 100
    assert calls == 1
    app.add_module('Sub Change()\nRange("A2:A100").Value=2\nEnd Sub', name='ChangeModule')
    app.run('Change')
    assert vector(sheet, 100) == '0' + '1' * 99
    assert calls == 2
    assert conditional_number_format(sheet, 1, 1) == 'General'
    assert calls == 3


def test_overlap_union_does_not_count_same_cell_twice() -> None:
    _, sheet = setup('Range("A1").Value=1\nRange("A2").Value=2\nRange("A3").Value=3',
                      kind='uniqueValues', areas=[Area(1, 1, 2, 1), Area(2, 1, 3, 1)])
    assert vector(sheet, 3) == '111'


@pytest.mark.parametrize(('values', 'expected'), [
    (['1', '01', '1.0', '1e0', '+1', ' 1 ', '100%', '$1.00'], '11111111'),
    (['I', 'i', '\u0131', '\u0130'], '1100'),
    (['Stra\u00dfe', 'STRASSE'], '11'),
    (['\u00e9', 'e\u0301'], '11'),
    (['a*', 'abc'], '10'),
    (['a?', 'ab'], '10'),
    (['a~*', 'a*'], '01'),
    (['=abc', 'abc', 'ABC'], '011'),
    (['=a*', 'abc', 'abd'], '100'),
    (['<>a*', 'abc', 'xyz'], '100'),
    (['>a*', 'abc', 'xyz'], '100'),
    (['~~', '~', '~'], '011'),
])
def test_native_numeric_unicode_and_asymmetric_wildcard_duplicates(values: list[str], expected: str) -> None:
    source = '\n'.join(f'Range("A{i}").NumberFormat="@"\nRange("A{i}").Value="{value}"'
                       for i, value in enumerate(values, 1))
    _, sheet = setup(source, kind='duplicateValues', areas=[Area(1, 1, len(values), 1)])
    assert vector(sheet, len(values)) == expected


def test_duplicate_number_precision_uses_fifteen_significant_digits() -> None:
    _, sheet = setup('Range("A1").Value=1\nRange("A2").Value=1.000000000000001\nRange("A3").Value=1.00000000000001',
                      kind='duplicateValues', areas=[Area(1, 1, 3, 1)])
    assert vector(sheet, 3) == '110'


@pytest.mark.parametrize(('pattern', 'expected'), [('?', '00'), ('??', '10'), ('*', '10')])
def test_native_wildcards_count_real_emoji_utf16_units(pattern: str, expected: str) -> None:
    _, sheet = setup(f'Range("A1").Value="{pattern}"\nRange("A2").Value=ChrW(-10179) & ChrW(-8704)',
                      kind='duplicateValues', areas=[Area(1, 1, 2, 1)])
    assert vector(sheet, 2) == expected


def test_unicode_comparison_canonicalizes_explicit_and_scalar_surrogates() -> None:
    _, sheet = setup('Range("A1").Value=ChrW(-10179) & ChrW(-8704)\nRange("A2").Value="\U0001f600"',
                      kind='duplicateValues', areas=[Area(1, 1, 2, 1)])
    assert vector(sheet, 2) == '11'


def test_mutating_rule_inside_one_read_is_reported_instead_of_using_stale_population() -> None:
    _, sheet = setup('Range("A1").Value=1\nRange("A2").Value=2', kind='top10', attributes='rank="1"',
                      areas=[Area(1, 1, 2, 1)])
    with stats.conditional_read():
        assert display_style(sheet, 2, 1).font.bold
        sheet.conditional_formats[0].areas = [Area(1, 1, 1, 1)]
        with pytest.raises(VBAUnsupportedError, match='Changing conditional rules'):
            display_style(sheet, 1, 1)
