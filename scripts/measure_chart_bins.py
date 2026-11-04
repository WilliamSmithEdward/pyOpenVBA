"""A chart group's histogram bin members on the kinds of chart the model draws: what each reads, and the chart part Excel
saves. The cases run as scripts/measure_chart_members_more.py runs its own.

    python scripts/measure_chart_bins.py

writes tests/fixtures/chart_bins.json, which tests/test_excel_chart_bins.py replays.
"""
from __future__ import annotations

from pathlib import Path

import measure_chart_members_more as runner

G = 'ch.ChartGroups(1)'
MEMBERS = {'BinsType': ('0', '1', '2', '3', '9'), 'BinWidthValue': ('2', '0.5', '0', '-1'), 'BinsCountValue': ('3', '0', '-1'), 'BinsOverflowEnabled': ('True', 'False'), 'BinsOverflowValue': ('5', '-1'),
           'BinsUnderflowEnabled': ('True', 'False'), 'BinsUnderflowValue': ('5', '-1')}
CASES: list[tuple[str, str]] = []
for kind in ('column', 'line', 'pie'):
    for name in MEMBERS:
        CASES.append((kind, f'r = {G}.{name}'))
for name, values in MEMBERS.items():
    CASES += [('column', f'{G}.{name} = {value}: r = {G}.{name}') for value in values]
CASES += [('column', f'{G}.BinsType = 1: r = {G}.BinWidthValue'), ('column', f'{G}.BinWidthValue = 2: r = {G}.BinsType'), ('column', f'{G}.BinsCountValue = 3: r = {G}.BinsType'), ('column', f'{G}.BinsOverflowValue = 5: r = {G}.BinsOverflowEnabled'), ('column', f'{G}.BinsUnderflowValue = 5: r = {G}.BinsUnderflowEnabled'),
          ('column', f'{G}.BinWidthValue = 2: ch.ChartType = xlLine: r = {G}.BinWidthValue')]
for owner, before in (('ch.Axes(xlValue).MajorGridlines', ''), ('ch.SeriesCollection(1).Trendlines(1)', 'ch.SeriesCollection(1).Trendlines.Add: ')):
    CASES += [('column', f'{before}r = {owner}.GetProperty("Name")'), ('column', f'{before}r = {owner}.GetProperty("x")'), ('column', f'{before}{owner}.SetProperty "Name", "z": r = {owner}.Name'), ('column', f'{before}{owner}.SetProperty "x", 1: r = 1')]


def main() -> None:
    runner.OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_bins.json'
    runner.CASES[:] = CASES
    runner.main()


if __name__ == '__main__':
    main()
