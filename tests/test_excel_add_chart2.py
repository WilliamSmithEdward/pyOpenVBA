"""Shapes.AddChart2 and Shapes.AddChart, replayed from live Excel (scripts/measure_add_chart2.py,
tests/fixtures/add_chart2.json): what each case answers and reads, and every chart part it saves, its style and colours
among them. A chart's place when none is given is not stable in Excel, and is left out of the comparison."""

from __future__ import annotations

import hashlib
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._chart_style_table import COLORS_PART, PROVEN
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_chart_markup import normalized

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/add_chart2.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Several areas selected; the chart types the engine does not write.
LEFT_OUT: tuple[str, ...] = (
    'A1:A5,C1:C5',
    'xlBubble',
    'xlRadar',
    'xl3DColumn',
    'ChartObjects(1).RoundedCorners',
    # A size the style does not give, which the engine refuses to read.
    'AddChart2(353, xlColumnClustered): r = n.Chart.ChartTitle.Font.Size',
    # Its legend's paragraph says rtl only after other cases in one run: alone, it does not (live, 2026-11-03).
    'ws.Range("A1:E2").Select',
)
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in LEFT_OUT)]


def _source() -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case '):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept:
            continue
        written.append(line)
    return '\n'.join(written)


def masked(answer: str) -> str:
    """The answer without a place Excel chose."""
    return re.sub(r'\|(?:4\.724409E-04|192)\|(?:4\.724409E-04|103\.125)\|', '|place|place|', answer)


def parts_of(path: Path) -> dict[str, str]:
    """Every part under xl/charts, as the fixture holds them."""
    found: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith('xl/charts/'):
                text = package.read(name).decode('utf-8')
                found[name] = normalized(text) if re.fullmatch(r'xl/charts/chart\d+\.xml', name) else text
    return found


def wanted_parts(held: dict[str, str]) -> dict[str, str]:
    """A case's parts, with the style parts the fixture holds once put back."""
    return {
        name: RECORD['styles'][text] if re.fullmatch(r'xl/charts/(style|colors)\d+\.xml', name) else text
        for name, text in held.items()
    }


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('add_chart2')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='AddChartProbe')
    return {index: str(app.run('AddChartProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert masked(run[0][index]) == masked(RECORD['cases'][index][1])


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    held = RECORD['cases'][index][2]
    if not held:
        return
    assert parts_of(run[1] / f'c{index}.xlsx') == wanted_parts(held)


#: scripts/measure_chart_style_parts.py: the chart part of each type the engine writes at each style a type takes by
#: default.
STYLES: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_style_parts.json').read_text(encoding='utf-8')
)
#: What the table says the engine matches, and what it says it does not, which must refuse.
CHECKED = [key for key in STYLES['checks'] if int(key.split(',')[1]) in PROVEN.get(int(key.split(',')[0]), ())]
REFUSED = [key for key in STYLES['checks'] if key not in CHECKED]
STYLE_PROBE = """Public Function Probe(ByVal path As String, ByVal style As Long, ByVal kind As Long) As String
    Dim wb As Workbook, ws As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:A5").Value = Application.Transpose(Array("k", "a", "b", "c", "d"))
    ws.Range("B1:B5").Value = Application.Transpose(Array("v", 1, 2, 3, 4))
    ws.Range("C1:C5").Value = Application.Transpose(Array("w", 5, 6, 7, 8))
    ws.Range("A1:C5").Select
    ws.Shapes.AddChart2 style, kind
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Probe = "ok"
End Function
"""


@pytest.fixture(scope='module')
def style_app() -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(STYLE_PROBE, name='StyleProbe')
    return app


@pytest.mark.parametrize('key', CHECKED)
def test_each_type_at_each_default_style_as_excel_saved_it(
    style_app: ExcelApplication, tmp_path: Path, key: str
) -> None:
    kind, style = key.split(',')
    style_app.run('StyleProbe.Probe', str(tmp_path / 'p.xlsx'), int(style), int(kind))
    with zipfile.ZipFile(tmp_path / 'p.xlsx') as package:
        part = normalized(package.read('xl/charts/chart1.xml').decode('utf-8'))
    assert hashlib.sha256(part.encode('utf-8')).hexdigest() == STYLES['checks'][key]


def test_most_styles_are_proven_for_the_common_types() -> None:
    assert all(len(PROVEN[kind]) >= 75 for kind in (51, 57, 4, 65, 1, 5, -4120, -4169, 52))


@pytest.mark.parametrize('key', REFUSED[::40])
def test_a_pair_not_proven_says_so(style_app: ExcelApplication, tmp_path: Path, key: str) -> None:
    kind, style = key.split(',')
    with pytest.raises(VBAUnsupportedError, match='is not implemented'):
        style_app.run('StyleProbe.Probe', str(tmp_path / 'p.xlsx'), int(style), int(kind))


@pytest.mark.parametrize('style', sorted(set(range(201, 498)) - PROVEN[51])[:4])
def test_the_other_styles_say_they_are_not_implemented(style: int) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(
        f"""Public Function Probe() As String
    ActiveSheet.Shapes.AddChart2 {style}, xlColumnClustered
End Function
""",
        name='Refused',
    )
    with pytest.raises(VBAUnsupportedError, match=f'Chart style {style} is not implemented'):
        app.run('Refused.Probe')


def test_the_colours_part_is_one_for_every_style() -> None:
    assert COLORS_PART == next(text for key, text in RECORD['styles'].items() if key.startswith('colors'))
