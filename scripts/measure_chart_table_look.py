"""The Border and Format of a chart's data table: what each reads, and the chart part Excel saves. The cases run as
scripts/measure_chart_members_more.py runs its own.

    python scripts/measure_chart_table_look.py

writes tests/fixtures/chart_table_look.json, which tests/test_excel_chart_table_look.py replays.
"""
from __future__ import annotations

from pathlib import Path

import measure_chart_members_more as runner

ON = 'ch.HasDataTable = True: '
T = 'ch.DataTable'
B = f'{T}.Border'
F = f'{T}.Format'
SHOWN = f'{B}.Color & " " & {B}.ColorIndex & " " & {B}.Weight & " " & {B}.LineStyle'
LINE = f'{F}.Line.Visible & " " & {F}.Line.ForeColor.RGB & " " & {F}.Line.Weight & " " & {F}.Line.DashStyle & " " & {F}.Line.Transparency'
FILL = f'{F}.Fill.Visible & " " & {F}.Fill.ForeColor.RGB & " " & {F}.Fill.Type & " " & {F}.Fill.Transparency'
CASES: list[tuple[str, str]] = [('column', f'{ON}r = TypeName({B}) & " " & TypeName({B}.Parent) & " " & TypeName({F}) & " " & TypeName({F}.Parent)'), ('column', f'{ON}r = {SHOWN}'), ('column', f'{ON}r = {LINE}'), ('column', f'{ON}r = {FILL}'), ('line', f'{ON}r = {SHOWN} & " / " & {LINE}')]
CASES += [('column', f'{ON}{B}{change}: r = {SHOWN} & " / " & {LINE}') for change in ('.Color = 255', '.ColorIndex = 3', '.ColorIndex = xlNone', '.ColorIndex = xlAutomatic', '.Weight = xlHairline', '.Weight = xlThin', '.Weight = xlMedium', '.Weight = xlThick', '.LineStyle = xlNone', '.LineStyle = xlContinuous', '.LineStyle = xlDash', '.LineStyle = xlDot', '.LineStyle = xlAutomatic')]
CASES += [('column', f'{ON}{F}.Line{change}: r = {LINE} & " / " & {SHOWN}') for change in ('.ForeColor.RGB = 255', '.Weight = 3', '.DashStyle = msoLineDash', '.Visible = msoFalse', '.Visible = msoTrue', '.Transparency = 0.5')]
CASES += [('column', f'{ON}{F}.Fill{change}: r = {FILL}') for change in ('.ForeColor.RGB = 255', '.Visible = msoFalse', '.Visible = msoTrue', '.Transparency = 0.5', '.Solid')]
CASES += [('column', f'{ON}{F}.Fill.ForeColor.RGB = 255: {F}.Line.ForeColor.RGB = 65280: r = {FILL} & " / " & {LINE}'), ('column', f'{ON}{B}.Color = 255: {T}.Font.Bold = True: r = {SHOWN}'), ('column', f'{ON}{B}.Color = 255: ch.HasDataTable = False: ch.HasDataTable = True: r = {SHOWN}'), ('column', f'{ON}{B}.Color = 255: {T}.Delete: ch.HasDataTable = True: r = {SHOWN}'),
          ('column', f'{ON}{B}.Color = 255: {T}.HasBorderOutline = False: r = {SHOWN} & " " & {T}.HasBorderOutline'), ('column', f'{ON}{T}.HasBorderHorizontal = False: r = {SHOWN}'), ('column', f'{ON}{F}.Glow.Radius = 5: r = {F}.Glow.Radius'), ('column', f'{ON}r = TypeName({F}.TextFrame2) & " " & {F}.TextFrame2.HasText'), ('column', f'{ON}r = {F}.ThreeD.Visible'), ('column', f'{ON}r = TypeName({F}.PictureFormat)'), ('column', f'{ON}r = {F}.AutoShapeType'), ('column', f'{ON}r = {F}.Adjustments.Count')]


def main() -> None:
    runner.OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_table_look.json'
    runner.CASES[:] = CASES
    runner.main()


if __name__ == '__main__':
    main()
